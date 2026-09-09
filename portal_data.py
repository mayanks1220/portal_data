"""
portal_data.py
==============================
COMBINED / FINAL sandbox module — merges the best of:
  - module_portal_merge_test.py  (multi-sheet-per-file reading, shift(N)
    continuation-tab consolidation, sheet consolidation report)
  - portal_chat_test.py          (Centre Code padding, tab-level audit,
    final shift-level cross-check, default-column export, grouped
    multi-tab workbook builder)

Nothing from either source file's LOGIC was removed — this file keeps:
  - FR Score descending sort (blank scores sort last)
  - Roll No + Shift as the Unique Key, duplicate flagging on that key
  - Scribers logic (rank 1 = highest FR Score per Unique Key; rank > 1 =
    Scribers, i.e. rows that would NOT be the one merged/kept)
  - Shift extraction from tab name first, filename as fallback (so both
    the OLD one-file-per-shift export and the NEW one-file-multi-tab
    export work without asking which one you uploaded)
  - shift_name(2) / shift_name (2) / shift_name(3) continuation-tab
    folding into one logical shift

WHAT'S NEW IN THIS FILE (per your requirements):
  1. Centre Code padding — independent Yes/No + digit-length control,
     exactly like Roll No's.
  2. Per-tab audit table: File, Tab Name, Logical Shift, Total, Unique,
     Duplicate, Status (OK/ERROR) — one row per raw sheet read.
  3. Final shift-level cross-check: for every LOGICAL shift — how many
     raw tabs fed into it, their names, source-tab total vs. final total,
     unique/duplicate counts, and a MATCH/CHECK flag. This is the "cross
     check everything after merge, before download" table you asked for.
  4. TWO DISTINCT downloads, no hard 10,00,000-row block on either:
       - Whole Data: every shift MIXED TOGETHER in one continuous sheet
         (e.g. "05 Jul'26 S1" and "05 Jul'26 S2" rows sit in the same
         tab). Only spills onto "All Data (2)", "All Data (3)", ... if
         the TOTAL row count crosses the cap — a pure row-count overflow
         split, never a shift-based split.
       - Shift-wise: pick one or more shifts via multiselect; EACH gets
         its own tab(s) — "ShiftName", "ShiftName (2)", "ShiftName (3)",
         ... — so a single shift with 12,00,000+ rows still downloads
         fine in one workbook. This is the button for that exact case.
     (Earlier versions of this file had Whole Data ALSO grouping by
     shift internally, making it look identical to picking every shift
     via Shift-wise — that's fixed now, and a duplicate hidden 'shift'
     column that used to sneak into every export is also fixed.)
  6. Default-column export: Roll No, Name, Centre Code, Enrollment Time,
     Device, BioDevice, Photo Matched, FR Score, Bio Count, Operator —
     PLUS every calculated column (Attendance, Unique Key, Duplicate
     Flag, Shift, EXAM_CODE, Source File, Source Sheet) automatically.
     Any other column is available via multiselect and gets appended if
     you pick it.
  7. Progress bar on every "Prepare & Download" action, showing % while
     the file is being built server-side. (Honesty note: this reflects
     file-generation progress on the server — Streamlit has no way to
     hook into the browser's actual download-transfer progress, since
     the file is already fully built in memory before the download
     button appears and the browser downloads it instantly.)
  8. Defensive reading: a broken/corrupted sheet no longer crashes the
     whole run — it's caught, marked ERROR in the tab audit, and the run
     continues with everything that DID read cleanly.

Run standalone with: streamlit run module_portal_merge_final.py
"""

import io
import os
import re
import math
from concurrent.futures import ThreadPoolExecutor, as_completed
from contextlib import contextmanager

import numpy as np
import pandas as pd
import streamlit as st
import openpyxl

# ══════════════════════════════════════════════════════════════════════
# CONSTANTS
# ══════════════════════════════════════════════════════════════════════
APP_TITLE = "Exam Portal Data Processing, Validation & Reporting System"
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

# Hard Excel per-sheet row cap is 1,048,576 including the header row.
# We leave a little headroom and let the user tune the practical cap.
EXCEL_HARD_ROW_LIMIT = 1_048_576
DEFAULT_MAX_ROWS_PER_TAB = 1_000_000

# Case/space-insensitive column alias matching (same idea as utils.ALIAS_MAP)
ALIASES = {
    "roll no": ["roll no", "roll_no", "rollno", "roll number", "roll_number"],
    "name": ["name", "candidate name", "candidate_name", "student name", "full name"],
    "centre code": ["centre code", "center code", "centre_code", "center_code", "centrecode", "centercode"],
    "center name": ["center name", "centre name", "centername", "centrename"],
    "enrollment time": ["enrollment time", "enrolment time", "enrollment_time", "enrolment_time"],
    "photo matched": ["photo matched", "photo_matched", "photomatched", "photo match"],
    "fr score": ["fr score", "fr_score", "frscore"],
    "shift": ["shift", "exam shift"],
    "device": ["device", "device name", "devicename"],
    "biodevice": ["biodevice", "bio device", "bio_device"],
    "bio count": ["bio count", "biocount", "bio_count"],
    "operator": ["operator", "operator name", "opname", "operatorname"],
}

# The exact default export column set you asked for, in this order.
# These are the CANONICAL (post-normalize) lowercase keys — see ALIASES.
DEFAULT_EXPORT_COLUMNS = [
    "roll no", "name", "centre code", "enrollment time",
    "device", "biodevice", "photo matched", "fr score", "bio count", "operator",
]

# Nicer display names applied ONLY on the exported copy (never on the
# working dataframe, so nothing downstream breaks).
EXPORT_DISPLAY_NAMES = {
    "roll no": "Roll No", "name": "Name", "centre code": "Centre Code",
    "enrollment time": "Enrollment Time", "device": "Device", "biodevice": "BioDevice",
    "photo matched": "Photo Matched", "fr score": "FR Score", "bio count": "Bio Count",
    "operator": "Operator", "shift": "Shift",
}

# Calculated / system columns — always included automatically, appended
# after the default + any extra user-picked columns.
SYSTEM_COLUMNS = ["Attendance", "Unique Key", "Duplicate Flag", "shift",
                   "EXAM_CODE", "Source File", "Source Sheet"]

SHIFT_REGEX = re.compile(r"(\d{1,2}\s+[A-Za-z]{3}'\d{2}(?:\s+[Ss]\d+)?)")
# Matches a trailing continuation suffix on a sheet/tab name, with or
# without a leading space: "21 Jun'26(2)" and "21 Jun'26 (2)" both fold
# into "21 Jun'26".
CONTINUATION_SUFFIX_REGEX = re.compile(r"\s*\(\d+\)\s*$")
INVALID_SHEET_CHARS = re.compile(r"[\[\]\:\*\?/\\]")


# ══════════════════════════════════════════════════════════════════════
# PAGE CONFIG + STYLE
# ══════════════════════════════════════════════════════════════════════
def _configure_page():
    st.set_page_config(page_title=APP_TITLE, page_icon="📊", layout="wide")
    st.markdown(
        """
        <style>
        /* ─────────────────────────────────────────────────────────────
           PROFESSIONAL UI THEME — presentation only.
           Processing / merge / export logic remains unchanged.
           Uses Streamlit theme variables so text remains readable in
           BOTH Light and Dark mode.
           ───────────────────────────────────────────────────────────── */
        :root {
            --pmf-primary: #2563eb;
            --pmf-primary-dark: #1d4ed8;
            --pmf-success: #16a34a;
            --pmf-danger: #dc2626;
            --pmf-warning: #d97706;
            --pmf-info: #0891b2;
            --pmf-border: rgba(128, 128, 128, 0.25);
            --pmf-card: var(--secondary-background-color);
            --pmf-text: var(--text-color);
            --pmf-muted: rgba(128, 128, 128, 0.95);
        }

        html, body, [class*="css"] {
            font-family: "Segoe UI", Inter, Arial, sans-serif;
        }

        .stApp {
            background: var(--background-color);
            color: var(--pmf-text);
        }

        .main .block-container {
            max-width: 1500px;
            padding: 1.5rem 2rem 3rem;
        }

        /* Header */
        .app-title {
            font-size: clamp(26px, 3vw, 36px);
            line-height: 1.15;
            font-weight: 850;
            letter-spacing: -0.7px;
            color: var(--pmf-text) !important;
            margin: 0 0 5px;
        }

        .app-sub {
            color: var(--pmf-muted) !important;
            font-size: 13.5px;
            line-height: 1.55;
            margin-bottom: 18px;
        }

        /* Section headings */
        .sec-title {
            display: flex;
            align-items: center;
            min-height: 42px;
            box-sizing: border-box;
            font-size: 14px;
            font-weight: 850;
            text-transform: uppercase;
            letter-spacing: .75px;
            color: var(--pmf-text) !important;
            border: 1px solid var(--pmf-border);
            border-left: 5px solid var(--pmf-primary);
            border-radius: 10px;
            background: var(--pmf-card);
            padding: 9px 14px;
            margin: 24px 0 12px;
            box-shadow: 0 2px 10px rgba(0,0,0,.04);
        }

        /* KPI cards */
        .kpi-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(145px, 1fr));
            gap: 12px;
            margin: 8px 0 18px;
        }

        .kpi-card {
            position: relative;
            overflow: hidden;
            min-height: 92px;
            box-sizing: border-box;
            background: var(--pmf-card);
            color: var(--pmf-text) !important;
            border: 1px solid var(--pmf-border);
            border-radius: 12px;
            padding: 13px 14px 12px;
            box-shadow: 0 3px 12px rgba(0,0,0,.055);
            transition: transform .15s ease, box-shadow .15s ease;
        }

        .kpi-card:hover {
            transform: translateY(-1px);
            box-shadow: 0 6px 18px rgba(0,0,0,.09);
        }

        .kpi-card::before {
            content: "";
            position: absolute;
            top: 0; left: 0; right: 0;
            height: 4px;
        }

        .kpi-icon { font-size: 18px; line-height: 1; margin-bottom: 8px; }
        .kpi-label {
            font-size: 10px;
            font-weight: 750;
            text-transform: uppercase;
            letter-spacing: .45px;
            color: var(--pmf-muted) !important;
            margin-bottom: 3px;
            white-space: nowrap;
            overflow: hidden;
            text-overflow: ellipsis;
        }
        .kpi-value {
            font-size: 22px;
            line-height: 1.15;
            font-weight: 850;
            color: var(--pmf-text) !important;
        }

        .kpi-blue::before { background:#2563eb; }
        .kpi-green::before { background:#16a34a; }
        .kpi-red::before { background:#ef4444; }
        .kpi-purple::before { background:#8b5cf6; }
        .kpi-amber::before { background:#f59e0b; }
        .kpi-cyan::before { background:#06b6d4; }
        .kpi-slate::before { background:#475569; }
        .kpi-pink::before { background:#ec4899; }

        /* Cards / explanatory notes */
        .note-box {
            padding: 13px 16px;
            border: 1px solid rgba(37,99,235,.35);
            border-left: 5px solid var(--pmf-primary);
            border-radius: 10px;
            background: color-mix(in srgb, var(--pmf-primary) 8%, var(--pmf-card));
            font-size: 13px;
            line-height: 1.55;
            color: var(--pmf-text) !important;
            margin: 10px 0 14px;
        }

        .note-box b, .note-box code { color: var(--pmf-text) !important; }

        /* Global Streamlit text readability */
        .stMarkdown, .stMarkdown p, .stMarkdown span,
        .stCaption, .stCaption p,
        [data-testid="stWidgetLabel"], [data-testid="stWidgetLabel"] p,
        [data-testid="stFileUploaderDropzone"] p,
        label, .stRadio label, .stCheckbox label, .stSelectbox label,
        .stMultiSelect label, .stNumberInput label, .stTextInput label {
            color: var(--pmf-text) !important;
        }

        .stCaption, .stCaption p,
        [data-testid="stCaptionContainer"] {
            color: var(--pmf-muted) !important;
        }

        /* Inputs */
        div[data-baseweb="input"],
        div[data-baseweb="select"],
        div[data-baseweb="textarea"] {
            border-radius: 9px;
        }

        div[data-baseweb="input"] > div,
        div[data-baseweb="select"] > div,
        div[data-baseweb="textarea"] > div {
            border-color: var(--pmf-border) !important;
            background: var(--pmf-card) !important;
        }

        input, textarea {
            color: var(--pmf-text) !important;
            caret-color: var(--pmf-primary);
        }

        input::placeholder, textarea::placeholder {
            color: var(--pmf-muted) !important;
            opacity: .8;
        }

        /* Buttons */
        .stButton > button, .stDownloadButton > button {
            min-height: 42px;
            border-radius: 9px;
            border: 1px solid var(--pmf-primary);
            font-weight: 750;
            letter-spacing: .1px;
            transition: all .15s ease;
        }

        .stButton > button {
            background: var(--pmf-primary);
            color: #ffffff !important;
        }

        .stButton > button:hover {
            background: var(--pmf-primary-dark);
            border-color: var(--pmf-primary-dark);
            color: #ffffff !important;
            box-shadow: 0 4px 12px rgba(37,99,235,.25);
        }

        .stDownloadButton > button {
            background: transparent;
            color: var(--pmf-primary) !important;
        }

        .stDownloadButton > button:hover {
            background: rgba(37,99,235,.08);
            color: var(--pmf-primary-dark) !important;
        }

        /* Tables */
        [data-testid="stDataFrame"] {
            border: 1px solid var(--pmf-border);
            border-radius: 10px;
            overflow: hidden;
            box-shadow: 0 2px 10px rgba(0,0,0,.035);
        }

        /* File uploader */
        [data-testid="stFileUploaderDropzone"] {
            border: 1.5px dashed rgba(37,99,235,.45) !important;
            border-radius: 12px !important;
            background: color-mix(in srgb, var(--pmf-primary) 4%, var(--pmf-card)) !important;
            padding: 18px !important;
        }

        [data-testid="stFileUploaderDropzone"] button {
            border-radius: 8px !important;
        }

        /* Alerts */
        [data-testid="stAlert"] {
            border-radius: 10px;
        }

        /* Progress */
        [data-testid="stProgressBar"] {
            border-radius: 999px;
            overflow: hidden;
        }

        /* Expanders / containers */
        [data-testid="stExpander"] {
            border: 1px solid var(--pmf-border);
            border-radius: 10px;
            background: var(--pmf-card);
        }

        /* Radio / checkbox accent */
        [data-baseweb="radio"] div[role="radio"] {
            border-color: var(--pmf-primary);
        }

        /* Responsive spacing */
        @media (max-width: 900px) {
            .main .block-container { padding: 1rem 1rem 2rem; }
            .sec-title { font-size: 12.5px; }
            .kpi-grid { grid-template-columns: repeat(2, minmax(0, 1fr)); }
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def render_kpi_cards(items):
    html = ['<div class="kpi-grid">']
    for it in items:
        html.append(
            f'<div class="kpi-card kpi-{it["color"]}">'
            f'<div class="kpi-icon">{it["icon"]}</div>'
            f'<div class="kpi-label">{it["label"]}</div>'
            f'<div class="kpi-value">{it["value"]}</div>'
            f'</div>'
        )
    html.append("</div>")
    st.markdown("".join(html), unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════
# SMALL TEXT / NORMALIZATION HELPERS
# ══════════════════════════════════════════════════════════════════════
def txt(x) -> str:
    return "" if pd.isna(x) else str(x).strip()

def norm(x) -> str:
    return re.sub(r"\s+", " ", re.sub(r"[_\-]+", " ", txt(x).lower())).strip()

def build_reverse_alias_map(alias_map):
    reverse = {}
    for canonical, aliases in alias_map.items():
        for a in aliases:
            reverse[norm(a)] = canonical
    return reverse

REVERSE_ALIAS = build_reverse_alias_map(ALIASES)

def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Case/space-insensitive rename to canonical names. Guards against
    two different source columns both trying to map to the same
    canonical name (keeps the first, leaves the second as-is) so a rename
    collision can never silently drop data."""
    used = set()
    rename = {}
    for c in df.columns:
        canonical = REVERSE_ALIAS.get(norm(c))
        if canonical and canonical not in used:
            rename[c] = canonical
            used.add(canonical)
    return df.rename(columns=rename)

def resolve_column(df: pd.DataFrame, label: str):
    """Case/space-insensitive lookup of `label` among df's actual columns."""
    if label in df.columns:
        return label
    target = norm(label)
    for c in df.columns:
        if norm(c) == target:
            return c
    return None

def clean_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    for c in df.columns:
        if df[c].dtype == object:
            df[c] = df[c].map(txt)
            df[c] = df[c].replace({"nan": "", "None": ""})
    return df


# ══════════════════════════════════════════════════════════════════════
# ENGINE SELECTION
# ══════════════════════════════════════════════════════════════════════
def pick_read_engine() -> str:
    try:
        import python_calamine  # noqa: F401
        return "calamine"
    except ImportError:
        return "openpyxl"

def read_engine_info() -> str:
    if pick_read_engine() == "calamine":
        return "⚡ Using **calamine** engine (fast Rust-based reader)."
    return ("🐢 Using **openpyxl** engine. Install `python-calamine` for a "
            "big read-speed boost on large files.")


# ══════════════════════════════════════════════════════════════════════
# SHIFT / EXAM CODE EXTRACTION  (tab name first, filename fallback)
# ══════════════════════════════════════════════════════════════════════
def normalize_sheet_shift_name(sheet_name: str) -> str:
    """Strips a trailing continuation suffix — '(2)', ' (2)', '(3)', ... —
    so 'shift_name', 'shift_name(2)', 'shift_name (3)' all fold into the
    SAME logical shift. This exists purely because a shift's data got
    split across multiple tabs once it passed the per-tab row cap."""
    return CONTINUATION_SUFFIX_REGEX.sub("", txt(sheet_name)).strip()

def extract_shift(sheet_name: str, file_name: str) -> str:
    """Prefer the de-suffixed tab name if it looks like a real shift
    (new multi-tab-per-file format); otherwise fall back to the filename
    (old one-file-per-shift format) — so both formats work without
    telling the app which one you uploaded."""
    cleaned = normalize_sheet_shift_name(sheet_name)
    m = SHIFT_REGEX.search(cleaned)
    if m:
        return m.group(1).strip()
    m = SHIFT_REGEX.search(file_name)
    if m:
        return m.group(1).strip()
    return cleaned if cleaned.lower() not in ("sheet1", "sheet", "") else "Unknown"

def extract_exam_code(file_name: str) -> str:
    base = os.path.splitext(os.path.basename(file_name))[0]
    return base.split("_")[0] if base else "Unknown"


# ══════════════════════════════════════════════════════════════════════
# READING — every sheet in every file, defensively (bad sheet ≠ crash)
# ══════════════════════════════════════════════════════════════════════
def read_one_workbook(file_name: str, file_bytes: bytes, skip_rows: int) -> list:
    """Reads every sheet in one workbook. Returns a list of result dicts:
    {file, sheet, shift, df, rows, status, error}. A sheet that fails to
    read (corrupt, wrong shape, etc.) is captured with status='ERROR' and
    an error message instead of raising — one bad tab never kills the
    whole merge."""
    results = []
    try:
        engine = pick_read_engine()
        try:
            xls = pd.ExcelFile(io.BytesIO(file_bytes), engine=engine)
        except Exception:
            xls = pd.ExcelFile(io.BytesIO(file_bytes), engine="openpyxl")
    except Exception as e:
        return [{"file": file_name, "sheet": "(workbook)", "shift": "Unknown",
                  "df": None, "rows": 0, "status": "ERROR", "error": f"Could not open workbook: {e}"}]

    exam_code = extract_exam_code(file_name)
    for sheet_name in xls.sheet_names:
        shift = extract_shift(sheet_name, file_name)
        try:
            df = pd.read_excel(xls, sheet_name=sheet_name, skiprows=int(skip_rows), dtype=str)
            df = df.dropna(how="all")
            if df.empty:
                results.append({"file": file_name, "sheet": sheet_name, "shift": shift,
                                 "df": None, "rows": 0, "status": "EMPTY", "error": ""})
                continue
            df = normalize_columns(df)
            df = clean_dataframe(df)
            df["shift"] = shift
            df["EXAM_CODE"] = exam_code
            df["Source File"] = file_name
            df["Source Sheet"] = sheet_name
            results.append({"file": file_name, "sheet": sheet_name, "shift": shift,
                             "df": df, "rows": len(df), "status": "OK", "error": ""})
        except Exception as e:
            results.append({"file": file_name, "sheet": sheet_name, "shift": shift,
                             "df": None, "rows": 0, "status": "ERROR", "error": str(e)})
    return results


@st.cache_data(show_spinner=False)
def load_all_workbooks(files_payload, skip_rows: int) -> list:
    """files_payload: tuple of (name, bytes) — cached on (names, sizes, skip_rows)."""
    all_results = []
    with ThreadPoolExecutor(max_workers=min(8, max(1, len(files_payload)))) as ex:
        futures = {ex.submit(read_one_workbook, name, data, skip_rows): name for name, data in files_payload}
        for fut in as_completed(futures):
            all_results.extend(fut.result())
    order = {name: i for i, (name, _) in enumerate(files_payload)}
    all_results.sort(key=lambda x: (order.get(x["file"], 999999), str(x["sheet"]).casefold()))
    return all_results


def build_tab_audit(items: list) -> pd.DataFrame:
    """One row per raw sheet/tab read: File, Tab Name, Logical Shift,
    Total Count, Unique Roll No + Shift, Duplicate Count, Status."""
    rows = []
    for it in items:
        df = it["df"]
        status = it["status"]
        if df is None:
            rows.append({"File Name": it["file"], "Tab Name": it["sheet"], "Logical Shift": it["shift"],
                         "Total Count": 0, "Unique Roll No + Shift": 0, "Duplicate Count": 0,
                         "Status": status, "Note": it.get("error", "")})
            continue
        if "roll no" in df.columns:
            key = (df["roll no"].fillna("").astype(str).str.strip()
                   + df["shift"].fillna("").astype(str).str.strip())
            key = key[key != ""]
            u = int(key.nunique())
            d = int(len(key) - u)
        else:
            u = d = 0
        rows.append({"File Name": it["file"], "Tab Name": it["sheet"], "Logical Shift": it["shift"],
                     "Total Count": len(df), "Unique Roll No + Shift": u, "Duplicate Count": d,
                     "Status": status, "Note": ""})
    return pd.DataFrame(rows)


# ══════════════════════════════════════════════════════════════════════
# PROCESSING PIPELINE  (attendance, padding, unique key, dup flag, FR sort)
# ══════════════════════════════════════════════════════════════════════
def apply_attendance(df: pd.DataFrame) -> pd.DataFrame:
    if "enrollment time" not in df.columns:
        return df
    e = df["enrollment time"].fillna("").astype(str).str.strip().str.upper()
    df = df.copy()
    df["Attendance"] = np.where(e.eq("PENDING"), "Absent", "Present")
    return df

def apply_padding(df: pd.DataFrame, column: str, enabled: bool, digits: int) -> pd.DataFrame:
    """Generic left-zero padding — used for BOTH Roll No and Centre Code,
    each with its own independent Yes/No + digit-length control."""
    if column not in df.columns:
        return df
    df = df.copy()
    df[column] = df[column].fillna("").astype(str).str.strip()
    if enabled:
        df[column] = df[column].str.zfill(int(digits))
    return df

def add_unique_key(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    roll = df.get("roll no", pd.Series("", index=df.index)).fillna("").astype(str).str.strip()
    shift = df.get("shift", pd.Series("", index=df.index)).fillna("").astype(str).str.strip()
    df["Unique Key"] = roll + shift
    return df

def flag_duplicates(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    key = df.get("Unique Key", pd.Series("", index=df.index))
    df["Duplicate Flag"] = np.where((key != "") & key.duplicated(keep=False), "Yes", "No")
    return df

def sort_by_fr_score(df: pd.DataFrame) -> pd.DataFrame:
    """FR Score descending (highest confidence first); blank/unparseable
    scores sort last, never first."""
    if "fr score" not in df.columns:
        return df.reset_index(drop=True)
    work = df.copy()
    work["__fr__"] = pd.to_numeric(work["fr score"], errors="coerce")
    work = work.sort_values("__fr__", ascending=False, na_position="last", kind="stable")
    return work.drop(columns="__fr__").reset_index(drop=True)

def unique_roll_shift_count(df: pd.DataFrame) -> int:
    if "Unique Key" in df.columns:
        key = df["Unique Key"].replace("", np.nan)
        return int(key.nunique())
    return 0

def build_scribers_data(df: pd.DataFrame):
    """rank 1 (per Unique Key) = highest FR Score = the row that would
    actually be merged/kept. rank > 1 = Scribers — rows that did NOT win
    that slot. Returns (duplicate_ranked_df, scribers_df)."""
    if "Unique Key" not in df.columns:
        return pd.DataFrame(), pd.DataFrame()
    work = df.copy()
    work["__fr__"] = pd.to_numeric(work.get("fr score", pd.Series(dtype=str)), errors="coerce")
    counts = work["Unique Key"].value_counts()
    dup_keys = counts[counts > 1].index
    work = work[work["Unique Key"].isin(dup_keys)].copy()
    if work.empty:
        return pd.DataFrame(), pd.DataFrame()
    work = work.sort_values(["Unique Key", "__fr__"], ascending=[True, False], na_position="last", kind="stable")
    work["rank"] = work.groupby("Unique Key", sort=False).cumcount() + 1
    work["duplicate_count"] = work.groupby("Unique Key")["Unique Key"].transform("count")
    duplicate_df = work.drop(columns="__fr__")
    scribers_df = duplicate_df[duplicate_df["rank"] > 1].copy()
    return duplicate_df, scribers_df


# ══════════════════════════════════════════════════════════════════════
# CROSS-CHECK / SUMMARY BUILDERS
# ══════════════════════════════════════════════════════════════════════
def build_final_shift_audit(final_df: pd.DataFrame, tab_audit_df: pd.DataFrame) -> pd.DataFrame:
    """THE cross-check table: for every LOGICAL shift — how many raw tabs
    fed into it (and their names), the source-tab total vs. the final
    (post-merge) total, unique/duplicate counts, and a MATCH/CHECK flag.
    Use this to visually confirm shift(2)/shift(3) consolidation worked
    and no rows were silently lost or duplicated."""
    if final_df.empty:
        return pd.DataFrame()

    ok_tabs = tab_audit_df[tab_audit_df["Status"] == "OK"]
    if ok_tabs.empty:
        source_side = pd.DataFrame(columns=["Logical Shift", "Tab Count", "Available Tab Names", "Source Tab Total"])
    else:
        source_side = (
            ok_tabs.groupby("Logical Shift", sort=False)
            .agg(**{
                "Tab Count": ("Tab Name", "count"),
                "Available Tab Names": ("Tab Name", lambda x: " | ".join(map(str, x))),
                "Source Tab Total": ("Total Count", "sum"),
            })
            .reset_index()
        )

    final_rows = []
    for shift_val, g in final_df.groupby("shift", sort=False, dropna=False):
        key = g.get("Unique Key", pd.Series(dtype=str)).replace("", np.nan)
        u = int(key.nunique())
        final_rows.append({
            "Logical Shift": shift_val, "Final Total": len(g),
            "Unique Roll No + Shift": u, "Duplicate Count": max(len(g) - u, 0),
        })
    final_side = pd.DataFrame(final_rows)

    merged = source_side.merge(final_side, on="Logical Shift", how="outer")
    merged["Tab Count"] = merged["Tab Count"].fillna(0).astype(int)
    merged["Source Tab Total"] = merged["Source Tab Total"].fillna(0).astype(int)
    merged["Final Total"] = merged["Final Total"].fillna(0).astype(int)
    merged["Unique Roll No + Shift"] = merged["Unique Roll No + Shift"].fillna(0).astype(int)
    merged["Duplicate Count"] = merged["Duplicate Count"].fillna(0).astype(int)
    merged["Available Tab Names"] = merged["Available Tab Names"].fillna("")
    merged["Rows Difference"] = merged["Final Total"] - merged["Source Tab Total"]
    merged["Cross Check"] = np.where(merged["Rows Difference"] == 0, "✅ MATCH", "⚠️ CHECK")
    return merged.sort_values("Logical Shift").reset_index(drop=True)

def build_shift_summary(df: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for shift_val, g in df.groupby("shift", sort=False, dropna=False):
        key = g.get("Unique Key", pd.Series(dtype=str)).replace("", np.nan)
        u = int(key.nunique())
        rows.append({
            "Shift": shift_val, "Total Count": len(g), "Unique Roll No + Shift": u,
            "Duplicate Count": max(len(g) - u, 0),
            "Present": int((g.get("Attendance", pd.Series(dtype=str)) == "Present").sum()),
            "Absent": int((g.get("Attendance", pd.Series(dtype=str)) == "Absent").sum()),
        })
    return pd.DataFrame(rows).sort_values("Shift").reset_index(drop=True) if rows else pd.DataFrame()

def build_center_shift_pivot(df: pd.DataFrame) -> pd.DataFrame:
    if not {"centre code", "shift", "Attendance"}.issubset(df.columns):
        return pd.DataFrame()
    present = df[df["Attendance"] == "Present"]
    if present.empty:
        return pd.DataFrame()
    pivot = present.groupby(["centre code", "shift"]).size().unstack(fill_value=0).reset_index()
    shift_cols = [c for c in pivot.columns if c != "centre code"]
    pivot["Grand Total (Present)"] = pivot[shift_cols].sum(axis=1) if shift_cols else 0
    return pivot

def build_photo_matched_breakdown(df: pd.DataFrame) -> pd.DataFrame:
    if "photo matched" not in df.columns or "Unique Key" not in df.columns:
        return pd.DataFrame()
    pm = df["photo matched"].replace("", "Blank").fillna("Blank")
    return (
        df.assign(__pm__=pm).groupby("__pm__")["Unique Key"].nunique()
        .reset_index().rename(columns={"__pm__": "Photo Matched", "Unique Key": "Unique Roll No Count"})
        .sort_values("Unique Roll No Count", ascending=False).reset_index(drop=True)
    )


# ══════════════════════════════════════════════════════════════════════
# EXPORT COLUMN SELECTION
# ══════════════════════════════════════════════════════════════════════
def resolve_default_columns(df: pd.DataFrame) -> list:
    """Which of the default export columns actually exist in this data
    (missing ones — e.g. no 'Bio Count' column at all — are silently
    skipped rather than erroring)."""
    resolved = []
    for label in DEFAULT_EXPORT_COLUMNS:
        actual = resolve_column(df, label)
        if actual and actual not in resolved:
            resolved.append(actual)
    return resolved

def build_export_columns(df: pd.DataFrame, extra_selected: list) -> list:
    cols = resolve_default_columns(df)
    for c in extra_selected:
        if c in df.columns and c not in cols:
            cols.append(c)
    for c in SYSTEM_COLUMNS:
        if c in df.columns and c not in cols:
            cols.append(c)
    return cols

def prepare_export_dataframe(df: pd.DataFrame, extra_selected: list) -> pd.DataFrame:
    cols = build_export_columns(df, extra_selected)
    out = df[cols].copy()
    return out.rename(columns={k: v for k, v in EXPORT_DISPLAY_NAMES.items() if k in out.columns})


# ══════════════════════════════════════════════════════════════════════
# EXCEL WRITE HELPERS  — openpyxl WRITE-ONLY mode (fastest verified option)
# ══════════════════════════════════════════════════════════════════════
# Benchmarked against 3 alternatives on 300,000 rows × 13 columns:
#   pandas+xlsxwriter (no constant_memory): 53.4s
#   pandas+openpyxl (normal mode):          57.8s
#   direct xlsxwriter constant_memory:      31.3s
#   openpyxl write-only + ws.append():      27.2s   <- fastest AND correct
# xlsxwriter's constant_memory option (used in an earlier version of this
# file) was proven via round-trip test to silently blank out every column
# except the first on row 0 of every sheet when combined with pandas'
# to_excel() — a serious, silent data-corruption risk. openpyxl write-only
# mode sidesteps that entirely: values are appended as plain Python lists,
# verified correct on round-trip, and every cell stays a genuine STRING
# type (not a number), so leading zeros in Roll No / Centre Code are never
# stripped by Excel regardless of any number format.

def make_sheet_name(base: str, index: int, used: set) -> str:
    """Builds a valid Excel sheet name, folding to the 'Shift', 'Shift (2)'
    continuation convention, sanitizing invalid characters, truncating to
    31 chars, and avoiding collisions if two different shift names would
    otherwise truncate to the same 31-char string."""
    base_clean = INVALID_SHEET_CHARS.sub("_", txt(base)) or "Data"
    suffix = "" if index == 1 else f" ({index})"
    max_base_len = max(1, 31 - len(suffix))
    candidate = (base_clean[:max_base_len] + suffix)[:31]
    bump = 2
    while candidate.casefold() in used:
        extra = f" ~{bump}"
        candidate = (base_clean[: max(1, 31 - len(extra))] + extra)[:31]
        bump += 1
    used.add(candidate.casefold())
    return candidate

def _write_rows_streaming(ws, df: pd.DataFrame, progress_cb=None,
                           progress_start: float = 0.0, progress_span: float = 95.0,
                           sheet_label: str = "") -> None:
    """Streams a dataframe's header + rows into an openpyxl write-only
    worksheet as plain Python lists. Reports progress in ~25 small batches
    PER SHEET (not just once per sheet) so the bar keeps moving smoothly
    even while a single huge (e.g. 10-12 lakh row) sheet is being written,
    instead of jumping only at sheet boundaries and appearing to 'stall'."""
    ws.append(list(df.columns))
    n = len(df)
    if n == 0:
        return
    batch_size = max(2000, n // 25)
    values = df.values.tolist()
    for i, row in enumerate(values):
        ws.append(row)
        if progress_cb and ((i + 1) % batch_size == 0 or (i + 1) == n):
            frac = (i + 1) / n
            pct = progress_start + frac * progress_span
            label = f"{sheet_label} — " if sheet_label else ""
            progress_cb(pct, f"{label}row {i + 1:,}/{n:,}")

def excel_bytes_simple(df: pd.DataFrame, progress_cb=None) -> bytes:
    """Single-sheet export for small/aux tables (audits, summaries, previews)."""
    wb = openpyxl.Workbook(write_only=True)
    ws = wb.create_sheet("Data")
    _write_rows_streaming(ws, df.fillna(""), progress_cb=progress_cb, progress_start=0.0, progress_span=90.0)
    if progress_cb:
        progress_cb(95, "Saving file…")
    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    data = buffer.getvalue()
    if progress_cb:
        progress_cb(100, "✅ Ready!")
    return data

def build_single_sheet_workbook(export_df: pd.DataFrame, max_rows_per_tab: int,
                                 progress_cb=None, base_sheet_name: str = "All Data") -> bytes:
    """WHOLE-DATA download: every row, every shift, mixed together in ONE
    continuous sheet — exactly the current sort order (e.g. '05 Jul'26 S1'
    and '05 Jul'26 S2' rows sit right next to each other, not split into
    separate tabs). Only continues onto 'All Data (2)', 'All Data (3)', ...
    if the TOTAL row count crosses max_rows_per_tab — this is a pure
    row-count overflow split, never a shift-based split."""
    n = len(export_df)
    if n == 0:
        raise ValueError("No data available to export.")
    n_chunks = max(1, math.ceil(n / max_rows_per_tab))
    used_names = set()

    wb = openpyxl.Workbook(write_only=True)
    for i in range(n_chunks):
        start = i * max_rows_per_tab
        chunk = export_df.iloc[start:start + max_rows_per_tab].fillna("")
        sheet = make_sheet_name(base_sheet_name, i + 1, used_names)
        ws = wb.create_sheet(sheet)
        span_start = (i / n_chunks) * 95.0
        span_size = 95.0 / n_chunks
        _write_rows_streaming(ws, chunk, progress_cb=progress_cb, progress_start=span_start,
                               progress_span=span_size, sheet_label=f"'{sheet}' ({i + 1}/{n_chunks})")
    if progress_cb:
        progress_cb(96, "Saving workbook to file…")
    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    data = buffer.getvalue()
    if progress_cb:
        progress_cb(100, "✅ Ready!")
    return data


def build_grouped_workbook(export_df: pd.DataFrame, group_key: pd.Series, shift_filter,
                            max_rows_per_tab: int, progress_cb=None) -> bytes:
    """SHIFT-WISE download: writes each selected shift's rows into its own
    tab(s) — 'ShiftName', 'ShiftName (2)', 'ShiftName (3)', ... — so a
    single shift with 12,00,000+ rows still fits in one workbook by
    continuing onto extra tabs. This is THE fix for the "single shift is
    12 lakh rows" case.

    `group_key` is a Series of raw shift values, POSITION-ALIGNED with
    export_df (same length, same row order) but kept SEPARATE from
    export_df's own columns — this avoids writing a duplicate/hidden
    'shift' column into the sheet alongside the already-present, nicely
    labeled 'Shift' export column.

    shift_filter: a list of one or more shift names to include (required
    for this function — pass None/[] only if you genuinely want every
    shift each in its own tab, which is rarely what you want for a
    "whole data" download; use build_single_sheet_workbook for that).
    """
    df_reset = export_df.reset_index(drop=True)
    key_reset = pd.Series(group_key).reset_index(drop=True)

    groups = []
    for shift_val in pd.unique(key_reset):
        mask = (key_reset == shift_val).values
        groups.append((shift_val, df_reset[mask]))

    if shift_filter:
        wanted = {txt(s).casefold() for s in shift_filter}
        groups = [(s, g) for s, g in groups if txt(s).casefold() in wanted]

    plan = []
    total_units = 0
    for shift_val, g in groups:
        if g.empty:
            continue
        n_chunks = max(1, math.ceil(len(g) / max_rows_per_tab))
        plan.append((shift_val, g, n_chunks))
        total_units += n_chunks

    if total_units == 0:
        raise ValueError("No matching shift data found to export.")

    used_names = set()
    wb = openpyxl.Workbook(write_only=True)
    done_units = 0
    for shift_val, g, n_chunks in plan:
        shift_label = txt(shift_val) or "Unknown"
        for i in range(n_chunks):
            start = i * max_rows_per_tab
            chunk = g.iloc[start:start + max_rows_per_tab].fillna("")
            sheet = make_sheet_name(shift_label, i + 1, used_names)
            ws = wb.create_sheet(sheet)
            span_start = (done_units / total_units) * 95.0
            span_size = (1 / total_units) * 95.0
            _write_rows_streaming(ws, chunk, progress_cb=progress_cb, progress_start=span_start,
                                   progress_span=span_size, sheet_label=f"'{sheet}' ({done_units + 1}/{total_units})")
            done_units += 1
    if progress_cb:
        progress_cb(96, "Saving workbook to file…")
    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    data = buffer.getvalue()
    if progress_cb:
        progress_cb(100, "✅ Ready!")
    return data


# ══════════════════════════════════════════════════════════════════════
# PROGRESS-BAR DOWNLOAD HELPER  (used for every download in this app)
# ══════════════════════════════════════════════════════════════════════
def prepare_and_download(build_fn, label: str, filename: str, key: str, data_version,
                          mime: str = XLSX_MIME, help_text: str = None):
    """Renders a 'Prepare' button; on click, runs build_fn(progress_cb)
    while updating a live percentage progress bar, then shows a normal
    st.download_button once the bytes are ready. Result is cached in
    session_state keyed by data_version so switching tabs / re-running
    the script doesn't rebuild unless the underlying data/options changed.

    NOTE: this progress bar reflects SERVER-SIDE file generation. Once
    the bytes are ready, the browser download itself is instant (Streamlit
    has no hook into actual browser download-transfer progress)."""
    ready = (st.session_state.get(f"{key}_ready")
             and st.session_state.get(f"{key}_version") == data_version)

    if st.button(f"🛠️ Prepare: {label}", key=f"{key}_btn", use_container_width=True, help=help_text):
        # Two visible indicators: the bar itself (with % baked into its own
        # text) PLUS a large, unmissable percentage readout right above it —
        # so the number is always visible regardless of how the bar's own
        # text renders. Previously the bar's text showed the % ONLY when no
        # message was supplied; once real messages started arriving ("row
        # 150,000/300,000") the % silently disappeared from view. Fixed by
        # always prefixing the % onto the bar text too.
        pct_display = st.empty()
        progress = st.progress(0, text="0% — Starting…")
        pct_display.markdown("**0%** — Starting…")

        def cb(pct, msg=None):
            pct_int = min(max(int(round(pct)), 0), 100)
            full_text = f"{pct_int}% — {msg}" if msg else f"{pct_int}% complete…"
            progress.progress(pct_int, text=full_text)
            pct_display.markdown(f"**{pct_int}%** — {msg or 'Working…'}")

        try:
            data = build_fn(cb)
        except Exception as e:
            progress.empty()
            pct_display.empty()
            st.error(f"❌ Failed to prepare file: {e}")
            return

        progress.progress(100, text="100% — ✅ Done!")
        pct_display.markdown("**100%** — ✅ Done — ready to download!")
        st.session_state[f"{key}_bytes"] = data
        st.session_state[f"{key}_version"] = data_version
        st.session_state[f"{key}_ready"] = True
        ready = True

    if ready:
        st.download_button(
            f"⬇️ Download {filename}",
            data=st.session_state[f"{key}_bytes"],
            file_name=filename,
            mime=mime,
            use_container_width=True,
            key=f"{key}_dl",
        )
    else:
        st.caption("Click *Prepare* above to build the file, then download it.")

def quick_table_builder(df: pd.DataFrame):
    """Wraps excel_bytes_simple so small aux-table downloads report REAL
    per-batch progress (not just two fixed placeholder points)."""
    def _build(cb):
        return excel_bytes_simple(df, progress_cb=cb)
    return _build


@contextmanager
def processing_status(initial_label: str):
    status = st.status(initial_label, expanded=False, state="running")
    def _step(label, state="running"):
        status.update(label=label, state=state)
    try:
        yield _step
    except Exception:
        status.update(label="❌ Something went wrong", state="error")
        raise
    else:
        status.update(state="complete")


# ══════════════════════════════════════════════════════════════════════
# PAGE RENDER
# ══════════════════════════════════════════════════════════════════════
def render_portal_merge_final():
    st.markdown(f'<div class="app-title">📊 {APP_TITLE}</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="app-sub">Multi-tab-per-shift reading · Roll No & Centre Code padding · '
        'tab-level + shift-level cross-checks · no-limit multi-tab downloads with live progress</div>',
        unsafe_allow_html=True,
    )

    # ── 1. Upload ──────────────────────────────────────────────────────
    st.markdown('<div class="sec-title">1 · Upload Portal Workbooks</div>', unsafe_allow_html=True)
    c1, c2 = st.columns([3, 1])
    with c1:
        uploaded_files = st.file_uploader(
            "Drag & drop .xlsx workbooks — works with EITHER one-file-per-shift "
            "OR one-file-with-multiple-shift-tabs exports",
            type=["xlsx"], accept_multiple_files=True, key="pmf_files",
        )
    with c2:
        skip_rows = st.number_input("Header rows to skip (Note: It will remove from all files and tabs)", min_value=0, max_value=20, value=5, step=1, key="pmf_skip")

    # ── 2. Formatting (Roll No AND Centre Code padding) ─────────────────
    st.markdown('<div class="sec-title">2 · Roll Number & Centre Code Formatting</div>', unsafe_allow_html=True)
    f1, f2 = st.columns(2)
    with f1:
        st.markdown("**Roll No**")
        roll_pad_choice = st.radio("Pad Roll No with leading zeros?", ["No", "Yes"], horizontal=True, index=0, key="pmf_roll_pad")
        roll_digits = st.number_input("Roll No digit length", min_value=1, max_value=20, value=7, step=1, key="pmf_roll_digits") if roll_pad_choice == "Yes" else 7
        if roll_pad_choice == "No":
            st.caption("Roll No kept exactly as read from source (as text).")
    with f2:
        st.markdown("**Centre Code**")
        centre_pad_choice = st.radio("Pad Centre Code with leading zeros?", ["No", "Yes"], horizontal=True, index=0, key="pmf_centre_pad")
        centre_digits = st.number_input("Centre Code digit length", min_value=1, max_value=20, value=7, step=1, key="pmf_centre_digits") if centre_pad_choice == "Yes" else 7
        if centre_pad_choice == "No":
            st.caption("Centre Code kept exactly as read from source (as text).")

    if not uploaded_files:
        st.info("Upload one or more .xlsx workbooks to begin.")
        return

    st.caption(read_engine_info())

    files_payload = [(f.name, f.getvalue()) for f in uploaded_files]

    # ── Read every sheet in every file ──────────────────────────────────
    with st.spinner("📂 Reading every workbook and every tab…"):
        items = load_all_workbooks(tuple(files_payload), skip_rows)

    tab_audit_df = build_tab_audit(items)
    error_rows = tab_audit_df[tab_audit_df["Status"] == "ERROR"]
    empty_rows = tab_audit_df[tab_audit_df["Status"] == "EMPTY"]

    # ── 3. Tab-level cross check ────────────────────────────────────────
    st.markdown('<div class="sec-title">3 · Tab / Sheet Level Cross Check</div>', unsafe_allow_html=True)
    ok_count = int((tab_audit_df["Status"] == "OK").sum())
    st.write(
        f"**{len(uploaded_files):,} workbook(s)** · **{len(tab_audit_df):,} tab(s) found** · "
        f"**{ok_count:,} tab(s) read OK** · **{int(tab_audit_df['Total Count'].sum()):,} source rows**"
    )
    st.dataframe(tab_audit_df, use_container_width=True, hide_index=True)
    prepare_and_download(
        quick_table_builder(tab_audit_df), "Tab Audit", "tab_level_cross_check.xlsx", "pmf_tabaudit",
        data_version=("tabaudit", len(tab_audit_df)),
    )
    if not error_rows.empty:
        st.error(f"⚠️ {len(error_rows):,} tab(s) could not be read — see Status=ERROR / Note column above. "
                 "Everything else was still processed normally.")
        with st.expander("Show error detail"):
            st.dataframe(error_rows, use_container_width=True, hide_index=True)
    if not empty_rows.empty:
        st.caption(f"ℹ️ {len(empty_rows):,} tab(s) were completely empty and were skipped.")

    # ── Build the merged dataframe ──────────────────────────────────────
    frames = [it["df"] for it in items if it["df"] is not None]
    if not frames:
        st.error("No readable sheets found in any uploaded file.")
        return

    with processing_status("🧮 Merging and calculating…") as step:
        df = pd.concat(frames, ignore_index=True, sort=False)
        source_total = len(df)
        step("🧮 Applying attendance rules…")
        df = apply_attendance(df)
        step("🔢 Formatting Roll No…")
        df = apply_padding(df, "roll no", roll_pad_choice == "Yes", int(roll_digits))
        step("🔢 Formatting Centre Code…")
        df = apply_padding(df, "centre code", centre_pad_choice == "Yes", int(centre_digits))
        step("🔑 Building Unique Key…")
        df = add_unique_key(df)
        step("♻️ Flagging duplicates…")
        df = flag_duplicates(df)
        step("↕️ Sorting by FR Score (highest first)…")
        df = sort_by_fr_score(df)
        step("✅ Ready")

    st.markdown(
        f'<div class="note-box">✅ Processing complete — Source rows: <b>{source_total:,}</b> · '
        f'Final rows: <b>{len(df):,}</b> · Difference: <b>{len(df) - source_total:,}</b> '
        f'(should be 0 — a non-zero difference here would mean rows were gained/lost during merge)</div>',
        unsafe_allow_html=True,
    )

    # ── 4. Overall KPI summary ──────────────────────────────────────────
    st.markdown('<div class="sec-title">4 · Overall Summary</div>', unsafe_allow_html=True)
    total_rows = len(df)
    unique_keys = unique_roll_shift_count(df)
    dup_count = int((df.get("Duplicate Flag", pd.Series(dtype=str)) == "Yes").sum())
    centers = df["centre code"].replace("", np.nan).nunique() if "centre code" in df.columns else 0
    present = int((df.get("Attendance", pd.Series(dtype=str)) == "Present").sum())
    absent = int((df.get("Attendance", pd.Series(dtype=str)) == "Absent").sum())
    logical_shifts = df["shift"].replace("", np.nan).nunique() if "shift" in df.columns else 0
    exam_codes = df["EXAM_CODE"].replace("", np.nan).nunique() if "EXAM_CODE" in df.columns else 0
    raw_tabs = ok_count

    render_kpi_cards([
        {"icon": "🧾", "label": "Total Rows", "value": f"{total_rows:,}", "color": "blue"},
        {"icon": "🆔", "label": "Unique Roll No + Shift", "value": f"{unique_keys:,}", "color": "cyan"},
        {"icon": "♻️", "label": "Duplicate Rows", "value": f"{dup_count:,}", "color": "red"},
        {"icon": "🏢", "label": "Centers", "value": f"{centers:,}", "color": "purple"},
        {"icon": "✅", "label": "Present", "value": f"{present:,}", "color": "green"},
        {"icon": "❌", "label": "Absent", "value": f"{absent:,}", "color": "amber"},
        {"icon": "🕐", "label": "Logical Shifts", "value": f"{logical_shifts:,}", "color": "slate"},
        {"icon": "📁", "label": "Exam Codes", "value": f"{exam_codes:,}", "color": "pink"},
        {"icon": "📑", "label": "Raw Tabs Read OK", "value": f"{raw_tabs:,}", "color": "slate"},
    ])

    # ── 5. Final shift-level cross check ────────────────────────────────
    st.markdown('<div class="sec-title">5 · Final Shift-wise Cross Check (After Merge, Before Download)</div>', unsafe_allow_html=True)
    st.caption(
        "For every logical shift: how many raw tabs fed into it (and their names), source-tab "
        "total vs. final total after merge, unique/duplicate counts, and a MATCH/CHECK flag. "
        "Use this to confirm shift(2)/shift(3) continuation tabs consolidated correctly."
    )
    shift_audit_df = build_final_shift_audit(df, tab_audit_df)
    st.dataframe(shift_audit_df, use_container_width=True, hide_index=True)
    prepare_and_download(
        quick_table_builder(shift_audit_df), "Final Shift Cross Check", "final_shift_cross_check.xlsx", "pmf_shiftaudit",
        data_version=("shiftaudit", len(shift_audit_df), total_rows),
    )
    if not shift_audit_df.empty and (shift_audit_df["Cross Check"] == "✅ MATCH").all():
        st.success("✅ All logical shifts match their source tab totals exactly.")
    elif not shift_audit_df.empty:
        st.warning("⚠️ One or more shifts show a CHECK status above — a rows-difference other than 0 usually means "
                   "a tab failed to read (see Section 3) or a shift name didn't consolidate as expected.")

    # ── 6. Shift-wise summary ───────────────────────────────────────────
    shift_summary_df = build_shift_summary(df)
    if not shift_summary_df.empty:
        st.markdown('<div class="sec-title">6 · Shift-wise Summary</div>', unsafe_allow_html=True)
        st.dataframe(shift_summary_df, use_container_width=True, hide_index=True)
        prepare_and_download(
            quick_table_builder(shift_summary_df), "Shift Summary", "shift_wise_summary.xlsx", "pmf_shiftsum",
            data_version=("shiftsum", len(shift_summary_df), total_rows),
        )

    # ── 7. Photo Matched breakdown ──────────────────────────────────────
    pm_df = build_photo_matched_breakdown(df)
    if not pm_df.empty:
        st.markdown('<div class="sec-title">7 · Photo Matched Breakdown</div>', unsafe_allow_html=True)
        st.dataframe(pm_df, use_container_width=True, hide_index=True)
        prepare_and_download(
            quick_table_builder(pm_df), "Photo Matched Breakdown", "photo_matched_breakdown.xlsx", "pmf_pm",
            data_version=("pm", len(pm_df), total_rows),
        )

    # ── 8. Centre Code × Shift pivot ────────────────────────────────────
    pivot_df = build_center_shift_pivot(df)
    if not pivot_df.empty:
        st.markdown('<div class="sec-title">8 · Centre Code × Shift — Present Count</div>', unsafe_allow_html=True)
        st.dataframe(pivot_df, use_container_width=True, hide_index=True)
        prepare_and_download(
            quick_table_builder(pivot_df), "Centre × Shift Pivot", "centre_shift_summary.xlsx", "pmf_pivot",
            data_version=("pivot", len(pivot_df), total_rows),
        )

    # ── 9. Duplicate / Scribers ─────────────────────────────────────────
    st.markdown('<div class="sec-title">9 · Duplicate / Scribers Check</div>', unsafe_allow_html=True)
    dup_ranked, scribers_df = build_scribers_data(df)
    if not dup_ranked.empty:
        st.markdown(f"**⚠️ Duplicate Roll No + Shift groups — {len(dup_ranked):,} row(s), with rank**")
        st.caption("rank 1 = highest FR Score in that Unique Key group (the row that would be kept/merged).")
        st.dataframe(dup_ranked, use_container_width=True, hide_index=True)
        prepare_and_download(
            quick_table_builder(dup_ranked), "Duplicate Rows (ranked)", "duplicate_roll_no_rows.xlsx", "pmf_dup",
            data_version=("dup", len(dup_ranked), total_rows),
        )
    else:
        st.success("✅ 0 duplicate Roll No + Shift groups found.")

    if not scribers_df.empty:
        st.markdown(f"**📝 Scribers Data — {len(scribers_df):,} row(s) (rank > 1)**")
        st.dataframe(scribers_df, use_container_width=True, hide_index=True)
        prepare_and_download(
            quick_table_builder(scribers_df), "Scribers Data", "scribers_data.xlsx", "pmf_scribers",
            data_version=("scribers", len(scribers_df), total_rows),
        )
    else:
        st.success("✅ 0 Scribers rows.")

    # ── 10. Search ───────────────────────────────────────────────────────
    st.markdown('<div class="sec-title">10 · Search by Roll No</div>', unsafe_allow_html=True)
    query = st.text_input("🔍 Search Roll No", key="pmf_search", placeholder="Type a Roll No and press Enter…")
    if query and query.strip() and "roll no" in df.columns:
        matches = df[df["roll no"].astype(str).str.contains(query.strip(), case=False, na=False, regex=False)]
        if matches.empty:
            st.warning(f"No rows found for Roll No containing '{query.strip()}'.")
        else:
            st.markdown(f"**Found {len(matches):,} row(s)**")
            st.dataframe(matches, use_container_width=True, hide_index=True)
            prepare_and_download(
                quick_table_builder(matches), "Search Results", f"search_{query.strip()}.xlsx", "pmf_searchdl",
                data_version=("search", query.strip(), len(matches)),
            )

    # ── 11. Preview ──────────────────────────────────────────────────────
    st.markdown('<div class="sec-title">11 · Data Preview</div>', unsafe_allow_html=True)
    with st.container(border=True):
        st.markdown("**First 100 rows (highest FR Score first)**")
        st.dataframe(df.head(100), use_container_width=True, hide_index=True)
        prepare_and_download(
            quick_table_builder(df.head(100)), "First 100 rows", "preview_first_100.xlsx", "pmf_top100",
            data_version=("top100", total_rows),
        )
    with st.container(border=True):
        st.markdown("**Last 100 rows (lowest FR Score)**")
        st.dataframe(df.tail(100), use_container_width=True, hide_index=True)
        prepare_and_download(
            quick_table_builder(df.tail(100)), "Last 100 rows", "preview_last_100.xlsx", "pmf_bottom100",
            data_version=("bottom100", total_rows),
        )

    # ── 12. Download column selection ───────────────────────────────────
    st.markdown('<div class="sec-title">12 · Download Column Selection</div>', unsafe_allow_html=True)
    default_cols_resolved = resolve_default_columns(df)
    default_labels_present = [EXPORT_DISPLAY_NAMES.get(c, c) for c in default_cols_resolved]
    st.caption(
        "Always included by default: **" + ", ".join(default_labels_present) + "** — "
        "plus every calculated column (Attendance, Unique Key, Duplicate Flag, Shift, "
        "EXAM_CODE, Source File, Source Sheet) automatically."
    )
    optional_cols = [c for c in df.columns if c not in default_cols_resolved and c not in SYSTEM_COLUMNS]
    extra_selected = st.multiselect(
        "Need any other column too? Pick it here — it'll be appended to the download.",
        optional_cols, key="pmf_extra_cols",
    )
    export_cols_final = build_export_columns(df, extra_selected)
    st.caption(f"→ {len(export_cols_final)} column(s) will be included in every download below.")

    # ── 13. Download — Whole Data (single sheet) & Shift-wise (per-shift tabs) ─
    st.markdown('<div class="sec-title">13 · Download</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="note-box">📌 <b>Whole Data</b> = every shift mixed together in ONE sheet '
        '(e.g. <code>05 Jul\'26 S1</code> and <code>05 Jul\'26 S2</code> rows sit in the same tab) — '
        'it only spills onto <code>All Data (2)</code>, <code>All Data (3)</code>… if the TOTAL row '
        'count crosses the cap below.<br>📌 <b>Shift-wise</b> = pick a shift and it gets its OWN tab(s) — '
        'this is the one to use when a single shift alone has 12,00,000+ rows: it continues onto '
        '<code>ShiftName (2)</code>, <code>ShiftName (3)</code>… within one workbook.</div>',
        unsafe_allow_html=True,
    )

    max_rows_per_tab = st.number_input(
        "Max rows per tab (Whole Data: splits the combined sheet by this many rows · "
        "Shift-wise: a shift beyond this continues onto ShiftName (2), ShiftName (3), ...)",
        min_value=10_000, max_value=EXCEL_HARD_ROW_LIMIT - 76, value=DEFAULT_MAX_ROWS_PER_TAB,
        step=10_000, key="pmf_maxrows",
    )

    # Clean export dataframe — 'Shift' appears exactly once (properly
    # labeled), no hidden duplicate 'shift' column in the actual sheet.
    export_df_display = prepare_export_dataframe(df, extra_selected)
    # Raw shift values, position-aligned with export_df_display, kept OUT
    # of the exported columns — used only internally to group rows for
    # the Shift-wise download.
    shift_group_key = df["shift"].reset_index(drop=True)

    exam_codes_present = sorted(c for c in df.get("EXAM_CODE", pd.Series(dtype=str)).dropna().unique() if c and c != "Unknown")
    file_prefix = "_".join(exam_codes_present) if exam_codes_present else "portal_merged"

    d1, d2 = st.columns(2)
    with d1:
        st.markdown("**Whole Data — one sheet, all shifts combined**")
        st.caption(f"{total_rows:,} row(s) across {logical_shifts:,} shift(s) — all in one continuous sheet.")
        whole_filename = f"{file_prefix}_combined_all_data.xlsx"
        prepare_and_download(
            lambda cb: build_single_sheet_workbook(export_df_display, int(max_rows_per_tab), cb, base_sheet_name="All Data"),
            "Whole Data Workbook", whole_filename, "pmf_whole",
            data_version=("whole", total_rows, tuple(export_cols_final), max_rows_per_tab,
                          roll_pad_choice, roll_digits, centre_pad_choice, centre_digits),
            help_text="All rows, all shifts, one continuous sheet — splits only on total row-count overflow.",
        )

    with d2:
        st.markdown("**Shift-wise — pick one or more shifts, each its own tab(s)**")
        shift_options = sorted([s for s in df["shift"].dropna().unique().tolist() if s]) if "shift" in df.columns else []
        selected_shifts = st.multiselect("Choose shift(s) to download", shift_options, key="pmf_shift_select")
        if selected_shifts:
            sel_row_count = int(df[df["shift"].isin(selected_shifts)].shape[0])
            n_tabs_est = sum(max(1, math.ceil(len(df[df["shift"] == s]) / max_rows_per_tab)) for s in selected_shifts)
            st.caption(f"{sel_row_count:,} row(s) across {len(selected_shifts)} shift(s) → ~{n_tabs_est} tab(s) total.")
            safe_name = re.sub(r'[\\/*?:"<>|]', "_", "_".join(selected_shifts))[:80]
            prepare_and_download(
                lambda cb: build_grouped_workbook(export_df_display, shift_group_key, selected_shifts, int(max_rows_per_tab), cb),
                "Selected Shift(s) Workbook", f"{safe_name}_multitab.xlsx", "pmf_shiftdl",
                data_version=("shiftdl", tuple(selected_shifts), sel_row_count, tuple(export_cols_final),
                              max_rows_per_tab, roll_pad_choice, roll_digits, centre_pad_choice, centre_digits),
                help_text="Each selected shift gets its own tab, auto-continuing if it exceeds the row cap.",
            )
        else:
            st.caption("Pick at least one shift above to enable this download.")

    # ── Suggestions ──────────────────────────────────────────────────────
    with st.expander("💡 Suggestions before rolling this into module1_portal.py"):
        st.markdown(
            """
- **Promote shared pieces into `utils.py` first.** Centre Code padding, the
  tab-level audit, the final shift-level cross-check, `make_sheet_name`, and
  `build_grouped_workbook` are all generic enough to be used by every module
  (1, 2, 4, 5), not just Portal. Moving them into `utils.py` once (instead of
  copy-pasting into `module1_portal.py`) means one fix applies everywhere,
  same as the rest of the app already does.
- **Make the row-cap-per-tab a shared constant** (e.g. `utils.MAX_ROWS_PER_TAB`)
  so every module's shift-wise/whole-data export uses the same default and you
  only tune it in one place if Excel's limit or your infra changes.
- **Progress bar caveat:** the bars here reflect *server-side* file generation
  only — Streamlit can't hook into the browser's actual download-transfer
  progress, since the file is already fully built in memory before the
  download button appears. That said, for 25L+ row exports the generation
  step itself can take real time, so the bar is still meaningful.
- **Sample fixture files** — before wiring this into production, it's worth
  keeping 2–3 small sample workbooks (one old one-file-per-shift, one new
  multi-tab, one with a deliberately corrupted sheet) as a quick regression
  check whenever this logic changes.
- **Write speed, resolved:** benchmarked 4 approaches on 300,000 rows —
  `xlsxwriter`'s `constant_memory` option was ~2x faster than plain
  pandas+xlsxwriter, but proven (via round-trip test) to silently corrupt
  row 0 of every sheet when combined with pandas' `to_excel()`. **openpyxl's
  write-only mode with `ws.append()`** (used everywhere in this file now)
  came out fastest of all AND correct, with no xlsxwriter dependency needed.
- **Please also check `utils.py`'s `df_to_excel_bytes`** (used across every
  other module) and `module_portal_merge_test.py`'s `build_shiftwise_multitab_excel`
  — both still enable `xlsxwriter`'s `constant_memory` option above ~100,000
  rows and are likely hitting the identical row-0 corruption bug on any large
  export today. Worth porting the openpyxl write-only approach from this file
  into `utils.py` before anything else rolls into production, since the bug
  fails silently (dropped data) rather than raising an error.
- **Column aliasing:** if new source files start naming Device / BioDevice /
  Bio Count / Operator differently, add the new spelling to `ALIASES` here
  (and to `utils.ALIAS_MAP` once merged) rather than special-casing it in the
  page code.
            """
        )


if __name__ == "__main__":
    _configure_page()
    render_portal_merge_final()
