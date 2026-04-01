"""
Equipment Accountability Form Generator
BMG Outsourcing — redesigned UI with per-monitor cable assignment
Supports both Work From Home and Work On Site templates
"""

import io
import copy
import base64
import re
import unicodedata
import zipfile
from datetime import datetime
from pathlib import Path
from lxml import etree

import pandas as pd
import streamlit as st

# ─────────────────────────────────────────────
# PATHS  —  checked in priority order
# ─────────────────────────────────────────────

WFH_TEMPLATE_NAME    = "Equipment Accountability Form (Work From Home).dotx"
ONSITE_TEMPLATE_NAME = "Equipment Accountability Form (Work On Site).dotx"

def _find_template(filename: str) -> Path | None:
    candidates = [
        Path(__file__).parent / "src" / filename,
        Path.cwd() / "src" / filename,
        Path(__file__).parent / filename,
        Path.cwd() / filename,
    ]
    for p in candidates:
        if p.exists():
            return p
    return None

def _find_logo() -> Path | None:
    candidates = [
        Path(__file__).parent / "images" / "logo.png",
        Path.cwd() / "images" / "logo.png",
        Path(__file__).parent / "logo.png",
        Path.cwd() / "logo.png",
    ]
    for p in candidates:
        if p.exists():
            return p
    return None

LOGO_PATH = _find_logo()

# ─────────────────────────────────────────────
# PAGE CONFIG
# ─────────────────────────────────────────────

st.set_page_config(
    page_title="Asset Accountability System",
    layout="centered",
)

# ─────────────────────────────────────────────
# LOGO HELPER
# ─────────────────────────────────────────────

def get_logo_b64() -> str:
    if LOGO_PATH and LOGO_PATH.exists():
        return base64.b64encode(LOGO_PATH.read_bytes()).decode()
    return ""

LOGO_B64 = get_logo_b64()
LOGO_HTML = (
    f'<img src="data:image/png;base64,{LOGO_B64}" style="height:72px;object-fit:contain;">'
    if LOGO_B64 else
    '<span style="font-size:1.5rem;font-weight:900;color:#fff;letter-spacing:-1px;">BMG</span>'
)

# ─────────────────────────────────────────────
# STYLES
# ─────────────────────────────────────────────

st.markdown(f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:ital,opsz,wght@0,9..40,300;0,9..40,400;0,9..40,500;0,9..40,600;0,9..40,700;1,9..40,400&display=swap');

:root {{
    --navy:       #0d2545;
    --blue:       #1565c0;
    --blue-lt:    #1e88e5;
    --blue-pale:  #e8f0fc;
    --orange:     #e65c00;
    --orange-lt:  #fff3ec;
    --green:      #43a832;
    --green-lt:   #edf7ea;
    --teal:       #00796b;
    --teal-lt:    #e0f2f1;
    --teal-pale:  #b2dfdb;
    --bg:         #f0f4fa;
    --card:       #ffffff;
    --border:     #d0dce8;
    --text:       #0d2545;
    --muted:      #5a6e8a;
    --radius:     10px;
    --shadow:     0 2px 12px rgba(13,37,69,.09);
    --shadow-md:  0 4px 24px rgba(13,37,69,.14);
}}

html, body, [class*="css"] {{
    font-family: 'DM Sans', sans-serif !important;
    background: var(--bg) !important;
    color: var(--text) !important;
}}

[data-testid="stSidebar"] {{ display: none !important; }}
header[data-testid="stHeader"] {{ display: none !important; }}
[data-testid="stDecoration"] {{ display: none !important; }}

.main .block-container {{
    max-width: 860px !important;
    padding: 2rem 1.5rem 3rem !important;
}}

/* ── Header ── */
.bmg-header {{
    background: linear-gradient(135deg, #0d2545 0%, #1565c0 100%);
    border-radius: 0 0 14px 14px;
    padding: 2rem 2rem;
    display: flex;
    align-items: center;
    justify-content: space-between;
    margin: -2rem -1.5rem 2.5rem -1.5rem;
    box-shadow: var(--shadow-md);
}}
.bmg-header-title {{
    color: #fff;
    font-size: 1.1rem;
    font-weight: 700;
    letter-spacing: -0.01em;
}}
.bmg-header-sub {{
    color: rgba(255,255,255,.6);
    font-size: 0.76rem;
    font-weight: 400;
    margin-top: 2px;
}}

/* ── Form type selector ── */
.form-type-selector {{
    display: flex;
    gap: 1rem;
    margin-bottom: 0.6rem;
}}
.form-type-card {{
    flex: 1;
    border: 2px solid var(--border);
    border-radius: var(--radius);
    padding: 1.1rem 1.3rem;
    background: var(--card);
    box-shadow: var(--shadow);
    text-align: left;
}}
.form-type-card.active-wfh {{
    border-color: var(--blue);
    background: var(--blue-pale);
    box-shadow: 0 0 0 3px rgba(21,101,192,.12), var(--shadow);
}}
.form-type-card.active-onsite {{
    border-color: var(--teal);
    background: var(--teal-lt);
    box-shadow: 0 0 0 3px rgba(0,121,107,.12), var(--shadow);
}}
.form-type-icon {{
    font-size: 1.6rem;
    margin-bottom: 0.3rem;
    display: block;
}}
.form-type-label {{
    font-size: 0.92rem;
    font-weight: 700;
    color: var(--navy);
    margin-bottom: 2px;
}}
.form-type-desc {{
    font-size: 0.75rem;
    color: var(--muted);
    line-height: 1.45;
}}
.form-type-badge {{
    display: inline-block;
    font-size: 0.65rem;
    font-weight: 700;
    border-radius: 4px;
    padding: 2px 7px;
    margin-top: 5px;
    letter-spacing: 0.06em;
    text-transform: uppercase;
}}
.badge-wfh    {{ background: var(--blue); color: #fff; }}
.badge-onsite {{ background: var(--teal); color: #fff; }}

/* ── Step card ── */
.step-card {{
    background: var(--card);
    border: 1px solid var(--border);
    border-radius: var(--radius);
    padding: 1.4rem 1.6rem 1.1rem;
    margin-bottom: 1rem;
    box-shadow: var(--shadow);
}}
.step-label {{
    display: flex;
    align-items: center;
    gap: 0.65rem;
    margin-bottom: 0.6rem;
}}
.step-badge {{
    background: linear-gradient(135deg, #1565c0, #1e88e5);
    color: #fff;
    font-size: 0.68rem;
    font-weight: 700;
    width: 22px;
    height: 22px;
    border-radius: 50%;
    display: inline-flex;
    align-items: center;
    justify-content: center;
    flex-shrink: 0;
}}
.step-title {{
    font-size: 0.93rem;
    font-weight: 700;
    color: var(--navy);
    letter-spacing: 0.005em;
}}
.step-desc {{
    font-size: 0.8rem;
    color: var(--muted);
    margin-bottom: 0.9rem;
    line-height: 1.55;
}}

/* ── Prepared by card ── */
.preparedby-card {{
    background: linear-gradient(135deg, #f0f7ff 0%, #e8f0fc 100%);
    border: 1.5px solid #b3c8f0;
    border-radius: var(--radius);
    padding: 1.2rem 1.6rem 1rem;
    margin-bottom: 1.2rem;
    box-shadow: var(--shadow);
    display: flex;
    align-items: center;
    gap: 1rem;
}}
.preparedby-icon {{ font-size: 1.6rem; flex-shrink: 0; }}
.preparedby-label {{
    font-size: 0.72rem;
    font-weight: 700;
    color: var(--blue);
    text-transform: uppercase;
    letter-spacing: 0.07em;
    margin-bottom: 2px;
}}
.preparedby-hint {{
    font-size: 0.76rem;
    color: var(--muted);
    line-height: 1.4;
}}

/* ── Inputs ── */
.stTextInput input,
.stSelectbox > div > div,
.stDateInput input {{
    border: 1.5px solid var(--border) !important;
    border-radius: 8px !important;
    box-shadow: none !important;
    background: #fff !important;
    color: var(--text) !important;
    font-family: 'DM Sans', sans-serif !important;
    font-size: 0.87rem !important;
}}
.stTextInput input:focus {{
    border-color: var(--blue) !important;
    box-shadow: 0 0 0 3px rgba(21,101,192,.12) !important;
}}

/* ── Labels ── */
label,
.stTextInput label,
.stDateInput label,
.stSelectbox label {{
    font-family: 'DM Sans', sans-serif !important;
    font-size: 0.72rem !important;
    font-weight: 600 !important;
    color: var(--muted) !important;
    text-transform: uppercase !important;
    letter-spacing: 0.06em !important;
}}

/* ── Primary button ── */
.stButton > button {{
    background: linear-gradient(135deg, #1565c0, #1e88e5) !important;
    color: #fff !important;
    border: none !important;
    border-radius: 8px !important;
    font-family: 'DM Sans', sans-serif !important;
    font-size: 0.87rem !important;
    font-weight: 600 !important;
    letter-spacing: 0.01em !important;
    padding: 0.55rem 1.4rem !important;
    box-shadow: 0 2px 8px rgba(21,101,192,.3) !important;
    transition: all .15s !important;
}}
.stButton > button:hover {{
    background: linear-gradient(135deg, #0d47a1, #1565c0) !important;
    box-shadow: 0 4px 16px rgba(21,101,192,.4) !important;
    transform: translateY(-1px) !important;
}}

/* ── Small outline buttons ── */
.small-btn .stButton > button {{
    background: #fff !important;
    color: var(--blue) !important;
    border: 1.5px solid var(--blue) !important;
    font-size: 0.76rem !important;
    padding: 0.28rem 0.85rem !important;
    box-shadow: none !important;
    font-weight: 600 !important;
    letter-spacing: 0.02em !important;
}}
.small-btn .stButton > button:hover {{
    background: var(--blue-pale) !important;
    transform: none !important;
    box-shadow: none !important;
}}

/* ── Download button ── */
[data-testid="stDownloadButton"] button {{
    background: linear-gradient(135deg, #e65c00, #ff8f00) !important;
    color: #fff !important;
    border: none !important;
    border-radius: 8px !important;
    font-family: 'DM Sans', sans-serif !important;
    font-size: 0.87rem !important;
    font-weight: 600 !important;
    letter-spacing: 0.01em !important;
    box-shadow: 0 2px 8px rgba(230,92,0,.35) !important;
    transition: all .15s !important;
}}
[data-testid="stDownloadButton"] button:hover {{
    background: linear-gradient(135deg, #bf360c, #e65c00) !important;
    box-shadow: 0 4px 16px rgba(230,92,0,.45) !important;
    transform: translateY(-1px) !important;
}}

/* ── Checkboxes ── */
.stCheckbox {{
    background: #f7f9fd;
    border: 1px solid var(--border);
    border-radius: 8px;
    padding: 0.5rem 0.8rem;
    margin-bottom: 0.3rem !important;
    transition: background .12s;
}}
.stCheckbox:hover {{ background: var(--blue-pale); border-color: #b3c8f0; }}
.stCheckbox label {{
    font-size: 0.83rem !important;
    font-weight: 500 !important;
    text-transform: none !important;
    letter-spacing: 0 !important;
    color: var(--text) !important;
}}

/* ── Monitor cable group ── */
.monitor-group-header {{
    background: var(--blue-pale);
    border: 1px solid #b3c8f0;
    border-radius: 8px 8px 0 0;
    padding: 0.5rem 0.9rem;
    font-size: 0.78rem;
    font-weight: 700;
    color: var(--navy);
    display: flex;
    align-items: center;
    gap: 0.5rem;
    margin-top: 0.6rem;
    margin-bottom: 0;
}}
.monitor-group-body {{
    background: #f7faff;
    border: 1px solid #b3c8f0;
    border-top: none;
    border-radius: 0 0 8px 8px;
    padding: 0.55rem 0.9rem 0.4rem;
    margin-bottom: 0.5rem;
}}

/* ── Sequence badge ── */
.seq-badge {{
    display: inline-block;
    background: var(--blue-pale);
    color: var(--blue);
    font-size: 0.68rem;
    font-weight: 700;
    border-radius: 4px;
    padding: 1px 6px;
    margin-right: 6px;
    vertical-align: middle;
}}

/* ── Alerts ── */
.stAlert {{
    border-radius: 8px !important;
    font-family: 'DM Sans', sans-serif !important;
    font-size: 0.83rem !important;
}}

/* ── Metrics ── */
[data-testid="stMetric"] {{
    background: var(--card);
    border: 1px solid var(--border);
    border-radius: var(--radius);
    padding: 1rem 1.2rem;
    box-shadow: var(--shadow);
}}
[data-testid="stMetricValue"] {{
    font-family: 'DM Sans', sans-serif !important;
    font-weight: 700 !important;
    font-size: 1.7rem !important;
    color: var(--navy) !important;
}}
[data-testid="stMetricLabel"] {{
    font-family: 'DM Sans', sans-serif !important;
    font-size: 0.72rem !important;
    font-weight: 600 !important;
    text-transform: uppercase !important;
    letter-spacing: 0.06em !important;
    color: var(--muted) !important;
}}

/* ── Divider ── */
hr {{ border-color: var(--border) !important; margin: 1rem 0 !important; }}

/* ── Caption ── */
.stCaption, small, [data-testid="stCaptionContainer"] {{
    font-family: 'DM Sans', sans-serif !important;
    color: var(--muted) !important;
    font-size: 0.76rem !important;
}}

/* ── Search hint ── */
.search-hint {{
    font-size: 0.77rem;
    color: var(--blue);
    background: var(--blue-pale);
    border-left: 3px solid var(--blue);
    border-radius: 0 6px 6px 0;
    padding: 0.55rem 0.75rem;
    margin-bottom: 0.85rem;
    font-weight: 400;
    line-height: 1.6;
}}
.search-hint strong {{ font-weight: 700; color: var(--navy); }}

/* ── File uploader ── */
[data-testid="stFileUploader"] section {{
    border: 2px dashed var(--border) !important;
    border-radius: var(--radius) !important;
    background: #f7f9fd !important;
}}

/* ── Footer ── */
.bmg-footer {{
    text-align: center;
    color: var(--muted);
    font-size: 0.71rem;
    margin-top: 2rem;
    padding-top: 1rem;
    border-top: 1px solid var(--border);
}}

/* ── Preview table ── */
.preview-table {{
    width: 100%;
    border-collapse: collapse;
    font-size: 0.8rem;
    margin-top: 0.5rem;
}}
.preview-table th {{
    background: var(--navy);
    color: #fff;
    font-weight: 600;
    padding: 0.45rem 0.7rem;
    text-align: left;
    font-size: 0.72rem;
    letter-spacing: 0.04em;
    text-transform: uppercase;
}}
.preview-table td {{
    padding: 0.4rem 0.7rem;
    border-bottom: 1px solid var(--border);
    color: var(--text);
}}
.preview-table tr:nth-child(even) td {{ background: #f7f9fd; }}
.preview-table tr:hover td {{ background: var(--blue-pale); }}
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
# CONFIG
# ─────────────────────────────────────────────

CSV_COLUMNS = {
    "asset_tag":      "Asset Tag",
    "content_type":   "Content Type",
    "brand":          "Brand",
    "model":          "Model",
    "serial":         "Serial Code",
    "condition":      "Condition",
    "usage_status":   "Usage Status",
    "current_user":   "Current User",
    "previous_owner": "Previous Owner",
    "client":         "Client",
    "remarks":        "Remark(s)",
}

PREPARED_BY_OPTIONS = [
    "IT Intern",
    "Jiro Macabitas",
    "Angelo Forbes",
    "Bryan Odero",
]

ITEM_SEQUENCE_ORDER = {
    "laptop":                  1,
    "computer":                1,
    "desktop":                 1,
    "charger":                 2,
    "monitor":                 3,
    "hdmi":                    5,
    "keyboard":                6,
    "mouse":                   7,
    "headset":                 8,
    "headphone":               8,
    "usb":                     9,
    "usb peripheral":          9,
    "vga":                     10,
    "dvi":                     11,
    "displayport":             12,
    "type-c":                  13,
    "type c":                  13,
    "adapter":                 13,
    "converter":               13,
    "docking station":         14,
    "docking":                 14,
    "webcam":                  16,
    "speaker":                 17,
    "ethernet adapter":        18,
}

SIMPLE_PERIPHERALS = [
    ("Charger",         "Charger"),
    ("Keyboard",        "Keyboard"),
    ("Mouse",           "Mouse"),
    ("Headset",         "Headset"),
    ("Docking Station", "Docking Station"),
    ("Webcam",          "Webcam"),
    ("Speakers",        "Audio"),
    ("Ethernet Adapter","Network Adapter"),
]

MONITOR_PERIPHERALS = [
    ("Monitor Power Cable", "Monitor Power Cable", 3.5),
    ("HDMI Cable",          "HDMI Cable",          5.0),
    ("VGA Cable",           "VGA Cable",           10.0),
    ("DVI Cable",           "DVI Cable",           11.0),
    ("DisplayPort Cable",   "DisplayPort Cable",   12.0),
    ("Type-C Adapter",      "Type-C Adapter",      13.0),
]

_POSITION_SP_SUFFIXES = (":position", ":jobtitle", ":job title", ":title", ":designation")
_POSITION_KEYWORDS    = ("position", "job title", "jobtitle", "designation", "title")

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_FONT_NAME  = "Calibri"
_FONT_SIZE  = "20"
_TABLE_SIZE = "18"

# ─────────────────────────────────────────────
# DATA HELPERS
# ─────────────────────────────────────────────

def load_csv(file_obj):
    for enc in ("utf-8-sig", "latin-1"):
        try:
            file_obj.seek(0)
            return pd.read_csv(file_obj, encoding=enc), None
        except UnicodeDecodeError:
            continue
        except Exception as e:
            return pd.DataFrame(), str(e)
    return pd.DataFrame(), "Could not decode the file."

def safe_str(val) -> str:
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return ""
    if isinstance(val, (datetime, pd.Timestamp)):
        return val.strftime("%B %d, %Y")
    return str(val).strip()

def detect_columns(df: pd.DataFrame) -> dict:
    actual = [str(c) for c in df.columns]
    result = {}
    for key, expected in CSV_COLUMNS.items():
        if expected in actual:
            result[key] = expected
        else:
            matches = [c for c in actual if expected.lower() in c.lower()]
            result[key] = matches[0] if matches else None
    return result

def detect_position_column(df: pd.DataFrame) -> str | None:
    cols = [str(c) for c in df.columns]
    for col in cols:
        col_l = col.lower()
        for suffix in _POSITION_SP_SUFFIXES:
            if col_l.endswith(suffix):
                return col
    exact_names = ("position", "job title", "jobtitle", "designation", "title", "role")
    for col in cols:
        if col.lower() in exact_names:
            return col
    for col in cols:
        col_l = col.lower()
        for kw in ("position", "job title", "jobtitle", "designation"):
            if kw in col_l:
                return col
    return None

def get_position_value(row, df_columns, position_col: str | None) -> str:
    if position_col and position_col in df_columns:
        val = safe_str(row.get(position_col, ""))
        if val:
            return val
    for col in df_columns:
        col_lower = col.lower()
        if any(col_lower.endswith(suffix) for suffix in _POSITION_SP_SUFFIXES):
            val = safe_str(row.get(col, ""))
            if val:
                return val
        if any(kw in col_lower for kw in ("position", "job title", "jobtitle", "designation")):
            val = safe_str(row.get(col, ""))
            if val:
                return val
    return ""

# ─────────────────────────────────────────────
# SEQUENCE HELPERS
# ─────────────────────────────────────────────

def _get_sequence_key(equipment_text: str) -> float:
    text_lower = equipment_text.lower()
    if "monitor power" in text_lower or (
        "power cable" in text_lower and "monitor" in text_lower
    ):
        return 3.5
    sorted_keys = sorted(ITEM_SEQUENCE_ORDER.keys(), key=lambda k: -len(k))
    for keyword in sorted_keys:
        if keyword in text_lower:
            return float(ITEM_SEQUENCE_ORDER[keyword])
    return 99.0

def _is_monitor(equipment_text: str) -> bool:
    t = equipment_text.lower()
    return "monitor" in t and "power" not in t and "cable" not in t

def _build_equipment_label(brand: str, model: str, content: str, asset_tag: str) -> str:
    parts = [p for p in [brand, model, content] if p]
    return " ".join(parts) if parts else asset_tag

def sort_assets_by_sequence(
    assets_df: pd.DataFrame,
    col_map: dict,
    simple_extra_rows: list[dict],
    monitor_cable_assignments: list[dict],
) -> list[dict]:
    rows: list[dict] = []

    for _, row in assets_df.iterrows():
        c         = col_map
        content   = safe_str(row.get(c.get("content_type") or "", ""))
        brand     = safe_str(row.get(c.get("brand")        or "", ""))
        model     = safe_str(row.get(c.get("model")        or "", ""))
        serial    = safe_str(row.get(c.get("serial")       or "", ""))
        asset_tag = safe_str(row.get(c.get("asset_tag")    or "", ""))
        remarks   = safe_str(row.get(c.get("remarks")      or "", ""))
        equipment = _build_equipment_label(brand, model, content, asset_tag)
        seq_key   = _get_sequence_key(equipment)
        rows.append({
            "equipment":   equipment,
            "serial":      serial,
            "asset_tag":   asset_tag,
            "remarks":     remarks,
            "_seq_key":    seq_key,
            "_is_monitor": _is_monitor(equipment),
        })

    for extra in simple_extra_rows:
        periph_name = extra.get("name", "")
        equipment   = _build_equipment_label(
            extra.get("brand", ""), extra.get("model", ""),
            extra.get("type",  ""), periph_name,
        )
        rows.append({
            "equipment":   equipment,
            "serial":      extra.get("serial",    ""),
            "asset_tag":   extra.get("asset_tag", ""),
            "remarks":     extra.get("remarks",   ""),
            "_seq_key":    _get_sequence_key(periph_name),
            "_is_monitor": False,
        })

    rows.sort(key=lambda r: r["_seq_key"])

    from collections import defaultdict
    cables_by_monitor: dict[int, list[tuple[float, str]]] = defaultdict(list)
    for assignment in monitor_cable_assignments:
        cables_by_monitor[assignment["monitor_idx"]].append(
            (assignment["cable_seq"], assignment["cable_name"])
        )
    for idx in cables_by_monitor:
        cables_by_monitor[idx].sort(key=lambda x: x[0])

    result: list[dict] = []
    monitor_counter = 0
    for row in rows:
        result.append(row)
        if row.get("_is_monitor"):
            for cable_seq, cable_name in cables_by_monitor.get(monitor_counter, []):
                result.append({
                    "equipment":   cable_name,
                    "serial":      "",
                    "asset_tag":   "",
                    "remarks":     "",
                    "_seq_key":    cable_seq,
                    "_is_monitor": False,
                })
            monitor_counter += 1

    return [{k: v for k, v in r.items() if not k.startswith("_")} for r in result]

# ─────────────────────────────────────────────
# SMART SEARCH
# ─────────────────────────────────────────────

def _normalize(text: str) -> str:
    text = unicodedata.normalize("NFD", text)
    text = "".join(c for c in text if unicodedata.category(c) != "Mn")
    text = re.sub(r"[^\w\s]", " ", text)
    return re.sub(r"\s+", " ", text).strip().lower()

def _score_match(query: str, name: str) -> tuple:
    q_norm   = _normalize(query)
    n_norm   = _normalize(name)
    q_tokens = q_norm.split()
    n_tokens = n_norm.split()
    if not q_tokens or not n_tokens:
        return (None, 0.0)
    if q_norm == n_norm:
        return (0, 1.0)
    n_q            = len(q_tokens)
    exact_matches  = sum(1 for qt in q_tokens if qt in n_tokens)
    substr_matches = sum(1 for qt in q_tokens if any(qt in nt for nt in n_tokens))
    full_substr    = q_norm in n_norm
    coverage       = exact_matches  / n_q
    substr_cov     = substr_matches / n_q
    if exact_matches  == n_q:                         return (1, coverage)
    if substr_matches == n_q:                         return (2, substr_cov)
    if exact_matches  >= max(1, round(n_q * 0.6)):   return (3, coverage)
    if substr_matches >= max(1, round(n_q * 0.6)):   return (4, substr_cov)
    if exact_matches  >= 1:                           return (5, coverage)
    if substr_matches >= 1 or full_substr:            return (6, max(substr_cov, 0.1))
    return (None, 0.0)

def smart_search(df: pd.DataFrame, user_col: str, query: str):
    query = query.strip()
    if not query:
        return [], {}
    q_tokens   = _normalize(query).split()
    q_variants = [query]
    if len(q_tokens) > 1:
        q_variants.append(" ".join(reversed(q_tokens)))
    all_names = df[user_col].dropna().astype(str).str.strip().unique().tolist()
    scored = []
    for name in all_names:
        if not name:
            continue
        best_tier, best_score = None, 0.0
        for qv in q_variants:
            tier, score = _score_match(qv, name)
            if tier is not None:
                if best_tier is None or tier < best_tier or (tier == best_tier and score > best_score):
                    best_tier, best_score = tier, score
        if best_tier is not None:
            scored.append({"name": name, "tier": best_tier, "score": best_score})
    if not scored:
        return [], {}
    scored.sort(key=lambda x: (x["tier"], -x["score"]))
    tier_labels = {
        0: "Exact match", 1: "All words matched", 2: "All words found",
        3: "Mostly matched", 4: "Mostly found", 5: "Partial match", 6: "Similar",
    }
    for item in scored:
        item["label"] = tier_labels.get(item["tier"], "Similar")
    df_by_name = {
        item["name"]: df[df[user_col].astype(str).str.strip() == item["name"]].copy()
        for item in scored
    }
    return scored, df_by_name

# ─────────────────────────────────────────────
# WORD TEMPLATE FILLER
# ─────────────────────────────────────────────

_CT_DOTX = "application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml"
_CT_DOCX = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"

_ROW_HEIGHT_DXA = "180"
_CELL_PAD_TOP   = "0"
_CELL_PAD_BTM   = "0"
_CELL_PAD_LEFT  = "60"
_CELL_PAD_RIGHT = "60"


def _make_rPr(bold: bool = False, size: str = _FONT_SIZE) -> etree._Element:
    rPr = etree.Element(f"{{{W}}}rPr")
    if bold:
        etree.SubElement(rPr, f"{{{W}}}b")
        etree.SubElement(rPr, f"{{{W}}}bCs")
    fonts = etree.SubElement(rPr, f"{{{W}}}rFonts")
    fonts.set(f"{{{W}}}ascii",    _FONT_NAME)
    fonts.set(f"{{{W}}}hAnsi",    _FONT_NAME)
    fonts.set(f"{{{W}}}cs",       _FONT_NAME)
    fonts.set(f"{{{W}}}eastAsia", _FONT_NAME)
    sz   = etree.SubElement(rPr, f"{{{W}}}sz");   sz.set(f"{{{W}}}val",   size)
    szCs = etree.SubElement(rPr, f"{{{W}}}szCs"); szCs.set(f"{{{W}}}val", size)
    return rPr

def _make_t(text: str) -> etree._Element:
    t = etree.Element(f"{{{W}}}t")
    t.text = text
    if text and (text[0] == " " or text[-1] == " "):
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    return t

def _patch_content_types(data: bytes) -> bytes:
    return data.replace(_CT_DOTX.encode(), _CT_DOCX.encode())

def _patch_app_xml(data: bytes) -> bytes:
    try:
        root = etree.fromstring(data)
        ns = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
        for el in root.findall(f"{{{ns}}}Templates"):
            root.remove(el)
        return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    except Exception:
        return data

def _set_sdt_value(body, tag_val: str, new_text: str, bold: bool = True,
                   size: str = _FONT_SIZE) -> bool:
    for sdt in body.iter(f"{{{W}}}sdt"):
        sdtPr = sdt.find(f"{{{W}}}sdtPr")
        if sdtPr is None:
            continue
        tag_el = sdtPr.find(f"{{{W}}}tag")
        if tag_el is None or tag_el.get(f"{{{W}}}val") != tag_val:
            continue
        showing = sdtPr.find(f"{{{W}}}showingPlcHdr")
        if showing is not None:
            sdtPr.remove(showing)
        sdtContent = sdt.find(f"{{{W}}}sdtContent")
        if sdtContent is None:
            sdtContent = etree.SubElement(sdt, f"{{{W}}}sdtContent")
        for ch in list(sdtContent):
            sdtContent.remove(ch)
        p   = etree.SubElement(sdtContent, f"{{{W}}}p")
        pPr = etree.SubElement(p,          f"{{{W}}}pPr")
        jc  = etree.SubElement(pPr, f"{{{W}}}jc")
        jc.set(f"{{{W}}}val", "left")
        sp  = etree.SubElement(pPr, f"{{{W}}}spacing")
        sp.set(f"{{{W}}}before",   "0")
        sp.set(f"{{{W}}}after",    "0")
        sp.set(f"{{{W}}}line",     "240")
        sp.set(f"{{{W}}}lineRule", "auto")
        r = etree.SubElement(p, f"{{{W}}}r")
        r.append(_make_rPr(bold=bold, size=size))
        r.append(_make_t(new_text))
        return True
    return False

def _fill_position_sdt(body, position: str) -> bool:
    return _set_sdt_value(body, "Contact No.", position, bold=True, size=_FONT_SIZE)

def _fill_prepared_by(body, prepared_by: str) -> bool:
    if not prepared_by:
        return False
    tables = list(body.iter(f"{{{W}}}tbl"))
    if len(tables) < 3:
        return _fill_prepared_by_fallback(body, prepared_by)
    sig_table = tables[2]
    rows = sig_table.findall(f"{{{W}}}tr")
    if len(rows) < 5:
        return _fill_prepared_by_fallback(body, prepared_by)
    cells = rows[4].findall(f"{{{W}}}tc")
    if len(cells) < 2:
        return _fill_prepared_by_fallback(body, prepared_by)
    replaced = False
    for t_el in cells[1].iter(f"{{{W}}}t"):
        if t_el.text and "[STAFF NAME]" in t_el.text:
            t_el.text = t_el.text.replace("[STAFF NAME]", prepared_by)
            replaced = True
    return replaced

def _fill_prepared_by_fallback(body, prepared_by: str) -> bool:
    replaced = False
    for t_el in body.iter(f"{{{W}}}t"):
        if t_el.text and "[STAFF NAME]" in t_el.text:
            t_el.text = t_el.text.replace("[STAFF NAME]", prepared_by)
            replaced = True
    return replaced

def _get_equipment_table(body):
    for tbl in body.iter(f"{{{W}}}tbl"):
        tblGrid = tbl.find(f"{{{W}}}tblGrid")
        if tblGrid is not None and len(tblGrid.findall(f"{{{W}}}gridCol")) == 5:
            return tbl
    return None

def _set_cell_text(cell_el, text: str):
    p_list = cell_el.findall(f"{{{W}}}p")
    p_el   = p_list[0] if p_list else etree.SubElement(cell_el, f"{{{W}}}p")
    for tag in [f"{{{W}}}r", f"{{{W}}}sdt"]:
        for el in p_el.findall(tag):
            p_el.remove(el)
    pPr = p_el.find(f"{{{W}}}pPr")
    if pPr is None:
        pPr = etree.SubElement(p_el, f"{{{W}}}pPr")
        p_el.insert(0, pPr)
    for tag in [f"{{{W}}}jc", f"{{{W}}}spacing", f"{{{W}}}contextualSpacing"]:
        for el in pPr.findall(tag):
            pPr.remove(el)
    jc = etree.SubElement(pPr, f"{{{W}}}jc")
    jc.set(f"{{{W}}}val", "left")
    sp = etree.SubElement(pPr, f"{{{W}}}spacing")
    sp.set(f"{{{W}}}before",   "0")
    sp.set(f"{{{W}}}after",    "0")
    sp.set(f"{{{W}}}line",     "240")
    sp.set(f"{{{W}}}lineRule", "auto")
    r_el = etree.SubElement(p_el, f"{{{W}}}r")
    r_el.append(_make_rPr(bold=False, size=_TABLE_SIZE))
    r_el.append(_make_t(text))

def _compact_row(row_el):
    trPr = row_el.find(f"{{{W}}}trPr")
    if trPr is None:
        trPr = etree.SubElement(row_el, f"{{{W}}}trPr")
        row_el.insert(0, trPr)
    for trH in trPr.findall(f"{{{W}}}trHeight"):
        trPr.remove(trH)
    trH = etree.SubElement(trPr, f"{{{W}}}trHeight")
    trH.set(f"{{{W}}}val",   _ROW_HEIGHT_DXA)
    trH.set(f"{{{W}}}hRule", "exact")
    for tc in row_el.iter(f"{{{W}}}tc"):
        tcPr = tc.find(f"{{{W}}}tcPr")
        if tcPr is None:
            tcPr = etree.SubElement(tc, f"{{{W}}}tcPr")
            tc.insert(0, tcPr)
        tcMar = tcPr.find(f"{{{W}}}tcMar")
        if tcMar is None:
            tcMar = etree.SubElement(tcPr, f"{{{W}}}tcMar")
        for side, val in [
            ("top",    _CELL_PAD_TOP),  ("bottom", _CELL_PAD_BTM),
            ("left",   _CELL_PAD_LEFT), ("right",  _CELL_PAD_RIGHT),
        ]:
            el = tcMar.find(f"{{{W}}}{side}")
            if el is None:
                el = etree.SubElement(tcMar, f"{{{W}}}{side}")
            el.set(f"{{{W}}}w",    val)
            el.set(f"{{{W}}}type", "dxa")
        for va in tcPr.findall(f"{{{W}}}vAlign"):
            tcPr.remove(va)
        vAlign = etree.SubElement(tcPr, f"{{{W}}}vAlign")
        vAlign.set(f"{{{W}}}val", "top")
        for p in tc.iter(f"{{{W}}}p"):
            pPr = p.find(f"{{{W}}}pPr")
            if pPr is not None:
                for spacing in pPr.findall(f"{{{W}}}spacing"):
                    pPr.remove(spacing)
                for cs in pPr.findall(f"{{{W}}}contextualSpacing"):
                    pPr.remove(cs)
            for r in p.findall(f"{{{W}}}r"):
                rPr = r.find(f"{{{W}}}rPr")
                if rPr is not None:
                    for spacing in rPr.findall(f"{{{W}}}spacing"):
                        rPr.remove(spacing)

def _compact_page_margins(body):
    sectPr = body.find(f"{{{W}}}sectPr")
    if sectPr is None:
        return
    pgMar = sectPr.find(f"{{{W}}}pgMar")
    if pgMar is None:
        pgMar = etree.SubElement(sectPr, f"{{{W}}}pgMar")
    limits = {"top": 720, "bottom": 720, "left": 720, "right": 720}
    for side, max_val in limits.items():
        current = pgMar.get(f"{{{W}}}{side}")
        try:
            if current is None or int(current) > max_val:
                pgMar.set(f"{{{W}}}{side}", str(max_val))
        except ValueError:
            pgMar.set(f"{{{W}}}{side}", str(max_val))

def _shrink_header_sdt_cells(body):
    tables = list(body.iter(f"{{{W}}}tbl"))
    if not tables:
        return
    header_tbl        = tables[0]
    HEADER_ROW_HEIGHT = "250"
    LABEL_CELL_WIDTH  = "1500"
    VALUE_CELL_WIDTH  = "3000"
    for tr in header_tbl.findall(f"{{{W}}}tr"):
        trPr = tr.find(f"{{{W}}}trPr")
        if trPr is None:
            trPr = etree.SubElement(tr, f"{{{W}}}trPr")
            tr.insert(0, trPr)
        for trH in trPr.findall(f"{{{W}}}trHeight"):
            trPr.remove(trH)
        trH = etree.SubElement(trPr, f"{{{W}}}trHeight")
        trH.set(f"{{{W}}}val",   HEADER_ROW_HEIGHT)
        trH.set(f"{{{W}}}hRule", "exact")
        for tc in tr.findall(f"{{{W}}}tc"):
            has_sdt = bool(tc.find(f".//{{{W}}}sdt"))
            tcPr = tc.find(f"{{{W}}}tcPr")
            if tcPr is None:
                tcPr = etree.SubElement(tc, f"{{{W}}}tcPr")
                tc.insert(0, tcPr)
            tcW = tcPr.find(f"{{{W}}}tcW")
            if tcW is None:
                tcW = etree.SubElement(tcPr, f"{{{W}}}tcW")
            tcW.set(f"{{{W}}}w",    VALUE_CELL_WIDTH if has_sdt else LABEL_CELL_WIDTH)
            tcW.set(f"{{{W}}}type", "dxa")
            tcMar = tcPr.find(f"{{{W}}}tcMar")
            if tcMar is None:
                tcMar = etree.SubElement(tcPr, f"{{{W}}}tcMar")
            for side in ("top", "bottom", "left", "right"):
                el = tcMar.find(f"{{{W}}}{side}")
                if el is None:
                    el = etree.SubElement(tcMar, f"{{{W}}}{side}")
                el.set(f"{{{W}}}w",    "40" if side in ("left", "right") else "0")
                el.set(f"{{{W}}}type", "dxa")
            for va in tcPr.findall(f"{{{W}}}vAlign"):
                tcPr.remove(va)
            vAlign = etree.SubElement(tcPr, f"{{{W}}}vAlign")
            vAlign.set(f"{{{W}}}val", "top")
            for p in tc.iter(f"{{{W}}}p"):
                pPr = p.find(f"{{{W}}}pPr")
                if pPr is None:
                    pPr = etree.SubElement(p, f"{{{W}}}pPr")
                    p.insert(0, pPr)
                for sp in pPr.findall(f"{{{W}}}spacing"):
                    pPr.remove(sp)
                sp_el = etree.SubElement(pPr, f"{{{W}}}spacing")
                sp_el.set(f"{{{W}}}before",   "0")
                sp_el.set(f"{{{W}}}after",    "0")
                sp_el.set(f"{{{W}}}line",     "240")
                sp_el.set(f"{{{W}}}lineRule", "auto")

def _fill_equipment_row(row_el, equipment, serial, asset_tag, remarks):
    cells = []
    for ch in row_el:
        if ch.tag == f"{{{W}}}tc":
            cells.append(ch)
        elif ch.tag == f"{{{W}}}sdt":
            sc = ch.find(f"{{{W}}}sdtContent")
            if sc is not None:
                for tc in sc.findall(f"{{{W}}}tc"):
                    cells.append(tc)
    for i, (cell, text) in enumerate(zip(cells, ["", equipment, serial, asset_tag, remarks])):
        if i == 0:
            continue
        _set_cell_text(cell, text)

def fill_template(
    sorted_rows: list[dict],
    employee_name: str,
    client: str,
    position: str,
    date_str: str,
    prepared_by: str = "",
    form_type: str = "wfh",
) -> bytes:
    template_filename = WFH_TEMPLATE_NAME if form_type == "wfh" else ONSITE_TEMPLATE_NAME
    tpl = _find_template(template_filename)
    if tpl is None:
        raise FileNotFoundError(
            f"Template not found: '{template_filename}'. "
            f"Place it inside the src/ folder next to this script."
        )
    with zipfile.ZipFile(io.BytesIO(tpl.read_bytes())) as zin:
        files = {n: zin.read(n) for n in zin.namelist()}
    if "[Content_Types].xml" in files:
        files["[Content_Types].xml"] = _patch_content_types(files["[Content_Types].xml"])
    if "docProps/app.xml" in files:
        files["docProps/app.xml"] = _patch_app_xml(files["docProps/app.xml"])
    root = etree.fromstring(files["word/document.xml"])
    body = root.find(f"{{{W}}}body")

    _set_sdt_value(body, "Name",   employee_name, bold=True, size=_FONT_SIZE)
    _set_sdt_value(body, "Client", client,        bold=True, size=_FONT_SIZE)
    _set_sdt_value(body, "Date",   date_str,      bold=True, size=_FONT_SIZE)
    _fill_position_sdt(body, position)
    _shrink_header_sdt_cells(body)
    _compact_page_margins(body)
    if prepared_by:
        _fill_prepared_by(body, prepared_by)

    eq_table = _get_equipment_table(body)
    if eq_table is not None:
        all_rows     = eq_table.findall(f"{{{W}}}tr")
        data_rows    = all_rows[1:]
        template_row = copy.deepcopy(data_rows[0]) if data_rows else None
        num_assets   = len(sorted_rows)

        for i, asset in enumerate(sorted_rows):
            if i < len(data_rows):
                row_el = data_rows[i]
                _compact_row(row_el)
                _fill_equipment_row(row_el, asset["equipment"], asset["serial"],
                                    asset["asset_tag"], asset["remarks"])
            elif template_row is not None:
                new_row = copy.deepcopy(template_row)
                _compact_row(new_row)
                eq_table.append(new_row)
                _fill_equipment_row(new_row, asset["equipment"], asset["serial"],
                                    asset["asset_tag"], asset["remarks"])

        all_rows_now = eq_table.findall(f"{{{W}}}tr")
        for extra_row in all_rows_now[1 + num_assets:]:
            eq_table.remove(extra_row)
        for data_row in eq_table.findall(f"{{{W}}}tr")[1:]:
            _compact_row(data_row)

    files["word/document.xml"] = etree.tostring(
        root, xml_declaration=True, encoding="UTF-8", standalone=True)
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in files.items():
            zout.writestr(name, data)
    return out.getvalue()

# ─────────────────────────────────────────────
# UI HELPERS
# ─────────────────────────────────────────────

def render_header():
    st.markdown(f"""
    <div class="bmg-header">
      <div>
        <div class="bmg-header-title">Asset Accountability System</div>
        <div class="bmg-header-sub">Equipment Accountability Form Generator</div>
      </div>
      {LOGO_HTML}
    </div>
    """, unsafe_allow_html=True)

def step_open(num: int, title: str, desc: str = ""):
    desc_html = f'<div class="step-desc">{desc}</div>' if desc else ""
    st.markdown(f"""
    <div class="step-card">
      <div class="step-label">
        <span class="step-badge">{num}</span>
        <span class="step-title">{title}</span>
      </div>
      {desc_html}
    """, unsafe_allow_html=True)

def step_close():
    st.markdown("</div>", unsafe_allow_html=True)

def _safe_periph_key(name: str) -> str:
    return re.sub(r"[^a-z0-9_]", "_", name.lower())

# ─────────────────────────────────────────────
# FORM TYPE SELECTOR
# ─────────────────────────────────────────────

def render_form_type_selector() -> str:
    """
    Visual card + radio selector for WFH vs On-Site.
    Returns "wfh" or "onsite".
    Switching clears any previously generated document so the download
    button always matches the selected template.
    """
    if "form_type" not in st.session_state:
        st.session_state["form_type"] = "wfh"

    current = st.session_state["form_type"]

    wfh_tpl    = _find_template(WFH_TEMPLATE_NAME)
    onsite_tpl = _find_template(ONSITE_TEMPLATE_NAME)

    def _missing(found):
        return "" if found else \
            " <span style='color:#e65c00;font-size:0.7rem;font-weight:400;'>(template missing)</span>"

    wfh_css    = "active-wfh"    if current == "wfh"    else ""
    onsite_css = "active-onsite" if current == "onsite" else ""
    wfh_badge    = '<span class="form-type-badge badge-wfh">Selected</span>'    if current == "wfh"    else ""
    onsite_badge = '<span class="form-type-badge badge-onsite">Selected</span>' if current == "onsite" else ""

    st.markdown(f"""
    <div class="form-type-selector">
      <div class="form-type-card {wfh_css}">
        <span class="form-type-icon">🏠</span>
        <div class="form-type-label">Work From Home{_missing(wfh_tpl)}</div>
        <div class="form-type-desc">For employees taking equipment home to work remotely.</div>
        {wfh_badge}
      </div>
      <div class="form-type-card {onsite_css}">
        <span class="form-type-icon">🏢</span>
        <div class="form-type-label">Work On Site{_missing(onsite_tpl)}</div>
        <div class="form-type-desc">For employees assigned equipment at the office.</div>
        {onsite_badge}
      </div>
    </div>
    """, unsafe_allow_html=True)

    choice = st.radio(
        "Select form type",
        options=["🏠  Work From Home", "🏢  Work On Site"],
        index=0 if current == "wfh" else 1,
        horizontal=True,
        label_visibility="collapsed",
        key="form_type_radio",
    )

    new_type = "wfh" if "Home" in choice else "onsite"
    if new_type != current:
        st.session_state["form_type"] = new_type
        # Clear stale generated document when template switches
        for k in ("docx", "form_name", "form_client", "form_date",
                  "prepared_by", "form_type_used"):
            st.session_state.pop(k, None)
        st.rerun()

    return st.session_state["form_type"]

# ─────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────

def main():
    render_header()

    # ── Form Type Selector ────────────────────────────────────────────────────
    st.markdown(
        '<div style="font-size:0.72rem;font-weight:700;color:#5a6e8a;'
        'text-transform:uppercase;letter-spacing:0.07em;margin-bottom:0.35rem;">'
        'Form Type</div>'
        '<div style="font-size:0.8rem;color:#5a6e8a;margin-bottom:0.6rem;line-height:1.5;">'
        'Choose which accountability form template to generate.'
        '</div>',
        unsafe_allow_html=True,
    )

    form_type = render_form_type_selector()

    # Resolve template path for the selected type
    if form_type == "wfh":
        tpl_path       = _find_template(WFH_TEMPLATE_NAME)
        type_label     = "Work From Home"
        type_color     = "#1565c0"
        type_bg        = "#e8f0fc"
        type_icon      = "🏠"
        type_short     = "WFH"
    else:
        tpl_path       = _find_template(ONSITE_TEMPLATE_NAME)
        type_label     = "Work On Site"
        type_color     = "#00796b"
        type_bg        = "#e0f2f1"
        type_icon      = "🏢"
        type_short     = "On Site"

    if tpl_path is None:
        st.error(
            f"⚠️ Template not found for **{type_label}**. "
            f"Place the `.dotx` file inside the **src/** folder and refresh.\n\n"
            f"Expected: `{'Equipment Accountability Form (Work From Home).dotx' if form_type == 'wfh' else 'Equipment Accountability Form (Work On Site).dotx'}`"
        )
    else:
        st.markdown(
            f'<div style="font-size:0.76rem;color:{type_color};background:{type_bg};'
            f'border-left:3px solid {type_color};border-radius:0 6px 6px 0;'
            f'padding:0.4rem 0.75rem;margin-bottom:0.8rem;">'
            f'{type_icon} Using template: <strong>{type_label}</strong>'
            f'</div>',
            unsafe_allow_html=True,
        )

    st.markdown("<hr>", unsafe_allow_html=True)

    # ── Prepared By ───────────────────────────────────────────────────────────
    st.markdown("""
    <div class="preparedby-card">
      <div class="preparedby-icon">🖊️</div>
      <div>
        <div class="preparedby-label">Prepared By</div>
        <div class="preparedby-hint">Select the IT staff member preparing this form — their name will be auto-filled in the <strong>Prepared by</strong> section of the document.</div>
      </div>
    </div>
    """, unsafe_allow_html=True)

    prepared_by = st.selectbox(
        "Select IT staff",
        options=PREPARED_BY_OPTIONS,
        index=0,
        label_visibility="collapsed",
        key="prepared_by_select",
    )

    st.markdown("<div style='height:0.5rem'></div>", unsafe_allow_html=True)

    # ── Step 1: Upload ────────────────────────────────────────────────────────
    step_open(1, "Upload Asset List",
              "Export your SharePoint asset list as CSV (Export → Export to CSV) and upload it below.")
    uploaded = st.file_uploader("CSV file", type=["csv"], label_visibility="collapsed")
    step_close()

    if not uploaded:
        st.markdown("""
        <div style="text-align:center;padding:2.5rem 1rem;color:#5a6e8a;font-size:0.84rem;
                    background:#fff;border:1px solid #d0dce8;border-radius:10px;margin-top:.5rem;">
            Upload your SharePoint CSV export above to get started
        </div>
        """, unsafe_allow_html=True)
        return

    with st.spinner("Reading file…"):
        df, error = load_csv(uploaded)

    if error:
        st.error(f"Could not read the file: {error}")
        return
    if df.empty:
        st.warning("The file is empty. Please check your export and try again.")
        return

    auto    = detect_columns(df)
    col_map = {
        key: (auto.get(key) if auto.get(key) and auto[key] in df.columns else None)
        for key in CSV_COLUMNS
    }
    user_col     = col_map.get("current_user")
    client_col   = col_map.get("client")
    position_col = detect_position_column(df)

    if not user_col:
        st.error('Could not find a "Current User" column. Make sure your CSV was exported from SharePoint with standard column names.')
        return

    st.success(f"✓ &nbsp;**{len(df):,}** records loaded from **{uploaded.name}**")
    if position_col:
        st.caption(f"Position column detected: **{position_col}**")
    else:
        st.caption("ℹ️ No Position/Job Title column detected in this CSV.")

    # ── Step 2: Search ────────────────────────────────────────────────────────
    step_open(2, "Find Employee",
              "Type any part of the name — first name, last name, or partial words.")
    st.markdown(
        '<div class="search-hint">'
        '💡 <strong>How search works:</strong> You can type any part of a name — '
        'first name only, last name only, or a combination in any order. '
        'Partial words are supported, so short fragments will still return results. '
        'Names stored with commas (e.g. Last, First Middle format) are matched even if you '
        'type the first name first. Results are ranked by match quality — '
        'closest matches appear at the top.'
        '</div>',
        unsafe_allow_html=True,
    )
    search = st.text_input(
        "Search", placeholder="Type a name to search…",
        label_visibility="collapsed",
    )
    step_close()

    if not search.strip():
        st.info("Type an employee name above to continue.")
        return

    results, df_by_name = smart_search(df, user_col, search)

    if not results:
        st.warning(f'No records found for "{search.strip()}". Try a shorter name or check spelling.')
        return

    top_results    = [r for r in results if r["tier"] <= 1]
    strong_results = [r for r in results if 2 <= r["tier"] <= 3]
    weak_results   = [r for r in results if r["tier"] >= 4]
    st.caption(f"{len(results)} result(s) — select an employee to continue")

    def make_label(item):
        count = len(df_by_name.get(item["name"], pd.DataFrame()))
        return f"{item['name']}  [{item['label']}]  ({count} asset{'s' if count != 1 else ''})"

    ordered  = top_results + strong_results + weak_results
    options  = [make_label(r) for r in ordered]
    name_map = {make_label(r): r["name"] for r in ordered}

    if len(options) == 1:
        chosen_display = options[0]
        if ordered[0]["tier"] <= 1:
            st.success("✓ &nbsp;One match found.")
        else:
            st.info(f"Closest result found ({ordered[0]['label']}). Confirm the selection below.")
    else:
        chosen_display = st.selectbox(
            "Select employee", options,
            help="Names ranked by match quality — exact matches appear first.",
        )

    chosen_name = name_map[chosen_display]
    df_filtered = df_by_name.get(chosen_name, pd.DataFrame())

    if df_filtered.empty:
        st.warning("No assets found for the selected employee.")
        return

    chosen_result = next((r for r in results if r["name"] == chosen_name), None)
    if chosen_result and chosen_result["tier"] >= 4:
        st.info(f"Showing results for the selected name ({chosen_result['label']} for your search). Not the right person? Select a different name above.")
    else:
        st.success(f"✓ &nbsp;**{len(df_filtered)}** asset(s) found.")

    # ── Step 3: Select CSV assets ─────────────────────────────────────────────
    step_open(3, "Select Assets from SharePoint",
              "Choose the equipment items from the asset list to include on the accountability form.")

    sel_key = f"sel_{chosen_name.lower().replace(' ', '_')}"
    if sel_key not in st.session_state:
        st.session_state[sel_key] = {idx: False for idx in df_filtered.index}
    for idx in df_filtered.index:
        if idx not in st.session_state[sel_key]:
            st.session_state[sel_key][idx] = False

    col_a, col_b, _ = st.columns([1, 1, 5])
    with col_a:
        st.markdown('<div class="small-btn">', unsafe_allow_html=True)
        if st.button("Select All", key="sp_select_all"):
            for idx in df_filtered.index:
                st.session_state[sel_key][idx] = True
            st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)
    with col_b:
        st.markdown('<div class="small-btn">', unsafe_allow_html=True)
        if st.button("Clear All", key="sp_clear_all"):
            for idx in df_filtered.index:
                st.session_state[sel_key][idx] = False
            st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)

    checked = []
    for idx, row in df_filtered.iterrows():
        tag   = safe_str(row.get(col_map.get("asset_tag")    or "", ""))
        ct    = safe_str(row.get(col_map.get("content_type") or "", ""))
        brand = safe_str(row.get(col_map.get("brand")        or "", ""))
        model = safe_str(row.get(col_map.get("model")        or "", ""))
        sn    = safe_str(row.get(col_map.get("serial")       or "", ""))
        cond  = safe_str(row.get(col_map.get("condition")    or "", ""))
        parts = [p for p in [brand, model, ct] if p]
        desc  = " ".join(parts) if parts else ""
        meta  = "  ·  ".join(p for p in [f"S/N: {sn}" if sn else "", cond] if p)
        label = f"**{tag}**" + (f" — {desc}" if desc else "") + (f"  ·  {meta}" if meta else "")
        val = st.checkbox(
            label,
            value=st.session_state[sel_key].get(idx, False),
            key=f"chk_{idx}_{sel_key}",
        )
        st.session_state[sel_key][idx] = val
        if val:
            checked.append(idx)

    step_close()

    df_selected = df_filtered.loc[checked].copy()

    # ── Identify selected monitors ────────────────────────────────────────────
    selected_monitors: list[dict] = []
    temp_rows_for_monitors = []
    for _, row in df_selected.iterrows():
        ct    = safe_str(row.get(col_map.get("content_type") or "", ""))
        brand = safe_str(row.get(col_map.get("brand")        or "", ""))
        model = safe_str(row.get(col_map.get("model")        or "", ""))
        eq    = _build_equipment_label(brand, model, ct, "")
        seq   = _get_sequence_key(eq)
        temp_rows_for_monitors.append((seq, eq))
    temp_rows_for_monitors.sort(key=lambda x: x[0])
    monitor_counter = 0
    for _, eq in temp_rows_for_monitors:
        if _is_monitor(eq):
            label = eq if eq.strip() else f"Monitor {monitor_counter + 1}"
            selected_monitors.append({"label": label, "idx": monitor_counter})
            monitor_counter += 1

    # ── Step 4: Peripherals ───────────────────────────────────────────────────
    step_open(4, "Add Peripherals",
              "Select simple accessories and assign cables to specific monitors.")

    st.markdown("**Generic Accessories**")

    for periph_name, _ in SIMPLE_PERIPHERALS:
        pkey = f"chk_simple_{_safe_periph_key(periph_name)}"
        if pkey not in st.session_state:
            st.session_state[pkey] = False

    col_sa, col_ca, _ = st.columns([1, 1, 5])
    with col_sa:
        st.markdown('<div class="small-btn">', unsafe_allow_html=True)
        if st.button("Select All", key="simple_select_all"):
            for periph_name, _ in SIMPLE_PERIPHERALS:
                st.session_state[f"chk_simple_{_safe_periph_key(periph_name)}"] = True
            st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)
    with col_ca:
        st.markdown('<div class="small-btn">', unsafe_allow_html=True)
        if st.button("Clear All", key="simple_clear_all"):
            for periph_name, _ in SIMPLE_PERIPHERALS:
                st.session_state[f"chk_simple_{_safe_periph_key(periph_name)}"] = False
            st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)

    simple_extra_rows: list[dict] = []
    col1, col2, col3 = st.columns(3)
    grid_cols = [col1, col2, col3]
    for i, (periph_name, periph_type) in enumerate(SIMPLE_PERIPHERALS):
        pkey = f"chk_simple_{_safe_periph_key(periph_name)}"
        with grid_cols[i % 3]:
            is_checked = st.checkbox(periph_name, key=pkey,
                                     value=st.session_state.get(pkey, False))
        if is_checked:
            simple_extra_rows.append({
                "name": periph_name, "brand": "", "model": "",
                "type": periph_type, "serial": "", "asset_tag": "", "remarks": "",
            })

    # ── Per-monitor cables ────────────────────────────────────────────────────
    monitor_cable_assignments: list[dict] = []

    if selected_monitors:
        st.markdown(
            "<div style='margin-top:1.1rem;margin-bottom:0.3rem;"
            "font-size:0.9rem;font-weight:700;color:var(--navy,#0d2545);'>"
            "Monitor Cables &amp; Adapters</div>",
            unsafe_allow_html=True,
        )
        st.caption("Each monitor is listed separately. Tick which cables/adapters came with that specific monitor.")

        for mon in selected_monitors:
            mon_label = mon["label"]
            mon_idx   = mon["idx"]
            st.markdown(
                f'<div class="monitor-group-header">'
                f'🖥️ Monitor {mon_idx + 1}'
                f'<span style="font-weight:400;color:#5a6e8a;margin-left:4px;">— {mon_label}</span>'
                f'</div>',
                unsafe_allow_html=True,
            )
            st.markdown('<div class="monitor-group-body">', unsafe_allow_html=True)
            cable_cols = st.columns(3)
            for ci, (cable_name, cable_type, cable_seq) in enumerate(MONITOR_PERIPHERALS):
                ckey = f"chk_mon{mon_idx}_{_safe_periph_key(cable_name)}"
                if ckey not in st.session_state:
                    st.session_state[ckey] = False
                with cable_cols[ci % 3]:
                    is_checked = st.checkbox(cable_name, key=ckey,
                                             value=st.session_state.get(ckey, False))
                if is_checked:
                    monitor_cable_assignments.append({
                        "monitor_label": mon_label,
                        "monitor_idx":   mon_idx,
                        "cable_name":    cable_name,
                        "cable_seq":     cable_seq,
                    })
            st.markdown('</div>', unsafe_allow_html=True)

    elif not df_selected.empty:
        st.caption("ℹ️ No monitors in the selected assets — monitor cable section not shown.")

    step_close()

    # ── Build sorted rows ─────────────────────────────────────────────────────
    sorted_rows = sort_assets_by_sequence(
        df_selected, col_map, simple_extra_rows, monitor_cable_assignments,
    )

    if not sorted_rows:
        st.info("Select at least one asset or peripheral to continue.")
        return

    cable_count = len(monitor_cable_assignments)
    st.caption(
        f"**{len(checked)}** SharePoint asset(s)"
        + f" + **{len(simple_extra_rows)}** accessory/accessories"
        + (f" + **{cable_count}** monitor cable(s)" if cable_count else "")
        + f" = **{len(sorted_rows)}** total item(s)"
    )

    with st.expander("📋 Preview — Sorted Form Order", expanded=False):
        st.markdown(
            '<table class="preview-table"><thead><tr>'
            '<th>#</th><th>Equipment (Brand Model Type)</th>'
            '<th>Serial</th><th>Asset Tag</th><th>Remarks</th>'
            '</tr></thead><tbody>' +
            "".join(
                f'<tr>'
                f'<td><span class="seq-badge">{i+1}</span></td>'
                f'<td>{row["equipment"]}</td><td>{row["serial"]}</td>'
                f'<td>{row["asset_tag"]}</td><td>{row["remarks"]}</td>'
                f'</tr>'
                for i, row in enumerate(sorted_rows)
            ) +
            "</tbody></table>",
            unsafe_allow_html=True,
        )

    # ── Step 5: Form details ──────────────────────────────────────────────────
    step_open(5, "Form Details",
              "These fields appear in the header of the generated document. Edit as needed.")

    default_client = ""
    default_pos    = ""

    if not df_selected.empty:
        first_row = df_selected.iloc[0]
        if client_col and client_col in df_selected.columns:
            default_client = safe_str(first_row.get(client_col, ""))
        default_pos = get_position_value(first_row, list(df_selected.columns), position_col)
        if not default_pos:
            for col in df_selected.columns:
                col_l = col.lower()
                if any(kw in col_l for kw in ("position", "job", "designation", "title", "role")):
                    val = safe_str(first_row.get(col, ""))
                    if val:
                        default_pos = val
                        break

    f1, f2 = st.columns(2)
    with f1:
        form_name     = st.text_input("Full Name", value=chosen_name)
        form_client   = st.text_input("Client",    value=default_client)
    with f2:
        form_position = st.text_input("Position",  value=default_pos)
        form_date     = st.date_input("Date",       value=datetime.today())

    step_close()
    form_date_str = form_date.strftime("%B %d, %Y")

    # ── Step 6: Generate ──────────────────────────────────────────────────────
    step_open(6, "Generate Document", "")

    st.markdown(
        f'<div class="search-hint" style="margin-bottom:1rem;">'
        f'✏️ <strong>Prepared by:</strong> {prepared_by} &nbsp;·&nbsp; '
        f'<strong>Employee:</strong> {chosen_name} &nbsp;·&nbsp; '
        f'<strong>Items:</strong> {len(sorted_rows)} &nbsp;·&nbsp; '
        f'{type_icon} <strong>Form:</strong> {type_label}'
        f'</div>',
        unsafe_allow_html=True,
    )

    if tpl_path is None:
        st.error(
            f"Cannot generate — **{type_label}** template not found. "
            f"Place the .dotx file in the **src/** folder and refresh."
        )
        step_close()
        return

    if st.button("⬇  Generate Word Document", use_container_width=True, type="primary"):
        with st.spinner("Filling the form…"):
            try:
                docx = fill_template(
                    sorted_rows,
                    form_name, form_client, form_position, form_date_str,
                    prepared_by=prepared_by,
                    form_type=form_type,
                )
                st.session_state["docx"]          = docx
                st.session_state["form_name"]     = form_name
                st.session_state["form_client"]   = form_client
                st.session_state["form_date"]     = form_date
                st.session_state["prepared_by"]   = prepared_by
                st.session_state["form_type_used"] = form_type
                st.success("✓ &nbsp;Document ready — click Download below.")
            except Exception as e:
                st.error(f"Error: {e}")

    if "docx" in st.session_state:
        _fname       = st.session_state.get("form_name")      or "Employee"
        _fclient     = st.session_state.get("form_client")    or ""
        _fdate       = st.session_state.get("form_date")      or datetime.now()
        _ftype_used  = st.session_state.get("form_type_used") or "wfh"
        _client_part = f" ({_fclient})" if _fclient else ""
        _day         = str(int(_fdate.strftime("%d")))
        _date_part   = _day + _fdate.strftime(" %B %Y")
        _type_part   = "WFH" if _ftype_used == "wfh" else "On Site"
        _filename    = (
            f"Employee Copy - Equipment Accountability Form ({_type_part})"
            f" - {_fname}{_client_part} _ {_date_part}.docx"
        )
        st.download_button(
            "📄  Download Word (.docx)",
            data=st.session_state["docx"],
            file_name=_filename,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
        )

    step_close()

    # ── Summary metrics ───────────────────────────────────────────────────────
    st.markdown("<div style='height:.4rem'></div>", unsafe_allow_html=True)
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total Records",   f"{len(df):,}")
    c2.metric("Matched",         f"{len(df_filtered):,}")
    c3.metric("From SharePoint", f"{len(df_selected):,}")
    c4.metric("On This Form",    f"{len(sorted_rows):,}")

    st.markdown(
        '<div class="bmg-footer">BMG Outsourcing, Inc. &nbsp;·&nbsp; Asset Accountability System</div>',
        unsafe_allow_html=True,
    )


if __name__ == "__main__":
    main()
