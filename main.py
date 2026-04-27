"""
Equipment Accountability Form Generator
BMG Outsourcing — CSV upload version.
Supports both Work From Home and Work On Site templates.
"""

from __future__ import annotations

import io
import copy
import base64
import re
import unicodedata
import zipfile
import json
from datetime import datetime
from pathlib import Path
from lxml import etree

import pandas as pd
import streamlit as st

# ─────────────────────────────────────────────
# PATHS
# ─────────────────────────────────────────────

WFH_TEMPLATE_NAME    = "Equipment Accountability Form (Work From Home).dotx"
ONSITE_TEMPLATE_NAME = "Equipment Accountability Form (Work On Site).dotx"

REGISTERED_USERS_PATH = Path(__file__).parent / "registered_users.json"


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


LOGO_B64  = get_logo_b64()
LOGO_HTML = (
    f'<img src="data:image/png;base64,{LOGO_B64}" style="height:64px;object-fit:contain;">'
    if LOGO_B64 else
    '<span style="font-size:1.4rem;font-weight:900;color:#fff;letter-spacing:-1px;">BMG</span>'
)

# ─────────────────────────────────────────────
# REGISTERED USERS HELPERS
# ─────────────────────────────────────────────

DEFAULT_PREPARED_BY = [
    "IT Intern",
    "Jiro Macabitas",
    "Angelo Forbes",
    "Bryan Odero",
]


def load_registered_users() -> list[str]:
    try:
        if REGISTERED_USERS_PATH.exists():
            data = json.loads(REGISTERED_USERS_PATH.read_text(encoding="utf-8"))
            extras = [u for u in data.get("users", []) if u not in DEFAULT_PREPARED_BY]
            return DEFAULT_PREPARED_BY + extras
    except Exception:
        pass
    return list(DEFAULT_PREPARED_BY)


def save_registered_users(users: list[str]):
    extras = [u for u in users if u not in DEFAULT_PREPARED_BY]
    try:
        REGISTERED_USERS_PATH.write_text(
            json.dumps({"users": extras}, indent=2, ensure_ascii=False),
            encoding="utf-8",
        )
    except Exception:
        pass


def get_prepared_by_options() -> list[str]:
    if "prepared_by_options" not in st.session_state:
        st.session_state["prepared_by_options"] = load_registered_users()
    return st.session_state["prepared_by_options"]


def add_prepared_by_user(name: str) -> tuple[bool, str]:
    name = name.strip()
    if not name:
        return False, "Name cannot be empty."
    options = get_prepared_by_options()
    if name.lower() in [o.lower() for o in options]:
        return False, f'"{name}" is already in the list.'
    options.append(name)
    st.session_state["prepared_by_options"] = options
    save_registered_users(options)
    return True, f'"{name}" has been added.'


def remove_prepared_by_user(name: str) -> tuple[bool, str]:
    if name in DEFAULT_PREPARED_BY:
        return False, f'"{name}" is a default user and cannot be removed.'
    options = get_prepared_by_options()
    if name not in options:
        return False, f'"{name}" not found.'
    options.remove(name)
    st.session_state["prepared_by_options"] = options
    save_registered_users(options)
    return True, f'"{name}" has been removed.'


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
    --green:      #2e7d32;
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

[data-testid="stSidebar"]   {{ display: none !important; }}
header[data-testid="stHeader"] {{ display: none !important; }}
[data-testid="stDecoration"]   {{ display: none !important; }}

.main .block-container {{
    max-width: 860px !important;
    padding: 2rem 1.5rem 3rem !important;
}}

/* ── Header ── */
.bmg-header {{
    background: linear-gradient(135deg, #0d2545 0%, #1565c0 100%);
    border-radius: 0 0 14px 14px;
    padding: 1.6rem 2rem;
    display: flex;
    align-items: center;
    justify-content: space-between;
    margin: -2rem -1.5rem 2rem -1.5rem;
    box-shadow: var(--shadow-md);
}}
.bmg-header-title {{
    color: #fff;
    font-size: 1.05rem;
    font-weight: 700;
    letter-spacing: -0.01em;
}}
.bmg-header-sub {{
    color: rgba(255,255,255,.55);
    font-size: 0.74rem;
    font-weight: 400;
    margin-top: 2px;
}}

/* ── Form type toggle ── */
.form-toggle-btn .stButton > button {{
    width: 100% !important;
    padding: 0.85rem 1rem !important;
    font-size: 0.85rem !important;
    font-weight: 700 !important;
    border-radius: 10px !important;
    letter-spacing: 0.01em !important;
    transition: all .18s !important;
}}
.form-toggle-btn.wfh-active .stButton > button {{
    background: linear-gradient(135deg, #1565c0, #1e88e5) !important;
    color: #fff !important;
    border: 2px solid #1565c0 !important;
    box-shadow: 0 0 0 3px rgba(21,101,192,.18), 0 2px 10px rgba(21,101,192,.25) !important;
}}
.form-toggle-btn.wfh-inactive .stButton > button {{
    background: #fff !important;
    color: var(--navy) !important;
    border: 2px solid var(--border) !important;
    box-shadow: var(--shadow) !important;
}}
.form-toggle-btn.wfh-inactive .stButton > button:hover {{
    background: var(--blue-pale) !important;
    border-color: #b3c8f0 !important;
}}
.form-toggle-btn.onsite-active .stButton > button {{
    background: linear-gradient(135deg, #00796b, #26a69a) !important;
    color: #fff !important;
    border: 2px solid #00796b !important;
    box-shadow: 0 0 0 3px rgba(0,121,107,.18), 0 2px 10px rgba(0,121,107,.25) !important;
}}
.form-toggle-btn.onsite-inactive .stButton > button {{
    background: #fff !important;
    color: var(--navy) !important;
    border: 2px solid var(--border) !important;
    box-shadow: var(--shadow) !important;
}}
.form-toggle-btn.onsite-inactive .stButton > button:hover {{
    background: var(--teal-lt) !important;
    border-color: var(--teal-pale) !important;
}}

/* ── Step card ── */
.step-card {{
    background: var(--card);
    border: 1px solid var(--border);
    border-radius: var(--radius);
    padding: 1.3rem 1.5rem 1rem;
    margin-bottom: 1rem;
    box-shadow: var(--shadow);
}}
.step-label {{
    display: flex;
    align-items: center;
    gap: 0.6rem;
    margin-bottom: 0.5rem;
}}
.step-badge {{
    background: linear-gradient(135deg, #1565c0, #1e88e5);
    color: #fff;
    font-size: 0.67rem;
    font-weight: 700;
    width: 20px;
    height: 20px;
    border-radius: 50%;
    display: inline-flex;
    align-items: center;
    justify-content: center;
    flex-shrink: 0;
}}
.step-title {{
    font-size: 0.9rem;
    font-weight: 700;
    color: var(--navy);
}}
.step-desc {{
    font-size: 0.79rem;
    color: var(--muted);
    margin-bottom: 0.85rem;
    line-height: 1.55;
}}

/* ── CSV source pill bar ── */
.csv-source-bar {{
    display: flex;
    align-items: center;
    gap: 0.6rem;
    background: var(--green-lt);
    border: 1px solid #c3e6c0;
    border-radius: 8px;
    padding: 0.6rem 0.9rem;
    font-size: 0.8rem;
    color: var(--green);
    font-weight: 600;
    height: 38px;
    box-sizing: border-box;
}}
.csv-stale-bar {{
    display: flex;
    align-items: center;
    gap: 0.6rem;
    background: #fff8e1;
    border: 1px solid #ffe082;
    border-radius: 8px;
    padding: 0.6rem 0.9rem;
    font-size: 0.8rem;
    color: #b07d00;
    font-weight: 500;
    height: 38px;
    box-sizing: border-box;
}}

/* ── Prepared by card ── */
.preparedby-card {{
    background: linear-gradient(135deg, #f0f7ff 0%, #e8f0fc 100%);
    border: 1.5px solid #b3c8f0;
    border-radius: var(--radius);
    padding: 1rem 1.4rem 0.85rem;
    margin-bottom: 1rem;
    box-shadow: var(--shadow);
    display: flex;
    align-items: center;
    gap: 0.9rem;
}}
.preparedby-icon {{ font-size: 1.4rem; flex-shrink: 0; }}
.preparedby-label {{
    font-size: 0.7rem;
    font-weight: 700;
    color: var(--blue);
    text-transform: uppercase;
    letter-spacing: 0.07em;
    margin-bottom: 2px;
}}
.preparedby-hint {{
    font-size: 0.75rem;
    color: var(--muted);
    line-height: 1.4;
}}

/* ── User registration panel ── */
.user-reg-panel {{
    background: #f7f9fd;
    border: 1.5px solid var(--border);
    border-radius: var(--radius);
    padding: 1rem 1.2rem;
    margin-top: 0.6rem;
    margin-bottom: 0.4rem;
}}
.user-reg-title {{
    font-size: 0.72rem;
    font-weight: 700;
    color: var(--navy);
    text-transform: uppercase;
    letter-spacing: 0.07em;
    margin-bottom: 0.65rem;
    display: flex;
    align-items: center;
    gap: 0.4rem;
}}
.user-chip-wrap {{
    display: flex;
    flex-wrap: wrap;
    gap: 6px;
    margin-bottom: 0.55rem;
}}
.user-chip {{
    display: inline-flex;
    align-items: center;
    gap: 5px;
    font-size: 0.76rem;
    font-weight: 500;
    color: var(--navy);
    background: #fff;
    border: 1px solid var(--border);
    border-radius: 100px;
    padding: 3px 10px;
}}
.user-chip.default-chip {{
    color: var(--muted);
    background: #f0f4fa;
    border-style: dashed;
}}

/* ── Inline row alignment helper ── */
[data-testid="stHorizontalBlock"] > div {{
    display: flex;
    flex-direction: column;
    justify-content: flex-end;
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
    font-size: 0.71rem !important;
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

/* ── Outline-blue buttons ── */
.btn-outline-blue .stButton > button {{
    background: #fff !important;
    color: var(--blue) !important;
    border: 1.5px solid var(--blue) !important;
    font-size: 0.75rem !important;
    padding: 0.25rem 0.8rem !important;
    box-shadow: none !important;
    font-weight: 600 !important;
    height: 38px !important;
    min-height: 38px !important;
}}
.btn-outline-blue .stButton > button:hover {{
    background: var(--blue-pale) !important;
    transform: none !important;
    box-shadow: none !important;
}}

/* ── Danger/outline-red buttons ── */
.btn-danger .stButton > button {{
    background: #fff !important;
    color: #c62828 !important;
    border: 1.5px solid #ef9a9a !important;
    font-size: 0.75rem !important;
    padding: 0.25rem 0.8rem !important;
    box-shadow: none !important;
    font-weight: 600 !important;
    height: 38px !important;
    min-height: 38px !important;
}}
.btn-danger .stButton > button:hover {{
    background: #ffebee !important;
    border-color: #c62828 !important;
    transform: none !important;
    box-shadow: none !important;
}}

/* ── Column-scoped small button sizing ── */
[data-testid="stHorizontalBlock"] [data-testid="stColumn"] .stButton > button {{
    height: 38px !important;
    min-height: 38px !important;
    padding-top: 0 !important;
    padding-bottom: 0 !important;
    font-size: 0.78rem !important;
    line-height: 1 !important;
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
    padding: 0.6rem 0.85rem !important;
    margin-bottom: 0.35rem !important;
    min-height: 2.4rem !important;
    transition: background .12s;
    display: flex !important;
    align-items: flex-start !important;
}}
.stCheckbox:hover {{ background: var(--blue-pale); border-color: #b3c8f0; }}
.stCheckbox label {{
    font-size: 0.84rem !important;
    font-weight: 500 !important;
    text-transform: none !important;
    letter-spacing: 0 !important;
    color: var(--text) !important;
    line-height: 1.5 !important;
    white-space: normal !important;
    word-break: break-word !important;
}}

/* ── Monitor block ── */
.monitor-block {{
    border: 1px solid var(--border);
    border-radius: 12px;
    overflow: hidden;
    margin-bottom: 0.85rem;
    background: var(--card);
    box-shadow: var(--shadow);
}}
.monitor-block-header {{
    display: flex;
    align-items: center;
    gap: 9px;
    padding: 0.65rem 0.9rem;
    background: #f0f4fa;
    border-bottom: 1px solid var(--border);
}}
.monitor-block-icon {{
    width: 28px;
    height: 28px;
    border-radius: 7px;
    background: var(--blue-pale);
    display: flex;
    align-items: center;
    justify-content: center;
    font-size: 0.9rem;
    flex-shrink: 0;
}}
.monitor-block-title {{
    font-size: 0.83rem;
    font-weight: 700;
    color: var(--navy);
}}
.monitor-block-sub {{
    font-size: 0.73rem;
    color: var(--muted);
    margin-top: 1px;
}}

/* ── Adapter chip panel ── */
.adapter-chip-panel {{
    margin: 0 0.9rem 0.65rem 2.4rem;
    background: #f7f9fd;
    border: 1px solid #dde5f0;
    border-radius: 8px;
    overflow: hidden;
}}
.adapter-chip-header {{
    display: flex;
    align-items: center;
    gap: 5px;
    padding: 0.32rem 0.7rem;
    border-bottom: 1px solid #e4ebf5;
    background: #eef2fb;
}}
.adapter-chip-dot {{
    width: 5px; height: 5px;
    border-radius: 50%;
    background: var(--muted);
    flex-shrink: 0;
}}
.adapter-chip-label {{
    font-size: 0.65rem;
    font-weight: 700;
    color: var(--muted);
    text-transform: uppercase;
    letter-spacing: 0.07em;
}}
.adapter-chips-wrap {{
    display: flex;
    flex-wrap: wrap;
    gap: 5px;
    padding: 0.5rem 0.7rem;
}}
.adapter-chip {{
    font-size: 0.73rem;
    padding: 3px 10px;
    border-radius: 100px;
    border: 1px solid var(--border);
    cursor: pointer;
    color: var(--muted);
    background: #fff;
    font-family: 'DM Sans', sans-serif;
    font-weight: 500;
    transition: all 0.12s;
    white-space: nowrap;
}}
.adapter-chip:hover {{ border-color: var(--blue); color: var(--blue); background: var(--blue-pale); }}
.adapter-chip-none {{
    font-size: 0.73rem;
    padding: 3px 10px;
    border-radius: 100px;
    border: 1px dashed var(--border);
    cursor: pointer;
    color: var(--muted);
    background: transparent;
    font-family: 'DM Sans', sans-serif;
    font-weight: 500;
    transition: all 0.12s;
    white-space: nowrap;
}}
.adapter-chip-none:hover {{ border-color: var(--muted); background: #f0f4fa; }}
.adapter-chip.chip-selected {{
    background: var(--blue-pale); color: var(--blue); border-color: var(--blue); font-weight: 600;
}}
.adapter-chip-none.chip-selected {{
    background: #f0f4fa; color: var(--muted); border-style: solid; border-color: var(--muted); font-weight: 600;
}}

/* ── Charger badge ── */
.charger-badge {{
    display: inline-flex;
    align-items: center;
    gap: 0.4rem;
    font-size: 0.75rem;
    font-weight: 600;
    color: var(--green);
    background: var(--green-lt);
    border: 1px solid #c3e6c0;
    border-radius: 6px;
    padding: 0.3rem 0.65rem;
    margin-bottom: 0.75rem;
}}

/* ── Sequence badge ── */
.seq-badge {{
    display: inline-block;
    background: var(--blue-pale);
    color: var(--blue);
    font-size: 0.67rem;
    font-weight: 700;
    border-radius: 4px;
    padding: 1px 5px;
    margin-right: 5px;
    vertical-align: middle;
}}

/* ── Alerts ── */
.stAlert {{
    border-radius: 8px !important;
    font-family: 'DM Sans', sans-serif !important;
    font-size: 0.82rem !important;
}}

/* ── Metrics ── */
[data-testid="stMetric"] {{
    background: var(--card);
    border: 1px solid var(--border);
    border-radius: var(--radius);
    padding: 0.9rem 1.1rem;
    box-shadow: var(--shadow);
}}
[data-testid="stMetricValue"] {{
    font-family: 'DM Sans', sans-serif !important;
    font-weight: 700 !important;
    font-size: 1.6rem !important;
    color: var(--navy) !important;
}}
[data-testid="stMetricLabel"] {{
    font-family: 'DM Sans', sans-serif !important;
    font-size: 0.7rem !important;
    font-weight: 600 !important;
    text-transform: uppercase !important;
    letter-spacing: 0.06em !important;
    color: var(--muted) !important;
}}

hr {{ border-color: var(--border) !important; margin: 0.9rem 0 !important; }}

.stCaption, small {{
    font-family: 'DM Sans', sans-serif !important;
    color: var(--muted) !important;
    font-size: 0.75rem !important;
}}

/* ── Info hint ── */
.info-hint {{
    font-size: 0.77rem;
    color: var(--blue);
    background: var(--blue-pale);
    border-left: 3px solid var(--blue);
    border-radius: 0 6px 6px 0;
    padding: 0.5rem 0.7rem;
    margin-bottom: 0.8rem;
    font-weight: 400;
    line-height: 1.6;
}}
.info-hint strong {{ font-weight: 700; color: var(--navy); }}

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
    font-size: 0.7rem;
    margin-top: 1.8rem;
    padding-top: 0.9rem;
    border-top: 1px solid var(--border);
}}

/* ── Preview table ── */
.preview-table {{
    width: 100%;
    border-collapse: collapse;
    font-size: 0.79rem;
    margin-top: 0.5rem;
}}
.preview-table th {{
    background: var(--navy);
    color: #fff;
    font-weight: 600;
    padding: 0.4rem 0.65rem;
    text-align: left;
    font-size: 0.7rem;
    letter-spacing: 0.04em;
    text-transform: uppercase;
}}
.preview-table td {{
    padding: 0.38rem 0.65rem;
    border-bottom: 1px solid var(--border);
    color: var(--text);
}}
.preview-table tr:nth-child(even) td {{ background: #f7f9fd; }}
.preview-table tr:hover td {{ background: var(--blue-pale); }}

/* ── Remarks field ── */
.remarks-wrap {{
    background: #fffdf0;
    border: 1.5px solid #ffe082;
    border-radius: 8px;
    padding: 0.8rem 0.95rem 0.55rem;
    margin-top: 0.3rem;
    margin-bottom: 0.5rem;
}}
.remarks-label {{
    font-size: 0.7rem;
    font-weight: 700;
    color: #b07d00;
    text-transform: uppercase;
    letter-spacing: 0.07em;
    margin-bottom: 0.3rem;
}}
.remarks-hint {{
    font-size: 0.73rem;
    color: var(--muted);
    margin-bottom: 0.45rem;
    line-height: 1.45;
}}

/* ── User manager label spacer ── */
.user-mgr-btn-spacer {{
    height: 2.35rem;
    display: block;
}}

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

ITEM_SEQUENCE_ORDER = {
    "laptop": 1, "computer": 1, "desktop": 1,
    "charger": 2, "monitor": 3,
    "hdmi": 5, "keyboard": 6, "mouse": 7,
    "headset": 8, "headphone": 8, "usb": 9,
    "usb peripheral": 9, "vga": 10, "dvi": 11,
    "displayport": 12, "type-c": 13, "type c": 13,
    "adapter": 13, "converter": 13, "docking station": 14,
    "docking": 14, "webcam": 16, "speaker": 17,
    "ethernet adapter": 18,
}

MONITOR_PERIPHERALS = [
    ("Monitor Power Cable",  3.5,  None),
    ("HDMI Cable",           5.0,  ["HDMI to VGA", "HDMI to DisplayPort", "HDMI to DVI", "HDMI to USB-C"]),
    ("VGA Cable",            10.0, ["VGA to HDMI", "VGA to DisplayPort", "VGA to DVI"]),
    ("DVI Cable",            11.0, ["DVI to HDMI", "DVI to VGA", "DVI to DisplayPort"]),
    ("DisplayPort Cable",    12.0, ["DisplayPort to HDMI", "DisplayPort to VGA", "DisplayPort to DVI", "DisplayPort to USB-C"]),
    ("USB-C Cable",          13.0, ["USB-C to HDMI", "USB-C to VGA", "USB-C to DisplayPort", "USB-C to DVI"]),
]

_POSITION_SP_SUFFIXES = (":position", ":jobtitle", ":job title", ":title", ":designation")
_POSITION_KEYWORDS    = ("position", "job title", "jobtitle", "designation", "title")

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_FONT_NAME  = "Calibri"
_FONT_SIZE  = "20"
_TABLE_SIZE = "18"

_JS_ONCLICK_REGEX = r"/'([^']+)'\s*\)$/"

# ─────────────────────────────────────────────
# DATA HELPERS
# ─────────────────────────────────────────────

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
    if "monitor power" in text_lower or ("power cable" in text_lower and "monitor" in text_lower):
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
    monitor_cable_assignments: list[dict],
    shared_remarks: str = "",
) -> list[dict]:
    rows: list[dict] = []
    for _, row in assets_df.iterrows():
        c         = col_map
        content   = safe_str(row.get(c.get("content_type") or "", ""))
        brand     = safe_str(row.get(c.get("brand")        or "", ""))
        model     = safe_str(row.get(c.get("model")        or "", ""))
        serial    = safe_str(row.get(c.get("serial")       or "", ""))
        asset_tag = safe_str(row.get(c.get("asset_tag")    or "", ""))
        equipment = _build_equipment_label(brand, model, content, asset_tag)
        seq_key   = _get_sequence_key(equipment)
        rows.append({
            "equipment":   equipment,
            "serial":      serial,
            "asset_tag":   asset_tag,
            "remarks":     "",
            "_seq_key":    seq_key,
            "_is_monitor": _is_monitor(equipment),
        })

    rows.append({
        "equipment": "Charger", "serial": "", "asset_tag": "",
        "remarks": "", "_seq_key": 2.0, "_is_monitor": False,
    })
    rows.sort(key=lambda r: r["_seq_key"])

    from collections import defaultdict
    cables_by_monitor: dict[int, list] = defaultdict(list)
    for assignment in monitor_cable_assignments:
        cables_by_monitor[assignment["monitor_idx"]].append(
            (assignment["cable_seq"], assignment["cable_name"], assignment.get("adapter_name", ""))
        )
    for idx in cables_by_monitor:
        cables_by_monitor[idx].sort(key=lambda x: x[0])

    result: list[dict] = []
    monitor_counter = 0
    for row in rows:
        result.append(row)
        if row.get("_is_monitor"):
            for cable_seq, cable_name, adapter_name in cables_by_monitor.get(monitor_counter, []):
                result.append({
                    "equipment": cable_name, "serial": "", "asset_tag": "", "remarks": "",
                    "_seq_key": cable_seq, "_is_monitor": False,
                })
                if adapter_name and adapter_name != "No Adapter Needed":
                    result.append({
                        "equipment": adapter_name, "serial": "", "asset_tag": "", "remarks": "",
                        "_seq_key": cable_seq + 0.1, "_is_monitor": False,
                    })
            monitor_counter += 1

    final = [{k: v for k, v in r.items() if not k.startswith("_")} for r in result]
    if final and shared_remarks:
        final[0]["remarks"] = shared_remarks
    return final

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
    if exact_matches  == n_q:                       return (1, coverage)
    if substr_matches == n_q:                       return (2, substr_cov)
    if exact_matches  >= max(1, round(n_q * 0.6)):  return (3, coverage)
    if substr_matches >= max(1, round(n_q * 0.6)):  return (4, substr_cov)
    if exact_matches  >= 1:                         return (5, coverage)
    if substr_matches >= 1 or full_substr:          return (6, max(substr_cov, 0.1))
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
_ROW_HEIGHT_DXA = "220"
_CELL_PAD_TOP   = "40"
_CELL_PAD_BTM   = "40"
_CELL_PAD_LEFT  = "80"
_CELL_PAD_RIGHT = "80"


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


def _set_sdt_value(body, tag_val: str, new_text: str, bold: bool = True, size: str = _FONT_SIZE) -> bool:
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
        jc  = etree.SubElement(pPr, f"{{{W}}}jc"); jc.set(f"{{{W}}}val", "left")
        sp  = etree.SubElement(pPr, f"{{{W}}}spacing")
        sp.set(f"{{{W}}}before", "0"); sp.set(f"{{{W}}}after", "0")
        sp.set(f"{{{W}}}line",   "240"); sp.set(f"{{{W}}}lineRule", "auto")
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
    jc = etree.SubElement(pPr, f"{{{W}}}jc"); jc.set(f"{{{W}}}val", "left")
    sp = etree.SubElement(pPr, f"{{{W}}}spacing")
    sp.set(f"{{{W}}}before", "0"); sp.set(f"{{{W}}}after", "0")
    sp.set(f"{{{W}}}line",   "240"); sp.set(f"{{{W}}}lineRule", "auto")
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
    trH.set(f"{{{W}}}val", _ROW_HEIGHT_DXA)
    trH.set(f"{{{W}}}hRule", "atLeast")
    for tc in row_el.iter(f"{{{W}}}tc"):
        tcPr = tc.find(f"{{{W}}}tcPr")
        if tcPr is None:
            tcPr = etree.SubElement(tc, f"{{{W}}}tcPr")
            tc.insert(0, tcPr)
        tcMar = tcPr.find(f"{{{W}}}tcMar")
        if tcMar is None:
            tcMar = etree.SubElement(tcPr, f"{{{W}}}tcMar")
        for side, val in [("top", _CELL_PAD_TOP), ("bottom", _CELL_PAD_BTM),
                          ("left", _CELL_PAD_LEFT), ("right", _CELL_PAD_RIGHT)]:
            el = tcMar.find(f"{{{W}}}{side}")
            if el is None:
                el = etree.SubElement(tcMar, f"{{{W}}}{side}")
            el.set(f"{{{W}}}w", val); el.set(f"{{{W}}}type", "dxa")
        for va in tcPr.findall(f"{{{W}}}vAlign"):
            tcPr.remove(va)
        vAlign = etree.SubElement(tcPr, f"{{{W}}}vAlign"); vAlign.set(f"{{{W}}}val", "center")
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
    for side, max_val in [("top", 720), ("bottom", 720), ("left", 720), ("right", 720)]:
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
            trPr = etree.SubElement(tr, f"{{{W}}}trPr"); tr.insert(0, trPr)
        for trH in trPr.findall(f"{{{W}}}trHeight"):
            trPr.remove(trH)
        trH = etree.SubElement(trPr, f"{{{W}}}trHeight")
        trH.set(f"{{{W}}}val", HEADER_ROW_HEIGHT); trH.set(f"{{{W}}}hRule", "exact")
        for tc in tr.findall(f"{{{W}}}tc"):
            has_sdt = bool(tc.find(f".//{{{W}}}sdt"))
            tcPr = tc.find(f"{{{W}}}tcPr")
            if tcPr is None:
                tcPr = etree.SubElement(tc, f"{{{W}}}tcPr"); tc.insert(0, tcPr)
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
            vAlign = etree.SubElement(tcPr, f"{{{W}}}vAlign"); vAlign.set(f"{{{W}}}val", "top")
            for p in tc.iter(f"{{{W}}}p"):
                pPr = p.find(f"{{{W}}}pPr")
                if pPr is None:
                    pPr = etree.SubElement(p, f"{{{W}}}pPr"); p.insert(0, pPr)
                for sp in pPr.findall(f"{{{W}}}spacing"):
                    pPr.remove(sp)
                sp_el = etree.SubElement(pPr, f"{{{W}}}spacing")
                sp_el.set(f"{{{W}}}before", "0"); sp_el.set(f"{{{W}}}after", "0")
                sp_el.set(f"{{{W}}}line",   "240"); sp_el.set(f"{{{W}}}lineRule", "auto")


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
            "Place it inside the src/ folder next to this script."
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
# PREPARED-BY USER MANAGEMENT UI
# ─────────────────────────────────────────────

def render_prepared_by_section() -> str:
    st.markdown("""
    <div class="preparedby-card">
      <div class="preparedby-icon">&#x270F;&#xFE0F;</div>
      <div>
        <div class="preparedby-label">Prepared By</div>
        <div class="preparedby-hint">IT staff member preparing this form — auto-filled in the document.</div>
      </div>
    </div>
    """, unsafe_allow_html=True)

    options = get_prepared_by_options()
    col_sel, col_mgr = st.columns([3, 1])
    with col_sel:
        prepared_by = st.selectbox(
            "Select IT staff", options=options, index=0,
            label_visibility="collapsed", key="prepared_by_select",
        )
    with col_mgr:
        if "show_user_mgr" not in st.session_state:
            st.session_state["show_user_mgr"] = False
        mgr_label = "▲ Manage" if st.session_state["show_user_mgr"] else "▼ Manage"
        if st.button(mgr_label, key="btn_toggle_user_mgr", use_container_width=True):
            st.session_state["show_user_mgr"] = not st.session_state["show_user_mgr"]
            st.rerun()

    if st.session_state.get("show_user_mgr"):
        _render_user_manager(options)

    return prepared_by


def _render_user_manager(options: list[str]):
    st.markdown("""
    <div class="user-reg-panel">
      <div class="user-reg-title">&#x1F464; Manage IT Staff List</div>
    </div>
    """, unsafe_allow_html=True)

    chips_html = ""
    for u in options:
        is_default = u in DEFAULT_PREPARED_BY
        chip_cls   = "user-chip default-chip" if is_default else "user-chip"
        lock_icon  = " 🔒" if is_default else ""
        chips_html += f'<span class="{chip_cls}">{u}{lock_icon}</span>'
    st.markdown(
        f'<div style="margin-bottom:0.55rem">'
        f'<div style="font-size:0.7rem;font-weight:700;color:#5a6e8a;text-transform:uppercase;'
        f'letter-spacing:.06em;margin-bottom:6px;">Current staff</div>'
        f'<div class="user-chip-wrap">{chips_html}</div>'
        f'</div>',
        unsafe_allow_html=True,
    )
    st.markdown("<hr>", unsafe_allow_html=True)

    # ── Add New Staff ─────────────────────────────────────────────────────────
    add_col, add_btn_col = st.columns([3, 1])
    with add_col:
        new_name = st.text_input(
            "Add New Staff",
            placeholder="Full name, e.g. Maria Santos",
            key="new_user_name_input",
        )
    with add_btn_col:
        st.write("")
        st.markdown('<div class="btn-outline-blue">', unsafe_allow_html=True)
        if st.button("Add", key="btn_add_user", use_container_width=True):
            ok, msg = add_prepared_by_user(new_name)
            if ok:
                st.success(msg)
                st.session_state.pop("new_user_name_input", None)
                st.rerun()
            else:
                st.error(msg)
        st.markdown('</div>', unsafe_allow_html=True)

    # ── Remove Staff ──────────────────────────────────────────────────────────
    removable = [u for u in options if u not in DEFAULT_PREPARED_BY]
    if removable:
        st.markdown("<hr>", unsafe_allow_html=True)
        rm_col, rm_btn_col = st.columns([3, 1])
        with rm_col:
            to_remove = st.selectbox(
                "Remove Staff",
                options=removable,
                key="remove_user_select",
            )
        with rm_btn_col:
            st.write("")
            st.markdown('<div class="btn-danger">', unsafe_allow_html=True)
            if st.button("Remove", key="btn_remove_user", use_container_width=True):
                ok, msg = remove_prepared_by_user(to_remove)
                if ok:
                    st.success(msg); st.rerun()
                else:
                    st.error(msg)
            st.markdown('</div>', unsafe_allow_html=True)
    else:
        st.caption("Only default staff shown — add a custom name above to enable removal.")

# ─────────────────────────────────────────────
# CSV UPLOAD UI
# ─────────────────────────────────────────────

def render_csv_upload() -> tuple[pd.DataFrame | None, str]:
    uploaded = st.file_uploader(
        "Upload asset list",
        type=["csv", "xlsx", "xls"],
        label_visibility="collapsed",
        key="csv_upload",
    )

    if uploaded is None:
        st.markdown(
            '<div class="info-hint">'
            '<strong>Tip:</strong> Export your asset list from SharePoint or Excel as a '
            '<strong>.csv</strong> or <strong>.xlsx</strong> file, then upload it here.'
            '</div>',
            unsafe_allow_html=True,
        )
        return None, ""

    try:
        if uploaded.name.lower().endswith((".xlsx", ".xls")):
            df = pd.read_excel(uploaded)
            label = f"{uploaded.name} · {len(df):,} rows"
        else:
            raw = uploaded.read()
            for enc in ("utf-8-sig", "utf-8", "latin-1", "cp1252"):
                try:
                    df = pd.read_csv(io.BytesIO(raw), encoding=enc)
                    break
                except UnicodeDecodeError:
                    continue
            else:
                st.error("Could not decode the CSV file. Try saving it as UTF-8.")
                return None, ""
            label = f"{uploaded.name} · {len(df):,} rows"
    except Exception as e:
        st.error(f"Failed to read file: {e}")
        return None, ""

    if df.empty:
        st.warning("The uploaded file appears to be empty.")
        return None, ""

    st.markdown(
        f'<div class="csv-source-bar">✅&nbsp; <span>{label}</span></div>',
        unsafe_allow_html=True,
    )
    return df, label

# ─────────────────────────────────────────────
# FORM TYPE SELECTOR
# ─────────────────────────────────────────────

def render_form_type_selector() -> str:
    if "form_type" not in st.session_state:
        st.session_state["form_type"] = "wfh"
    current = st.session_state["form_type"]

    wfh_tpl    = _find_template(WFH_TEMPLATE_NAME)
    onsite_tpl = _find_template(ONSITE_TEMPLATE_NAME)

    col_wfh, col_onsite = st.columns(2)
    with col_wfh:
        wfh_css = "wfh-active" if current == "wfh" else "wfh-inactive"
        st.markdown(f'<div class="form-toggle-btn {wfh_css}">', unsafe_allow_html=True)
        wfh_label = "Work From Home" + ("" if wfh_tpl else "  (template missing)")
        if st.button(wfh_label, key="btn_wfh", use_container_width=True):
            if current != "wfh":
                st.session_state["form_type"] = "wfh"
                for k in ("docx", "form_name", "form_client", "form_date",
                          "prepared_by", "form_type_used"):
                    st.session_state.pop(k, None)
                st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)

    with col_onsite:
        onsite_css = "onsite-active" if current == "onsite" else "onsite-inactive"
        st.markdown(f'<div class="form-toggle-btn {onsite_css}">', unsafe_allow_html=True)
        onsite_label = "Work On Site" + ("" if onsite_tpl else "  (template missing)")
        if st.button(onsite_label, key="btn_onsite", use_container_width=True):
            if current != "onsite":
                st.session_state["form_type"] = "onsite"
                for k in ("docx", "form_name", "form_client", "form_date",
                          "prepared_by", "form_type_used"):
                    st.session_state.pop(k, None)
                st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)

    return st.session_state["form_type"]

# ─────────────────────────────────────────────
# MONITOR CABLE + ADAPTER RENDERER
# ─────────────────────────────────────────────

def render_monitor_cable_block(mon_idx: int, mon_label: str) -> list[dict]:
    assignments: list[dict] = []

    st.markdown(f"""
    <div class="monitor-block">
      <div class="monitor-block-header">
        <div class="monitor-block-icon">&#x1F5A5;</div>
        <div>
          <div class="monitor-block-title">Monitor {mon_idx + 1}</div>
          <div class="monitor-block-sub">{mon_label}</div>
        </div>
      </div>
    </div>
    """, unsafe_allow_html=True)

    with st.container():
        st.markdown(
            '<div style="margin-top:-0.6rem;border:1px solid #d0dce8;'
            'border-top:none;border-radius:0 0 12px 12px;'
            'background:#ffffff;overflow:hidden;margin-bottom:0.85rem;">',
            unsafe_allow_html=True,
        )

        for cable_name, cable_seq, adapter_options in MONITOR_PERIPHERALS:
            ckey = f"cable_mon{mon_idx}_{_safe_periph_key(cable_name)}"
            if ckey not in st.session_state:
                st.session_state[ckey] = False

            st.markdown('<div style="border-bottom:1px solid #edf1f7;">', unsafe_allow_html=True)
            is_checked = st.checkbox(cable_name, key=ckey)
            st.markdown('</div>', unsafe_allow_html=True)

            chosen_adapter = ""
            if is_checked and adapter_options:
                akey    = f"adapter_mon{mon_idx}_{_safe_periph_key(cable_name)}"
                sel_key = f"sel_{akey}"
                if sel_key not in st.session_state:
                    st.session_state[sel_key] = "none"
                current_adapter = st.session_state.get(sel_key, "none")

                chips_html = ""
                none_cls = "adapter-chip-none chip-selected" if current_adapter == "none" else "adapter-chip-none"
                chips_html += (
                    f'<button class="{none_cls}" '
                    f'onclick="setAdapter(\'{akey}\', \'none\')">No adapter</button>'
                )
                for opt in adapter_options:
                    opt_key = opt.replace("'", "\\'")
                    cls = "adapter-chip chip-selected" if current_adapter == opt else "adapter-chip"
                    chips_html += (
                        f'<button class="{cls}" '
                        f'onclick="setAdapter(\'{akey}\', \'{opt_key}\')">{opt}</button>'
                    )

                st.markdown(f"""
                <div class="adapter-chip-panel" id="panel_{akey}">
                  <div class="adapter-chip-header">
                    <div class="adapter-chip-dot"></div>
                    <span class="adapter-chip-label">Adapter needed?</span>
                  </div>
                  <div class="adapter-chips-wrap" id="chips_{akey}">
                    {chips_html}
                  </div>
                </div>
                <script>
                function setAdapter(key, val) {{
                    const wrap = document.getElementById('chips_' + key);
                    if (!wrap) return;
                    wrap.querySelectorAll('button').forEach(btn => btn.classList.remove('chip-selected'));
                    event.target.classList.add('chip-selected');
                    sessionStorage.setItem('adapter_' + key, val);
                }}
                (function() {{
                    const stored = sessionStorage.getItem('adapter_{akey}');
                    if (!stored) return;
                    const wrap = document.getElementById('chips_{akey}');
                    if (!wrap) return;
                    wrap.querySelectorAll('button').forEach(btn => {{
                        btn.classList.remove('chip-selected');
                        const m = btn.getAttribute('onclick').match({_JS_ONCLICK_REGEX});
                        if (m && m[1] === stored) btn.classList.add('chip-selected');
                    }});
                }})();
                </script>
                """, unsafe_allow_html=True)

                st.markdown('<div style="display:none;">', unsafe_allow_html=True)
                adapter_opts_full = ["none"] + list(adapter_options)
                chosen_raw = st.selectbox(
                    f"_adapter_{akey}", options=adapter_opts_full,
                    key=sel_key, label_visibility="collapsed",
                )
                chosen_adapter = "" if chosen_raw == "none" else chosen_raw
                st.markdown('</div>', unsafe_allow_html=True)

            if is_checked:
                assignments.append({
                    "monitor_label": mon_label,
                    "monitor_idx":   mon_idx,
                    "cable_name":    cable_name,
                    "cable_seq":     cable_seq,
                    "adapter_name":  chosen_adapter,
                })

        st.markdown('</div>', unsafe_allow_html=True)

    return assignments

# ─────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────

def main():
    render_header()

    # ── Form Type ─────────────────────────────────────────────────────────────
    st.markdown(
        '<div style="font-size:0.7rem;font-weight:700;color:#5a6e8a;'
        'text-transform:uppercase;letter-spacing:0.07em;margin-bottom:0.3rem;">'
        'Form Type</div>',
        unsafe_allow_html=True,
    )
    form_type = render_form_type_selector()

    if form_type == "wfh":
        tpl_path   = _find_template(WFH_TEMPLATE_NAME)
        type_label = "Work From Home"
        type_color = "#1565c0"
        type_bg    = "#e8f0fc"
    else:
        tpl_path   = _find_template(ONSITE_TEMPLATE_NAME)
        type_label = "Work On Site"
        type_color = "#00796b"
        type_bg    = "#e0f2f1"

    if tpl_path is None:
        st.error(
            f"Template not found for **{type_label}**. "
            "Place the `.dotx` file inside the **src/** folder and refresh."
        )
    else:
        st.markdown(
            f'<div style="font-size:0.75rem;color:{type_color};background:{type_bg};'
            f'border-left:3px solid {type_color};border-radius:0 6px 6px 0;'
            f'padding:0.38rem 0.7rem;margin:0.5rem 0 0.8rem;">'
            f'Using template: <strong>{type_label}</strong>'
            f'</div>',
            unsafe_allow_html=True,
        )

    st.markdown("<hr>", unsafe_allow_html=True)

    # ── Prepared By ───────────────────────────────────────────────────────────
    prepared_by = render_prepared_by_section()
    st.markdown("<div style='height:0.4rem'></div>", unsafe_allow_html=True)

    # ── Step 1: Upload CSV ────────────────────────────────────────────────────
    step_open(1, "Upload Asset List",
              "Upload a CSV or Excel file exported from SharePoint or your asset tracker.")
    df, csv_label = render_csv_upload()
    step_close()

    if df is None:
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
        st.error(
            'Could not find a "Current User" column. '
            'Make sure your file has a column named "Current User" (or similar).'
        )
        return

    if position_col:
        st.caption(f"{len(df):,} records · Position column: {position_col}")
    else:
        st.caption(f"{len(df):,} records · No position column detected")

    # ── Step 2: Search ────────────────────────────────────────────────────────
    step_open(2, "Find Employee",
              "Type any part of a name — partial words and any order are supported.")
    search = st.text_input(
        "Search", placeholder="e.g. Juan Dela Cruz or just 'juan'...",
        label_visibility="collapsed",
    )
    step_close()

    if not search.strip():
        st.info("Type an employee name above to continue.")
        return

    results, df_by_name = smart_search(df, user_col, search)

    if not results:
        st.warning(f'No records found for "{search.strip()}". Try a shorter or different name.')
        return

    top_results    = [r for r in results if r["tier"] <= 1]
    strong_results = [r for r in results if 2 <= r["tier"] <= 3]
    weak_results   = [r for r in results if r["tier"] >= 4]

    def make_label(item):
        count = len(df_by_name.get(item["name"], pd.DataFrame()))
        return f"{item['name']}  [{item['label']}]  ({count} asset{'s' if count != 1 else ''})"

    ordered       = top_results + strong_results + weak_results
    options       = [make_label(r) for r in ordered]
    name_map      = {make_label(r): r["name"] for r in ordered}

    if len(options) == 1 and ordered[0]["tier"] <= 1:
        chosen_display = options[0]
        st.success(f"Matched: **{ordered[0]['name']}**")
    else:
        st.caption(f"{len(results)} result(s)")
        chosen_display = st.selectbox(
            "Select employee", options,
            help="Ranked by match quality — exact matches first.",
        )

    chosen_name = name_map[chosen_display]
    df_filtered = df_by_name.get(chosen_name, pd.DataFrame())

    if df_filtered.empty:
        st.warning("No assets found for the selected employee.")
        return

    chosen_result = next((r for r in results if r["name"] == chosen_name), None)
    if chosen_result and chosen_result["tier"] >= 4:
        st.info("Closest result for your search. Not right? Select a different name above.")
    else:
        st.success(f"**{len(df_filtered)}** asset(s) found for {chosen_name}.")

    # ── Step 3: Select Assets ─────────────────────────────────────────────────
    step_open(3, "Select Assets",
              "Choose the items from the asset list to include on the form.")

    sel_key = f"sel_{chosen_name.lower().replace(' ', '_')}"
    if sel_key not in st.session_state:
        st.session_state[sel_key] = {idx: False for idx in df_filtered.index}
    for idx in df_filtered.index:
        if idx not in st.session_state[sel_key]:
            st.session_state[sel_key][idx] = False

    col_a, col_b, _ = st.columns([1, 1, 5])
    with col_a:
        st.markdown('<div class="btn-outline-blue">', unsafe_allow_html=True)
        if st.button("Select All", key="sp_select_all"):
            for idx in df_filtered.index:
                st.session_state[sel_key][idx] = True
            st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)
    with col_b:
        st.markdown('<div class="btn-outline-blue">', unsafe_allow_html=True)
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

    # ── Step 4: Cables & Adapters ─────────────────────────────────────────────
    step_open(4, "Monitor Cables & Adapters",
              "A charger is automatically included. Select cables and adapters for each monitor.")
    st.markdown('<div class="charger-badge">Charger automatically included</div>', unsafe_allow_html=True)

    monitor_cable_assignments: list[dict] = []
    if selected_monitors:
        st.caption("Tick which cables came with each monitor, then pick an adapter if needed.")
        for mon in selected_monitors:
            assignments = render_monitor_cable_block(mon["idx"], mon["label"])
            monitor_cable_assignments.extend(assignments)
    elif not df_selected.empty:
        st.markdown(
            '<div style="font-size:0.8rem;color:#5a6e8a;padding:0.4rem 0;">'
            'No monitors in the selected assets — cable section not applicable.'
            '</div>',
            unsafe_allow_html=True,
        )

    step_close()

    # ── Remarks ───────────────────────────────────────────────────────────────
    st.markdown("""
    <div class="remarks-wrap">
      <div class="remarks-label">Remarks</div>
      <div class="remarks-hint">Appears on the first row of the form, covering the entire equipment list.</div>
    </div>
    """, unsafe_allow_html=True)
    shared_remarks = st.text_input(
        "Remarks",
        placeholder="e.g. All items in good condition, for WFH use",
        label_visibility="collapsed",
        key="shared_remarks",
    )

    # ── Build sorted rows ─────────────────────────────────────────────────────
    sorted_rows = sort_assets_by_sequence(
        df_selected, col_map, monitor_cable_assignments, shared_remarks,
    )

    if not sorted_rows:
        st.info("Select at least one asset to continue.")
        return

    cable_count   = len(monitor_cable_assignments)
    adapter_count = sum(
        1 for a in monitor_cable_assignments
        if a.get("adapter_name") and a["adapter_name"] not in ("", "none", "No Adapter Needed")
    )
    summary_parts = [f"**{len(checked)}** asset(s)", "**1** charger"]
    if cable_count:
        summary_parts.append(f"**{cable_count}** cable(s)")
    if adapter_count:
        summary_parts.append(f"**{adapter_count}** adapter(s)")
    st.caption(" + ".join(summary_parts) + f" = **{len(sorted_rows)}** total items on form")

    with st.expander("Preview — Form Order", expanded=False):
        st.markdown(
            '<table class="preview-table"><thead><tr>'
            '<th>#</th><th>Equipment</th><th>Serial</th><th>Asset Tag</th><th>Remarks</th>'
            '</tr></thead><tbody>' +
            "".join(
                f'<tr><td><span class="seq-badge">{i+1}</span></td>'
                f'<td>{row["equipment"]}</td><td>{row["serial"]}</td>'
                f'<td>{row["asset_tag"]}</td><td>{row["remarks"]}</td></tr>'
                for i, row in enumerate(sorted_rows)
            ) +
            "</tbody></table>",
            unsafe_allow_html=True,
        )

    # ── Step 5: Form Details ──────────────────────────────────────────────────
    step_open(5, "Form Details",
              "Header fields for the generated document — auto-filled from the asset data.")

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
        f'<div class="info-hint">'
        f'<strong>Prepared by:</strong> {prepared_by} &nbsp;·&nbsp; '
        f'<strong>Employee:</strong> {chosen_name} &nbsp;·&nbsp; '
        f'<strong>Items:</strong> {len(sorted_rows)} &nbsp;·&nbsp; '
        f'<strong>Form:</strong> {type_label}'
        f'</div>',
        unsafe_allow_html=True,
    )

    if tpl_path is None:
        st.error(
            f"Cannot generate — **{type_label}** template not found. "
            "Place the .dotx file in the **src/** folder and refresh."
        )
        step_close()
        return

    if st.button("Generate Word Document", use_container_width=True, type="primary"):
        with st.spinner("Filling the form..."):
            try:
                docx_bytes = fill_template(
                    sorted_rows,
                    form_name, form_client, form_position, form_date_str,
                    prepared_by=prepared_by,
                    form_type=form_type,
                )
                st.session_state["docx"]           = docx_bytes
                st.session_state["form_name"]      = form_name
                st.session_state["form_client"]    = form_client
                st.session_state["form_date"]      = form_date
                st.session_state["prepared_by"]    = prepared_by
                st.session_state["form_type_used"] = form_type
                st.success("Document ready — click Download below.")
            except Exception as e:
                st.error(f"Error: {e}")

    if "docx" in st.session_state:
        _fname      = st.session_state.get("form_name")      or "Employee"
        _fclient    = st.session_state.get("form_client")    or ""
        _fdate      = st.session_state.get("form_date")      or datetime.now()
        _ftype_used = st.session_state.get("form_type_used") or "wfh"
        _client_p   = f" ({_fclient})" if _fclient else ""
        _day        = str(int(_fdate.strftime("%d")))
        _date_p     = _day + _fdate.strftime(" %B %Y")
        _type_p     = "WFH" if _ftype_used == "wfh" else "On Site"
        _filename   = (
            f"Employee Copy - Equipment Accountability Form ({_type_p})"
            f" - {_fname}{_client_p} _ {_date_p}.docx"
        )
        st.download_button(
            "⬇️ Download Word (.docx)",
            data=st.session_state["docx"],
            file_name=_filename,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
        )

    step_close()

    # ── Summary metrics ───────────────────────────────────────────────────────
    st.markdown("<div style='height:.3rem'></div>", unsafe_allow_html=True)
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total Records", f"{len(df):,}")
    c2.metric("Matched",       f"{len(df_filtered):,}")
    c3.metric("Selected",      f"{len(df_selected):,}")
    c4.metric("On This Form",  f"{len(sorted_rows):,}")

    st.markdown(
        '<div class="bmg-footer">BMG Outsourcing, Inc. &nbsp;·&nbsp; Asset Accountability System</div>',
        unsafe_allow_html=True,
    )


if __name__ == "__main__":
    main()
