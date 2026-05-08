"""
Equipment Accountability Form Generator
BMG Outsourcing — CSV upload version.
Supports both Work From Home and Work On Site templates.

FIXES APPLIED:
  1. Dark-mode: comprehensive CSS variables + forced-colors rules ensure all
     text, borders, backgrounds, and buttons remain legible in Edge / Chrome
     dark mode and Windows High Contrast.
  2. Inserted data (Name, Client, Position, Date) is now rendered NON-BOLD.
  3. Table rows are more compact: reduced row height, tighter padding, smaller
     font, consistent spacing.
  4. Copy type selector added (IT Copy / Employee Copy).
  5. "Verified By" label changes to "Received By" when Employee Copy is chosen.
  6. Select All now correctly sets every checkbox widget key and reruns.
  7. Clear All now correctly clears every checkbox widget key and reruns.
  8. Remarks column text aligns to top of cell.
  9. Empty equipment/serial/asset_tag fields replaced with "-".
 10. Signature image support: if a PNG/JPG exists in src/signatures/<Name>.png,
     it is embedded above the staff name in the prepared-by signature cell.

FOLDER STRUCTURE REQUIRED:
  your_project/
  ├── app.py                          ← this file
  ├── registered_users.json           ← auto-created on first run
  ├── src/
  │   ├── Equipment Accountability Form (Work From Home).dotx
  │   ├── Equipment Accountability Form (Work On Site).dotx
  │   └── signatures/
  │       ├── Jiro Macabitas.png      ← signature image for Jiro
  │       ├── Angelo Forbes.png       ← signature image for Angelo (optional)
  │       └── <Any Staff Name>.png    ← one file per staff member (optional)
  └── images/
      └── logo.png                    ← app logo (optional)

SIGNATURE IMAGE NOTES:
  - File name must exactly match the staff name in the "Prepared By" dropdown.
  - PNG with transparent background is recommended (e.g. 300x80px).
  - JPG/JPEG is also supported.
  - If no signature file is found for the selected staff, the form generates
    normally with just the text name — no error is raised.
"""

from __future__ import annotations

import io
import copy
import base64
import re
import struct
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


def _find_signature_image(prepared_by: str) -> Path | None:
    """
    Look for a signature image for the given staff name.
    Searches src/signatures/<Name>.png (or .jpg/.jpeg).
    File name must exactly match the prepared_by string.
    """
    sig_dir_candidates = [
        Path(__file__).parent / "src" / "signatures",
        Path.cwd() / "src" / "signatures",
        Path(__file__).parent / "signatures",
        Path.cwd() / "signatures",
    ]
    for sig_dir in sig_dir_candidates:
        if not sig_dir.exists():
            continue
        for ext in (".png", ".jpg", ".jpeg"):
            candidate = sig_dir / f"{prepared_by}{ext}"
            if candidate.exists():
                return candidate
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
    --input-bg:   #ffffff;
    --input-text: #0d2545;
    --chip-bg:    #ffffff;
    --step-bg:    #ffffff;
    --remark-bg:  #fffdf0;
    --remark-bdr: #ffe082;
    --remark-txt: #b07d00;
    --reg-bg:     #f7f9fd;
    --hint-bg:    #e8f0fc;
    --hint-txt:   #1565c0;
    --table-odd:  #f7f9fd;
    --table-hdr:  #0d2545;
    --mon-hdr-bg: #f0f4fa;
    --adapter-bg: #f7f9fd;
    --adapter-hdr:#eef2fb;
    --adapter-bdr:#dde5f0;
    --adapter-sub:#e4ebf5;
    --chip-def-bg:#f0f4fa;
    --prepby-bg:  #f0f7ff;
    --prepby-bdr: #b3c8f0;
    --copy-bg:    #f8f9ff;
    --copy-bdr:   #c5d4f0;
    --radius:     10px;
    --shadow:     0 2px 12px rgba(13,37,69,.09);
    --shadow-md:  0 4px 24px rgba(13,37,69,.14);
}}

@media (prefers-color-scheme: dark) {{
    :root {{
        --navy:       #c8d9f5;
        --blue:       #90baf9;
        --blue-lt:    #64b5f6;
        --blue-pale:  #1a2d4d;
        --orange:     #ffab76;
        --green:      #81c784;
        --green-lt:   #1b2e1c;
        --teal:       #4db6ac;
        --teal-lt:    #1b2e2b;
        --teal-pale:  #2a4a47;
        --bg:         #0e1621;
        --card:       #162033;
        --border:     #2a3d5a;
        --text:       #d6e4f7;
        --muted:      #7a9cc0;
        --input-bg:   #1a2840;
        --input-text: #d6e4f7;
        --chip-bg:    #1a2840;
        --step-bg:    #162033;
        --remark-bg:  #1e1a0a;
        --remark-bdr: #7a5a00;
        --remark-txt: #f0c040;
        --reg-bg:     #111d2e;
        --hint-bg:    #1a2d4d;
        --hint-txt:   #90baf9;
        --table-odd:  #111d2e;
        --table-hdr:  #0d1e33;
        --mon-hdr-bg: #111d2e;
        --adapter-bg: #111d2e;
        --adapter-hdr:#0d1e33;
        --adapter-bdr:#2a3d5a;
        --adapter-sub:#1a2840;
        --chip-def-bg:#0e1621;
        --prepby-bg:  #111d2e;
        --prepby-bdr: #2a4d80;
        --copy-bg:    #111d2e;
        --copy-bdr:   #2a4d80;
        --shadow:     0 2px 12px rgba(0,0,0,.4);
        --shadow-md:  0 4px 24px rgba(0,0,0,.6);
    }}
}}

@media (forced-colors: active) {{
    * {{ border-color: ButtonText !important; color: ButtonText !important;
         background: ButtonFace !important; }}
    a {{ color: LinkText !important; }}
    button {{ forced-color-adjust: none; }}
}}

html, body, [class*="css"] {{
    font-family: 'DM Sans', sans-serif !important;
    background: var(--bg) !important;
    color: var(--text) !important;
}}

[data-testid="stSidebar"]         {{ display: none !important; }}
header[data-testid="stHeader"]    {{ display: none !important; }}
[data-testid="stDecoration"]      {{ display: none !important; }}

[data-testid="stAppViewContainer"],
[data-testid="stMain"],
section.main,
.stApp {{
    background: var(--bg) !important;
    color: var(--text) !important;
}}

.main .block-container {{
    max-width: 860px !important;
    padding: 2rem 1.5rem 3rem !important;
}}

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
    color: rgba(255,255,255,.65);
    font-size: 0.74rem;
    font-weight: 400;
    margin-top: 2px;
}}

.copy-type-card {{
    background: var(--copy-bg);
    border: 1.5px solid var(--copy-bdr);
    border-radius: var(--radius);
    padding: 0.85rem 1.2rem;
    margin-bottom: 1rem;
    box-shadow: var(--shadow);
    color: var(--text);
}}
.copy-type-label {{
    font-size: 0.7rem;
    font-weight: 700;
    color: var(--blue);
    text-transform: uppercase;
    letter-spacing: 0.07em;
    margin-bottom: 0.5rem;
    display: flex;
    align-items: center;
    gap: 0.4rem;
}}
.copy-toggle-btn .stButton > button {{
    width: 100% !important;
    padding: 0.65rem 1rem !important;
    font-size: 0.82rem !important;
    font-weight: 700 !important;
    border-radius: 8px !important;
    letter-spacing: 0.01em !important;
    transition: all .18s !important;
}}
.copy-toggle-btn.it-active .stButton > button {{
    background: linear-gradient(135deg, #1565c0, #1e88e5) !important;
    color: #fff !important;
    border: 2px solid #1565c0 !important;
    box-shadow: 0 0 0 3px rgba(21,101,192,.25) !important;
}}
.copy-toggle-btn.it-inactive .stButton > button {{
    background: var(--card) !important;
    color: var(--navy) !important;
    border: 2px solid var(--border) !important;
    box-shadow: var(--shadow) !important;
}}
.copy-toggle-btn.it-inactive .stButton > button:hover {{
    background: var(--blue-pale) !important;
    border-color: var(--blue) !important;
    color: var(--navy) !important;
}}
.copy-toggle-btn.emp-active .stButton > button {{
    background: linear-gradient(135deg, #e65c00, #ff8f00) !important;
    color: #fff !important;
    border: 2px solid #e65c00 !important;
    box-shadow: 0 0 0 3px rgba(230,92,0,.25) !important;
}}
.copy-toggle-btn.emp-inactive .stButton > button {{
    background: var(--card) !important;
    color: var(--navy) !important;
    border: 2px solid var(--border) !important;
    box-shadow: var(--shadow) !important;
}}
.copy-toggle-btn.emp-inactive .stButton > button:hover {{
    background: #fff3e0 !important;
    border-color: #e65c00 !important;
    color: var(--navy) !important;
}}

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
    box-shadow: 0 0 0 3px rgba(21,101,192,.25), 0 2px 10px rgba(21,101,192,.3) !important;
}}
.form-toggle-btn.wfh-inactive .stButton > button {{
    background: var(--card) !important;
    color: var(--navy) !important;
    border: 2px solid var(--border) !important;
    box-shadow: var(--shadow) !important;
}}
.form-toggle-btn.wfh-inactive .stButton > button:hover {{
    background: var(--blue-pale) !important;
    border-color: var(--blue) !important;
    color: var(--navy) !important;
}}
.form-toggle-btn.onsite-active .stButton > button {{
    background: linear-gradient(135deg, #00796b, #26a69a) !important;
    color: #fff !important;
    border: 2px solid #00796b !important;
    box-shadow: 0 0 0 3px rgba(0,121,107,.25), 0 2px 10px rgba(0,121,107,.3) !important;
}}
.form-toggle-btn.onsite-inactive .stButton > button {{
    background: var(--card) !important;
    color: var(--navy) !important;
    border: 2px solid var(--border) !important;
    box-shadow: var(--shadow) !important;
}}
.form-toggle-btn.onsite-inactive .stButton > button:hover {{
    background: var(--teal-lt) !important;
    border-color: var(--teal) !important;
    color: var(--navy) !important;
}}

.step-card {{
    background: var(--card);
    border: 1px solid var(--border);
    border-radius: var(--radius);
    padding: 1.3rem 1.5rem 1rem;
    margin-bottom: 1rem;
    box-shadow: var(--shadow);
    color: var(--text);
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

.preparedby-card {{
    background: var(--prepby-bg);
    border: 1.5px solid var(--prepby-bdr);
    border-radius: var(--radius);
    padding: 1rem 1.4rem 0.85rem;
    margin-bottom: 1rem;
    box-shadow: var(--shadow);
    display: flex;
    align-items: center;
    gap: 0.9rem;
    color: var(--text);
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

.sig-preview-badge {{
    display: inline-flex;
    align-items: center;
    gap: 0.4rem;
    font-size: 0.72rem;
    font-weight: 600;
    color: var(--green);
    background: var(--green-lt);
    border: 1px solid #c3e6c0;
    border-radius: 6px;
    padding: 0.25rem 0.6rem;
    margin-top: 0.35rem;
}}
.sig-missing-badge {{
    display: inline-flex;
    align-items: center;
    gap: 0.4rem;
    font-size: 0.72rem;
    font-weight: 600;
    color: var(--muted);
    background: var(--reg-bg);
    border: 1px dashed var(--border);
    border-radius: 6px;
    padding: 0.25rem 0.6rem;
    margin-top: 0.35rem;
}}

.user-reg-panel {{
    background: var(--reg-bg);
    border: 1.5px solid var(--border);
    border-radius: var(--radius);
    padding: 1rem 1.2rem;
    margin-top: 0.6rem;
    margin-bottom: 0.4rem;
    color: var(--text);
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
    color: var(--text);
    background: var(--chip-bg);
    border: 1px solid var(--border);
    border-radius: 100px;
    padding: 3px 10px;
}}
.user-chip.default-chip {{
    color: var(--muted);
    background: var(--chip-def-bg);
    border-style: dashed;
}}

[data-testid="stHorizontalBlock"] > div {{
    display: flex;
    flex-direction: column;
    justify-content: flex-end;
}}

.stTextInput input,
.stSelectbox > div > div,
.stDateInput input {{
    border: 1.5px solid var(--border) !important;
    border-radius: 8px !important;
    box-shadow: none !important;
    background: var(--input-bg) !important;
    color: var(--input-text) !important;
    font-family: 'DM Sans', sans-serif !important;
    font-size: 0.87rem !important;
}}
.stTextInput input:focus {{
    border-color: var(--blue) !important;
    box-shadow: 0 0 0 3px rgba(21,101,192,.15) !important;
}}
[data-baseweb="select"] * {{
    background: var(--input-bg) !important;
    color: var(--input-text) !important;
}}
[data-baseweb="popover"] [role="option"] {{
    background: var(--input-bg) !important;
    color: var(--input-text) !important;
}}
[data-baseweb="popover"] [role="option"]:hover,
[data-baseweb="popover"] [aria-selected="true"] {{
    background: var(--blue-pale) !important;
    color: var(--navy) !important;
}}

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
    box-shadow: 0 4px 16px rgba(21,101,192,.45) !important;
    transform: translateY(-1px) !important;
    color: #fff !important;
}}

.btn-outline-blue .stButton > button {{
    background: var(--card) !important;
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
    color: var(--navy) !important;
    transform: none !important;
    box-shadow: none !important;
}}

.btn-danger .stButton > button {{
    background: var(--card) !important;
    color: #ef5350 !important;
    border: 1.5px solid #ef9a9a !important;
    font-size: 0.75rem !important;
    padding: 0.25rem 0.8rem !important;
    box-shadow: none !important;
    font-weight: 600 !important;
    height: 38px !important;
    min-height: 38px !important;
}}
.btn-danger .stButton > button:hover {{
    background: #3e1111 !important;
    border-color: #ef5350 !important;
    color: #ef9a9a !important;
    transform: none !important;
    box-shadow: none !important;
}}

[data-testid="stHorizontalBlock"] [data-testid="stColumn"] .stButton > button {{
    height: 38px !important;
    min-height: 38px !important;
    padding-top: 0 !important;
    padding-bottom: 0 !important;
    font-size: 0.78rem !important;
    line-height: 1 !important;
}}

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

.stCheckbox {{
    background: var(--reg-bg);
    border: 1px solid var(--border);
    border-radius: 8px;
    padding: 0.55rem 0.85rem !important;
    margin-bottom: 0.3rem !important;
    min-height: 2.2rem !important;
    transition: background .12s;
    display: flex !important;
    align-items: flex-start !important;
}}
.stCheckbox:hover {{ background: var(--blue-pale); border-color: var(--blue); }}
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
    background: var(--mon-hdr-bg);
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

.adapter-chip-panel {{
    margin: 0 0.9rem 0.65rem 2.4rem;
    background: var(--adapter-bg);
    border: 1px solid var(--adapter-bdr);
    border-radius: 8px;
    overflow: hidden;
}}
.adapter-chip-header {{
    display: flex;
    align-items: center;
    gap: 5px;
    padding: 0.32rem 0.7rem;
    border-bottom: 1px solid var(--adapter-sub);
    background: var(--adapter-hdr);
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
    background: var(--chip-bg);
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
.adapter-chip-none:hover {{ border-color: var(--muted); background: var(--blue-pale); }}
.adapter-chip.chip-selected {{
    background: var(--blue-pale); color: var(--blue); border-color: var(--blue); font-weight: 600;
}}
.adapter-chip-none.chip-selected {{
    background: var(--chip-def-bg); color: var(--muted); border-style: solid;
    border-color: var(--muted); font-weight: 600;
}}

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

.stAlert {{
    border-radius: 8px !important;
    font-family: 'DM Sans', sans-serif !important;
    font-size: 0.82rem !important;
    color: var(--text) !important;
}}

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

.info-hint {{
    font-size: 0.77rem;
    color: var(--hint-txt);
    background: var(--hint-bg);
    border-left: 3px solid var(--blue);
    border-radius: 0 6px 6px 0;
    padding: 0.5rem 0.7rem;
    margin-bottom: 0.8rem;
    font-weight: 400;
    line-height: 1.6;
}}
.info-hint strong {{ font-weight: 700; color: var(--navy); }}

[data-testid="stFileUploader"] section {{
    border: 2px dashed var(--border) !important;
    border-radius: var(--radius) !important;
    background: var(--reg-bg) !important;
    color: var(--text) !important;
}}
[data-testid="stFileUploader"] * {{
    color: var(--text) !important;
}}

.bmg-footer {{
    text-align: center;
    color: var(--muted);
    font-size: 0.7rem;
    margin-top: 1.8rem;
    padding-top: 0.9rem;
    border-top: 1px solid var(--border);
}}

.preview-table {{
    width: 100%;
    border-collapse: collapse;
    font-size: 0.79rem;
    margin-top: 0.5rem;
    color: var(--text);
}}
.preview-table th {{
    background: var(--table-hdr);
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
.preview-table tr:nth-child(even) td {{ background: var(--table-odd); }}
.preview-table tr:hover td {{ background: var(--blue-pale); }}

.remarks-wrap {{
    background: var(--remark-bg);
    border: 1.5px solid var(--remark-bdr);
    border-radius: 8px;
    padding: 0.8rem 0.95rem 0.55rem;
    margin-top: 0.3rem;
    margin-bottom: 0.5rem;
}}
.remarks-label {{
    font-size: 0.7rem;
    font-weight: 700;
    color: var(--remark-txt);
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

[data-testid="stExpander"] {{
    background: var(--card) !important;
    border: 1px solid var(--border) !important;
    border-radius: var(--radius) !important;
    color: var(--text) !important;
}}
[data-testid="stExpander"] summary {{
    color: var(--text) !important;
}}

p, span, div, li {{
    color: var(--text);
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
    " laptop charger": 2, "monitor": 3,
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
_TABLE_SIZE = "20"

_JS_ONCLICK_REGEX = r"/'([^']+)'\s*\)$/"

# Word/OOXML relationship namespaces
_RELS_NS  = "http://schemas.openxmlformats.org/package/2006/relationships"
_IMG_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
_WP_NS    = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
_A_NS     = "http://schemas.openxmlformats.org/drawingml/2006/main"
_PIC_NS   = "http://schemas.openxmlformats.org/drawingml/2006/picture"
_R_NS     = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

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
        "equipment": " Laptop Charger", "serial": "", "asset_tag": "",
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

_ROW_HEIGHT_DXA = "280"
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


def _set_sdt_value(
    body, tag_val: str, new_text: str,
    bold: bool = False,
    size: str = _FONT_SIZE,
) -> bool:
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
    return _set_sdt_value(body, "Contact No.", position, bold=False, size=_FONT_SIZE)


def _patch_copy_label(body, copy_type: str):
    for t_el in body.iter(f"{{{W}}}t"):
        if t_el.text:
            if copy_type == "employee":
                t_el.text = t_el.text.replace("Verified By", "Received By")
                t_el.text = t_el.text.replace("VERIFIED BY", "RECEIVED BY")
                t_el.text = t_el.text.replace("Verified by", "Received by")
                t_el.text = t_el.text.replace("IT Copy", "Employee Copy")
                t_el.text = t_el.text.replace("IT COPY", "EMPLOYEE COPY")
                t_el.text = t_el.text.replace("It Copy", "Employee Copy")
            else:
                t_el.text = t_el.text.replace("Verified By", "Verified By")
                t_el.text = t_el.text.replace("Employee Copy", "IT Copy")
                t_el.text = t_el.text.replace("EMPLOYEE COPY", "IT COPY")

def _get_equipment_table(body):
    for tbl in body.iter(f"{{{W}}}tbl"):
        tblGrid = tbl.find(f"{{{W}}}tblGrid")
        if tblGrid is not None and len(tblGrid.findall(f"{{{W}}}gridCol")) == 5:
            return tbl
    return None


def _set_cell_text(cell_el, text: str, top_align: bool = False):
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
    # vertical alignment on the cell
    tcPr = cell_el.find(f"{{{W}}}tcPr")
    if tcPr is None:
        tcPr = etree.SubElement(cell_el, f"{{{W}}}tcPr")
        cell_el.insert(0, tcPr)
    for va in tcPr.findall(f"{{{W}}}vAlign"):
        tcPr.remove(va)
    vAlign = etree.SubElement(tcPr, f"{{{W}}}vAlign")
    vAlign.set(f"{{{W}}}val", "top" if top_align else "center")


def _apply_compact_cell(tc, col_idx, pad_top, pad_btm, pad_left, pad_right,
                         font_size, remarks_col_index):
    """Apply compact formatting to a single table cell."""
    tcPr = tc.find(f"{{{W}}}tcPr")
    if tcPr is None:
        tcPr = etree.SubElement(tc, f"{{{W}}}tcPr")
        tc.insert(0, tcPr)
    noWrap = tcPr.find(f"{{{W}}}noWrap")
    if noWrap is None:
        etree.SubElement(tcPr, f"{{{W}}}noWrap")
    tcMar = tcPr.find(f"{{{W}}}tcMar")
    if tcMar is None:
        tcMar = etree.SubElement(tcPr, f"{{{W}}}tcMar")
    for side, val in [("top", pad_top), ("bottom", pad_btm),
                      ("left", pad_left), ("right", pad_right)]:
        el = tcMar.find(f"{{{W}}}{side}")
        if el is None:
            el = etree.SubElement(tcMar, f"{{{W}}}{side}")
        el.set(f"{{{W}}}w", val)
        el.set(f"{{{W}}}type", "dxa")
    for va in tcPr.findall(f"{{{W}}}vAlign"):
        tcPr.remove(va)
    vAlign = etree.SubElement(tcPr, f"{{{W}}}vAlign")
    # remarks column top-aligned, all others center
    vAlign.set(f"{{{W}}}val", "top" if col_idx == remarks_col_index else "center")
    for p in tc.iter(f"{{{W}}}p"):
        pPr = p.find(f"{{{W}}}pPr")
        if pPr is not None:
            for spacing in pPr.findall(f"{{{W}}}spacing"):
                pPr.remove(spacing)
            for cs in pPr.findall(f"{{{W}}}contextualSpacing"):
                pPr.remove(cs)
            sp_el = etree.SubElement(pPr, f"{{{W}}}spacing")
            sp_el.set(f"{{{W}}}before",   "0")
            sp_el.set(f"{{{W}}}after",    "0")
            sp_el.set(f"{{{W}}}line",     "240")
            sp_el.set(f"{{{W}}}lineRule", "auto")
        for r in p.findall(f"{{{W}}}r"):
            rPr = r.find(f"{{{W}}}rPr")
            if rPr is None:
                rPr = etree.SubElement(r, f"{{{W}}}rPr")
                r.insert(0, rPr)
            for sz_tag in (f"{{{W}}}sz", f"{{{W}}}szCs"):
                sz_el = rPr.find(sz_tag)
                if sz_el is None:
                    sz_el = etree.SubElement(rPr, sz_tag)
                sz_el.set(f"{{{W}}}val", font_size)
            for spacing in rPr.findall(f"{{{W}}}spacing"):
                rPr.remove(spacing)


def _compact_row(
    row_el,
    row_height: str       = _ROW_HEIGHT_DXA,
    pad_top:    str       = _CELL_PAD_TOP,
    pad_btm:    str       = _CELL_PAD_BTM,
    pad_left:   str       = _CELL_PAD_LEFT,
    pad_right:  str       = _CELL_PAD_RIGHT,
    font_size:  str       = _TABLE_SIZE,
    remarks_col_index: int = 4,
):
    trPr = row_el.find(f"{{{W}}}trPr")
    if trPr is None:
        trPr = etree.SubElement(row_el, f"{{{W}}}trPr")
        row_el.insert(0, trPr)
    for trH in trPr.findall(f"{{{W}}}trHeight"):
        trPr.remove(trH)
    trH = etree.SubElement(trPr, f"{{{W}}}trHeight")
    trH.set(f"{{{W}}}val",   row_height)
    trH.set(f"{{{W}}}hRule", "exact")

    # Walk direct children only — same traversal as _fill_equipment_row
    col_idx = 0
    for ch in row_el:
        if ch.tag == f"{{{W}}}tc":
            _apply_compact_cell(ch, col_idx, pad_top, pad_btm, pad_left, pad_right,
                                font_size, remarks_col_index)
            col_idx += 1
        elif ch.tag == f"{{{W}}}sdt":
            sc = ch.find(f"{{{W}}}sdtContent")
            if sc is None:
                continue
            for tc in sc.findall(f"{{{W}}}tc"):
                _apply_compact_cell(tc, col_idx, pad_top, pad_btm, pad_left, pad_right,
                                    font_size, remarks_col_index)
                col_idx += 1


def _compact_page_margins(body, num_rows: int = 0):
    sectPr = body.find(f"{{{W}}}sectPr")
    if sectPr is None:
        return
    pgMar = sectPr.find(f"{{{W}}}pgMar")
    if pgMar is None:
        pgMar = etree.SubElement(sectPr, f"{{{W}}}pgMar")
    if num_rows > 15:
        v_margin = 320
    elif num_rows > 10:
        v_margin = 460
    else:
        v_margin = 640
    h_margin = 540
    for side, max_val in [("top", v_margin), ("bottom", v_margin),
                          ("left", h_margin), ("right", h_margin)]:
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

    # Replace blank values with "-"
    equipment = equipment.strip() if equipment and equipment.strip() else "-"
    serial    = serial.strip()    if serial    and serial.strip()    else "-"
    asset_tag = asset_tag.strip() if asset_tag and asset_tag.strip() else "-"

    for i, (cell, text, top) in enumerate(zip(
        cells,
        ["", equipment, serial, asset_tag, remarks],
        [False, False,  False,  False,     True],
    )):
        if i == 0:
            continue
        _set_cell_text(cell, text, top_align=top)


# ─────────────────────────────────────────────
# SIGNATURE IMAGE — WORD EMBEDDING
# ─────────────────────────────────────────────

def _get_png_dimensions(data: bytes) -> tuple[int, int]:
    """Read width, height from a PNG header. Returns (w, h) in pixels."""
    try:
        if data[:8] == b'\x89PNG\r\n\x1a\n':
            w, h = struct.unpack('>II', data[16:24])
            return w, h
    except Exception:
        pass
    return 300, 80   # safe fallback


def _get_jpeg_dimensions(data: bytes) -> tuple[int, int]:
    """Read width, height from a JPEG. Returns (w, h) in pixels."""
    try:
        i = 2
        while i < len(data):
            marker = data[i:i+2]
            i += 2
            if marker in (b'\xff\xc0', b'\xff\xc1', b'\xff\xc2'):
                h, w = struct.unpack('>HH', data[i+3:i+7])
                return w, h
            length = struct.unpack('>H', data[i:i+2])[0]
            i += length
    except Exception:
        pass
    return 300, 80


def _add_image_to_zip(files: dict, image_bytes: bytes, ext: str) -> str:
    """
    Register the signature image in the zip and its relationship.
    Returns the relationship ID string (e.g. 'rId901').
    """
    rId            = "rId901"
    media_path     = f"word/media/sig_prepared_by{ext}"
    rels_path      = "word/_rels/document.xml.rels"

    # Save image bytes into the zip dict
    files[media_path] = image_bytes

    # Parse existing rels, add new one
    mime = "image/png" if ext == ".png" else "image/jpeg"
    rel_xml = files.get(rels_path, b'')
    try:
        rels_root = etree.fromstring(rel_xml)
    except Exception:
        rels_root = etree.Element(f"{{{_RELS_NS}}}Relationships")

    # Remove any stale sig rel from a previous generation
    for old in rels_root.findall(f"{{{_RELS_NS}}}Relationship[@Id='{rId}']"):
        rels_root.remove(old)

    rel_el = etree.SubElement(rels_root, f"{{{_RELS_NS}}}Relationship")
    rel_el.set("Id",         rId)
    rel_el.set("Type",       _IMG_TYPE)
    rel_el.set("Target",     f"media/sig_prepared_by{ext}")

    files[rels_path] = etree.tostring(
        rels_root, xml_declaration=True, encoding="UTF-8", standalone=True)
    return rId


def _make_inline_image_run(rId: str, img_w_px: int, img_h_px: int,
                            max_w_emu: int = 1_200_000) -> etree._Element:
    """
    Build a <w:r> containing an inline drawing for the signature image.
    Sizes the image proportionally, capped at max_w_emu (default ~3.2 cm).
    EMU = English Metric Units; 1 px at 96dpi ≈ 9525 EMU.
    """
    PX_TO_EMU = 9525
    w_emu = img_w_px * PX_TO_EMU
    h_emu = img_h_px * PX_TO_EMU

    if w_emu > max_w_emu:
        scale = max_w_emu / w_emu
        w_emu = int(w_emu * scale)
        h_emu = int(h_emu * scale)

    r = etree.Element(f"{{{W}}}r")

    drawing = etree.SubElement(r, f"{{{W}}}drawing")
    inline  = etree.SubElement(drawing, f"{{{_WP_NS}}}inline")
    inline.set("distT", "0"); inline.set("distB", "0")
    inline.set("distL", "0"); inline.set("distR", "0")

    extent = etree.SubElement(inline, f"{{{_WP_NS}}}extent")
    extent.set("cx", str(w_emu))
    extent.set("cy", str(h_emu))

    effectExtent = etree.SubElement(inline, f"{{{_WP_NS}}}effectExtent")
    effectExtent.set("l", "0"); effectExtent.set("t", "0")
    effectExtent.set("r", "0"); effectExtent.set("b", "0")

    docPr = etree.SubElement(inline, f"{{{_WP_NS}}}docPr")
    docPr.set("id", "901"); docPr.set("name", "SignatureImage")

    cNvGraphicFramePr = etree.SubElement(inline, f"{{{_WP_NS}}}cNvGraphicFramePr")
    graphicFrameLocks = etree.SubElement(
        cNvGraphicFramePr,
        "{http://schemas.openxmlformats.org/drawingml/2006/main}graphicFrameLocks"
    )
    graphicFrameLocks.set("noChangeAspect", "1")

    graphic = etree.SubElement(
        inline, "{http://schemas.openxmlformats.org/drawingml/2006/main}graphic")
    graphicData = etree.SubElement(graphic,
        "{http://schemas.openxmlformats.org/drawingml/2006/main}graphicData")
    graphicData.set(
        "uri", "http://schemas.openxmlformats.org/drawingml/2006/picture")

    pic = etree.SubElement(graphicData,
        "{http://schemas.openxmlformats.org/drawingml/2006/picture}pic")

    nvPicPr = etree.SubElement(pic,
        "{http://schemas.openxmlformats.org/drawingml/2006/picture}nvPicPr")
    cNvPr = etree.SubElement(nvPicPr,
        "{http://schemas.openxmlformats.org/drawingml/2006/picture}cNvPr")
    cNvPr.set("id", "0"); cNvPr.set("name", "SignatureImage")
    cNvPicPr = etree.SubElement(nvPicPr,
        "{http://schemas.openxmlformats.org/drawingml/2006/picture}cNvPicPr")

    blipFill = etree.SubElement(pic,
        "{http://schemas.openxmlformats.org/drawingml/2006/picture}blipFill")
    blip = etree.SubElement(blipFill,
        "{http://schemas.openxmlformats.org/drawingml/2006/main}blip")
    blip.set(f"{{{_R_NS}}}embed", rId)

    stretch = etree.SubElement(blipFill,
        "{http://schemas.openxmlformats.org/drawingml/2006/main}stretch")
    etree.SubElement(stretch,
        "{http://schemas.openxmlformats.org/drawingml/2006/main}fillRect")

    spPr = etree.SubElement(pic,
        "{http://schemas.openxmlformats.org/drawingml/2006/picture}spPr")
    xfrm = etree.SubElement(spPr,
        "{http://schemas.openxmlformats.org/drawingml/2006/main}xfrm")
    off = etree.SubElement(xfrm,
        "{http://schemas.openxmlformats.org/drawingml/2006/main}off")
    off.set("x", "0"); off.set("y", "0")
    ext2 = etree.SubElement(xfrm,
        "{http://schemas.openxmlformats.org/drawingml/2006/main}ext")
    ext2.set("cx", str(w_emu)); ext2.set("cy", str(h_emu))
    prstGeom = etree.SubElement(spPr,
        "{http://schemas.openxmlformats.org/drawingml/2006/main}prstGeom")
    prstGeom.set("prst", "rect")
    etree.SubElement(prstGeom,
        "{http://schemas.openxmlformats.org/drawingml/2006/main}avLst")

    return r


def _insert_signature_into_cell(cell_el, rId: str, img_bytes: bytes, ext: str):
    """
    Insert the signature image as a floating image (wrap in front of text),
    positioned to the right side of the signature cell.
    """
    if ext == ".png":
        w_px, h_px = _get_png_dimensions(img_bytes)
    else:
        w_px, h_px = _get_jpeg_dimensions(img_bytes)

    PX_TO_EMU = 9525
    max_w_emu = 2_000_000
    w_emu = w_px * PX_TO_EMU
    h_emu = h_px * PX_TO_EMU

    if w_emu > max_w_emu:
        scale = max_w_emu / w_emu
        w_emu = int(w_emu * scale)
        h_emu = int(h_emu * scale)

    # Horizontal offset from left edge of cell — push image to the right
    pos_x_emu = 600_000   # increase this value to move further right
    pos_y_emu = -150_000

    r = etree.Element(f"{{{W}}}r")
    drawing = etree.SubElement(r, f"{{{W}}}drawing")

    anchor = etree.SubElement(drawing, f"{{{_WP_NS}}}anchor")
    anchor.set("distT", "0")
    anchor.set("distB", "0")
    anchor.set("distL", "0")
    anchor.set("distR", "0")
    anchor.set("simplePos", "0")
    anchor.set("relativeHeight", "251658240")
    anchor.set("behindDoc", "0")      # 0 = in front of text
    anchor.set("locked", "0")
    anchor.set("layoutInCell", "1")
    anchor.set("allowOverlap", "1")

    simplePos = etree.SubElement(anchor, f"{{{_WP_NS}}}simplePos")
    simplePos.set("x", "0")
    simplePos.set("y", "0")

    posH = etree.SubElement(anchor, f"{{{_WP_NS}}}positionH")
    posH.set("relativeFrom", "column")
    posOffset_h = etree.SubElement(posH, f"{{{_WP_NS}}}posOffset")
    posOffset_h.text = str(pos_x_emu)

    posV = etree.SubElement(anchor, f"{{{_WP_NS}}}positionV")
    posV.set("relativeFrom", "paragraph")
    posOffset_v = etree.SubElement(posV, f"{{{_WP_NS}}}posOffset")
    posOffset_v.text = str(pos_y_emu)

    extent = etree.SubElement(anchor, f"{{{_WP_NS}}}extent")
    extent.set("cx", str(w_emu))
    extent.set("cy", str(h_emu))

    effectExtent = etree.SubElement(anchor, f"{{{_WP_NS}}}effectExtent")
    effectExtent.set("l", "0")
    effectExtent.set("t", "0")
    effectExtent.set("r", "0")
    effectExtent.set("b", "0")

    wrapNone = etree.SubElement(anchor, f"{{{_WP_NS}}}wrapNone")

    docPr = etree.SubElement(anchor, f"{{{_WP_NS}}}docPr")
    docPr.set("id", "901")
    docPr.set("name", "SignatureImage")

    cNvGraphicFramePr = etree.SubElement(anchor, f"{{{_WP_NS}}}cNvGraphicFramePr")
    graphicFrameLocks = etree.SubElement(
        cNvGraphicFramePr,
        "{http://schemas.openxmlformats.org/drawingml/2006/main}graphicFrameLocks"
    )
    graphicFrameLocks.set("noChangeAspect", "1")

    graphic = etree.SubElement(
        anchor, "{http://schemas.openxmlformats.org/drawingml/2006/main}graphic")
    graphicData = etree.SubElement(graphic,
        "{http://schemas.openxmlformats.org/drawingml/2006/main}graphicData")
    graphicData.set("uri", "http://schemas.openxmlformats.org/drawingml/2006/picture")

    pic = etree.SubElement(graphicData,
        "{http://schemas.openxmlformats.org/drawingml/2006/picture}pic")

    nvPicPr = etree.SubElement(pic,
        "{http://schemas.openxmlformats.org/drawingml/2006/picture}nvPicPr")
    cNvPr = etree.SubElement(nvPicPr,
        "{http://schemas.openxmlformats.org/drawingml/2006/picture}cNvPr")
    cNvPr.set("id", "0")
    cNvPr.set("name", "SignatureImage")
    etree.SubElement(nvPicPr,
        "{http://schemas.openxmlformats.org/drawingml/2006/picture}cNvPicPr")

    blipFill = etree.SubElement(pic,
        "{http://schemas.openxmlformats.org/drawingml/2006/picture}blipFill")
    blip = etree.SubElement(blipFill,
        "{http://schemas.openxmlformats.org/drawingml/2006/main}blip")
    blip.set(f"{{{_R_NS}}}embed", rId)

    stretch = etree.SubElement(blipFill,
        "{http://schemas.openxmlformats.org/drawingml/2006/main}stretch")
    etree.SubElement(stretch,
        "{http://schemas.openxmlformats.org/drawingml/2006/main}fillRect")

    spPr = etree.SubElement(pic,
        "{http://schemas.openxmlformats.org/drawingml/2006/picture}spPr")
    xfrm = etree.SubElement(spPr,
        "{http://schemas.openxmlformats.org/drawingml/2006/main}xfrm")
    off = etree.SubElement(xfrm,
        "{http://schemas.openxmlformats.org/drawingml/2006/main}off")
    off.set("x", "0")
    off.set("y", "0")
    ext2 = etree.SubElement(xfrm,
        "{http://schemas.openxmlformats.org/drawingml/2006/main}ext")
    ext2.set("cx", str(w_emu))
    ext2.set("cy", str(h_emu))
    prstGeom = etree.SubElement(spPr,
        "{http://schemas.openxmlformats.org/drawingml/2006/main}prstGeom")
    prstGeom.set("prst", "rect")
    etree.SubElement(prstGeom,
        "{http://schemas.openxmlformats.org/drawingml/2006/main}avLst")

    # Insert the floating image run into the first paragraph of the cell
    first_p = cell_el.find(f"{{{W}}}p")
    if first_p is not None:
        first_p.append(r)
    else:
        p = etree.SubElement(cell_el, f"{{{W}}}p")
        p.append(r)


def _fill_prepared_by(body, prepared_by: str, files: dict | None = None) -> bool:
    """
    Fill the [STAFF NAME] placeholder in the signature table.
    If `files` is provided and a signature image exists, embeds it above the name.
    """
    if not prepared_by:
        return False

    sig_path  = _find_signature_image(prepared_by)
    rId       = None
    img_bytes = None
    img_ext   = None

    if sig_path and files is not None:
        try:
            img_bytes = sig_path.read_bytes()
            img_ext   = sig_path.suffix.lower()
            rId       = _add_image_to_zip(files, img_bytes, img_ext)
        except Exception:
            rId = None

    # Find the signature cell (3rd table, 5th row, 2nd cell)
    tables = list(body.iter(f"{{{W}}}tbl"))
    sig_cell = None
    if len(tables) >= 3:
        sig_table = tables[2]
        rows = sig_table.findall(f"{{{W}}}tr")
        if len(rows) >= 5:
            cells = rows[4].findall(f"{{{W}}}tc")
            if len(cells) >= 2:
                sig_cell = cells[1]

    if sig_cell is None:
        return _fill_prepared_by_fallback(body, prepared_by)

    # Replace [STAFF NAME] text
    replaced = False
    for t_el in sig_cell.iter(f"{{{W}}}t"):
        if t_el.text and "[STAFF NAME]" in t_el.text:
            t_el.text = t_el.text.replace("[STAFF NAME]", prepared_by)
            replaced = True

    # If we have a signature image, embed it before the text
    if rId and img_bytes and img_ext:
        _insert_signature_into_cell(sig_cell, rId, img_bytes, img_ext)

    return replaced


def _fill_prepared_by_fallback(body, prepared_by: str) -> bool:
    replaced = False
    for t_el in body.iter(f"{{{W}}}t"):
        if t_el.text and "[STAFF NAME]" in t_el.text:
            t_el.text = t_el.text.replace("[STAFF NAME]", prepared_by)
            replaced = True
    return replaced
def _apply_remarks_rowspan(eq_table, num_data_rows: int):
    """
    Merges the remarks cell (col index 4) across all data rows
    so it appears as one unified cell instead of being split by row borders.
    """
    all_rows = eq_table.findall(f"{{{W}}}tr")
    data_rows = all_rows[1:]  # skip header row

    for row_i, row_el in enumerate(data_rows[:num_data_rows]):
        cells = []
        for ch in row_el:
            if ch.tag == f"{{{W}}}tc":
                cells.append(ch)
            elif ch.tag == f"{{{W}}}sdt":
                sc = ch.find(f"{{{W}}}sdtContent")
                if sc is not None:
                    for tc in sc.findall(f"{{{W}}}tc"):
                        cells.append(tc)

        if len(cells) < 5:
            continue

        remarks_cell = cells[4]
        tcPr = remarks_cell.find(f"{{{W}}}tcPr")
        if tcPr is None:
            tcPr = etree.SubElement(remarks_cell, f"{{{W}}}tcPr")
            remarks_cell.insert(0, tcPr)

        # Remove any existing vMerge
        for vm in tcPr.findall(f"{{{W}}}vMerge"):
            tcPr.remove(vm)

        vMerge = etree.SubElement(tcPr, f"{{{W}}}vMerge")

        if row_i == 0:
            # First data row: start the merge + write the remarks text
            vMerge.set(f"{{{W}}}val", "restart")
        else:
            # All subsequent rows: continuation cell must be empty
            for p in remarks_cell.findall(f"{{{W}}}p"):
                remarks_cell.remove(p)
            empty_p = etree.SubElement(remarks_cell, f"{{{W}}}p")
            etree.SubElement(empty_p, f"{{{W}}}pPr")

def fill_template(
    sorted_rows: list[dict],
    employee_name: str,
    client: str,
    position: str,
    date_str: str,
    prepared_by: str = "",
    form_type: str = "wfh",
    copy_type: str = "it",
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
        # Also register the image content type if not already present
        ct_root = etree.fromstring(files["[Content_Types].xml"])
        ct_ns   = "http://schemas.openxmlformats.org/package/2006/content-types"
        existing_exts = {
            el.get("Extension", "")
            for el in ct_root.findall(f"{{{ct_ns}}}Default")
        }
        for ext_str, mime_str in [("png", "image/png"), ("jpeg", "image/jpeg"), ("jpg", "image/jpeg")]:
            if ext_str not in existing_exts:
                def_el = etree.SubElement(ct_root, f"{{{ct_ns}}}Default")
                def_el.set("Extension",   ext_str)
                def_el.set("ContentType", mime_str)
        files["[Content_Types].xml"] = etree.tostring(
            ct_root, xml_declaration=True, encoding="UTF-8", standalone=True)

    if "docProps/app.xml" in files:
        files["docProps/app.xml"] = _patch_app_xml(files["docProps/app.xml"])

    root = etree.fromstring(files["word/document.xml"])
    body = root.find(f"{{{W}}}body")

    _set_sdt_value(body, "Name",   employee_name, bold=False, size=_FONT_SIZE)
    _set_sdt_value(body, "Client", client,        bold=False, size=_FONT_SIZE)
    _set_sdt_value(body, "Date",   date_str,      bold=False, size=_FONT_SIZE)
    _fill_position_sdt(body, position)
    _shrink_header_sdt_cells(body)
    _patch_copy_label(body, copy_type)

    num_rows = len(sorted_rows)

    if num_rows <= 10:
        row_h     = "280"
        pad_v     = "40"
        pad_h     = "80"
        font_size = "20"
    elif num_rows <= 14:
        row_h     = "240"
        pad_v     = "30"
        pad_h     = "70"
        font_size = "18"
    elif num_rows <= 18:
        row_h     = "200"
        pad_v     = "20"
        pad_h     = "60"
        font_size = "16"
    else:
        row_h     = "180"
        pad_v     = "12"
        pad_h     = "50"
        font_size = "15"

    _compact_page_margins(body, num_rows=num_rows)

    # Pass `files` so signature image can be embedded into the zip
    if prepared_by:
        _fill_prepared_by(body, prepared_by, files=files)

    eq_table = _get_equipment_table(body)
    if eq_table is not None:
        all_rows     = eq_table.findall(f"{{{W}}}tr")
        data_rows    = all_rows[1:]
        template_row = copy.deepcopy(data_rows[0]) if data_rows else None
        num_assets   = len(sorted_rows)

        geo = dict(
            row_height = row_h,
            pad_top    = pad_v,
            pad_btm    = pad_v,
            pad_left   = pad_h,
            pad_right  = pad_h,
            font_size  = font_size,
        )

        for i, asset in enumerate(sorted_rows):
            if i < len(data_rows):
                row_el = data_rows[i]
                _compact_row(row_el, **geo)
                _fill_equipment_row(row_el, asset["equipment"], asset["serial"],
                                    asset["asset_tag"], asset["remarks"])
            elif template_row is not None:
                new_row = copy.deepcopy(template_row)
                _compact_row(new_row, **geo)
                eq_table.append(new_row)
                _fill_equipment_row(new_row, asset["equipment"], asset["serial"],
                                    asset["asset_tag"], asset["remarks"])

        all_rows_now = eq_table.findall(f"{{{W}}}tr")
        for extra_row in all_rows_now[1 + num_assets:]:
            eq_table.remove(extra_row)
        _apply_remarks_rowspan(eq_table, num_assets)

        for data_row in eq_table.findall(f"{{{W}}}tr")[1:]:
            _compact_row(data_row, **geo)

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


def render_copy_type_selector() -> str:
    if "copy_type" not in st.session_state:
        st.session_state["copy_type"] = "it"
    current = st.session_state["copy_type"]

    st.markdown("""
    <div class="copy-type-card">
      <div class="copy-type-label">📋 Copy Type</div>
    </div>
    """, unsafe_allow_html=True)

    col_it, col_emp = st.columns(2)
    with col_it:
        it_css = "it-active" if current == "it" else "it-inactive"
        st.markdown(f'<div class="copy-toggle-btn {it_css}">', unsafe_allow_html=True)
        if st.button("🖥️  IT Copy", key="btn_copy_it", use_container_width=True):
            if current != "it":
                st.session_state["copy_type"] = "it"
                st.session_state.pop("docx", None)
                st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)

    with col_emp:
        emp_css = "emp-active" if current == "employee" else "emp-inactive"
        st.markdown(f'<div class="copy-toggle-btn {emp_css}">', unsafe_allow_html=True)
        if st.button("👤  Employee Copy", key="btn_copy_emp", use_container_width=True):
            if current != "employee":
                st.session_state["copy_type"] = "employee"
                st.session_state.pop("docx", None)
                st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)

    if current == "employee":
        st.markdown(
            '<div style="font-size:0.74rem;color:var(--orange);background:#fff3e0;'
            'border-left:3px solid #e65c00;border-radius:0 6px 6px 0;'
            'padding:0.38rem 0.7rem;margin-top:0.4rem;">'
            '⚠️ "Verified By" will be changed to "Received By" in the document.'
            '</div>',
            unsafe_allow_html=True,
        )
    return current


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

    # Show signature status badge
    if prepared_by:
        sig_path = _find_signature_image(prepared_by)
        if sig_path:
            st.markdown(
                f'<div class="sig-preview-badge">'
                f'✅ Signature image found — will be embedded in document'
                f'</div>',
                unsafe_allow_html=True,
            )
        else:
            st.markdown(
                f'<div class="sig-missing-badge">'
                f'○ No signature image — add <code>src/signatures/{prepared_by}.png</code> to enable'
                f'</div>',
                unsafe_allow_html=True,
            )

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
        f'<div style="font-size:0.7rem;font-weight:700;color:var(--muted);text-transform:uppercase;'
        f'letter-spacing:.06em;margin-bottom:6px;">Current staff</div>'
        f'<div class="user-chip-wrap">{chips_html}</div>'
        f'</div>',
        unsafe_allow_html=True,
    )
    st.markdown("<hr>", unsafe_allow_html=True)

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

_DOWNSTREAM_KEYS = (
    "df_loaded", "csv_file_id",
    "docx", "form_name", "form_client", "form_date",
    "prepared_by", "form_type_used", "copy_type_used",
    "shared_remarks",
)


def _clear_downstream():
    for k in list(st.session_state.keys()):
        if k.startswith("sel_") or k.startswith("chk_"):
            del st.session_state[k]
    for k in _DOWNSTREAM_KEYS:
        st.session_state.pop(k, None)


DEFAULT_CSV_PATH = Path(__file__).resolve().parent / "src" / "assets.csv"

def render_csv_upload() -> tuple[pd.DataFrame | None, str]:
    # Try to load the default CSV if no file has been uploaded yet
    default_df = None
    default_label = ""
    if DEFAULT_CSV_PATH.exists():
        try:
            default_df = pd.read_csv(DEFAULT_CSV_PATH, encoding="utf-8-sig")
            default_label = f"{DEFAULT_CSV_PATH.name} · {len(default_df):,} rows (default)"
        except Exception:
            default_df = None

    uploaded = st.file_uploader(
        "Upload asset list",
        type=["csv", "xlsx", "xls"],
        label_visibility="collapsed",
        key="csv_upload",
    )

    if uploaded is None:
        # No file uploaded — use default if available
        if st.session_state.get("csv_file_id") not in (None, "__default__"):
            _clear_downstream()

        if default_df is not None:
            if st.session_state.get("csv_file_id") != "__default__":
                _clear_downstream()
                st.session_state["csv_file_id"] = "__default__"
                st.session_state["df_loaded"] = default_df

            st.markdown(
                f'<div class="csv-source-bar">📂&nbsp; <span>{default_label}</span></div>',
                unsafe_allow_html=True,
            )
            st.caption("Using default asset list — upload a file above to replace it.")
            return st.session_state["df_loaded"], default_label

        # No default either
        if st.session_state.get("csv_file_id") is not None:
            _clear_downstream()
        st.markdown(
            '<div class="info-hint">'
            '<strong>Tip:</strong> Export your asset list from SharePoint or Excel as a '
            '<strong>.csv</strong> or <strong>.xlsx</strong> file, then upload it here.'
            '</div>',
            unsafe_allow_html=True,
        )
        return None, ""

    # A file was uploaded — treat it normally
    file_id = f"{uploaded.name}::{uploaded.size}"

    if st.session_state.get("csv_file_id") != file_id:
        _clear_downstream()
        st.session_state["csv_file_id"] = file_id

    if "df_loaded" in st.session_state:
        df    = st.session_state["df_loaded"]
        label = f"{uploaded.name} · {len(df):,} rows"
        st.markdown(
            f'<div class="csv-source-bar">✅&nbsp; <span>{label}</span></div>',
            unsafe_allow_html=True,
        )
        return df, label

    try:
        if uploaded.name.lower().endswith((".xlsx", ".xls")):
            df = pd.read_excel(uploaded)
        else:
            raw = uploaded.read()
            df  = None
            for enc in ("utf-8-sig", "utf-8", "latin-1", "cp1252"):
                try:
                    df = pd.read_csv(io.BytesIO(raw), encoding=enc)
                    break
                except UnicodeDecodeError:
                    continue
            if df is None:
                st.error("Could not decode the CSV file. Try saving it as UTF-8.")
                return None, ""
    except Exception as e:
        st.error(f"Failed to read file: {e}")
        return None, ""

    if df.empty:
        st.warning("The uploaded file appears to be empty.")
        return None, ""

    st.session_state["df_loaded"] = df
    label = f"{uploaded.name} · {len(df):,} rows"
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
            '<div style="margin-top:-0.6rem;border:1px solid var(--border);'
            'border-top:none;border-radius:0 0 12px 12px;'
            'background:var(--card);overflow:hidden;margin-bottom:0.85rem;">',
            unsafe_allow_html=True,
        )

        for cable_name, cable_seq, adapter_options in MONITOR_PERIPHERALS:
            ckey = f"cable_mon{mon_idx}_{_safe_periph_key(cable_name)}"
            if ckey not in st.session_state:
                st.session_state[ckey] = False

            st.markdown('<div style="border-bottom:1px solid var(--border);">', unsafe_allow_html=True)
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
    if "csv_file_id" not in st.session_state and DEFAULT_CSV_PATH.exists():
        try:
            df = pd.read_csv(DEFAULT_CSV_PATH, encoding="utf-8-sig")
            st.session_state["csv_file_id"] = "__default__"
            st.session_state["df_loaded"] = df
        except Exception:
            pass

    render_header()

    st.markdown(
        '<div style="font-size:0.7rem;font-weight:700;color:var(--muted);'
        'text-transform:uppercase;letter-spacing:0.07em;margin-bottom:0.3rem;">'
        'Form Type</div>',
        unsafe_allow_html=True,
    )
    form_type = render_form_type_selector()

    if form_type == "wfh":
        tpl_path   = _find_template(WFH_TEMPLATE_NAME)
        type_label = "Work From Home"
        type_color_light = "#1565c0"
        type_bg_light    = "#e8f0fc"
    else:
        tpl_path   = _find_template(ONSITE_TEMPLATE_NAME)
        type_label = "Work On Site"
        type_color_light = "#00796b"
        type_bg_light    = "#e0f2f1"

    if tpl_path is None:
        st.error(
            f"Template not found for **{type_label}**. "
            "Place the `.dotx` file inside the **src/** folder and refresh."
        )
    else:
        st.markdown(
            f'<div style="font-size:0.75rem;'
            f'color:{type_color_light};background:{type_bg_light};'
            f'border-left:3px solid {type_color_light};border-radius:0 6px 6px 0;'
            f'padding:0.38rem 0.7rem;margin:0.5rem 0 0.8rem;">'
            f'Using template: <strong>{type_label}</strong>'
            f'</div>',
            unsafe_allow_html=True,
        )

    st.markdown("<div style='height:0.3rem'></div>", unsafe_allow_html=True)
    copy_type = render_copy_type_selector()

    st.markdown("<hr>", unsafe_allow_html=True)

    prepared_by = render_prepared_by_section()
    st.markdown("<div style='height:0.4rem'></div>", unsafe_allow_html=True)

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

    step_open(3, "Select Assets",
              "Choose the items from the asset list to include on the form.")

    sel_key = f"sel_{chosen_name.lower().replace(' ', '_')}"
    row_widget_keys = {idx: f"chk_{idx}_{sel_key}" for idx in df_filtered.index}

    for wkey in row_widget_keys.values():
        if wkey not in st.session_state:
            st.session_state[wkey] = False

    col_a, col_b, _ = st.columns([1, 1, 5])
    with col_a:
        st.markdown('<div class="btn-outline-blue">', unsafe_allow_html=True)
        if st.button("Select All", key="sp_select_all"):
            for wkey in row_widget_keys.values():
                st.session_state[wkey] = True
            st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)
    with col_b:
        st.markdown('<div class="btn-outline-blue">', unsafe_allow_html=True)
        if st.button("Clear All", key="sp_clear_all"):
            for wkey in row_widget_keys.values():
                st.session_state[wkey] = False
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

        wkey = row_widget_keys[idx]
        val  = st.checkbox(label, key=wkey)
        if val:
            checked.append(idx)

    step_close()

    df_selected = df_filtered.loc[checked].copy()

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

    step_open(4, "Monitor Cables & Adapters",
              "A charger is automatically included. Select cables and adapters for each monitor.")
    st.markdown('<div class="charger-badge">✅ Charger automatically included</div>', unsafe_allow_html=True)

    monitor_cable_assignments: list[dict] = []
    if selected_monitors:
        st.caption("Tick which cables came with each monitor, then pick an adapter if needed.")
        for mon in selected_monitors:
            assignments = render_monitor_cable_block(mon["idx"], mon["label"])
            monitor_cable_assignments.extend(assignments)
    elif not df_selected.empty:
        st.markdown(
            '<div style="font-size:0.8rem;color:var(--muted);padding:0.4rem 0;">'
            'No monitors in the selected assets — cable section not applicable.'
            '</div>',
            unsafe_allow_html=True,
        )

    step_close()

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

    step_open(6, "Generate Document", "")

    copy_label_display = "IT Copy" if copy_type == "it" else "Employee Copy"
    st.markdown(
        f'<div class="info-hint">'
        f'<strong>Prepared by:</strong> {prepared_by} &nbsp;·&nbsp; '
        f'<strong>Employee:</strong> {chosen_name} &nbsp;·&nbsp; '
        f'<strong>Items:</strong> {len(sorted_rows)} &nbsp;·&nbsp; '
        f'<strong>Form:</strong> {type_label} &nbsp;·&nbsp; '
        f'<strong>Copy:</strong> {copy_label_display}'
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
                    copy_type=copy_type,
                )
                st.session_state["docx"]            = docx_bytes
                st.session_state["form_name"]       = form_name
                st.session_state["form_client"]     = form_client
                st.session_state["form_date"]       = form_date
                st.session_state["prepared_by"]     = prepared_by
                st.session_state["form_type_used"]  = form_type
                st.session_state["copy_type_used"]  = copy_type
                st.success("Document ready — click Download below.")
            except Exception as e:
                st.error(f"Error: {e}")

    if "docx" in st.session_state:
        _fname      = st.session_state.get("form_name")      or "Employee"
        _fclient    = st.session_state.get("form_client")    or ""
        _fdate      = st.session_state.get("form_date")      or datetime.now()
        _ftype_used = st.session_state.get("form_type_used") or "wfh"
        _fcopy_used = st.session_state.get("copy_type_used") or "it"
        _client_p   = f" ({_fclient})" if _fclient else ""
        _day        = str(int(_fdate.strftime("%d")))
        _date_p     = _day + _fdate.strftime(" %B %Y")
        _type_p     = "WFH" if _ftype_used == "wfh" else "On Site"
        _copy_p     = "IT Copy" if _fcopy_used == "it" else "Employee Copy"
        _filename   = (
            f"{_copy_p} - Equipment Accountability Form"
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
