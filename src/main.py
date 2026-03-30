"""
Equipment Accountability Form Generator
BMG Outsourcing — redesigned UI
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
# PATHS
# ─────────────────────────────────────────────

BASE_DIR      = Path(__file__).parent
TEMPLATE_PATH = BASE_DIR / "src" / "Equipment Accountability Form (Work From Home).dotx"
LOGO_PATH     = BASE_DIR / "images" / "logo.png"

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
    if LOGO_PATH.exists():
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
    max-width: 800px !important;
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

/* ── Small outline buttons (Select All / Clear All) ── */
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

/* ── Caption / small text ── */
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
.search-hint strong {{
    font-weight: 700;
    color: var(--navy);
}}

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

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

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

def get_template_path() -> Path | None:
    return TEMPLATE_PATH if TEMPLATE_PATH.exists() else None

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
    if exact_matches == n_q:                          return (1, coverage)
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

_FONT_NAME  = "Times New Roman"
_FONT_SIZE  = "20"
_TABLE_SIZE = "18"

def _make_rPr(bold: bool = False, size: str = _FONT_SIZE) -> etree._Element:
    rPr = etree.Element(f"{{{W}}}rPr")
    if bold:
        etree.SubElement(rPr, f"{{{W}}}b")
        etree.SubElement(rPr, f"{{{W}}}bCs")
    fonts = etree.SubElement(rPr, f"{{{W}}}rFonts")
    fonts.set(f"{{{W}}}ascii",     _FONT_NAME)
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

def _set_sdt_value(body, tag_val: str, new_text: str, bold: bool = True) -> bool:
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
        p = etree.SubElement(sdtContent, f"{{{W}}}p")
        r = etree.SubElement(p,          f"{{{W}}}r")
        r.append(_make_rPr(bold=bold, size=_FONT_SIZE))
        r.append(_make_t(new_text))
        return True
    return False

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
    r_el = etree.SubElement(p_el, f"{{{W}}}r")
    r_el.append(_make_rPr(bold=True, size=_TABLE_SIZE))
    r_el.append(_make_t(text))

def _compact_row(row_el):
    trPr = row_el.find(f"{{{W}}}trPr")
    if trPr is not None:
        for trH in trPr.findall(f"{{{W}}}trHeight"):
            trPr.remove(trH)
    for tc in row_el.iter(f"{{{W}}}tc"):
        tcPr = tc.find(f"{{{W}}}tcPr")
        if tcPr is None:
            tcPr = etree.SubElement(tc, f"{{{W}}}tcPr")
            tc.insert(0, tcPr)
        tcMar = tcPr.find(f"{{{W}}}tcMar")
        if tcMar is None:
            tcMar = etree.SubElement(tcPr, f"{{{W}}}tcMar")
        for side, val in [("top", "0"), ("bottom", "0"), ("left", "108"), ("right", "108")]:
            el = tcMar.find(f"{{{W}}}{side}")
            if el is None:
                el = etree.SubElement(tcMar, f"{{{W}}}{side}")
            el.set(f"{{{W}}}w",    val)
            el.set(f"{{{W}}}type", "dxa")
        for p in tc.iter(f"{{{W}}}p"):
            pPr = p.find(f"{{{W}}}pPr")
            if pPr is not None:
                for spacing in pPr.findall(f"{{{W}}}spacing"):
                    pPr.remove(spacing)
                for cs in pPr.findall(f"{{{W}}}contextualSpacing"):
                    pPr.remove(cs)

def _compact_page_margins(body):
    sectPr = body.find(f"{{{W}}}sectPr")
    if sectPr is None:
        return
    pgMar = sectPr.find(f"{{{W}}}pgMar")
    if pgMar is None:
        pgMar = etree.SubElement(sectPr, f"{{{W}}}pgMar")
    limits = {"top": 720, "bottom": 720, "left": 900, "right": 900}
    for side, max_val in limits.items():
        current = pgMar.get(f"{{{W}}}{side}")
        try:
            if current is None or int(current) > max_val:
                pgMar.set(f"{{{W}}}{side}", str(max_val))
        except ValueError:
            pgMar.set(f"{{{W}}}{side}", str(max_val))

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

def fill_template(assets_df, col_map, employee_name, client, position, date_str) -> bytes:
    tpl = get_template_path()
    if tpl is None:
        raise FileNotFoundError("Template .dotx not found.")
    with zipfile.ZipFile(io.BytesIO(tpl.read_bytes())) as zin:
        files = {n: zin.read(n) for n in zin.namelist()}
    if "[Content_Types].xml" in files:
        files["[Content_Types].xml"] = _patch_content_types(files["[Content_Types].xml"])
    if "docProps/app.xml" in files:
        files["docProps/app.xml"] = _patch_app_xml(files["docProps/app.xml"])
    root = etree.fromstring(files["word/document.xml"])
    body = root.find(f"{{{W}}}body")
    _set_sdt_value(body, "Name",        employee_name, bold=True)
    _set_sdt_value(body, "Client",      client,        bold=True)
    _set_sdt_value(body, "Position",    position,      bold=True)
    _set_sdt_value(body, "Date",        date_str,      bold=True)
    _set_sdt_value(body, "Contact No.", "",            bold=True)
    _compact_page_margins(body)
    eq_table = _get_equipment_table(body)
    if eq_table is not None:
        all_rows     = eq_table.findall(f"{{{W}}}tr")
        data_rows    = all_rows[1:]
        template_row = copy.deepcopy(data_rows[0]) if data_rows else None
        num_assets   = len(assets_df)
        for i, (_, row) in enumerate(assets_df.iterrows()):
            c         = col_map
            content   = safe_str(row.get(c.get("content_type") or "", ""))
            brand     = safe_str(row.get(c.get("brand")        or "", ""))
            model     = safe_str(row.get(c.get("model")        or "", ""))
            serial    = safe_str(row.get(c.get("serial")       or "", ""))
            asset_tag = safe_str(row.get(c.get("asset_tag")    or "", ""))
            remarks   = safe_str(row.get(c.get("remarks")      or "", ""))
            parts     = [p for p in [content, model] if p]
            equipment = " ".join(parts) + (f" ({brand})" if brand else "")
            if not equipment:
                equipment = asset_tag
            if i < len(data_rows):
                row_el = data_rows[i]
                _compact_row(row_el)
                _fill_equipment_row(row_el, equipment, serial, asset_tag, remarks)
            elif template_row is not None:
                new_row = copy.deepcopy(template_row)
                _compact_row(new_row)
                eq_table.append(new_row)
                _fill_equipment_row(new_row, equipment, serial, asset_tag, remarks)
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

# ─────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────

def main():
    render_header()

    # ── Step 1: Upload ────────────────────────
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
    user_col   = col_map.get("current_user")
    client_col = col_map.get("client")

    if not user_col:
        st.error('Could not find a "Current User" column. Make sure your CSV was exported from SharePoint with standard column names.')
        return

    st.success(f"✓ &nbsp;**{len(df):,}** records loaded from **{uploaded.name}**")

    # ── Step 2: Search ────────────────────────
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
    total_found    = len(results)

    st.caption(f"{total_found} result(s) — select an employee to continue")

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

    # ── Step 3: Select assets ─────────────────
    step_open(3, "Select Assets",
              "Choose the equipment items to include on the accountability form.")

    sel_key = f"sel_{chosen_name.lower().replace(' ', '_')}"
    if sel_key not in st.session_state:
        st.session_state[sel_key] = {idx: False for idx in df_filtered.index}
    for idx in df_filtered.index:
        if idx not in st.session_state[sel_key]:
            st.session_state[sel_key][idx] = False

    col_a, col_b, _ = st.columns([1, 1, 5])
    with col_a:
        st.markdown('<div class="small-btn">', unsafe_allow_html=True)
        if st.button("Select All"):
            for idx in df_filtered.index:
                st.session_state[sel_key][idx] = True
            st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)
    with col_b:
        st.markdown('<div class="small-btn">', unsafe_allow_html=True)
        if st.button("Clear All"):
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
        desc  = " ".join(p for p in [ct, model] if p)
        if brand:
            desc += f" ({brand})"
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

    if not checked:
        st.info("Select at least one asset to continue.")
        return

    st.caption(f"**{len(checked)}** asset(s) selected")

    # ── Step 4: Form details ──────────────────
    step_open(4, "Form Details",
              "These fields appear in the header of the generated document. Edit as needed.")

    default_client, default_pos = "", ""
    if not df_selected.empty:
        if client_col:
            default_client = safe_str(df_selected.iloc[0].get(client_col, ""))
        for c in df_selected.columns:
            if any(kw in c.lower() for kw in ("position", "job title", "designation")):
                default_pos = safe_str(df_selected.iloc[0].get(c, ""))
                break

    f1, f2 = st.columns(2)
    with f1:
        form_name     = st.text_input("Full Name", value=chosen_name)
        form_client   = st.text_input("Client",    value=default_client)
    with f2:
        form_position = st.text_input("Position",  value=default_pos)
        form_date     = st.date_input("Date",      value=datetime.today())

    step_close()
    form_date_str = form_date.strftime("%B %d, %Y")

    # ── Step 5: Generate ──────────────────────
    step_open(5, "Generate Document", "")

    tpl = get_template_path()
    if tpl is None:
        st.error("Word template (.dotx) not found. Make sure the file is inside the src/ folder.")
        step_close()
        return

    if st.button("⬇  Generate Word Document", use_container_width=True, type="primary"):
        with st.spinner("Filling the form…"):
            try:
                docx = fill_template(
                    df_selected, col_map,
                    form_name, form_client, form_position, form_date_str,
                )
                st.session_state["docx"]        = docx
                st.session_state["form_name"]   = form_name
                st.session_state["form_client"] = form_client
                st.session_state["form_date"]   = form_date
                st.success("✓ &nbsp;Document ready — click Download below.")
            except Exception as e:
                st.error(f"Error: {e}")

    if "docx" in st.session_state:
        _fname       = st.session_state.get("form_name")   or "Employee"
        _fclient     = st.session_state.get("form_client") or ""
        _fdate       = st.session_state.get("form_date")   or datetime.now()
        _client_part = f" ({_fclient})" if _fclient else ""
        _day         = str(int(_fdate.strftime("%d")))
        _date_part   = _day + _fdate.strftime(" %B %Y")
        _filename    = (
            f"Employee Copy - Equipment Accountability Form"
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

    # ── Summary metrics ───────────────────────
    st.markdown("<div style='height:.4rem'></div>", unsafe_allow_html=True)
    c1, c2, c3 = st.columns(3)
    c1.metric("Total Records", f"{len(df):,}")
    c2.metric("Matched",       f"{len(df_filtered):,}")
    c3.metric("On This Form",  f"{len(df_selected):,}")

    st.markdown(
        '<div class="bmg-footer">BMG Outsourcing, Inc. &nbsp;·&nbsp; Asset Accountability System</div>',
        unsafe_allow_html=True,
    )


if __name__ == "__main__":
    main()