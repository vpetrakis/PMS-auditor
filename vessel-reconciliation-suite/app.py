import streamlit as st
import pandas as pd
import numpy as np
import io
import re
import hashlib
from datetime import datetime
import base64
import traceback
import warnings
from zipfile import ZipFile
from xml.etree import ElementTree as ET

warnings.filterwarnings("ignore")

st.set_page_config(page_title="POSEIDON | Recon Suite", page_icon="⚓", layout="wide", initial_sidebar_state="collapsed")

# ═══════════════════════════════════════════════════════════════════════════════
# 1. ENTERPRISE PAGE CONFIGURATION & POSEIDON TITAN CSS
# ═══════════════════════════════════════════════════════════════════════════════

def _u(s):
    return f"data:image/svg+xml;base64,{base64.b64encode(s.encode()).decode()}"

LOGO_SVG = base64.b64encode(b'<svg viewBox="0 0 48 48" xmlns="http://www.w3.org/2000/svg"><defs><linearGradient id="pg" x1="0" y1="0" x2="1" y2="1"><stop offset="0%" stop-color="#c9a84c"/><stop offset="50%" stop-color="#00e0b0"/><stop offset="100%" stop-color="#fff"/></linearGradient></defs><circle cx="24" cy="24" r="22" fill="none" stroke="url(#pg)" stroke-width="0.8" opacity=".3"/><path d="M24 6L24 42" stroke="url(#pg)" stroke-width="1.5" stroke-linecap="round"/><path d="M12 24Q24 32 36 24" fill="none" stroke="url(#pg)" stroke-width="1.5" stroke-linecap="round"/></svg>').decode()

ICONS = {
    "VERIFIED": _u('<svg viewBox="0 0 28 28" xmlns="http://www.w3.org/2000/svg"><circle cx="14" cy="14" r="12" fill="none" stroke="#00e0b0" stroke-width="1" opacity=".2"/><circle cx="14" cy="14" r="7.5" fill="#061a14" stroke="#00e0b0" stroke-width="1.5"/><polyline points="10,14.5 12.8,17 18,10.5" fill="none" stroke="#00e0b0" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"/></svg>'),
    "DRIFT": _u('<svg viewBox="0 0 28 28" xmlns="http://www.w3.org/2000/svg"><circle cx="14" cy="14" r="12" fill="none" stroke="#ff2a55" stroke-width="1" stroke-dasharray="4 3"/><circle cx="14" cy="14" r="7.5" fill="#1a0508" stroke="#ff2a55" stroke-width="1.5"/><g stroke="#ff2a55" stroke-width="2.5" stroke-linecap="round"><line x1="11" y1="11" x2="17" y2="17"/><line x1="17" y1="11" x2="11" y2="17"/></g></svg>'),
    "CLOCK": _u('<svg viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg"><path fill="#3b82f6" d="M11.99 2C6.47 2 2 6.48 2 12s4.47 10 9.99 10C17.52 22 22 17.52 22 12S17.52 2 11.99 2zM12 20c-4.42 0-8-3.58-8-8s3.58-8 8-8 8 3.58 8 8-3.58 8-8 8zm.5-13H11v6l5.25 3.15.75-1.23-4.5-2.67z"/></svg>'),
    "SEAL": _u('<svg viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg"><path fill="#7b68ee" d="M18 8h-1V6c0-2.76-2.24-5-5-5S7 3.24 7 6v2H6c-1.1 0-2 .9-2 2v10c0 1.1.9 2 2 2h12c1.1 0 2-.9 2-2V10c0-1.1-.9-2-2-2zM9 6c0-1.66 1.34-3 3-3s3 1.34 3 3v2H9V6zm9 14H6V10h12v10zm-6-3c1.1 0 2-.9 2-2s-.9-2-2-2-2 .9-2 2 .9 2 2 2z"/></svg>')
}

st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;800&family=JetBrains+Mono:wght@400;700&display=swap');
    .stApp { background-color: #060b13; font-family: 'Inter', sans-serif; color: #f8fafc; }
    #MainMenu {visibility: hidden;} footer {visibility: hidden;} header {visibility: hidden;}

    :root { --glass-bg: rgba(13, 21, 34, 0.6); --glass-border: rgba(255,255,255,0.05); }

    .hero { display: flex; align-items: center; justify-content: space-between; padding-bottom: 20px; border-bottom: 1px solid var(--glass-border); margin-bottom: 30px; }
    .hero-left { display: flex; align-items: center; gap: 20px; }
    .hero-logo { width: 55px; height: 55px; }
    .hero-title { font-size: 2.2rem; font-weight: 800; background: linear-gradient(90deg, #ffffff, #8ba1b5); -webkit-background-clip: text; -webkit-text-fill-color: transparent; letter-spacing: -1px; line-height: 1.1;}
    .hero-sub { font-size: 0.85rem; color: #64748b; font-weight: 600; text-transform: uppercase; letter-spacing: 2px; }
    .hero-badge { text-align: right; font-family: 'JetBrains Mono', monospace; font-size: 0.75rem; color: #64748b; line-height: 1.6; }

    .stFileUploader > div > div { background-color: var(--glass-bg) !important; border: 1px dashed rgba(255,255,255,0.1) !important; border-radius: 12px !important; transition: all 0.3s ease; }
    .stFileUploader > div > div:hover { border-color: #00e0b0 !important; background-color: rgba(0, 224, 176, 0.05) !important; }

    .hud-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(220px, 1fr)); gap: 20px; margin-bottom: 30px; }
    .hud-card { background: rgba(13, 21, 34, 0.8); border: 1px solid var(--glass-border); border-radius: 8px; padding: 20px; display: flex; flex-direction: column; box-shadow: 0 10px 30px -10px rgba(0,0,0,0.5); }
    .hud-header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 10px; }
    .hud-title { font-size: 0.75rem; color: #64748b; text-transform: uppercase; letter-spacing: 1.2px; font-weight: 600; }
    .hud-icon img { width: 24px; height: 24px; }
    .hud-val { font-size: 2rem; font-weight: 800; color: #ffffff; line-height: 1.1; font-family: 'JetBrains Mono', monospace; }
    .hud-sub { font-size: 0.75rem; color: #475569; margin-top: 5px; font-weight: 500; }

    .stTabs [data-baseweb="tab-list"] { gap: 24px; }
    .stTabs [data-baseweb="tab"] { height: 50px; white-space: pre-wrap; background-color: transparent; border-radius: 4px; color: #64748b; font-weight: 600; }
    .stTabs [aria-selected="true"] { color: #00e0b0; border-bottom: 2px solid #00e0b0; }
    </style>
""", unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════════════════════
# 2. HELPERS
# ═══════════════════════════════════════════════════════════════════════════════

def normalize_text(s):
    s = str(s).upper().strip()
    s = s.replace("\xa0", " ")
    s = re.sub(r"\s+", " ", s)
    return s

def extract_first_float(value):
    if pd.isna(value):
        return np.nan
    s = str(value).replace(",", "")
    m = re.search(r"-?\d+(?:\.\d+)?", s)
    return float(m.group()) if m else np.nan

def robust_parse_date(value):
    if pd.isna(value):
        return pd.NaT
    s = str(value).strip()
    if not s:
        return pd.NaT

    direct = pd.to_datetime(s, errors="coerce", dayfirst=True)
    if pd.notna(direct):
        return direct

    cleaned = re.sub(r"\s+", " ", s)
    fmts = [
        "%Y-%m-%d %H%M%S",
        "%Y-%m-%d",
        "%d-%m-%Y",
        "%d/%m/%Y",
        "%d.%m.%Y",
        "%d-%b-%Y",
        "%d-%B-%Y",
    ]
    for fmt in fmts:
        try:
            return pd.to_datetime(cleaned, format=fmt, errors="raise", dayfirst=True)
        except Exception:
            pass
    return pd.NaT

def is_date_like(value):
    return pd.notna(robust_parse_date(value))

def find_target_sheet(xls):
    preferred = []
    for s in xls.sheet_names:
        su = s.upper()
        score = 0
        if "DAILY" in su:
            score += 3
        if "OPERATING" in su:
            score += 3
        if "HOURS" in su:
            score += 2
        if score > 0:
            preferred.append((score, s))
    if preferred:
        preferred.sort(reverse=True)
        return preferred[0][1]
    return xls.sheet_names[0]

# ═══════════════════════════════════════════════════════════════════════════════
# 3. EXCEL PARSER - FIXED WITHOUT BREAKING APP INTEGRITY
# ═══════════════════════════════════════════════════════════════════════════════

@st.cache_data(show_spinner=False)
def parse_daily_hours_excel(file_bytes):
    """
    Robust parser for the DAILY OPERATING HOURS workbook.
    Returns:
      - clean_df with Date + ME_Hours + DG1_Hours + DG2_Hours + DG3_Hours
      - raw diagnostic sheet
      - col_map used in extraction
    """
    try:
        xls = pd.ExcelFile(io.BytesIO(file_bytes), engine="openpyxl")
        target_sheet = find_target_sheet(xls)
        df_raw = pd.read_excel(io.BytesIO(file_bytes), sheet_name=target_sheet, header=None, dtype=object, engine="openpyxl")
    except Exception:
        try:
            df_raw = pd.read_excel(io.BytesIO(file_bytes), header=None, dtype=object, engine="openpyxl")
        except Exception:
            return pd.DataFrame(), pd.DataFrame(), {}

    if df_raw.empty:
        return pd.DataFrame(), df_raw, {}

    diag_raw = df_raw.copy()

    header_scan_rows = min(35, len(df_raw))
    header_block = df_raw.head(header_scan_rows).copy().ffill(axis=1).ffill(axis=0)

    date_col = None
    for col_idx in range(header_block.shape[1]):
        col_text = " | ".join([normalize_text(x) for x in header_block.iloc[:, col_idx].tolist() if pd.notna(x)])
        if "DATE" in col_text or re.search(r"\bDAY\b", col_text):
            date_col = col_idx
            break

    if date_col is None:
        for col_idx in range(df_raw.shape[1]):
            probe = df_raw.iloc[:80, col_idx]
            hits = sum(is_date_like(v) for v in probe)
            if hits >= 3:
                date_col = col_idx
                break

    col_map = {}
    if date_col is not None:
        col_map["Date"] = date_col

    first_date_row = None
    if date_col is not None:
        for i in range(len(df_raw)):
            if is_date_like(df_raw.iloc[i, date_col]):
                first_date_row = i
                break

    if first_date_row is None:
        return pd.DataFrame(), diag_raw, col_map

    header_rows_to_use = list(range(max(0, first_date_row - 4), first_date_row + 1))
    working_header = df_raw.iloc[header_rows_to_use, :].copy().ffill(axis=1).ffill(axis=0)

    def classify_col(col_idx):
        vals = [normalize_text(v) for v in working_header.iloc[:, col_idx].tolist() if pd.notna(v)]
        txt = " | ".join(vals)
        if "MAIN ENGINE" in txt or "M/E" in txt:
            return "ME_Hours"
        if "DIESEL GENERATOR NO.1" in txt or "DIESEL GENERATOR NO 1" in txt or "DG 1" in txt or "D/G 1" in txt:
            return "DG1_Hours"
        if "DIESEL GENERATOR NO.2" in txt or "DIESEL GENERATOR NO 2" in txt or "DG 2" in txt or "D/G 2" in txt:
            return "DG2_Hours"
        if "DIESEL GENERATOR NO.3" in txt or "DIESEL GENERATOR NO 3" in txt or "DG 3" in txt or "D/G 3" in txt:
            return "DG3_Hours"
        return None

    found_cols = {"ME_Hours": [], "DG1_Hours": [], "DG2_Hours": [], "DG3_Hours": []}
    for col_idx in range(df_raw.shape[1]):
        cls = classify_col(col_idx)
        if cls:
            found_cols[cls].append(col_idx)

    for key, candidates in found_cols.items():
        if candidates:
            col_map[key] = min(candidates)

    df = df_raw.iloc[first_date_row:].copy().reset_index(drop=True)
    clean_df = pd.DataFrame()
    clean_df["Date"] = df.iloc[:, col_map["Date"]].apply(robust_parse_date)

    for sys_col in ["ME_Hours", "DG1_Hours", "DG2_Hours", "DG3_Hours"]:
        if sys_col in col_map:
            clean_df[sys_col] = df.iloc[:, col_map[sys_col]].apply(extract_first_float)
        else:
            clean_df[sys_col] = np.nan

    raw_date_col = df.iloc[:, col_map["Date"]].astype(str).str.upper()
    non_total_mask = ~raw_date_col.str.contains("TOTAL MONTHLY", na=False)
    clean_df = clean_df[non_total_mask].copy()

    clean_df = clean_df[clean_df["Date"].notna()].copy()

    hour_cols = [c for c in ["ME_Hours", "DG1_Hours", "DG2_Hours", "DG3_Hours"] if c in clean_df.columns]
    for c in hour_cols:
        clean_df[c] = pd.to_numeric(clean_df[c], errors="coerce").fillna(0.0)

    clean_df = clean_df.reset_index(drop=True)
    return clean_df, diag_raw, col_map

# ═══════════════════════════════════════════════════════════════════════════════
# 4. DOC / DOCX PARSER
# ═══════════════════════════════════════════════════════════════════════════════

@st.cache_data(show_spinner=False)
def parse_pms_binary_doc(file_bytes, file_name=""):
    """
    Supports legacy .doc parsing and safe .docx text extraction.
    Keeps the original return contract: (clean_df, raw_cells_df)
    """
    fname = (file_name or "").lower()

    def parse_doc_binary(bytes_):
        raw_text = bytes_.decode("latin-1", errors="ignore").replace("\x00", "")
        cells = [c.strip() for c in raw_text.split("\x07") if c.strip()]
        return cells

    def parse_docx_xml(bytes_):
        texts = []
        try:
            with ZipFile(io.BytesIO(bytes_)) as z:
                xml = z.read("word/document.xml")
            root = ET.fromstring(xml)
            ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
            for para in root.findall(".//w:p", ns):
                runs = []
                for t in para.findall(".//w:t", ns):
                    if t.text:
                        runs.append(t.text)
                line = " ".join(runs).strip()
                if line:
                    texts.append(line)
        except Exception:
            pass
        return texts

    if fname.endswith(".docx"):
        cells = parse_docx_xml(file_bytes)
    else:
        cells = parse_doc_binary(file_bytes)

    extracted_data = []

    COMPONENTS = [
        'CYLINDER COVER', 'PISTON ASSY', 'PISTON ASSEMBLY', 'STUFFING BOX', 'PISTON CROWN', 'CYLINDER LINER',
        'EXHAUST VALVE', 'EXAUST VALVE', 'STARTING VALVE', 'SAFETY VALVE', 'FUEL VALVE', 'FUEL VALVES',
        'FUEL PUMP', 'SUCTION VALVE', 'PUNCTURE VALVE', 'CROSSHEAD BEARING', 'CROSSHEAD BEARINGS',
        'BOTTOM END BEARING', 'BOTTOM END BEARINGS', 'MAIN BEARING', 'MAIN BEARINGS', 'CYLINDER HEAD',
        'PISTON', 'CONNECTING ROD', 'TURBOCHARGER', 'AIR COOLER', 'COOLING WATER PUMP', 'THERMOSTAT VALVE',
        'THRUST BEARING'
    ]

    system = "ME"
    sub_units = ["Cyl 1", "Cyl 2", "Cyl 3", "Cyl 4", "Cyl 5", "Cyl 6"]
    target_timeline = "ME_Hours"

    i = 0
    while i < len(cells):
        cell = normalize_text(cells[i])

        if "MAIN ENGINE" in cell:
            system = "ME"
            target_timeline = "ME_Hours"
            sub_units = ["Cyl 1", "Cyl 2", "Cyl 3", "Cyl 4", "Cyl 5", "Cyl 6"]
        elif "AUX. ENGINE" in cell or "AUX ENGINE" in cell or "D/G" in cell or "DG " in cell or "DIESEL GENERATOR" in cell:
            system = "DG"
            sub_units = ["DG1", "DG2", "DG3"]

        is_comp = any(c in cell for c in COMPONENTS) and len(cell) < 80
        if is_comp:
            comp_name = cell
            dates, hours = [], []

            j = i + 1
            while j < min(i + 20, len(cells)):
                if str(cells[j]).strip() == '1':
                    for k in range(len(sub_units)):
                        if j + 1 + k < len(cells):
                            dates.append(cells[j + 1 + k])
                    break
                j += 1

            j = i + 1
            while j < min(i + 35, len(cells)):
                if str(cells[j]).strip() == '2':
                    for k in range(len(sub_units)):
                        if j + 1 + k < len(cells):
                            hours.append(cells[j + 1 + k])
                    break
                j += 1

            for idx, su in enumerate(sub_units):
                if idx < len(dates) and idx < len(hours):
                    d = robust_parse_date(dates[idx])
                    h = extract_first_float(hours[idx])
                    if pd.notna(d) and pd.notna(h):
                        curr_timeline = target_timeline if system == 'ME' else f"{su}_Hours"
                        extracted_data.append({
                            'System': system,
                            'Target_Timeline': curr_timeline,
                            'Component': f"[{system}] {comp_name} ({su})",
                            'Last_Overhaul': d,
                            'Claimed_Hours': float(h)
                        })
        i += 1

    raw_cells_df = pd.DataFrame(cells, columns=["ASCII Extracted Text Blocks"])
    if extracted_data:
        clean_df = pd.DataFrame(extracted_data).dropna(subset=['Last_Overhaul']).reset_index(drop=True)
        return clean_df, raw_cells_df

    return pd.DataFrame(), raw_cells_df

# ═══════════════════════════════════════════════════════════════════════════════
# 5. SAFE TARGET-TIMELINE RESOLUTION
# ═══════════════════════════════════════════════════════════════════════════════

def resolve_timeline_column(target_col, component_name):
    valid = {"ME_Hours", "DG1_Hours", "DG2_Hours", "DG3_Hours"}
    if target_col in valid:
        return target_col

    comp_u = str(component_name).upper()
    if "DG1" in comp_u or "(DG1)" in comp_u:
        return "DG1_Hours"
    if "DG2" in comp_u or "(DG2)" in comp_u:
        return "DG2_Hours"
    if "DG3" in comp_u or "(DG3)" in comp_u:
        return "DG3_Hours"
    return "ME_Hours"

# ═══════════════════════════════════════════════════════════════════════════════
# 6. MAIN FRONTEND ORCHESTRATOR
# ═══════════════════════════════════════════════════════════════════════════════

st.markdown(f"""
<div class="hero">
    <div class="hero-left">
        <img src="data:image/svg+xml;base64,{LOGO_SVG}" class="hero-logo" alt=""/>
        <div>
            <div class="hero-title">POSEIDON RECON</div>
            <div class="hero-sub">Enterprise Forensic Auditor</div>
        </div>
    </div>
    <div class="hero-badge">
        <span style="color:#00e0b0">KERNEL</span>&ensp;Zero-Trust Triangulation<br>
        <span style="color:#00e0b0">DECODER</span>&ensp;X-Ray Matrix Sweep<br>
        <span style="color:#fff">BUILD</span>&ensp;v10.2.0 Stable
    </div>
</div>
""", unsafe_allow_html=True)

col1, col2 = st.columns(2)
with col1:
    st.markdown("<div style='color:#8ba1b5; font-size:0.9rem; font-weight:600; margin-bottom:10px;'>1. COMPONENT OVERHAULS (.doc / .docx)</div>", unsafe_allow_html=True)
    pms_file = st.file_uploader("Upload Word Binary", type=["doc", "docx"], key="pms_box")

with col2:
    st.markdown("<div style='color:#8ba1b5; font-size:0.9rem; font-weight:600; margin-bottom:10px;'>2. DAILY OPERATING HOURS (Excel)</div>", unsafe_allow_html=True)
    logs_file = st.file_uploader("Upload Excel Timeline", type=["xlsx", "xls"], key="logs_box")

if pms_file and logs_file:
    with st.spinner("Executing X-Ray Matrix Extractor & Triangulation..."):
        try:
            pms_df, diag_pms = parse_pms_binary_doc(pms_file.getvalue(), pms_file.name)
            timeline_df, diag_timeline, col_map = parse_daily_hours_excel(logs_file.getvalue())

            total_days = len(timeline_df) if not timeline_df.empty else 0
            audit_results = []

            if not timeline_df.empty and not pms_df.empty:
                for _, row in pms_df.iterrows():
                    comp = row['Component']
                    oh_date = row['Last_Overhaul']
                    legacy_hrs = row['Claimed_Hours']
                    target_col = resolve_timeline_column(row['Target_Timeline'], comp)

                    if target_col in timeline_df.columns:
                        mask = timeline_df['Date'] >= oh_date
                        verified_hrs = timeline_df.loc[mask, target_col].sum()
                        delta = verified_hrs - legacy_hrs

                        audit_results.append({
                            "Component": comp,
                            "Overhaul Date": oh_date.strftime('%d-%b-%Y') if pd.notna(oh_date) else "",
                            "Claimed (Doc)": int(round(float(legacy_hrs))) if pd.notna(legacy_hrs) else 0,
                            "Verified (Excel)": int(round(float(verified_hrs))) if pd.notna(verified_hrs) else 0,
                            "Delta (Drift)": int(round(float(delta))) if pd.notna(delta) else 0,
                            "Status": "VERIFIED" if int(round(float(delta))) == 0 else "DRIFT DETECTED"
                        })
                    else:
                        audit_results.append({
                            "Component": comp,
                            "Overhaul Date": oh_date.strftime('%d-%b-%Y') if pd.notna(oh_date) else "",
                            "Claimed (Doc)": int(round(float(legacy_hrs))) if pd.notna(legacy_hrs) else 0,
                            "Verified (Excel)": 0,
                            "Delta (Drift)": 0,
                            "Status": "NO TIMELINE DATA"
                        })

            st.markdown("<br>", unsafe_allow_html=True)
            t1, t2, t3 = st.tabs(["📊 IMMUTABLE LEDGER", "📈 FORENSIC PLOT", "🔎 GLASS-BOX DIAGNOSTICS"])

            with t1:
                if audit_results:
                    res_df = pd.DataFrame(audit_results)
                    errors_corrected = len(res_df[res_df['Status'] == 'DRIFT DETECTED'])
                    digital_seal = hashlib.sha256(res_df.to_json(orient='records').encode()).hexdigest()

                    st.markdown(f"""
                    <div class="hud-grid">
                        <div class="hud-card" style="border-bottom: 3px solid #00e0b0;">
                            <div class="hud-header">
                                <div class="hud-title">Components Audited</div>
                                <div class="hud-icon"><img src="{ICONS['VERIFIED']}"></div>
                            </div>
                            <div class="hud-val">{len(res_df)}</div>
                            <div class="hud-sub">Extracted via Matrix Resolver</div>
                        </div>
                        <div class="hud-card" style="border-bottom: 3px solid {'#ff2a55' if errors_corrected > 0 else '#00e0b0'};">
                            <div class="hud-header">
                                <div class="hud-title">Drift Anomalies</div>
                                <div class="hud-icon"><img src="{ICONS['DRIFT'] if errors_corrected > 0 else ICONS['VERIFIED']}"></div>
                            </div>
                            <div class="hud-val" style="color: {'#ff2a55' if errors_corrected > 0 else '#00e0b0'};">{errors_corrected}</div>
                            <div class="hud-sub">Mathematical deviations found</div>
                        </div>
                        <div class="hud-card" style="border-bottom: 3px solid #3b82f6;">
                            <div class="hud-header">
                                <div class="hud-title">Timeline Stitched</div>
                                <div class="hud-icon"><img src="{ICONS['CLOCK']}"></div>
                            </div>
                            <div class="hud-val">{total_days}</div>
                            <div class="hud-sub">Chronological Days Verified</div>
                        </div>
                        <div class="hud-card" style="border-bottom: 3px solid #7b68ee;">
                            <div class="hud-header">
                                <div class="hud-title">Digital Seal (SHA-256)</div>
                                <div class="hud-icon"><img src="{ICONS['SEAL']}"></div>
                            </div>
                            <div class="hud-val" style="font-size: 1.4rem; margin-top: 8px;">{digital_seal[:12]}</div>
                            <div class="hud-sub">Data integrity locked</div>
                        </div>
                    </div>
                    """, unsafe_allow_html=True)

                    def style_dataframe(row):
                        if row['Status'] == 'DRIFT DETECTED':
                            return ['background-color: rgba(255, 42, 85, 0.1); color: #ff8a9f'] * len(row)
                        elif row['Status'] == 'NO TIMELINE DATA':
                            return ['background-color: rgba(201, 168, 76, 0.1); color: #c9a84c'] * len(row)
                        return ['color: #00e0b0'] * len(row)

                    st.dataframe(res_df.style.apply(style_dataframe, axis=1), use_container_width=True, hide_index=True)

                    csv_data = res_df.to_csv(index=False).encode('utf-8')
                    st.download_button("⬇️ EXPORT FORENSIC LEDGER (.CSV)", data=csv_data, file_name="Reconciliation_Audit.csv", mime='text/csv')
                else:
                    st.error("Audit Could Not Complete. Check the Diagnostics tab to ensure the files contain the correct data.")

            with t2:
                if audit_results:
                    st.markdown("### Claimed vs Verified Running Hours")
                    st.markdown("<span style='color:#64748b; font-size:0.85rem;'>Native Streamlit rendering (Zero external chart dependencies)</span><br><br>", unsafe_allow_html=True)
                    res_df = pd.DataFrame(audit_results)
                    plot_df = res_df[['Component', 'Claimed (Doc)', 'Verified (Excel)']].set_index('Component')
                    st.bar_chart(plot_df, color=["#c9a84c", "#00e0b0"], height=500)

            with t3:
                st.markdown("### The Transparency Engine")
                c1, c2 = st.columns(2)
                with c1:
                    st.markdown("<div style='color:#8ba1b5; font-size:0.8rem; font-weight:600; margin-bottom:10px;'>RAW ASCII CELLS (.doc/.docx)</div>", unsafe_allow_html=True)
                    if diag_pms is not None and not diag_pms.empty:
                        st.dataframe(diag_pms, use_container_width=True, height=500)
                    else:
                        st.warning("No data extracted from the Overhaul file.")

                with c2:
                    st.markdown("<div style='color:#8ba1b5; font-size:0.8rem; font-weight:600; margin-bottom:10px;'>X-RAY TIMELINE MATRIX (.xlsx)</div>", unsafe_allow_html=True)
                    if diag_timeline is not None and not diag_timeline.empty:
                        st.code(f"Engine Column Map Generated:\n{col_map}", language="json")
                        st.write(f"Parsed timeline rows: {len(timeline_df)}")
                        st.dataframe(timeline_df.head(50), use_container_width=True, height=450)
                    else:
                        st.warning("No data extracted from the Timeline file.")

        except Exception as e:
            st.error(f"🚨 Pipeline Execution Halted: {str(e)}")
            st.info(traceback.format_exc())
