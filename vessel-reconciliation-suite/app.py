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
from difflib import SequenceMatcher

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
    "ORACLE": _u('<svg viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg"><path fill="#f59e0b" d="M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm-1 15h2v2h-2v-2zm1-12C9.24 5 7 7.24 7 10h2c0-1.66 1.34-3 3-3s3 1.34 3 3c0 2-3 1.75-3 5h2c0-2.25 3-2.5 3-5 0-2.76-2.24-5-5-5z"/></svg>'),
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
    
    .truth-index-box { background: linear-gradient(135deg, rgba(245, 158, 11, 0.1), rgba(245, 158, 11, 0.02)); border: 1px solid rgba(245, 158, 11, 0.3); border-radius: 12px; padding: 30px; text-align: center; margin-bottom: 20px; }
    .truth-index-val { font-size: 4rem; font-weight: 800; color: #f59e0b; line-height: 1; font-family: 'JetBrains Mono', monospace; }
    .truth-index-lbl { font-size: 0.85rem; color: #cbd5e1; text-transform: uppercase; letter-spacing: 2px; font-weight: 600; margin-top: 10px; }
    </style>
""", unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════════════════════
# 2. CORE DATA UTILITIES
# ═══════════════════════════════════════════════════════════════════════════════

def normalize_text(s):
    s = str(s).upper().strip()
    s = s.replace("\xa0", " ")
    return re.sub(r"\s+", " ", s)

def extract_first_float(value):
    if pd.isna(value): return np.nan
    s = str(value).replace(",", "")
    m = re.search(r"-?\d+(?:\.\d+)?", s)
    return float(m.group()) if m else np.nan

def robust_parse_date(value):
    if pd.isna(value): return pd.NaT
    s = str(value).strip()
    if not s: return pd.NaT
    direct = pd.to_datetime(s, errors="coerce", dayfirst=True)
    if pd.notna(direct): return direct
    cleaned = re.sub(r"\s+", " ", s)
    for fmt in ["%Y-%m-%d %H%M%S", "%Y-%m-%d", "%d-%m-%Y", "%d/%m/%Y", "%d.%m.%Y", "%d-%b-%Y", "%d-%B-%Y"]:
        try:
            return pd.to_datetime(cleaned, format=fmt, errors="raise", dayfirst=True)
        except: pass
    return pd.NaT

# ═══════════════════════════════════════════════════════════════════════════════
# 3. EXCEL MASTER-LOG PARSER & WORD DOCUMENT PARSER (IMMUTABLE)
# ═══════════════════════════════════════════════════════════════════════════════

@st.cache_data(show_spinner=False)
def parse_master_pms_excel(file_bytes):
    try:
        xls = pd.ExcelFile(io.BytesIO(file_bytes), engine="openpyxl")
        target_sheet = xls.sheet_names[0]
        for s in xls.sheet_names:
            if "PARTS LOG" in s.upper() or "PMS" in s.upper():
                target_sheet = s
                break
        df_raw = pd.read_excel(io.BytesIO(file_bytes), sheet_name=target_sheet, header=None, dtype=object, engine="openpyxl")
    except Exception:
        try:
            df_raw = pd.read_excel(io.BytesIO(file_bytes), header=None, dtype=object, engine="openpyxl")
        except: return []

    item_col, hours_col, header_row = -1, -1, -1
    for i in range(min(30, len(df_raw))):
        row_joined = " | ".join([str(x).upper() for x in df_raw.iloc[i].values if pd.notna(x)])
        if ("ITEM" in row_joined or "DESCRIPTION" in row_joined) and "CURRENT" in row_joined and "HOURS" in row_joined:
            header_row = i
            for j, val in enumerate(df_raw.iloc[i].values):
                val_u = str(val).upper()
                if "ITEM" in val_u or "DESCRIPTION" in val_u: item_col = j
                elif "CURRENT" in val_u and "HOURS" in val_u: hours_col = j
            break

    if item_col == -1 or hours_col == -1: return []

    excel_records = []
    curr_sys, curr_unit = "ME", ""

    for i in range(header_row + 1, len(df_raw)):
        item_val, hours_val = df_raw.iloc[i, item_col], df_raw.iloc[i, hours_col]
        if pd.isna(item_val): continue
        item_str = normalize_text(item_val)

        if "MAIN ENGINE" in item_str or "M/E" in item_str: curr_sys, curr_unit = "ME", ""
        elif any(x in item_str for x in ["GENERATOR NO.1", "DG 1", "D/G 1", "NO 1", "NO.1"]): curr_sys, curr_unit = "DG", "DG1"
        elif any(x in item_str for x in ["GENERATOR NO.2", "DG 2", "D/G 2", "NO 2", "NO.2"]): curr_sys, curr_unit = "DG", "DG2"
        elif any(x in item_str for x in ["GENERATOR NO.3", "DG 3", "D/G 3", "NO 3", "NO.3"]): curr_sys, curr_unit = "DG", "DG3"

        if "CYLINDER NO" in item_str or "CYL NO" in item_str:
            m = re.search(r'NO[\.\:\s]*(\d+)', item_str)
            if m: curr_unit = f"Cyl {m.group(1)}"

        h = extract_first_float(hours_val)
        if pd.notna(h) and len(item_str) > 2 and "HOURS" not in item_str and "DATE" not in item_str:
            excel_records.append({'System': curr_sys, 'Unit': curr_unit, 'ExcelComponent': item_str, 'ExcelHours': h})

    return excel_records

@st.cache_data(show_spinner=False)
def parse_pms_binary_doc(file_bytes, file_name=""):
    fname = (file_name or "").lower()
    def parse_docx_xml(bytes_):
        texts = []
        try:
            with ZipFile(io.BytesIO(bytes_)) as z: root = ET.fromstring(z.read("word/document.xml"))
            ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
            for para in root.findall(".//w:p", ns):
                line = " ".join([t.text for t in para.findall(".//w:t", ns) if t.text]).strip()
                if line: texts.append(line)
        except: pass
        return texts

    cells = parse_docx_xml(file_bytes) if fname.endswith(".docx") else [c.strip() for c in file_bytes.decode("latin-1", errors="ignore").replace("\x00", "").split("\x07") if c.strip()]
    extracted_data = []

    COMPONENTS = ['CYLINDER COVER', 'PISTON ASSY', 'PISTON ASSEMBLY', 'STUFFING BOX', 'PISTON CROWN', 'CYLINDER LINER', 'EXHAUST VALVE', 'EXAUST VALVE', 'STARTING VALVE', 'SAFETY VALVE', 'FUEL VALVE', 'FUEL VALVES', 'FUEL PUMP', 'SUCTION VALVE', 'PUNCTURE VALVE', 'CROSSHEAD BEARING', 'CROSSHEAD BEARINGS', 'BOTTOM END BEARING', 'BOTTOM END BEARINGS', 'MAIN BEARING', 'MAIN BEARINGS', 'CYLINDER HEAD', 'PISTON', 'CONNECTING ROD', 'TURBOCHARGER', 'AIR COOLER', 'COOLING WATER PUMP', 'THERMOSTAT VALVE', 'THRUST BEARING']
    system, sub_units = "ME", ["Cyl 1", "Cyl 2", "Cyl 3", "Cyl 4", "Cyl 5", "Cyl 6"]

    i = 0
    while i < len(cells):
        cell = normalize_text(cells[i])
        if "MAIN ENGINE" in cell: system, sub_units = "ME", ["Cyl 1", "Cyl 2", "Cyl 3", "Cyl 4", "Cyl 5", "Cyl 6"]
        elif any(x in cell for x in ["AUX. ENGINE", "AUX ENGINE", "D/G", "DG ", "DIESEL GENERATOR"]): system, sub_units = "DG", ["DG1", "DG2", "DG3"]

        if any(c in cell for c in COMPONENTS) and len(cell) < 80:
            dates, hours = [], []
            for shift, target, arr in [(15, '1', dates), (35, '2', hours)]:
                j = i + 1
                while j < min(i + shift, len(cells)):
                    if str(cells[j]).strip() == target:
                        for k in range(len(sub_units)):
                            if j + 1 + k < len(cells): arr.append(cells[j + 1 + k])
                        break
                    j += 1

            for idx, su in enumerate(sub_units):
                if idx < len(dates) and idx < len(hours):
                    d, h = robust_parse_date(dates[idx]), extract_first_float(hours[idx])
                    if pd.notna(d) and pd.notna(h):
                        extracted_data.append({'System': system, 'Unit': su, 'BaseComponent': cell, 'Component': f"[{system}] {cell} ({su})", 'Last_Overhaul': d, 'Claimed_Hours': float(h)})
        i += 1
    return pd.DataFrame(extracted_data).dropna(subset=['Last_Overhaul']).reset_index(drop=True), pd.DataFrame(cells, columns=["Raw Doc Cells"])

def get_verified_hours(doc_row, excel_records):
    doc_sys, doc_unit = doc_row['System'], doc_row['Unit']
    w_comp = re.sub(r'[^A-Z]', '', doc_row['BaseComponent'].replace("ASSY", "ASSEMBLY"))
    best_score, best_hours = 0, 0
    for er in excel_records:
        if er['System'] != doc_sys or (er['Unit'] != doc_unit and er['Unit'] != ""): continue
        e_comp = re.sub(r'[^A-Z]', '', er['ExcelComponent'].replace("ASSY", "ASSEMBLY"))
        score = SequenceMatcher(None, w_comp, e_comp).ratio()
        if w_comp in e_comp or e_comp in w_comp: score += 0.3
        if score > best_score: best_score, best_hours = score, er['ExcelHours']
    return best_hours if best_score > 0.55 else None

# ═══════════════════════════════════════════════════════════════════════════════
# 4. THE ORACLE (MIDDLEWARE ANALYTICS ENGINE)
# ═══════════════════════════════════════════════════════════════════════════════

def run_oracle_diagnostics(audit_results, excel_records, doc_df):
    """Applies Physics Laws, MAD Z-Scores, and Ghost Sweeps to generate the Truth Index."""
    
    # 1. Calculate Robust Statistics (MAD Z-Score)
    valid_deltas = [r['Delta (Drift)'] for r in audit_results if r['Status'] != "MISSING FROM EXCEL"]
    if valid_deltas:
        median_delta = np.median(valid_deltas)
        mad = np.median([abs(d - median_delta) for d in valid_deltas])
        if mad == 0: mad = 1e-6
    else:
        median_delta, mad = 0, 1e-6

    threat_matrix = []
    physics_violations = 0
    severe_outliers = 0
    zero_drift_count = 0
    now = pd.Timestamp.now()

    for r in audit_results:
        if r['Status'] == "MISSING FROM EXCEL":
            r['Severity'], r['Z-Score'], r['Anomaly'] = 0, 0.0, "MISSING MASTER RECORD"
            threat_matrix.append(r)
            continue
            
        delta = r['Delta (Drift)']
        claimed = r['Claimed (Doc)']
        if delta == 0: zero_drift_count += 1
        
        anomaly = "NONE"
        oh_date_str = r['Overhaul Date']
        time_violation = False
        
        # A. Physics Check: Time Warp (Claimed hours exceed calendar hours)
        if oh_date_str:
            oh_date = pd.to_datetime(oh_date_str, errors='coerce')
            if pd.notna(oh_date):
                days_elapsed = max(1, (now - oh_date).days)
                max_hours = days_elapsed * 24
                if claimed > max_hours:
                    anomaly = f"TIME VIOLATION (Max {max_hours}h)"
                    physics_violations += 1
                    time_violation = True

        # B. Physics Check: Claim Exceeds Master Truth
        if delta < 0 and not time_violation:
            anomaly = f"CLAIM EXCEEDS MASTER BY {abs(delta)}h"
            physics_violations += 1
            
        # C. Statistical Integrity (MAD Z-Score)
        z_score = 0.6745 * (delta - median_delta) / mad
        abs_z = abs(z_score)
        
        if abs_z < 1: sev = 1
        elif abs_z < 2: sev = 2
        elif abs_z < 3.5: sev = 3
        elif abs_z < 5: sev = 4
        else: sev = 5
        
        if sev >= 4 and anomaly == "NONE":
            anomaly = f"SEVERE OUTLIER (Z={z_score:.1f})"
            severe_outliers += 1
            
        r['Severity'] = sev
        r['Z-Score'] = round(z_score, 2)
        r['Anomaly'] = anomaly
        
        if anomaly != "NONE":
            threat_matrix.append(r)
            
    # 2. Ghost Sweep Engine (Reverse Look-up)
    ghosts = []
    doc_comps = [re.sub(r'[^A-Z]', '', normalize_text(c).replace("ASSY", "ASSEMBLY")) for c in doc_df['BaseComponent'].tolist()] if not doc_df.empty else []
    major_kws = ['LINER', 'PISTON', 'BEARING', 'PUMP', 'VALVE', 'COOLER', 'TURBO', 'COVER', 'HEAD']
    
    for er in excel_records:
        e_comp_clean = re.sub(r'[^A-Z]', '', er['ExcelComponent'].replace("ASSY", "ASSEMBLY"))
        if any(k in er['ExcelComponent'].upper() for k in major_kws):
            found = False
            for w_comp_clean in doc_comps:
                if SequenceMatcher(None, e_comp_clean, w_comp_clean).ratio() > 0.55 or w_comp_clean in e_comp_clean or e_comp_clean in w_comp_clean:
                    found = True
                    break
            if not found:
                ghosts.append({
                    "System": er['System'],
                    "Unit": er['Unit'],
                    "Component": er['ExcelComponent'],
                    "Master Hours": er['ExcelHours'],
                    "Anomaly": "GHOST (OMITTED FROM REPORT)"
                })

    # 3. POSEIDON Truth Index
    total_audited = len(audit_results)
    if total_audited == 0:
        truth_index = 0
    else:
        score_acc = 0.40 * ((zero_drift_count / total_audited) * 100)
        score_phys = 0.30 * max(0, 100 - (physics_violations * 20))
        score_ghost = 0.20 * max(0, 100 - (len(ghosts) * 5))
        score_stat = 0.10 * max(0, 100 - (severe_outliers * 10))
        truth_index = int(score_acc + score_phys + score_ghost + score_stat)
        
    return audit_results, threat_matrix, ghosts, truth_index

# ═══════════════════════════════════════════════════════════════════════════════
# 5. MAIN FRONTEND ORCHESTRATOR
# ═══════════════════════════════════════════════════════════════════════════════

st.markdown(f"""
<div class="hero">
    <div class="hero-left">
        <img src="data:image/svg+xml;base64,{LOGO_SVG}" class="hero-logo" alt=""/>
        <div>
            <div class="hero-title">POSEIDON RECON</div>
            <div class="hero-sub">Predictive Forensics Engine</div>
        </div>
    </div>
    <div class="hero-badge">
        <span style="color:#00e0b0">KERNEL</span>&ensp;Fuzzy Vlookup Bridge<br>
        <span style="color:#f59e0b">ORACLE</span>&ensp;Physics & MAD Statistics<br>
        <span style="color:#fff">BUILD</span>&ensp;v11.0.0 The Oracle
    </div>
</div>
""", unsafe_allow_html=True)

col1, col2 = st.columns(2)
with col1:
    st.markdown("<div style='color:#8ba1b5; font-size:0.9rem; font-weight:600; margin-bottom:10px;'>1. COMPONENT OVERHAULS (.doc)</div>", unsafe_allow_html=True)
    pms_file = st.file_uploader("Upload Word Report", type=["doc", "docx"], key="pms_box")

with col2:
    st.markdown("<div style='color:#8ba1b5; font-size:0.9rem; font-weight:600; margin-bottom:10px;'>2. EXCEL MASTER LOG (PMS)</div>", unsafe_allow_html=True)
    logs_file = st.file_uploader("Upload Excel PMS Log", type=["xlsx", "xls"], key="logs_box")

if pms_file and logs_file:
    with st.spinner("Executing Mathematical Cross-Reference & Guardrails..."):
        try:
            doc_df, diag_pms = parse_pms_binary_doc(pms_file.getvalue(), pms_file.name)
            excel_records = parse_master_pms_excel(logs_file.getvalue())

            raw_audit_results = []
            if not doc_df.empty and excel_records:
                for _, row in doc_df.iterrows():
                    comp, oh_date, legacy_hrs = row['Component'], row['Last_Overhaul'], row['Claimed_Hours']
                    verified_hrs = get_verified_hours(row, excel_records)

                    if verified_hrs is not None:
                        delta = verified_hrs - legacy_hrs
                        raw_audit_results.append({
                            "Component": comp,
                            "Overhaul Date": oh_date.strftime('%d-%b-%Y') if pd.notna(oh_date) else "",
                            "Claimed (Doc)": int(round(float(legacy_hrs))),
                            "Verified (Excel)": int(round(float(verified_hrs))),
                            "Delta (Drift)": int(round(float(delta))),
                            "Status": "VERIFIED" if int(round(float(delta))) == 0 else "DRIFT DETECTED"
                        })
                    else:
                        raw_audit_results.append({
                            "Component": comp,
                            "Overhaul Date": oh_date.strftime('%d-%b-%Y') if pd.notna(oh_date) else "",
                            "Claimed (Doc)": int(round(float(legacy_hrs))),
                            "Verified (Excel)": 0,
                            "Delta (Drift)": 0,
                            "Status": "MISSING FROM EXCEL"
                        })

            # THE ORACLE (Middleware Execution)
            audit_results, threat_matrix, ghosts, truth_index = run_oracle_diagnostics(raw_audit_results, excel_records, doc_df)

            st.markdown("<br>", unsafe_allow_html=True)
            t1, t2, t3 = st.tabs(["📊 IMMUTABLE LEDGER", "👁️ THE ORACLE (ANALYTICS)", "🔎 MASTER DATABASE X-RAY"])

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
                            <div class="hud-sub">Cross-Referenced via Vlookup</div>
                        </div>
                        <div class="hud-card" style="border-bottom: 3px solid {'#ff2a55' if errors_corrected > 0 else '#00e0b0'};">
                            <div class="hud-header">
                                <div class="hud-title">Drift Anomalies</div>
                                <div class="hud-icon"><img src="{ICONS['DRIFT'] if errors_corrected > 0 else ICONS['VERIFIED']}"></div>
                            </div>
                            <div class="hud-val" style="color: {'#ff2a55' if errors_corrected > 0 else '#00e0b0'};">{errors_corrected}</div>
                            <div class="hud-sub">Mathematical deviations found</div>
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
                        if row['Status'] == 'DRIFT DETECTED': return ['background-color: rgba(255, 42, 85, 0.1); color: #ff8a9f'] * len(row)
                        elif row['Status'] == 'MISSING FROM EXCEL': return ['background-color: rgba(201, 168, 76, 0.1); color: #c9a84c'] * len(row)
                        return ['color: #00e0b0'] * len(row)
                    
                    st.dataframe(res_df.drop(columns=['Severity', 'Z-Score', 'Anomaly']).style.apply(style_dataframe, axis=1), use_container_width=True, hide_index=True)
                    csv_data = res_df.to_csv(index=False).encode('utf-8')
                    st.download_button("⬇️ EXPORT FORENSIC LEDGER (.CSV)", data=csv_data, file_name="Reconciliation_Audit.csv", mime='text/csv')
                else:
                    st.error("Audit Could Not Complete. Check the Diagnostics tab to ensure the Excel file contains the Master Parts Log.")

            with t2:
                if audit_results:
                    st.markdown(f"""
                    <div class="truth-index-box">
                        <div class="truth-index-val">{truth_index}%</div>
                        <div class="truth-index-lbl">POSEIDON Truth Index</div>
                        <div style="font-size:0.75rem; color:#64748b; margin-top:5px;">Mathematical grade based on Zero-Drift, Physics Violations, Statistical Outliers, and Ghost Sweeps.</div>
                    </div>
                    """, unsafe_allow_html=True)

                    c1, c2 = st.columns(2)
                    with c1:
                        st.markdown("<h4 style='color:#f59e0b;'>⚠️ Critical Threat Matrix</h4>", unsafe_allow_html=True)
                        st.markdown("<span style='color:#64748b; font-size:0.85rem;'>Displays severe statistical outliers and physical time violations. Routine typos are hidden.</span><br><br>", unsafe_allow_html=True)
                        if threat_matrix:
                            tm_df = pd.DataFrame(threat_matrix)[['Component', 'Delta (Drift)', 'Z-Score', 'Anomaly']]
                            st.dataframe(tm_df, use_container_width=True, hide_index=True)
                        else:
                            st.success("Zero severe mathematical anomalies detected.")

                    with c2:
                        st.markdown("<h4 style='color:#00e0b0;'>👻 Ghost Component Sweep</h4>", unsafe_allow_html=True)
                        st.markdown("<span style='color:#64748b; font-size:0.85rem;'>Major equipment found in the Excel Master Database but completely omitted from the Word report.</span><br><br>", unsafe_allow_html=True)
                        if ghosts:
                            st.dataframe(pd.DataFrame(ghosts), use_container_width=True, hide_index=True)
                        else:
                            st.success("No major Ghost Components detected. Word report is structurally complete.")

            with t3:
                c1, c2 = st.columns(2)
                with c1:
                    st.markdown("<div style='color:#8ba1b5; font-size:0.8rem; font-weight:600; margin-bottom:10px;'>DOC REPORT (Extracted Components)</div>", unsafe_allow_html=True)
                    if not doc_df.empty:
                        st.dataframe(doc_df[['Component', 'Claimed_Hours', 'Last_Overhaul']], use_container_width=True, height=500)
                    else:
                        st.warning("No data extracted from the Word file.")
                        
                with c2:
                    st.markdown("<div style='color:#8ba1b5; font-size:0.8rem; font-weight:600; margin-bottom:10px;'>EXCEL MASTER (Extracted Databases)</div>", unsafe_allow_html=True)
                    if excel_records:
                        st.dataframe(pd.DataFrame(excel_records), use_container_width=True, height=500)
                    else:
                        st.warning("No Master Data extracted from the Excel file. Is it the right sheet?")

        except Exception as e:
            st.error(f"🚨 Pipeline Execution Halted: {str(e)}")
            st.info(traceback.format_exc())
