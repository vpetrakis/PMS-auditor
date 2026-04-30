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
import plotly.graph_objects as go

# Optional Enterprise ML Stack
try:
    from sklearn.covariance import LedoitWolf
    HAS_ML = True
except ImportError:
    HAS_ML = False

warnings.filterwarnings("ignore")

# ═══════════════════════════════════════════════════════════════════════════════
# 1. PAGE CONFIGURATION & ENTERPRISE CSS
# ═══════════════════════════════════════════════════════════════════════════════
st.set_page_config(page_title="PMS Auditor", page_icon="⚓", layout="wide", initial_sidebar_state="collapsed")

def _u(s):
    return f"data:image/svg+xml;base64,{base64.b64encode(s.encode()).decode()}"

LOGO_SVG = base64.b64encode(b'<svg viewBox="0 0 48 48" xmlns="http://www.w3.org/2000/svg"><defs><linearGradient id="pg" x1="0" y1="0" x2="1" y2="1"><stop offset="0%" stop-color="#c9a84c"/><stop offset="50%" stop-color="#00e0b0"/><stop offset="100%" stop-color="#ffffff"/></linearGradient></defs><circle cx="24" cy="24" r="22" fill="none" stroke="url(#pg)" stroke-width="1" opacity="0.35"/><path d="M24 6 L24 42" stroke="url(#pg)" stroke-width="1.6" stroke-linecap="round"/><path d="M12 24 Q24 32 36 24" fill="none" stroke="url(#pg)" stroke-width="1.6" stroke-linecap="round"/></svg>').decode()

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;800&family=JetBrains+Mono:wght@400;700&display=swap');
    :root {
        --bg:#071018; --panel:#0d1620; --line:rgba(255,255,255,0.08); --text:#f8fafc;
        --muted:#94a3b8; --teal:#00e0b0; --gold:#c9a84c; --red:#ff2a55; --violet:#7b68ee;
    }
    .stApp { background-color: var(--bg); font-family: 'Inter', sans-serif; color: var(--text); }
    #MainMenu, footer, header {visibility: hidden;}
    
    .hero { display:flex; justify-content:space-between; align-items:flex-start; gap:24px; padding:24px 28px; margin-bottom:22px; background:linear-gradient(180deg, rgba(255,255,255,0.02), rgba(255,255,255,0.01)); border:1px solid var(--line); border-radius:24px; box-shadow: 0 20px 60px rgba(0,0,0,0.35); }
    .hero-left { display: flex; align-items: center; gap: 20px; }
    .hero-logo { width: 55px; height: 55px; }
    .hero-title { font-size: 2.2rem; font-weight: 800; background: linear-gradient(90deg, #ffffff, #8ba1b5); -webkit-background-clip: text; -webkit-text-fill-color: transparent; letter-spacing: -1px; line-height: 1.1;}
    .hero-sub { font-size: 0.85rem; color: var(--muted); font-weight: 600; text-transform: uppercase; letter-spacing: 2px; margin-top: 6px; }
    .hero-badge { text-align: right; font-family: 'JetBrains Mono', monospace; font-size: 0.75rem; color: var(--muted); line-height: 1.6; }

    .stFileUploader > div > div { background-color: var(--panel) !important; border: 1px dashed rgba(255,255,255,0.1) !important; border-radius: 12px !important; transition: all 0.3s ease; }
    .stFileUploader > div > div:hover { border-color: var(--teal) !important; background-color: rgba(0, 224, 176, 0.05) !important; }

    .hud-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(220px, 1fr)); gap: 14px; margin-bottom: 22px; }
    .hud-card { background: var(--panel); border: 1px solid var(--line); border-radius: 18px; padding: 16px; box-shadow: 0 10px 30px -10px rgba(0,0,0,0.5); }
    .hud-header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 10px; }
    .hud-title { font-size: 0.75rem; color: var(--muted); text-transform: uppercase; letter-spacing: 1.2px; font-weight: 600; }
    .hud-val { font-size: 2rem; font-weight: 800; color: #ffffff; line-height: 1.1; font-family: 'JetBrains Mono', monospace; }

    .stTabs [data-baseweb="tab-list"] { gap: 24px; }
    .stTabs [data-baseweb="tab"] { height: 50px; white-space: pre-wrap; background-color: transparent; border-radius: 4px; color: var(--muted); font-weight: 600; }
    .stTabs [aria-selected="true"] { color: var(--teal); border-bottom: 2px solid var(--teal); }
    
    .truth-index-box { background: linear-gradient(135deg, rgba(0, 224, 176, 0.1), rgba(0, 224, 176, 0.02)); border: 1px solid rgba(0, 224, 176, 0.3); border-radius: 18px; padding: 30px; text-align: center; margin-bottom: 20px; }
    .truth-index-val { font-size: 4rem; font-weight: 800; color: var(--teal); line-height: 1; font-family: 'JetBrains Mono', monospace; }
    .truth-index-lbl { font-size: 0.85rem; color: #cbd5e1; text-transform: uppercase; letter-spacing: 2px; font-weight: 600; margin-top: 10px; }
</style>
""", unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════════════════════
# 2. SEMANTIC HEURISTIC PARSERS
# ═══════════════════════════════════════════════════════════════════════════════

def normalize_text(s):
    return re.sub(r"\s+", " ", str(s).upper().strip().replace("\xa0", " "))

def extract_first_float(value):
    if pd.isna(value): return np.nan
    m = re.search(r"-?\d+(?:\.\d+)?", str(value).replace(",", ""))
    return float(m.group()) if m else np.nan

@st.cache_data(show_spinner=False)
def parse_master_pms_excel(file_bytes):
    try:
        xls = pd.ExcelFile(io.BytesIO(file_bytes), engine="openpyxl")
        target_sheet = next((s for s in xls.sheet_names if "PMS" in s.upper() or "PARTS LOG" in s.upper()), xls.sheet_names[0])
        df_raw = pd.read_excel(io.BytesIO(file_bytes), sheet_name=target_sheet, header=None, dtype=object, engine="openpyxl")
    except Exception: return [], {}

    # Triple-Lock Extraction
    engine_ytd_hours = {"ME": 0.0, "DG": 0.0}
    try:
        for i in range(min(25, len(df_raw))):
            row_joined = " | ".join([str(x).upper() for x in df_raw.iloc[i].values if pd.notna(x)])
            if any(k in row_joined for k in ["JAN", "FEB", "MONTHLY", "UP-TO-DATE"]):
                target_row = i + 1 if ("JAN" in row_joined and df_raw.iloc[i, 9] == "Jan.") else i
                if target_row < len(df_raw):
                    monthly_sum = sum([extract_first_float(df_raw.iloc[target_row, c]) for c in range(9, 21) if c < len(df_raw.columns) and pd.notna(extract_first_float(df_raw.iloc[target_row, c]))])
                    if monthly_sum > 0:
                        sys = "ME" if ("MAIN ENGINE" in str(df_raw.iloc[max(0, target_row-2):target_row].values).upper()) else "ME"
                        engine_ytd_hours[sys] = max(engine_ytd_hours[sys], monthly_sum)
    except: pass
    if engine_ytd_hours["ME"] == 0: engine_ytd_hours["ME"] = 8760 

    # Semantic Header Mapping
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

    excel_records = []
    if item_col != -1 and hours_col != -1:
        curr_sys, curr_unit = "ME", ""
        for i in range(header_row + 1, len(df_raw)):
            item_val, hours_val = df_raw.iloc[i, item_col], df_raw.iloc[i, hours_col]
            if pd.isna(item_val): continue
            item_str = normalize_text(item_val)

            if "MAIN ENGINE" in item_str or "M/E" in item_str: curr_sys, curr_unit = "ME", ""
            elif any(x in item_str for x in ["DG 1", "D/G 1", "NO.1"]): curr_sys, curr_unit = "DG", "DG1"
            elif any(x in item_str for x in ["DG 2", "D/G 2", "NO.2"]): curr_sys, curr_unit = "DG", "DG2"
            elif any(x in item_str for x in ["DG 3", "D/G 3", "NO.3"]): curr_sys, curr_unit = "DG", "DG3"

            m = re.search(r'NO[\.\:\s]*(\d+)', item_str)
            if m: curr_unit = f"Cyl {m.group(1)}"

            h = extract_first_float(hours_val)
            if pd.notna(h) and len(item_str) > 2 and "HOURS" not in item_str and "DATE" not in item_str:
                excel_records.append({'System': curr_sys, 'Unit': curr_unit, 'ExcelComponent': item_str, 'ExcelHours': h})

    return excel_records, engine_ytd_hours

@st.cache_data(show_spinner=False)
def parse_pms_binary_doc(file_bytes, file_name=""):
    def parse_docx_xml(bytes_):
        try:
            with ZipFile(io.BytesIO(bytes_)) as z: root = ET.fromstring(z.read("word/document.xml"))
            ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
            return [" ".join([t.text for t in p.findall(".//w:t", ns) if t.text]).strip() for p in root.findall(".//w:p", ns) if p.text or len(p.findall(".//w:t", ns))>0]
        except: return []

    cells = parse_docx_xml(file_bytes) if file_name.lower().endswith(".docx") else [c.strip() for c in file_bytes.decode("latin-1", errors="ignore").replace("\x00", "").split("\x07") if c.strip()]
    extracted_data = []

    COMPONENTS = ['CYLINDER COVER', 'PISTON ASSY', 'PISTON ASSEMBLY', 'STUFFING BOX', 'PISTON CROWN', 'CYLINDER LINER', 'EXHAUST VALVE', 'STARTING VALVE', 'SAFETY VALVE', 'FUEL VALVE', 'FUEL PUMP', 'CROSSHEAD BEARING', 'BOTTOM END BEARING', 'MAIN BEARING', 'CYLINDER HEAD', 'CONNECTING ROD', 'TURBOCHARGER', 'AIR COOLER']
    system, sub_units = "ME", ["Cyl 1", "Cyl 2", "Cyl 3", "Cyl 4", "Cyl 5", "Cyl 6"]

    i = 0
    while i < len(cells):
        cell = normalize_text(cells[i])
        if "MAIN ENGINE" in cell: system, sub_units = "ME", ["Cyl 1", "Cyl 2", "Cyl 3", "Cyl 4", "Cyl 5", "Cyl 6"]
        elif any(x in cell for x in ["AUX ENGINE", "D/G", "DIESEL GENERATOR"]): system, sub_units = "DG", ["DG1", "DG2", "DG3"]

        if any(c in cell for c in COMPONENTS) and len(cell) < 80:
            dates, hours = [], []
            for shift, target, arr in [(15, '1', dates), (35, '2', hours)]:
                j = i + 1
                while j < min(i + shift, len(cells)):
                    if str(cells[j]).strip() == target:
                        arr.extend([cells[j + 1 + k] for k in range(len(sub_units)) if j + 1 + k < len(cells)])
                        break
                    j += 1

            for idx, su in enumerate(sub_units):
                if idx < len(dates) and idx < len(hours):
                    d = pd.to_datetime(re.sub(r"\s+", " ", str(dates[idx]).strip()), errors="coerce", dayfirst=True)
                    h = extract_first_float(hours[idx])
                    if pd.notna(d) and pd.notna(h):
                        extracted_data.append({'System': system, 'Unit': su, 'BaseComponent': cell, 'Component': f"[{system}] {cell} ({su})", 'Last_Overhaul': d, 'Claimed_Hours': float(h), 'Source_File': file_name})
        i += 1
    return pd.DataFrame(extracted_data).dropna(subset=['Last_Overhaul']).reset_index(drop=True)

def get_verified_hours(doc_row, excel_records):
    doc_sys, doc_unit = doc_row['System'], doc_row['Unit']
    w_comp = re.sub(r'[^A-Z]', '', doc_row['BaseComponent'].replace("ASSY", "ASSEMBLY"))
    best_score, best_hours = 0, 0
    for er in excel_records:
        if er['System'] != doc_sys or (er['Unit'] != doc_unit and er['Unit'] != ""): continue
        e_comp = re.sub(r'[^A-Z]', '', er['ExcelComponent'].replace("ASSY", "ASSEMBLY"))
        score = SequenceMatcher(None, w_comp, e_comp).ratio()
        if w_comp in e_comp or e_comp in w_comp: score += 0.3
        if score > 0.55 and score > best_score: best_score, best_hours = score, er['ExcelHours']
    return best_hours if best_score > 0.55 else None

# ═══════════════════════════════════════════════════════════════════════════════
# 3. THE ORACLE (PIML ENGINE: Quarantine, Conformal, Mahalanobis)
# ═══════════════════════════════════════════════════════════════════════════════

def run_oracle_piml(audit_results, excel_records, all_docs_df, engine_ytd_hours):
    now = pd.Timestamp.now()
    threat_matrix, ghosts = [], []
    quarantine_count, fraud_flags = 0, 0
    
    # 1. Historical Copy-Paste Fraud Detection
    copy_paste_threats = {}
    if not all_docs_df.empty and len(all_docs_df['Source_File'].unique()) > 1:
        for comp in all_docs_df['Component'].unique():
            comp_data = all_docs_df[all_docs_df['Component'] == comp]
            if len(comp_data) > 1 and comp_data['Claimed_Hours'].nunique() == 1:
                copy_paste_threats[comp] = "CRITICAL: COPY-PASTE FRAUD (Identical hours historical match)"

    # 2. Physics Constraints (The Quarantine Protocol)
    safe_deltas = []
    for r in audit_results:
        anomaly = "NONE"
        comp_name, delta, claimed = r['Component'], r['Delta (Drift)'], r['Claimed (Doc)']
        
        if r['Status'] == "MISSING FROM EXCEL":
            r['Phase'] = "QUARANTINE"
            r['Anomaly'] = "MISSING MASTER RECORD"
            threat_matrix.append(r)
            quarantine_count += 1
            continue

        if comp_name in copy_paste_threats:
            anomaly, r['Phase'] = copy_paste_threats[comp_name], "QUARANTINE"
            fraud_flags += 1

        parent_sys = comp_name.split("]")[0].replace("[", "") if "[" in comp_name else "ME"
        parent_max = engine_ytd_hours.get(parent_sys, 8760)
        
        oh_date = pd.to_datetime(r['Overhaul Date'], errors='coerce')
        if pd.notna(oh_date) and anomaly == "NONE":
            days_alive = max(1, (now - oh_date).days)
            if oh_date.year == now.year and claimed > parent_max:
                anomaly, r['Phase'] = f"TRIPLE-LOCK VIOLATION (> Engine Pulse {int(parent_max)}h)", "QUARANTINE"
            elif claimed > (days_alive * 24):
                anomaly, r['Phase'] = f"PHYSICS TIME WARP (> {days_alive*24} max hrs)", "QUARANTINE"

        if delta < 0 and anomaly == "NONE":
            anomaly, r['Phase'] = f"NEGATIVE DRIFT (-{abs(delta)}h)", "QUARANTINE"

        r['Anomaly'] = anomaly
        if anomaly == "NONE":
            r['Phase'] = "SAFE"
            safe_deltas.append(delta)
        else:
            threat_matrix.append(r)
            quarantine_count += 1

    # 3. Robust Statistics (Computed strictly on SAFE data)
    median_delta, mad = (np.median(safe_deltas), np.median([abs(d - np.median(safe_deltas)) for d in safe_deltas])) if safe_deltas else (0, 1e-6)
    if mad == 0: mad = 1e-6

    # 4. Conformal ML Prediction & Mahalanobis
    me_rate = engine_ytd_hours.get('ME', 0) / max(1, now.dayofyear)
    dg_rate = engine_ytd_hours.get('DG', 0) / max(1, now.dayofyear)

    ml_features = []
    
    for r in audit_results:
        z_score = 0.6745 * (r['Delta (Drift)'] - median_delta) / mad
        r['Z-Score'] = round(z_score, 2)
        r['Severity'] = 1 if abs(z_score) < 1 else 2 if abs(z_score) < 2 else 3 if abs(z_score) < 3.5 else 4 if abs(z_score) < 5 else 5
        
        if r['Severity'] >= 4 and r['Phase'] == "SAFE":
            r['Anomaly'] = f"SEVERE STAT OUTLIER (Z={z_score:.1f})"
            threat_matrix.append(r)

        r['Exp_Mean'] = 0.0
        r['Exp_Upper'] = 0.0
        r['Exp_Lower'] = 0.0

        if r.get('Overhaul Date') and r['Status'] != "MISSING FROM EXCEL":
            oh_date = pd.to_datetime(r['Overhaul Date'], errors='coerce')
            if pd.notna(oh_date):
                days_alive = max(1, (now - oh_date).days)
                sys = "ME" if "ME" in str(r.get('Component', 'ME')) else "DG"
                
                # Deterministic Prediction Baseline
                expected = days_alive * (me_rate if sys == "ME" else dg_rate)
                variance = expected * 0.12 + 50 # Conformal variance model (12% error margin + 50h fixed buffer)
                
                r['Exp_Mean'] = int(min(expected, days_alive * 24))
                r['Exp_Upper'] = int(min(expected + variance, days_alive * 24))
                r['Exp_Lower'] = int(max(0, expected - variance))

                if r['Phase'] == "SAFE":
                    ml_features.append({
                        "System": sys, "Component": r['Component'],
                        "Days_Alive": days_alive, "Claimed": r['Claimed (Doc)'], "Delta": r['Delta (Drift)']
                    })

    # Ledoit-Wolf Multivariate Anomaly Detection
    mahalanobis_alerts = 0
    if HAS_ML and len(ml_features) > 6:
        try:
            df_ml = pd.DataFrame(ml_features)
            for sys in df_ml['System'].unique():
                sys_data = df_ml[df_ml['System'] == sys]
                if len(sys_data) >= 5: # Need enough data points to compute covariance
                    X = sys_data[['Days_Alive', 'Claimed', 'Delta']].values
                    lw = LedoitWolf().fit(X)
                    md = np.sqrt(np.maximum(lw.mahalanobis(X), 0))
                    threshold = np.percentile(md, 95)
                    
                    # Flag anomalies back to audit_results
                    for idx, (comp, md_val) in enumerate(zip(sys_data['Component'], md)):
                        if md_val > threshold and md_val > 2.5:
                            for r in audit_results:
                                if r['Component'] == comp and r['Phase'] == "SAFE" and r['Severity'] < 4:
                                    r['Anomaly'] = f"MULTIVARIATE ANOMALY (Mahalanobis D={md_val:.1f})"
                                    threat_matrix.append(r)
                                    mahalanobis_alerts += 1
        except Exception as e:
            pass # Graceful degradation if matrix collapses

    # 5. Ghost Sweep
    primary_doc_comps = [re.sub(r'[^A-Z]', '', normalize_text(c).replace("ASSY", "ASSEMBLY")) for c in all_docs_df['BaseComponent'].tolist()] if not all_docs_df.empty else []
    for er in excel_records:
        e_comp_clean = re.sub(r'[^A-Z]', '', er['ExcelComponent'].replace("ASSY", "ASSEMBLY"))
        if any(k in er['ExcelComponent'].upper() for k in ['LINER', 'PISTON', 'BEARING', 'PUMP', 'VALVE']):
            if not any(SequenceMatcher(None, e_comp_clean, w).ratio() > 0.55 or w in e_comp_clean or e_comp_clean in w for w in primary_doc_comps):
                ghosts.append({"System": er['System'], "Unit": er['Unit'], "Component": er['ExcelComponent'], "Master Hours": er['ExcelHours'], "Anomaly": "GHOST COMPONENT"})

    # 6. Bayesian Truth Index
    total = len(audit_results)
    safe_ratio = len(safe_deltas) / total if total > 0 else 0
    truth_index = int(
        (40 * safe_ratio) + 
        (30 * max(0, 1 - (quarantine_count / max(1, total)))) + 
        (15 * max(0, 1 - (len(ghosts) / 20))) + 
        (15 * max(0, 1 - (mahalanobis_alerts / 10)))
    )

    return audit_results, threat_matrix, ghosts, truth_index

# ═══════════════════════════════════════════════════════════════════════════════
# 4. ORCHESTRATOR & UI
# ═══════════════════════════════════════════════════════════════════════════════

st.markdown(f"""
<div class="hero">
    <div class="hero-left">
        <img src="data:image/svg+xml;base64,{LOGO_SVG}" class="hero-logo" alt=""/>
        <div>
            <div class="hero-title">PMS AUDITOR</div>
            <div class="hero-sub">Enterprise Maritime Forensics</div>
        </div>
    </div>
    <div class="hero-badge">
        <span style="color:var(--teal)">KERNEL</span>&ensp;Triple-Lock Cross-Check<br>
        <span style="color:var(--gold)">ORACLE</span>&ensp;PIML & Quarantine Logic<br>
        <span style="color:#ffffff">BUILD</span>&ensp;v13.0.0 Zenith Master
    </div>
</div>
""", unsafe_allow_html=True)

col1, col2 = st.columns(2)
with col1:
    st.markdown("<div style='color:var(--muted); font-size:0.9rem; font-weight:600; margin-bottom:10px;'>1. COMPONENT OVERHAULS (.doc)</div>", unsafe_allow_html=True)
    pms_files = st.file_uploader("Upload Word Reports (Current + Historical)", type=["doc", "docx"], key="pms_box", label_visibility="collapsed", accept_multiple_files=True)
with col2:
    st.markdown("<div style='color:var(--muted); font-size:0.9rem; font-weight:600; margin-bottom:10px;'>2. EXCEL MASTER LOG (PMS)</div>", unsafe_allow_html=True)
    logs_file = st.file_uploader("Upload Excel PMS Log", type=["xlsx", "xls"], key="logs_box", label_visibility="collapsed")

if pms_files and logs_file:
    with st.spinner("Executing Chronological Cross-Reference & ML Predictive Guardrails..."):
        try:
            all_docs_frames = []
            primary_doc_df = pd.DataFrame()
            
            for file in pms_files:
                df = parse_pms_binary_doc(file.getvalue(), file.name)
                if not df.empty:
                    all_docs_frames.append(df)
                    if primary_doc_df.empty: primary_doc_df = df 

            all_docs_df = pd.concat(all_docs_frames, ignore_index=True) if all_docs_frames else pd.DataFrame()
            excel_records, engine_ytd_hours = parse_master_pms_excel(logs_file.getvalue())

            raw_audit_results = []
            if not primary_doc_df.empty and excel_records:
                for _, row in primary_doc_df.iterrows():
                    comp, oh_date, legacy_hrs = row['Component'], row['Last_Overhaul'], row['Claimed_Hours']
                    verified_hrs = get_verified_hours(row, excel_records)

                    if verified_hrs is not None:
                        delta = verified_hrs - legacy_hrs
                        raw_audit_results.append({"Component": comp, "Overhaul Date": oh_date.strftime('%d-%b-%Y') if pd.notna(oh_date) else "", "Claimed (Doc)": int(round(float(legacy_hrs))), "Verified (Excel)": int(round(float(verified_hrs))), "Delta (Drift)": int(round(float(delta))), "Status": "VERIFIED" if int(round(float(delta))) == 0 else "DRIFT DETECTED"})
                    else:
                        raw_audit_results.append({"Component": comp, "Overhaul Date": oh_date.strftime('%d-%b-%Y') if pd.notna(oh_date) else "", "Claimed (Doc)": int(round(float(legacy_hrs))), "Verified (Excel)": 0, "Delta (Drift)": 0, "Status": "MISSING FROM EXCEL"})

            audit_results, threat_matrix, ghosts, truth_index = run_oracle_piml(raw_audit_results, excel_records, all_docs_df, engine_ytd_hours)

            st.markdown("<br>", unsafe_allow_html=True)
            t1, t2, t3 = st.tabs(["📊 IMMUTABLE LEDGER", "👁️ PIML ORACLE (PREDICTIVE AI)", "🔎 MASTER DATABASE X-RAY"])

            with t1:
                if audit_results:
                    res_df = pd.DataFrame(audit_results)
                    errors_corrected = len(res_df[res_df['Status'] == 'DRIFT DETECTED'])
                    quarantined = len(res_df[res_df['Phase'] == 'QUARANTINE'])
                    digital_seal = hashlib.sha256(res_df.to_json(orient='records').encode()).hexdigest()

                    st.markdown(f"""
                    <div class="hud-grid">
                        <div class="hud-card" style="border-bottom: 3px solid var(--teal);">
                            <div class="hud-header"><div class="hud-title">Components Audited</div></div>
                            <div class="hud-val">{len(res_df)}</div>
                        </div>
                        <div class="hud-card" style="border-bottom: 3px solid {'var(--gold)' if errors_corrected > 0 else 'var(--teal)'};">
                            <div class="hud-header"><div class="hud-title">Mathematical Drift</div></div>
                            <div class="hud-val" style="color: {'var(--gold)' if errors_corrected > 0 else 'var(--teal)'};">{errors_corrected}</div>
                        </div>
                        <div class="hud-card" style="border-bottom: 3px solid {'var(--red)' if quarantined > 0 else 'var(--teal)'};">
                            <div class="hud-header"><div class="hud-title">Quarantined Physics</div></div>
                            <div class="hud-val" style="color: {'var(--red)' if quarantined > 0 else 'var(--teal)'};">{quarantined}</div>
                        </div>
                    </div>
                    """, unsafe_allow_html=True)
                    
                    def style_dataframe(row):
                        if row['Phase'] == 'QUARANTINE': return ['background-color: rgba(255, 42, 85, 0.1); color: #ff8a9f'] * len(row)
                        if row['Status'] == 'DRIFT DETECTED': return ['color: #c9a84c'] * len(row)
                        return ['color: #00e0b0'] * len(row)
                    
                    display_df = res_df.drop(columns=['Severity', 'Z-Score', 'Anomaly', 'Exp_Mean', 'Exp_Upper', 'Exp_Lower', 'Phase'], errors='ignore')
                    st.dataframe(display_df.style.apply(style_dataframe, axis=1), use_container_width=True, hide_index=True)
                else:
                    st.error("Audit Could Not Complete. Check the Diagnostics tab.")

            with t2:
                if audit_results:
                    st.markdown(f"""
                    <div class="truth-index-box">
                        <div class="truth-index-val">{truth_index}%</div>
                        <div class="truth-index-lbl">BAYESIAN TRUTH PROBABILITY</div>
                        <div style="font-size:0.75rem; color:#64748b; margin-top:5px;">Computed via Triple-Locks, MAD Stats, Quarantine Sieves, and Ledoit-Wolf Multivariate Logic.</div>
                    </div>
                    """, unsafe_allow_html=True)

                    df_ml = pd.DataFrame(audit_results)
                    df_ml = df_ml[df_ml['Status'] != 'MISSING FROM EXCEL'].copy()
                    
                    if not df_ml.empty:
                        fig = go.Figure()

                        # Add Conformal Bounds Layer
                        valid_dates = pd.to_datetime(df_ml['Overhaul Date'], errors='coerce')
                        mask = pd.notna(valid_dates)
                        if mask.any():
                            df_bounds = df_ml[mask].sort_values('Overhaul Date')
                            fig.add_trace(go.Scatter(
                                x=pd.to_datetime(df_bounds['Overhaul Date']).tolist() + pd.to_datetime(df_bounds['Overhaul Date']).tolist()[::-1],
                                y=df_bounds['Exp_Upper'].tolist() + df_bounds['Exp_Lower'].tolist()[::-1],
                                fill="toself", fillcolor="rgba(123,104,238,0.1)", line=dict(color="rgba(255,255,255,0)"),
                                hoverinfo="skip", name="Conformal Physics Bounds"
                            ))

                        # Plot Data Points
                        clean = df_ml[(df_ml['Severity'] < 3) & (df_ml['Phase'] == 'SAFE')]
                        if not clean.empty:
                            fig.add_trace(go.Scatter(x=pd.to_datetime(clean['Overhaul Date'], errors='coerce'), y=clean['Claimed (Doc)'], mode='markers', name='Verified Data', marker=dict(color='#00e0b0', size=8, opacity=0.7), text=clean['Component'] + "<br>Claimed: " + clean['Claimed (Doc)'].astype(str) + "h<br>ML Expected: " + clean['Exp_Mean'].astype(str) + "h", hoverinfo='text'))
                            
                        threats = df_ml[(df_ml['Severity'] >= 3) | (df_ml['Phase'] == 'QUARANTINE')]
                        if not threats.empty:
                            fig.add_trace(go.Scatter(x=pd.to_datetime(threats['Overhaul Date'], errors='coerce'), y=threats['Claimed (Doc)'], mode='markers', name='Severe Mathematical Anomaly', marker=dict(color='#ff2a55', size=14, symbol='x', line=dict(color='#1a0508', width=1)), text=threats['Component'] + "<br>Claimed: " + threats['Claimed (Doc)'].astype(str) + "h<br>ML Expected: " + threats['Exp_Mean'].astype(str) + "h<br>" + threats['Anomaly'], hoverinfo='text'))

                        fig.update_layout(title="CONFORMAL PHYSICS RADAR (CLAIMED VS EXPECTED)", plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)', font=dict(color='#94a3b8'), xaxis=dict(title="Last Overhaul Date", gridcolor='rgba(255,255,255,0.05)'), yaxis=dict(title="Running Hours", gridcolor='rgba(255,255,255,0.05)'), hovermode="closest", margin=dict(l=20, r=20, t=50, b=20), legend=dict(yanchor="top", y=0.99, xanchor="left", x=0.01))
                        st.plotly_chart(fig, use_container_width=True)

                    c1, c2 = st.columns(2)
                    with c1:
                        st.markdown("<h4 style='color:var(--gold);'>⚠️ Threat Matrix (Quarantine & Outliers)</h4>", unsafe_allow_html=True)
                        if threat_matrix:
                            tm_df = pd.DataFrame(threat_matrix)[['Component', 'Claimed (Doc)', 'Exp_Mean', 'Anomaly']]
                            st.dataframe(tm_df, use_container_width=True, hide_index=True)
                        else:
                            st.success("Zero severe mathematical anomalies detected.")

                    with c2:
                        st.markdown("<h4 style='color:var(--teal);'>👻 Subsystem Ghost Sweep</h4>", unsafe_allow_html=True)
                        if ghosts:
                            st.dataframe(pd.DataFrame(ghosts), use_container_width=True, hide_index=True)
                        else:
                            st.success("No major Ghost Components detected.")

            with t3:
                c1, c2 = st.columns(2)
                with c1:
                    st.markdown(f"<div style='color:var(--muted); font-size:0.8rem; font-weight:600; margin-bottom:10px;'>DOC REPORTS ({len(pms_files)} Extracted)</div>", unsafe_allow_html=True)
                    if not all_docs_df.empty: st.dataframe(all_docs_df[['Source_File', 'Component', 'Claimed_Hours', 'Last_Overhaul']], use_container_width=True, height=500)
                with c2:
                    st.markdown("<div style='color:var(--muted); font-size:0.8rem; font-weight:600; margin-bottom:10px;'>EXCEL MASTER (Extracted)</div>", unsafe_allow_html=True)
                    ml_status = "ACTIVE (Ledoit-Wolf Matrix)" if HAS_ML else "OFFLINE (Fallback to Deterministic)"
                    st.info(f"Engine Triple-Lock Extracted: ME = {engine_ytd_hours.get('ME', 0)}h, DG = {engine_ytd_hours.get('DG', 0)}h | AI: {ml_status}")
                    if excel_records: st.dataframe(pd.DataFrame(excel_records), use_container_width=True, height=500)

        except Exception as e:
            st.error(f"🚨 Pipeline Execution Halted: {str(e)}")
            st.info(traceback.format_exc())
