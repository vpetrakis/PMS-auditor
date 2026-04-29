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
    "LINK": _u('<svg viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg"><path fill="#3b82f6" d="M3.9 12c0-1.71 1.39-3.1 3.1-3.1h4V7H7c-2.76 0-5 2.24-5 5s2.24 5 5 5h4v-1.9H7c-1.71 0-3.1-1.39-3.1-3.1zM8 13h8v-2H8v2zm9-6h-4v1.9h4c1.71 0 3.1 1.39 3.1 3.1s-1.39 3.1-3.1 3.1h-4V17h4c2.76 0 5-2.24 5-5s-2.24-5-5-5z"/></svg>'),
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
# 3. THE ENTERPRISE EXCEL MASTER-LOG PARSER
# ═══════════════════════════════════════════════════════════════════════════════

@st.cache_data(show_spinner=False)
def parse_master_pms_excel(file_bytes):
    """
    Scans the Excel file for the Master PMS Log (ME & DG MAIN PARTS LOG).
    Extracts every single component and its exact "CURRENT OPERATING HOURS".
    """
    try:
        xls = pd.ExcelFile(io.BytesIO(file_bytes), engine="openpyxl")
        
        # 1. Target the Master Sheet logically
        target_sheet = xls.sheet_names[0]
        for s in xls.sheet_names:
            if "PARTS LOG" in s.upper() or "PMS" in s.upper():
                target_sheet = s
                break
                
        df_raw = pd.read_excel(io.BytesIO(file_bytes), sheet_name=target_sheet, header=None, dtype=object, engine="openpyxl")
    except Exception:
        try:
            df_raw = pd.read_excel(io.BytesIO(file_bytes), header=None, dtype=object, engine="openpyxl")
        except:
            return []

    item_col = -1
    hours_col = -1
    header_row = -1

    # 2. X-Ray the top rows to find the exact Data Columns
    for i in range(min(30, len(df_raw))):
        row_joined = " | ".join([str(x).upper() for x in df_raw.iloc[i].values if pd.notna(x)])
        
        if ("ITEM" in row_joined or "DESCRIPTION" in row_joined) and "CURRENT" in row_joined and "HOURS" in row_joined:
            header_row = i
            for j, val in enumerate(df_raw.iloc[i].values):
                val_u = str(val).upper()
                if "ITEM" in val_u or "DESCRIPTION" in val_u:
                    item_col = j
                elif "CURRENT" in val_u and "HOURS" in val_u:
                    hours_col = j
            break

    # Fallback to pure column scanning if they are on different rows
    if item_col == -1 or hours_col == -1:
        for i in range(min(30, len(df_raw))):
            for j, val in enumerate(df_raw.iloc[i].values):
                val_u = str(val).upper()
                if item_col == -1 and ("ITEM" in val_u or "DESCRIPTION" in val_u): item_col = j
                if hours_col == -1 and ("CURRENT" in val_u and "HOURS" in val_u): hours_col = j

    if item_col == -1 or hours_col == -1:
        return []

    excel_records = []
    curr_sys = "ME"
    curr_unit = ""

    # 3. Sweep Downwards: Tracking Context (System/Unit) and extracting Hours
    for i in range(header_row + 1, len(df_raw)):
        item_val = df_raw.iloc[i, item_col]
        hours_val = df_raw.iloc[i, hours_col]

        if pd.isna(item_val): continue
        item_str = normalize_text(item_val)

        # Update State Context
        if "MAIN ENGINE" in item_str or "M/E" in item_str:
            curr_sys = "ME"
            curr_unit = ""
        elif any(x in item_str for x in ["GENERATOR NO.1", "DG 1", "D/G 1", "NO 1", "NO.1"]):
            curr_sys = "DG"
            curr_unit = "DG1"
        elif any(x in item_str for x in ["GENERATOR NO.2", "DG 2", "D/G 2", "NO 2", "NO.2"]):
            curr_sys = "DG"
            curr_unit = "DG2"
        elif any(x in item_str for x in ["GENERATOR NO.3", "DG 3", "D/G 3", "NO 3", "NO.3"]):
            curr_sys = "DG"
            curr_unit = "DG3"

        # Update Cylinder Context
        if "CYLINDER NO" in item_str or "CYL NO" in item_str:
            m = re.search(r'NO[\.\:\s]*(\d+)', item_str)
            if m: curr_unit = f"Cyl {m.group(1)}"

        # Validate Component Data
        h = extract_first_float(hours_val)
        if pd.notna(h) and len(item_str) > 2 and "HOURS" not in item_str and "DATE" not in item_str:
            excel_records.append({
                'System': curr_sys,
                'Unit': curr_unit,
                'ExcelComponent': item_str,
                'ExcelHours': h
            })

    return excel_records

# ═══════════════════════════════════════════════════════════════════════════════
# 4. WORD DOCUMENT PARSER (ASCII Failsafe)
# ═══════════════════════════════════════════════════════════════════════════════

@st.cache_data(show_spinner=False)
def parse_pms_binary_doc(file_bytes, file_name=""):
    """Extracts claims from legacy .doc and modern .docx."""
    fname = (file_name or "").lower()

    def parse_doc_binary(bytes_):
        raw_text = bytes_.decode("latin-1", errors="ignore").replace("\x00", "")
        return [c.strip() for c in raw_text.split("\x07") if c.strip()]

    def parse_docx_xml(bytes_):
        texts = []
        try:
            with ZipFile(io.BytesIO(bytes_)) as z:
                xml = z.read("word/document.xml")
            root = ET.fromstring(xml)
            ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
            for para in root.findall(".//w:p", ns):
                runs = [t.text for t in para.findall(".//w:t", ns) if t.text]
                line = " ".join(runs).strip()
                if line: texts.append(line)
        except: pass
        return texts

    cells = parse_docx_xml(file_bytes) if fname.endswith(".docx") else parse_doc_binary(file_bytes)
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

    i = 0
    while i < len(cells):
        cell = normalize_text(cells[i])

        if "MAIN ENGINE" in cell:
            system = "ME"
            sub_units = ["Cyl 1", "Cyl 2", "Cyl 3", "Cyl 4", "Cyl 5", "Cyl 6"]
        elif any(x in cell for x in ["AUX. ENGINE", "AUX ENGINE", "D/G", "DG ", "DIESEL GENERATOR"]):
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
                        if j + 1 + k < len(cells): dates.append(cells[j + 1 + k])
                    break
                j += 1

            j = i + 1
            while j < min(i + 35, len(cells)):
                if str(cells[j]).strip() == '2':
                    for k in range(len(sub_units)):
                        if j + 1 + k < len(cells): hours.append(cells[j + 1 + k])
                    break
                j += 1

            for idx, su in enumerate(sub_units):
                if idx < len(dates) and idx < len(hours):
                    d = robust_parse_date(dates[idx])
                    h = extract_first_float(hours[idx])
                    if pd.notna(d) and pd.notna(h):
                        extracted_data.append({
                            'System': system,
                            'Unit': su,
                            'BaseComponent': comp_name,
                            'Component': f"[{system}] {comp_name} ({su})",
                            'Last_Overhaul': d,
                            'Claimed_Hours': float(h)
                        })
        i += 1

    return pd.DataFrame(extracted_data).dropna(subset=['Last_Overhaul']).reset_index(drop=True), pd.DataFrame(cells, columns=["Raw Doc Cells"])

# ═══════════════════════════════════════════════════════════════════════════════
# 5. THE FUZZY VLOOKUP BRIDGE
# ═══════════════════════════════════════════════════════════════════════════════

def get_verified_hours(doc_row, excel_records):
    """Mathematically links the Word component to the exact Excel component row."""
    doc_sys = doc_row['System']
    doc_unit = doc_row['Unit']
    
    # Normalize strings for scoring
    w_comp = re.sub(r'[^A-Z]', '', doc_row['BaseComponent'].replace("ASSY", "ASSEMBLY"))
    
    best_score = 0
    best_hours = 0
    
    for er in excel_records:
        # Strict System Match
        if er['System'] != doc_sys:
            continue
            
        # Unit Match (Permissive if Excel unit is blank)
        if er['Unit'] != doc_unit and er['Unit'] != "":
            continue
            
        e_comp = re.sub(r'[^A-Z]', '', er['ExcelComponent'].replace("ASSY", "ASSEMBLY"))
        
        # Calculate mathematical string similarity
        score = SequenceMatcher(None, w_comp, e_comp).ratio()
        
        # Substring Override (e.g., 'CYLINDERCOVER' inside 'MECYLINDERCOVER')
        if w_comp in e_comp or e_comp in w_comp:
            score += 0.3
            
        if score > best_score:
            best_score = score
            best_hours = er['ExcelHours']

    # Threshold (0.55 guarantees a solid match without requiring perfection)
    if best_score > 0.55:
        return best_hours
    return None

# ═══════════════════════════════════════════════════════════════════════════════
# 6. MAIN FRONTEND ORCHESTRATOR
# ═══════════════════════════════════════════════════════════════════════════════

st.markdown(f"""
<div class="hero">
    <div class="hero-left">
        <img src="data:image/svg+xml;base64,{LOGO_SVG}" class="hero-logo" alt=""/>
        <div>
            <div class="hero-title">POSEIDON RECON</div>
            <div class="hero-sub">Vlookup Cross-Reference Engine</div>
        </div>
    </div>
    <div class="hero-badge">
        <span style="color:#00e0b0">KERNEL</span>&ensp;Fuzzy Data Bridge<br>
        <span style="color:#00e0b0">SOURCE</span>&ensp;Master Parts Log<br>
        <span style="color:#fff">BUILD</span>&ensp;v10.3.0 Zenith
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
    with st.spinner("Executing Master Cross-Reference Matcher..."):
        try:
            # 1. Isolate the Extractions
            doc_df, diag_pms = parse_pms_binary_doc(pms_file.getvalue(), pms_file.name)
            excel_records = parse_master_pms_excel(logs_file.getvalue())

            audit_results = []

            # 2. Execute the VLOOKUP Mathematical Bridge
            if not doc_df.empty and excel_records:
                for _, row in doc_df.iterrows():
                    comp = row['Component']
                    oh_date = row['Last_Overhaul']
                    legacy_hrs = row['Claimed_Hours']
                    
                    verified_hrs = get_verified_hours(row, excel_records)

                    if verified_hrs is not None:
                        delta = verified_hrs - legacy_hrs
                        audit_results.append({
                            "Component": comp,
                            "Overhaul Date": oh_date.strftime('%d-%b-%Y') if pd.notna(oh_date) else "",
                            "Claimed (Doc)": int(round(float(legacy_hrs))),
                            "Verified (Excel)": int(round(float(verified_hrs))),
                            "Delta (Drift)": int(round(float(delta))),
                            "Status": "VERIFIED" if int(round(float(delta))) == 0 else "DRIFT DETECTED"
                        })
                    else:
                        audit_results.append({
                            "Component": comp,
                            "Overhaul Date": oh_date.strftime('%d-%b-%Y') if pd.notna(oh_date) else "",
                            "Claimed (Doc)": int(round(float(legacy_hrs))),
                            "Verified (Excel)": 0,
                            "Delta (Drift)": 0,
                            "Status": "MISSING FROM EXCEL"
                        })

            # ═══════════════════════════════════════════════════════════════════════════════
            # 7. DASHBOARD RENDERING
            # ═══════════════════════════════════════════════════════════════════════════════
            st.markdown("<br>", unsafe_allow_html=True)
            t1, t2, t3 = st.tabs(["📊 IMMUTABLE LEDGER", "📈 FORENSIC PLOT", "🔎 MASTER DATABASE X-RAY"])

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
                            <div class="hud-sub">Math/Log deviations found</div>
                        </div>
                        <div class="hud-card" style="border-bottom: 3px solid #3b82f6;">
                            <div class="hud-header">
                                <div class="hud-title">Master Log Linked</div>
                                <div class="hud-icon"><img src="{ICONS['LINK']}"></div>
                            </div>
                            <div class="hud-val">{len(excel_records)}</div>
                            <div class="hud-sub">Excel Database Rows Scanned</div>
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
                    
                    st.dataframe(res_df.style.apply(style_dataframe, axis=1), use_container_width=True, hide_index=True)
                    
                    csv_data = res_df.to_csv(index=False).encode('utf-8')
                    st.download_button("⬇️ EXPORT FORENSIC LEDGER (.CSV)", data=csv_data, file_name="Reconciliation_Audit.csv", mime='text/csv')
                else:
                    st.error("Audit Could Not Complete. Check the Diagnostics tab to ensure the Excel file contains the Master Parts Log.")

            with t2:
                if audit_results:
                    st.markdown("### Claimed vs Verified Master Hours")
                    st.markdown("<span style='color:#64748b; font-size:0.85rem;'>Native Streamlit rendering (Zero external chart dependencies)</span><br><br>", unsafe_allow_html=True)
                    plot_df = res_df[['Component', 'Claimed (Doc)', 'Verified (Excel)']].set_index('Component')
                    st.bar_chart(plot_df, color=["#c9a84c", "#00e0b0"], height=500)

            with t3:
                st.markdown("### The Transparency Engine")
                
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
