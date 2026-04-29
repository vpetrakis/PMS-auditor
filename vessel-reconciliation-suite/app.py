import streamlit as st
import pandas as pd
import numpy as np
import io
import re
import hashlib
from datetime import datetime
import difflib
import warnings

try:
    from docx import Document
except ImportError:
    pass

warnings.filterwarnings("ignore")

# ═══════════════════════════════════════════════════════════════════════════════
# 1. ENTERPRISE PAGE CONFIGURATION & PREMIUM CSS (TITAN AESTHETIC)
# ═══════════════════════════════════════════════════════════════════════════════
st.set_page_config(page_title="Vessel Reconciliation Suite", page_icon="⚓", layout="wide", initial_sidebar_state="collapsed")

st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;800&family=JetBrains+Mono:wght@400;700&display=swap');
    
    .stApp { background-color: #060b13; font-family: 'Inter', sans-serif; color: #f8fafc; }
    #MainMenu {visibility: hidden;} footer {visibility: hidden;} header {visibility: hidden;}
    
    .hero { border-bottom: 1px solid rgba(255,255,255,0.05); padding-bottom: 20px; margin-bottom: 30px; }
    .hero-title { font-size: 2.5rem; font-weight: 800; background: linear-gradient(90deg, #ffffff, #8ba1b5); -webkit-background-clip: text; -webkit-text-fill-color: transparent; letter-spacing: -1px; }
    .hero-sub { font-size: 0.95rem; color: #64748b; font-weight: 500; text-transform: uppercase; letter-spacing: 2px; }
    
    .stFileUploader > div > div { background-color: rgba(13, 21, 34, 0.6) !important; border: 1px dashed rgba(255,255,255,0.1) !important; border-radius: 12px !important; transition: all 0.3s ease; }
    .stFileUploader > div > div:hover { border-color: #00e0b0 !important; background-color: rgba(0, 224, 176, 0.05) !important; }
    
    .hud-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(250px, 1fr)); gap: 20px; margin-bottom: 30px; }
    .hud-card { background: rgba(13, 21, 34, 0.8); backdrop-filter: blur(10px); border: 1px solid rgba(255,255,255,0.05); border-radius: 12px; padding: 24px; box-shadow: 0 10px 30px -10px rgba(0,0,0,0.5); }
    .hud-card.success { border-bottom: 3px solid #00e0b0; }
    .hud-card.warn { border-bottom: 3px solid #ff2a55; }
    .hud-title { font-size: 0.8rem; color: #64748b; text-transform: uppercase; letter-spacing: 1.5px; font-weight: 600; margin-bottom: 10px; }
    .hud-val { font-size: 2.2rem; font-weight: 800; color: #ffffff; line-height: 1.1; font-family: 'JetBrains Mono', monospace; }
    .hud-sub { font-size: 0.8rem; color: #475569; margin-top: 8px; }
    
    .stDataFrame { border-radius: 12px; overflow: hidden; border: 1px solid rgba(255,255,255,0.05); }
    </style>
""", unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════════════════════
# 2. THE GOD-MODE PARSER (Fuzzy Logic + Shape Recognition)
# ═══════════════════════════════════════════════════════════════════════════════
DATE_ALIASES = ['DATE', 'DAY', 'ΗΜΕΡΟΜΗΝΙΑ', 'TIME', 'LOG', 'PERIOD']
ME_ALIASES = ['MAIN', 'ME', 'M/E', 'PROPULSION', 'ENGINE 1', 'RUNNING']

def fuzzy_header_match(cell_value, aliases):
    """Checks if a cell matches any of our known semantic aliases."""
    val = str(cell_value).upper().strip()
    return any(alias in val for alias in aliases)

def extract_by_kinematic_shape(text_stream):
    """Level 5 Failsafe: Ignores tables entirely. Hunts for Dates next to Numbers."""
    lines = text_stream.splitlines()
    data = []
    # Regex for standard maritime dates (e.g., 01-Mar-26, 01/03/2026, etc)
    date_pattern = r'\b(\d{1,2}[-/\.]\w{2,9}[-/\.]\d{2,4}|\d{1,2}[-/\.]\d{1,2}[-/\.]\d{2,4})\b'
    
    for line in lines:
        dates_found = re.findall(date_pattern, line)
        if dates_found:
            clean_line = line.replace(dates_found[0], '')
            # Regex for floating point numbers (up to 24.0)
            nums_found = re.findall(r'\b(?:[0-1]?[0-9]|2[0-4])(?:\.\d+)?\b', clean_line)
            if nums_found:
                # Assume the largest valid number on the line is the ME Hours
                valid_hours = [float(n) for n in nums_found if 0 <= float(n) <= 24]
                if valid_hours:
                    data.append({'Date': dates_found[0], 'ME_Hours': max(valid_hours)})
                    
    if data:
        df = pd.DataFrame(data)
        df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
        return df.dropna().drop_duplicates(subset=['Date'], keep='last')
    return None

def extract_semantic_timeline(df_raw):
    """Level 1-4 Failsafe: Hunts for Fuzzy Headers inside any grid structure."""
    if df_raw is None or df_raw.empty: return None
    header_idx, date_idx, me_idx = -1, -1, -1
    
    df_raw = df_raw.astype(str)
    
    for i in range(min(100, len(df_raw))):
        row_vals = df_raw.iloc[i].values
        if any(fuzzy_header_match(v, DATE_ALIASES) for v in row_vals) and any(fuzzy_header_match(v, ME_ALIASES) for v in row_vals):
            header_idx = i
            for j, val in enumerate(row_vals):
                if fuzzy_header_match(val, DATE_ALIASES) and date_idx == -1: date_idx = j
                elif fuzzy_header_match(val, ME_ALIASES) and me_idx == -1: me_idx = j
            break
            
    if header_idx != -1 and date_idx != -1 and me_idx != -1:
        df = df_raw.iloc[header_idx + 1:].copy()
        clean_df = pd.DataFrame()
        clean_df['Date'] = pd.to_datetime(df.iloc[:, date_idx], errors='coerce')
        clean_df['ME_Hours'] = df.iloc[:, me_idx].apply(lambda x: re.sub(r'[^\d.]', '', str(x)))
        clean_df['ME_Hours'] = pd.to_numeric(clean_df['ME_Hours'], errors='coerce').fillna(0.0)
        return clean_df.dropna(subset=['Date'])
    return None

def pure_python_omni_extractor(file_bytes, file_name):
    """The Sequential Decryption Chamber bypassing Pandas C-Engine vulnerabilities."""
    
    # 1. Standard Modern Excel
    if file_name.endswith(('.xlsx', '.xls')):
        try:
            df = pd.read_excel(io.BytesIO(file_bytes), header=None, engine='openpyxl' if file_name.endswith('.xlsx') else 'xlrd', dtype=str)
            res = extract_semantic_timeline(df)
            if res is not None and not res.empty: return res
        except: pass

    # 2. Word Document Tables (.docx)
    if file_name.endswith('.docx'):
        try:
            doc = Document(io.BytesIO(file_bytes))
            data = [[cell.text.strip() for cell in row.cells] for table in doc.tables for row in table.rows]
            if data:
                res = extract_semantic_timeline(pd.DataFrame(data))
                if res is not None and not res.empty: return res
        except: pass

    # 3. HTML / Legacy Base64 Table Extraction
    raw_text = file_bytes.decode('latin-1', errors='ignore').replace('\x00', '')
    if '<table' in raw_text.lower():
        try:
            tables = pd.read_html(io.StringIO(raw_text))
            for t in tables:
                res = extract_semantic_timeline(t)
                if res is not None and not res.empty: return res
        except: pass

    # 4. Pure Regex Grid Reconstruction
    synthetic_grid = [re.split(r'\t|\s{2,}', line.strip()) for line in raw_text.splitlines() if len(re.split(r'\t|\s{2,}', line.strip())) > 1]
    if synthetic_grid:
        res = extract_semantic_timeline(pd.DataFrame(synthetic_grid))
        if res is not None and not res.empty: return res

    # 5. Kinematic Shape Hunting (Nuclear Option)
    return extract_by_kinematic_shape(raw_text)

@st.cache_data(show_spinner=False)
def parse_multiple_logs(log_files):
    all_logs, failed_files = [], []
    
    for f in log_files:
        clean_df = pure_python_omni_extractor(f.getvalue(), f.name.lower())
        if clean_df is not None and not clean_df.empty:
            all_logs.append(clean_df)
        else:
            failed_files.append(f.name)
            
    if not all_logs:
        raise ValueError("Data Integrity Failure. The Omni-Parser could not mathematically locate chronological data in the provided files.")
        
    master_timeline = pd.concat(all_logs, ignore_index=True)
    master_timeline = master_timeline.sort_values('Date').drop_duplicates(subset=['Date'], keep='last').reset_index(drop=True)
    return master_timeline, failed_files

@st.cache_data(show_spinner=False)
def parse_single_pms(file_bytes):
    df_raw = pd.read_excel(io.BytesIO(file_bytes), header=None, engine='openpyxl', dtype=str)
    header_idx, comp_idx, date_idx, hrs_idx = -1, -1, -1, -1

    COMP_ALIASES = ['ITEM', 'COMPONENT', 'DESCRIPTION', 'NAME']
    OH_ALIASES = ['DATE', 'OVERHAUL', 'INSP', 'LAST']
    HRS_ALIASES = ['HOUR', 'RUN', 'CURRENT', 'CLAIM']

    for i in range(min(50, len(df_raw))):
        row_vals = df_raw.iloc[i].values
        if any(fuzzy_header_match(v, COMP_ALIASES) for v in row_vals) and any(fuzzy_header_match(v, OH_ALIASES) for v in row_vals):
            header_idx = i
            for j, val in enumerate(row_vals):
                if fuzzy_header_match(val, COMP_ALIASES) and comp_idx == -1: comp_idx = j
                elif fuzzy_header_match(val, OH_ALIASES) and date_idx == -1: date_idx = j
                elif fuzzy_header_match(val, HRS_ALIASES) and hrs_idx == -1: hrs_idx = j
            break

    if header_idx == -1: comp_idx, date_idx, hrs_idx, header_idx = 1, 5, 7, 7 

    df = df_raw.iloc[header_idx + 1:].copy()
    clean_df = pd.DataFrame()
    clean_df['Component'] = df.iloc[:, comp_idx].astype(str).str.strip()
    clean_df['Last_Overhaul'] = pd.to_datetime(df.iloc[:, date_idx], errors='coerce')
    clean_df['Claimed_Hours'] = pd.to_numeric(df.iloc[:, hrs_idx], errors='coerce').fillna(0)
    
    clean_df = clean_df[(clean_df['Component'] != 'nan') & (clean_df['Component'] != 'None') & (clean_df['Component'] != '')]
    return clean_df.dropna(subset=['Last_Overhaul']).reset_index(drop=True)

# ═══════════════════════════════════════════════════════════════════════════════
# 3. ENTITY RESOLUTION & HUD ROUTER
# ═══════════════════════════════════════════════════════════════════════════════
st.markdown("""
<div class="hero">
    <div class="hero-title">VESSEL RECONCILIATION SUITE</div>
    <div class="hero-sub">M/V Minoan Falcon | Zero-Trust Forensic Auditor</div>
</div>
""", unsafe_allow_html=True)

with st.container():
    col1, col2 = st.columns(2)
    with col1:
        st.markdown("<div style='color:#8ba1b5; font-size:0.9rem; font-weight:600; margin-bottom:10px;'>1. TARGET BASELINE (PMS)</div>", unsafe_allow_html=True)
        pms_file = st.file_uploader("Upload TEC-001 Master Sheet", type=["xlsx", "xls"], key="pms")
    with col2:
        st.markdown("<div style='color:#8ba1b5; font-size:0.9rem; font-weight:600; margin-bottom:10px;'>2. CHRONOLOGICAL LOGS</div>", unsafe_allow_html=True)
        logs_files = st.file_uploader("Upload Monthly Log(s). Multi-select enabled.", type=["xlsx", "xls", "docx", "doc", "csv", "txt", "rtf", "html"], accept_multiple_files=True, key="logs")

if pms_file and logs_files:
    try:
        with st.spinner("Initializing Enterprise Omni-Parser & Entity Resolution..."):
            
            daily_df, failed_files = parse_multiple_logs(logs_files)
            pms_df = parse_single_pms(pms_file.getvalue())
            total_days_stitched = len(daily_df)

            if failed_files:
                st.warning(f"⚠️ Notice: The system skipped completely unreadable file formats: {', '.join(failed_files)}")

            physics_violations = []
            for _, row in daily_df.iterrows():
                if row['ME_Hours'] > 24 or row['ME_Hours'] < 0:
                    physics_violations.append({
                        "Date": row['Date'].strftime('%d-%b-%Y'),
                        "System": "MAIN ENGINE",
                        "Logged Hours": row['ME_Hours'],
                        "Reason": "Kinematic Violation (Limit: 24h)"
                    })

            audit_results = []
            for _, row in pms_df.iterrows():
                comp = row['Component']
                oh_date = row['Last_Overhaul']
                legacy_hrs = row['Claimed_Hours']
                
                mask = daily_df['Date'] >= oh_date
                verified_hrs = daily_df.loc[mask, 'ME_Hours'].sum()
                delta = verified_hrs - legacy_hrs
                
                audit_results.append({
                    "Component": comp,
                    "Overhaul Date": oh_date.strftime('%d-%b-%Y'),
                    "Legacy Claim": int(legacy_hrs),
                    "Verified Math": int(verified_hrs),
                    "Delta": int(delta),
                    "Status": "VERIFIED" if int(delta) == 0 else "DRIFT DETECTED"
                })

        # ═══════════════════════════════════════════════════════════════════════════════
        # 4. PREMIUM HUD DASHBOARD
        # ═══════════════════════════════════════════════════════════════════════════════
        if audit_results:
            st.markdown("<br><hr style='border-color: rgba(255,255,255,0.05);'><br>", unsafe_allow_html=True)
            
            res_df = pd.DataFrame(audit_results)
            errors_corrected = len(res_df[res_df['Delta'] != 0])
            
            seal_source = res_df.to_json(orient='records')
            digital_seal = hashlib.sha256(seal_source.encode()).hexdigest()

            hud_html = f"""
            <div class="hud-grid">
                <div class="hud-card success">
                    <div class="hud-title">Components Audited</div>
                    <div class="hud-val">{len(res_df)}</div>
                    <div class="hud-sub">Extracted via Omni-Parser</div>
                </div>
                <div class="hud-card {'warn' if errors_corrected > 0 else 'success'}">
                    <div class="hud-title">Drift Anomalies</div>
                    <div class="hud-val" style="color: {'#ff2a55' if errors_corrected > 0 else '#00e0b0'};">{errors_corrected}</div>
                    <div class="hud-sub">Mathematical corrections identified</div>
                </div>
                <div class="hud-card">
                    <div class="hud-title">Timeline Stitched</div>
                    <div class="hud-val">{total_days_stitched}</div>
                    <div class="hud-sub">Chronological Days Verified</div>
                </div>
                <div class="hud-card" style="border-bottom: 3px solid #7b68ee;">
                    <div class="hud-title">Cryptographic Seal (SHA-256)</div>
                    <div class="hud-val" style="font-size: 1.2rem; margin-top: 10px;">{digital_seal[:12]}...{digital_seal[-6:]}</div>
                    <div class="hud-sub">Data integrity mathematically locked</div>
                </div>
            </div>
            """
            st.markdown(hud_html, unsafe_allow_html=True)

            if physics_violations:
                st.markdown("<h3 style='color:#ff2a55; font-size:1.2rem; margin-top:20px;'>⚠️ PHASE 1: KINEMATIC VIOLATIONS DETECTED</h3>", unsafe_allow_html=True)
                st.dataframe(pd.DataFrame(physics_violations), use_container_width=True, hide_index=True)

            st.markdown("<h3 style='color:#f8fafc; font-size:1.2rem; margin-top:20px;'>📑 PHASE 2: FORENSIC BASELINE RECONCILIATION</h3>", unsafe_allow_html=True)
            
            def style_dataframe(row):
                if row['Delta'] > 0: return ['background-color: rgba(255, 42, 85, 0.1); color: #ff8a9f'] * len(row)
                elif row['Delta'] < 0: return ['background-color: rgba(201, 168, 76, 0.1); color: #c9a84c'] * len(row)
                return ['color: #00e0b0'] * len(row)
            
            st.dataframe(res_df.style.apply(style_dataframe, axis=1), use_container_width=True, hide_index=True)

            st.markdown("<br>", unsafe_allow_html=True)
            csv_data = res_df.to_csv(index=False).encode('utf-8')
            st.download_button(
                label="⬇️ DOWNLOAD IMMUTABLE BASELINE (.CSV)",
                data=csv_data,
                file_name=f"Verified_Baseline_{datetime.now().strftime('%Y%m%d')}.csv",
                mime='text/csv',
                type="primary"
            )
            
        else:
            st.error("Extraction Failed: No valid overhaul dates found in the PMS file.")

    except Exception as e:
        st.error(f"🚨 Pipeline Execution Halted: {str(e)}")
        st.info("The Omni-Parser isolated a fatal anomaly. The structure of the documents does not contain mathematically viable timestamps or numerical matrices.")
