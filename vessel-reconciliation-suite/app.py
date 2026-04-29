import streamlit as st
import pandas as pd
import numpy as np
import io
import re
import hashlib
from datetime import datetime
import warnings

try:
    from docx import Document
except ImportError:
    pass

warnings.filterwarnings("ignore")

# ═══════════════════════════════════════════════════════════════════════════════
# 1. ENTERPRISE PAGE CONFIGURATION & PREMIUM CSS
# ═══════════════════════════════════════════════════════════════════════════════
st.set_page_config(page_title="Minoan Falcon | Recon Suite", page_icon="⚓", layout="wide", initial_sidebar_state="collapsed")

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
    
    .hud-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 15px; margin-bottom: 30px; }
    .hud-card { background: rgba(13, 21, 34, 0.8); border: 1px solid rgba(255,255,255,0.05); border-radius: 8px; padding: 20px; }
    .hud-card.success { border-bottom: 3px solid #00e0b0; }
    .hud-card.warn { border-bottom: 3px solid #ff2a55; }
    .hud-card.info { border-bottom: 3px solid #3b82f6; }
    .hud-title { font-size: 0.75rem; color: #64748b; text-transform: uppercase; letter-spacing: 1px; font-weight: 600; margin-bottom: 5px; }
    .hud-val { font-size: 2rem; font-weight: 800; color: #ffffff; line-height: 1.1; font-family: 'JetBrains Mono', monospace; }
    
    .stTabs [data-baseweb="tab-list"] { gap: 24px; }
    .stTabs [data-baseweb="tab"] { height: 50px; white-space: pre-wrap; background-color: transparent; border-radius: 4px; color: #64748b; font-weight: 600; }
    .stTabs [aria-selected="true"] { color: #00e0b0; border-bottom: 2px solid #00e0b0; }
    </style>
""", unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════════════════════
# 2. THE GLASS-BOX OMNI-PARSER (Aggressive Semantic Hunting)
# ═══════════════════════════════════════════════════════════════════════════════
# Expanded Dictionaries to catch legacy and foreign headers
DATE_ALIASES = ['DATE', 'DAY', 'ΗΜΕΡΟΜΗΝΙΑ', 'TIME', 'LOG', 'PERIOD', 'DT']
ME_ALIASES = ['MAIN', 'ME', 'M/E', 'PROPULSION', 'ENGINE', 'RUNNING', 'HRS', 'HOURS']

def fuzzy_match(cell_val, aliases):
    val = str(cell_val).upper().strip()
    return any(alias in val for alias in aliases)

def extract_semantic_timeline(df_raw):
    """Hunts for Dates and Hours. Returns both the cleaned data AND the raw data for transparency."""
    if df_raw is None or df_raw.empty: return None, None
    header_idx, date_idx, me_idx = -1, -1, -1
    df_str = df_raw.astype(str)
    
    for i in range(min(150, len(df_str))):
        row_vals = df_str.iloc[i].values
        if any(fuzzy_match(v, DATE_ALIASES) for v in row_vals) and any(fuzzy_match(v, ME_ALIASES) for v in row_vals):
            header_idx = i
            for j, val in enumerate(row_vals):
                if fuzzy_match(val, DATE_ALIASES) and date_idx == -1: date_idx = j
                elif fuzzy_match(val, ME_ALIASES) and me_idx == -1: me_idx = j
            break
            
    if header_idx != -1 and date_idx != -1 and me_idx != -1:
        df = df_raw.iloc[header_idx + 1:].copy()
        clean_df = pd.DataFrame()
        # Aggressive date coercion
        clean_df['Date'] = pd.to_datetime(df.iloc[:, date_idx], errors='coerce', dayfirst=True)
        # Strip everything except numbers and decimals
        clean_df['ME_Hours'] = df.iloc[:, me_idx].apply(lambda x: re.sub(r'[^\d.]', '', str(x)))
        clean_df['ME_Hours'] = pd.to_numeric(clean_df['ME_Hours'], errors='coerce').fillna(0.0)
        
        # Return the cleaned data, AND the raw table so the user can inspect it
        return clean_df.dropna(subset=['Date']), df_raw
    return None, df_raw

def pure_python_omni_extractor(file_bytes, file_name):
    raw_tables_scanned = []
    
    # 1. Standard Modern Excel
    if file_name.endswith(('.xlsx', '.xls')):
        try:
            df = pd.read_excel(io.BytesIO(file_bytes), header=None, engine='openpyxl' if file_name.endswith('.xlsx') else 'xlrd', dtype=str)
            clean, raw = extract_semantic_timeline(df)
            raw_tables_scanned.append(raw)
            if clean is not None and not clean.empty: return clean, raw_tables_scanned
        except: pass

    # 2. HTML / Legacy Base64 Table Extraction (Cracks fake .doc files)
    raw_text = file_bytes.decode('latin-1', errors='ignore').replace('\x00', '')
    if '<table' in raw_text.lower():
        try:
            tables = pd.read_html(io.StringIO(raw_text))
            for t in tables:
                raw_tables_scanned.append(t)
                clean, raw = extract_semantic_timeline(t)
                if clean is not None and not clean.empty: return clean, raw_tables_scanned
        except: pass

    # 3. Pure Regex Grid Reconstruction
    synthetic_grid = [re.split(r'\t|\s{2,}', line.strip()) for line in raw_text.splitlines() if len(re.split(r'\t|\s{2,}', line.strip())) > 1]
    if synthetic_grid:
        df_grid = pd.DataFrame(synthetic_grid)
        raw_tables_scanned.append(df_grid)
        clean, raw = extract_semantic_timeline(df_grid)
        if clean is not None and not clean.empty: return clean, raw_tables_scanned

    return None, raw_tables_scanned

@st.cache_data(show_spinner=False)
def parse_multiple_logs(log_files):
    all_logs = []
    diagnostic_data = {}
    
    for f in log_files:
        clean_df, raw_scanned = pure_python_omni_extractor(f.getvalue(), f.name.lower())
        diagnostic_data[f.name] = raw_scanned # Save raw data for the Glass Box
        if clean_df is not None and not clean_df.empty:
            all_logs.append(clean_df)
            
    if not all_logs:
        return pd.DataFrame(), diagnostic_data # Return empty to trigger Glass Box review
        
    master_timeline = pd.concat(all_logs, ignore_index=True)
    master_timeline = master_timeline.sort_values('Date').drop_duplicates(subset=['Date'], keep='last').reset_index(drop=True)
    return master_timeline, diagnostic_data

@st.cache_data(show_spinner=False)
def parse_single_pms(file_bytes):
    df_raw = pd.read_excel(io.BytesIO(file_bytes), header=None, engine='openpyxl', dtype=str)
    header_idx, comp_idx, date_idx, hrs_idx = -1, -1, -1, -1

    COMP_ALIASES = ['ITEM', 'COMPONENT', 'DESCRIPTION', 'NAME', 'ΕΞΑΡΤΗΜΑ']
    OH_ALIASES = ['DATE', 'OVERHAUL', 'INSP', 'LAST', 'ΗΜΕΡ']
    HRS_ALIASES = ['HOUR', 'RUN', 'CURRENT', 'CLAIM', 'ΩΡΕΣ']

    for i in range(min(50, len(df_raw))):
        row_vals = df_raw.iloc[i].values
        if any(fuzzy_match(v, COMP_ALIASES) for v in row_vals) and any(fuzzy_match(v, OH_ALIASES) for v in row_vals):
            header_idx = i
            for j, val in enumerate(row_vals):
                if fuzzy_match(val, COMP_ALIASES) and comp_idx == -1: comp_idx = j
                elif fuzzy_match(val, OH_ALIASES) and date_idx == -1: date_idx = j
                elif fuzzy_match(val, HRS_ALIASES) and hrs_idx == -1: hrs_idx = j
            break

    if header_idx == -1: comp_idx, date_idx, hrs_idx, header_idx = 1, 5, 7, 7 

    df = df_raw.iloc[header_idx + 1:].copy()
    clean_df = pd.DataFrame()
    clean_df['Component'] = df.iloc[:, comp_idx].astype(str).str.strip()
    clean_df['Last_Overhaul'] = pd.to_datetime(df.iloc[:, date_idx], errors='coerce', dayfirst=True)
    clean_df['Claimed_Hours'] = pd.to_numeric(df.iloc[:, hrs_idx], errors='coerce').fillna(0)
    
    clean_df = clean_df[(clean_df['Component'] != 'nan') & (clean_df['Component'] != 'None') & (clean_df['Component'] != '')]
    return clean_df.dropna(subset=['Last_Overhaul']).reset_index(drop=True), df_raw

# ═══════════════════════════════════════════════════════════════════════════════
# 3. FRONTEND UI & LOGIC ROUTER
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
        with st.spinner("Initializing Enterprise Omni-Parser..."):
            daily_df, diag_logs = parse_multiple_logs(logs_files)
            pms_df, diag_pms = parse_single_pms(pms_file.getvalue())
            
            total_days_stitched = len(daily_df) if not daily_df.empty else 0
            
            audit_results = []
            physics_violations = []

            if total_days_stitched > 0 and not pms_df.empty:
                for _, row in daily_df.iterrows():
                    if row['ME_Hours'] > 24 or row['ME_Hours'] < 0:
                        physics_violations.append({"Date": row['Date'].strftime('%d-%b-%Y'), "System": "MAIN ENGINE", "Logged Hours": row['ME_Hours']})

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
        # 4. TABBED PREMIUM DASHBOARD
        # ═══════════════════════════════════════════════════════════════════════════════
        st.markdown("<br>", unsafe_allow_html=True)
        tab1, tab2 = st.tabs(["📊 AUDIT DASHBOARD", "🔎 GLASS-BOX DIAGNOSTICS"])

        with tab1:
            if audit_results:
                res_df = pd.DataFrame(audit_results)
                errors_corrected = len(res_df[res_df['Delta'] != 0])
                digital_seal = hashlib.sha256(res_df.to_json(orient='records').encode()).hexdigest()

                hud_html = f"""
                <div class="hud-grid">
                    <div class="hud-card success"><div class="hud-title">Components Audited</div><div class="hud-val">{len(res_df)}</div></div>
                    <div class="hud-card {'warn' if errors_corrected > 0 else 'success'}"><div class="hud-title">Drift Anomalies</div><div class="hud-val" style="color: {'#ff2a55' if errors_corrected > 0 else '#00e0b0'};">{errors_corrected}</div></div>
                    <div class="hud-card info"><div class="hud-title">Timeline Stitched</div><div class="hud-val">{total_days_stitched}</div></div>
                    <div class="hud-card" style="border-bottom: 3px solid #7b68ee;"><div class="hud-title">Digital Seal (SHA-256)</div><div class="hud-val" style="font-size: 1.2rem; margin-top: 10px;">{digital_seal[:12]}...</div></div>
                </div>
                """
                st.markdown(hud_html, unsafe_allow_html=True)

                if physics_violations:
                    st.error(f"⚠️ Phase 1: {len(physics_violations)} Kinematic Violations Detected (Hours > 24)")
                
                def style_dataframe(row):
                    if row['Delta'] > 0: return ['background-color: rgba(255, 42, 85, 0.1); color: #ff8a9f'] * len(row)
                    elif row['Delta'] < 0: return ['background-color: rgba(201, 168, 76, 0.1); color: #c9a84c'] * len(row)
                    return ['color: #00e0b0'] * len(row)
                
                st.dataframe(res_df.style.apply(style_dataframe, axis=1), use_container_width=True, hide_index=True)
                
                csv_data = res_df.to_csv(index=False).encode('utf-8')
                st.download_button("⬇️ DOWNLOAD IMMUTABLE BASELINE (.CSV)", data=csv_data, file_name=f"Verified_Baseline_{datetime.now().strftime('%Y%m%d')}.csv", mime='text/csv', type="primary")
            else:
                st.error("Audit Failed: The engine could not extract enough cross-referencing data. Check the Glass-Box Diagnostics tab.")

        with tab2:
            st.markdown("### Transparency Engine")
            st.markdown("<span style='color:#64748b;'>If your audit results are showing 1 Component or missing days, look at the raw data below. This is exactly what the machine ripped out of your files *before* filtering. Look for strange headers, merged text, or weird date formats.</span>", unsafe_allow_html=True)
            
            st.subheader("1. Raw Log Extraction (What the machine saw in the .doc files)")
            for filename, tables in diag_logs.items():
                st.markdown(f"**File:** `{filename}`")
                if tables:
                    for i, t in enumerate(tables):
                        st.caption(f"Table Matrix {i+1}")
                        st.dataframe(t.head(15), use_container_width=True) # Show first 15 rows
                else:
                    st.warning(f"Engine could not find ANY tables or grids inside {filename}.")

            st.subheader("2. Raw PMS Extraction (What the machine saw in TEC-001)")
            st.dataframe(diag_pms.head(15), use_container_width=True)

    except Exception as e:
        st.error(f"🚨 Fatal Anomaly: {str(e)}")
