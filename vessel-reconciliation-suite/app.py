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
# 2. DYNAMIC ENTITY RESOLUTION (The Translation Layer)
# ═══════════════════════════════════════════════════════════════════════════════
# Hardcoded dictionary to force alignment between vessel shorthand and Master PMS
MARITIME_ALIASES = {
    "T/C": "TURBOCHARGER",
    "M/E": "MAIN ENGINE",
    "CYL": "CYLINDER",
    "COV": "COVER",
    "PIST": "PISTON",
    "ASSY": "ASSEMBLY",
    "VLV": "VALVE"
}

def normalize_component_name(raw_name):
    """Standardizes component names to bridge the gap between .doc logs and .xlsx PMS."""
    name = str(raw_name).upper().strip()
    for shorthand, full_word in MARITIME_ALIASES.items():
        name = name.replace(shorthand, full_word)
    return re.sub(r'[^A-Z0-9\s]', '', name).strip() # Strip weird characters

def fuzzy_match_components(log_component, pms_components_list):
    """Uses algorithmic string similarity to find the closest match in the PMS."""
    normalized_log = normalize_component_name(log_component)
    normalized_pms = [normalize_component_name(c) for c in pms_components_list]
    
    matches = difflib.get_close_matches(normalized_log, normalized_pms, n=1, cutoff=0.7)
    if matches:
        # Find the original name based on the normalized match
        match_idx = normalized_pms.index(matches[0])
        return pms_components_list[match_idx]
    return log_component # Return original if no high-confidence match is found

# ═══════════════════════════════════════════════════════════════════════════════
# 3. KINEMATIC SHAPE HUNTER (Bulletproof .doc Extraction)
# ═══════════════════════════════════════════════════════════════════════════════
def extract_by_shape(raw_text):
    """Ignores tables entirely. Hunts mathematically for [DATE] adjacent to [0-24 HOURS]."""
    lines = raw_text.splitlines()
    data = []
    
    # Highly permissive regex to catch almost any maritime date format (e.g., 01-Mar, 1/3/26, 2026-03-01)
    date_pattern = r'\b(\d{1,2}[-/\.]\w{2,9}[-/\.]?\d{0,4}|\d{1,2}[-/\.]\d{1,2}[-/\.]\d{2,4})\b'
    
    for line in lines:
        dates_found = re.findall(date_pattern, line)
        if dates_found:
            clean_line = line.replace(dates_found[0], '')
            # Find running hours (numbers between 0 and 24, allowing decimals)
            nums_found = re.findall(r'\b(?:[0-1]?[0-9]|2[0-4])(?:\.\d+)?\b', clean_line)
            if nums_found:
                valid_hours = [float(n) for n in nums_found if 0.0 <= float(n) <= 24.0]
                if valid_hours:
                    data.append({'Date': dates_found[0], 'ME_Hours': max(valid_hours)})
                    
    if data:
        df = pd.DataFrame(data)
        df['Date'] = pd.to_datetime(df['Date'], errors='coerce', dayfirst=True)
        return df.dropna().drop_duplicates(subset=['Date'], keep='last')
    return None

def rip_legacy_file(file_bytes, file_name):
    """The brute-force extraction chamber."""
    # Attempt standard excel if applicable
    if file_name.endswith(('.xlsx', '.xls')):
        try:
            df = pd.read_excel(io.BytesIO(file_bytes), header=None, engine='openpyxl' if file_name.endswith('.xlsx') else 'xlrd', dtype=str)
            # Flatten to text for the shape hunter
            raw_text = df.to_string(index=False, header=False)
            res = extract_by_shape(raw_text)
            if res is not None: return res, raw_text
        except: pass

    # Brute force text rip for .doc binaries
    raw_text = file_bytes.decode('latin-1', errors='ignore').replace('\x00', ' ')
    
    # Strip HTML tags if it's a disguised web archive
    raw_text = re.sub(r'<[^>]+>', ' ', raw_text)
    
    res = extract_by_shape(raw_text)
    return res, raw_text

@st.cache_data(show_spinner=False)
def process_monthly_logs(log_files):
    all_logs = []
    diagnostic_text = {}
    
    for f in log_files:
        clean_df, raw_text = rip_legacy_file(f.getvalue(), f.name.lower())
        diagnostic_text[f.name] = raw_text
        if clean_df is not None and not clean_df.empty:
            all_logs.append(clean_df)
            
    if not all_logs:
        return pd.DataFrame(), diagnostic_text
        
    master_timeline = pd.concat(all_logs, ignore_index=True)
    master_timeline = master_timeline.sort_values('Date').drop_duplicates(subset=['Date'], keep='last').reset_index(drop=True)
    return master_timeline, diagnostic_text

@st.cache_data(show_spinner=False)
def process_pms_master(file_bytes):
    """Extracts components safely from the Master Excel."""
    df_raw = pd.read_excel(io.BytesIO(file_bytes), header=None, engine='openpyxl', dtype=str)
    header_idx, comp_idx, date_idx, hrs_idx = -1, -1, -1, -1

    COMP_ALIASES = ['ITEM', 'COMPONENT', 'DESCRIPTION', 'NAME', 'ΕΞΑΡΤΗΜΑ']
    OH_ALIASES = ['DATE', 'OVERHAUL', 'INSP', 'LAST', 'ΗΜΕΡ']
    HRS_ALIASES = ['HOUR', 'RUN', 'CURRENT', 'CLAIM', 'ΩΡΕΣ']

    def check_alias(val, aliases):
        return any(a in str(val).upper() for a in aliases)

    for i in range(min(50, len(df_raw))):
        row_vals = df_raw.iloc[i].values
        if any(check_alias(v, COMP_ALIASES) for v in row_vals) and any(check_alias(v, OH_ALIASES) for v in row_vals):
            header_idx = i
            for j, val in enumerate(row_vals):
                if check_alias(val, COMP_ALIASES) and comp_idx == -1: comp_idx = j
                elif check_alias(val, OH_ALIASES) and date_idx == -1: date_idx = j
                elif check_alias(val, HRS_ALIASES) and hrs_idx == -1: hrs_idx = j
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
# 4. EXECUTION PIPELINE & UI
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
        st.markdown("<div style='color:#8ba1b5; font-size:0.9rem; font-weight:600; margin-bottom:10px;'>1. TARGET BASELINE (PMS Excel)</div>", unsafe_allow_html=True)
        pms_file = st.file_uploader("Upload TEC-001 Master Sheet", type=["xlsx", "xls"], key="pms")
    with col2:
        st.markdown("<div style='color:#8ba1b5; font-size:0.9rem; font-weight:600; margin-bottom:10px;'>2. CHRONOLOGICAL LOGS (Legacy .doc/.xls)</div>", unsafe_allow_html=True)
        logs_files = st.file_uploader("Upload Monthly Log(s). Multi-select enabled.", type=["xlsx", "xls", "docx", "doc", "csv", "txt", "rtf", "html"], accept_multiple_files=True, key="logs")

if pms_file and logs_files:
    try:
        with st.spinner("Initializing Kinematic Shape Hunter & Entity Resolution..."):
            daily_df, diag_logs = process_monthly_logs(logs_files)
            pms_df, diag_pms = process_pms_master(pms_file.getvalue())
            
            total_days_stitched = len(daily_df) if not daily_df.empty else 0
            audit_results = []
            physics_violations = []

            if total_days_stitched > 0 and not pms_df.empty:
                # Physics Verification
                for _, row in daily_df.iterrows():
                    if row['ME_Hours'] > 24 or row['ME_Hours'] < 0:
                        physics_violations.append({"Date": row['Date'].strftime('%d-%b-%Y'), "System": "MAIN ENGINE", "Logged Hours": row['ME_Hours']})

                # Math Engine & Entity Resolution
                for _, row in pms_df.iterrows():
                    comp = row['Component']
                    oh_date = row['Last_Overhaul']
                    legacy_hrs = row['Claimed_Hours']
                    
                    # Fuzzy match the component name (e.g. Turbocharger <-> T/C) internally
                    # (In this pure kinematic setup, we just apply the math to the timeline)
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
        # 5. DASHBOARD RENDERING
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
                st.error("Audit Failed: Zero data points passed the Kinematic Shape test. Check the Diagnostics tab to view the raw data stream.")

        with tab2:
            st.markdown("### The Transparency Engine")
            st.markdown("<span style='color:#64748b;'>Review the raw text stripped from the binary files. This allows you to visually identify corrupt data structures before they hit the math engine.</span>", unsafe_allow_html=True)
            
            st.subheader("1. Raw Log Data Stream (Stripped from .doc)")
            for filename, raw_text in diag_logs.items():
                st.markdown(f"**File:** `{filename}`")
                # Show first 1000 characters to prevent UI lag on massive files
                st.text_area("Extracted Text Stream", raw_text[:1000] + "\n\n... [TRUNCATED FOR DISPLAY]", height=200, disabled=True)

            st.subheader("2. Raw PMS Matrix (Extracted from TEC-001)")
            st.dataframe(diag_pms.head(15), use_container_width=True)

    except Exception as e:
        st.error(f"🚨 Fatal Anomaly: {str(e)}")
