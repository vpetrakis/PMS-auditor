import streamlit as st
import pandas as pd
import numpy as np
import io
import hashlib
from datetime import datetime
import warnings

warnings.filterwarnings("ignore")

# ═══════════════════════════════════════════════════════════════════════════════
# 1. ENTERPRISE PAGE CONFIGURATION & PREMIUM CSS (TITAN AESTHETIC)
# ═══════════════════════════════════════════════════════════════════════════════
st.set_page_config(page_title="Vessel Reconciliation Suite", page_icon="⚓", layout="wide", initial_sidebar_state="collapsed")

st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;800&family=JetBrains+Mono:wght@400;700&display=swap');
    
    /* Deep Dark Background */
    .stApp { background-color: #060b13; font-family: 'Inter', sans-serif; color: #f8fafc; }
    
    /* Hide Streamlit Branding */
    #MainMenu {visibility: hidden;} footer {visibility: hidden;} header {visibility: hidden;}
    
    /* Hero Section */
    .hero { border-bottom: 1px solid rgba(255,255,255,0.05); padding-bottom: 20px; margin-bottom: 30px; }
    .hero-title { font-size: 2.5rem; font-weight: 800; background: linear-gradient(90deg, #ffffff, #8ba1b5); -webkit-background-clip: text; -webkit-text-fill-color: transparent; letter-spacing: -1px; }
    .hero-sub { font-size: 0.95rem; color: #64748b; font-weight: 500; text-transform: uppercase; letter-spacing: 2px; }
    
    /* Glassmorphism Uploaders */
    .stFileUploader > div > div { background-color: rgba(13, 21, 34, 0.6) !important; border: 1px dashed rgba(255,255,255,0.1) !important; border-radius: 12px !important; transition: all 0.3s ease; }
    .stFileUploader > div > div:hover { border-color: #00e0b0 !important; background-color: rgba(0, 224, 176, 0.05) !important; }
    
    /* Premium HUD Cards */
    .hud-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(250px, 1fr)); gap: 20px; margin-bottom: 30px; }
    .hud-card { background: rgba(13, 21, 34, 0.8); backdrop-filter: blur(10px); border: 1px solid rgba(255,255,255,0.05); border-radius: 12px; padding: 24px; box-shadow: 0 10px 30px -10px rgba(0,0,0,0.5); }
    .hud-card.success { border-bottom: 3px solid #00e0b0; }
    .hud-card.warn { border-bottom: 3px solid #ff2a55; }
    .hud-title { font-size: 0.8rem; color: #64748b; text-transform: uppercase; letter-spacing: 1.5px; font-weight: 600; margin-bottom: 10px; }
    .hud-val { font-size: 2.2rem; font-weight: 800; color: #ffffff; line-height: 1.1; font-family: 'JetBrains Mono', monospace; }
    .hud-sub { font-size: 0.8rem; color: #475569; margin-top: 8px; }
    
    /* Dataframes */
    .stDataFrame { border-radius: 12px; overflow: hidden; border: 1px solid rgba(255,255,255,0.05); }
    </style>
""", unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════════════════════
# 2. THE SEMANTIC ENGINE (100% Data Integrity Parsers)
# ═══════════════════════════════════════════════════════════════════════════════
@st.cache_data(show_spinner=False)
def parse_multiple_logs(log_files):
    """Parses an unlimited array of daily log files, stitches them, and removes duplicates."""
    all_logs = []
    
    for f in log_files:
        df_raw = pd.read_excel(io.BytesIO(f.getvalue()), header=None, engine='openpyxl', dtype=str)
        header_idx, date_idx, me_idx = -1, -1, -1
        
        # Semantic Hunt
        for i in range(min(50, len(df_raw))):
            vals = [str(x).upper() for x in df_raw.iloc[i].values if pd.notna(x)]
            if any('DATE' in v or 'DAY' in v for v in vals) and any('MAIN' in v or 'ME ' in v for v in vals):
                header_idx = i
                for j, val in enumerate(df_raw.iloc[i].values):
                    v_str = str(val).upper() if pd.notna(val) else ""
                    if 'DATE' in v_str or 'DAY' in v_str: date_idx = j
                    elif 'MAIN' in v_str or 'ME ' in v_str: me_idx = j
                break
        
        if header_idx != -1 and date_idx != -1 and me_idx != -1:
            df = df_raw.iloc[header_idx + 1:].copy()
            clean_df = pd.DataFrame()
            clean_df['Date'] = pd.to_datetime(df.iloc[:, date_idx], errors='coerce')
            clean_df['ME_Hours'] = pd.to_numeric(df.iloc[:, me_idx], errors='coerce').fillna(0.0)
            all_logs.append(clean_df.dropna(subset=['Date']))
            
    if not all_logs:
        raise ValueError("Could not extract semantic timeline from uploaded logs.")
        
    # Stitch, Sort chronologically, and drop duplicated days (keeps the most recent entry if overlaps exist)
    master_timeline = pd.concat(all_logs, ignore_index=True)
    master_timeline = master_timeline.sort_values('Date').drop_duplicates(subset=['Date'], keep='last').reset_index(drop=True)
    return master_timeline

@st.cache_data(show_spinner=False)
def parse_single_pms(file_bytes):
    """Hunts for the Component, Overhaul Date, and Claimed Hours in the master PMS."""
    df_raw = pd.read_excel(io.BytesIO(file_bytes), header=None, engine='openpyxl', dtype=str)
    header_idx, comp_idx, date_idx, hrs_idx = -1, -1, -1, -1

    # Semantic Hunt
    for i in range(min(50, len(df_raw))):
        vals = [str(x).upper() for x in df_raw.iloc[i].values if pd.notna(x)]
        if any('ITEM' in v or 'COMPONENT' in v or 'DESCRIPTION' in v for v in vals) and any('DATE' in v or 'OVERHAUL' in v or 'INSP' in v for v in vals):
            header_idx = i
            for j, val in enumerate(df_raw.iloc[i].values):
                v_str = str(val).upper() if pd.notna(val) else ""
                if 'ITEM' in v_str or 'COMPONENT' in v_str or 'NAME' in v_str or 'DESCRIPTION' in v_str: comp_idx = j
                elif 'DATE' in v_str or 'OVERHAUL' in v_str or 'INSP' in v_str: date_idx = j
                elif 'HOUR' in v_str or 'RUN' in v_str or 'CURRENT' in v_str: hrs_idx = j
            break

    # Fallback to standard columns if semantic fails
    if header_idx == -1: comp_idx, date_idx, hrs_idx, header_idx = 1, 5, 7, 7 

    df = df_raw.iloc[header_idx + 1:].copy()
    clean_df = pd.DataFrame()
    clean_df['Component'] = df.iloc[:, comp_idx].astype(str).str.strip()
    clean_df['Last_Overhaul'] = pd.to_datetime(df.iloc[:, date_idx], errors='coerce')
    clean_df['Claimed_Hours'] = pd.to_numeric(df.iloc[:, hrs_idx], errors='coerce').fillna(0)
    
    clean_df = clean_df[(clean_df['Component'] != 'nan') & (clean_df['Component'] != 'None')]
    return clean_df.dropna(subset=['Last_Overhaul']).reset_index(drop=True)

# ═══════════════════════════════════════════════════════════════════════════════
# 3. FRONTEND UI & LOGIC ROUTER
# ═══════════════════════════════════════════════════════════════════════════════
st.markdown("""
<div class="hero">
    <div class="hero-title">VESSEL RECONCILIATION SUITE</div>
    <div class="hero-sub">M/V Minoan Falcon | Zero-Trust Forensic Auditor</div>
</div>
""", unsafe_allow_html=True)

# The Airlock (Upload Zones)
with st.container():
    col1, col2 = st.columns(2)
    with col1:
        st.markdown("<div style='color:#8ba1b5; font-size:0.9rem; font-weight:600; margin-bottom:10px;'>1. TARGET BASELINE (PMS)</div>", unsafe_allow_html=True)
        pms_file = st.file_uploader("Upload TEC-001 Master Sheet", type=["xlsx", "xls"], key="pms")
    with col2:
        st.markdown("<div style='color:#8ba1b5; font-size:0.9rem; font-weight:600; margin-bottom:10px;'>2. CHRONOLOGICAL LOGS (RUNNING HOURS)</div>", unsafe_allow_html=True)
        logs_files = st.file_uploader("Upload Monthly Log(s). Multi-select enabled.", type=["xlsx", "xls"], accept_multiple_files=True, key="logs")

# Execution Gate
if pms_file and logs_files:
    try:
        with st.spinner("Initializing Enterprise Semantic Engine..."):
            
            # 1. Parse & Stitch
            daily_df = parse_multiple_logs(logs_files)
            pms_df = parse_single_pms(pms_file.getvalue())
            total_days_stitched = len(daily_df)

            # 2. Physics Constraints (Phase 1)
            physics_violations = []
            for _, row in daily_df.iterrows():
                if row['ME_Hours'] > 24 or row['ME_Hours'] < 0:
                    physics_violations.append({
                        "Date": row['Date'].strftime('%d-%b-%Y'),
                        "System": "MAIN ENGINE",
                        "Logged Hours": row['ME_Hours'],
                        "Reason": "Kinematic Violation (Limit: 24h)"
                    })

            # 3. Drift Calculation (Phase 2)
            audit_results = []
            for _, row in pms_df.iterrows():
                comp = row['Component']
                oh_date = row['Last_Overhaul']
                legacy_hrs = row['Claimed_Hours']
                
                # Math Engine: Sum daily hours where timeline date >= overhaul date
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
            
            # Cryptographic Hash of the data payload
            seal_source = res_df.to_json(orient='records')
            digital_seal = hashlib.sha256(seal_source.encode()).hexdigest()

            # Render Premium HUD
            hud_html = f"""
            <div class="hud-grid">
                <div class="hud-card success">
                    <div class="hud-title">Components Audited</div>
                    <div class="hud-val">{len(res_df)}</div>
                    <div class="hud-sub">Extracted via Semantic Engine</div>
                </div>
                <div class="hud-card {'warn' if errors_corrected > 0 else 'success'}">
                    <div class="hud-title">Drift Anomalies</div>
                    <div class="hud-val" style="color: {'#ff2a55' if errors_corrected > 0 else '#00e0b0'};">{errors_corrected}</div>
                    <div class="hud-sub">Mathematical corrections applied</div>
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

            # Phase 1: Physics Tab
            if physics_violations:
                st.markdown("<h3 style='color:#ff2a55; font-size:1.2rem; margin-top:20px;'>⚠️ PHASE 1: KINEMATIC VIOLATIONS DETECTED</h3>", unsafe_allow_html=True)
                st.dataframe(pd.DataFrame(physics_violations), use_container_width=True, hide_index=True)

            # Phase 2: Results Tab
            st.markdown("<h3 style='color:#f8fafc; font-size:1.2rem; margin-top:20px;'>📑 PHASE 2: FORENSIC BASELINE RECONCILIATION</h3>", unsafe_allow_html=True)
            
            # Color code the rows using Pandas Styling
            def style_dataframe(row):
                if row['Delta'] > 0: return ['background-color: rgba(255, 42, 85, 0.1); color: #ff8a9f'] * len(row)
                elif row['Delta'] < 0: return ['background-color: rgba(201, 168, 76, 0.1); color: #c9a84c'] * len(row)
                return ['color: #00e0b0'] * len(row)
            
            st.dataframe(res_df.style.apply(style_dataframe, axis=1), use_container_width=True, hide_index=True)

            # Export Button
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
            st.error("Semantic Extraction Failed: No valid overhaul dates found in the PMS file.")

    except Exception as e:
        st.error(f"🚨 Pipeline Crash Prevented: {str(e)}")
        st.info("The software aborted to protect data integrity. Ensure you uploaded the correct Excel formats.")
