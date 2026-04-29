import streamlit as st
import pandas as pd
import numpy as np
import io
import re
import hashlib
from datetime import datetime
import warnings

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
# 2. THE TARGETED EXTRACTION ENGINES
# ═══════════════════════════════════════════════════════════════════════════════

def parse_daily_hours_excel(file_bytes):
    """Engine 1: Designed STRICTLY for the Excel Timeline (TEC-001 format)."""
    df_raw = pd.read_excel(io.BytesIO(file_bytes), header=None, engine='openpyxl', dtype=str)
    header_idx, date_idx, hrs_idx = -1, -1, -1
    
    # Hunt for exact terminology provided by user
    for i in range(min(50, len(df_raw))):
        vals = [str(x).upper() for x in df_raw.iloc[i].values if pd.notna(x)]
        if any('DATE' in v for v in vals) and any('OPERATING HOURS' in v for v in vals):
            header_idx = i
            for j, val in enumerate(df_raw.iloc[i].values):
                v_str = str(val).upper() if pd.notna(val) else ""
                if 'DATE' in v_str: date_idx = j
                elif 'OPERATING HOURS' in v_str: hrs_idx = j
            break

    if header_idx != -1 and date_idx != -1 and hrs_idx != -1:
        df = df_raw.iloc[header_idx + 1:].copy()
        clean_df = pd.DataFrame()
        clean_df['Date'] = pd.to_datetime(df.iloc[:, date_idx], errors='coerce', dayfirst=True)
        # Strip text, keep only numbers
        clean_df['ME_Hours'] = df.iloc[:, hrs_idx].apply(lambda x: re.sub(r'[^\d.]', '', str(x)))
        clean_df['ME_Hours'] = pd.to_numeric(clean_df['ME_Hours'], errors='coerce').fillna(0.0)
        return clean_df.dropna(subset=['Date']).reset_index(drop=True), df_raw
    
    return pd.DataFrame(), df_raw

def parse_pms_binary_doc(file_bytes):
    """Engine 2: The ASCII Bell Ripper. Designed STRICTLY for the .doc Overhaul file."""
    # Decode the legacy binary
    raw_text = file_bytes.decode('latin-1', errors='ignore')
    clean_text = raw_text.replace('\x00', '')
    
    # \x07 is the ASCII Bell character Microsoft Word uses to separate table cells
    cells = [c.strip() for c in clean_text.split('\x07') if c.strip()]
    
    extracted_data = []
    
    # Maritime Component Anchor List
    COMPONENTS = ['CYLINDER COVER', 'PISTON ASSEMBLY', 'STUFFING BOX', 'PISTON CROWN', 
                  'CYLINDER LINER', 'EXAUST VALVE', 'EXHAUST VALVE', 'STARTING VALVE', 
                  'SAFETY VALVE', 'FUEL VALVES', 'FUEL PUMP']
    
    date_pattern = r'\b(\d{1,2}[-/\.]\w{2,9}[-/\.]?\d{2,4}|\d{1,2}[-/\.]\d{1,2}[-/\.]\d{2,4})\b'
    
    current_comp = None
    comp_dates = []
    comp_hours = []

    for cell in cells:
        cell_upper = cell.upper()
        
        # 1. Is this cell a component name?
        if any(c in cell_upper for c in COMPONENTS) and len(cell) < 30:
            if current_comp and comp_dates and comp_hours:
                # Save the PREVIOUS component block before starting the new one
                extracted_data.append({
                    'Component': current_comp,
                    # Grab the most recent overhaul date
                    'Last_Overhaul': pd.to_datetime(comp_dates[-1], errors='coerce', dayfirst=True),
                    # Grab the highest running hours recorded in that block
                    'Claimed_Hours': max(comp_hours)
                })
            current_comp = cell_upper
            comp_dates = []
            comp_hours = []
            continue
            
        # 2. If we are tracking a component, hunt for dates and hours
        if current_comp:
            dates_found = re.findall(date_pattern, cell)
            if dates_found:
                comp_dates.extend(dates_found)
                
            # Look for large running hours (ignoring cylinder numbers like '1' or '2')
            nums_found = re.findall(r'\b\d{3,6}\b', cell)
            if nums_found:
                comp_hours.extend([float(n) for n in nums_found])

    # Save the very last component in the file
    if current_comp and comp_dates and comp_hours:
        extracted_data.append({
            'Component': current_comp,
            'Last_Overhaul': pd.to_datetime(comp_dates[-1], errors='coerce', dayfirst=True),
            'Claimed_Hours': max(comp_hours)
        })

    raw_cells_df = pd.DataFrame(cells, columns=["ASCII Extracted Cells"])

    if extracted_data:
        clean_df = pd.DataFrame(extracted_data).dropna(subset=['Last_Overhaul']).reset_index(drop=True)
        return clean_df, raw_cells_df
    
    return pd.DataFrame(), raw_cells_df

# ═══════════════════════════════════════════════════════════════════════════════
# 3. MANUAL INGESTION UI (Restoring Control to the User)
# ═══════════════════════════════════════════════════════════════════════════════
st.markdown("""
<div class="hero">
    <div class="hero-title">VESSEL RECONCILIATION SUITE</div>
    <div class="hero-sub">M/V Minoan Falcon | Zero-Trust Forensic Auditor</div>
</div>
""", unsafe_allow_html=True)

col1, col2 = st.columns(2)

with col1:
    st.markdown("<div style='color:#8ba1b5; font-size:0.9rem; font-weight:600; margin-bottom:10px;'>1. COMPONENT OVERHAUL REPORT (.doc)</div>", unsafe_allow_html=True)
    pms_file = st.file_uploader("Upload the file containing Cylinder Covers, Last O/H, etc.", type=["doc", "docx"], key="pms_box")

with col2:
    st.markdown("<div style='color:#8ba1b5; font-size:0.9rem; font-weight:600; margin-bottom:10px;'>2. DAILY OPERATING HOURS (Excel)</div>", unsafe_allow_html=True)
    logs_file = st.file_uploader("Upload the file containing Dates and Daily Operating Hours", type=["xlsx", "xls"], key="logs_box")

# ═══════════════════════════════════════════════════════════════════════════════
# 4. TRIANGULATION MATH ENGINE
# ═══════════════════════════════════════════════════════════════════════════════
if pms_file and logs_file:
    with st.spinner("Initializing ASCII Bell Ripper & Triangulation Engine..."):
        
        # Execute targeted parsers based on the box the user chose
        pms_df, diag_pms = parse_pms_binary_doc(pms_file.getvalue())
        timeline_df, diag_timeline = parse_daily_hours_excel(logs_file.getvalue())
        
        total_days = len(timeline_df) if not timeline_df.empty else 0
        audit_results = []
        physics_violations = []

        if not timeline_df.empty and not pms_df.empty:
            # 1. Physics Check
            for _, row in timeline_df.iterrows():
                if row['ME_Hours'] > 24 or row['ME_Hours'] < 0:
                    physics_violations.append({"Date": row['Date'].strftime('%d-%b-%Y'), "System": "MAIN ENGINE", "Logged Hours": row['ME_Hours']})

            # 2. Drift Calculation
            for _, row in pms_df.iterrows():
                comp = row['Component']
                oh_date = row['Last_Overhaul']
                legacy_hrs = row['Claimed_Hours']
                
                # Filter the timeline starting from the Overhaul Date
                mask = timeline_df['Date'] >= oh_date
                verified_hrs = timeline_df.loc[mask, 'ME_Hours'].sum()
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
                    <div class="hud-card info"><div class="hud-title">Timeline Stitched</div><div class="hud-val">{total_days}</div></div>
                    <div class="hud-card" style="border-bottom: 3px solid #7b68ee;"><div class="hud-title">Digital Seal (SHA-256)</div><div class="hud-val" style="font-size: 1.2rem; margin-top: 10px;">{digital_seal[:10]}...</div></div>
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
                st.error("Audit Could Not Complete. Please check the Diagnostics tab to ensure the correct files were placed in the correct boxes.")

        with tab2:
            st.markdown("### The Transparency Engine")
            st.markdown("<span style='color:#64748b;'>Review the raw data extracted by the ASCII Bell Ripper and the Timeline Engine. If data is missing from the Audit Dashboard, look here to see what the machine failed to catch.</span>", unsafe_allow_html=True)
            
            c1, c2 = st.columns(2)
            with c1:
                st.subheader("Raw Component Cells (Extracted from .doc)")
                if diag_pms is not None and not diag_pms.empty:
                    st.dataframe(diag_pms, use_container_width=True, height=500)
                else:
                    st.warning("No data extracted from the Component file.")
                    
            with c2:
                st.subheader("Raw Timeline Matrix (Extracted from Excel)")
                if diag_timeline is not None and not diag_timeline.empty:
                    st.dataframe(diag_timeline.head(30), use_container_width=True, height=500)
                else:
                    st.warning("No data extracted from the Timeline file.")
