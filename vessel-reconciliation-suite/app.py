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
# 2. THE INTELLIGENT AUTO-ROUTER & EXTRACTION ENGINES
# ═══════════════════════════════════════════════════════════════════════════════

# Known components to anchor the ASCII Bell Ripper
KNOWN_COMPONENTS = ['CYLINDER COVER', 'PISTON ASSEMBLY', 'STUFFING BOX', 'PISTON CROWN', 'CYLINDER LINER', 'EXAUST VALVE', 'STARTING VALVE', 'SAFETY VALVE', 'FUEL VALVES', 'FUEL PUMP']

def parse_timeline_excel(file_bytes):
    """Extracts Date and Hours from the 'DAILY OPERATING HOURS' Excel file."""
    df_raw = pd.read_excel(io.BytesIO(file_bytes), header=None, engine='openpyxl', dtype=str)
    header_idx, date_idx, hrs_idx = -1, -1, -1
    
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
        clean_df['ME_Hours'] = df.iloc[:, hrs_idx].apply(lambda x: re.sub(r'[^\d.]', '', str(x)))
        clean_df['ME_Hours'] = pd.to_numeric(clean_df['ME_Hours'], errors='coerce').fillna(0.0)
        return clean_df.dropna(subset=['Date']), df_raw
    return pd.DataFrame(), df_raw

def parse_pms_binary_doc(file_bytes):
    """The ASCII Bell Ripper: Cracks the 1997 OLE2 .doc table using the hidden \x07 delimiter."""
    raw_text = file_bytes.decode('latin-1', errors='ignore')
    clean_text = raw_text.replace('\x00', '')
    
    # Split the binary file entirely by the Microsoft Word Table Cell Delimiter (\x07)
    cells = [c.strip() for c in clean_text.split('\x07') if c.strip()]
    
    extracted_data = []
    current_component = None
    dates_buffer = []
    hours_buffer = []

    date_pattern = r'\b(\d{1,2}[-/\.]\w{2,9}[-/\.]?\d{2,4}|\d{1,2}[-/\.]\d{1,2}[-/\.]\d{2,4})\b'
    
    for cell in cells:
        # 1. Identify Component
        if any(comp in cell.upper() for comp in KNOWN_COMPONENTS) and len(cell) < 30:
            # If we already had a component, save it before starting the next one
            if current_component and dates_buffer and hours_buffer:
                extracted_data.append({
                    'Component': current_component,
                    # Safe date parsing
                    'Last_Overhaul': pd.to_datetime(dates_buffer[-1], errors='coerce', dayfirst=True),
                    'Claimed_Hours': max(hours_buffer) # Running hours are usually the largest number
                })
            current_component = cell.upper()
            dates_buffer = []
            hours_buffer = []
            continue
            
        if current_component:
            # 2. Extract Dates for this component
            dates_found = re.findall(date_pattern, cell)
            if dates_found:
                dates_buffer.extend(dates_found)
                
            # 3. Extract Running Hours (>100 to avoid confusing with cylinder numbers like '1' or '2')
            nums_found = re.findall(r'\b\d{3,6}\b', cell)
            if nums_found:
                hours_buffer.extend([float(n) for n in nums_found])

    # Append the last component in the file
    if current_component and dates_buffer and hours_buffer:
        extracted_data.append({
            'Component': current_component,
            'Last_Overhaul': pd.to_datetime(dates_buffer[-1], errors='coerce', dayfirst=True),
            'Claimed_Hours': max(hours_buffer)
        })

    if extracted_data:
        df = pd.DataFrame(extracted_data).dropna(subset=['Last_Overhaul'])
        return df, pd.DataFrame(cells, columns=["Raw ASCII Bytes"])
    
    return pd.DataFrame(), pd.DataFrame(cells, columns=["Raw ASCII Bytes"])


# ═══════════════════════════════════════════════════════════════════════════════
# 3. MAIN PIPELINE EXECUTION
# ═══════════════════════════════════════════════════════════════════════════════
st.markdown("""
<div class="hero">
    <div class="hero-title">VESSEL RECONCILIATION SUITE</div>
    <div class="hero-sub">M/V Minoan Falcon | Zero-Trust Forensic Auditor</div>
</div>
""", unsafe_allow_html=True)

st.markdown("<div style='color:#8ba1b5; font-size:0.9rem; font-weight:600; margin-bottom:10px;'>SECURE INGESTION ZONE</div>", unsafe_allow_html=True)
st.markdown("<span style='color:#475569; font-size:0.8rem;'>Drag and drop ALL files (Daily Hours Excel & Overhaul .doc) into this single bucket. The Auto-Router will classify them.</span>", unsafe_allow_html=True)

all_files = st.file_uploader("Drop Files Here", type=["xlsx", "xls", "docx", "doc", "csv"], accept_multiple_files=True, label_visibility="collapsed")

if all_files:
    with st.spinner("Initializing Auto-Router & ASCII Bell Ripper..."):
        
        timeline_df = pd.DataFrame()
        pms_df = pd.DataFrame()
        diag_timeline = None
        diag_pms = None
        
        # --- THE AUTO-ROUTER ---
        for f in all_files:
            file_bytes = f.getvalue()
            raw_text = file_bytes.decode('latin-1', errors='ignore').upper()
            
            # Classification Logic
            if "DAILY OPERATING HOURS" in raw_text:
                st.toast(f"Routed '{f.name}' to Timeline Engine", icon="⏱️")
                timeline_df, diag_timeline = parse_timeline_excel(file_bytes)
            elif "LAST O/H" in raw_text or "CYLINDER COVER" in raw_text:
                st.toast(f"Routed '{f.name}' to PMS Engine", icon="⚙️")
                pms_df, diag_pms = parse_pms_binary_doc(file_bytes)
            else:
                st.warning(f"File '{f.name}' could not be classified. Skipping.")

        # --- MATH & TRIANGULATION ENGINE ---
        total_days = len(timeline_df) if not timeline_df.empty else 0
        audit_results = []
        physics_violations = []

        if not timeline_df.empty and not pms_df.empty:
            for _, row in timeline_df.iterrows():
                if row['ME_Hours'] > 24 or row['ME_Hours'] < 0:
                    physics_violations.append({"Date": row['Date'].strftime('%d-%b-%Y'), "System": "MAIN ENGINE", "Logged Hours": row['ME_Hours']})

            for _, row in pms_df.iterrows():
                comp = row['Component']
                oh_date = row['Last_Overhaul']
                legacy_hrs = row['Claimed_Hours']
                
                # Math: Filter Timeline >= Overhaul Date
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
        # 4. DASHBOARD RENDERING
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
                st.error("Audit Could Not Complete. Please check the Diagnostics tab.")

        with tab2:
            st.markdown("### The Transparency Engine")
            st.markdown("<span style='color:#64748b;'>Review the raw data extracted by the ASCII Bell Ripper and the Timeline Engine.</span>", unsafe_allow_html=True)
            
            c1, c2 = st.columns(2)
            with c1:
                st.subheader("Raw Component Cells (.doc)")
                if diag_pms is not None:
                    st.dataframe(diag_pms, use_container_width=True, height=400)
                else:
                    st.info("No .doc PMS file routed yet.")
                    
            with c2:
                st.subheader("Raw Timeline Matrix (Excel)")
                if diag_timeline is not None:
                    st.dataframe(diag_timeline.head(20), use_container_width=True, height=400)
                else:
                    st.info("No Excel Timeline file routed yet.")
