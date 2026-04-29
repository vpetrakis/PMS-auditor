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

warnings.filterwarnings("ignore")

# ═══════════════════════════════════════════════════════════════════════════════
# 1. ENTERPRISE PAGE CONFIGURATION & POSEIDON TITAN CSS
# ═══════════════════════════════════════════════════════════════════════════════
st.set_page_config(page_title="POSEIDON | Recon Suite", page_icon="⚓", layout="wide", initial_sidebar_state="collapsed")

def _u(s): return f"data:image/svg+xml;base64,{base64.b64encode(s.encode()).decode()}"

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
# 2. THE REVERSE-ANCHOR EXTRACTION ENGINES
# ═══════════════════════════════════════════════════════════════════════════════

@st.cache_data(show_spinner=False)
def parse_daily_hours_excel(file_bytes):
    """The Reverse-Anchor Plumb Line Engine."""
    try:
        df_raw = pd.read_excel(io.BytesIO(file_bytes), sheet_name='DAILY OPERATING HOURS', header=None, engine='openpyxl', dtype=str)
    except:
        df_raw = pd.read_excel(io.BytesIO(file_bytes), header=None, engine='openpyxl', dtype=str)
        
    col_map = {}
    header_row_2 = -1
    
    # 1. Sweep vertically to find the exact row that contains DATE and OPERATING HOURS
    for i in range(min(60, len(df_raw))):
        row_str = " ".join([str(x).upper() for x in df_raw.iloc[i].values if pd.notna(x)])
        if 'DATE' in row_str and ('OPERATING' in row_str or 'HOURS' in row_str):
            header_row_2 = i
            break

    # 2. Reverse-Anchor: Look one row up to establish the Parents and forward-fill
    if header_row_2 != -1:
        header_row_1 = max(0, header_row_2 - 1)
        
        parent_headers = df_raw.iloc[header_row_1].ffill().fillna("").astype(str).str.upper()
        sub_headers = df_raw.iloc[header_row_2].fillna("").astype(str).str.upper()
        
        for col_idx in range(len(df_raw.columns)):
            parent = parent_headers.iloc[col_idx]
            sub = sub_headers.iloc[col_idx]
            
            if 'DATE' in sub or 'DATE' in parent or 'DAY' in sub:
                col_map['Date'] = col_idx
            elif 'OPERATING' in sub or 'HOURS' in sub or 'RUN' in sub:
                if 'MAIN' in parent or 'M/E' in parent or 'ME ' in parent:
                    col_map['ME_Hours'] = col_idx
                elif 'GEN' in parent and ('1' in parent or 'ONE' in parent):
                    col_map['DG1_Hours'] = col_idx
                elif 'GEN' in parent and ('2' in parent or 'TWO' in parent):
                    col_map['DG2_Hours'] = col_idx
                elif 'GEN' in parent and ('3' in parent or 'THREE' in parent):
                    col_map['DG3_Hours'] = col_idx

    # 3. Extract Timeline Data
    if 'Date' in col_map:
        df = df_raw.iloc[header_row_2 + 1:].copy()
        clean_df = pd.DataFrame()
        clean_df['Date'] = pd.to_datetime(df.iloc[:, col_map['Date']], errors='coerce', dayfirst=True)
        
        for sys_col in ['ME_Hours', 'DG1_Hours', 'DG2_Hours', 'DG3_Hours']:
            if sys_col in col_map:
                clean_df[sys_col] = df.iloc[:, col_map[sys_col]].apply(lambda x: re.sub(r'[^\d.]', '', str(x)))
                clean_df[sys_col] = pd.to_numeric(clean_df[sys_col], errors='coerce').fillna(0.0)
            else:
                clean_df[sys_col] = 0.0
                
        return clean_df.dropna(subset=['Date']).reset_index(drop=True), df_raw, col_map
    
    return pd.DataFrame(), df_raw, col_map

@st.cache_data(show_spinner=False)
def parse_pms_binary_doc(file_bytes):
    """The Horizontal Matrix Resolver (ASCII Bell Ripper with State Machine)."""
    raw_text = file_bytes.decode('latin-1', errors='ignore').replace('\x00', '')
    cells = [c.strip() for c in raw_text.split('\x07') if c.strip()]
    extracted_data = []
    
    COMPONENTS = ['CYLINDER COVER', 'PISTON ASSEMBLY', 'STUFFING BOX', 'PISTON CROWN', 'CYLINDER LINER', 
                  'EXAUST VALVE', 'STARTING VALVE', 'SAFETY VALVE', 'FUEL VALVES', 'FUEL PUMP', 
                  'SUCTION VALVE', 'PUNCTURE VALVE', 'CROSSHEAD BEARINGS', 'BOTTOM END BEARINGS', 
                  'MAIN BEARINGS', 'CYLINDER HEAD', 'PISTON', 'CONNECTING ROD', 'TURBOCHARGER', 
                  'AIR COOLER', 'COOLING WATER PUMP', 'THERMOSTAT VALVE', 'THRUST BEARING']
                  
    system = "ME"
    sub_units = ["Cyl 1", "Cyl 2", "Cyl 3", "Cyl 4", "Cyl 5", "Cyl 6"]
    target_timeline = "ME_Hours"
    
    i = 0
    while i < len(cells):
        cell = cells[i].upper()
        
        # Phase 1: Parent State Tracker
        if "MAIN ENGINE" in cell:
            system = "ME"
            target_timeline = "ME_Hours"
            sub_units = ["Cyl 1", "Cyl 2", "Cyl 3", "Cyl 4", "Cyl 5", "Cyl 6"]
        elif "AUX. ENGINE" in cell or "D/G" in cell:
            system = "DG"
            sub_units = ["DG1", "DG2", "DG3"]
            
        # Phase 2: Component Lock-On
        is_comp = any(c in cell for c in COMPONENTS) and len(cell) < 40
        if is_comp:
            comp_name = cell
            dates, hours = [], []
            
            j = i + 1
            while j < min(i + 15, len(cells)):
                if cells[j].strip() == '1':
                    for k in range(len(sub_units)):
                        if j + 1 + k < len(cells): dates.append(cells[j + 1 + k])
                    break
                j += 1
                
            j = i + 1
            while j < min(i + 30, len(cells)):
                if cells[j].strip() == '2':
                    for k in range(len(sub_units)):
                        if j + 1 + k < len(cells): hours.append(cells[j + 1 + k])
                    break
                j += 1
                
            # Phase 3: Horizontal Cylinder Matrix Resolver
            for idx, su in enumerate(sub_units):
                if idx < len(dates) and idx < len(hours):
                    d_match = re.search(r'\b(\d{1,2}[-/\.]\w{2,9}[-/\.]?\d{2,4}|\d{1,2}[-/\.]\d{1,2}[-/\.]\d{2,4})\b', dates[idx])
                    h_clean = re.sub(r'[^\d.]', '', hours[idx])
                    
                    if d_match and h_clean:
                        curr_timeline = target_timeline if system == 'ME' else f"{su}_Hours"
                        extracted_data.append({
                            'System': system,
                            'Target_Timeline': curr_timeline,
                            'Component': f"[{system}] {comp_name} ({su})",
                            'Last_Overhaul': pd.to_datetime(d_match.group(1), errors='coerce', dayfirst=True),
                            'Claimed_Hours': float(h_clean) if h_clean.replace('.','',1).isdigit() else 0.0
                        })
            i += 1 
        else:
            i += 1
            
    raw_cells_df = pd.DataFrame(cells, columns=["ASCII Extracted Text Blocks"])

    if extracted_data:
        clean_df = pd.DataFrame(extracted_data).dropna(subset=['Last_Overhaul']).reset_index(drop=True)
        return clean_df, raw_cells_df
    
    return pd.DataFrame(), raw_cells_df

# ═══════════════════════════════════════════════════════════════════════════════
# 3. MAIN FRONTEND ORCHESTRATOR
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
        <span style="color:#00e0b0">DECODER</span>&ensp;State-Machine Matrix<br>
        <span style="color:#fff">BUILD</span>&ensp;v10.0.3 Anchor Edition
    </div>
</div>
""", unsafe_allow_html=True)

col1, col2 = st.columns(2)
with col1:
    st.markdown("<div style='color:#8ba1b5; font-size:0.9rem; font-weight:600; margin-bottom:10px;'>1. COMPONENT OVERHAULS (.doc)</div>", unsafe_allow_html=True)
    pms_file = st.file_uploader("Upload Word Binary", type=["doc", "docx"], key="pms_box")

with col2:
    st.markdown("<div style='color:#8ba1b5; font-size:0.9rem; font-weight:600; margin-bottom:10px;'>2. DAILY OPERATING HOURS (Excel)</div>", unsafe_allow_html=True)
    logs_file = st.file_uploader("Upload Excel Timeline", type=["xlsx", "xls"], key="logs_box")

if pms_file and logs_file:
    with st.spinner("Executing State-Machine Extractor & Triangulation..."):
        try:
            pms_df, diag_pms = parse_pms_binary_doc(pms_file.getvalue())
            timeline_df, diag_timeline, col_map = parse_daily_hours_excel(logs_file.getvalue())
            
            total_days = len(timeline_df) if not timeline_df.empty else 0
            audit_results = []
            
            if not timeline_df.empty and not pms_df.empty:
                for _, row in pms_df.iterrows():
                    comp = row['Component']
                    oh_date = row['Last_Overhaul']
                    legacy_hrs = row['Claimed_Hours']
                    target_col = row['Target_Timeline']
                    
                    if target_col in timeline_df.columns:
                        mask = timeline_df['Date'] >= oh_date
                        verified_hrs = timeline_df.loc[mask, target_col].sum()
                        delta = verified_hrs - legacy_hrs
                        
                        audit_results.append({
                            "Component": comp,
                            "Overhaul Date": oh_date.strftime('%d-%b-%Y'),
                            "Claimed (Doc)": int(legacy_hrs),
                            "Verified (Excel)": int(verified_hrs),
                            "Delta (Drift)": int(delta),
                            "Status": "VERIFIED" if int(delta) == 0 else "DRIFT DETECTED"
                        })
                    else:
                        audit_results.append({
                            "Component": comp,
                            "Overhaul Date": oh_date.strftime('%d-%b-%Y'),
                            "Claimed (Doc)": int(legacy_hrs),
                            "Verified (Excel)": 0,
                            "Delta (Drift)": 0,
                            "Status": "NO TIMELINE DATA"
                        })

            # ═══════════════════════════════════════════════════════════════════════════════
            # 4. DASHBOARD RENDERING
            # ═══════════════════════════════════════════════════════════════════════════════
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
                        if row['Status'] == 'DRIFT DETECTED': return ['background-color: rgba(255, 42, 85, 0.1); color: #ff8a9f'] * len(row)
                        elif row['Status'] == 'NO TIMELINE DATA': return ['background-color: rgba(201, 168, 76, 0.1); color: #c9a84c'] * len(row)
                        return ['color: #00e0b0'] * len(row)
                    
                    st.dataframe(res_df.style.apply(style_dataframe, axis=1), use_container_width=True, hide_index=True)
                    
                    csv_data = res_df.to_csv(index=False).encode('utf-8')
                    st.download_button("⬇️ EXPORT FORENSIC LEDGER (.CSV)", data=csv_data, file_name=f"Reconciliation_Audit.csv", mime='text/csv')
                else:
                    st.error("Audit Could Not Complete. Check the Diagnostics tab to ensure the files contain the correct data.")

            with t2:
                if audit_results:
                    st.markdown("### Claimed vs Verified Running Hours")
                    st.markdown("<span style='color:#64748b; font-size:0.85rem;'>Native Streamlit rendering (Zero external chart dependencies)</span><br><br>", unsafe_allow_html=True)
                    plot_df = res_df[['Component', 'Claimed (Doc)', 'Verified (Excel)']].set_index('Component')
                    st.bar_chart(plot_df, color=["#c9a84c", "#00e0b0"], height=500)

            with t3:
                st.markdown("### The Transparency Engine")
                
                c1, c2 = st.columns(2)
                with c1:
                    st.markdown("<div style='color:#8ba1b5; font-size:0.8rem; font-weight:600; margin-bottom:10px;'>RAW ASCII CELLS (.doc)</div>", unsafe_allow_html=True)
                    if diag_pms is not None and not diag_pms.empty:
                        st.dataframe(diag_pms, use_container_width=True, height=500)
                    else:
                        st.warning("No data extracted from the Overhaul file.")
                        
                with c2:
                    st.markdown("<div style='color:#8ba1b5; font-size:0.8rem; font-weight:600; margin-bottom:10px;'>RAW TIMELINE MATRIX (.xlsx)</div>", unsafe_allow_html=True)
                    if diag_timeline is not None and not diag_timeline.empty:
                        # NEW: Display exactly how the engine mapped the columns
                        st.code(f"Engine Column Map Generated:\n{col_map}", language="json")
                        st.dataframe(diag_timeline.head(50), use_container_width=True, height=450)
                    else:
                        st.warning("No data extracted from the Timeline file.")

        except Exception as e:
            st.error(f"🚨 Pipeline Execution Halted: {str(e)}")
            st.info(traceback.format_exc())
