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

warnings.filterwarnings("ignore")

# ═══════════════════════════════════════════════════════════════════════════════
# 1. PREMIUM FRONTEND & SVG ENGINE (NO EMOJIS)
# ═══════════════════════════════════════════════════════════════════════════════
st.set_page_config(page_title="PMS Auditor", layout="wide", initial_sidebar_state="collapsed")

def _svg(svg_string):
    return f"data:image/svg+xml;base64,{base64.b64encode(svg_string.encode()).decode()}"

# High-Fidelity SVG Assets
SVG_LOGO = _svg('<svg viewBox="0 0 48 48" xmlns="http://www.w3.org/2000/svg"><defs><linearGradient id="g1" x1="0" y1="0" x2="1" y2="1"><stop offset="0%" stop-color="#00e0b0"/><stop offset="100%" stop-color="#0284c7"/></linearGradient></defs><circle cx="24" cy="24" r="20" fill="none" stroke="url(#g1)" stroke-width="1.5" opacity="0.3"/><path d="M24 10 L24 38" stroke="url(#g1)" stroke-width="2.5" stroke-linecap="round"/><path d="M14 24 L34 24" stroke="url(#g1)" stroke-width="2.5" stroke-linecap="round"/><circle cx="24" cy="24" r="6" fill="#030712" stroke="url(#g1)" stroke-width="2"/></svg>')
SVG_SHIELD = _svg('<svg viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg"><path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z" fill="none" stroke="#00e0b0" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"/><path d="M9 12l2 2 4-4" fill="none" stroke="#00e0b0" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"/></svg>')
SVG_ALERT = _svg('<svg viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg"><path d="M10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0zM12 9v4M12 17h.01" fill="none" stroke="#ff2a55" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"/></svg>')
SVG_MATH = _svg('<svg viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg"><path d="M4 7h16M4 17h16M14 4l-4 16" fill="none" stroke="#c9a84c" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"/></svg>')
SVG_LEDGER = _svg('<svg viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg"><rect x="3" y="4" width="18" height="16" rx="2" ry="2" fill="none" stroke="#64748b" stroke-width="1.5"/><line x1="3" y1="10" x2="21" y2="10" stroke="#64748b" stroke-width="1.5"/><line x1="8" y1="10" x2="8" y2="20" stroke="#64748b" stroke-width="1.5"/></svg>')

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;600;800&family=JetBrains+Mono:wght@400;700&display=swap');
    
    :root {
        --bg: #030712; --panel: rgba(17, 24, 39, 0.6); --border: rgba(255, 255, 255, 0.08); 
        --text: #f8fafc; --muted: #64748b; 
        --cyan: #00e0b0; --red: #ff2a55; --gold: #c9a84c; --blue: #0284c7;
    }
    
    .stApp { background-color: var(--bg); font-family: 'Inter', sans-serif; color: var(--text); }
    #MainMenu, footer, header {visibility: hidden;}

    /* Animations */
    @keyframes pulse-glow { 0% { box-shadow: 0 0 0 0 rgba(0, 224, 176, 0.2); } 70% { box-shadow: 0 0 20px 10px rgba(0, 224, 176, 0); } 100% { box-shadow: 0 0 0 0 rgba(0, 224, 176, 0); } }
    @keyframes slide-up { from { opacity: 0; transform: translateY(10px); } to { opacity: 1; transform: translateY(0); } }

    /* High-End Hero Section */
    .hero {
        display: flex; justify-content: space-between; align-items: center;
        background: radial-gradient(circle at top right, rgba(0, 224, 176, 0.05), transparent 40%),
                    linear-gradient(180deg, rgba(255,255,255,0.03), rgba(255,255,255,0.01));
        border: 1px solid var(--border); border-radius: 16px;
        padding: 32px 40px; margin-bottom: 30px;
        backdrop-filter: blur(20px); box-shadow: 0 10px 30px rgba(0,0,0,0.5);
        animation: slide-up 0.6s ease-out forwards;
    }
    .hero-title { font-size: 2.2rem; font-weight: 800; letter-spacing: -0.04em; background: linear-gradient(90deg, #fff, #94a3b8); -webkit-background-clip: text; -webkit-text-fill-color: transparent; }
    .hero-sub { font-size: 0.8rem; color: var(--cyan); text-transform: uppercase; letter-spacing: 0.2em; font-weight: 600; margin-top: 4px; }
    
    /* HUD Grid */
    .hud-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(240px, 1fr)); gap: 16px; margin-bottom: 30px; animation: slide-up 0.8s ease-out forwards; }
    .hud-card {
        background: var(--panel); border: 1px solid var(--border); border-radius: 12px; padding: 20px;
        backdrop-filter: blur(10px); transition: all 0.3s ease;
    }
    .hud-card:hover { transform: translateY(-3px); border-color: rgba(255,255,255,0.15); box-shadow: 0 10px 20px rgba(0,0,0,0.3); }
    .hud-header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 12px; }
    .hud-title { font-size: 0.75rem; color: var(--muted); text-transform: uppercase; letter-spacing: 0.1em; font-weight: 600; }
    .hud-icon { width: 24px; height: 24px; }
    .hud-val { font-family: 'JetBrains Mono', monospace; font-size: 2.2rem; font-weight: 700; color: #fff; line-height: 1; }
    
    /* Custom Tabs */
    .stTabs [data-baseweb="tab-list"] { gap: 40px; border-bottom: 1px solid var(--border); }
    .stTabs [data-baseweb="tab"] { height: 60px; font-weight: 600; color: var(--muted); text-transform: uppercase; letter-spacing: 0.05em; font-size: 0.85rem; background: transparent; transition: color 0.3s; }
    .stTabs [aria-selected="true"] { color: var(--cyan); border-bottom: 2px solid var(--cyan); }
    
    /* Uploader Overrides */
    .stFileUploader > div > div { background-color: rgba(17,24,39,0.4) !important; border: 1px dashed rgba(255,255,255,0.1) !important; border-radius: 12px !important; }
    .stFileUploader > div > div:hover { border-color: var(--cyan) !important; background-color: rgba(0, 224, 176, 0.02) !important; }
</style>
""", unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════════════════════
# 2. THE UNBREAKABLE IMMUTABLE KERNEL (REVERTED TO PROVEN LOGIC)
# ═══════════════════════════════════════════════════════════════════════════════

def normalize_text(s): return re.sub(r"\s+", " ", str(s).upper().strip().replace("\xa0", " "))

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
def parse_pms_binary_doc(file_bytes):
    def parse_docx_xml(bytes_):
        try:
            with ZipFile(io.BytesIO(bytes_)) as z: root = ET.fromstring(z.read("word/document.xml"))
            ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
            return [" ".join([t.text for t in p.findall(".//w:t", ns) if t.text]).strip() for p in root.findall(".//w:p", ns) if p.text or len(p.findall(".//w:t", ns))>0]
        except: return []

    cells = parse_docx_xml(file_bytes)
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
                        extracted_data.append({'System': system, 'Unit': su, 'BaseComponent': cell, 'Component': f"[{system}] {cell} ({su})", 'Last_Overhaul': d, 'Claimed_Hours': float(h)})
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
# 3. THE PURE MATH ORACLE (BURN RATE VS PHYSICS)
# ═══════════════════════════════════════════════════════════════════════════════

def execute_kinematic_math(doc_df, excel_records, engine_ytd_hours):
    now = pd.Timestamp.now()
    day_of_year = max(1, now.dayofyear)
    
    # Calculate True Engine Pulse (Hours accumulated per day)
    engine_pulse = {
        "ME": engine_ytd_hours.get("ME", 0) / day_of_year,
        "DG": engine_ytd_hours.get("DG", 0) / day_of_year
    }

    results = []
    
    for _, row in doc_df.iterrows():
        comp = row['Component']
        claimed = row['Claimed_Hours']
        oh_date = row['Last_Overhaul']
        sys_type = "ME" if "ME" in str(row['System']).upper() else "DG"
        
        # 1. Verification Logic
        verified = get_verified_hours(row, excel_records)
        status = "ORPHANED DATA" if verified is None else ("VERIFIED" if int(verified) == int(claimed) else "LEDGER DESYNC")
        drift = (verified - claimed) if verified is not None else 0
        
        # 2. Pure Physics Logic
        days_alive = max(1, (now - oh_date).days)
        burn_rate = claimed / days_alive
        pulse_limit = engine_pulse.get(sys_type, 16.5)
        
        violation = "NONE"
        alert_level = 0
        
        if status == "ORPHANED DATA":
            violation = "Missing from Master Ledger"
            alert_level = 1
        elif drift < 0:
            violation = "Time Reversal (Negative Drift)"
            alert_level = 3
        elif burn_rate > 24.0:
            violation = f"Physics Breach (Burn Rate: {burn_rate:.1f} > 24h/d)"
            alert_level = 3
        elif burn_rate > (pulse_limit * 1.3): # 30% margin for logging delay
            violation = f"Kinematic Breach (Burn Rate: {burn_rate:.1f} > Engine Pulse: {pulse_limit:.1f})"
            alert_level = 2
            
        results.append({
            "Component": comp,
            "Last Overhaul": oh_date.strftime('%d-%b-%Y'),
            "Claimed (h)": int(claimed),
            "Ledger (h)": int(verified) if verified is not None else 0,
            "Drift (h)": int(drift),
            "Burn Rate (h/d)": round(burn_rate, 2),
            "Engine Pulse (h/d)": round(pulse_limit, 2),
            "Status": status,
            "Physics Violation": violation,
            "AlertLevel": alert_level # Hidden column for coloring
        })

    return pd.DataFrame(results), engine_pulse

# ═══════════════════════════════════════════════════════════════════════════════
# 4. ORCHESTRATOR & VISUALS
# ═══════════════════════════════════════════════════════════════════════════════

st.markdown(f"""
<div class="hero">
    <div class="hero-left">
        <img src="{SVG_LOGO}" class="hero-logo" alt=""/>
        <div>
            <div class="hero-title">PMS AUDITOR</div>
            <div class="hero-sub">Kinematic Mathematics Engine</div>
        </div>
    </div>
    <div class="hero-badge">
        <span style="color:var(--cyan)">KERNEL</span>&ensp;Immutable Sync<br>
        <span style="color:var(--gold)">ORACLE</span>&ensp;Pure Physics<br>
        <span style="color:#ffffff">BUILD</span>&ensp;v15.0.0
    </div>
</div>
""", unsafe_allow_html=True)

col1, col2 = st.columns(2)
with col1:
    st.markdown("<div style='color:var(--muted); font-size:0.85rem; font-weight:600; margin-bottom:10px; letter-spacing: 0.1em;'>REPORT INGESTION (.DOCX)</div>", unsafe_allow_html=True)
    doc_file = st.file_uploader("Upload Word Report", type=["doc", "docx"], label_visibility="collapsed")
with col2:
    st.markdown("<div style='color:var(--muted); font-size:0.85rem; font-weight:600; margin-bottom:10px; letter-spacing: 0.1em;'>MASTER LEDGER INGESTION (.XLSX)</div>", unsafe_allow_html=True)
    excel_file = st.file_uploader("Upload Master Database", type=["xlsx", "xls"], label_visibility="collapsed")

if doc_file and excel_file:
    with st.spinner("Compiling Mathematical Topology..."):
        try:
            # 1. Unbreakable Parsing
            doc_df = parse_pms_binary_doc(doc_file.getvalue())
            excel_records, engine_ytd_hours = parse_master_pms_excel(excel_file.getvalue())
            
            if doc_df.empty:
                st.error("Extraction Core Halted: No recognized component data found in Word document.")
                st.stop()

            # 2. Mathematical Execution
            master_df, engine_pulse = execute_kinematic_math(doc_df, excel_records, engine_ytd_hours)

            # 3. HUD Metrics
            total_audited = len(master_df)
            desync_count = len(master_df[master_df['Status'] == 'LEDGER DESYNC'])
            physics_breaches = len(master_df[master_df['AlertLevel'] >= 2])
            
            st.markdown(f"""
            <div class="hud-grid">
                <div class="hud-card" style="border-bottom: 2px solid var(--blue);">
                    <div class="hud-header">
                        <div class="hud-title">Components Audited</div>
                        <img src="{SVG_LEDGER}" class="hud-icon"/>
                    </div>
                    <div class="hud-val">{total_audited}</div>
                </div>
                <div class="hud-card" style="border-bottom: 2px solid {'var(--gold)' if desync_count > 0 else 'var(--cyan)'};">
                    <div class="hud-header">
                        <div class="hud-title">Ledger Desyncs</div>
                        <img src="{SVG_MATH}" class="hud-icon"/>
                    </div>
                    <div class="hud-val" style="color:{'var(--gold)' if desync_count > 0 else 'var(--cyan)'};">{desync_count}</div>
                </div>
                <div class="hud-card" style="border-bottom: 2px solid {'var(--red)' if physics_breaches > 0 else 'var(--cyan)'};">
                    <div class="hud-header">
                        <div class="hud-title">Physics Breaches</div>
                        <img src="{SVG_ALERT}" class="hud-icon"/>
                    </div>
                    <div class="hud-val" style="color:{'var(--red)' if physics_breaches > 0 else 'var(--text)'};">{physics_breaches}</div>
                </div>
            </div>
            """, unsafe_allow_html=True)

            t1, t2, t3 = st.tabs(["IMMUTABLE LEDGER", "KINEMATIC ORACLE", "RAW EXTRACTION"])

            with t1:
                # Safe Display Logic
                display_cols = [c for c in master_df.columns if c != 'AlertLevel']
                view_df = master_df[display_cols]
                
                def color_rows(df_view):
                    style = pd.DataFrame('', index=df_view.index, columns=df_view.columns)
                    for idx in df_view.index:
                        lvl = master_df.loc[idx, 'AlertLevel']
                        if lvl == 3: style.loc[idx, :] = 'background-color: rgba(255, 42, 85, 0.1); color: #ff8a9f;'
                        elif lvl == 2: style.loc[idx, :] = 'color: #c9a84c;'
                        elif lvl == 1: style.loc[idx, :] = 'color: #94a3b8;'
                        else: style.loc[idx, :] = 'color: #00e0b0;'
                    return style

                st.dataframe(view_df.style.apply(color_rows, axis=None), use_container_width=True, hide_index=True)

            with t2:
                st.markdown("<h4 style='color:var(--text); font-weight:600; margin-bottom:15px;'>The Kinematic Limit Horizon</h4>", unsafe_allow_html=True)
                
                plot_df = master_df[master_df['AlertLevel'] != 1].copy() # Remove orphans for graphing
                if not plot_df.empty:
                    # Sort to bring highest burn rates to top
                    plot_df = plot_df.sort_values('Burn Rate (h/d)', ascending=True).tail(25)
                    
                    colors = ["#ff2a55" if lvl == 3 else "#c9a84c" if lvl == 2 else "#00e0b0" for lvl in plot_df['AlertLevel']]
                    
                    fig = go.Figure()
                    fig.add_trace(go.Bar(
                        y=plot_df['Component'], x=plot_df['Burn Rate (h/d)'], orientation='h',
                        marker_color=colors, marker_line_color="#030712", marker_line_width=1,
                        text=plot_df['Burn Rate (h/d)'].astype(str) + " h/d", textposition='outside'
                    ))
                    
                    me_pulse = engine_pulse.get("ME", 16.5)
                    fig.add_vline(x=me_pulse, line_width=2, line_dash="dash", line_color="#c9a84c", annotation_text=f"Engine Pulse ({me_pulse:.1f}h)", annotation_font_color="#c9a84c", annotation_position="top right")
                    fig.add_vline(x=24.0, line_width=2, line_dash="solid", line_color="#ff2a55", annotation_text="Physics Limit (24h)", annotation_font_color="#ff2a55", annotation_position="top right")
                    
                    fig.update_layout(
                        plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)', font=dict(family="Inter", color='#94a3b8'),
                        xaxis=dict(title="Accumulated Burn Rate (Hours/Day)", gridcolor='rgba(255,255,255,0.05)', zerolinecolor="rgba(255,255,255,0.05)"),
                        yaxis=dict(gridcolor='rgba(0,0,0,0)'),
                        height=650, margin=dict(l=10, r=40, t=30, b=20)
                    )
                    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})

            with t3:
                c1, c2 = st.columns(2)
                with c1:
                    st.markdown("<div style='color:var(--muted); font-size:0.8rem; font-weight:600; margin-bottom:10px;'>PARSED DOC REPORT</div>", unsafe_allow_html=True)
                    st.dataframe(doc_df[['Component', 'Claimed_Hours', 'Last_Overhaul']], use_container_width=True, height=500)
                with c2:
                    st.markdown("<div style='color:var(--muted); font-size:0.8rem; font-weight:600; margin-bottom:10px;'>PARSED EXCEL LEDGER</div>", unsafe_allow_html=True)
                    st.json({"Calculated_Engine_Pulse": engine_pulse, "Raw_YTD_Hours": engine_ytd_hours})
                    st.dataframe(pd.DataFrame(excel_records)[['Component', 'ExcelHours']], use_container_width=True, height=380)

        except Exception as e:
            st.error("🚨 Critical System Halt.")
            st.code(traceback.format_exc(), language="bash")
