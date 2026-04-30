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

st.set_page_config(page_title="PMS Auditor | Enterprise Recon", layout="wide", initial_sidebar_state="collapsed")

# ═══════════════════════════════════════════════════════════════════════════════
# 1. ENTERPRISE CSS & HIGH-END SVG ENGINE
# ═══════════════════════════════════════════════════════════════════════════════

def _svg(s):
    return f"data:image/svg+xml;base64,{base64.b64encode(s.encode()).decode()}"

# High-Fidelity SVG Paths (No standard emojis allowed)
SVG_LOGO = _svg('<svg viewBox="0 0 48 48" xmlns="http://www.w3.org/2000/svg"><defs><linearGradient id="pg" x1="0" y1="0" x2="1" y2="1"><stop offset="0%" stop-color="#00e0b0"/><stop offset="100%" stop-color="#0284c7"/></linearGradient></defs><circle cx="24" cy="24" r="22" fill="none" stroke="url(#pg)" stroke-width="1.5" opacity="0.3"/><path d="M24 6L24 42" stroke="url(#pg)" stroke-width="2" stroke-linecap="round"/><path d="M10 24L38 24" stroke="url(#pg)" stroke-width="2" stroke-linecap="round"/><circle cx="24" cy="24" r="4" fill="#030712" stroke="url(#pg)" stroke-width="2"/></svg>')
SVG_SYNC = _svg('<svg viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg"><path d="M21 12c0 4.97-4.03 9-9 9s-9-4.03-9-9 4.03-9 9-9c2.12 0 4.07.74 5.61 1.97" fill="none" stroke="#00e0b0" stroke-width="2" stroke-linecap="round"/><path d="M16 4v5h-5" fill="none" stroke="#00e0b0" stroke-width="2" stroke-linecap="round"/><path d="M10 14l2 2 4-4" fill="none" stroke="#00e0b0" stroke-width="2" stroke-linecap="round"/></svg>')
SVG_ALERT = _svg('<svg viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg"><path d="M12 22C6.477 22 2 17.523 2 12S6.477 2 12 2s10 4.477 10 10-4.477 10-10 10zm-1-11v6h2v-6h-2zm0-4v2h2V7h-2z" fill="#ff2a55"/></svg>')
SVG_MATH = _svg('<svg viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg"><path d="M4 7h16M4 17h16M14 4l-4 16" fill="none" stroke="#c9a84c" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/></svg>')
SVG_SEAL = _svg('<svg viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg"><path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z" fill="none" stroke="#7b68ee" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/><circle cx="12" cy="12" r="3" fill="none" stroke="#7b68ee" stroke-width="2"/></svg>')

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;600;800&family=JetBrains+Mono:wght@400;700&display=swap');
    
    :root {
        --bg: #030712; --panel: rgba(15, 23, 42, 0.6); --border: rgba(255, 255, 255, 0.08); 
        --text: #f8fafc; --muted: #64748b; 
        --cyan: #00e0b0; --red: #ff2a55; --gold: #c9a84c; --purple: #7b68ee;
    }
    
    .stApp { background-color: var(--bg); font-family: 'Inter', sans-serif; color: var(--text); }
    #MainMenu, footer, header {visibility: hidden;}

    /* Animations */
    @keyframes slide-up { from { opacity: 0; transform: translateY(15px); } to { opacity: 1; transform: translateY(0); } }

    /* Glassmorphism Hero */
    .hero {
        display: flex; justify-content: space-between; align-items: center;
        background: radial-gradient(circle at top right, rgba(0, 224, 176, 0.05), transparent 50%),
                    linear-gradient(180deg, rgba(15,23,42,0.8), rgba(15,23,42,0.4));
        border: 1px solid var(--border); border-radius: 16px;
        padding: 32px 40px; margin-bottom: 30px;
        backdrop-filter: blur(20px); box-shadow: 0 10px 40px -10px rgba(0,0,0,0.5);
        animation: slide-up 0.5s ease-out forwards;
    }
    .hero-left { display: flex; align-items: center; gap: 24px; }
    .hero-logo { width: 56px; height: 56px; }
    .hero-title { font-size: 2.2rem; font-weight: 800; letter-spacing: -0.03em; color: #fff; line-height: 1.1;}
    .hero-sub { font-size: 0.85rem; color: var(--cyan); text-transform: uppercase; letter-spacing: 0.2em; font-weight: 600; margin-top: 4px; }
    .hero-badge { text-align: right; font-family: 'JetBrains Mono', monospace; font-size: 0.75rem; color: var(--muted); line-height: 1.6; }
    
    /* HUD Grid */
    .hud-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(220px, 1fr)); gap: 16px; margin-bottom: 30px; animation: slide-up 0.7s ease-out forwards; }
    .hud-card {
        background: var(--panel); border: 1px solid var(--border); border-radius: 12px; padding: 24px;
        backdrop-filter: blur(10px); transition: all 0.3s ease; display: flex; flex-direction: column;
    }
    .hud-card:hover { transform: translateY(-4px); border-color: rgba(255,255,255,0.2); box-shadow: 0 12px 30px -10px rgba(0,0,0,0.6); }
    .hud-header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 12px; }
    .hud-title { font-size: 0.75rem; color: var(--muted); text-transform: uppercase; letter-spacing: 0.1em; font-weight: 600; }
    .hud-icon { width: 24px; height: 24px; }
    .hud-val { font-family: 'JetBrains Mono', monospace; font-size: 2.4rem; font-weight: 800; color: #fff; line-height: 1; }
    .hud-sub { font-size: 0.75rem; color: var(--muted); margin-top: 8px; font-weight: 500; }
    
    /* Custom Tabs */
    .stTabs [data-baseweb="tab-list"] { gap: 32px; border-bottom: 1px solid var(--border); }
    .stTabs [data-baseweb="tab"] { height: 56px; font-weight: 600; color: var(--muted); text-transform: uppercase; letter-spacing: 0.05em; font-size: 0.85rem; background: transparent; transition: color 0.3s; }
    .stTabs [aria-selected="true"] { color: var(--cyan); border-bottom: 2px solid var(--cyan); }
    
    /* Uploader Overrides */
    .stFileUploader > div > div { background-color: rgba(15,23,42,0.3) !important; border: 1px dashed rgba(255,255,255,0.15) !important; border-radius: 12px !important; transition: all 0.3s ease; }
    .stFileUploader > div > div:hover { border-color: var(--cyan) !important; background-color: rgba(0, 224, 176, 0.03) !important; }
</style>
""", unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════════════════════
# 2. CORE DATA EXTRACTION (PROVEN / UNALTERED LOGIC)
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
        try: return pd.to_datetime(cleaned, format=fmt, errors="raise", dayfirst=True)
        except: pass
    return pd.NaT

@st.cache_data(show_spinner=False)
def parse_master_pms_excel(file_bytes):
    try:
        xls = pd.ExcelFile(io.BytesIO(file_bytes), engine="openpyxl")
        target_sheet = next((s for s in xls.sheet_names if "PMS" in s.upper() or "PARTS LOG" in s.upper()), xls.sheet_names[0])
        df_raw = pd.read_excel(io.BytesIO(file_bytes), sheet_name=target_sheet, header=None, dtype=object, engine="openpyxl")
    except Exception:
        try: df_raw = pd.read_excel(io.BytesIO(file_bytes), header=None, dtype=object, engine="openpyxl")
        except: return []

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

    for i in range(header_row + 1, len(df_raw)):
        item_val = df_raw.iloc[i, item_col]
        hours_val = df_raw.iloc[i, hours_col]

        if pd.isna(item_val): continue
        item_str = normalize_text(item_val)

        if "MAIN ENGINE" in item_str or "M/E" in item_str:
            curr_sys, curr_unit = "ME", ""
        elif any(x in item_str for x in ["GENERATOR NO.1", "DG 1", "D/G 1", "NO 1", "NO.1"]):
            curr_sys, curr_unit = "DG", "DG1"
        elif any(x in item_str for x in ["GENERATOR NO.2", "DG 2", "D/G 2", "NO 2", "NO.2"]):
            curr_sys, curr_unit = "DG", "DG2"
        elif any(x in item_str for x in ["GENERATOR NO.3", "DG 3", "D/G 3", "NO 3", "NO.3"]):
            curr_sys, curr_unit = "DG", "DG3"

        if "CYLINDER NO" in item_str or "CYL NO" in item_str:
            m = re.search(r'NO[\.\:\s]*(\d+)', item_str)
            if m: curr_unit = f"Cyl {m.group(1)}"

        h = extract_first_float(hours_val)
        if pd.notna(h) and len(item_str) > 2 and "HOURS" not in item_str and "DATE" not in item_str:
            excel_records.append({
                'System': curr_sys,
                'Unit': curr_unit,
                'ExcelComponent': item_str,
                'ExcelHours': h
            })

    return excel_records

@st.cache_data(show_spinner=False)
def parse_pms_binary_doc(file_bytes, file_name=""):
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
                line = " ".join([t.text for t in para.findall(".//w:t", ns) if t.text]).strip()
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
            system, sub_units = "ME", ["Cyl 1", "Cyl 2", "Cyl 3", "Cyl 4", "Cyl 5", "Cyl 6"]
        elif any(x in cell for x in ["AUX. ENGINE", "AUX ENGINE", "D/G", "DG ", "DIESEL GENERATOR"]):
            system, sub_units = "DG", ["DG1", "DG2", "DG3"]

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

    # Failsafe logic to prevent KeyError: ['Last_Overhaul']
    df = pd.DataFrame(extracted_data)
    if df.empty:
        return df, pd.DataFrame(cells, columns=["Raw Doc Cells"])
        
    return df.dropna(subset=['Last_Overhaul']).reset_index(drop=True), pd.DataFrame(cells, columns=["Raw Doc Cells"])

def get_verified_hours(doc_row, excel_records):
    doc_sys = doc_row['System']
    doc_unit = doc_row['Unit']
    w_comp = re.sub(r'[^A-Z]', '', doc_row['BaseComponent'].replace("ASSY", "ASSEMBLY"))
    
    best_score = 0
    best_hours = 0
    
    for er in excel_records:
        if er['System'] != doc_sys: continue
        if er['Unit'] != doc_unit and er['Unit'] != "": continue
            
        e_comp = re.sub(r'[^A-Z]', '', er['ExcelComponent'].replace("ASSY", "ASSEMBLY"))
        score = SequenceMatcher(None, w_comp, e_comp).ratio()
        
        if w_comp in e_comp or e_comp in w_comp: score += 0.3
            
        if score > best_score:
            best_score = score
            best_hours = er['ExcelHours']

    if best_score > 0.55:
        return best_hours
    return None

# ═══════════════════════════════════════════════════════════════════════════════
# 3. MAIN FRONTEND ORCHESTRATOR
# ═══════════════════════════════════════════════════════════════════════════════

st.markdown(f"""
<div class="hero">
    <div class="hero-left">
        <img src="{SVG_LOGO}" class="hero-logo" alt=""/>
        <div>
            <div class="hero-title">PMS AUDITOR</div>
            <div class="hero-sub">Pure Data Integrity Engine</div>
        </div>
    </div>
    <div class="hero-badge">
        <span style="color:var(--cyan)">KERNEL</span>&ensp;Fuzzy Data Bridge<br>
        <span style="color:var(--cyan)">SOURCE</span>&ensp;Master Parts Log<br>
        <span style="color:#fff">BUILD</span>&ensp;v15.0 Zenith
    </div>
</div>
""", unsafe_allow_html=True)

col1, col2 = st.columns(2)
with col1:
    st.markdown("<div style='color:var(--muted); font-size:0.85rem; font-weight:600; margin-bottom:8px; letter-spacing:0.1em;'>1. COMPONENT OVERHAULS (.DOCX)</div>", unsafe_allow_html=True)
    pms_file = st.file_uploader("Upload Word Report", type=["doc", "docx"], key="pms_box", label_visibility="collapsed")

with col2:
    st.markdown("<div style='color:var(--muted); font-size:0.85rem; font-weight:600; margin-bottom:8px; letter-spacing:0.1em;'>2. EXCEL MASTER LOG (.XLSX)</div>", unsafe_allow_html=True)
    logs_file = st.file_uploader("Upload Excel PMS Log", type=["xlsx", "xls"], key="logs_box", label_visibility="collapsed")

if pms_file and logs_file:
    with st.spinner("Executing Master Cross-Reference Matcher..."):
        try:
            # Extraction Phase
            doc_df, diag_pms = parse_pms_binary_doc(pms_file.getvalue(), pms_file.name)
            excel_records = parse_master_pms_excel(logs_file.getvalue())

            if doc_df.empty:
                st.error("Extraction Halted: No recognized component data found in the Word document. Please verify the template format.")
                st.stop()

            # Mathematical Analysis
            audit_results = []
            
            if excel_records:
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

            # Dashboard Rendering
            st.markdown("<br>", unsafe_allow_html=True)
            t1, t2, t3 = st.tabs(["📊 IMMUTABLE LEDGER", "⚖️ INTEGRITY ORACLE", "🔎 RAW EXTRACTION DATA"])

            res_df = pd.DataFrame(audit_results)

            with t1:
                if not res_df.empty:
                    errors_corrected = len(res_df[res_df['Status'] == 'DRIFT DETECTED'])
                    digital_seal = hashlib.sha256(res_df.to_json(orient='records').encode()).hexdigest()

                    st.markdown(f"""
                    <div class="hud-grid">
                        <div class="hud-card" style="border-bottom: 3px solid var(--cyan);">
                            <div class="hud-header">
                                <div class="hud-title">Components Audited</div>
                                <img src="{SVG_SYNC}" class="hud-icon"/>
                            </div>
                            <div class="hud-val">{len(res_df)}</div>
                            <div class="hud-sub">Mathematically Cross-Referenced</div>
                        </div>
                        <div class="hud-card" style="border-bottom: 3px solid {'var(--red)' if errors_corrected > 0 else 'var(--cyan)'};">
                            <div class="hud-header">
                                <div class="hud-title">Integrity Breaches</div>
                                <img src="{SVG_ALERT if errors_corrected > 0 else SVG_SYNC}" class="hud-icon"/>
                            </div>
                            <div class="hud-val" style="color: {'var(--red)' if errors_corrected > 0 else 'var(--cyan)'};">{errors_corrected}</div>
                            <div class="hud-sub">Mathematical Deviations Found</div>
                        </div>
                        <div class="hud-card" style="border-bottom: 3px solid var(--gold);">
                            <div class="hud-header">
                                <div class="hud-title">Master Log Sourced</div>
                                <img src="{SVG_MATH}" class="hud-icon"/>
                            </div>
                            <div class="hud-val">{len(excel_records)}</div>
                            <div class="hud-sub">Excel Database Rows Scanned</div>
                        </div>
                        <div class="hud-card" style="border-bottom: 3px solid var(--purple);">
                            <div class="hud-header">
                                <div class="hud-title">Digital Seal (SHA-256)</div>
                                <img src="{SVG_SEAL}" class="hud-icon"/>
                            </div>
                            <div class="hud-val" style="font-size: 1.4rem; margin-top: 8px;">{digital_seal[:12]}</div>
                            <div class="hud-sub">Cryptographic Data Lock</div>
                        </div>
                    </div>
                    """, unsafe_allow_html=True)
                    
                    # Safe Column Configuration to prevent Drop Errors
                    st.dataframe(
                        res_df,
                        column_config={
                            "Status": st.column_config.TextColumn("System Flag")
                        },
                        hide_index=True,
                        use_container_width=True
                    )
                else:
                    st.error("Audit Could Not Complete. Please check the 'Raw Extraction' tab.")

            with t2:
                if not res_df.empty:
                    st.markdown("<h4 style='color:var(--text); font-weight:600; margin-bottom:15px;'>Forensic Divergence Matrix</h4>", unsafe_allow_html=True)
                    st.markdown("<span style='color:var(--muted); font-size:0.85rem;'>Displays the exact mathematical drift (in hours) between the Word claim and the Master Ledger. Perfect alignment sits at 0.</span><br><br>", unsafe_allow_html=True)
                    
                    # Filter out missing records for graphing
                    plot_df = res_df[res_df['Status'] != 'MISSING FROM EXCEL'].copy()
                    
                    if not plot_df.empty:
                        plot_df = plot_df.sort_values('Delta (Drift)', ascending=True)
                        colors = ["#00e0b0" if d == 0 else "#ff2a55" for d in plot_df['Delta (Drift)']]
                        
                        fig = go.Figure()
                        fig.add_trace(go.Bar(
                            y=plot_df['Component'], x=plot_df['Delta (Drift)'], orientation='h',
                            marker_color=colors, marker_line_width=0,
                            text=plot_df['Delta (Drift)'].astype(str) + "h", textposition='outside'
                        ))
                        
                        fig.add_vline(x=0, line_width=2, line_color="#f8fafc")
                        
                        fig.update_layout(
                            plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)', font=dict(family="Inter", color='#94a3b8'),
                            xaxis=dict(title="Delta Drift (Verified Hours - Claimed Hours)", gridcolor='rgba(255,255,255,0.05)', zeroline=False),
                            yaxis=dict(gridcolor='rgba(0,0,0,0)'), height=max(400, len(plot_df)*30), margin=dict(l=10, r=40, t=20, b=20)
                        )
                        st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})

            with t3:
                c1, c2 = st.columns(2)
                with c1:
                    st.markdown("<div style='color:var(--muted); font-size:0.8rem; font-weight:600; margin-bottom:10px;'>DOC REPORT (Extracted Components)</div>", unsafe_allow_html=True)
                    if not doc_df.empty:
                        st.dataframe(doc_df[['Component', 'Claimed_Hours', 'Last_Overhaul']], use_container_width=True, height=500)
                    else:
                        st.warning("No data extracted from the Word file.")
                        
                with c2:
                    st.markdown("<div style='color:var(--muted); font-size:0.8rem; font-weight:600; margin-bottom:10px;'>EXCEL MASTER (Extracted Databases)</div>", unsafe_allow_html=True)
                    if excel_records:
                        st.dataframe(pd.DataFrame(excel_records)[['Component', 'ExcelHours']], use_container_width=True, height=500)
                    else:
                        st.warning("No Master Data extracted from the Excel file.")

        except Exception as e:
            st.error(f"🚨 Pipeline Execution Halted: {str(e)}")
            st.code(traceback.format_exc(), language="bash")
