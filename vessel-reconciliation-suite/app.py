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
import plotly.express as px

warnings.filterwarnings("ignore")

# ═══════════════════════════════════════════════════════════════════════════════
# 1. PAGE CONFIGURATION & ENTERPRISE CSS
# ═══════════════════════════════════════════════════════════════════════════════
st.set_page_config(page_title="PMS Auditor", page_icon="⚓", layout="wide", initial_sidebar_state="collapsed")

def _u(s):
    return f"data:image/svg+xml;base64,{base64.b64encode(s.encode()).decode()}"

LOGO_SVG = base64.b64encode(b'<svg viewBox="0 0 48 48" xmlns="http://www.w3.org/2000/svg"><defs><linearGradient id="pg" x1="0" y1="0" x2="1" y2="1"><stop offset="0%" stop-color="#c9a84c"/><stop offset="50%" stop-color="#00e0b0"/><stop offset="100%" stop-color="#ffffff"/></linearGradient></defs><circle cx="24" cy="24" r="22" fill="none" stroke="url(#pg)" stroke-width="1" opacity="0.35"/><path d="M24 6 L24 42" stroke="url(#pg)" stroke-width="1.6" stroke-linecap="round"/><path d="M12 24 Q24 32 36 24" fill="none" stroke="url(#pg)" stroke-width="1.6" stroke-linecap="round"/></svg>').decode()

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;800&family=JetBrains+Mono:wght@400;700&display=swap');
    :root {
        --bg:#071018; --panel:#0d1620; --line:rgba(255,255,255,0.08); --text:#f8fafc;
        --muted:#94a3b8; --teal:#00e0b0; --gold:#c9a84c; --red:#ff2a55; --violet:#7b68ee;
    }
    .stApp { background-color: var(--bg); font-family: 'Inter', sans-serif; color: var(--text); }
    #MainMenu, footer, header {visibility: hidden;}
    
    .hero { display:flex; justify-content:space-between; align-items:flex-start; gap:24px; padding:24px 28px; margin-bottom:22px; background:linear-gradient(180deg, rgba(255,255,255,0.02), rgba(255,255,255,0.01)); border:1px solid var(--line); border-radius:24px; box-shadow: 0 20px 60px rgba(0,0,0,0.35); }
    .hero-left { display: flex; align-items: center; gap: 20px; }
    .hero-logo { width: 55px; height: 55px; }
    .hero-title { font-size: 2.2rem; font-weight: 800; background: linear-gradient(90deg, #ffffff, #8ba1b5); -webkit-background-clip: text; -webkit-text-fill-color: transparent; letter-spacing: -1px; line-height: 1.1;}
    .hero-sub { font-size: 0.85rem; color: var(--muted); font-weight: 600; text-transform: uppercase; letter-spacing: 2px; margin-top: 6px; }
    .hero-badge { text-align: right; font-family: 'JetBrains Mono', monospace; font-size: 0.75rem; color: var(--muted); line-height: 1.6; }

    .stFileUploader > div > div { background-color: var(--panel) !important; border: 1px dashed rgba(255,255,255,0.1) !important; border-radius: 12px !important; transition: all 0.3s ease; }
    .stFileUploader > div > div:hover { border-color: var(--teal) !important; background-color: rgba(0, 224, 176, 0.05) !important; }

    .hud-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(220px, 1fr)); gap: 14px; margin-bottom: 22px; }
    .hud-card { background: var(--panel); border: 1px solid var(--line); border-radius: 18px; padding: 16px; box-shadow: 0 10px 30px -10px rgba(0,0,0,0.5); }
    .hud-header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 10px; }
    .hud-title { font-size: 0.75rem; color: var(--muted); text-transform: uppercase; letter-spacing: 1.2px; font-weight: 600; }
    .hud-val { font-size: 2rem; font-weight: 800; color: #ffffff; line-height: 1.1; font-family: 'JetBrains Mono', monospace; }

    .stTabs [data-baseweb="tab-list"] { gap: 24px; }
    .stTabs [data-baseweb="tab"] { height: 50px; white-space: pre-wrap; background-color: transparent; border-radius: 4px; color: var(--muted); font-weight: 600; }
    .stTabs [aria-selected="true"] { color: var(--teal); border-bottom: 2px solid var(--teal); }
</style>
""", unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════════════════════
# 2. SEMANTIC HEURISTIC PARSERS
# ═══════════════════════════════════════════════════════════════════════════════

def normalize_text(s):
    return re.sub(r"\s+", " ", str(s).upper().strip().replace("\xa0", " "))

def extract_first_float(value):
    if pd.isna(value): return np.nan
    m = re.search(r"-?\d+(?:\.\d+)?", str(value).replace(",", ""))
    return float(m.group()) if m else np.nan

def extract_report_date(filename):
    m = re.search(r'(JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC)[a-z]*[\s\.\-\_]*(\d{4})', filename, re.IGNORECASE)
    if m:
        try: return pd.to_datetime(f"01 {m.group(1)} {m.group(2)}")
        except: pass
    return pd.Timestamp.now()

@st.cache_data(show_spinner=False)
def parse_master_pms_excel(file_bytes):
    try:
        xls = pd.ExcelFile(io.BytesIO(file_bytes), engine="openpyxl")
        target_sheet = next((s for s in xls.sheet_names if "PMS" in s.upper() or "PARTS LOG" in s.upper()), xls.sheet_names[0])
        df_raw = pd.read_excel(io.BytesIO(file_bytes), sheet_name=target_sheet, header=None, dtype=object, engine="openpyxl")
    except Exception: return [], {}

    # Extract Vessel Pulse (YTD and Max Monthly Hours)
    engine_stats = {"ME_YTD": 0.0, "DG_YTD": 0.0, "ME_Pulse_Daily": 16.5, "DG_Pulse_Daily": 10.0}
    try:
        for i in range(min(25, len(df_raw))):
            row_joined = " | ".join([str(x).upper() for x in df_raw.iloc[i].values if pd.notna(x)])
            if any(k in row_joined for k in ["JAN", "FEB", "MONTHLY", "UP-TO-DATE"]):
                target_row = i + 1 if ("JAN" in row_joined and df_raw.iloc[i, 9] == "Jan.") else i
                if target_row < len(df_raw):
                    monthly_vals = [extract_first_float(df_raw.iloc[target_row, c]) for c in range(9, 21) if c < len(df_raw.columns) and pd.notna(extract_first_float(df_raw.iloc[target_row, c]))]
                    if monthly_vals:
                        sys = "ME" if ("MAIN ENGINE" in str(df_raw.iloc[max(0, target_row-2):target_row].values).upper()) else "ME"
                        engine_stats[f"{sys}_YTD"] = sum(monthly_vals)
                        # Assume the highest logged month represents 30 days of standard steaming tempo
                        engine_stats[f"{sys}_Pulse_Daily"] = max(monthly_vals) / 30.0 
    except: pass

    # Semantic Header Mapping
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

    return excel_records, engine_stats

@st.cache_data(show_spinner=False)
def parse_pms_binary_doc(file_bytes, file_name=""):
    def parse_docx_xml(bytes_):
        try:
            with ZipFile(io.BytesIO(bytes_)) as z: root = ET.fromstring(z.read("word/document.xml"))
            ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
            return [" ".join([t.text for t in p.findall(".//w:t", ns) if t.text]).strip() for p in root.findall(".//w:p", ns) if p.text or len(p.findall(".//w:t", ns))>0]
        except: return []

    cells = parse_docx_xml(file_bytes) if file_name.lower().endswith(".docx") else [c.strip() for c in file_bytes.decode("latin-1", errors="ignore").replace("\x00", "").split("\x07") if c.strip()]
    extracted_data = []
    report_date = extract_report_date(file_name)

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
                        extracted_data.append({'System': system, 'Unit': su, 'BaseComponent': cell, 'Component': f"[{system}] {cell} ({su})", 'Last_Overhaul': d, 'Claimed_Hours': float(h), 'Report_Date': report_date, 'Source_File': file_name})
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
# 3. KINEMATIC ORACLE (RATE-OF-CHANGE ENGINE)
# ═══════════════════════════════════════════════════════════════════════════════

def run_kinematic_oracle(primary_df, historical_df, excel_records, engine_stats):
    results = []
    
    # 1. Establish the Timeline & Engine Limits
    now = pd.Timestamp.now()
    has_history = not historical_df.empty
    
    for _, row in primary_df.iterrows():
        comp = row['Component']
        claimed = row['Claimed_Hours']
        oh_date = row['Last_Overhaul']
        sys_key = "ME" if "ME" in str(row['System']).upper() else "DG"
        engine_pulse_daily = engine_stats.get(f"{sys_key}_Pulse_Daily", 16.5)
        
        # Cross-reference with Excel Master
        verified = get_verified_hours(row, excel_records)
        excel_status = "MISSING FROM EXCEL" if verified is None else ("VERIFIED" if int(verified) == int(claimed) else "DRIFT DETECTED")
        delta_excel = (verified - claimed) if verified is not None else 0
        
        # Core Kinematic Variables
        delta_h = 0.0
        delta_days = 0.0
        burn_rate = 0.0
        history_match = False
        
        if has_history:
            # Look for this component in the previous month
            hist_match = historical_df[historical_df['Component'] == comp].sort_values('Report_Date', ascending=False)
            if not hist_match.empty:
                prev_row = hist_match.iloc[0]
                delta_h = claimed - prev_row['Claimed_Hours']
                delta_days = max(1, (row['Report_Date'] - prev_row['Report_Date']).days)
                burn_rate = delta_h / delta_days
                history_match = True

        if not history_match:
            # Fallback: Calculate average burn rate since overhaul
            delta_h = claimed
            delta_days = max(1, (row['Report_Date'] - oh_date).days)
            burn_rate = delta_h / delta_days

        # Physics Gates (The Traps)
        anomaly = "NONE"
        phase = "SAFE"
        
        if excel_status == "MISSING FROM EXCEL":
            anomaly, phase = "MISSING MASTER RECORD", "QUARANTINE"
        elif delta_h < 0:
            anomaly, phase = f"TIME REVERSAL (Negative Burn: {delta_h}h)", "QUARANTINE"
        elif burn_rate > 24.0:
            anomaly, phase = f"TIME WARP (>24h/day Burn Rate: {burn_rate:.1f}h/d)", "QUARANTINE"
        elif burn_rate > (engine_pulse_daily * 1.3): # Allowing 30% margin for logging drift
            anomaly, phase = f"KINEMATIC VIOLATION (Burn {burn_rate:.1f} > Engine {engine_pulse_daily:.1f})", "QUARANTINE"
        elif history_match and delta_h == 0 and engine_pulse_daily > 0:
            anomaly, phase = "COPY-PASTE FRAUD (Zero burn logged despite engine activity)", "QUARANTINE"
            
        results.append({
            "Component": comp,
            "Last Overhaul": oh_date.strftime('%d-%b-%Y') if pd.notna(oh_date) else "N/A",
            "Claimed Hrs": int(claimed),
            "Master Hrs": int(verified) if verified is not None else 0,
            "Drift": int(delta_excel),
            "Burn Rate (h/d)": round(burn_rate, 1),
            "Engine Pulse (h/d)": round(engine_pulse_daily, 1),
            "Status": excel_status,
            "Anomaly": anomaly,
            "Phase": phase
        })

    return results

# ═══════════════════════════════════════════════════════════════════════════════
# 4. ORCHESTRATOR & UI
# ═══════════════════════════════════════════════════════════════════════════════

st.markdown(f"""
<div class="hero">
    <div class="hero-left">
        <img src="data:image/svg+xml;base64,{LOGO_SVG}" class="hero-logo" alt=""/>
        <div>
            <div class="hero-title">PMS AUDITOR</div>
            <div class="hero-sub">Kinematic Fraud Detection Engine</div>
        </div>
    </div>
    <div class="hero-badge">
        <span style="color:var(--teal)">KERNEL</span>&ensp;Triple-Lock Cross-Check<br>
        <span style="color:var(--gold)">ORACLE</span>&ensp;Kinematic Rate-of-Change<br>
        <span style="color:#ffffff">BUILD</span>&ensp;v13.0.1 Zenith Master
    </div>
</div>
""", unsafe_allow_html=True)

col1, col2 = st.columns(2)
with col1:
    st.markdown("<div style='color:var(--muted); font-size:0.9rem; font-weight:600; margin-bottom:10px;'>1. COMPONENT OVERHAULS (.doc)</div>", unsafe_allow_html=True)
    pms_files = st.file_uploader("Upload Word Reports (Current + Historical)", type=["doc", "docx"], key="pms_box", label_visibility="collapsed", accept_multiple_files=True)
with col2:
    st.markdown("<div style='color:var(--muted); font-size:0.9rem; font-weight:600; margin-bottom:10px;'>2. EXCEL MASTER LOG (PMS)</div>", unsafe_allow_html=True)
    logs_file = st.file_uploader("Upload Excel PMS Log", type=["xlsx", "xls"], key="logs_box", label_visibility="collapsed")

if pms_files and logs_file:
    with st.spinner("Executing Kinematic Cross-Reference & Burn Rate Analysis..."):
        try:
            # 1. Parse all Word Documents
            all_docs_frames = []
            for file in pms_files:
                df = parse_pms_binary_doc(file.getvalue(), file.name)
                if not df.empty: all_docs_frames.append(df)
            
            if not all_docs_frames:
                st.error("No valid data extracted from Word documents.")
                st.stop()
                
            all_docs_df = pd.concat(all_docs_frames, ignore_index=True)
            
            # Sort chronologically to find the "Current" (Primary) document
            all_docs_df.sort_values('Report_Date', ascending=False, inplace=True)
            latest_date = all_docs_df['Report_Date'].iloc[0]
            primary_doc_df = all_docs_df[all_docs_df['Report_Date'] == latest_date]
            historical_doc_df = all_docs_df[all_docs_df['Report_Date'] < latest_date]

            # 2. Parse Excel Master
            excel_records, engine_stats = parse_master_pms_excel(logs_file.getvalue())

            # 3. Run Kinematic Oracle
            audit_results = run_kinematic_oracle(primary_doc_df, historical_doc_df, excel_records, engine_stats)
            res_df = pd.DataFrame(audit_results)

            st.markdown("<br>", unsafe_allow_html=True)
            t1, t2, t3 = st.tabs(["📊 IMMUTABLE LEDGER", "👁️ KINEMATIC ORACLE (PHYSICS)", "🔎 CHRONOLOGICAL X-RAY"])

            with t1:
                errors_corrected = len(res_df[res_df['Status'] == 'DRIFT DETECTED'])
                quarantined = len(res_df[res_df['Phase'] == 'QUARANTINE'])
                digital_seal = hashlib.sha256(res_df.to_json(orient='records').encode()).hexdigest()

                st.markdown(f"""
                <div class="hud-grid">
                    <div class="hud-card" style="border-bottom: 3px solid var(--teal);">
                        <div class="hud-header"><div class="hud-title">Components Audited</div></div>
                        <div class="hud-val">{len(res_df)}</div>
                    </div>
                    <div class="hud-card" style="border-bottom: 3px solid {'var(--gold)' if errors_corrected > 0 else 'var(--teal)'};">
                        <div class="hud-header"><div class="hud-title">Mathematical Drift</div></div>
                        <div class="hud-val" style="color: {'var(--gold)' if errors_corrected > 0 else 'var(--teal)'};">{errors_corrected}</div>
                    </div>
                    <div class="hud-card" style="border-bottom: 3px solid {'var(--red)' if quarantined > 0 else 'var(--teal)'};">
                        <div class="hud-header"><div class="hud-title">Quarantined Physics</div></div>
                        <div class="hud-val" style="color: {'var(--red)' if quarantined > 0 else 'var(--teal)'};">{quarantined}</div>
                    </div>
                </div>
                """, unsafe_allow_html=True)
                
                # Safe Styler Architecture (Eliminates KeyError)
                hide_cols = ['Phase', 'Anomaly', 'Engine Pulse (h/d)']
                display_cols = [c for c in res_df.columns if c not in hide_cols]
                display_df = res_df[display_cols]
                
                def apply_row_styles(df_view):
                    style_df = pd.DataFrame('', index=df_view.index, columns=df_view.columns)
                    for idx in df_view.index:
                        phase = res_df.loc[idx, 'Phase']
                        status = res_df.loc[idx, 'Status']
                        if phase == 'QUARANTINE':
                            style_df.loc[idx, :] = 'background-color: rgba(255, 42, 85, 0.1); color: #ff8a9f'
                        elif status == 'DRIFT DETECTED':
                            style_df.loc[idx, :] = 'color: #c9a84c'
                        else:
                            style_df.loc[idx, :] = 'color: #00e0b0'
                    return style_df
                
                st.dataframe(display_df.style.apply(apply_row_styles, axis=None), use_container_width=True, hide_index=True)

            with t2:
                # --- VISUAL 1: THE BURN RATE IMPOSSIBILITY MATRIX ---
                st.markdown("<h4 style='color:var(--text); margin-bottom:20px;'>1. Burn Rate Impossibility Matrix</h4>", unsafe_allow_html=True)
                
                plot_df = res_df[res_df['Status'] != 'MISSING FROM EXCEL'].copy()
                if not plot_df.empty:
                    # Color coding logic
                    colors = []
                    for _, r in plot_df.iterrows():
                        if r['Burn Rate (h/d)'] > 24: colors.append('#ff2a55') # Impossible
                        elif r['Burn Rate (h/d)'] > (r['Engine Pulse (h/d)'] * 1.3): colors.append('#c9a84c') # Suspicious
                        else: colors.append('#00e0b0') # Verified
                    plot_df['Color'] = colors
                    
                    # Sort by Burn Rate descending for visual impact
                    plot_df = plot_df.sort_values('Burn Rate (h/d)', ascending=True).tail(20) # Show top 20 anomalies
                    
                    fig = go.Figure()
                    fig.add_trace(go.Bar(
                        y=plot_df['Component'], x=plot_df['Burn Rate (h/d)'], orientation='h',
                        marker_color=plot_df['Color'],
                        text=plot_df['Burn Rate (h/d)'].astype(str) + " h/d", textposition='outside'
                    ))
                    
                    # Engine Pulse Gold Line
                    avg_pulse = plot_df['Engine Pulse (h/d)'].mean()
                    fig.add_vline(x=avg_pulse, line_width=3, line_dash="dash", line_color="#c9a84c", annotation_text=f"Engine Pulse ({avg_pulse:.1f}h)", annotation_position="top right")
                    
                    # Physics Limit Red Line
                    fig.add_vline(x=24.0, line_width=3, line_dash="solid", line_color="#ff2a55", annotation_text="Physics Limit (24h)", annotation_position="top right")
                    
                    fig.update_layout(
                        plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)', font=dict(color='#94a3b8'),
                        xaxis=dict(title="Accumulated Hours Per Day", gridcolor='rgba(255,255,255,0.05)'),
                        height=600, margin=dict(l=20, r=40, t=40, b=20)
                    )
                    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})

                # --- VISUAL 2: THE THREAT MATRIX ---
                st.markdown("<h4 style='color:var(--gold); margin-top:30px;'>2. Kinematic Threat Matrix</h4>", unsafe_allow_html=True)
                threats = res_df[res_df['Phase'] == 'QUARANTINE'][['Component', 'Claimed Hrs', 'Burn Rate (h/d)', 'Anomaly']]
                if not threats.empty:
                    st.dataframe(threats, use_container_width=True, hide_index=True)
                else:
                    st.success("Zero Kinematic or Physics violations detected. Data aligns with the Engine Pulse.")

            with t3:
                c1, c2 = st.columns(2)
                with c1:
                    st.markdown(f"<div style='color:var(--muted); font-size:0.8rem; font-weight:600; margin-bottom:10px;'>CHRONOLOGICAL REPORTS ({len(pms_files)} Uploaded)</div>", unsafe_allow_html=True)
                    st.dataframe(all_docs_df[['Source_File', 'Report_Date', 'Component', 'Claimed_Hours']], use_container_width=True, height=500)
                with c2:
                    st.markdown("<div style='color:var(--muted); font-size:0.8rem; font-weight:600; margin-bottom:10px;'>EXCEL MASTER ENGINE STATS</div>", unsafe_allow_html=True)
                    st.json(engine_stats)
                    st.markdown("<div style='color:var(--muted); font-size:0.8rem; font-weight:600; margin-top:20px; margin-bottom:10px;'>RAW MASTER LOGS</div>", unsafe_allow_html=True)
                    st.dataframe(pd.DataFrame(excel_records), use_container_width=True, height=350)

        except Exception as e:
            st.error(f"🚨 Pipeline Execution Halted: {str(e)}")
            st.info(traceback.format_exc())
