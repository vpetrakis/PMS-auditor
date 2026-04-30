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
# 1. ENTERPRISE UI & HIGH-END CSS ANIMATIONS
# ═══════════════════════════════════════════════════════════════════════════════
st.set_page_config(page_title="PMS Auditor | Temporal Suite", layout="wide", initial_sidebar_state="collapsed")

def _svg(svg_string):
    return f"data:image/svg+xml;base64,{base64.b64encode(svg_string.encode()).decode()}"

# Premium Animated SVGs (No standard emojis)
LOGO = _svg('<svg viewBox="0 0 48 48" xmlns="http://www.w3.org/2000/svg"><defs><linearGradient id="pg" x1="0" y1="0" x2="1" y2="1"><stop offset="0%" stop-color="#7b68ee"/><stop offset="50%" stop-color="#00e0b0"/><stop offset="100%" stop-color="#ffffff"/></linearGradient></defs><circle cx="24" cy="24" r="22" fill="none" stroke="url(#pg)" stroke-width="1.5" opacity="0.4"/><path d="M24 8 L24 40" stroke="url(#pg)" stroke-width="2" stroke-linecap="round"/><path d="M10 24 Q24 34 38 24" fill="none" stroke="url(#pg)" stroke-width="2" stroke-linecap="round"/></svg>')
ICON_SHIELD = _svg('<svg viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg"><path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z" fill="none" stroke="#00e0b0" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"/></svg>')
ICON_BREACH = _svg('<svg viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg"><path d="M10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0zM12 9v4M12 17h.01" fill="none" stroke="#ff2a55" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"/></svg>')
ICON_HASH = _svg('<svg viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg"><path d="M4 9h16M4 15h16M10 3L8 21M16 3l-2 18" fill="none" stroke="#7b68ee" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"/></svg>')

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;600;800&family=JetBrains+Mono:wght@400;700&display=swap');
    
    :root {
        --bg: #030712; --panel: rgba(17, 24, 39, 0.7); --border: rgba(255,255,255,0.05); 
        --text: #f8fafc; --muted: #64748b; --accent-teal: #00e0b0; --accent-red: #ff2a55; --accent-purple: #7b68ee;
    }
    
    .stApp { background-color: var(--bg); font-family: 'Inter', sans-serif; color: var(--text); }
    #MainMenu, footer, header {visibility: hidden;}

    /* Premium Glassmorphism Hero */
    .hero {
        display: flex; justify-content: space-between; align-items: center;
        background: radial-gradient(circle at top left, rgba(123, 104, 238, 0.05), transparent 40%),
                    linear-gradient(180deg, rgba(255,255,255,0.03), rgba(255,255,255,0.01));
        border: 1px solid var(--border); border-radius: 20px;
        padding: 30px 40px; margin-bottom: 30px;
        backdrop-filter: blur(20px); box-shadow: 0 20px 40px -10px rgba(0,0,0,0.5);
    }
    .hero-title { font-size: 2.5rem; font-weight: 800; letter-spacing: -0.05em; background: linear-gradient(90deg, #fff, #94a3b8); -webkit-background-clip: text; -webkit-text-fill-color: transparent; line-height: 1.1; }
    .hero-sub { font-size: 0.85rem; color: var(--muted); text-transform: uppercase; letter-spacing: 0.2em; font-weight: 600; margin-top: 8px; }
    
    /* Animated Data HUD */
    .hud-container { display: grid; grid-template-columns: repeat(auto-fit, minmax(250px, 1fr)); gap: 20px; margin-bottom: 30px; }
    .hud-card {
        background: var(--panel); border: 1px solid var(--border); border-radius: 16px; padding: 24px;
        backdrop-filter: blur(10px); transition: transform 0.3s ease, border-color 0.3s ease;
    }
    .hud-card:hover { transform: translateY(-2px); border-color: rgba(255,255,255,0.15); }
    .hud-header { display: flex; justify-content: space-between; align-items: flex-start; margin-bottom: 15px; }
    .hud-title { font-size: 0.75rem; color: var(--muted); text-transform: uppercase; letter-spacing: 0.1em; font-weight: 600; }
    .hud-icon { width: 28px; height: 28px; }
    .hud-value { font-family: 'JetBrains Mono', monospace; font-size: 2.2rem; font-weight: 700; color: #fff; line-height: 1; }
    
    /* Tab Styling */
    .stTabs [data-baseweb="tab-list"] { gap: 30px; border-bottom: 1px solid var(--border); }
    .stTabs [data-baseweb="tab"] { height: 60px; font-weight: 600; color: var(--muted); text-transform: uppercase; letter-spacing: 0.05em; font-size: 0.85rem; background: transparent; }
    .stTabs [aria-selected="true"] { color: var(--accent-teal); border-bottom: 2px solid var(--accent-teal); }
    
    /* Dataframe overrides for seamless integration */
    [data-testid="stDataFrame"] { border-radius: 12px; overflow: hidden; border: 1px solid var(--border); }
</style>
""", unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════════════════════
# 2. SEMANTIC PARSERS (FOCUSED STRICTLY ON RUNNING HOURS)
# ═══════════════════════════════════════════════════════════════════════════════

def normalize(s): return re.sub(r"\s+", " ", str(s).upper().strip().replace("\xa0", " "))

def ext_float(val):
    if pd.isna(val): return np.nan
    m = re.search(r"-?\d+(?:\.\d+)?", str(val).replace(",", ""))
    return float(m.group()) if m else np.nan

def extract_doc_date(filename):
    m = re.search(r'(JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC)[a-z]*[\s\.\-\_]*(\d{4})', filename, re.IGNORECASE)
    if m:
        try: return pd.to_datetime(f"01 {m.group(1)} {m.group(2)}")
        except: pass
    return pd.Timestamp.now()

@st.cache_data(show_spinner=False)
def parse_excel_ledger(file_bytes):
    try:
        xls = pd.ExcelFile(io.BytesIO(file_bytes), engine="openpyxl")
        target_sheet = next((s for s in xls.sheet_names if "PMS" in s.upper() or "PARTS" in s.upper()), xls.sheet_names[0])
        df = pd.read_excel(io.BytesIO(file_bytes), sheet_name=target_sheet, header=None, dtype=object, engine="openpyxl")
    except: return [], {}

    engine_pulse = {"ME": 8760.0, "DG": 8760.0} # Fallback absolute limit
    try:
        # Scan for Engine Monthly Hours Matrix (Usually Row 14/15)
        for i in range(min(30, len(df))):
            row_str = " ".join([str(x).upper() for x in df.iloc[i].values if pd.notna(x)])
            if "JAN" in row_str and "FEB" in row_str:
                target_row = i + 1 if df.iloc[i, 9] == "Jan." else i
                if target_row < len(df):
                    monthly_sums = [ext_float(df.iloc[target_row, c]) for c in range(9, 21) if c < len(df.columns) and pd.notna(ext_float(df.iloc[target_row, c]))]
                    if monthly_sums:
                        sys = "ME" if "MAIN" in str(df.iloc[max(0, target_row-2):target_row].values).upper() else "ME"
                        engine_pulse[sys] = sum(monthly_sums)
    except: pass

    # Find Component Data
    item_col, hrs_col, head_row = -1, -1, -1
    for i in range(min(40, len(df))):
        row_str = " ".join([str(x).upper() for x in df.iloc[i].values if pd.notna(x)])
        if ("ITEM" in row_str or "DESCRIPTION" in row_str) and "HOURS" in row_str:
            head_row = i
            for j, val in enumerate(df.iloc[i].values):
                vu = str(val).upper()
                if "ITEM" in vu or "DESCRIPTION" in vu: item_col = j
                elif "CURRENT" in vu and "HOURS" in vu: hrs_col = j
            break

    records = []
    if item_col != -1 and hrs_col != -1:
        curr_sys, curr_unit = "ME", ""
        for i in range(head_row + 1, len(df)):
            item, hrs = df.iloc[i, item_col], df.iloc[i, hrs_col]
            if pd.isna(item): continue
            item_str = normalize(item)

            if "MAIN ENGINE" in item_str or "M/E" in item_str: curr_sys, curr_unit = "ME", ""
            elif any(x in item_str for x in ["DG 1", "D/G 1", "NO.1"]): curr_sys, curr_unit = "DG", "DG1"
            elif any(x in item_str for x in ["DG 2", "D/G 2", "NO.2"]): curr_sys, curr_unit = "DG", "DG2"
            
            m = re.search(r'NO[\.\:\s]*(\d+)', item_str)
            if m: curr_unit = f"Cyl {m.group(1)}"

            h = ext_float(hrs)
            if pd.notna(h) and len(item_str) > 2 and "HOURS" not in item_str and "DATE" not in item_str:
                records.append({'System': curr_sys, 'Unit': curr_unit, 'Component': item_str, 'Master_Hours': h})

    return records, engine_pulse

@st.cache_data(show_spinner=False)
def parse_doc_reports(file_bytes, filename=""):
    try:
        with ZipFile(io.BytesIO(file_bytes)) as z: root = ET.fromstring(z.read("word/document.xml"))
        ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
        cells = [" ".join([t.text for t in p.findall(".//w:t", ns) if t.text]).strip() for p in root.findall(".//w:p", ns) if p.text or len(p.findall(".//w:t", ns))>0]
    except: return pd.DataFrame()

    ext_data = []
    report_date = extract_doc_date(filename)
    KWS = ['CYLINDER COVER', 'PISTON ASSY', 'PISTON ASSEMBLY', 'STUFFING BOX', 'PISTON CROWN', 'CYLINDER LINER', 'EXHAUST VALVE', 'STARTING VALVE', 'SAFETY VALVE', 'FUEL VALVE', 'FUEL PUMP', 'CROSSHEAD BEARING', 'BOTTOM END BEARING', 'MAIN BEARING']
    sys, units = "ME", ["Cyl 1", "Cyl 2", "Cyl 3", "Cyl 4", "Cyl 5", "Cyl 6"]

    i = 0
    while i < len(cells):
        cell = normalize(cells[i])
        if "MAIN ENGINE" in cell: sys, units = "ME", ["Cyl 1", "Cyl 2", "Cyl 3", "Cyl 4", "Cyl 5", "Cyl 6"]
        elif any(x in cell for x in ["AUX", "D/G", "DIESEL"]): sys, units = "DG", ["DG1", "DG2", "DG3"]

        if any(c in cell for c in KWS) and len(cell) < 80:
            dates, hours = [], []
            for shift, target, arr in [(15, '1', dates), (35, '2', hours)]:
                j = i + 1
                while j < min(i + shift, len(cells)):
                    if str(cells[j]).strip() == target:
                        arr.extend([cells[j + 1 + k] for k in range(len(units)) if j + 1 + k < len(cells)])
                        break
                    j += 1

            for idx, u in enumerate(units):
                if idx < len(dates) and idx < len(hours):
                    d = pd.to_datetime(re.sub(r"\s+", " ", str(dates[idx]).strip()), errors="coerce", dayfirst=True)
                    h = ext_float(hours[idx])
                    if pd.notna(d) and pd.notna(h):
                        ext_data.append({'System': sys, 'Unit': u, 'Base': cell, 'Component': f"[{sys}] {cell} ({u})", 'Overhaul': d, 'Claimed_Hours': float(h), 'Report_Date': report_date, 'File': filename})
        i += 1
    return pd.DataFrame(ext_data).dropna(subset=['Overhaul']).reset_index(drop=True)

# ═══════════════════════════════════════════════════════════════════════════════
# 3. TEMPORAL CONTINUITY ORACLE (100% INTEGRITY ENGINE)
# ═══════════════════════════════════════════════════════════════════════════════

def get_master_hours(doc_comp, doc_sys, doc_unit, excel_records):
    c_clean = re.sub(r'[^A-Z]', '', doc_comp.replace("ASSY", "ASSEMBLY"))
    best_score, best_h = 0, None
    for er in excel_records:
        if er['System'] != doc_sys or (er['Unit'] != doc_unit and er['Unit'] != ""): continue
        e_clean = re.sub(r'[^A-Z]', '', er['Component'].replace("ASSY", "ASSEMBLY"))
        score = SequenceMatcher(None, c_clean, e_clean).ratio()
        if c_clean in e_clean or e_clean in c_clean: score += 0.3
        if score > 0.6 and score > best_score: best_score, best_h = score, er['Master_Hours']
    return best_h

def execute_temporal_oracle(all_docs, excel_records, engine_pulse):
    results = []
    
    # Sort documents chronologically
    all_docs.sort_values('Report_Date', ascending=False, inplace=True)
    latest_date = all_docs['Report_Date'].iloc[0]
    primary_doc = all_docs[all_docs['Report_Date'] == latest_date]
    hist_doc = all_docs[all_docs['Report_Date'] < latest_date]
    
    now = pd.Timestamp.now()

    for _, row in primary_doc.iterrows():
        comp = row['Component']
        claimed = row['Claimed_Hours']
        oh_date = row['Overhaul']
        sys = "ME" if "ME" in str(row['System']).upper() else "DG"
        
        master_hrs = get_master_hours(row['Base'], sys, row['Unit'], excel_records)
        status = "MISSING MASTER" if master_hrs is None else ("VERIFIED" if int(master_hrs) == int(claimed) else "LEDGER DESYNC")
        drift = (master_hrs - claimed) if master_hrs is not None else 0
        
        # Temporal Logic Variables
        hist_claimed = None
        delta_hrs = claimed
        days_elapsed = max(1, (latest_date - oh_date).days)
        
        if not hist_doc.empty:
            match = hist_doc[hist_doc['Component'] == comp].sort_values('Report_Date', ascending=False)
            if not match.empty:
                hist_claimed = match.iloc[0]['Claimed_Hours']
                delta_hrs = claimed - hist_claimed
                days_elapsed = max(1, (latest_date - match.iloc[0]['Report_Date']).days)

        engine_max_allowed = engine_pulse.get(sys, 8760) / 12.0 # Roughly a month's absolute limit
        
        # The Integrity Gates
        breach = "NONE"
        if status == "MISSING MASTER":
            breach = "DATA ORPHAN (Not in Master Ledger)"
        elif delta_hrs < 0:
            breach = f"TEMPORAL REVERSAL (Hours decreased by {abs(delta_hrs)}h)"
        elif delta_hrs == 0 and engine_pulse.get(sys, 0) > 0 and hist_claimed is not None:
            breach = "STATIC FRAUD (Zero hours logged despite Engine run)"
        elif claimed > (days_elapsed * 24) and hist_claimed is None:
            breach = f"PHYSICS BREACH (Claimed {claimed}h > {days_elapsed*24}h max possible)"

        results.append({
            "Component": comp,
            "Last Overhaul": oh_date.strftime('%d-%b-%Y'),
            "Historical Hrs": int(hist_claimed) if hist_claimed is not None else "N/A",
            "Claimed Hrs": int(claimed),
            "Master Hrs": int(master_hrs) if master_hrs is not None else "N/A",
            "Drift": int(drift),
            "Status": status,
            "Integrity Breach": breach
        })

    return pd.DataFrame(results)

# ═══════════════════════════════════════════════════════════════════════════════
# 4. FRONTEND ORCHESTRATOR
# ═══════════════════════════════════════════════════════════════════════════════

st.markdown(f"""
<div class="hero">
    <div class="hero-left">
        <img src="{LOGO}" class="hero-logo" alt=""/>
        <div>
            <div class="hero-title">PMS AUDITOR</div>
            <div class="hero-sub">Temporal Data Integrity Suite</div>
        </div>
    </div>
    <div class="hero-badge">
        <span style="color:var(--accent-teal)">SYSTEM</span>&ensp;Continuity Oracle<br>
        <span style="color:var(--muted)">VERSION</span>&ensp;v14.0.0 Zenith
    </div>
</div>
""", unsafe_allow_html=True)

c1, c2 = st.columns(2)
with c1:
    docs_files = st.file_uploader("Upload Word Reports (.doc / .docx) - Supports multiple for timeline analysis", accept_multiple_files=True)
with c2:
    excel_file = st.file_uploader("Upload Master Database (.xlsx / .xls)")

if docs_files and excel_file:
    with st.spinner("Executing Cryptographic Ledger Sync & Temporal Analysis..."):
        try:
            # Parse Documents
            all_dfs = [parse_doc_reports(f.getvalue(), f.name) for f in docs_files]
            all_dfs = [df for df in all_dfs if not df.empty]
            if not all_dfs:
                st.error("Extraction Failed: No valid Component Running Hours found in Word documents.")
                st.stop()
            docs_df = pd.concat(all_dfs, ignore_index=True)

            # Parse Master
            excel_records, engine_pulse = parse_excel_ledger(excel_file.getvalue())

            # Run Integrity Oracle
            results_df = execute_temporal_oracle(docs_df, excel_records, engine_pulse)

            # --- HUD METRICS ---
            total_audited = len(results_df)
            verified_count = len(results_df[results_df['Status'] == 'VERIFIED'])
            breach_count = len(results_df[results_df['Integrity Breach'] != 'NONE'])
            data_hash = hashlib.sha256(results_df.to_json().encode()).hexdigest()[:12].upper()

            st.markdown(f"""
            <div class="hud-container">
                <div class="hud-card" style="border-top: 3px solid var(--accent-purple);">
                    <div class="hud-header">
                        <div class="hud-title">Components Audited</div>
                        <img src="{ICON_HASH}" class="hud-icon"/>
                    </div>
                    <div class="hud-value">{total_audited}</div>
                </div>
                <div class="hud-card" style="border-top: 3px solid {'var(--accent-teal)' if verified_count == total_audited else 'var(--muted)'};">
                    <div class="hud-header">
                        <div class="hud-title">Ledger Synchronized</div>
                        <img src="{ICON_SHIELD}" class="hud-icon"/>
                    </div>
                    <div class="hud-value" style="color:var(--accent-teal);">{verified_count}</div>
                </div>
                <div class="hud-card" style="border-top: 3px solid {'var(--accent-red)' if breach_count > 0 else 'var(--muted)'};">
                    <div class="hud-header">
                        <div class="hud-title">Temporal Breaches</div>
                        <img src="{ICON_BREACH}" class="hud-icon"/>
                    </div>
                    <div class="hud-value" style="color:{'var(--accent-red)' if breach_count > 0 else '#fff'};">{breach_count}</div>
                </div>
                <div class="hud-card" style="border-top: 3px solid var(--muted);">
                    <div class="hud-header">
                        <div class="hud-title">Cryptographic Lock</div>
                        <img src="{ICON_HASH}" class="hud-icon"/>
                    </div>
                    <div class="hud-value" style="font-size:1.6rem; color:var(--muted); margin-top:8px;">{data_hash}</div>
                </div>
            </div>
            """, unsafe_allow_html=True)

            t1, t2, t3 = st.tabs(["MASTER LEDGER", "TEMPORAL TOPOLOGY", "DATA EXTRACTS"])

            with t1:
                # Safe Styler architecture
                def colorize_rows(df_view):
                    style = pd.DataFrame('', index=df_view.index, columns=df_view.columns)
                    for idx in df_view.index:
                        status = results_df.loc[idx, 'Status']
                        breach = results_df.loc[idx, 'Integrity Breach']
                        if breach != 'NONE':
                            style.loc[idx, :] = 'background-color: rgba(255, 42, 85, 0.08); color: #ff8a9f;'
                        elif status == 'LEDGER DESYNC':
                            style.loc[idx, :] = 'color: #c9a84c;'
                        else:
                            style.loc[idx, :] = 'color: #00e0b0;'
                    return style

                st.dataframe(results_df.style.apply(colorize_rows, axis=None), use_container_width=True, hide_index=True)

            with t2:
                # --- PREMIUM VISUAL: THE TEMPORAL ENVELOPE (AREA CHART) ---
                plot_df = results_df[results_df['Integrity Breach'] != 'DATA ORPHAN (Not in Master Ledger)'].copy()
                if not plot_df.empty:
                    # Sort for cleaner visual flow
                    plot_df = plot_df.sort_values('Claimed Hrs')
                    
                    fig = go.Figure()

                    # The Envelope (Safe Zone)
                    fig.add_trace(go.Scatter(
                        x=plot_df['Component'], y=plot_df['Master Hrs'],
                        name="Master Ledger Baseline", mode="lines",
                        line=dict(color="rgba(0, 224, 176, 0.5)", width=2),
                        fill='tozeroy', fillcolor="rgba(0, 224, 176, 0.05)"
                    ))

                    # The Claimed Data
                    fig.add_trace(go.Scatter(
                        x=plot_df['Component'], y=plot_df['Claimed Hrs'],
                        name="Reported Hours", mode="markers+lines",
                        line=dict(color="#ffffff", width=1),
                        marker=dict(
                            color=["#ff2a55" if b != "NONE" else "#00e0b0" for b in plot_df['Integrity Breach']],
                            size=[12 if b != "NONE" else 8 for b in plot_df['Integrity Breach']],
                            line=dict(color="#000", width=1)
                        ),
                        text=plot_df['Integrity Breach'],
                        hovertemplate="<b>%{x}</b><br>Claimed: %{y}h<br>Status: %{text}<extra></extra>"
                    ))

                    fig.update_layout(
                        title=dict(text="TEMPORAL CONTINUITY ENVELOPE", font=dict(color="#f8fafc", size=18)),
                        plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)',
                        font=dict(family="Inter", color="#64748b"),
                        xaxis=dict(showgrid=False, showticklabels=False, title="Audited Components (Sorted by Time)"),
                        yaxis=dict(gridcolor="rgba(255,255,255,0.05)", title="Accumulated Hours"),
                        hovermode="x unified", margin=dict(l=10, r=10, t=60, b=20),
                        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
                    )
                    st.plotly_chart(fig, use_container_width=True, config={'displayModeBar': False})

                # Threat Matrix Output
                breaches = results_df[results_df['Integrity Breach'] != 'NONE'][['Component', 'Claimed Hrs', 'Historical Hrs', 'Integrity Breach']]
                if not breaches.empty:
                    st.markdown("<h4 style='color:var(--accent-red); margin-top:20px; font-weight:600;'>Critical Integrity Breaches</h4>", unsafe_allow_html=True)
                    st.dataframe(breaches, use_container_width=True, hide_index=True)

            with t3:
                c1, c2 = st.columns(2)
                with c1:
                    st.markdown("<div style='color:var(--muted); font-size:0.8rem; font-weight:600; margin-bottom:10px;'>PARSED WORD DOCUMENTS</div>", unsafe_allow_html=True)
                    st.dataframe(docs_df[['Report_Date', 'Component', 'Claimed_Hours']], use_container_width=True, height=400)
                with c2:
                    st.markdown("<div style='color:var(--muted); font-size:0.8rem; font-weight:600; margin-bottom:10px;'>PARSED EXCEL LEDGER</div>", unsafe_allow_html=True)
                    st.dataframe(pd.DataFrame(excel_records)[['Component', 'Master_Hours']], use_container_width=True, height=400)

        except Exception as e:
            st.error(f"🚨 Fatal System Error: {str(e)}")
            st.code(traceback.format_exc(), language="bash")
