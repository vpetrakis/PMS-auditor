import streamlit as st
import pandas as pd
import numpy as np
import re
import io
import hashlib
import warnings
from datetime import datetime

try:
    from docx import Document
    HAS_DOCX = True
except ImportError:
    HAS_DOCX = False

warnings.filterwarnings("ignore")
st.set_page_config(page_title="Vessel Reconciliation Suite", page_icon="⚓", layout="wide", initial_sidebar_state="collapsed")

# ═══════════════════════════════════════════════════════════════════════════════
# CSS
# ═══════════════════════════════════════════════════════════════════════════════
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;600&family=Sora:wght@400;500;600;700;800&display=swap');
.stApp { background:#f8fafc; font-family:'Sora',sans-serif; color:#1e293b; }
#MainMenu,footer,header {visibility:hidden;}
.hero { border-bottom:2px solid #e2e8f0; padding-bottom:20px; margin-bottom:30px; }
.hero-title { font-size:2rem; font-weight:800; color:#1e293b; letter-spacing:-0.03em; }
.hero-sub { font-size:0.8rem; color:#94a3b8; font-weight:600; text-transform:uppercase; letter-spacing:0.12em; margin-top:4px; }
.kpi-grid { display:grid; grid-template-columns:repeat(auto-fit,minmax(200px,1fr)); gap:16px; margin-bottom:28px; }
.kpi-card { background:#fff; border:1px solid #e2e8f0; border-radius:12px; padding:20px; box-shadow:0 1px 4px rgba(0,0,0,0.04); border-top:4px solid #3b82f6; }
.kpi-card.red { border-top-color:#ef4444; }
.kpi-card.amber { border-top-color:#f59e0b; }
.kpi-card.green { border-top-color:#10b981; }
.kpi-card.violet { border-top-color:#8b5cf6; }
.kpi-label { font-size:0.7rem; color:#94a3b8; text-transform:uppercase; letter-spacing:0.12em; font-weight:600; margin-bottom:8px; }
.kpi-val { font-size:2rem; font-weight:800; color:#1e293b; font-family:'IBM Plex Mono',monospace; }
.kpi-sub { font-size:0.72rem; color:#cbd5e1; margin-top:6px; }
.seal-box { background:#fff; border:1px solid #e2e8f0; border-radius:12px; padding:16px 20px; font-family:'IBM Plex Mono',monospace; font-size:0.72rem; color:#64748b; margin-bottom:20px; word-break:break-all; }
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="hero">
    <div class="hero-title">⚓ VESSEL RECONCILIATION SUITE</div>
    <div class="hero-sub">Zero-Trust Forensic Running Hours Auditor</div>
</div>
""", unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════════════════════
# SAFE NUMBER — identical to Poseidon Titan's _sn()
# ═══════════════════════════════════════════════════════════════════════════════
def _sn(val):
    """Strip all non-numeric chars and return float; NaN on garbage."""
    if pd.isna(val): return np.nan
    s = str(val).strip().upper()
    if s in ['NIL','N/A','NA','XXX','NONE','UNKNOWN','BLANK','-','X','','NULL']: return np.nan
    s = re.sub(r'[^\d.\-]', '', s)
    try:
        return float(s) if s and s not in ('.', '-', '-.') else np.nan
    except ValueError:
        return np.nan

def _sn0(val):
    v = _sn(val)
    return 0.0 if np.isnan(v) else v

# ═══════════════════════════════════════════════════════════════════════════════
# SAFE DATE — handles all formats seen in maritime Excel files
# ═══════════════════════════════════════════════════════════════════════════════
def _safe_date(val):
    if val is None or (isinstance(val, float) and np.isnan(val)): return pd.NaT
    if isinstance(val, (pd.Timestamp, datetime)): return pd.Timestamp(val)
    
    s = str(val).strip()
    if not s or s.lower() in ['-','n/a','nil','none','null','']: return pd.NaT
    
    # Typo corrections (Poseidon-style)
    s = re.sub(r'20224', '2024', s)
    s = re.sub(r'20023', '2023', s)
    s = re.sub(r'20225', '2025', s)
    
    # "15 Jan 2024" or "15 Jan. 2024" → normalise
    m = re.match(r'^(\d{1,2})\s+([A-Za-z]{3,9})\.?\s+(\d{4})$', s)
    if m:
        s = f"{m.group(3)}-{m.group(2)[:3]}-{m.group(1).zfill(2)}"
    
    d = pd.to_datetime(s, errors='coerce', dayfirst=True)
    if pd.notna(d): return d
    
    # DD/MM/YYYY
    m2 = re.match(r'^(\d{1,2})[\/\-\.](\d{1,2})[\/\-\.](\d{2,4})$', s)
    if m2:
        dd, mm, yy = m2.group(1), m2.group(2), m2.group(3)
        yr = (1900 if int(yy) > 50 else 2000) + int(yy) if len(yy) == 2 else int(yy)
        d = pd.to_datetime(f"{yr}-{mm}-{dd}", errors='coerce')
        if pd.notna(d): return d
    
    return pd.NaT

# ═══════════════════════════════════════════════════════════════════════════════
# FORWARD-FILL ROW — Poseidon's pd.Series.ffill() equivalent
# Propagates merged-cell header text to the right.
# ═══════════════════════════════════════════════════════════════════════════════
def _fwdfill(series):
    last = ''
    result = []
    for v in series:
        s = str(v).strip().upper() if pd.notna(v) else ''
        if s and s not in ['NAN','NONE']:
            last = s
        result.append(last)
    return result

# ═══════════════════════════════════════════════════════════════════════════════
# COLUMN MAPPER — LOG FILES
# Mirrors Poseidon Titan's _map_columns() adapted for running-hours format.
# c1 = forward-filled top header, c2 = sub-header below it
# ═══════════════════════════════════════════════════════════════════════════════
def _map_log_columns(top_filled, bottom, n):
    cols = {}
    for j in range(n):
        c1 = str(top_filled[j]).upper().strip() if j < len(top_filled) else ''
        c2 = str(bottom[j]).upper().strip()     if j < len(bottom)     else ''
        cb = f"{c1} {c2}".strip()

        if 'date' not in cols and ('DATE' in cb or 'DAY' in cb):
            cols['date'] = j
        if 'me' not in cols and (
            'MAIN ENGINE' in cb or 'MAIN ENG' in cb or
            'M/E' in c2 or 'M.E.' in c2 or c2 == 'ME' or
            ('MAIN' in c1 and ('ENG' in c2 or c2 in ['','HRS','HOURS']))
        ):
            cols['me'] = j
        if 'dg1' not in cols and (
            'DG1' in cb or 'D/G 1' in cb or 'D.G.1' in cb or
            'DG NO.1' in cb or 'GEN 1' in cb or 'G/E 1' in cb or 'AUX 1' in cb or
            ('D/G' in c1 and ('1' == c2 or 'NO.1' in c2 or '1' in c2.split()))
        ):
            cols['dg1'] = j
        if 'dg2' not in cols and (
            'DG2' in cb or 'D/G 2' in cb or 'D.G.2' in cb or
            'DG NO.2' in cb or 'GEN 2' in cb or 'G/E 2' in cb or 'AUX 2' in cb or
            ('D/G' in c1 and ('2' == c2 or 'NO.2' in c2 or '2' in c2.split()))
        ):
            cols['dg2'] = j
        if 'dg3' not in cols and (
            'DG3' in cb or 'D/G 3' in cb or 'D.G.3' in cb or
            'DG NO.3' in cb or 'GEN 3' in cb or 'G/E 3' in cb or 'AUX 3' in cb or
            ('D/G' in c1 and ('3' == c2 or 'NO.3' in c2 or '3' in c2.split()))
        ):
            cols['dg3'] = j
    return cols

# ═══════════════════════════════════════════════════════════════════════════════
# COLUMN MAPPER — PMS FILES
# ═══════════════════════════════════════════════════════════════════════════════
def _map_pms_columns(top_filled, bottom, n):
    cols = {}
    for j in range(n):
        c1 = str(top_filled[j]).upper().strip() if j < len(top_filled) else ''
        c2 = str(bottom[j]).upper().strip()     if j < len(bottom)     else ''
        cb = f"{c1} {c2}".strip()

        if 'component' not in cols and (
            'COMPONENT' in cb or 'EQUIPMENT' in cb or 'DESCRIPTION' in cb or
            'ITEM' in cb or 'JOB' in cb or 'TASK' in cb or 'NAME' in cb
        ):
            cols['component'] = j
        if 'ohdate' not in cols and (
            'OVERHAUL' in cb or 'LAST OH' in cb or 'O/H DATE' in cb or
            'LAST O/H' in cb or 'INSP DATE' in cb or 'LAST INSP' in cb or
            'LAST INSPECTION' in cb or 'COMPLETED' in cb or
            ('DATE' in cb and 'component' in cols)
        ):
            cols['ohdate'] = j
        if 'hours' not in cols and (
            'RUNNING HOURS' in cb or 'R/H' in cb or 'HRS SINCE' in cb or
            'CURRENT HRS' in cb or 'CLAIMED' in cb or 'HOURS SINCE' in cb or
            'RUN HRS' in cb or 'CURRENT RUNNING' in cb or
            ('HOURS' in cb and j > 0 and 'ohdate' in cols)
        ):
            cols['hours'] = j
        if 'system' not in cols and (
            'SYSTEM' in cb or 'MACHINERY' in cb or 'DEPT' in cb or
            'CATEGORY' in cb or 'PARENT' in cb
        ):
            cols['system'] = j
    return cols

# ═══════════════════════════════════════════════════════════════════════════════
# INFER SYSTEM (ME vs DG1/2/3) from component name + system column
# ═══════════════════════════════════════════════════════════════════════════════
def _infer_system(name, sys_val=''):
    s = f"{str(name).upper()} {str(sys_val).upper()}"
    if any(k in s for k in ['DG3','D/G 3','GEN 3','G/E 3','AUX 3','GENERATOR 3']): return 'DG3'
    if any(k in s for k in ['DG2','D/G 2','GEN 2','G/E 2','AUX 2','GENERATOR 2']): return 'DG2'
    if any(k in s for k in ['DG1','D/G 1','GEN 1','G/E 1','AUX 1','GENERATOR 1']): return 'DG1'
    if any(k in s for k in ['DG','D/G','DIESEL GEN','GENERATOR','GENSET','AUX ENGINE']): return 'DG1'
    return 'MAIN_ENGINE'

# ═══════════════════════════════════════════════════════════════════════════════
# HEADER SCANNER — finds the row index with the key sentinel words
# Mirrors Poseidon's "for i in range(min(150, len(df_raw)))" loop
# ═══════════════════════════════════════════════════════════════════════════════
def _find_header_row(df_raw, sentinel_groups, max_scan=80):
    """
    Returns (header_idx, top_filled, bottom) or (-1, None, None).
    sentinel_groups: list of lists — every group must have at least one match.
    """
    for i in range(min(max_scan, len(df_raw))):
        vals = [str(x).upper() for x in df_raw.iloc[i].values if pd.notna(x) and str(x).strip() not in ('','NAN')]
        all_match = all(
            any(kw in v for v in vals for kw in group)
            for group in sentinel_groups
        )
        if not all_match:
            continue
        
        # Forward-fill this row (handles merged cells)
        top_filled = _fwdfill(df_raw.iloc[i])
        # Sub-header on the next row
        bottom = [str(x).upper().strip() if pd.notna(x) else ''
                  for x in (df_raw.iloc[i + 1].values if i + 1 < len(df_raw) else [])]
        return i, top_filled, bottom
    
    return -1, None, None

# ═══════════════════════════════════════════════════════════════════════════════
# PARSE LOG FILE — uses Poseidon's dual-header approach
# ═══════════════════════════════════════════════════════════════════════════════
@st.cache_data(show_spinner=False)
def parse_log_file(file_bytes, file_name):
    """Returns (daily_df, error_string_or_None)"""
    try:
        if file_name.lower().endswith(('.xlsx', '.xls')):
            wb = pd.read_excel(io.BytesIO(file_bytes), sheet_name=None, header=None, dtype=str, engine='openpyxl' if file_name.endswith('.xlsx') else 'xlrd')
        else:
            # CSV / TXT fallback
            raw = file_bytes.decode('latin-1', errors='replace')
            df_raw = pd.read_csv(io.StringIO(raw), header=None, on_bad_lines='skip', dtype=str)
            wb = {'Sheet1': df_raw}
        
        for sheet_name, df_raw in wb.items():
            if df_raw is None or len(df_raw) < 4:
                continue
            
            header_idx, top_filled, bottom = _find_header_row(
                df_raw,
                sentinel_groups=[
                    ['DATE', 'DAY'],
                    ['MAIN', 'M/E', 'M.E.', 'MAIN ENGINE', 'ENGINE'],
                ]
            )
            if header_idx == -1:
                continue
            
            n = len(df_raw.columns)
            bottom_padded = bottom + [''] * (n - len(bottom))
            cols = _map_log_columns(top_filled, bottom_padded, n)
            
            if 'date' not in cols or 'me' not in cols:
                continue
            
            # Determine where data actually begins
            # If bottom row has meaningful sub-headers, skip it too
            has_subheader = any(
                c and c not in ['DATE','DAY','MAIN','M/E','ENGINE']
                for c in bottom_padded[:n]
            )
            data_start = header_idx + (2 if has_subheader else 1)
            
            df_data = df_raw.iloc[data_start:].copy().reset_index(drop=True)
            
            rows = []
            for _, row in df_data.iterrows():
                vals = list(row.values)
                date = _safe_date(vals[cols['date']] if cols['date'] < len(vals) else None)
                if pd.isna(date): continue
                if not (1980 <= date.year <= 2100): continue
                
                rows.append({
                    'Date': date,
                    'ME_Hours':  _sn0(vals[cols['me']]  if cols['me']  < len(vals) else None),
                    'DG1_Hours': _sn0(vals[cols['dg1']] if 'dg1' in cols and cols['dg1'] < len(vals) else None),
                    'DG2_Hours': _sn0(vals[cols['dg2']] if 'dg2' in cols and cols['dg2'] < len(vals) else None),
                    'DG3_Hours': _sn0(vals[cols['dg3']] if 'dg3' in cols and cols['dg3'] < len(vals) else None),
                })
            
            if rows:
                return pd.DataFrame(rows), None
        
        return None, f"No sheet in '{file_name}' contained a recognisable DATE + MAIN ENGINE header row."
    
    except Exception as e:
        return None, f"Parse error in '{file_name}': {str(e)}"

# ═══════════════════════════════════════════════════════════════════════════════
# PARSE PMS FILE — same dual-header approach
# ═══════════════════════════════════════════════════════════════════════════════
@st.cache_data(show_spinner=False)
def parse_pms_file(file_bytes, file_name):
    """Returns (pms_df, error_string_or_None)"""
    try:
        wb = pd.read_excel(io.BytesIO(file_bytes), sheet_name=None, header=None, dtype=str, engine='openpyxl')
        
        for sheet_name, df_raw in wb.items():
            if df_raw is None or len(df_raw) < 4:
                continue
            
            header_idx, top_filled, bottom = _find_header_row(
                df_raw,
                sentinel_groups=[
                    ['COMPONENT', 'EQUIPMENT', 'DESCRIPTION', 'ITEM', 'NAME', 'JOB'],
                    ['DATE', 'OVERHAUL', 'OH', 'INSP', 'O/H', 'LAST'],
                ]
            )
            
            n = len(df_raw.columns)
            
            if header_idx == -1:
                # TEC-001 hard fallback: data from row 8, col 1=name, col 5=date, col 7=hours
                if len(df_raw) > 10 and n >= 8:
                    header_idx = 7
                    top_filled = [''] * n
                    bottom = [''] * n
                    cols = {'component': 1, 'ohdate': 5, 'hours': 7}
                else:
                    continue
            else:
                bottom_padded = bottom + [''] * (n - len(bottom))
                cols = _map_pms_columns(top_filled, bottom_padded, n)
                if 'component' not in cols: cols['component'] = 1
                if 'ohdate'    not in cols: cols['ohdate']    = 5
                if 'hours'     not in cols: cols['hours']     = 7
            
            has_subheader = any(
                str(x).strip() and str(x).upper() not in ['NAN','']
                for x in (df_raw.iloc[header_idx + 1].values if header_idx + 1 < len(df_raw) else [])
            )
            data_start = header_idx + (2 if has_subheader else 1)
            df_data = df_raw.iloc[data_start:].copy().reset_index(drop=True)
            
            rows = []
            for _, row in df_data.iterrows():
                vals = list(row.values)
                
                raw_name = vals[cols['component']] if cols['component'] < len(vals) else None
                if not raw_name or str(raw_name).strip().upper() in ['NAN','NONE','N/A','','-']:
                    continue
                
                comp_name = str(raw_name).strip()
                oh_date   = _safe_date(vals[cols['ohdate']] if cols['ohdate'] < len(vals) else None)
                if pd.isna(oh_date): continue
                if not (1980 <= oh_date.year <= 2100): continue
                
                hours      = _sn0(vals[cols['hours']] if cols['hours'] < len(vals) else None)
                sys_val    = vals[cols.get('system', -1)] if cols.get('system', -1) >= 0 and cols.get('system', -1) < len(vals) else ''
                parent_sys = _infer_system(comp_name, sys_val)
                
                rows.append({
                    'Component':   comp_name,
                    'OHDate':      oh_date,
                    'LegacyHours': hours,
                    'System':      parent_sys,
                })
            
            if rows:
                return pd.DataFrame(rows), None
        
        return None, f"No sheet in '{file_name}' contained a valid component + overhaul date structure."
    
    except Exception as e:
        return None, f"PMS parse error in '{file_name}': {str(e)}"

# ═══════════════════════════════════════════════════════════════════════════════
# MASTER TIMELINE — stitch N log DataFrames, dedup by date, sort
# ═══════════════════════════════════════════════════════════════════════════════
def build_timeline(log_dfs):
    combined = pd.concat(log_dfs, ignore_index=True)
    combined['DateKey'] = combined['Date'].dt.strftime('%Y-%m-%d')
    combined = combined.sort_values('Date')
    combined = combined.drop_duplicates(subset='DateKey', keep='last').reset_index(drop=True)
    return combined.drop(columns='DateKey')

def detect_gaps(timeline_df):
    gaps = []
    dates = timeline_df['Date'].sort_values().reset_index(drop=True)
    for i in range(1, len(dates)):
        diff = (dates.iloc[i] - dates.iloc[i-1]).days
        if diff > 1:
            gaps.append({'From': dates.iloc[i-1], 'To': dates.iloc[i], 'Missing Days': diff - 1})
    return gaps

# ═══════════════════════════════════════════════════════════════════════════════
# PHYSICS VIOLATIONS — hours > 24 or < 0 in any system
# ═══════════════════════════════════════════════════════════════════════════════
def detect_physics_violations(timeline_df):
    violations = []
    sys_cols = [('ME_Hours','MAIN ENGINE'), ('DG1_Hours','DG1'), ('DG2_Hours','DG2'), ('DG3_Hours','DG3')]
    for col, label in sys_cols:
        if col not in timeline_df.columns: continue
        bad = timeline_df[(timeline_df[col] > 24) | (timeline_df[col] < 0)]
        for _, r in bad.iterrows():
            h = r[col]
            violations.append({
                'Date':   r['Date'].strftime('%d %b %Y'),
                'System': label,
                'Hours':  round(h, 2),
                'Reason': f"Exceeds 24h max" if h > 24 else "Negative hours"
            })
    return pd.DataFrame(violations) if violations else pd.DataFrame(columns=['Date','System','Hours','Reason'])

# ═══════════════════════════════════════════════════════════════════════════════
# DISCREPANCY ENGINE
# ═══════════════════════════════════════════════════════════════════════════════
def _col_for_system(sys):
    return {'MAIN_ENGINE':'ME_Hours','DG1':'DG1_Hours','DG2':'DG2_Hours','DG3':'DG3_Hours'}.get(sys,'ME_Hours')

def calculate_discrepancies(pms_df, timeline_df):
    results = []
    timeline_start = timeline_df['Date'].min() if not timeline_df.empty else pd.NaT
    
    for _, comp in pms_df.iterrows():
        oh_date = comp['OHDate']
        col     = _col_for_system(comp['System'])
        
        mask    = timeline_df['Date'] >= oh_date
        subset  = timeline_df.loc[mask, col] if col in timeline_df.columns else pd.Series([], dtype=float)
        verified = round(subset.sum(), 1) if not subset.empty else 0.0
        legacy   = comp['LegacyHours']
        delta    = round(verified - legacy, 1)
        
        # Confidence
        conf, note = 'HIGH', ''
        if pd.isna(timeline_start) or timeline_start > oh_date:
            gap = int((timeline_start - oh_date).days) if pd.notna(timeline_start) else 0
            conf = 'LOW' if gap > 30 else 'MEDIUM'
            note = f"Log data starts {gap}d after overhaul"
        elif detect_gaps(timeline_df[timeline_df['Date'] >= oh_date]):
            conf = 'MEDIUM'
            note = "Gaps in timeline after overhaul date"
        
        results.append({
            'Component':      comp['Component'],
            'System':         comp['System'].replace('_',' '),
            'Overhaul Date':  oh_date.strftime('%d %b %Y'),
            'Legacy (h)':     legacy,
            'Verified (h)':   verified,
            'Delta (h)':      delta,
            'Confidence':     conf,
            'Status':         'VERIFIED' if abs(delta) <= 0.5 else 'DRIFT DETECTED',
            'Note':           note,
        })
    
    return pd.DataFrame(results)

# ═══════════════════════════════════════════════════════════════════════════════
# DIGITAL SEAL — SHA-256 of canonical result JSON
# ═══════════════════════════════════════════════════════════════════════════════
def compute_seal(results_df):
    payload = results_df[['Component','Legacy (h)','Verified (h)','Delta (h)','Status']].to_json(orient='records')
    return hashlib.sha256(payload.encode()).hexdigest()

# ═══════════════════════════════════════════════════════════════════════════════
# FRONTEND
# ═══════════════════════════════════════════════════════════════════════════════
col1, col2 = st.columns(2)
with col1:
    st.markdown("<div style='font-size:0.75rem;font-weight:700;color:#64748b;text-transform:uppercase;letter-spacing:0.1em;margin-bottom:8px'>① PMS BASELINE (TEC-001)</div>", unsafe_allow_html=True)
    pms_file = st.file_uploader("Upload TEC-001 Master Sheet", type=["xlsx","xls"], key="pms", label_visibility="collapsed")

with col2:
    st.markdown("<div style='font-size:0.75rem;font-weight:700;color:#64748b;text-transform:uppercase;letter-spacing:0.1em;margin-bottom:8px'>② OPERATING HOURS LOGS (multi-select)</div>", unsafe_allow_html=True)
    log_files = st.file_uploader("Upload Log Files", type=["xlsx","xls","csv","txt"], accept_multiple_files=True, key="logs", label_visibility="collapsed")

if not pms_file and not log_files:
    st.info("Upload your PMS baseline (TEC-001) and one or more monthly running-hours log files to begin the forensic audit.")
    st.stop()

if pms_file is None:
    st.warning("Waiting for PMS file.")
    st.stop()

if not log_files:
    st.warning("Waiting for at least one log file.")
    st.stop()

# Parse PMS
with st.spinner("Parsing PMS baseline…"):
    pms_df, pms_err = parse_pms_file(pms_file.getvalue(), pms_file.name)

if pms_err or pms_df is None or pms_df.empty:
    st.error(f"❌ PMS Parse Failed: {pms_err or 'No data extracted.'}")
    st.stop()

st.success(f"✓ PMS: {len(pms_df)} components extracted from **{pms_file.name}**")

# Parse all log files
all_log_dfs, failed = [], []
for f in log_files:
    with st.spinner(f"Parsing {f.name}…"):
        df, err = parse_log_file(f.getvalue(), f.name)
    if err or df is None or df.empty:
        failed.append(f"{f.name}: {err or 'No data extracted'}")
    else:
        all_log_dfs.append(df)
        st.success(f"✓ Log: **{f.name}** — {len(df)} daily entries")

if failed:
    st.warning("⚠️ Skipped files: " + " | ".join(failed))

if not all_log_dfs:
    st.error("❌ No valid log data could be extracted. Verify file formats.")
    st.stop()

# Build timeline
timeline_df = build_timeline(all_log_dfs)
gaps        = detect_gaps(timeline_df)
violations  = detect_physics_violations(timeline_df)
results_df  = calculate_discrepancies(pms_df, timeline_df)
seal        = compute_seal(results_df)

drift_count = len(results_df[results_df['Status'] == 'DRIFT DETECTED'])
viol_count  = len(violations)

# ── KPI Cards ──────────────────────────────────────────────────────────────────
st.markdown("<br>", unsafe_allow_html=True)
kpi_html = f"""
<div class="kpi-grid">
  <div class="kpi-card green">
    <div class="kpi-label">Components Audited</div>
    <div class="kpi-val">{len(results_df)}</div>
    <div class="kpi-sub">Extracted from PMS baseline</div>
  </div>
  <div class="kpi-card {'red' if drift_count > 0 else 'green'}">
    <div class="kpi-label">Drift Anomalies</div>
    <div class="kpi-val" style="color:{'#ef4444' if drift_count > 0 else '#10b981'}">{drift_count}</div>
    <div class="kpi-sub">Hours do not match legacy claim</div>
  </div>
  <div class="kpi-card violet">
    <div class="kpi-label">Days in Timeline</div>
    <div class="kpi-val">{len(timeline_df)}</div>
    <div class="kpi-sub">{timeline_df['Date'].min().strftime('%d %b %Y')} → {timeline_df['Date'].max().strftime('%d %b %Y')}</div>
  </div>
  <div class="kpi-card {'amber' if viol_count > 0 else 'green'}">
    <div class="kpi-label">Physics Violations</div>
    <div class="kpi-val" style="color:{'#f59e0b' if viol_count > 0 else '#10b981'}">{viol_count}</div>
    <div class="kpi-sub">Hours &gt; 24h or &lt; 0 in any system</div>
  </div>
  <div class="kpi-card {'amber' if gaps else 'green'}">
    <div class="kpi-label">Timeline Gaps</div>
    <div class="kpi-val">{len(gaps)}</div>
    <div class="kpi-sub">{sum(g['Missing Days'] for g in gaps)} missing day(s) total</div>
  </div>
  <div class="kpi-card">
    <div class="kpi-label">Log Files Merged</div>
    <div class="kpi-val">{len(all_log_dfs)}</div>
    <div class="kpi-sub">Deduplicated & stitched</div>
  </div>
</div>
"""
st.markdown(kpi_html, unsafe_allow_html=True)

# ── Digital Seal ───────────────────────────────────────────────────────────────
st.markdown(f'<div class="seal-box">🔐 SHA-256 INTEGRITY SEAL: {seal}</div>', unsafe_allow_html=True)

# ── Tabs ───────────────────────────────────────────────────────────────────────
t1, t2, t3, t4 = st.tabs(["📑 FORENSIC RECONCILIATION", "⚠️ PHYSICS VIOLATIONS", "📅 TIMELINE", "🔍 RAW TIMELINE DATA"])

with t1:
    st.markdown("**[START ROB] + [BUNKERS] − [END ROB] = [PHYSICAL BURN]**" if False else "**Verified Hours** are summed from the master timeline from each component's last overhaul date onwards.")

    def _style_row(row):
        if row['Status'] == 'DRIFT DETECTED': return ['background-color: #fef2f2; color: #991b1b'] * len(row)
        if row['Confidence'] == 'LOW':        return ['background-color: #fffbeb; color: #92400e'] * len(row)
        if row['Confidence'] == 'MEDIUM':     return ['background-color: #fefce8; color: #854d0e'] * len(row)
        return ['color: #065f46'] * len(row)

    st.dataframe(
        results_df.style.apply(_style_row, axis=1),
        use_container_width=True, hide_index=True, height=500,
        column_config={
            'Legacy (h)':   st.column_config.NumberColumn(format='%.1f'),
            'Verified (h)': st.column_config.NumberColumn(format='%.1f'),
            'Delta (h)':    st.column_config.NumberColumn(format='%.1f'),
        }
    )

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine='openpyxl') as wr:
        results_df.to_excel(wr, index=False, sheet_name='Audit Results')
        if not violations.empty:
            violations.to_excel(wr, index=False, sheet_name='Physics Violations')
        pd.DataFrame([{'Seal': seal, 'Generated': datetime.now().isoformat(), 'Components': len(results_df)}]).to_excel(wr, index=False, sheet_name='Digital Seal')
    buf.seek(0)
    st.download_button("⬇️ Download Audit Workbook (.xlsx)", data=buf, file_name=f"VRS_Audit_{datetime.now().strftime('%Y%m%d')}.xlsx", mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

with t2:
    if violations.empty:
        st.success("✅ Zero physics violations. All daily entries are within the 0–24h physical bounds.")
    else:
        st.error(f"⚠️ {len(violations)} impossible values detected. These are **included** in the audit arithmetic — not discarded.")
        st.dataframe(violations, use_container_width=True, hide_index=True)

with t3:
    if gaps:
        gap_df = pd.DataFrame(gaps)
        gap_df['From'] = gap_df['From'].dt.strftime('%d %b %Y')
        gap_df['To']   = gap_df['To'].dt.strftime('%d %b %Y')
        st.warning(f"⚠️ {len(gaps)} chronological gap(s) found — {sum(g['Missing Days'] for g in gaps)} total missing days.")
        st.dataframe(gap_df, use_container_width=True, hide_index=True)
    else:
        st.success("✅ No chronological gaps in the master timeline.")

with t4:
    st.caption(f"Master timeline: {len(timeline_df)} days · {len(all_log_dfs)} source file(s) merged")
    st.dataframe(
        timeline_df.rename(columns={'Date':'Date','ME_Hours':'ME (h)','DG1_Hours':'DG1 (h)','DG2_Hours':'DG2 (h)','DG3_Hours':'DG3 (h)'}),
        use_container_width=True, hide_index=True, height=400
    )
