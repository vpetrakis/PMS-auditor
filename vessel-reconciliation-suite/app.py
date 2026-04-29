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

warnings.filterwarnings("ignore")

st.set_page_config(page_title="POSEIDON | Recon Suite", page_icon="⚓", layout="wide", initial_sidebar_state="collapsed")

# ═══════════════════════════════════════════════════════════════════════════════
# 1. ENTERPRISE PAGE CONFIGURATION & POSEIDON TITAN CSS
# ═══════════════════════════════════════════════════════════════════════════════

def _u(s):
    return f"data:image/svg+xml;base64,{base64.b64encode(s.encode()).decode()}"

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
# 2. HELPERS
# ═══════════════════════════════════════════════════════════════════════════════

def normalize_text(s):
    s = str(s).upper().strip()
    s = s.replace("\xa0", " ")
    s = re.sub(r"\s+", " ", s)
    return s


def extract_first_float(value):
    if pd.isna(value):
        return np.nan
    s = str(value).replace(",", "")
    m = re.search(r"-?\d+(?:\.\d+)?", s)
    return float(m.group()) if m else np.nan


def robust_parse_date(value):
    if pd.isna(value):
        return pd.NaT
    s = str(value).strip()
    if not s:
        return pd.NaT

    direct = pd.to_datetime(s, errors="coerce", dayfirst=True)
    if pd.notna(direct):
        return direct

    cleaned = re.sub(r"\s+", " ", s)
    fmts = [
        "%Y-%m-%d %H%M%S",
        "%Y-%m-%d",
        "%d-%m-%Y",
        "%d/%m/%Y",
        "%d.%m.%Y",
        "%d-%b-%Y",
        "%d-%B-%Y",
    ]
    for fmt in fmts:
        try:
            return pd.to_datetime(cleaned, format=fmt, errors="raise", dayfirst=True)
        except Exception:
            pass
    return pd.NaT


def is_date_like(value):
    return pd.notna(robust_parse_date(value))


def find_target_sheet(xls):
    preferred = []
    for s in xls.sheet_names:
        su = s.upper()
        score = 0
        if "DAILY" in su:
            score += 3
        if "OPERATING" in su:
            score += 3
        if "HOURS" in su:
            score += 2
        if score > 0:
            preferred.append((score, s))
    if preferred:
        preferred.sort(reverse=True)
        return preferred[0][1]
    return xls.sheet_names[0]

# ═══════════════════════════════════════════════════════════════════════════════
# 3. EXCEL PARSER - FIXED WITHOUT BREAKING APP INTEGRITY
# ═══════════════════════════════════════════════════════════════════════════════

@st.cache_data(show_spinner=False)
def parse_daily_hours_excel(file_bytes):
    """
    Robust parser for the DAILY OPERATING HOURS workbook.
    Returns:
      - clean_df with Date + ME_Hours + DG1_Hours + DG2_Hours + DG3_Hours
      - raw diagnostic sheet
      - col_map used in extraction
    """
    try:
        xls = pd.ExcelFile(io.BytesIO(file_bytes), engine="openpyxl")
        target_sheet = find_target_sheet(xls)
        df_raw = pd.read_excel(io.BytesIO(file_bytes), sheet_name=target_sheet, header=None, dtype=object, engine="openpyxl")
    except Exception:
        try:
            df_raw = pd.read_excel(io.BytesIO(file_bytes), header=None, dtype=object, engine="openpyxl")
        except Exception:
            return pd.DataFrame(), pd.DataFrame(), {}

    if df_raw.empty:
        return pd.DataFrame(), df_raw, {}

    diag_raw = df_raw.copy()

    header_scan_rows = min(35, len(df_raw))
    header_block = df_raw.head(header_scan_rows).copy().ffill(axis=1).ffill(axis=0)

    # Detect date column first
    date_col = None
    for col_idx in range(header_block.shape[1]):
        col_text = " | ".join([normalize_text(x) for x in header_block.iloc[:, col_idx].tolist() if pd.notna(x)])
        if "DATE" in col_text or re.search(r"\bDAY\b", col_text):
            date_col = col_idx
            break

    if date_col is None:
        # fallback: first column with at least 3 date-like cells in first 80 rows
        for col_idx in range(df_raw.shape[1]):
            probe = df_raw.iloc[:80, col_idx]
            hits = sum(is_date_like(v) for v in probe)
            if hits >= 3:
                date_col = col_idx
                break

    col_map = {}
    if date_col is not None:
        col_map["Date"] = date_col

    # Identify engine columns by the header row nearest to the first real date row
    first_date_row = None
    if date_col is not None:
        for i in range(len(df_raw)):
            if is_date_like(df_raw.iloc[i, date_col]):
                first_date_row = i
                break

    if first_date_row is None:
        return pd.DataFrame(), diag_raw, col_map

    header_rows_to_use = list(range(max(0, first_date_row - 4), first_date_row + 1))
    working_header = df_raw.iloc[header_rows_to_use, :].copy().ffill(axis=1).ffill(axis=0)

    def classify_col(col_idx):
        vals = [normalize_text(v) for v in working_header.iloc[:, col_idx].tolist() if pd.notna(v)]
        txt = " | ".join(vals)
        if "MAIN ENGINE" in txt or "M/E" in txt:
            return "ME_Hours"
        if "DIESEL GENERATOR NO.1" in txt or "DIESEL GENERATOR NO 1" in txt or "DG 1" in txt or "D/G 1" in txt:
            return "DG1_Hours"
        if "DIESEL GENERATOR NO.2" in txt or "DIESEL GENERATOR NO 2" in txt or "DG 2" in txt or "D/G 2" in txt:
            return "DG2_Hours"
        if "DIESEL GENERATOR NO.3" in txt or "DIESEL GENERATOR
