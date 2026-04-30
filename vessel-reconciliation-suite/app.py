import streamlit as st
import pandas as pd
import numpy as np
import re
import io
import math
import traceback
import base64
import warnings
from pathlib import Path
import plotly.graph_objects as go
from plotly.subplots import make_subplots

# ═══════════════════════════════════════════════════════════════════════════════
# DEPENDENCIES & SETUP
# ═══════════════════════════════════════════════════════════════════════════════
try:
    from xgboost import XGBRegressor
    from sklearn.covariance import LedoitWolf
    from sklearn.model_selection import KFold
    import shap
    HAS_ML = True
except ImportError:
    HAS_ML = False

warnings.filterwarnings("ignore")
st.set_page_config(page_title="POSEIDON TITAN", page_icon="⚓", layout="wide", initial_sidebar_state="collapsed")

# ═══════════════════════════════════════════════════════════════════════════════
# CSS LOADER & PREMIUM ANIMATED SVG ASSETS
# ═══════════════════════════════════════════════════════════════════════════════
def load_local_css():
    css_path = Path(__file__).parent / "assets" / "style.css"
    if css_path.exists():
        st.markdown(f"<style>{css_path.read_text()}</style>", unsafe_allow_html=True)

load_local_css()

def _u(s):
    return f"data:image/svg+xml;base64,{base64.b64encode(s.encode()).decode()}"

LOGO_SVG = base64.b64encode(
    b'<svg viewBox="0 0 48 48" xmlns="http://www.w3.org/2000/svg"><defs><linearGradient id="pg" x1="0" y1="0" x2="1" y2="1"><stop offset="0%" stop-color="#c9a84c"/><stop offset="50%" stop-color="#00e0b0"/><stop offset="100%" stop-color="#fff"/></linearGradient></defs><circle cx="24" cy="24" r="22" fill="none" stroke="url(#pg)" stroke-width="0.8" opacity=".3"/><path d="M24 6L24 42" stroke="url(#pg)" stroke-width="1.5" stroke-linecap="round"/><path d="M12 24Q24 32 36 24" fill="none" stroke="url(#pg)" stroke-width="1.5" stroke-linecap="round"/></svg>'
).decode()

VERIFIED_SVG = """<svg viewBox="0 0 28 28" xmlns="http://www.w3.org/2000/svg"><style>@keyframes pulse { 0% { r: 12; opacity: 0.2; } 50% { r: 13.5; opacity: 0.6; } 100% { r: 12; opacity: 0.2; } } .p { animation: pulse 2s infinite ease-in-out; }</style><circle cx="14" cy="14" r="12" fill="none" stroke="#00e0b0" stroke-width="1" class="p"/><circle cx="14" cy="14" r="7.5" fill="#061a14" stroke="#00e0b0" stroke-width="1.5"/><polyline points="10,14.5 12.8,17 18,10.5" fill="none" stroke="#00e0b0" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"/></svg>"""
GHOST_SVG = """<svg viewBox="0 0 28 28" xmlns="http://www.w3.org/2000/svg"><style>@keyframes flash { 0%, 100% { opacity: 1; } 50% { opacity: 0.3; } } @keyframes spin { 100% { transform: rotate(360deg); } } .f { animation: flash 1s infinite; } .s { transform-origin: center; animation: spin 5s linear infinite; }</style><circle cx="14" cy="14" r="12" fill="none" stroke="#ff2a55" stroke-width="1" stroke-dasharray="4 3" class="s"/><circle cx="14" cy="14" r="7.5" fill="#1a0508" stroke="#ff2a55" stroke-width="1.5" class="f"/><g stroke="#ff2a55" stroke-width="2.5" stroke-linecap="round" class="f"><line x1="11" y1="11" x2="17" y2="17"/><line x1="17" y1="11" x2="11" y2="17"/></g></svg>"""
OUTLIER_SVG = """<svg viewBox="0 0 28 28" xmlns="http://www.w3.org/2000/svg"><style>@keyframes breathe { 0%, 100% { stroke-width: 1.2; transform: scale(1); } 50% { stroke-width: 2.2; transform: scale(1.05); } } .b { transform-origin: center; animation: breathe 2s infinite ease-in-out; }</style><rect x="4" y="4" width="20" height="20" rx="5" fill="none" stroke="#c9a84c" stroke-width="1.2" class="b"/><circle cx="14" cy="14" r="4.5" fill="#0e0a1e" stroke="#c9a84c" stroke-width="1.5"/><circle cx="14" cy="14" r="1.8" fill="#c9a84c"/></svg>"""

ICONS = {
    "VERIFIED": _u(VERIFIED_SVG),
    "GHOST BUNKER": _u(GHOST_SVG),
    "STAT OUTLIER": _u(OUTLIER_SVG)
}

STATUS_COLORS = {
    "VERIFIED": "#00e0b0",
    "GHOST BUNKER": "#ff2a55",
    "STAT OUTLIER": "#c9a84c"
}

REQUIRED_RAW_COLS = [
    "FO_A", "FO_L", "MGO_A", "MGO_L", "Bunk_FO", "Bunk_MGO", "Bunk_MELO", "Bunk_HSCYLO",
    "Bunk_LSCYLO", "Bunk_GELO", "Bunk_CYLO", "MELO_R", "HSCYLO_R", "LSCYLO_R", "GELO_R",
    "CYLO_R", "Speed", "DistLeg", "TotalDist", "CargoQty", "Voy", "Port", "AD", "Date", "Time"
]

# ═══════════════════════════════════════════════════════════════════════════════
# FORENSIC UTILITIES & LEXICAL SIEVE
# ═══════════════════════════════════════════════════════════════════════════════
@st.cache_data(show_spinner=False)
def load_fleet_master():
    db_path = Path(__file__).parent / "fleet_master.csv"
    if db_path.exists():
        try:
            return pd.read_csv(db_path).set_index("Vessel_Name")
        except Exception:
            pass
    return pd.DataFrame(columns=["Min_Speed_kn", "Ghost_Tol_Sea", "Ghost_Tol_Port"])

fleet_db = load_fleet_master()

def _sn(val):
    if pd.isna(val):
        return np.nan
    s = str(val).strip().upper()
    if s in ["NIL", "N/A", "NA", "XXX", "NONE", "UNKNOWN", "BLANK", "-", "X", "", "NULL"]:
        return np.nan
    s = re.sub(r"[^\d.\-]", "", s)
    try:
        return float(s) if s and s not in (".", "-", "-.") else np.nan
    except ValueError:
        return np.nan

def _sn0(val):
    v = _sn(val)
    return 0.0 if np.isnan(v) else v

def _parse_dt(d_val, t_val):
    try:
        if pd.isna(d_val) or str(d_val).strip() == "":
            return pd.NaT
        ds = str(d_val).strip()
        ds = re.sub(r"20224", "2024", ds)
        ds = re.sub(r"20023", "2023", ds)
        ds = re.sub(
            r"(\d+)\s+([A-Za-z]+)\.?\s+(\d{4})",
            lambda m: f"{m.group(3)}-{m.group(2)[:3]}-{m.group(1).zfill(2)}",
            ds
        )
        p = pd.to_datetime(ds, errors="coerce")
        if pd.isna(p):
            return pd.NaT
        d_str = p.strftime("%Y-%m-%d")
        t_str = "00:00"
        if pd.notna(t_val) and str(t_val).strip() != "":
            tr = re.sub(r"[HhLlTtUuCc\s]", "", str(t_val).strip())
            m = re.match(r"^(\d{1,2}):(\d{2})", tr)
            if m:
                t_str = f"{m.group(1).zfill(2)}:{m.group(2)}"
        return pd.to_datetime(f"{d_str} {t_str}", errors="coerce")
    except Exception:
        return pd.NaT

def compute_dqi(r1, r2, days, phys_burn, drift, ghost_tol):
    if days <= 0 or pd.isna(phys_burn):
        return 0
    scores = [100.0]
    scores.append(100.0 if phys_burn >= ghost_tol else max(0.0, 100 - abs(phys_burn) * 5))
    tol = max(30.0, 0.03 * max(_sn0(r1.get("FO_A")), _sn0(r2.get("FO_A"))))
    scores.append(math.exp(-0.5 * ((drift) / tol) ** 2) * 100 if tol > 0 else 0.0)
    return int(sum(scores) / len(scores))

# ═══════════════════════════════════════════════════════════════════════════════
# THE ROUTER: CONFIGURATION-DRIVEN MANIFEST MAPPING (DECOUPLED EXTRACTION)
# ═══════════════════════════════════════════════════════════════════════════════
MULTI_VERSION_MAP = {
    "COURAGE": 118,
    "DIGNITY": 128,
    "FALCON": 32,
    "GEORGIAT": 175,
    "GEORGIA T": 175,
    "STEFANOST": 201,
    "STEFANOS T": 201,
    "CHRISTIANNA": 85
}

def _map_columns(top_header, bottom_header, num_cols):
    cols_found = {}
    for j in range(num_cols):
        c1 = str(top_header.iloc[j]).upper().strip() if pd.notna(top_header.iloc[j]) else ""
        c2 = str(bottom_header.iloc[j]).upper().strip() if pd.notna(bottom_header.iloc[j]) else ""
        c_comb = f"{c1} {c2}".strip()

        if "VOY" in c_comb:
            cols_found["Voy"] = j
        elif "PORT" in c_comb or "LOC" in c_comb:
            cols_found["Port"] = j
        elif "A/D" in c_comb or c_comb == "AD" or "STATUS" in c_comb:
            cols_found["AD"] = j
        elif "SPEED" in c_comb:
            cols_found["Speed"] = j
        elif "CARGO" in c_comb or "QTY" in c_comb:
            cols_found["CargoQty"] = j
        elif "DATE" in c_comb or "DAY" in c_comb:
            cols_found["Date"] = j
        elif "TIME" in c_comb and "TOTAL" not in c_comb:
            cols_found["Time"] = j
        elif "DIST" in c_comb and "LEG" in c_comb:
            cols_found["DistLeg"] = j
        elif "DIST" in c_comb and "TOTAL" in c_comb:
            cols_found["TotalDist"] = j
        elif "BUNKER" in c1 or "RECEIV" in c1:
            if "FO" in c2 and "MGO" not in c2:
                cols_found["Bunk_FO"] = j
            elif "MGO" in c2:
                cols_found["Bunk_MGO"] = j
            elif "MELO" in c2:
                cols_found["Bunk_MELO"] = j
            elif "HSCYLO" in c2 or "HS CYL" in c2:
                cols_found["Bunk_HSCYLO"] = j
            elif "LSCYLO" in c2 or "LS CYL" in c2:
                cols_found["Bunk_LSCYLO"] = j
            elif "CYLO" in c2 or "CYL OIL" in c2:
                cols_found["Bunk_CYLO"] = j
            elif "GELO" in c2:
                cols_found["Bunk_GELO"] = j
        elif "ROB" in c1 or "REMAIN" in c1:
            if "FO A" in c2 or "FO ACT" in c2:
                cols_found["FO_A"] = j
            elif "FO L" in c2 or "FO LED" in c2:
                cols_found["FO_L"] = j
            elif "MGO A" in c2:
                cols_found["MGO_A"] = j
            elif "MGO L" in c2:
                cols_found["MGO_L"] = j
            elif "MELO" in c2:
                cols_found["MELO_R"] = j
            elif "HSCYLO" in c2 or "HS CYL" in c2:
                cols_found["HSCYLO_R"] = j
            elif "LSCYLO" in c2 or "LS CYL" in c2:
                cols_found["LSCYLO_R"] = j
            elif "CYLO" in c2 or "CYL OIL" in c2:
                cols_found["CYLO_R"] = j
            elif "GELO" in c2:
                cols_found["GELO_R"] = j
    return cols_found

def _parse_standard(df_raw):
    header_idx = -1
    cols_found = {}
    for i in range(min(150, len(df_raw))):
        vals = [str(x).upper() for x in df_raw.iloc[i].values if pd.notna(x)]
        if any(k in v for v in vals for k in ["DATE", "DAY"]) and any(k in v for v in vals for k in ["PORT", "LOC"]):
            header_idx = i
            top_header = df_raw.iloc[i].ffill()
            bottom_header = df_raw.iloc[i + 1] if i + 1 < len(df_raw) else pd.Series([np.nan] * len(df_raw.columns))
            cols_found = _map_columns(top_header, bottom_header, len(df_raw.columns))

    if header_idx == -1:
        raise ValueError("Matrix Lock Failed: No valid headers found.")

    df = df_raw.iloc[header_idx + 1:].copy().reset_index(drop=True)
    for std_name, exc_idx in cols_found.items():
        df[std_name] = df.iloc[:, exc_idx]
    return df

def _parse_manifest(df_raw, start_row):
    idx = max(0, int(start_row) - 1)

    header_idx = -1
    cols_found = {}
    for i in range(min(15, len(df_raw))):
        vals = [str(x).upper() for x in df_raw.iloc[i].values if pd.notna(x)]
        if any(k in v for v in vals for k in ["DATE", "DAY"]) and any(k in v for v in vals for k in ["PORT", "LOC"]):
            header_idx = i
            top_header = df_raw.iloc[i].ffill()
            bottom_header = df_raw.iloc[i + 1] if i + 1 < len(df_raw) else pd.Series([np.nan] * len(df_raw.columns))
            cols_found = _map_columns(top_header, bottom_header, len(df_raw.columns))
            break

    if header_idx == -1:
        raise ValueError("Global Header Check Failed: Ensure 'DATE' and 'PORT' exist in the top rows of the file.")

    clean_data_chunk = df_raw.iloc[idx:].copy().reset_index(drop=True)

    df = pd.DataFrame()
    for std_name, exc_idx in cols_found.items():
        df[std_name] = clean_data_chunk.iloc[:, exc_idx]

    return df

def semantic_parse(file_bytes, file_name):
    vn_raw = re.sub(r"\.[^.]+$", "", file_name).strip()
    vname = re.sub(r"[_\-]+", " ", vn_raw).upper()

    if file_name.lower().endswith(".xlsx"):
        df_raw = pd.read_excel(io.BytesIO(file_bytes), header=None, engine="openpyxl", dtype=str)
    else:
        df_raw = pd.read_csv(io.StringIO(file_bytes.decode("latin-1", errors="replace")), header=None, on_bad_lines="skip", dtype=str)

    if df_raw.empty or len(df_raw) < 4:
        raise ValueError("File is empty or severely malformed.")

    is_multi_version = False
    split_row = 0
    for vessel, row in MULTI_VERSION_MAP.items():
        if vessel in vname:
            is_multi_version = True
            split_row = row
            break

    if is_multi_version:
        df = _parse_manifest(df_raw, split_row)
    else:
        df = _parse_standard(df_raw)

    missing = [col for col in REQUIRED_RAW_COLS if col not in df.columns]
    for req in missing:
        df[req] = np.nan

    math_cols = [
        "FO_A", "FO_L", "MGO_A", "MGO_L", "Bunk_FO", "Bunk_MGO", "Bunk_MELO", "Bunk_HSCYLO",
        "Bunk_LSCYLO", "Bunk_GELO", "Bunk_CYLO", "MELO_R", "HSCYLO_R", "LSCYLO_R", "GELO_R",
        "CYLO_R", "Speed", "DistLeg", "TotalDist", "CargoQty"
    ]
    for col in math_cols:
        df[col] = df[col].apply(_sn)

    string_cols = ["Voy", "Port", "AD", "Date", "Time"]
    for col in string_cols:
        df[col] = df[col].fillna("").astype(str).str.strip()

    df["Datetime"] = df.apply(lambda r: _parse_dt(r.get("Date"), r.get("Time")), axis=1)
    df = df.dropna(subset=["Datetime"]).sort_values("Datetime").reset_index(drop=True)
    df["AD"] = df["AD"].apply(lambda v: "D" if v.upper() in ["D", "DEP", "SBE", "FAOP"] else ("A" if v.upper().startswith("A") else v))
    return df, vname

# ═══════════════════════════════════════════════════════════════════════════════
# TRI-STATE AD-TO-AD STATE MACHINE (KINEMATIC IMPUTATION PROTOCOL)
# ═══════════════════════════════════════════════════════════════════════════════
def build_state_machine(df, min_speed, ghost_sea, ghost_port):
    ad_events = df[df["AD"].isin(["A", "D"])].copy()
    if len(ad_events) < 2:
        raise ValueError("Insufficient A/D events to construct a timeline.")

    ad_events["Prev_AD"] = ad_events["AD"].shift(1)
    ad_events = ad_events[ad_events["AD"] != ad_events["Prev_AD"]].drop(columns=["Prev_AD"]).copy()

    trips, cum_drift = [], []
    for i in range(len(ad_events) - 1):
        r1, r2 = ad_events.iloc[i], ad_events.iloc[i + 1]
        idx1, idx2 = r1.name, r2.name
        status, flags = "VERIFIED", []
        phys_burn, log_burn, drift, daily_burn, days = np.nan, np.nan, np.nan, np.nan, 0.0

        phase = "SEA" if r1["AD"] == "D" else "PORT"
        days = (r2["Datetime"] - r1["Datetime"]).total_seconds() / 86400.0
        if days <= 0:
            days, flags = 0.02, flags + ["Time Delta Fallback"]

        start_rob, end_rob = r1.get("FO_A"), r2.get("FO_A")
        if pd.isna(start_rob) or pd.isna(end_rob):
            status, flags = "QUARANTINE_ROB", flags + ["Missing Sounding"]

        if r1["AD"] == "D" and not pd.isna(start_rob):
            fol = r1.get("FO_L")
            cum_drift.append({
                "dt": r1["Datetime"],
                "gap": start_rob - (fol if not pd.isna(fol) else start_rob),
                "port": r1.get("Port", "")[:20]
            })

        window = df.loc[idx1 + 1:idx2]
        if phase == "PORT":
            bfo = df.loc[idx1:idx2, "Bunk_FO"].sum(skipna=True)
            b_melo = df.loc[idx1:idx2, "Bunk_MELO"].sum(skipna=True)
            b_hscylo = df.loc[idx1:idx2, "Bunk_HSCYLO"].sum(skipna=True)
            b_lscylo = df.loc[idx1:idx2, "Bunk_LSCYLO"].sum(skipna=True)
            b_cylo = df.loc[idx1:idx2, "Bunk_CYLO"].sum(skipna=True)
            b_gelo = df.loc[idx1:idx2, "Bunk_GELO"].sum(skipna=True)
        else:
            bfo = window["Bunk_FO"].sum(skipna=True)
            b_melo = window["Bunk_MELO"].sum(skipna=True)
            b_hscylo = window["Bunk_HSCYLO"].sum(skipna=True)
            b_lscylo = window["Bunk_LSCYLO"].sum(skipna=True)
            b_cylo = window["Bunk_CYLO"].sum(skipna=True)
            b_gelo = window["Bunk_GELO"].sum(skipna=True)

        speed = window["Speed"].replace(0, np.nan).mean() if not window["Speed"].empty else np.nan
        dist = window["DistLeg"].sum(skipna=True)

        if dist <= 0 and phase == "SEA":
            dist = max(0, _sn0(r2.get("TotalDist")) - _sn0(r1.get("TotalDist")))
            if dist <= 0 and not pd.isna(speed):
                dist = speed * (days * 24.0)
                flags.append("Distance Imputed from Kinematics")

        if pd.isna(speed):
            speed = dist / (days * 24.0) if days > 0 else 0.0

        melo_c = max(0, (_sn0(r1.get("MELO_R")) - _sn0(r2.get("MELO_R"))) + b_melo)
        hscylo_c = max(0, (_sn0(r1.get("HSCYLO_R")) - _sn0(r2.get("HSCYLO_R"))) + b_hscylo)
        lscylo_c = max(0, (_sn0(r1.get("LSCYLO_R")) - _sn0(r2.get("LSCYLO_R"))) + b_lscylo)
        cylo_gen_c = max(0, (_sn0(r1.get("CYLO_R")) - _sn0(r2.get("CYLO_R"))) + b_cylo)
        gelo_c = max(0, (_sn0(r1.get("GELO_R")) - _sn0(r2.get("GELO_R"))) + b_gelo)

        dqi = 0
        if status == "VERIFIED" or "QUARANTINE" not in status:
            phys_burn = (start_rob - end_rob) + bfo
            log_start = r1.get("FO_L") if not pd.isna(r1.get("FO_L")) else start_rob
            log_end = r2.get("FO_L") if not pd.isna(r2.get("FO_L")) else end_rob
            log_burn = (log_start - log_end) + bfo
            drift = phys_burn - log_burn
            daily_burn = phys_burn / days

            if bfo < 0:
                status, flags = "QUARANTINE", flags + ["Negative Bunker Input"]
            if abs(drift) > 20 and abs(abs(drift) - abs(bfo)) < 5.0:
                status, flags = "QUARANTINE", flags + ["Mass Imbalance"]
            if daily_burn > 250:
                status, flags = "QUARANTINE", flags + ["MCR Limit Exceeded"]
            if phase == "PORT" and phys_burn < ghost_port and "QUARANTINE" not in status:
                status, flags = "GHOST BUNKER", flags + ["Missing Receipt"]
            elif phase == "SEA" and phys_burn < ghost_sea and "QUARANTINE" not in status:
                status, flags = "GHOST BUNKER", flags + ["Negative Burn"]

            dqi = compute_dqi(r1, r2, days, phys_burn, drift, ghost_tol=(ghost_port if phase == "PORT" else ghost_sea))

        trips.append({
            "Indicator": ICONS.get(status, ICONS["VERIFIED"]) if "QUARANTINE" not in status else "⛔",
            "Timeline": f"{r1['Datetime'].strftime('%d %b %y')} → {r2['Datetime'].strftime('%d %b %y')}",
            "Date_Start_TS": r1["Datetime"],
            "Phase": phase,
            "Condition": "LADEN" if _sn0(r1.get("CargoQty", 0)) > 100 else "BALLAST",
            "Voy": r1.get("Voy", ""),
            "Route": f"{r1.get('Port','')[:15]} → {r2.get('Port','')[:15]}" if phase == "SEA" else f"Port Idle: {r1.get('Port','')[:15]}",
            "Days": round(days, 2),
            "Dist_NM": round(dist, 0),
            "Speed_kn": round(speed, 1),
            "CargoQty": _sn0(r1.get("CargoQty", 0)),
            "FO_A_Start": start_rob if status == "VERIFIED" else np.nan,
            "Bunk_FO": bfo,
            "FO_A_End": end_rob if status == "VERIFIED" else np.nan,
            "Phys_Burn": round(phys_burn, 1),
            "Log_Burn": round(log_burn, 1),
            "Drift_MT": round(drift, 1),
            "Daily_Burn": round(daily_burn, 1) if status == "VERIFIED" else np.nan,
            "MELO_L": round(melo_c, 0),
            "HSCYLO_L": round(hscylo_c, 0),
            "LSCYLO_L": round(lscylo_c, 0),
            "CYLO_GEN_L": round(cylo_gen_c, 0),
            "GELO_L": round(gelo_c, 0),
            "Total_CYLO": round(hscylo_c + lscylo_c + cylo_gen_c, 0),
            "DQI": int(dqi),
            "Status": status,
            "Flags": ", ".join(flags) if flags else ""
        })

    trip_df = pd.DataFrame(trips)
    if len(trip_df) >= 4:
        for cond in ["LADEN", "BALLAST"]:
            ver = trip_df[
                (trip_df["Status"] == "VERIFIED") &
                (trip_df["Phase"] == "SEA") &
                (trip_df["Phys_Burn"] > 0) &
                (trip_df["Condition"] == cond)
            ]
            if len(ver) >= 4:
                q1, q3 = ver["Daily_Burn"].quantile(0.25), ver["Daily_Burn"].quantile(0.75)
                iqr = q3 - q1
                if iqr > 0:
                    lo, hi = q1 - 2.0 * iqr, q3 + 2.0 * iqr
                    mask = (
                        (trip_df["Status"] == "VERIFIED") &
                        (trip_df["Phase"] == "SEA") &
                        (trip_df["Condition"] == cond) &
                        ((trip_df["Daily_Burn"] < lo) | (trip_df["Daily_Burn"] > hi))
                    )
                    trip_df.loc[mask, "Status"] = "STAT OUTLIER"
                    trip_df.loc[mask, "Indicator"] = ICONS["STAT OUTLIER"]

    return trip_df, cum_drift

# ═══════════════════════════════════════════════════════════════════════════════
# FULL DATA-DRIVEN PIML — Ledoit-Wolf + K-Fold Conformal
# ═══════════════════════════════════════════════════════════════════════════════
def execute_ai_physics(trip_df, min_speed):
    ai_status_msg = "Enterprise AI Optimized."
    if not HAS_ML:
        return trip_df, "AI Offline: Missing sklearn/xgboost."
    if trip_df.empty:
        return trip_df, "AI Offline: Empty ledger."

    cols_to_add = [
        "AI_Exp", "HM_Base", "Stoch_Var", "SHAP_Base", "SHAP_Propulsion", "SHAP_Mass",
        "SHAP_Kinematics", "SHAP_Season", "SHAP_Degradation", "Exp_Lower", "Exp_Upper",
        "Mahalanobis", "MD_Threshold", "P_Value"
    ]
    for col in cols_to_add:
        if col not in trip_df.columns:
            trip_df[col] = np.nan

    try:
        sea_mask = (
            (trip_df["Phase"] == "SEA") &
            (trip_df["Status"] == "VERIFIED") &
            (trip_df["Speed_kn"] >= min_speed)
        )
        if sea_mask.sum() < 8:
            raise ValueError(f"Insufficient valid Sea Legs ({sea_mask.sum()}). Min 8 req.")

        ml = trip_df.loc[sea_mask].copy()
        ml["True_Mass"] = (ml["CargoQty"].fillna(0) + ml["FO_A_Start"].fillna(0)).clip(lower=0.1)
        ml["SOG"] = ml["Dist_NM"] / np.maximum(ml["Days"] * 24, 0.1)
        ml["Kin_Delta"] = (ml["Speed_kn"] - ml["SOG"]).clip(-3.0, 3.0)
        ml["Accel_Penalty"] = ml["Speed_kn"].diff().fillna(0.0).clip(-2.0, 2.0)
        ml["Speed_Cubed"] = ml["Speed_kn"] ** 3
        ml["Season_Sin"] = np.sin(2 * np.pi * ml["Date_Start_TS"].dt.month.fillna(6) / 12.0)
        ml["Season_Cos"] = np.cos(2 * np.pi * ml["Date_Start_TS"].dt.month.fillna(6) / 12.0)
        epoch = trip_df["Date_Start_TS"].min()
        ml["Days_Since_Epoch"] = (ml["Date_Start_TS"] - epoch).dt.total_seconds() / 86400.0

        features = ["Speed_kn", "Speed_Cubed", "True_Mass", "Kin_Delta", "Accel_Penalty", "Season_Sin", "Season_Cos", "Days_Since_Epoch"]
        maha_features = ["Speed_kn", "True_Mass", "Accel_Penalty", "Season_Sin", "Season_Cos", "Days_Since_Epoch"]
        ml[features] = ml[features].fillna(0.0)

        k_array = ml["Daily_Burn"] / ((ml["True_Mass"] ** (2 / 3)) * ml["Speed_Cubed"] + 1e-6)
        best_k = np.median(k_array[k_array <= np.percentile(k_array, 25)])
        ml["HM_Base"] = best_k * (ml["True_Mass"] ** (2 / 3)) * ml["Speed_Cubed"]
        trip_df.loc[sea_mask, "HM_Base"] = ml["HM_Base"]

        y_delta = ml["Daily_Burn"] - ml["HM_Base"]
        X_train, weights = ml[features], ml["Days"].clip(0.1, 30.0)
        if y_delta.var() < 0.05:
            raise ValueError("Target variance too low.")

        kf = KFold(n_splits=min(5, len(X_train)), shuffle=True, random_state=42)
        oof_preds = np.zeros(len(X_train))
        for train_idx, val_idx in kf.split(X_train):
            m = XGBRegressor(n_estimators=100, max_depth=3, learning_rate=0.06, random_state=42)
            m.fit(X_train.iloc[train_idx], y_delta.iloc[train_idx], sample_weight=weights.iloc[train_idx])
            oof_preds[val_idx] = m.predict(X_train.iloc[val_idx])

        oof_residuals = np.abs(y_delta - oof_preds)
        model = XGBRegressor(n_estimators=100, max_depth=3, learning_rate=0.06, random_state=42)
        model.fit(X_train, y_delta, sample_weight=weights)
        preds = ml["HM_Base"] + model.predict(X_train)

        var_model = XGBRegressor(n_estimators=40, max_depth=2, learning_rate=0.05, random_state=42)
        var_model.fit(X_train, oof_residuals, sample_weight=weights)
        var_preds_train = np.maximum(var_model.predict(X_train), 0.01)
        conformal_scores = oof_residuals / var_preds_train

        n = len(conformal_scores)
        q90 = np.quantile(conformal_scores, min(1.0, np.ceil((n + 1) * 0.90) / n) if n > 0 else 0.90)
        stoch_margin = np.maximum(var_model.predict(X_train) * q90, 0.5)

        p_vals = [
            (1.0 - (np.sum(conformal_scores <= (np.abs(ml.loc[idx, "Daily_Burn"] - preds.iloc[i]) / var_preds_train[i])) / len(conformal_scores))) * 100
            for i, idx in enumerate(ml.index)
        ]
        trip_df.loc[sea_mask, "P_Value"] = p_vals

        X_maha = ml[maha_features].values
        lw = LedoitWolf().fit(X_maha)
        md = np.sqrt(np.maximum(lw.mahalanobis(X_maha), 0))
        trip_df.loc[sea_mask, "Mahalanobis"] = md
        trip_df.loc[sea_mask, "MD_Threshold"] = np.percentile(md, 95)

        explainer = shap.TreeExplainer(model)
        sv = explainer.shap_values(X_train)
        base_val = explainer.expected_value[0] if isinstance(explainer.expected_value, np.ndarray) else explainer.expected_value

        trip_df.loc[sea_mask, "AI_Exp"] = preds.round(1)
        trip_df.loc[sea_mask, "Stoch_Var"] = stoch_margin.round(1)
        trip_df.loc[sea_mask, "SHAP_Base"] = base_val
        trip_df.loc[sea_mask, "SHAP_Propulsion"] = sv[:, 0] + sv[:, 1]
        trip_df.loc[sea_mask, "SHAP_Mass"] = sv[:, 2]
        trip_df.loc[sea_mask, "SHAP_Kinematics"] = sv[:, 3] + sv[:, 4]
        trip_df.loc[sea_mask, "SHAP_Season"] = sv[:, 5] + sv[:, 6]
        trip_df.loc[sea_mask, "SHAP_Degradation"] = sv[:, 7]
        trip_df.loc[sea_mask, "Exp_Lower"] = preds - stoch_margin
        trip_df.loc[sea_mask, "Exp_Upper"] = preds + stoch_margin

        outlier_mask = sea_mask & (
            (trip_df["Daily_Burn"] < trip_df["Exp_Lower"]) |
            (trip_df["Daily_Burn"] > trip_df["Exp_Upper"])
        )
        trip_df.loc[outlier_mask, "Status"] = "STAT OUTLIER"

    except ValueError as e:
        ai_status_msg = f"AI Offline: {str(e)}"
    except Exception as e:
        ai_status_msg = f"AI Exception: {str(e)}"
        print(traceback.format_exc())

    return trip_df, ai_status_msg
    # ═══════════════════════════════════════════════════════════════════════════════
# PIPELINE ORCHESTRATOR
# ═══════════════════════════════════════════════════════════════════════════════
@st.cache_data(show_spinner=False)
def run_pipeline(file_bytes, filename, min_speed, ghost_sea, ghost_port):
    try:
        parsed_df, vname = semantic_parse(file_bytes, filename)
        trip_df, cum_drift = build_state_machine(parsed_df, min_speed, ghost_sea, ghost_port)
        trip_df, ai_msg = execute_ai_physics(trip_df, min_speed)

        quarantined = len(trip_df[trip_df["Status"].str.contains("QUARANTINE")])
        valid_sea = trip_df[(trip_df["Phase"] == "SEA") & (trip_df["Status"] == "VERIFIED")]
        avg_sea = valid_sea["Phys_Burn"].sum() / valid_sea["Days"].sum() if valid_sea["Days"].sum() > 0 else 0.0

        trip_df["Total_CYLO"] = (
            trip_df.get("HSCYLO_L", pd.Series([0], dtype=float)) +
            trip_df.get("LSCYLO_L", pd.Series([0], dtype=float)) +
            trip_df.get("CYLO_GEN_L", pd.Series([0], dtype=float))
        )

        summary = {
            "vname": vname,
            "integrity": round((len(trip_df) - quarantined) / len(trip_df) * 100, 1) if not trip_df.empty else 0,
            "avg_dqi": round(trip_df["DQI"].mean(), 0) if not trip_df.empty else 0,
            "total_fuel": round(trip_df["Phys_Burn"].sum(skipna=True), 1),
            "avg_sea_burn": round(avg_sea, 1),
            "total_nm": round(trip_df["Dist_NM"].sum(), 0),
            "total_days": round(trip_df["Days"].sum(), 1),
            "total_melo": round(trip_df.get("MELO_L", pd.Series([0])).sum(), 0),
            "total_cylo": round(trip_df["Total_CYLO"].sum(), 0),
            "cycles": len(trip_df),
            "quarantined": quarantined,
            "anomalies": len(trip_df[trip_df["Status"].isin(["GHOST BUNKER", "STAT OUTLIER"])]),
            "ai_msg": ai_msg
        }
        return trip_df, summary, cum_drift, None
    except ValueError as e:
        return pd.DataFrame(), None, None, f"Parsing Rejected: {str(e)}"
    except Exception as e:
        return pd.DataFrame(), None, None, f"System Crash: {str(e)}"

# ═══════════════════════════════════════════════════════════════════════════════
# PLOTLY RENDER ENGINE
# ═══════════════════════════════════════════════════════════════════════════════
_BL = dict(
    paper_bgcolor="rgba(0,0,0,0)",
    plot_bgcolor="rgba(0,0,0,0)",
    hovermode="x unified",
    hoverlabel=dict(
        bgcolor="rgba(6,12,18,0.97)",
        bordercolor="rgba(0,224,176,0.55)",
        font=dict(family="Geist Mono", color="#f8fafc", size=13)
    ),
    font=dict(family="Hanken Grotesk", color="#f8fafc"),
    transition=dict(duration=800, easing="cubic-in-out")
)

_M = dict(l=15, r=15, t=85, b=30)
_AX = dict(
    gridcolor="rgba(255,255,255,0.02)",
    zerolinecolor="rgba(255,255,255,0.05)",
    tickfont=dict(family="Geist Mono", size=11, color="#475569"),
    showspikes=True,
    spikecolor="rgba(0,224,176,0.6)",
    spikethickness=1,
    spikedash="solid"
)

def chart_fuel(df):
    sea = df[(df["Phase"] == "SEA") & (~df["Status"].str.contains("QUARANTINE"))]
    port = df[(df["Phase"] == "PORT") & (~df["Status"].str.contains("QUARANTINE"))]

    fig = make_subplots(
        rows=2,
        cols=1,
        shared_xaxes=True,
        row_heights=[0.7, 0.3],
        vertical_spacing=0.08
    )

    if not sea.empty:
        fig.add_trace(
            go.Bar(
                x=sea["Timeline"],
                y=sea["Phys_Burn"],
                name="Sea Fuel",
                marker_color="rgba(0,224,176,0.15)",
                marker_line_color="#00e0b0",
                marker_line_width=1.5
            ),
            row=1,
            col=1
        )

        fig.add_trace(
            go.Scatter(
                x=sea["Timeline"],
                y=sea["Daily_Burn"],
                name="Sea MT/day",
                mode="lines+markers",
                line=dict(color="#00e0b0", width=3, shape="spline"),
                fill="tozeroy",
                fillcolor="rgba(0,224,176,0.05)",
                marker=dict(size=8, color="#051014", line=dict(color="#00e0b0", width=2))
            ),
            row=1,
            col=1
        )

        fig.add_trace(
            go.Scatter(
                x=sea["Timeline"],
                y=sea["Speed_kn"],
                name="Sea Speed",
                mode="lines+markers",
                line=dict(color="#c9a84c", width=3, shape="spline"),
                fill="tozeroy",
                fillcolor="rgba(201,168,76,0.05)",
                marker=dict(size=8, color="#051014", line=dict(color="#c9a84c", width=2))
            ),
            row=2,
            col=1
        )

    if not port.empty:
        fig.add_trace(
            go.Bar(
                x=port["Timeline"],
                y=port["Phys_Burn"],
                name="Port Fuel",
                marker_color="rgba(255,42,85,0.15)",
                marker_line_color="#ff2a55",
                marker_line_width=1.5
            ),
            row=1,
            col=1
        )

    fig.update_layout(
        **_BL,
        margin=_M,
        title=dict(
            text="Tri-State Fuel Consumption & Kinematics",
            font=dict(size=24, family="Bricolage Grotesque", color="#fff")
        ),
        barmode="group",
        showlegend=True,
        height=700,
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
    )
    fig.update_xaxes(tickangle=-45, automargin=True, **_AX)
    fig.update_yaxes(**_AX)
    return fig

def chart_lube(df):
    fig = go.Figure()

    if df.get("MELO_L", pd.Series([0])).sum() > 0:
        fig.add_trace(
            go.Bar(
                x=df["Timeline"], y=df["MELO_L"], 
