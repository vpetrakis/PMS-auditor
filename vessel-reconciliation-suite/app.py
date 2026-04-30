import io
import math
from dataclasses import dataclass
from typing import Dict, List, Optional, Tuple

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
import streamlit as st

try:
    from sklearn.covariance import LedoitWolf
    SKLEARN_OK = True
except Exception:
    SKLEARN_OK = False

st.set_page_config(
    page_title="Oracle PMS Intelligence",
    page_icon="◈",
    layout="wide",
    initial_sidebar_state="expanded",
)

MONTH_MAP = {
    "Jan": 1, "Feb": 2, "Mar": 3, "Apr": 4, "May": 5, "Jun": 6,
    "Jul": 7, "Aug": 8, "Sep": 9, "Oct": 10, "Nov": 11, "Dec": 12
}
MONTHS = list(MONTH_MAP.keys())
ENGINE_COLS = [
    "Main Engine",
    "Diesel Generator No.1",
    "Diesel Generator No.2",
    "Diesel Generator No.3",
    "Diesel Generator No.4",
]

DEFAULT_CONFIG = pd.DataFrame([
    {"pattern": "CYLINDER COVER", "reset_mode": "OH", "family": "Cylinder Cover", "criticality": "High"},
    {"pattern": "CYLINDER COVER COOLING JACKET", "reset_mode": "OH", "family": "Cover Jacket", "criticality": "Medium"},
    {"pattern": "CYL. LINER", "reset_mode": "RENEWAL_OR_OBSERVATION", "family": "Cylinder Liner", "criticality": "High"},
    {"pattern": "CYL LINER COOLING JACKET", "reset_mode": "OH", "family": "Liner Jacket", "criticality": "Medium"},
    {"pattern": "PISTON ASSY", "reset_mode": "OH", "family": "Piston Assembly", "criticality": "High"},
    {"pattern": "PISTON CROWN", "reset_mode": "RENEWAL", "family": "Piston Crown", "criticality": "High"},
    {"pattern": "STUFFING BOX", "reset_mode": "OH", "family": "Stuffing Box", "criticality": "Medium"},
    {"pattern": "EXHAUST VALVE", "reset_mode": "OH", "family": "Exhaust Valve", "criticality": "High"},
    {"pattern": "STARTING VALVE", "reset_mode": "OH", "family": "Starting Valve", "criticality": "Medium"},
    {"pattern": "SAFETY VALVE", "reset_mode": "OH", "family": "Safety Valve", "criticality": "Medium"},
    {"pattern": "FUEL VALVE", "reset_mode": "OH", "family": "Fuel Valve", "criticality": "High"},
    {"pattern": "FUEL PUMP", "reset_mode": "OH", "family": "Fuel Pump", "criticality": "High"},
    {"pattern": "PLUNGER BARREL", "reset_mode": "RENEWAL", "family": "Plunger Barrel", "criticality": "High"},
    {"pattern": "SUCTION VALVE", "reset_mode": "OH", "family": "Suction Valve", "criticality": "Medium"},
    {"pattern": "PUNCTURE VALVE", "reset_mode": "OH", "family": "Puncture Valve", "criticality": "Medium"},
    {"pattern": "BEARING CLEARANCES", "reset_mode": "CALENDAR", "family": "Bearing Clearance", "criticality": "Medium"},
    {"pattern": "GUIDE SHOE", "reset_mode": "CALENDAR", "family": "Guide Shoe", "criticality": "Medium"},
    {"pattern": "AIR COOLER", "reset_mode": "OH", "family": "Air Cooler", "criticality": "High"},
    {"pattern": "L.O. COOLER", "reset_mode": "OH", "family": "LO Cooler", "criticality": "Medium"},
])


@dataclass
class AuditResult:
    daily_df: pd.DataFrame
    monthly_df: pd.DataFrame
    pms_df: pd.DataFrame
    oracle_df: pd.DataFrame
    issues_df: pd.DataFrame
    meta: Dict[str, object]


CSS = """
<style>
:root {
  --bg:#07111a;
  --bg2:#0c1722;
  --card:rgba(14,24,35,0.72);
  --card2:rgba(18,30,44,0.84);
  --line:rgba(255,255,255,0.08);
  --text:#eef4ff;
  --muted:#90a4bd;
  --teal:#73f5d0;
  --gold:#d6b36a;
  --rose:#ff5f86;
  --vio:#9f8cff;
}
.stApp {
  background:
    radial-gradient(circle at 10% 0%, rgba(115,245,208,0.10), transparent 30%),
    radial-gradient(circle at 100% 0%, rgba(214,179,106,0.10), transparent 28%),
    linear-gradient(180deg, #07111a 0%, #08131d 35%, #0a1520 100%);
  color:var(--text);
}
.block-container {padding-top:1.2rem;padding-bottom:1.5rem;max-width: 1500px;}
[data-testid="stSidebar"] {background:linear-gradient(180deg,#09131d 0%, #0d1723 100%); border-right:1px solid var(--line);}
h1,h2,h3 {color:var(--text); letter-spacing:-0.02em;}
.hero {
  position:relative; overflow:hidden; border:1px solid var(--line); border-radius:28px;
  background:linear-gradient(135deg, rgba(13,22,34,0.92), rgba(8,17,26,0.78));
  padding:28px 30px 24px 30px; margin-bottom:20px; box-shadow:0 20px 60px rgba(0,0,0,0.28);
}
.hero:before {
  content:""; position:absolute; inset:-2px; background:linear-gradient(120deg, rgba(115,245,208,0.18), transparent 30%, rgba(214,179,106,0.18), transparent 70%, rgba(159,140,255,0.18));
  filter:blur(24px); z-index:0; animation:drift 10s ease-in-out infinite alternate;
}
@keyframes drift { from {transform:translateX(-2%) translateY(-1%);} to {transform:translateX(2%) translateY(1%);} }
.hero-inner {position:relative; z-index:1; display:flex; justify-content:space-between; gap:18px; align-items:flex-end; flex-wrap:wrap;}
.kicker {font-size:0.73rem; letter-spacing:0.22em; text-transform:uppercase; color:var(--gold); margin-bottom:10px;}
.title {font-size:2.5rem; font-weight:800; line-height:1.02; color:var(--text);}
.subtitle {font-size:1rem; color:var(--muted); max-width:820px; margin-top:10px;}
.pillbar {display:flex; gap:10px; flex-wrap:wrap; margin-top:16px;}
.pill {padding:8px 12px; border-radius:999px; border:1px solid var(--line); background:rgba(255,255,255,0.03); color:var(--text); font-size:0.78rem;}
.metric-grid {display:grid; grid-template-columns:repeat(4,minmax(0,1fr)); gap:14px; margin:18px 0 24px;}
.metric-card {
  background:linear-gradient(180deg, rgba(18,30,44,0.86), rgba(13,22,32,0.90));
  border:1px solid var(--line); border-radius:22px; padding:18px 18px 14px;
  box-shadow:0 14px 30px rgba(0,0,0,0.18); position:relative; overflow:hidden;
}
.metric-card:after {content:""; position:absolute; left:0; top:0; width:100%; height:1px; background:linear-gradient(90deg, transparent, rgba(255,255,255,0.35), transparent);}
.metric-label {font-size:0.8rem; letter-spacing:0.08em; text-transform:uppercase; color:var(--muted);}
.metric-value {font-size:2rem; font-weight:800; color:var(--text); margin-top:6px;}
.metric-sub {font-size:0.86rem; color:var(--muted); margin-top:4px;}
.glass {
  background:var(--card); border:1px solid var(--line); border-radius:22px; padding:16px 18px;
}
.status-ok {color:var(--teal);} .status-warn {color:var(--gold);} .status-bad {color:var(--rose);} .status-mid {color:var(--vio);}
.smallnote {font-size:0.9rem;color:var(--muted)}
.section-head {font-size:1rem; text-transform:uppercase; letter-spacing:0.18em; color:var(--gold); margin:10px 0 12px;}
.stTabs [data-baseweb="tab-list"] {gap:10px;}
.stTabs [data-baseweb="tab"] {background:rgba(255,255,255,0.03); border:1px solid var(--line); border-radius:14px; padding:10px 16px;}
.stTabs [aria-selected="true"] {background:rgba(115,245,208,0.08)!important; border-color:rgba(115,245,208,0.30)!important;}
</style>
"""

st.markdown(CSS, unsafe_allow_html=True)


def normalize_text(v: object) -> str:
    if pd.isna(v):
        return ""
    return " ".join(str(v).replace("\n", " ").replace("\r", " ").split()).strip()


def try_parse_date(v: object) -> pd.Timestamp:
    if pd.isna(v) or normalize_text(v) == "":
        return pd.NaT
    try:
        ts = pd.to_datetime(v, errors="coerce", dayfirst=True)
        return ts.normalize() if pd.notna(ts) else pd.NaT
    except Exception:
        return pd.NaT


def find_sheet_name(excel_file, candidates: List[str]) -> Optional[str]:
    xl = pd.ExcelFile(excel_file)
    names = xl.sheet_names
    for cand in candidates:
        for real in names:
            if cand.lower() in real.lower():
                return real
    return names[0] if names else None


def infer_engine_from_code(code: str) -> str:
    c = normalize_text(code).upper()
    if c.startswith("ME"):
        return "Main Engine"
    if c.startswith("DG-01") or c.startswith("DG1"):
        return "Diesel Generator No.1"
    if c.startswith("DG-02") or c.startswith("DG2"):
        return "Diesel Generator No.2"
    if c.startswith("DG-03") or c.startswith("DG3"):
        return "Diesel Generator No.3"
    if c.startswith("DG-04") or c.startswith("DG4"):
        return "Diesel Generator No.4"
    return "Main Engine"


def family_from_code(code: str) -> str:
    c = normalize_text(code).upper()
    if c.startswith("ME-") and len(c) >= 5:
        parts = c.split("-")
        return parts[0] + "-" + parts[1] if len(parts) > 1 else c
    if c.startswith("DG-"):
        parts = c.split("-")
        return parts[0] + "-" + parts[1] if len(parts) > 1 else c
    return c[:5]


def parse_daily_sheet(uploaded_file) -> Tuple[pd.DataFrame, pd.DataFrame, Dict[str, object]]:
    sheet_name = find_sheet_name(uploaded_file, ["daily operating hours", "daily"])
    raw = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=None)

    year = pd.Timestamp.today().year
    for _, row in raw.iterrows():
        vals = [normalize_text(v) for v in row.tolist()]
        if "Current Year" in vals:
            idx = vals.index("Current Year")
            y = pd.to_numeric(row.iloc[idx + 1] if idx + 1 < len(row) else np.nan, errors="coerce")
            if pd.notna(y):
                year = int(y)
                break

    records = []
    for _, row in raw.iterrows():
        c0 = row.iloc[0] if len(row) > 0 else None
        is_dt = isinstance(c0, pd.Timestamp)
        looks_dt = isinstance(c0, str) and any(ch.isdigit() for ch in c0)
        if not (is_dt or looks_dt):
            continue
        dt = try_parse_date(c0)
        if pd.isna(dt):
            continue
        vals = list(row.values)
        records.append({
            "date": dt,
            "month": dt.strftime("%b"),
            "Main Engine": float(pd.to_numeric(vals[1] if len(vals) > 1 else np.nan, errors="coerce") or 0),
            "Diesel Generator No.1": float(pd.to_numeric(vals[2] if len(vals) > 2 else np.nan, errors="coerce") or 0),
            "Diesel Generator No.2": float(pd.to_numeric(vals[4] if len(vals) > 4 else np.nan, errors="coerce") or 0),
            "Diesel Generator No.3": float(pd.to_numeric(vals[6] if len(vals) > 6 else np.nan, errors="coerce") or 0),
            "Diesel Generator No.4": float(pd.to_numeric(vals[8] if len(vals) > 8 else np.nan, errors="coerce") or 0),
        })

    daily_df = pd.DataFrame(records).sort_values("date").reset_index(drop=True)
    monthly_df = pd.DataFrame(columns=["month_num", "month"] + ENGINE_COLS)
    if not daily_df.empty:
        monthly_df = (
            daily_df.assign(month_num=daily_df["date"].dt.month)
            .groupby(["month_num", "month"], as_index=False)[ENGINE_COLS].sum()
            .sort_values("month_num")
            .reset_index(drop=True)
        )
    return daily_df, monthly_df, {"year": year, "daily_sheet": sheet_name}


def parse_pms_sheet(uploaded_file) -> Tuple[pd.DataFrame, Dict[str, object]]:
    sheet_name = find_sheet_name(uploaded_file, ["pms"])
    df = pd.read_excel(uploaded_file, sheet_name=sheet_name)
    cols = {str(c).lower().strip(): c for c in df.columns}

    def pick(*names):
        for n in names:
            for low, real in cols.items():
                if n in low:
                    return real
        return None

    code_col = pick("code no")
    item_col = pick("items")
    interval_col = pick("interval")
    last_date_col = pick("date of last inspection")
    last_year_col = pick("operating hours at the end of last year")
    current_col = pick("current operating hours")
    etd_col = pick("estimated date of next inspection")

    out = pd.DataFrame({
        "code": df[code_col] if code_col else "",
        "item": df[item_col] if item_col else "",
        "interval": df[interval_col] if interval_col else "",
        "last_inspection_date": df[last_date_col] if last_date_col else pd.NaT,
        "hours_end_last_year": df[last_year_col] if last_year_col else np.nan,
        "current_operating_hours": df[current_col] if current_col else np.nan,
        "estimated_next_date": df[etd_col] if etd_col else pd.NaT,
    })

    for m in MONTHS:
        matched = None
        for c in df.columns:
            if normalize_text(c).startswith(m):
                matched = c
                break
        out[m] = df[matched] if matched is not None else np.nan

    out["code"] = out["code"].map(normalize_text)
    out["item"] = out["item"].map(normalize_text)
    out["interval"] = out["interval"].map(normalize_text)
    out["last_inspection_date"] = out["last_inspection_date"].map(try_parse_date)
    out["estimated_next_date"] = out["estimated_next_date"].map(try_parse_date)
    out["hours_end_last_year"] = pd.to_numeric(out["hours_end_last_year"], errors="coerce")
    out["current_operating_hours"] = pd.to_numeric(out["current_operating_hours"], errors="coerce")
    for m in MONTHS:
        out[m] = pd.to_numeric(out[m], errors="coerce")

    out = out[(out["code"] != "") | (out["item"] != "")].copy()
    out["engine_group"] = out["code"].map(infer_engine_from_code)
    out["asset_family"] = out["code"].map(family_from_code)
    return out.reset_index(drop=True), {"pms_sheet": sheet_name}


def resolve_rule(item: str, config_df: pd.DataFrame) -> Optional[pd.Series]:
    item_u = normalize_text(item).upper()
    hits = []
    for _, row in config_df.iterrows():
        patt = normalize_text(row.get("pattern", "")).upper()
        if patt and patt in item_u:
            hits.append((len(patt), row))
    if not hits:
        return None
    hits.sort(key=lambda x: x[0], reverse=True)
    return hits[0][1]


def monthly_hours_since_date(daily_df: pd.DataFrame, engine_name: str, start_date: pd.Timestamp) -> Tuple[float, Dict[str, float]]:
    if engine_name not in daily_df.columns or pd.isna(start_date):
        return np.nan, {m: np.nan for m in MONTHS}
    filt = daily_df[daily_df["date"] > start_date]
    if filt.empty:
        return 0.0, {m: 0.0 for m in MONTHS}
    per_month = filt.groupby(filt["date"].dt.strftime("%b"))[engine_name].sum().to_dict()
    total = float(filt[engine_name].sum())
    return total, {m: float(per_month.get(m, 0.0)) for m in MONTHS}


def monthly_hours_from_year_start(daily_df: pd.DataFrame, engine_name: str) -> Tuple[float, Dict[str, float]]:
    if engine_name not in daily_df.columns:
        return np.nan, {m: np.nan for m in MONTHS}
    per_month = daily_df.groupby(daily_df["date"].dt.strftime("%b"))[engine_name].sum().to_dict()
    monthly = {m: float(per_month.get(m, 0.0)) for m in MONTHS}
    return float(sum(monthly.values())), monthly


def infer_oracle(pms_df: pd.DataFrame, daily_df: pd.DataFrame, config_df: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for _, r in pms_df.iterrows():
        if pd.isna(r.get("current_operating_hours")):
            continue

        rule = resolve_rule(r.get("item", ""), config_df)
        mode = rule.get("reset_mode", "OH") if rule is not None else "OH"
        family = rule.get("family", "Unmapped") if rule is not None else "Unmapped"
        criticality = rule.get("criticality", "Medium") if rule is not None else "Medium"
        engine = r.get("engine_group", "Main Engine")

        parent_total, parent_months = monthly_hours_from_year_start(daily_df, engine)
        post_event_total, post_event_months = monthly_hours_since_date(daily_df, engine, r.get("last_inspection_date"))

        if mode in ["OH", "RENEWAL", "RENEWAL_OR_OBSERVATION"] and pd.notna(r.get("last_inspection_date")):
            baseline = post_event_total
            month_map = post_event_months
            basis = f"post_{mode.lower()}"
        elif mode == "CALENDAR":
            baseline = parent_total
            month_map = {m: np.nan for m in MONTHS}
            basis = "calendar_control"
        else:
            baseline = parent_total
            month_map = parent_months
            basis = "parent_ytd"

        reported = float(r.get("current_operating_hours")) if pd.notna(r.get("current_operating_hours")) else np.nan
        engineering_gap = np.nan if pd.isna(baseline) or pd.isna(reported) else reported - baseline

        blended = baseline if pd.notna(baseline) else reported
        if pd.notna(baseline) and pd.notna(reported):
            blended = 0.7 * baseline + 0.3 * reported
            if abs(engineering_gap) <= 2:
                blended = baseline

        due_ratio = np.nan
        interval_text = normalize_text(r.get("interval", ""))
        interval_num = pd.to_numeric(interval_text.split(",")[0], errors="coerce") if interval_text else np.nan
        if pd.notna(interval_num) and pd.notna(blended) and mode != "CALENDAR":
            due_ratio = blended / interval_num if interval_num > 0 else np.nan

        monthly_delta_l1 = 0.0
        monthly_known = 0
        for m in MONTHS:
            a = pd.to_numeric(r.get(m), errors="coerce")
            e = month_map.get(m, np.nan)
            if pd.notna(a) and pd.notna(e):
                monthly_delta_l1 += abs(a - e)
                monthly_known += 1

        rows.append({
            "code": r.get("code", ""),
            "item": r.get("item", ""),
            "engine_group": engine,
            "asset_family": r.get("asset_family", ""),
            "rule_family": family,
            "criticality": criticality,
            "reset_mode": mode,
            "basis": basis,
            "last_inspection_date": r.get("last_inspection_date"),
            "estimated_next_date": r.get("estimated_next_date"),
            "reported_current_hours": reported,
            "engineering_baseline": baseline,
            "oracle_hours": blended,
            "engineering_gap": engineering_gap,
            "parent_ytd_hours": parent_total,
            "post_event_hours": post_event_total,
            "interval_hours": interval_num,
            "due_ratio": due_ratio,
            "monthly_delta_l1": monthly_delta_l1,
            "monthly_known": monthly_known,
            **{f"reported_{m}": pd.to_numeric(r.get(m), errors="coerce") for m in MONTHS},
            **{f"oracle_{m}": month_map.get(m, np.nan) for m in MONTHS},
        })

    out = pd.DataFrame(rows)
    if out.empty:
        return out

    out["family_median_gap"] = out.groupby(["engine_group", "rule_family"])["engineering_gap"].transform("median")
    out["family_gap_dev"] = (out["engineering_gap"] - out["family_median_gap"]).abs()

    feat = out[["reported_current_hours", "engineering_baseline", "engineering_gap", "monthly_delta_l1", "due_ratio", "family_gap_dev"]].copy()
    feat = feat.replace([np.inf, -np.inf], np.nan).fillna(feat.median(numeric_only=True)).fillna(0)

    if len(feat) >= 4:
        X = feat.values.astype(float)
        if SKLEARN_OK:
            lw = LedoitWolf().fit(X)
            md = np.sqrt(np.maximum(lw.mahalanobis(X), 0))
        else:
            z = (feat - feat.mean()) / feat.std(ddof=0).replace(0, 1)
            md = np.sqrt((z ** 2).sum(axis=1)).values
        out["anomaly_score"] = md
        out["anomaly_threshold"] = np.percentile(md, 90)
    else:
        out["anomaly_score"] = 0.0
        out["anomaly_threshold"] = 9.99

    gap_scale = max(float(np.nanmedian(np.abs(out["engineering_gap"].fillna(0)))), 1.0)
    monthly_scale = max(float(np.nanmedian(out["monthly_delta_l1"].fillna(0))), 1.0)
    anom_scale = max(float(np.nanmedian(out["anomaly_score"].fillna(0))), 1.0)

    out["trust_observation"] = (
        100
        - np.clip(np.abs(out["engineering_gap"].fillna(0)) / gap_scale * 18, 0, 60)
        - np.clip(out["monthly_delta_l1"].fillna(0) / monthly_scale * 10, 0, 28)
    )
    out["trust_inference"] = (
        100
        - np.clip(out["anomaly_score"].fillna(0) / max(anom_scale, 0.5) * 16, 0, 35)
        - np.where(out["reset_mode"].eq("CALENDAR"), 18, 0)
    )
    out["trust_decision"] = (0.45 * out["trust_observation"] + 0.55 * out["trust_inference"]).clip(0, 100)

    out["uncertainty_hours"] = (
        np.abs(out["engineering_gap"].fillna(0)) * 0.35
        + out["monthly_delta_l1"].fillna(0) * 0.10
        + np.maximum(out["anomaly_score"].fillna(0) - out["anomaly_threshold"].fillna(0), 0) * 4
    ).clip(lower=4)

    out["oracle_low"] = (out["oracle_hours"] - out["uncertainty_hours"]).clip(lower=0)
    out["oracle_high"] = out["oracle_hours"] + out["uncertainty_hours"]

    conditions = [
        out["trust_decision"] >= 82,
        (out["trust_decision"] >= 58) & (out["trust_decision"] < 82),
        out["trust_decision"] < 58,
    ]
    labels = ["Verified", "Adjusted", "Quarantined"]
    out["oracle_status"] = np.select(conditions, labels, default="Adjusted")

    out["decision_action"] = np.select(
        [
            (out["due_ratio"] >= 1.0) & (out["trust_decision"] >= 58),
            (out["due_ratio"] >= 0.85) & (out["trust_decision"] >= 58),
            out["trust_decision"] < 58,
        ],
        ["Requisition now", "Prepare requisition", "Escalate review"],
        default="Monitor",
    )

    return out.sort_values(["oracle_status", "trust_decision", "anomaly_score"], ascending=[True, False, False]).reset_index(drop=True)


def summarize_issues(oracle_df: pd.DataFrame) -> pd.DataFrame:
    issues = []
    for _, r in oracle_df.iterrows():
        if abs(r.get("engineering_gap", 0)) >= 0.5:
            sev = "high" if abs(r.get("engineering_gap", 0)) >= 24.5 else "medium"
            issues.append({
                "severity": sev,
                "code": r["code"],
                "item": r["item"],
                "issue": "Reported vs engineering baseline disagreement",
                "detail": f"reported={r['reported_current_hours']:.0f}, baseline={r['engineering_baseline']:.0f}, gap={r['engineering_gap']:.0f}",
            })
        if r.get("oracle_status") == "Quarantined":
            issues.append({
                "severity": "high",
                "code": r["code"],
                "item": r["item"],
                "issue": "Low decision confidence",
                "detail": f"trust={r['trust_decision']:.0f}, anomaly={r['anomaly_score']:.2f}",
            })

    out = pd.DataFrame(issues)
    if not out.empty:
        out["severity"] = pd.Categorical(out["severity"], ["high", "medium", "low"], ordered=True)
        out = out.sort_values(["severity", "code"]).reset_index(drop=True)
    return out


def export_audit_excel(result: AuditResult) -> bytes:
    bio = io.BytesIO()
    with pd.ExcelWriter(bio, engine="xlsxwriter") as writer:
        result.daily_df.to_excel(writer, sheet_name="daily_hours", index=False)
        result.monthly_df.to_excel(writer, sheet_name="monthly_hours", index=False)
        result.pms_df.to_excel(writer, sheet_name="pms_clean", index=False)
        result.oracle_df.to_excel(writer, sheet_name="oracle_engine", index=False)
        result.issues_df.to_excel(writer, sheet_name="issues", index=False)
        pd.DataFrame([result.meta]).to_excel(writer, sheet_name="meta", index=False)
    bio.seek(0)
    return bio.getvalue()


def build_result(uploaded_file, config_df: pd.DataFrame) -> AuditResult:
    daily_df, monthly_df, meta1 = parse_daily_sheet(uploaded_file)
    pms_df, meta2 = parse_pms_sheet(uploaded_file)
    oracle_df = infer_oracle(pms_df, daily_df, config_df)
    issues_df = summarize_issues(oracle_df)
    return AuditResult(daily_df, monthly_df, pms_df, oracle_df, issues_df, {**meta1, **meta2})


def render_hero():
    st.markdown(
        """
        <div class='hero'>
          <div class='hero-inner'>
            <div>
              <div class='kicker'>Oracle • PMS Intelligence Kernel</div>
              <div class='title'>Maintenance Oracle for Running-Hours Truth</div>
              <div class='subtitle'>A forensic-grade Streamlit application for reconciling PMS declarations, parent-engine operating hours, reset events, uncertainty bands, and anomaly geometry in one premium operational workspace.</div>
              <div class='pillbar'>
                <div class='pill'>Deterministic engineering baseline</div>
                <div class='pill'>Bayesian-style uncertainty bands</div>
                <div class='pill'>Ledoit-Wolf anomaly geometry</div>
                <div class='pill'>Premium review UX for GitHub/Streamlit</div>
              </div>
            </div>
          </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_metrics(oracle_df: pd.DataFrame):
    if oracle_df.empty:
        return
    verified = int((oracle_df["oracle_status"] == "Verified").sum())
    adjusted = int((oracle_df["oracle_status"] == "Adjusted").sum())
    quarantined = int((oracle_df["oracle_status"] == "Quarantined").sum())
    avg_trust = oracle_df["trust_decision"].mean()

    st.markdown(
        f"""
        <div class='metric-grid'>
          <div class='metric-card'><div class='metric-label'>Verified components</div><div class='metric-value status-ok'>{verified}</div><div class='metric-sub'>Directly trusted by the oracle</div></div>
          <div class='metric-card'><div class='metric-label'>Adjusted components</div><div class='metric-value status-mid'>{adjusted}</div><div class='metric-sub'>Inferred through baseline + reconciliation</div></div>
          <div class='metric-card'><div class='metric-label'>Quarantined components</div><div class='metric-value status-bad'>{quarantined}</div><div class='metric-sub'>Require engineering review</div></div>
          <div class='metric-card'><div class='metric-label'>Decision trust</div><div class='metric-value'>{avg_trust:.0f}</div><div class='metric-sub'>Average oracle confidence</div></div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def chart_monthly(monthly_df: pd.DataFrame):
    plot_df = monthly_df.melt(id_vars=["month_num", "month"], value_vars=ENGINE_COLS, var_name="engine", value_name="hours")
    fig = px.bar(plot_df, x="month", y="hours", color="engine", barmode="group")
    fig.update_layout(
        height=460,
        margin=dict(l=10, r=10, t=60, b=20),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(color="#eef4ff"),
    )
    return fig


def chart_family_heatmap(oracle_df: pd.DataFrame):
    tmp = oracle_df.copy()
    tmp["family_index"] = tmp["code"].astype(str).str.extract(r"-(\d{2})-")
    tmp["family_index"] = tmp["family_index"].fillna(tmp["code"])
    pivot = tmp.pivot_table(index="rule_family", columns="family_index", values="oracle_hours", aggfunc="mean")
    fig = go.Figure(
        data=go.Heatmap(
            z=pivot.values,
            x=pivot.columns.astype(str),
            y=pivot.index.astype(str),
            colorscale=[[0, "#0c1722"], [0.35, "#1a3240"], [0.7, "#2f807d"], [1, "#d6b36a"]],
            colorbar=dict(title="hrs"),
        )
    )
    fig.update_layout(
        height=520,
        margin=dict(l=10, r=10, t=50, b=20),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(color="#eef4ff"),
        title="Oracle family coherence map",
    )
    return fig


def chart_component_truth(oracle_df: pd.DataFrame, selected_code: str):
    row = oracle_df.loc[oracle_df["code"] == selected_code].iloc[0]
    fig = go.Figure()
    fig.add_trace(go.Scatter(
        x=[row["oracle_low"], row["oracle_high"]],
        y=["Oracle interval", "Oracle interval"],
        mode="lines",
        line=dict(color="#9f8cff", width=12),
        name="90% interval",
    ))
    fig.add_trace(go.Scatter(
        x=[row["oracle_hours"]],
        y=["Oracle interval"],
        mode="markers",
        marker=dict(size=16, color="#73f5d0"),
        name="Oracle mean",
    ))
    fig.add_trace(go.Scatter(
        x=[row["engineering_baseline"]],
        y=["Oracle interval"],
        mode="markers",
        marker=dict(size=14, color="#d6b36a", symbol="diamond"),
        name="Engineering baseline",
    ))
    fig.add_trace(go.Scatter(
        x=[row["reported_current_hours"]],
        y=["Oracle interval"],
        mode="markers",
        marker=dict(size=14, color="#ff5f86", symbol="x"),
        name="Reported",
    ))
    if pd.notna(row["interval_hours"]):
        fig.add_vline(x=row["interval_hours"], line_color="rgba(255,255,255,0.40)", line_dash="dash")
    fig.update_layout(
        height=250,
        margin=dict(l=10, r=10, t=50, b=20),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(color="#eef4ff"),
        title=f"Truth strip · {selected_code}",
    )
    return fig


def chart_waterfall(oracle_df: pd.DataFrame, selected_code: str):
    r = oracle_df.loc[oracle_df["code"] == selected_code].iloc[0]
    baseline = r["engineering_baseline"] if pd.notna(r["engineering_baseline"]) else 0
    gap = (r["reported_current_hours"] - baseline) if pd.notna(r["reported_current_hours"]) and pd.notna(baseline) else 0
    family_adj = -0.25 * r["family_gap_dev"] if pd.notna(r["family_gap_dev"]) else 0
    anomaly_adj = -0.20 * max(r["anomaly_score"] - r["anomaly_threshold"], 0) * 6 if pd.notna(r["anomaly_score"]) else 0
    target = r["oracle_hours"] if pd.notna(r["oracle_hours"]) else baseline
    residual = target - (baseline + gap + family_adj + anomaly_adj)

    fig = go.Figure(go.Waterfall(
        orientation="v",
        measure=["absolute", "relative", "relative", "relative", "relative", "total"],
        x=["Baseline", "Report shift", "Family adj.", "Anomaly adj.", "Residual", "Oracle"],
        y=[baseline, gap, family_adj, anomaly_adj, residual, 0],
        connector={"line": {"color": "rgba(255,255,255,0.18)", "dash": "dot"}},
        increasing={"marker": {"color": "#73f5d0"}},
        decreasing={"marker": {"color": "#ff5f86"}},
        totals={"marker": {"color": "#9f8cff"}},
    ))
    fig.update_layout(
        height=420,
        margin=dict(l=10, r=10, t=50, b=10),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(color="#eef4ff"),
        title="Evidence waterfall",
    )
    return fig


def render_overview(result: AuditResult):
    render_metrics(result.oracle_df)
    c1, c2 = st.columns([1.1, 0.9])
    with c1:
        st.plotly_chart(chart_monthly(result.monthly_df), use_container_width=True, config={"displayModeBar": False})
    with c2:
        st.plotly_chart(chart_family_heatmap(result.oracle_df), use_container_width=True, config={"displayModeBar": False})

    st.markdown("<div class='section-head'>Oracle table</div>", unsafe_allow_html=True)
    show = result.oracle_df[[
        "code", "item", "engine_group", "rule_family", "reset_mode", "reported_current_hours",
        "engineering_baseline", "oracle_hours", "oracle_low", "oracle_high", "trust_decision",
        "anomaly_score", "oracle_status", "decision_action"
    ]].copy()
    st.dataframe(show, use_container_width=True, height=520, hide_index=True)


def render_component_lab(result: AuditResult):
    if result.oracle_df.empty:
        st.info("No components available.")
        return

    code = st.selectbox("Select component", result.oracle_df["code"].tolist())
    row = result.oracle_df.loc[result.oracle_df["code"] == code].iloc[0]

    left, right = st.columns([0.95, 1.05])
    with left:
        st.markdown("<div class='glass'>", unsafe_allow_html=True)
        st.markdown(f"### {row['code']}")
        st.write(f"**Item:** {row['item']}")
        st.write(f"**Engine group:** {row['engine_group']}")
        st.write(f"**Rule family:** {row['rule_family']}")
        st.write(f"**Reset mode:** {row['reset_mode']}")
        st.write(f"**Status:** {row['oracle_status']}")
        st.write(f"**Action:** {row['decision_action']}")
        st.write(f"**Last event:** {row['last_inspection_date'].date() if pd.notna(row['last_inspection_date']) else '-'}")
        st.write(f"**Due ratio:** {row['due_ratio']:.2f}" if pd.notna(row['due_ratio']) else "**Due ratio:** -")
        st.markdown("</div>", unsafe_allow_html=True)
        st.plotly_chart(chart_component_truth(result.oracle_df, code), use_container_width=True, config={"displayModeBar": False})

    with right:
        st.plotly_chart(chart_waterfall(result.oracle_df, code), use_container_width=True, config={"displayModeBar": False})
        md = pd.DataFrame({
            "month": MONTHS,
            "reported": [row.get(f"reported_{m}") for m in MONTHS],
            "oracle": [row.get(f"oracle_{m}") for m in MONTHS],
        })
        fig = go.Figure()
        fig.add_trace(go.Bar(name="Reported", x=md["month"], y=md["reported"], marker_color="rgba(255,95,134,0.65)"))
        fig.add_trace(go.Bar(name="Oracle", x=md["month"], y=md["oracle"], marker_color="rgba(115,245,208,0.65)"))
        fig.update_layout(
            height=360,
            barmode="group",
            title="Monthly declaration vs oracle",
            margin=dict(l=10, r=10, t=45, b=10),
            paper_bgcolor="rgba(0,0,0,0)",
            plot_bgcolor="rgba(0,0,0,0)",
            font=dict(color="#eef4ff"),
        )
        st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})


def render_issues(result: AuditResult):
    st.markdown("<div class='section-head'>Exception log</div>", unsafe_allow_html=True)
    if result.issues_df.empty:
        st.success("No material exceptions detected.")
    else:
        st.dataframe(result.issues_df, use_container_width=True, hide_index=True, height=420)


def render_mapping(config_df: pd.DataFrame) -> pd.DataFrame:
    st.markdown("<div class='section-head'>Rule control matrix</div>", unsafe_allow_html=True)
    return st.data_editor(
        config_df,
        use_container_width=True,
        num_rows="dynamic",
        hide_index=True,
        column_config={
            "pattern": st.column_config.TextColumn("Pattern"),
            "reset_mode": st.column_config.SelectboxColumn("Reset mode", options=["OH", "RENEWAL", "RENEWAL_OR_OBSERVATION", "CALENDAR"]),
            "family": st.column_config.TextColumn("Family"),
            "criticality": st.column_config.SelectboxColumn("Criticality", options=["High", "Medium", "Low"]),
        },
        key="oracle_rule_editor",
    )


def render_raw(result: AuditResult):
    st.markdown("<div class='section-head'>Raw engineering layers</div>", unsafe_allow_html=True)
    a, b, c = st.tabs(["Daily hours", "Monthly hours", "PMS parse"])
    with a:
        st.dataframe(result.daily_df, use_container_width=True, height=420)
    with b:
        st.dataframe(result.monthly_df, use_container_width=True, height=320)
    with c:
        st.dataframe(result.pms_df, use_container_width=True, height=520)


def main():
    render_hero()

    with st.sidebar:
        st.markdown("### Control")
        uploaded = st.file_uploader("Upload PMS workbook", type=["xlsx", "xlsm"])
        vessel_name = st.text_input("Vessel label", value="MINOAN FALCON")
        st.markdown(
            "<div class='smallnote'>The app is designed for Streamlit/GitHub environments: no browser storage, no custom servers, and fully file-driven execution.</div>",
            unsafe_allow_html=True,
        )

    if "oracle_config_df" not in st.session_state:
        st.session_state.oracle_config_df = DEFAULT_CONFIG.copy()

    tabs = st.tabs(["Overview", "Oracle Lab", "Issues", "Rules", "Raw Data"])

    with tabs[3]:
        st.session_state.oracle_config_df = render_mapping(st.session_state.oracle_config_df)

    if uploaded is None:
        with tabs[0]:
            st.info("Upload the workbook to start the oracle engine.")
            st.markdown(
                """
                - Rebuilds monthly machine hours from the daily sheet.
                - Applies event-reset logic per component family.
                - Calculates engineering baselines, blended oracle estimates, uncertainty bands, and anomaly scores.
                - Produces premium review visuals and an exportable audit workbook.
                """
            )
        return

    result = build_result(uploaded, st.session_state.oracle_config_df)
    export_bytes = export_audit_excel(result)

    with st.sidebar:
        st.download_button(
            "Download oracle workbook",
            data=export_bytes,
            file_name=f"{vessel_name.lower().replace(' ', '_')}_oracle_audit.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        st.download_button(
            "Download rule matrix",
            data=st.session_state.oracle_config_df.to_csv(index=False).encode("utf-8"),
            file_name="oracle_rule_matrix.csv",
            mime="text/csv",
        )

    with tabs[0]:
        render_overview(result)
    with tabs[1]:
        render_component_lab(result)
    with tabs[2]:
        render_issues(result)
    with tabs[4]:
        render_raw(result)


if __name__ == "__main__":
    main()
