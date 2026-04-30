import io
import re
import zipfile
from dataclasses import dataclass
from typing import Dict, List, Optional, Tuple

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st

try:
    from sklearn.covariance import LedoitWolf
    SKLEARN_OK = True
except Exception:
    SKLEARN_OK = False


st.set_page_config(
    page_title="Oracle PMS vs Running Hours",
    page_icon="◈",
    layout="wide",
    initial_sidebar_state="expanded",
)

MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]

DEFAULT_RULES = pd.DataFrame([
    {"pattern": "CYLINDER COVER", "reset_mode": "OH", "family": "Cylinder Cover"},
    {"pattern": "CYLINDER COVER COOLING JACKET", "reset_mode": "OH", "family": "Cover Jacket"},
    {"pattern": "CYL. LINER", "reset_mode": "RENEWAL_OR_OBSERVATION", "family": "Cylinder Liner"},
    {"pattern": "CYL LINER COOLING JACKET", "reset_mode": "OH", "family": "Liner Jacket"},
    {"pattern": "PISTON ASSY", "reset_mode": "OH", "family": "Piston Assembly"},
    {"pattern": "PISTON CROWN", "reset_mode": "RENEWAL", "family": "Piston Crown"},
    {"pattern": "STUFFING BOX", "reset_mode": "OH", "family": "Stuffing Box"},
    {"pattern": "EXHAUST VALVE", "reset_mode": "OH", "family": "Exhaust Valve"},
    {"pattern": "STARTING VALVE", "reset_mode": "OH", "family": "Starting Valve"},
    {"pattern": "SAFETY VALVE", "reset_mode": "OH", "family": "Safety Valve"},
    {"pattern": "FUEL VALVE", "reset_mode": "OH", "family": "Fuel Valve"},
    {"pattern": "FUEL PUMP", "reset_mode": "OH", "family": "Fuel Pump"},
    {"pattern": "PLUNGER BARREL", "reset_mode": "RENEWAL", "family": "Plunger Barrel"},
    {"pattern": "SUCTION VALVE", "reset_mode": "OH", "family": "Suction Valve"},
    {"pattern": "PUNCTURE VALVE", "reset_mode": "OH", "family": "Puncture Valve"},
    {"pattern": "AIR COOLER", "reset_mode": "OH", "family": "Air Cooler"},
    {"pattern": "L.O. COOLER", "reset_mode": "OH", "family": "LO Cooler"},
])

CSS = """
<style>
:root {
  --bg:#07111a;
  --bg2:#0c1722;
  --card:rgba(14,24,35,0.72);
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
    radial-gradient(circle at 10% 0%, rgba(115,245,208,0.09), transparent 30%),
    radial-gradient(circle at 100% 0%, rgba(214,179,106,0.09), transparent 28%),
    linear-gradient(180deg, #07111a 0%, #09131d 40%, #0a1520 100%);
  color:var(--text);
}
.block-container {padding-top:1.0rem;padding-bottom:1.4rem;max-width:1500px;}
[data-testid="stSidebar"] {background:linear-gradient(180deg,#09131d 0%, #0d1723 100%); border-right:1px solid var(--line);}
.hero {
  position:relative; overflow:hidden; border:1px solid var(--line); border-radius:28px;
  background:linear-gradient(135deg, rgba(13,22,34,0.92), rgba(8,17,26,0.78));
  padding:28px 30px 24px 30px; margin-bottom:18px; box-shadow:0 20px 60px rgba(0,0,0,0.28);
}
.hero:before {
  content:""; position:absolute; inset:-2px;
  background:linear-gradient(120deg, rgba(115,245,208,0.18), transparent 30%, rgba(214,179,106,0.18), transparent 70%, rgba(159,140,255,0.18));
  filter:blur(24px); z-index:0; animation:drift 10s ease-in-out infinite alternate;
}
@keyframes drift { from {transform:translateX(-2%) translateY(-1%);} to {transform:translateX(2%) translateY(1%);} }
.hero-inner {position:relative; z-index:1;}
.kicker {font-size:0.73rem; letter-spacing:0.22em; text-transform:uppercase; color:var(--gold); margin-bottom:10px;}
.title {font-size:2.45rem; font-weight:800; line-height:1.02; color:var(--text);}
.subtitle {font-size:1rem; color:var(--muted); max-width:900px; margin-top:10px;}
.metric-grid {display:grid; grid-template-columns:repeat(4,minmax(0,1fr)); gap:14px; margin:18px 0 22px;}
.metric-card {
  background:linear-gradient(180deg, rgba(18,30,44,0.86), rgba(13,22,32,0.90));
  border:1px solid var(--line); border-radius:22px; padding:18px 18px 14px;
  box-shadow:0 14px 30px rgba(0,0,0,0.18);
}
.metric-label {font-size:0.8rem; letter-spacing:0.08em; text-transform:uppercase; color:var(--muted);}
.metric-value {font-size:2rem; font-weight:800; color:var(--text); margin-top:6px;}
.metric-sub {font-size:0.86rem; color:var(--muted); margin-top:4px;}
.section-head {font-size:1rem; text-transform:uppercase; letter-spacing:0.18em; color:var(--gold); margin:8px 0 12px;}
.status-ok {color:var(--teal);}
.status-bad {color:var(--rose);}
.status-mid {color:var(--vio);}
.status-warn {color:var(--gold);}
.smallnote {font-size:0.9rem;color:var(--muted)}
.stTabs [data-baseweb="tab-list"] {gap:10px;}
.stTabs [data-baseweb="tab"] {background:rgba(255,255,255,0.03); border:1px solid var(--line); border-radius:14px; padding:10px 16px;}
.stTabs [aria-selected="true"] {background:rgba(115,245,208,0.08)!important; border-color:rgba(115,245,208,0.30)!important;}
</style>
"""
st.markdown(CSS, unsafe_allow_html=True)


@dataclass
class ParseBundle:
    daily_df: pd.DataFrame
    monthly_df: pd.DataFrame
    pms_df: pd.DataFrame
    word_df: pd.DataFrame
    compare_df: pd.DataFrame
    issues_df: pd.DataFrame
    meta: Dict[str, str]


def norm(v) -> str:
    if pd.isna(v):
        return ""
    return " ".join(str(v).replace("\n", " ").replace("\r", " ").split()).strip()


def safe_date(v):
    if pd.isna(v) or norm(v) == "":
        return pd.NaT
    try:
        return pd.to_datetime(v, errors="coerce", dayfirst=True)
    except Exception:
        return pd.NaT


def infer_engine_from_code(code: str) -> str:
    c = norm(code).upper()
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
    c = norm(code).upper()
    parts = c.split("-")
    if len(parts) >= 2:
        return f"{parts[0]}-{parts[1]}"
    return c


def resolve_rule(item: str, rules_df: pd.DataFrame):
    item_u = norm(item).upper()
    best = None
    best_len = -1
    for _, row in rules_df.iterrows():
        p = norm(row["pattern"]).upper()
        if p and p in item_u and len(p) > best_len:
            best = row
            best_len = len(p)
    return best


def find_sheet_name(xls: pd.ExcelFile, hints: List[str]) -> str:
    for hint in hints:
        for name in xls.sheet_names:
            if hint.lower() in name.lower():
                return name
    return xls.sheet_names[0]


def parse_daily_sheet(uploaded_excel) -> Tuple[pd.DataFrame, pd.DataFrame, Dict[str, str]]:
    xls = pd.ExcelFile(uploaded_excel)
    sheet = find_sheet_name(xls, ["daily operating hours", "daily"])
    raw = pd.read_excel(xls, sheet_name=sheet, header=None)

    records = []
    for _, row in raw.iterrows():
        c0 = row.iloc[0] if len(row) else None
        dt = safe_date(c0)
        if pd.isna(dt):
            continue

        vals = list(row.values)
        def num(i):
            if i >= len(vals):
                return 0.0
            v = pd.to_numeric(vals[i], errors="coerce")
            return float(v) if pd.notna(v) else 0.0

        records.append({
            "date": dt,
            "month": dt.strftime("%b"),
            "Main Engine": num(1),
            "Diesel Generator No.1": num(2),
            "Diesel Generator No.2": num(4),
            "Diesel Generator No.3": num(6),
            "Diesel Generator No.4": num(8),
        })

    if not records:
        empty_daily = pd.DataFrame(columns=[
            "date", "month", "Main Engine", "Diesel Generator No.1",
            "Diesel Generator No.2", "Diesel Generator No.3", "Diesel Generator No.4"
        ])
        empty_monthly = pd.DataFrame(columns=[
            "month_num", "month", "Main Engine", "Diesel Generator No.1",
            "Diesel Generator No.2", "Diesel Generator No.3", "Diesel Generator No.4"
        ])
        return empty_daily, empty_monthly, {"daily_sheet": sheet, "daily_parse": "No daily rows detected"}

    daily_df = pd.DataFrame(records).sort_values("date").reset_index(drop=True)
    monthly_df = (
        daily_df.assign(month_num=daily_df["date"].dt.month)
        .groupby(["month_num", "month"], as_index=False)[[
            "Main Engine",
            "Diesel Generator No.1",
            "Diesel Generator No.2",
            "Diesel Generator No.3",
            "Diesel Generator No.4"
        ]]
        .sum()
        .sort_values("month_num")
        .reset_index(drop=True)
    )
    return daily_df, monthly_df, {"daily_sheet": sheet, "daily_parse": "OK"}


def parse_pms_sheet(uploaded_excel) -> Tuple[pd.DataFrame, Dict[str, str]]:
    xls = pd.ExcelFile(uploaded_excel)
    sheet = find_sheet_name(xls, ["pms"])
    df = pd.read_excel(xls, sheet_name=sheet)
    cols = {str(c).strip().lower(): c for c in df.columns}

    def pick(*keys):
        for key in keys:
            for low, real in cols.items():
                if key in low:
                    return real
        return None

    code_col = pick("code no", "code")
    item_col = pick("items", "item")
    interval_col = pick("interval")
    current_col = pick("current operating hours")
    date_col = pick("date of last inspection")
    est_col = pick("estimated date of next inspection")

    out = pd.DataFrame({
        "code": df[code_col] if code_col else "",
        "item": df[item_col] if item_col else "",
        "interval": df[interval_col] if interval_col else "",
        "current_operating_hours": df[current_col] if current_col else np.nan,
        "last_inspection_date": df[date_col] if date_col else pd.NaT,
        "estimated_next_date": df[est_col] if est_col else pd.NaT,
    })

    for m in MONTHS:
        matched = None
        for c in df.columns:
            if norm(c).startswith(m):
                matched = c
                break
        out[m] = df[matched] if matched is not None else np.nan

    out["code"] = out["code"].map(norm)
    out["item"] = out["item"].map(norm)
    out = out[(out["code"] != "") | (out["item"] != "")].copy()

    out["current_operating_hours"] = pd.to_numeric(out["current_operating_hours"], errors="coerce")
    out["last_inspection_date"] = out["last_inspection_date"].map(safe_date)
    out["estimated_next_date"] = out["estimated_next_date"].map(safe_date)
    out["engine_group"] = out["code"].map(infer_engine_from_code)
    out["asset_family"] = out["code"].map(family_from_code)

    for m in MONTHS:
        out[m] = pd.to_numeric(out[m], errors="coerce")

    return out.reset_index(drop=True), {"pms_sheet": sheet}


def read_docx_text(file_bytes: bytes) -> str:
    try:
        with zipfile.ZipFile(io.BytesIO(file_bytes)) as z:
            xml = z.read("word/document.xml").decode("utf-8", errors="ignore")
        text = re.sub(r"</w:p>", "\n", xml)
        text = re.sub(r"<[^>]+>", " ", text)
        text = re.sub(r"\s+", " ", text)
        return text
    except Exception:
        return ""


def clean_word_text_to_lines(text: str) -> List[str]:
    text = text.replace("\xa0", " ")
    text = re.sub(r"\s+", " ", text)
    rough = re.split(r"(ME-\d{2}-\d{2}|DG-\d{2}-\d{2}|DG\d-\d{2}-\d{2})", text)
    lines = []
    i = 0
    while i < len(rough):
        chunk = rough[i].strip()
        if re.fullmatch(r"(ME-\d{2}-\d{2}|DG-\d{2}-\d{2}|DG\d-\d{2}-\d{2})", chunk):
            nxt = rough[i + 1].strip() if i + 1 < len(rough) else ""
            lines.append(f"{chunk} {nxt}".strip())
            i += 2
        else:
            i += 1
    return lines


def parse_word_running_hours(uploaded_word) -> Tuple[pd.DataFrame, Dict[str, str]]:
    name = uploaded_word.name.lower()
    file_bytes = uploaded_word.getvalue()

    if name.endswith(".docx"):
        text = read_docx_text(file_bytes)
    elif name.endswith(".doc"):
        try:
            text = file_bytes.decode("latin-1", errors="ignore")
        except Exception:
            text = ""
    else:
        text = ""

    if not text:
        return pd.DataFrame(columns=["code", "word_running_hours", "word_line"]), {"word_parse": "No readable text detected"}

    lines = clean_word_text_to_lines(text)
    rows = []

    for line in lines:
        line_u = line.upper()
        code_match = re.search(r"(ME-\d{2}-\d{2}|DG-\d{2}-\d{2}|DG\d-\d{2}-\d{2})", line_u)
        if not code_match:
            continue
        code = code_match.group(1).replace("DG1", "DG-01").replace("DG2", "DG-02").replace("DG3", "DG-03").replace("DG4", "DG-04")
        nums = re.findall(r"(?<![\d/])\d{1,6}(?:\.\d+)?(?![\d/])", line_u)
        if not nums:
            continue

        numeric_values = [float(x) for x in nums]
        running_hours = numeric_values[-1]

        rows.append({
            "code": code,
            "word_running_hours": running_hours,
            "word_line": line[:240],
        })

    word_df = pd.DataFrame(rows)
    if word_df.empty:
        return word_df, {"word_parse": "No running-hours rows found"}

    word_df = (
        word_df.groupby("code", as_index=False)
        .agg({
            "word_running_hours": "max",
            "word_line": "first"
        })
        .sort_values("code")
        .reset_index(drop=True)
    )
    return word_df, {"word_parse": "OK"}


def compute_pms_oracle_hours(pms_df: pd.DataFrame, daily_df: pd.DataFrame, rules_df: pd.DataFrame) -> pd.DataFrame:
    def monthly_hours_since(engine: str, since_date) -> float:
        if daily_df.empty or engine not in daily_df.columns or pd.isna(since_date):
            return np.nan
        mask = daily_df["date"] > since_date
        return float(daily_df.loc[mask, engine].sum())

    def ytd_hours(engine: str) -> float:
        if daily_df.empty or engine not in daily_df.columns:
            return np.nan
        return float(daily_df[engine].sum())

    rows = []
    for _, r in pms_df.iterrows():
        rule = resolve_rule(r["item"], rules_df)
        mode = rule["reset_mode"] if rule is not None else "OH"
        family = rule["family"] if rule is not None else "Unmapped"
        engine = r["engine_group"]

        if mode in ["OH", "RENEWAL", "RENEWAL_OR_OBSERVATION"] and pd.notna(r["last_inspection_date"]):
            oracle_hours = monthly_hours_since(engine, r["last_inspection_date"])
            basis = "post_event"
        else:
            oracle_hours = ytd_hours(engine)
            basis = "parent_ytd"

        rows.append({
            "code": r["code"],
            "item": r["item"],
            "engine_group": engine,
            "rule_family": family,
            "reset_mode": mode,
            "current_operating_hours": r["current_operating_hours"],
            "oracle_running_hours": oracle_hours,
            "basis": basis,
            "last_inspection_date": r["last_inspection_date"],
            "estimated_next_date": r["estimated_next_date"],
        })

    out = pd.DataFrame(rows)
    if out.empty:
        return out

    out["pms_vs_oracle_gap"] = out["current_operating_hours"] - out["oracle_running_hours"]
    return out


def compare_pms_word(pms_oracle_df: pd.DataFrame, word_df: pd.DataFrame) -> pd.DataFrame:
    if pms_oracle_df.empty:
        return pd.DataFrame()

    out = pms_oracle_df.merge(word_df, on="code", how="left")
    out["pms_vs_word_gap"] = out["current_operating_hours"] - out["word_running_hours"]
    out["oracle_vs_word_gap"] = out["oracle_running_hours"] - out["word_running_hours"]

    feat = out[["current_operating_hours", "oracle_running_hours", "word_running_hours", "pms_vs_oracle_gap", "pms_vs_word_gap", "oracle_vs_word_gap"]].copy()
    feat = feat.replace([np.inf, -np.inf], np.nan).fillna(0)

    if len(feat) >= 4:
        if SKLEARN_OK:
            model = LedoitWolf().fit(feat.values.astype(float))
            score = np.sqrt(np.maximum(model.mahalanobis(feat.values.astype(float)), 0))
        else:
            z = (feat - feat.mean()) / feat.std(ddof=0).replace(0, 1)
            score = np.sqrt((z ** 2).sum(axis=1)).values
        out["anomaly_score"] = score
        threshold = np.percentile(score, 90)
    else:
        out["anomaly_score"] = 0.0
        threshold = 9.99

    out["anomaly_threshold"] = threshold

    trust = (
        100
        - np.clip(out["pms_vs_word_gap"].abs().fillna(30) * 0.8, 0, 45)
        - np.clip(out["pms_vs_oracle_gap"].abs().fillna(30) * 0.5, 0, 30)
        - np.clip(np.maximum(out["anomaly_score"] - threshold, 0) * 12, 0, 25)
    )
    out["trust_score"] = trust.clip(0, 100)

    out["status"] = np.select(
        [
            out["trust_score"] >= 82,
            (out["trust_score"] >= 58) & (out["trust_score"] < 82),
            out["trust_score"] < 58,
        ],
        ["Verified", "Adjusted", "Quarantined"],
        default="Adjusted",
    )

    out["word_match_flag"] = np.where(out["word_running_hours"].notna(), "Matched", "No Word Match")
    return out.sort_values(["status", "trust_score", "anomaly_score"], ascending=[True, False, False]).reset_index(drop=True)


def build_issues(compare_df: pd.DataFrame) -> pd.DataFrame:
    issues = []
    if compare_df.empty:
        return pd.DataFrame(columns=["severity", "code", "issue", "detail"])

    for _, r in compare_df.iterrows():
        if pd.isna(r["word_running_hours"]):
            issues.append({
                "severity": "high",
                "code": r["code"],
                "issue": "Missing Word match",
                "detail": "Component not found in running-hours monthly report",
            })
        elif abs(r["pms_vs_word_gap"]) > 1:
            sev = "high" if abs(r["pms_vs_word_gap"]) >= 24 else "medium"
            issues.append({
                "severity": sev,
                "code": r["code"],
                "issue": "PMS vs Word mismatch",
                "detail": f"PMS={r['current_operating_hours']:.0f}, Word={r['word_running_hours']:.0f}, gap={r['pms_vs_word_gap']:.0f}",
            })

        if r["status"] == "Quarantined":
            issues.append({
                "severity": "high",
                "code": r["code"],
                "issue": "Low confidence row",
                "detail": f"Trust={r['trust_score']:.0f}, anomaly={r['anomaly_score']:.2f}",
            })

    issues_df = pd.DataFrame(issues)
    if issues_df.empty:
        return pd.DataFrame(columns=["severity", "code", "issue", "detail"])

    issues_df["severity"] = pd.Categorical(issues_df["severity"], ["high", "medium", "low"], ordered=True)
    return issues_df.sort_values(["severity", "code"]).reset_index(drop=True)


def build_bundle(pms_file, word_file, rules_df: pd.DataFrame) -> ParseBundle:
    daily_df, monthly_df, daily_meta = parse_daily_sheet(pms_file)
    pms_df, pms_meta = parse_pms_sheet(pms_file)
    word_df, word_meta = parse_word_running_hours(word_file)
    pms_oracle_df = compute_pms_oracle_hours(pms_df, daily_df, rules_df)
    compare_df = compare_pms_word(pms_oracle_df, word_df)
    issues_df = build_issues(compare_df)
    meta = {**daily_meta, **pms_meta, **word_meta}
    return ParseBundle(daily_df, monthly_df, pms_df, word_df, compare_df, issues_df, meta)


def render_hero():
    st.markdown("""
    <div class='hero'>
      <div class='hero-inner'>
        <div class='kicker'>Oracle • PMS vs Running Hours Reconciliation</div>
        <div class='title'>Compare PMS hours directly against Word running-hours report</div>
        <div class='subtitle'>This version restores the original two-file workflow: one drag-and-drop for the PMS workbook and one for the monthly running-hours Word report, then computes row-level reconciliation, confidence, and anomaly scoring.</div>
      </div>
    </div>
    """, unsafe_allow_html=True)


def render_metrics(compare_df: pd.DataFrame):
    if compare_df.empty:
        return
    verified = int((compare_df["status"] == "Verified").sum())
    adjusted = int((compare_df["status"] == "Adjusted").sum())
    quarantined = int((compare_df["status"] == "Quarantined").sum())
    matched = int(compare_df["word_running_hours"].notna().sum())
    st.markdown(f"""
    <div class='metric-grid'>
      <div class='metric-card'><div class='metric-label'>Word matches</div><div class='metric-value'>{matched}</div><div class='metric-sub'>Rows aligned by component code</div></div>
      <div class='metric-card'><div class='metric-label'>Verified</div><div class='metric-value status-ok'>{verified}</div><div class='metric-sub'>High-trust PMS vs Word agreement</div></div>
      <div class='metric-card'><div class='metric-label'>Adjusted</div><div class='metric-value status-mid'>{adjusted}</div><div class='metric-sub'>Needs oracle support</div></div>
      <div class='metric-card'><div class='metric-label'>Quarantined</div><div class='metric-value status-bad'>{quarantined}</div><div class='metric-sub'>Needs manual review</div></div>
    </div>
    """, unsafe_allow_html=True)


def fig_gap_bar(compare_df: pd.DataFrame):
    top = compare_df.copy()
    top["abs_gap"] = top["pms_vs_word_gap"].abs().fillna(0)
    top = top.sort_values("abs_gap", ascending=False).head(20)
    fig = go.Figure()
    fig.add_trace(go.Bar(
        x=top["code"],
        y=top["pms_vs_word_gap"],
        marker_color=np.where(top["pms_vs_word_gap"].fillna(0) >= 0, "rgba(255,95,134,0.72)", "rgba(115,245,208,0.72)"),
        name="PMS - Word gap"
    ))
    fig.update_layout(
        title="Largest PMS vs Word gaps",
        height=430,
        margin=dict(l=10, r=10, t=55, b=20),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(color="#eef4ff"),
        xaxis_title="Code",
        yaxis_title="Hours gap",
    )
    return fig


def fig_truth_strip(compare_df: pd.DataFrame, code: str):
    r = compare_df.loc[compare_df["code"] == code].iloc[0]
    fig = go.Figure()
    y = ["Running hours"]
    if pd.notna(r["word_running_hours"]):
        fig.add_trace(go.Scatter(x=[r["word_running_hours"]], y=y, mode="markers", marker=dict(size=16, color="#73f5d0"), name="Word"))
    if pd.notna(r["current_operating_hours"]):
        fig.add_trace(go.Scatter(x=[r["current_operating_hours"]], y=y, mode="markers", marker=dict(size=16, color="#ff5f86", symbol="x"), name="PMS"))
    if pd.notna(r["oracle_running_hours"]):
        fig.add_trace(go.Scatter(x=[r["oracle_running_hours"]], y=y, mode="markers", marker=dict(size=16, color="#d6b36a", symbol="diamond"), name="Oracle"))
    fig.update_layout(
        title=f"Truth strip · {code}",
        height=250,
        margin=dict(l=10, r=10, t=50, b=20),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(color="#eef4ff"),
        xaxis_title="Running hours",
        yaxis_title="",
    )
    return fig


def fig_monthly_hours(monthly_df: pd.DataFrame):
    if monthly_df.empty:
        return go.Figure()
    fig = go.Figure()
    for col, color in [
        ("Main Engine", "#73f5d0"),
        ("Diesel Generator No.1", "#d6b36a"),
        ("Diesel Generator No.2", "#9f8cff"),
        ("Diesel Generator No.3", "#ff5f86"),
    ]:
        if col in monthly_df.columns:
            fig.add_trace(go.Bar(x=monthly_df["month"], y=monthly_df[col], name=col, marker_color=color))
    fig.update_layout(
        title="Monthly parent-machine hours from PMS daily sheet",
        barmode="group",
        height=430,
        margin=dict(l=10, r=10, t=55, b=20),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(color="#eef4ff"),
        xaxis_title="Month",
        yaxis_title="Hours",
    )
    return fig


def export_excel(bundle: ParseBundle) -> bytes:
    bio = io.BytesIO()
    with pd.ExcelWriter(bio, engine="xlsxwriter") as writer:
        bundle.daily_df.to_excel(writer, sheet_name="daily_hours", index=False)
        bundle.monthly_df.to_excel(writer, sheet_name="monthly_hours", index=False)
        bundle.pms_df.to_excel(writer, sheet_name="pms_clean", index=False)
        bundle.word_df.to_excel(writer, sheet_name="word_clean", index=False)
        bundle.compare_df.to_excel(writer, sheet_name="comparison", index=False)
        bundle.issues_df.to_excel(writer, sheet_name="issues", index=False)
        pd.DataFrame([bundle.meta]).to_excel(writer, sheet_name="meta", index=False)
    bio.seek(0)
    return bio.getvalue()


def main():
    render_hero()

    if "rules_df" not in st.session_state:
        st.session_state.rules_df = DEFAULT_RULES.copy()

    with st.sidebar:
        st.markdown("### Inputs")
        pms_file = st.file_uploader("Drag & drop PMS workbook", type=["xlsx", "xlsm"], key="pms_uploader")
        word_file = st.file_uploader("Drag & drop running-hours Word report", type=["doc", "docx"], key="word_uploader")
        vessel_name = st.text_input("Vessel label", value="MINOAN FALCON")
        st.markdown("<div class='smallnote'>The app now requires both sources again: PMS Excel and Word running-hours report.</div>", unsafe_allow_html=True)

    tabs = st.tabs(["Overview", "Component Lab", "Issues", "Rule Matrix", "Raw Data"])

    with tabs[3]:
        st.markdown("<div class='section-head'>Reset rule matrix</div>", unsafe_allow_html=True)
        st.session_state.rules_df = st.data_editor(
            st.session_state.rules_df,
            use_container_width=True,
            num_rows="dynamic",
            hide_index=True,
            key="rules_editor",
        )

    if pms_file is None or word_file is None:
        with tabs[0]:
            st.info("Upload both files to launch the reconciliation engine.")
            st.markdown("""
            - PMS workbook upload.
            - Word monthly running-hours report upload.
            - Direct PMS vs Word comparison by component code.
            - Oracle support from PMS daily-hours reconstruction.
            """)
        return

    bundle = build_bundle(pms_file, word_file, st.session_state.rules_df)
    export_bytes = export_excel(bundle)

    with st.sidebar:
        st.download_button(
            "Download reconciliation workbook",
            data=export_bytes,
            file_name=f"{vessel_name.lower().replace(' ', '_')}_reconciliation.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    with tabs[0]:
        render_metrics(bundle.compare_df)
        c1, c2 = st.columns([1, 1])
        with c1:
            st.plotly_chart(fig_gap_bar(bundle.compare_df), use_container_width=True, config={"displayModeBar": False})
        with c2:
            st.plotly_chart(fig_monthly_hours(bundle.monthly_df), use_container_width=True, config={"displayModeBar": False})

        st.markdown("<div class='section-head'>Comparison table</div>", unsafe_allow_html=True)
        if bundle.compare_df.empty:
            st.warning("No comparable rows found between PMS and Word report.")
        else:
            show = bundle.compare_df[[
                "code", "item", "engine_group", "rule_family", "reset_mode",
                "current_operating_hours", "word_running_hours", "oracle_running_hours",
                "pms_vs_word_gap", "pms_vs_oracle_gap", "oracle_vs_word_gap",
                "trust_score", "status", "word_match_flag"
            ]].copy()
            st.dataframe(show, use_container_width=True, hide_index=True, height=520)

    with tabs[1]:
        st.markdown("<div class='section-head'>Component oracle lab</div>", unsafe_allow_html=True)
        if bundle.compare_df.empty:
            st.info("No component rows available.")
        else:
            code = st.selectbox("Select component code", bundle.compare_df["code"].tolist())
            row = bundle.compare_df.loc[bundle.compare_df["code"] == code].iloc[0]
            l1, l2 = st.columns([0.95, 1.05])
            with l1:
                st.write(f"**Code:** {row['code']}")
                st.write(f"**Item:** {row['item']}")
                st.write(f"**Engine group:** {row['engine_group']}")
                st.write(f"**Rule family:** {row['rule_family']}")
                st.write(f"**Reset mode:** {row['reset_mode']}")
                st.write(f"**PMS current hours:** {row['current_operating_hours']}")
                st.write(f"**Word running hours:** {row['word_running_hours']}")
                st.write(f"**Oracle running hours:** {row['oracle_running_hours']}")
                st.write(f"**PMS vs Word gap:** {row['pms_vs_word_gap']}")
                st.write(f"**Trust score:** {row['trust_score']:.1f}")
                st.write(f"**Status:** {row['status']}")
                st.plotly_chart(fig_truth_strip(bundle.compare_df, code), use_container_width=True, config={"displayModeBar": False})
            with l2:
                vals = pd.DataFrame({
                    "Source": ["PMS", "Word", "Oracle"],
                    "Hours": [row["current_operating_hours"], row["word_running_hours"], row["oracle_running_hours"]],
                })
                fig = go.Figure()
                fig.add_trace(go.Bar(
                    x=vals["Source"],
                    y=vals["Hours"],
                    marker_color=["#ff5f86", "#73f5d0", "#d6b36a"]
                ))
                fig.update_layout(
                    title="Source comparison",
                    height=350,
                    margin=dict(l=10, r=10, t=55, b=20),
                    paper_bgcolor="rgba(0,0,0,0)",
                    plot_bgcolor="rgba(0,0,0,0)",
                    font=dict(color="#eef4ff"),
                    yaxis_title="Hours",
                )
                st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})

                fig2 = go.Figure()
                fig2.add_trace(go.Indicator(
                    mode="gauge+number",
                    value=float(row["trust_score"]),
                    title={"text": "Trust score"},
                    gauge={
                        "axis": {"range": [0, 100]},
                        "bar": {"color": "#73f5d0" if row["trust_score"] >= 82 else "#d6b36a" if row["trust_score"] >= 58 else "#ff5f86"},
                        "steps": [
                            {"range": [0, 58], "color": "rgba(255,95,134,0.18)"},
                            {"range": [58, 82], "color": "rgba(214,179,106,0.18)"},
                            {"range": [82, 100], "color": "rgba(115,245,208,0.18)"},
                        ],
                    }
                ))
                fig2.update_layout(
                    height=290,
                    margin=dict(l=20, r=20, t=40, b=10),
                    paper_bgcolor="rgba(0,0,0,0)",
                    font=dict(color="#eef4ff"),
                )
                st.plotly_chart(fig2, use_container_width=True, config={"displayModeBar": False})

    with tabs[2]:
        st.markdown("<div class='section-head'>Issues</div>", unsafe_allow_html=True)
        if bundle.issues_df.empty:
            st.success("No major exceptions found.")
        else:
            st.dataframe(bundle.issues_df, use_container_width=True, hide_index=True, height=450)

    with tabs[4]:
        st.markdown("<div class='section-head'>Raw parsed data</div>", unsafe_allow_html=True)
        a, b, c, d = st.tabs(["Daily", "PMS", "Word", "Meta"])
        with a:
            st.dataframe(bundle.daily_df, use_container_width=True, height=420)
        with b:
            st.dataframe(bundle.pms_df, use_container_width=True, height=520)
        with c:
            st.dataframe(bundle.word_df, use_container_width=True, height=420)
        with d:
            st.json(bundle.meta)


if __name__ == "__main__":
    main()
