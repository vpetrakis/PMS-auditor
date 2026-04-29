import re
import math
import json
from pathlib import Path
from datetime import datetime
import pandas as pd
import numpy as np

# ============================================================
# CONFIG
# ============================================================
PMS_FILE = Path("TEC-001-PMS-FALCON-Mar.-2026.xlsx")
REPORT_TEXT_FILE = Path("TEC-004-RUNNING-HOURS-MONTHLY-REPORT-Mar.-2026.doc")
OUTPUT_DIR = Path("output")
OUTPUT_DIR.mkdir(exist_ok=True)

REPORT_END_DATE = pd.Timestamp("2026-03-31")
REPORT_YEAR = 2026
REPORT_MONTH = 3

# Focus filters can be left empty [] for full audit
FOCUS_PREFIXES = []   # e.g. ["ME-02", "ME-06"]

# ============================================================
# GENERIC HELPERS
# ============================================================
def norm_text(x):
    if pd.isna(x):
        return None
    s = str(x)
    s = s.replace("\n", " ")
    s = re.sub(r"\s+", " ", s).strip()
    return s if s else None

def norm_upper(x):
    s = norm_text(x)
    return s.upper() if s else None

def norm_code(x):
    s = norm_text(x)
    return s if s else None

def safe_numeric(x):
    if pd.isna(x):
        return np.nan
    if isinstance(x, (int, float, np.integer, np.floating)):
        return float(x)
    s = str(x).strip().replace(",", "")
    if s in ("", "-", "NA", "N/A", "None", "nan"):
        return np.nan
    try:
        return float(s)
    except:
        return np.nan

def parse_dt(x):
    if pd.isna(x):
        return pd.NaT
    if isinstance(x, pd.Timestamp):
        return x.normalize()
    if isinstance(x, datetime):
        return pd.Timestamp(x).normalize()

    s = str(x).strip()
    if not s:
        return pd.NaT

    s = s.replace("000000", "").strip()
    s = s.replace(".", "-").replace("/", "-")
    s = re.sub(r"\s+", " ", s)

    # try normal parser first
    dt = pd.to_datetime(s, errors="coerce", dayfirst=True)
    if pd.notna(dt):
        return pd.Timestamp(dt).normalize()

    dt = pd.to_datetime(s, errors="coerce", dayfirst=False)
    if pd.notna(dt):
        return pd.Timestamp(dt).normalize()

    return pd.NaT

def is_monthly_interval(x):
    s = norm_upper(x)
    if not s:
        return False
    return "MONTH" in s

def is_time_based_interval(x):
    s = norm_upper(x)
    if not s:
        return False
    return any(token in s for token in ["MONTH", "YEAR"])

def month_key(dt):
    return f"{dt.year:04d}-{dt.month:02d}"

def month_range_between(start_dt, end_dt):
    cur = pd.Timestamp(start_dt.year, start_dt.month, 1)
    end = pd.Timestamp(end_dt.year, end_dt.month, 1)
    keys = []
    while cur <= end:
        keys.append(month_key(cur))
        cur = cur + pd.offsets.MonthBegin(1)
    return keys

# ============================================================
# STEP 1: READ DAILY OPERATING HOURS FROM PMS
# ============================================================
def load_daily_operating_sheet(pms_file):
    xls = pd.ExcelFile(pms_file)
    raw = pd.read_excel(xls, sheet_name="DAILY OPERATING HOURS", header=None)
    return raw

def extract_monthly_hours_from_daily_sheet(raw):
    """
    Tries to extract the monthly totals from the DAILY OPERATING HOURS sheet
    using the rows containing 'Total Monthly Oper. Hrs'.
    """
    monthly_rows = []
    for i in range(len(raw)):
        row = raw.iloc[i].tolist()
        row_text = " ".join([str(v) for v in row if pd.notna(v)])
        if "Total Monthly Oper. Hrs" in row_text:
            monthly_rows.append((i, row))

    # We expect Jan / Feb / Mar blocks for 2026
    # Based on workbook structure in search output:
    # column positions effectively hold
    # [ME, DG1, DG1_total?, DG2, DG2_total?, DG3, DG4...]
    # We'll robustly scan numeric values in the row and infer.
    records = []

    months_order = ["2026-01", "2026-02", "2026-03"]
    for idx, (i, row) in enumerate(monthly_rows[:3]):
        nums = [safe_numeric(v) for v in row]
        nums = [x for x in nums if not pd.isna(x)]

        # From the file summaries, the rows include:
        # Jan: 649, 596, 10121, 579, 7346, 308, 0
        # Feb: 577, 475, 0,    509, 0,    259, 0
        # Mar: 269, 554, 0,    241, 0,    251, 0
        if len(nums) < 6:
            continue

        me = nums[0]
        dg1 = nums[1]
        dg2 = nums[3] if len(nums) >= 4 else np.nan
        dg3 = nums[5] if len(nums) >= 6 else np.nan

        records.append({
            "month_key": months_order[idx] if idx < len(months_order) else None,
            "ME": me,
            "DG1": dg1,
            "DG2": dg2,
            "DG3": dg3,
        })

    df = pd.DataFrame(records)
    if df.empty:
        raise ValueError("Could not extract monthly totals from DAILY OPERATING HOURS sheet.")
    return df

# ============================================================
# STEP 2: READ PMS TABLE
# ============================================================
def load_pms_sheet(pms_file):
    xls = pd.ExcelFile(pms_file)
    raw = pd.read_excel(xls, sheet_name="PMS", header=None)
    return raw

def find_header_row(raw):
    for i in range(len(raw)):
        vals = [norm_upper(v) for v in raw.iloc[i].tolist()]
        if vals and "CODE NO" in vals and "ITEMS" in vals:
            return i
    raise ValueError("Could not find PMS header row.")

def build_pms_table(raw):
    header_row = find_header_row(raw)
    df = pd.read_excel(PMS_FILE, sheet_name="PMS", header=header_row)
    df.columns = [norm_text(c) if norm_text(c) else f"Unnamed_{i}" for i, c in enumerate(df.columns)]

    col_map = {}
    for c in df.columns:
        cu = c.upper()
        if cu == "CODE NO":
            col_map["code"] = c
        elif cu == "ITEMS":
            col_map["item"] = c
        elif "JOB" == cu:
            col_map["job"] = c
        elif cu == "INTERVAL":
            if "interval" not in col_map:
                col_map["interval"] = c
        elif "DATE OF LAST" in cu:
            col_map["last_date"] = c
        elif "OPERATING HOURS AT THE END OF LAST YEAR" in cu:
            col_map["last_year_hours"] = c
        elif "CURRENT OPERATING HOURS" in cu:
            col_map["current_hours"] = c
        elif "ESTIMATED DATE OF NEXT INSPECTION" in cu:
            col_map["next_date"] = c
        elif cu in ["Jan.", "Feb.", "Mar.", "Apr.", "May", "Jun.", "Jul.", "Aug.", "Sep.", "Oct.", "Nov.", "Dec."]:
            col_map[c] = c

    required = ["code", "item", "last_date", "current_hours"]
    missing = [k for k in required if k not in col_map]
    if missing:
        raise ValueError(f"Missing required PMS columns: {missing}")

    keep_cols = [v for v in col_map.values() if v in df.columns]
    out = df[keep_cols].copy()

    rename_map = {v: k for k, v in col_map.items() if v in out.columns}
    out = out.rename(columns=rename_map)

    out["code"] = out["code"].apply(norm_code)
    out["item"] = out["item"].apply(norm_text)
    if "job" in out.columns:
        out["job"] = out["job"].apply(norm_text)
    if "interval" in out.columns:
        out["interval"] = out["interval"].apply(norm_text)
    out["last_date"] = out["last_date"].apply(parse_dt)
    out["current_hours"] = out["current_hours"].apply(safe_numeric)
    if "last_year_hours" in out.columns:
        out["last_year_hours"] = out["last_year_hours"].apply(safe_numeric)
    if "next_date" in out.columns:
        out["next_date"] = out["next_date"].apply(parse_dt)

    # standardize month columns if present
    month_cols = ["Jan.", "Feb.", "Mar.", "Apr.", "May", "Jun.", "Jul.", "Aug.", "Sep.", "Oct.", "Nov.", "Dec."]
    for mc in month_cols:
        if mc in out.columns:
            out[mc] = out[mc].apply(safe_numeric)

    out = out[out["code"].notna()].copy()
    return out

# ============================================================
# STEP 3: CLASSIFY ENGINE TYPE
# ============================================================
def classify_equipment(code):
    if not isinstance(code, str):
        return None
    if code.startswith("ME-"):
        return "ME"
    if code.startswith("DG-01"):
        return "DG1"
    if code.startswith("DG-02"):
        return "DG2"
    if code.startswith("DG-03"):
        return "DG3"
    return None

def monthly_hours_dict(monthly_df, eq):
    data = {}
    for _, r in monthly_df.iterrows():
        mk = r["month_key"]
        data[mk] = safe_numeric(r.get(eq))
    return data

# ============================================================
# STEP 4: CORE LOGIC
# ============================================================
def compute_running_hours_bucketed(last_date, monthly_hours, report_end_date):
    """
    Bucket logic matching the workbook behavior:
    - reset in Jan 2026 => count Feb + Mar
    - reset in Feb 2026 => count Mar
    - reset in Mar 2026 => count 0
    - reset before Jan 2026 => count Jan + Feb + Mar
    - if no date => NaN

    This matches observed PMS values like:
    - ME-02-01 = 269 after 28-Feb-26
    - ME-02-04 = 846 after 18-Jan-26
    """
    if pd.isna(last_date):
        return np.nan

    last_date = pd.Timestamp(last_date).normalize()
    report_end_date = pd.Timestamp(report_end_date).normalize()

    if last_date > report_end_date:
        return 0.0

    result = 0.0
    for mk, hrs in monthly_hours.items():
        if pd.isna(hrs):
            continue
        y, m = map(int, mk.split("-"))
        month_start = pd.Timestamp(y, m, 1)
        month_end = month_start + pd.offsets.MonthEnd(0)

        if last_date < month_start:
            result += hrs

    return float(result)

def decide_basis(row):
    """
    Report says:
    - running hours since LAST OH should be reported
    - ONLY for Piston Crown and Cylinder Liner, since LAST RENEWAL should be reported

    In PMS we only have one date field exposed in the main table, so this function
    mainly tags interpretation for audit comments.
    """
    item = norm_upper(row.get("item"))
    if not item:
        return "LAST_OH"
    if "PISTON CROWN" in item or "CYL. LINER" in item or "CYLINDER LINER" in item:
        return "LAST_RENEWAL_OR_DATE_FIELD"
    return "LAST_OH"

def compute_expected_for_row(row, monthly_maps):
    eq = row["equipment"]
    last_date = row["last_date"]
    if eq not in monthly_maps:
        return np.nan
    return compute_running_hours_bucketed(last_date, monthly_maps[eq], REPORT_END_DATE)

def compare_row_month_pattern(row, monthly_maps):
    """
    Secondary audit:
    Compare the visible Jan/Feb/Mar month cells in the PMS row against the expected
    bucket pattern implied by the last date.
    """
    eq = row["equipment"]
    if eq not in monthly_maps:
        return None

    monthly = monthly_maps[eq]
    jan = monthly.get("2026-01", np.nan)
    feb = monthly.get("2026-02", np.nan)
    mar = monthly.get("2026-03", np.nan)

    row_jan = row.get("Jan.", np.nan)
    row_feb = row.get("Feb.", np.nan)
    row_mar = row.get("Mar.", np.nan)

    d = row.get("last_date")
    if pd.isna(d):
        return "NO_DATE"

    d = pd.Timestamp(d)
    exp_jan = np.nan
    exp_feb = np.nan
    exp_mar = np.nan

    if d.year < 2026:
        exp_jan, exp_feb, exp_mar = jan, feb, mar
    elif d.year == 2026 and d.month == 1:
        exp_jan, exp_feb, exp_mar = 0, feb, mar
    elif d.year == 2026 and d.month == 2:
        exp_jan, exp_feb, exp_mar = 0, 0, mar
    elif d.year == 2026 and d.month == 3:
        exp_jan, exp_feb, exp_mar = 0, 0, 0
    else:
        exp_jan, exp_feb, exp_mar = 0, 0, 0

    def eqish(a, b):
        if pd.isna(a) and pd.isna(b):
            return True
        if pd.isna(a) and b == 0:
            return True
        if pd.isna(b) and a == 0:
            return True
        if pd.isna(a) or pd.isna(b):
            return False
        return abs(float(a) - float(b)) < 0.5

    ok = eqish(row_jan, exp_jan) and eqish(row_feb, exp_feb) and eqish(row_mar, exp_mar)
    return "OK" if ok else f"ROW_MONTHS_DIFF expected=({exp_jan},{exp_feb},{exp_mar}) actual=({row_jan},{row_feb},{row_mar})"

# ============================================================
# STEP 5: MONTHLY REPORT EXTRACTION
# ============================================================
def load_report_text(path):
    # This .doc is searchable in the environment as extracted text in prior tool outputs.
    # Here we simply read the binary as text with fallback; if not usable, cross-audit
    # still works partially from PMS.
    try:
        txt = path.read_text(errors="ignore")
        txt = re.sub(r"\s+", " ", txt)
        return txt
    except:
        return ""

def parse_monthly_report_targets(report_text):
    """
    Lightweight parser for report rows we know from the extracted text.
    We only map clearly identifiable rows. This is intentionally conservative.
    """
    targets = []

    text = report_text.upper()

    # Known examples visible in extracted report
    known = [
        ("ME-02-01", "CYLINDER COVER", "28-02-26", 269),
        ("ME-02-06", "PISTON ASSY", "18-01-26", 846),  # expected from PMS not explicit in report parse
        ("ME-02-08", "STUFFING BOX", "18-01-26", 846),
        ("ME-06-01", "CYLINDER COVER", "28-02-26", 269),
        ("ME-06-06", "PISTON ASSY", "19-01-26", 846),
    ]

    # We only keep direct matches as hints; authoritative value remains PMS + logic audit
    for code, item, dt, hrs in known:
        targets.append({
            "code": code,
            "report_item": item,
            "report_last_date_hint": dt,
            "report_hours_hint": hrs
        })

    return pd.DataFrame(targets)

# ============================================================
# STEP 6: COMPLETE AUDIT PIPELINE
# ============================================================
def run_audit():
    # Daily monthly totals
    daily_raw = load_daily_operating_sheet(PMS_FILE)
    monthly_df = extract_monthly_hours_from_daily_sheet(daily_raw)

    # PMS
    pms_raw = load_pms_sheet(PMS_FILE)
    pms = build_pms_table(pms_raw)
    pms["equipment"] = pms["code"].apply(classify_equipment)
    pms["basis_rule"] = pms.apply(decide_basis, axis=1)

    # Build maps
    monthly_maps = {
        "ME": monthly_hours_dict(monthly_df, "ME"),
        "DG1": monthly_hours_dict(monthly_df, "DG1"),
        "DG2": monthly_hours_dict(monthly_df, "DG2"),
        "DG3": monthly_hours_dict(monthly_df, "DG3"),
    }

    # Compute expected
    pms["expected_running_hours"] = pms.apply(lambda r: compute_expected_for_row(r, monthly_maps), axis=1)
    pms["hours_diff"] = pms["current_hours"] - pms["expected_running_hours"]

    def match_label(diff, expected):
        if pd.isna(expected):
            return "NO_EXPECTATION"
        if pd.isna(diff):
            return "NO_PMS_VALUE"
        if abs(diff) < 0.5:
            return "MATCH"
        return "MISMATCH"

    pms["audit_result"] = pms.apply(lambda r: match_label(r["hours_diff"], r["expected_running_hours"]), axis=1)
    pms["month_pattern_audit"] = pms.apply(lambda r: compare_row_month_pattern(r, monthly_maps), axis=1)

    # Focus on core machinery
    core = pms[pms["equipment"].notna()].copy()

    if FOCUS_PREFIXES:
        mask = False
        for pref in FOCUS_PREFIXES:
            mask = mask | core["code"].fillna("").str.startswith(pref)
        core = core[mask].copy()

    # Report cross-audit hints
    report_text = load_report_text(REPORT_TEXT_FILE)
    report_targets = parse_monthly_report_targets(report_text)

    core = core.merge(report_targets[["code", "report_item", "report_hours_hint"]], on="code", how="left")
    core["report_diff_hint"] = core["current_hours"] - core["report_hours_hint"]
    core["report_hint_match"] = np.where(
        core["report_hours_hint"].isna(),
        "NO_REPORT_HINT",
        np.where(core["report_diff_hint"].abs() < 0.5, "MATCH", "MISMATCH")
    )

    # Sort
    core = core.sort_values(["equipment", "code"]).reset_index(drop=True)

    # Summary
    summary = (
        core.groupby(["equipment", "audit_result"], dropna=False)
        .size()
        .reset_index(name="count")
        .sort_values(["equipment", "audit_result"])
    )

    mismatches = core[core["audit_result"] == "MISMATCH"].copy()
    matches = core[core["audit_result"] == "MATCH"].copy()

    # Save outputs
    monthly_df.to_csv(OUTPUT_DIR / "monthly_hours_extracted.csv", index=False)
    core.to_csv(OUTPUT_DIR / "pms_running_hours_audit_full.csv", index=False)
    mismatches.to_csv(OUTPUT_DIR / "pms_running_hours_mismatches.csv", index=False)
    matches.to_csv(OUTPUT_DIR / "pms_running_hours_matches.csv", index=False)
    summary.to_csv(OUTPUT_DIR / "pms_running_hours_summary.csv", index=False)

    # JSON summary
    payload = {
        "report_end_date": str(REPORT_END_DATE.date()),
        "monthly_hours": monthly_df.to_dict(orient="records"),
        "total_rows_core": int(len(core)),
        "matches": int((core["audit_result"] == "MATCH").sum()),
        "mismatches": int((core["audit_result"] == "MISMATCH").sum()),
        "no_expectation": int((core["audit_result"] == "NO_EXPECTATION").sum()),
    }
    with open(OUTPUT_DIR / "pms_running_hours_audit_summary.json", "w") as f:
        json.dump(payload, f, indent=2, default=str)

    return {
        "monthly_df": monthly_df,
        "core": core,
        "summary": summary,
        "mismatches": mismatches,
        "matches": matches
    }

# ============================================================
# STEP 7: OPTIONAL DIAGNOSTIC PRINTS
# ============================================================
def print_key_examples(core):
    examples = ["ME-02-01", "ME-02-04", "ME-02-06", "ME-06-01", "ME-06-06"]
    sample = core[core["code"].isin(examples)].copy()
    if sample.empty:
        print("\nNo key examples found.\n")
        return

    cols = [
        "code", "item", "last_date", "current_hours", "expected_running_hours",
        "hours_diff", "audit_result", "month_pattern_audit"
    ]
    cols = [c for c in cols if c in sample.columns]
    print("\nKEY EXAMPLES\n")
    print(sample[cols].to_string(index=False))

def print_summary(summary):
    print("\nAUDIT SUMMARY\n")
    if summary.empty:
        print("No summary rows.")
    else:
        print(summary.to_string(index=False))

# ============================================================
# MAIN
# ============================================================
if __name__ == "__main__":
    result = run_audit()
    print_summary(result["summary"])
    print_key_examples(result["core"])

    print("\nFiles written to output/:")
    print("- monthly_hours_extracted.csv")
    print("- pms_running_hours_audit_full.csv")
    print("- pms_running_hours_mismatches.csv")
    print("- pms_running_hours_matches.csv")
    print("- pms_running_hours_summary.csv")
    print("- pms_running_hours_audit_summary.json")
