import streamlit as st
import pandas as pd
import numpy as np
import re
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta
from io import BytesIO

st.set_page_config(page_title="PMS Running Hours Checker", layout="wide")

# =========================================================
# Helpers
# =========================================================

MONTH_MAP = {
    "jan": 1, "january": 1,
    "feb": 2, "february": 2,
    "mar": 3, "march": 3,
    "apr": 4, "april": 4,
    "may": 5,
    "jun": 6, "june": 6,
    "jul": 7, "july": 7,
    "aug": 8, "august": 8,
    "sep": 9, "sept": 9, "september": 9,
    "oct": 10, "october": 10,
    "nov": 11, "november": 11,
    "dec": 12, "december": 12,
}

MONTH_COLS = ["Jan.", "Feb.", "Mar.", "Apr.", "May", "Jun.", "Jul.", "Aug.", "Sep.", "Oct.", "Nov.", "Dec."]
MONTH_COLS_ALT = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]

ENGINE_DAILY_MAP = {
    "MAIN ENGINE": "me_hrs",
    "DIESEL GENERATOR NO.1": "dg1_hrs",
    "DIESEL GENERATOR NO.2": "dg2_hrs",
    "DIESEL GENERATOR NO.3": "dg3_hrs",
    "DIESEL GENERATOR NO.4": "dg4_hrs",
}

ENTITY_TO_DAILY_COL = {
    "ME": "me_hrs",
    "DG1": "dg1_hrs",
    "DG2": "dg2_hrs",
    "DG3": "dg3_hrs",
    "DG4": "dg4_hrs",
}

def clean_text(x):
    if pd.isna(x):
        return ""
    return str(x).strip()

def safe_num(x):
    if pd.isna(x):
        return np.nan
    s = str(x).strip().replace(",", "")
    if s in ["", "nan", "None", "ERRORREF!", "ERRORVALUE!"]:
        return np.nan
    try:
        return float(s)
    except:
        return np.nan

def parse_date_flexible(x):
    if pd.isna(x):
        return pd.NaT
    if isinstance(x, pd.Timestamp):
        return x.normalize()
    s = str(x).strip()
    if s == "" or s.lower() == "nan":
        return pd.NaT

    try:
        dt = pd.to_datetime(s, dayfirst=True, errors="coerce")
        if pd.notna(dt):
            return dt.normalize()
    except:
        pass

    s2 = s.replace(" ", "").replace(".", "").replace("/", "-")
    s2 = s2.upper()

    m = re.match(r"(\d{1,2})-?([A-Z]{3,9})-?(\d{2,4})", s2)
    if m:
        d = int(m.group(1))
        mon_txt = m.group(2).lower()
        y = int(m.group(3))
        if y < 100:
            y += 2000
        mon = MONTH_MAP.get(mon_txt[:3], MONTH_MAP.get(mon_txt))
        if mon:
            try:
                return pd.Timestamp(year=y, month=mon, day=d)
            except:
                return pd.NaT

    return pd.NaT

def normalize_cols(df):
    cols = []
    for c in df.columns:
        c2 = str(c).replace("\n", " ").strip()
        cols.append(c2)
    df.columns = cols
    return df

def month_name_to_num(name):
    if not isinstance(name, str):
        return None
    s = name.lower().replace(".", "").strip()
    return MONTH_MAP.get(s)

def pretty_month(dt):
    if pd.isna(dt):
        return ""
    return dt.strftime("%b-%y")

def infer_entity_from_code(code):
    code = clean_text(code).upper()
    if code.startswith("ME-"):
        return "ME"
    if code.startswith("DG-01") or code.startswith("DG1") or "DG NO1" in code or "DG NO.1" in code:
        return "DG1"
    if code.startswith("DG-02") or code.startswith("DG2") or "DG NO2" in code or "DG NO.2" in code:
        return "DG2"
    if code.startswith("DG-03") or code.startswith("DG3") or "DG NO3" in code or "DG NO.3" in code:
        return "DG3"
    if code.startswith("DG-04") or code.startswith("DG4") or "DG NO4" in code or "DG NO.4" in code:
        return "DG4"
    return None

def infer_entity_from_item_or_section(code, item, current_section):
    text = f"{clean_text(code)} {clean_text(item)} {clean_text(current_section)}".upper()
    if "MAIN ENGINE" in text or text.startswith("ME-"):
        return "ME"
    if "DIESEL GENERATOR NO 1" in text or "DIESEL GENERATOR NO.1" in text or "DG NO1" in text or "DG NO.1" in text:
        return "DG1"
    if "DIESEL GENERATOR NO 2" in text or "DIESEL GENERATOR NO.2" in text or "DG NO2" in text or "DG NO.2" in text:
        return "DG2"
    if "DIESEL GENERATOR NO 3" in text or "DIESEL GENERATOR NO.3" in text or "DG NO3" in text or "DG NO.3" in text:
        return "DG3"
    if "DIESEL GENERATOR NO 4" in text or "DIESEL GENERATOR NO.4" in text or "DG NO4" in text or "DG NO.4" in text:
        return "DG4"
    return infer_entity_from_code(code)

def parse_interval_to_hours(interval_a, interval_b=None, estimated_per_year=5040):
    candidates = [interval_a, interval_b]
    for val in candidates:
        if pd.isna(val):
            continue
        s = str(val).strip().upper()
        if s == "":
            continue

        n = safe_num(s)
        if pd.notna(n):
            return float(n)

        m = re.match(r"(\d+(?:\.\d+)?)\s*(YEAR|YEARS|MONTH|MONTHS|DAY|DAYS)", s)
        if m:
            qty = float(m.group(1))
            unit = m.group(2)
            if "YEAR" in unit:
                return estimated_per_year * qty
            if "MONTH" in unit:
                return estimated_per_year / 12 * qty
            if "DAY" in unit:
                return estimated_per_year / 365 * qty

        if s == "MONTHLY":
            return estimated_per_year / 12

    return np.nan

def estimate_next_date(last_date, current_hours, interval_hours, estimated_per_year=5040):
    if pd.isna(last_date) or pd.isna(current_hours) or pd.isna(interval_hours) or estimated_per_year <= 0:
        return pd.NaT
    remaining = interval_hours - current_hours
    if remaining <= 0:
        return last_date
    days_needed = remaining / estimated_per_year * 365
    return last_date + pd.to_timedelta(days_needed, unit="D")

# =========================================================
# Daily hours parser
# =========================================================

def parse_daily_operating_hours(xls):
    raw = pd.read_excel(xls, sheet_name="DAILY OPERATING HOURS", header=None)
    raw = raw.copy()

    data_rows = []
    current_date = None

    for i in range(len(raw)):
        row = raw.iloc[i].tolist()
        row_txt = [clean_text(x) for x in row]

        date_val = row[0] if len(row) > 0 else None
        date_parsed = parse_date_flexible(date_val)

        if pd.notna(date_parsed):
            nums = [safe_num(x) for x in row]
            data_rows.append({
                "date": date_parsed,
                "me_hrs": nums[1] if len(nums) > 1 else np.nan,
                "dg1_hrs": nums[4] if len(nums) > 4 else np.nan,
                "dg2_hrs": nums[6] if len(nums) > 6 else np.nan,
                "dg3_hrs": nums[8] if len(nums) > 8 else np.nan,
                "dg4_hrs": nums[10] if len(nums) > 10 else np.nan,
            })

    df = pd.DataFrame(data_rows)
    if df.empty:
        return df, pd.DataFrame()

    for c in ["me_hrs", "dg1_hrs", "dg2_hrs", "dg3_hrs", "dg4_hrs"]:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)

    df["year"] = df["date"].dt.year
    df["month"] = df["date"].dt.month

    monthly = (
        df.groupby(["year", "month"], as_index=False)[["me_hrs", "dg1_hrs", "dg2_hrs", "dg3_hrs", "dg4_hrs"]]
        .sum()
        .sort_values(["year", "month"])
    )

    return df, monthly

# =========================================================
# PMS parser
# =========================================================

def parse_pms_sheet(xls):
    raw = pd.read_excel(xls, sheet_name="PMS", header=None)
    raw = raw.replace({np.nan: ""})

    rows = []
    current_section = ""

    for i in range(len(raw)):
        vals = raw.iloc[i].tolist()
        txt = [clean_text(v) for v in vals]

        if len(txt) < 3:
            continue

        first = txt[0].upper()
        second = txt[1].upper() if len(txt) > 1 else ""

        if "MAIN ENGINE Type" in " ".join(txt):
            current_section = "ME"
        elif "DIESEL GENERATOR No 1" in " ".join(txt) or "DIESEL GENERATOR No.1" in " ".join(txt):
            current_section = "DG1"
        elif "DIESEL GENERATOR No 2" in " ".join(txt) or "DIESEL GENERATOR No.2" in " ".join(txt):
            current_section = "DG2"
        elif "DIESEL GENERATOR No 3" in " ".join(txt) or "DIESEL GENERATOR No.3" in " ".join(txt):
            current_section = "DG3"
        elif "DIESEL GENERATOR No 4" in " ".join(txt) or "DIESEL GENERATOR No.4" in " ".join(txt):
            current_section = "DG4"

        if re.match(r"^[A-Z]{2,3}-\d{2}-\d{2}$", first):
            code = txt[0]
            item = txt[1] if len(txt) > 1 else ""
            job = txt[2] if len(txt) > 2 else ""
            interval1 = txt[3] if len(txt) > 3 else ""
            interval2 = txt[4] if len(txt) > 4 else ""
            last_inspection = parse_date_flexible(vals[5] if len(vals) > 5 else "")
            end_last_year_hours = safe_num(vals[6] if len(vals) > 6 else "")
            current_oper_hours = safe_num(vals[7] if len(vals) > 7 else "")
            est_next = parse_date_flexible(vals[8] if len(vals) > 8 else "")

            month_vals = {}
            for idx, m in enumerate(MONTH_COLS, start=9):
                month_vals[m] = safe_num(vals[idx] if len(vals) > idx else "")

            entity = infer_entity_from_item_or_section(code, item, current_section)

            rows.append({
                "code": code,
                "item": item,
                "job": job,
                "interval_1": interval1,
                "interval_2": interval2,
                "last_inspection_date": last_inspection,
                "oper_hours_end_last_year": end_last_year_hours,
                "current_oper_hours_pms": current_oper_hours,
                "estimated_next_inspection_pms": est_next,
                "entity": entity,
                **month_vals
            })

    return pd.DataFrame(rows)

# =========================================================
# Running logic
# =========================================================

def monthly_hours_since_date(monthly_df, entity_col, start_date, report_date):
    if pd.isna(start_date) or pd.isna(report_date):
        return 0.0, {}

    start_date = pd.Timestamp(start_date).normalize()
    report_date = pd.Timestamp(report_date).normalize()

    md = monthly_df.copy()
    md["month_start"] = pd.to_datetime(dict(year=md["year"], month=md["month"], day=1))
    md["month_end"] = md["month_start"] + pd.offsets.MonthEnd(1)

    total = 0.0
    parts = {}

    for _, r in md.iterrows():
        ms = r["month_start"]
        me = r["month_end"]

        if me < start_date or ms > report_date:
            continue

        month_hours = r[entity_col]

        # same-month start assumed to count from next month only for PMS pattern? no.
        # Based on the workbook, if last OH is on 28-Feb-26, March accumulates 269 and February contributes 0. [file:49]
        # So same-month contributions are ignored and accumulation starts from next month close.
        if ms.year == start_date.year and ms.month == start_date.month:
            contrib = 0.0
        else:
            contrib = month_hours

        parts[f"{ms.year}-{ms.month:02d}"] = contrib
        total += contrib

    return total, parts

def calculate_expected_hours(pms_df, monthly_df, report_date, estimated_per_year=5040):
    out = pms_df.copy()
    expected = []
    mismatches = []
    details = []
    next_dates = []
    intervals = []

    for _, r in out.iterrows():
        entity = r["entity"]
        daily_col = ENTITY_TO_DAILY_COL.get(entity)
        if not daily_col:
            expected.append(np.nan)
            mismatches.append(False)
            details.append("")
            next_dates.append(pd.NaT)
            intervals.append(np.nan)
            continue

        last_date = r["last_inspection_date"]
        exp_hours, parts = monthly_hours_since_date(monthly_df, daily_col, last_date, report_date)

        pms_hours = r["current_oper_hours_pms"]
        mismatch = False
        if pd.notna(pms_hours):
            mismatch = abs(exp_hours - pms_hours) > 0.5

        int_hours = parse_interval_to_hours(r["interval_1"], r["interval_2"], estimated_per_year)
        next_dt = estimate_next_date(last_date, exp_hours, int_hours, estimated_per_year)

        expected.append(exp_hours)
        mismatches.append(mismatch)
        details.append(", ".join([f"{k}:{v:.0f}" for k, v in parts.items() if abs(v) > 0]))
        next_dates.append(next_dt)
        intervals.append(int_hours)

    out["expected_hours"] = expected
    out["mismatch"] = mismatches
    out["calc_breakdown"] = details
    out["interval_hours"] = intervals
    out["estimated_next_by_app"] = next_dates
    out["delta_hours"] = out["current_oper_hours_pms"] - out["expected_hours"]

    return out

# =========================================================
# UI
# =========================================================

st.title("PMS Running Hours Checker")
st.caption("Checks PMS current running hours against monthly operating hours and reset dates")

uploaded_file = st.file_uploader("Upload PMS Excel file", type=["xlsx"])

default_report_date = pd.Timestamp("2026-03-31")
report_date = st.date_input("Report date", value=default_report_date)

estimated_per_year = st.number_input("Estimated steaming hours per year", min_value=1000, max_value=10000, value=5040, step=10)

if uploaded_file is not None:
    try:
        xls = pd.ExcelFile(uploaded_file)

        daily_df, monthly_df = parse_daily_operating_hours(xls)
        pms_df = parse_pms_sheet(xls)

        if daily_df.empty or pms_df.empty:
            st.error("Could not parse required sheets or data structure.")
            st.stop()

        result_df = calculate_expected_hours(
            pms_df=pms_df,
            monthly_df=monthly_df,
            report_date=pd.Timestamp(report_date),
            estimated_per_year=estimated_per_year
        )

        st.success("Workbook parsed successfully")

        tab1, tab2, tab3, tab4 = st.tabs(["Summary", "Mismatches", "All Items", "Monthly Hours"])

        with tab1:
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Parsed PMS Items", len(result_df))
            c2.metric("Mismatches", int(result_df["mismatch"].sum()))
            c3.metric("ME Mar-26 Hours", int(monthly_df.query("year==2026 and month==3")["me_hrs"].sum()) if not monthly_df.query("year==2026 and month==3").empty else 0)
            c4.metric("DG1 Mar-26 Hours", int(monthly_df.query("year==2026 and month==3")["dg1_hrs"].sum()) if not monthly_df.query("year==2026 and month==3").empty else 0)

            st.markdown("### Known examples from your file")
            ex = result_df[result_df["code"].isin(["ME-02-01", "ME-02-04", "ME-02-06", "ME-02-08"])][[
                "code", "item", "last_inspection_date", "current_oper_hours_pms", "expected_hours", "delta_hours", "calc_breakdown"
            ]]
            st.dataframe(ex, use_container_width=True)

        with tab2:
            mismatch_df = result_df[result_df["mismatch"]].copy()
            st.dataframe(
                mismatch_df[[
                    "entity", "code", "item", "last_inspection_date",
                    "current_oper_hours_pms", "expected_hours", "delta_hours",
                    "interval_1", "interval_2", "estimated_next_inspection_pms",
                    "estimated_next_by_app", "calc_breakdown"
                ]].sort_values(["entity", "code"]),
                use_container_width=True
            )

        with tab3:
            filt_entity = st.selectbox("Filter entity", ["ALL", "ME", "DG1", "DG2", "DG3", "DG4"])
            filt_text = st.text_input("Search code/item", "")

            view = result_df.copy()
            if filt_entity != "ALL":
                view = view[view["entity"] == filt_entity]
            if filt_text.strip():
                patt = filt_text.strip().upper()
                view = view[
                    view["code"].str.upper().str.contains(patt, na=False) |
                    view["item"].str.upper().str.contains(patt, na=False)
                ]

            st.dataframe(
                view[[
                    "entity", "code", "item", "job", "interval_1", "interval_2",
                    "last_inspection_date", "current_oper_hours_pms", "expected_hours",
                    "delta_hours", "mismatch", "estimated_next_inspection_pms",
                    "estimated_next_by_app", "calc_breakdown"
                ]].sort_values(["entity", "code"]),
                use_container_width=True
            )

        with tab4:
            monthly_show = monthly_df.copy()
            monthly_show["month_name"] = monthly_show["month"].map({
                1:"Jan", 2:"Feb", 3:"Mar", 4:"Apr", 5:"May", 6:"Jun",
                7:"Jul", 8:"Aug", 9:"Sep", 10:"Oct", 11:"Nov", 12:"Dec"
            })
            st.dataframe(monthly_show[["year", "month", "month_name", "me_hrs", "dg1_hrs", "dg2_hrs", "dg3_hrs", "dg4_hrs"]], use_container_width=True)

        # export
        export_cols = [
            "entity", "code", "item", "job", "interval_1", "interval_2",
            "last_inspection_date", "current_oper_hours_pms", "expected_hours",
            "delta_hours", "mismatch", "estimated_next_inspection_pms",
            "estimated_next_by_app", "calc_breakdown"
        ]
        export_df = result_df[export_cols].copy()

        output = BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter", datetime_format="dd-mm-yyyy") as writer:
            export_df.to_excel(writer, index=False, sheet_name="check_results")
            monthly_df.to_excel(writer, index=False, sheet_name="monthly_hours")
            pms_df.to_excel(writer, index=False, sheet_name="parsed_pms")

            wb = writer.book
            ws = writer.sheets["check_results"]

            red_fmt = wb.add_format({"bg_color": "#FFC7CE", "font_color": "#9C0006"})
            green_fmt = wb.add_format({"bg_color": "#C6EFCE", "font_color": "#006100"})
            num_fmt = wb.add_format({"num_format": "0"})
            date_fmt = wb.add_format({"num_format": "dd-mm-yyyy"})

            ws.set_column("A:A", 10)
            ws.set_column("B:B", 14)
            ws.set_column("C:C", 38)
            ws.set_column("D:D", 12)
            ws.set_column("E:F", 14)
            ws.set_column("G:G", 14, date_fmt)
            ws.set_column("H:J", 14, num_fmt)
            ws.set_column("K:K", 10)
            ws.set_column("L:M", 16, date_fmt)
            ws.set_column("N:N", 30)

            mismatch_col = export_df.columns.get_loc("mismatch")
            ws.conditional_format(1, mismatch_col, len(export_df), mismatch_col, {
                "type": "cell",
                "criteria": "==",
                "value": True,
                "format": red_fmt
            })
            ws.conditional_format(1, mismatch_col, len(export_df), mismatch_col, {
                "type": "cell",
                "criteria": "==",
                "value": False,
                "format": green_fmt
            })

        st.download_button(
            label="Download check results",
            data=output.getvalue(),
            file_name="pms_running_hours_check.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.exception(e)

else:
    st.info("Upload the PMS Excel workbook to start checking results.")

    st.markdown("""
### What this app checks

- Reads **DAILY OPERATING HOURS** monthly totals like ME March = 269, DG1 March = 554, DG2 March = 241, DG3 March = 251
- Reads **PMS** rows such as `ME-02-01`, `ME-02-04`, etc.
- Recomputes running hours from the **last inspection / OH date**
- Flags mismatches between PMS value and recalculated value
- Exports a clean Excel mismatch report

### Expected examples for your March 2026 file

- `ME-02-01` should be **269** because last OH is **28-Feb-26**
- `ME-02-04` should be **846** because last inspection is **18-Jan-26**, so Feb 577 + Mar 269
- Same logic applies to `ME-02-06` and `ME-02-08`
""")
