import io
import re
from dataclasses import dataclass
from typing import Dict, List, Optional, Tuple

import numpy as np
import pandas as pd
import streamlit as st

st.set_page_config(page_title="PMS Running Hours Reconciliation", layout="wide")

MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
DEFAULT_TOLERANCE = 0.01


@dataclass
class ParseResult:
    detail: pd.DataFrame
    summary: pd.DataFrame
    source_name: str


class ReconError(Exception):
    pass


@st.cache_data(show_spinner=False)
def read_excel_sheets(file_bytes: bytes) -> Dict[str, pd.DataFrame]:
    excel = pd.ExcelFile(io.BytesIO(file_bytes), engine="openpyxl")
    sheets = {}
    for sheet in excel.sheet_names:
        sheets[sheet] = excel.parse(sheet_name=sheet, header=None, dtype=object)
    return sheets


def normalize_text(value: object) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and np.isnan(value):
        return ""
    text = str(value).replace("\n", " ").replace("\r", " ")
    text = re.sub(r"\s+", " ", text).strip()
    return text


def normalize_name(value: object) -> str:
    text = normalize_text(value).upper()
    text = text.replace("CYL.", "CYLINDER")
    text = text.replace("ASSY", "ASSEMBLY")
    text = text.replace("LINNER", "LINER")
    text = re.sub(r"\bNO\.?\s*(\d+)\b", r"NO \1", text)
    text = re.sub(r"[^A-Z0-9 ]+", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def to_number(value: object) -> Optional[float]:
    if value is None:
        return None
    if isinstance(value, (int, float, np.integer, np.floating)):
        if pd.isna(value):
            return None
        return float(value)
    text = normalize_text(value)
    if not text:
        return None
    text = text.replace(",", "")
    try:
        return float(text)
    except ValueError:
        return None


def detect_sheet_name(sheets: Dict[str, pd.DataFrame], keywords: List[str]) -> Optional[str]:
    for name in sheets:
        upper = name.upper()
        if all(k.upper() in upper for k in keywords):
            return name
    return None


def find_month_header_row(df: pd.DataFrame) -> Optional[int]:
    for i in range(min(len(df), 200)):
        row = [normalize_text(x) for x in df.iloc[i].tolist()]
        count = sum(1 for cell in row if cell in MONTHS)
        if count >= 3:
            return i
    return None


def find_pms_header_row(df: pd.DataFrame) -> Optional[int]:
    for i in range(min(len(df), 120)):
        row = [normalize_text(x).upper() for x in df.iloc[i].tolist()]
        if "CODE NO" in row and "ITEMS" in row:
            return i
    return None


def extract_pms_summary(df: pd.DataFrame) -> pd.DataFrame:
    rows = []
    text_df = df.fillna("")
    for i in range(len(text_df)):
        row = [normalize_text(x) for x in text_df.iloc[i].tolist()]
        joined = " | ".join(row).upper()

        if "CURRENT YEAR UP-TO-DATE TOTAL WORKING HOURS" in joined:
            equipment = None
            prefix = joined.split("CURRENT YEAR UP-TO-DATE TOTAL WORKING HOURS")[0].strip(" |")

            for candidate in [
                "MAIN ENGINE",
                "DIESEL GENERATOR NO. 1",
                "DIESEL GENERATOR NO. 2",
                "DIESEL GENERATOR NO. 3",
                "DIESEL GENERATOR NO. 4",
            ]:
                if candidate in joined:
                    equipment = candidate
                    break

            if equipment is None and prefix:
                equipment = prefix

            nums = [to_number(x) for x in text_df.iloc[i].tolist()]
            nums = [x for x in nums if x is not None]
            if not nums:
                continue

            total = nums[0]
            month_values = nums[1:13]

            row_data = {
                "equipment": equipment or "UNKNOWN",
                "reported_total": total,
            }
            for idx, m in enumerate(MONTHS):
                row_data[f"reported_{m.lower()}"] = month_values[idx] if idx < len(month_values) else None

            rows.append(row_data)

    out = pd.DataFrame(rows)
    if out.empty:
        return out

    out = out.drop_duplicates(subset=["equipment"], keep="first").reset_index(drop=True)
    return out


def parse_pms_detail(df: pd.DataFrame) -> pd.DataFrame:
    header_row = find_pms_header_row(df)
    if header_row is None:
        raise ReconError("Could not locate the PMS item header row.")

    columns = [normalize_text(x) for x in df.iloc[header_row].tolist()]
    ncols = len(columns)
    col_idx = {name: idx for idx, name in enumerate(columns) if name}

    code_col = col_idx.get("Code No")
    item_col = col_idx.get("ITEMS")
    cur_col = col_idx.get("CURRENT OPERATING HOURS")

    if code_col is None or item_col is None or cur_col is None:
        raise ReconError("Required PMS columns are missing.")

    month_cols = {}
    for m in MONTHS:
        if m in col_idx:
            month_cols[m] = col_idx[m]

    records = []
    current_group = ""

    for i in range(header_row + 1, len(df)):
        row = df.iloc[i].tolist()

        code = normalize_text(row[code_col]) if code_col < ncols else ""
        item = normalize_text(row[item_col]) if item_col < ncols else ""

        if not code and item and not re.search(r"\d", item):
            current_group = item

        if not code:
            continue

        if not re.match(r"^[A-Z]{2,4}-\d{2}(?:\.\d+|-\d+)?(?:-\d+)?$", code, flags=re.IGNORECASE):
            if not re.match(r"^[A-Z]{2,4}-\d{2}-\d{2}$", code, flags=re.IGNORECASE):
                continue

        current_hours = to_number(row[cur_col]) if cur_col < ncols else None
        month_vals = {m: (to_number(row[idx]) if idx < ncols else None) for m, idx in month_cols.items()}

        rec = {
            "key_type": "detail",
            "code": code.upper(),
            "item": item,
            "item_norm": normalize_name(item),
            "group": current_group,
            "group_norm": normalize_name(current_group),
            "reported_current_hours": current_hours,
        }

        for m in MONTHS:
            rec[f"reported_{m.lower()}"] = month_vals.get(m)

        records.append(rec)

    out = pd.DataFrame(records)
    if out.empty:
        raise ReconError("No PMS detail rows were parsed.")

    out = out.drop_duplicates(subset=["code"], keep="first").reset_index(drop=True)
    return out


def parse_me_dg_main_parts_log(df: pd.DataFrame) -> pd.DataFrame:
    month_header_row = find_month_header_row(df)
    if month_header_row is None:
        raise ReconError("Could not locate the month header row in the running-hours report.")

    records = []
    current_equipment = ""

    for i in range(month_header_row + 1, len(df)):
        row = df.iloc[i].tolist()

        first = normalize_text(row[0]) if len(row) > 0 else ""
        second = normalize_text(row[1]) if len(row) > 1 else ""

        if second and any(
            tag in second.upper()
            for tag in [
                "MAIN ENGINE TYPE",
                "DIESEL GENERATOR NO 1 TYPE",
                "DIESEL GENERATOR NO 2 TYPE",
                "DIESEL GENERATOR NO 3 TYPE",
                "DIESEL GENERATOR NO 4 TYPE",
            ]
        ):
            current_equipment = second.upper().replace(" TYPE", "")
            continue

        if not first and not second:
            continue

        nums = [to_number(x) for x in row]
        nums = [x for x in nums if x is not None]

        if len(nums) < 4:
            continue

        name = first or second
        upper_name = name.upper()

        if upper_name in {"STANDARD", "UNDERSIZE"}:
            continue

        total = nums[0]
        month_vals = nums[1:13]

        rec = {
            "key_type": "detail",
            "code": None,
            "item": name,
            "item_norm": normalize_name(name),
            "group": current_equipment,
            "group_norm": normalize_name(current_equipment),
            "reported_current_hours": total,
        }

        for idx, m in enumerate(MONTHS):
            rec[f"reported_{m.lower()}"] = month_vals[idx] if idx < len(month_values) else None

        records.append(rec)

    out = pd.DataFrame(records)
    if out.empty:
        raise ReconError("No detail lines were parsed from the running-hours report.")

    return out


def build_report_summary(detail_df: pd.DataFrame) -> pd.DataFrame:
    summaries = []

    for group, sub in detail_df.groupby("group_norm", dropna=False):
        gname = sub["group"].dropna().iloc[0] if sub["group"].notna().any() else ""
        unique_vals = sub["reported_current_hours"].dropna().unique().tolist()
        total = unique_vals[0] if len(unique_vals) == 1 else None

        row = {"equipment": gname or group or "UNKNOWN", "reported_total": total}

        for m in MONTHS:
            vals = sub[f"reported_{m.lower()}"]
            uniq = vals.dropna().unique().tolist()
            row[f"reported_{m.lower()}"] = uniq[0] if len(uniq) == 1 else None

        summaries.append(row)

    return pd.DataFrame(summaries)


def parse_report(file_bytes: bytes, source_name: str) -> ParseResult:
    sheets = read_excel_sheets(file_bytes)
    pms_sheet = detect_sheet_name(sheets, ["PMS"])
    log_sheet = detect_sheet_name(sheets, ["ME", "DG", "MAIN", "PARTS", "LOG"])

    if pms_sheet is None:
        raise ReconError("The workbook does not contain a PMS sheet.")

    pms_df = sheets[pms_sheet]
    detail = parse_pms_detail(pms_df)
    summary = extract_pms_summary(pms_df)

    if log_sheet is not None:
        log_detail = parse_me_dg_main_parts_log(sheets[log_sheet])
        if summary.empty:
            summary = build_report_summary(log_detail)

    return ParseResult(detail=detail, summary=summary, source_name=source_name)


def parse_running_report(file_bytes: bytes, source_name: str) -> ParseResult:
    sheets = read_excel_sheets(file_bytes)
    log_sheet = detect_sheet_name(sheets, ["ME", "DG", "MAIN", "PARTS", "LOG"])
    pms_sheet = detect_sheet_name(sheets, ["PMS"])

    detail = pd.DataFrame()
    summary = pd.DataFrame()

    if log_sheet is not None:
        detail = parse_me_dg_main_parts_log(sheets[log_sheet])
        summary = build_report_summary(detail)

    if detail.empty and pms_sheet is not None:
        detail = parse_pms_detail(sheets[pms_sheet])
        summary = extract_pms_summary(sheets[pms_sheet])

    if detail.empty:
        raise ReconError("Could not parse a running-hours report from the uploaded file.")

    return ParseResult(detail=detail, summary=summary, source_name=source_name)


def equipment_alias(text: str) -> str:
    t = normalize_name(text)
    replacements = {
        "DIESEL GENERATOR NO 1": "DG1",
        "DIESEL GENERATOR NO 2": "DG2",
        "DIESEL GENERATOR NO 3": "DG3",
        "DIESEL GENERATOR NO 4": "DG4",
        "MAIN ENGINE": "ME",
        "DIESEL GENERATOR NO1": "DG1",
        "DIESEL GENERATOR NO2": "DG2",
        "DIESEL GENERATOR NO3": "DG3",
        "DIESEL GENERATOR NO4": "DG4",
    }
    for src, tgt in replacements.items():
        t = t.replace(src, tgt)
    return t.strip()


def infer_family_from_code(code: Optional[str]) -> str:
    if not code:
        return ""
    code = code.upper()
    if code.startswith("ME-"):
        return "ME"
    if code.startswith("DG-01"):
        return "DG1"
    if code.startswith("DG-02"):
        return "DG2"
    if code.startswith("DG-03"):
        return "DG3"
    if code.startswith("DG-04"):
        return "DG4"
    return ""


def enrich_pms_keys(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    out["family"] = out["code"].apply(infer_family_from_code)
    out["cyl_no"] = out["item"].astype(str).str.extract(r"(?:NO\.?|NO)\s*(\d+)", expand=False)
    out["match_name"] = out["item_norm"]
    return out


def enrich_report_keys(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    out["family"] = out["group"].apply(equipment_alias)
    out["cyl_no"] = out["item"].astype(str).str.extract(r"(?:NO\.?|NO)\s*(\d+)", expand=False)
    out["match_name"] = out["item_norm"]
    return out


def match_details(report_df: pd.DataFrame, pms_df: pd.DataFrame) -> pd.DataFrame:
    rep = enrich_report_keys(report_df)
    pms = enrich_pms_keys(pms_df)

    if rep.empty or pms.empty:
        return pd.DataFrame()

    pms_code = pms[pms["code"].notna()].copy()

    direct = rep.merge(
        pms_code,
        how="left",
        left_on=["family", "match_name", "cyl_no"],
        right_on=["family", "match_name", "cyl_no"],
        suffixes=("_report", "_pms"),
    )

    unmatched = direct[direct["code"].isna()].copy()
    matched = direct[direct["code"].notna()].copy()

    if not unmatched.empty:
        loose = unmatched[rep.columns].merge(
            pms_code,
            how="left",
            left_on=["family", "match_name"],
            right_on=["family", "match_name"],
            suffixes=("_report", "_pms"),
        )
        matched = pd.concat([matched, loose[loose["code"].notna()]], ignore_index=True)
        still_unmatched = loose[loose["code"].isna()][rep.columns].copy()
    else:
        still_unmatched = pd.DataFrame(columns=rep.columns)

    result = matched.copy()

    if not still_unmatched.empty:
        for col in pms.columns:
            if col not in result.columns:
                result[col] = np.nan
        for col in still_unmatched.columns:
            if col not in result.columns:
                result[col] = np.nan
        result = pd.concat([result, still_unmatched[result.columns]], ignore_index=True)

    return result


def reconcile_summary(report_summary: pd.DataFrame, pms_summary: pd.DataFrame, tolerance: float) -> pd.DataFrame:
    if report_summary.empty or pms_summary.empty:
        return pd.DataFrame()

    left = report_summary.copy()
    right = pms_summary.copy()

    left["equipment_key"] = left["equipment"].apply(equipment_alias)
    right["equipment_key"] = right["equipment"].apply(equipment_alias)

    merged = left.merge(right, on="equipment_key", how="outer", suffixes=("_report", "_pms"))
    merged["equipment"] = merged["equipment_report"].fillna(merged["equipment_pms"])
    merged["delta_total"] = merged["reported_total_report"] - merged["reported_total_pms"]

    for m in MONTHS:
        ml = f"reported_{m.lower()}_report"
        mr = f"reported_{m.lower()}_pms"
        merged[f"delta_{m.lower()}"] = merged[ml] - merged[mr]

    merged["status"] = np.where(
        merged["delta_total"].abs().fillna(np.inf) <= tolerance,
        "match",
        np.where(merged["delta_total"].isna(), "incomplete", "mismatch"),
    )

    order = ["equipment", "reported_total_report", "reported_total_pms", "delta_total"]
    for m in MONTHS:
        order += [f"reported_{m.lower()}_report", f"reported_{m.lower()}_pms", f"delta_{m.lower()}"]
    order += ["status"]

    return merged[order].sort_values(by=["status", "equipment"], na_position="last").reset_index(drop=True)


def reconcile_detail(report_detail: pd.DataFrame, pms_detail: pd.DataFrame, tolerance: float) -> pd.DataFrame:
    matched = match_details(report_detail, pms_detail)
    if matched.empty:
        return pd.DataFrame()

    matched["report_hours"] = (
        matched["reported_current_hours_report"]
        if "reported_current_hours_report" in matched.columns
        else matched.get("reported_current_hours")
    )

    if "reported_current_hours_pms" in matched.columns:
        matched["pms_hours"] = matched["reported_current_hours_pms"]
    else:
        matched["pms_hours"] = matched.get("reported_current_hours")

    matched["delta_hours"] = matched["report_hours"] - matched["pms_hours"]

    matched["status"] = np.where(
        matched["delta_hours"].abs().fillna(np.inf) <= tolerance,
        "match",
        np.where(matched["pms_hours"].isna(), "missing in PMS", "mismatch"),
    )

    report_item_col = "item_report" if "item_report" in matched.columns else "item"
    report_group_col = "group_report" if "group_report" in matched.columns else "group"

    out = pd.DataFrame(
        {
            "equipment": matched[report_group_col],
            "report_item": matched[report_item_col],
            "pms_code": matched.get("code"),
            "pms_item": matched.get("item_pms", matched.get("item")),
            "reported_hours": matched["report_hours"],
            "pms_hours": matched["pms_hours"],
            "delta_hours": matched["delta_hours"],
            "status": matched["status"],
        }
    )

    return out.sort_values(
        by=["delta_hours"],
        key=lambda s: s.abs(),
        ascending=False,
        na_position="last",
    ).reset_index(drop=True)


def dataframe_download(df: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="reconciliation")
    output.seek(0)
    return output.read()


def render_metric_block(summary_df: pd.DataFrame, detail_df: pd.DataFrame):
    total_rows = len(detail_df)
    mismatches = int((detail_df["status"] == "mismatch").sum()) if not detail_df.empty else 0
    matches = int((detail_df["status"] == "match").sum()) if not detail_df.empty else 0
    missing = int((detail_df["status"] == "missing in PMS").sum()) if not detail_df.empty else 0

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Detail rows", f"{total_rows:,}")
    c2.metric("Matches", f"{matches:,}")
    c3.metric("Mismatches", f"{mismatches:,}")
    c4.metric("Missing in PMS", f"{missing:,}")

    if not summary_df.empty:
        st.subheader("Equipment summary")
        st.dataframe(summary_df, use_container_width=True, hide_index=True)


def main():
    st.title("PMS Running Hours Reconciliation")
    st.caption("Compare reported running hours against PMS values and calculate the delta only.")

    with st.sidebar:
        st.header("Inputs")
        tolerance = st.number_input(
            "Tolerance (hrs)",
            min_value=0.0,
            value=DEFAULT_TOLERANCE,
            step=0.01,
            format="%.2f",
        )
        st.markdown(
            "Upload a reported running-hours workbook and a PMS workbook. "
            "If both datasets are in the same workbook, upload the same file twice."
        )
        report_file = st.file_uploader("Reported running-hours workbook", type=["xlsx", "xlsm", "xls"])
        pms_file = st.file_uploader("PMS workbook", type=["xlsx", "xlsm", "xls"])

    try:
        if report_file is None or pms_file is None:
            st.info("Upload both files to start the reconciliation.")
            return

        report_bytes, report_name = report_file.read(), report_file.name
        pms_bytes, pms_name = pms_file.read(), pms_file.name

        report = parse_running_report(report_bytes, report_name)
        pms = parse_report(pms_bytes, pms_name)

        summary_df = reconcile_summary(report.summary, pms.summary, tolerance)
        detail_df = reconcile_detail(report.detail, pms.detail, tolerance)

        render_metric_block(summary_df, detail_df)

        tab1, tab2, tab3 = st.tabs(["Detail reconciliation", "Summary reconciliation", "Downloads"])

        with tab1:
            st.subheader("Detail reconciliation")
            if detail_df.empty:
                st.warning("No detail reconciliation rows could be generated.")
            else:
                status_options = sorted(detail_df["status"].dropna().unique().tolist())
                equipment_options = sorted(detail_df["equipment"].dropna().astype(str).unique().tolist())

                status_filter = st.multiselect(
                    "Filter status",
                    options=status_options,
                    default=status_options,
                )
                equipment_filter = st.multiselect(
                    "Filter equipment",
                    options=equipment_options,
                    default=equipment_options,
                )

                view = detail_df.copy()
                if status_filter:
                    view = view[view["status"].isin(status_filter)]
                if equipment_filter:
                    view = view[view["equipment"].astype(str).isin(equipment_filter)]

                st.dataframe(view, use_container_width=True, hide_index=True)

        with tab2:
            st.subheader("Summary reconciliation")
            if summary_df.empty:
                st.warning("No summary reconciliation rows could be generated.")
            else:
                st.dataframe(summary_df, use_container_width=True, hide_index=True)

        with tab3:
            st.subheader("Downloads")
            if not detail_df.empty:
                st.download_button(
                    "Download detail reconciliation.xlsx",
                    data=dataframe_download(detail_df),
                    file_name="detail_reconciliation.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            if not summary_df.empty:
                st.download_button(
                    "Download summary reconciliation.xlsx",
                    data=dataframe_download(summary_df),
                    file_name="summary_reconciliation.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

    except ReconError as e:
        st.error(str(e))
    except Exception as e:
        st.exception(e)


if __name__ == "__main__":
    main()
