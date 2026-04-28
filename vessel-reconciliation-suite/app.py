import streamlit as st
import pandas as pd
from datetime import datetime
import hashlib

# --- 1. ENTERPRISE PAGE CONFIGURATION ---
st.set_page_config(
    page_title="Minoan Falcon | Bi-Directional Auditor",
    page_icon="🚢",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# --- 2. CORPORATE UI STYLING ---
st.markdown("""
    <style>
    .stApp { background-color: #f8fafc; font-family: 'Inter', sans-serif; }
    div[data-testid="metric-container"] {
        background-color: #ffffff;
        border: 1px solid #e2e8f0;
        padding: 20px;
        border-radius: 12px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.05);
    }
    h1, h2, h3 { color: #0f172a; }
    </style>
""", unsafe_allow_html=True)

# --- 3. HEADER & INGESTION PIPELINE ---
st.title("🚢 Vessel Reconciliation Suite")
st.markdown("**M/V Minoan Falcon | Single-File Bi-Directional Auditor**")
st.markdown("<span style='color: #64748b; font-size: 0.9rem;'>Zero-Trust processing: Cryptographic seal applied upon mathematical verification.</span>", unsafe_allow_html=True)
st.divider()

uploaded_file = st.file_uploader("Drop Master Workbook (TEC-001.xlsx)", type=["xlsx"])

if uploaded_file:
    try:
        with st.status("Executing Bi-Directional Audit...", expanded=True) as status:
            st.write("📥 Reading Master Workbook...")
            xls = pd.ExcelFile(uploaded_file)
            sheet_names = xls.sheet_names

            # --- DEFENSIVE CHECK: Validate Sheets ---
            required_pms = [s for s in sheet_names if 'PMS' in s.upper()]
            required_logs = [s for s in sheet_names if 'DAILY' in s.upper() or 'OPERATING' in s.upper()]

            if not required_pms or not required_logs:
                status.update(label="Audit Failed: Invalid File Schema", state="error")
                st.error("❌ Critical: Workbook must contain both a 'PMS' tab and a 'DAILY OPERATING HOURS' tab.")
                st.stop()

            st.write("⚙️ Extracting data arrays...")
            pms_df = pd.read_excel(xls, sheet_name=required_pms[0], skiprows=7)
            daily_df = pd.read_excel(xls, sheet_name=required_logs[0], skiprows=10)

            # Clean Daily Logs
            daily_df = daily_df.dropna(how='all') 
            daily_df.iloc[:, 0] = pd.to_datetime(daily_df.iloc[:, 0], errors='coerce')
            daily_df = daily_df.dropna(subset=[daily_df.columns[0]])
            
            for i in range(1, len(daily_df.columns)):
                daily_df.iloc[:, i] = pd.to_numeric(daily_df.iloc[:, i], errors='coerce').fillna(0)

            st.write("🔎 Phase 1: Running Physics Constraints Audit on Daily Logs...")
            physics_violations = []
            date_col = daily_df.columns[0]
            me_col = daily_df.columns[1]
            
            for index, row in daily_df.iterrows():
                logged_hrs = row[me_col]
                if logged_hrs > 24 or logged_hrs < 0:
                    physics_violations.append({
                        "Date": row[date_col].strftime('%d-%b-%Y'),
                        "System": "MAIN ENGINE",
                        "Logged Hours": logged_hrs,
                        "Reason": "Violation: Hours exceed 24h limit or are negative."
                    })

            st.write("🧮 Phase 2: Calculating PMS Baseline Drift...")
            audit_results = []
            
            for index, row in pms_df.iterrows():
                if len(row) < 8: continue
                
                component = str(row.iloc[1]).strip()
                last_oh_date = row.iloc[5]
                legacy_hrs = pd.to_numeric(row.iloc[7], errors='coerce')
                
                if component and component != 'nan' and isinstance(last_oh_date, datetime) and pd.notnull(last_oh_date):
                    legacy_hrs_val = float(legacy_hrs) if pd.notnull(legacy_hrs) else 0.0
                    
                    mask = daily_df[date_col] >= last_oh_date
                    verified_hrs = daily_df.loc[mask, me_col].sum()
                    delta = verified_hrs - legacy_hrs_val
                    
                    audit_results.append({
                        "Component": component,
                        "Overhaul Date": last_oh_date.strftime('%d-%b-%Y'),
                        "Legacy Claim": int(legacy_hrs_val),
                        "Verified Math": int(verified_hrs),
                        "Delta": int(delta),
                        "Status": "VERIFIED" if int(delta) == 0 else "OVERWRITTEN"
                    })

            status.update(label="Audit Complete. Generating Cryptographic Seal.", state="complete", expanded=False)

        # --- 4. DASHBOARD RENDERING ---
        if audit_results:
            res_df = pd.DataFrame(audit_results)
            errors_corrected = len(res_df[res_df['Delta'] != 0])
            
            seal_source = res_df.to_json(orient='records')
            digital_seal = hashlib.sha256(seal_source.encode()).hexdigest()

            c1, c2, c3 = st.columns(3)
            c1.metric("Components Audited", len(res_df))
            c2.metric("Discrepancies Corrected", errors_corrected, delta=f"{errors_corrected} Found", delta_color="inverse")
            c3.metric("SHA-256 Digital Seal", f"{digital_seal[:10]}...{digital_seal[-4:]}")

            st.divider()

            if physics_violations:
                st.subheader("⚠️ Phase 1: Daily Logs Physics Violations")
                st.error("The following entries in the Daily Logs violate the 24-hour limit and must be investigated.")
                st.dataframe(pd.DataFrame(physics_violations), use_container_width=True, hide_index=True)
                st.divider()

            st.subheader("📑 Phase 2: PMS Baseline Drift Reconciliation")
            
            def style_dataframe(row):
                if row['Delta'] != 0:
                    return ['background-color: #fff1f2; color: #9f1239'] * len(row) 
                return [''] * len(row)

            styled_df = res_df.style.apply(style_dataframe, axis=1)
            st.dataframe(styled_df, use_container_width=True, hide_index=True)

            st.markdown("### 💾 Export Immutable Baseline")
            st.download_button(
                label="⬇️ Download Verified Baseline (CSV)",
                data=res_df.to_csv(index=False).encode('utf-8'),
                file_name=f"Verified_Baseline_MinoanFalcon_{datetime.now().strftime('%Y%m%d')}.csv",
                mime='text/csv',
                type="primary"
            )

    except Exception as e:
        st.error(f"🚨 Pipeline Crash Prevented: {str(e)}")
        st.info("The software aborted the calculation due to an unreadable format in the Excel file.")

else:
    st.info("Engine Ready. Awaiting Master Workbook upload to commence Bi-Directional Audit.")
