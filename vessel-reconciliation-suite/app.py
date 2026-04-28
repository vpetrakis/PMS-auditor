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

st.markdown("""
    <style>
    .stApp { background-color: #f8fafc; font-family: 'Inter', sans-serif; }
    div[data-testid="metric-container"] {
        background-color: #ffffff; border: 1px solid #e2e8f0; padding: 20px; border-radius: 12px;
    }
    </style>
""", unsafe_allow_html=True)

# --- 2. HEADER ---
st.title("🚢 Vessel Reconciliation Suite")
st.markdown("**M/V Minoan Falcon | Dual-File Bi-Directional Auditor**")
st.divider()

# --- 3. DUAL FILE INGESTION ---
st.subheader("1. Secure Data Ingestion")
col1, col2 = st.columns(2)

with col1:
    pms_file = st.file_uploader("📥 Step 1: Upload TEC-001 (PMS Tab)", type=["xlsx", "xls"])
with col2:
    logs_file = st.file_uploader("📥 Step 2: Upload Daily Operating Hours", type=["xlsx", "xls"])

if pms_file and logs_file:
    try:
        with st.status("Executing Bi-Directional Audit...", expanded=True) as status:
            st.write("⚙️ Extracting data from both files...")
            
            # Use engine='openpyxl' to ensure modern Excel files read correctly
            pms_df = pd.read_excel(pms_file, skiprows=7, engine='openpyxl')
            daily_df = pd.read_excel(logs_file, skiprows=10, engine='openpyxl')

            st.write("🧹 Cleaning Data formatting...")
            
            # Safely parse the Daily Logs Date (Assuming Column 0 is the Date)
            date_col_name = daily_df.columns[0]
            # Coerce errors turns impossible dates into NaT (Not a Time), then we drop them.
            daily_df[date_col_name] = pd.to_datetime(daily_df[date_col_name], errors='coerce')
            daily_df = daily_df.dropna(subset=[date_col_name])
            
            # Safely parse the Running Hours (Assuming Column 1 is Main Engine)
            me_col_name = daily_df.columns[1]
            daily_df[me_col_name] = pd.to_numeric(daily_df[me_col_name], errors='coerce').fillna(0)

            st.write("🔎 Phase 1: Running Physics Constraints Audit on Daily Logs...")
            physics_violations = []
            
            for index, row in daily_df.iterrows():
                logged_hrs = row[me_col_name]
                if logged_hrs > 24 or logged_hrs < 0:
                    physics_violations.append({
                        "Date": row[date_col_name].strftime('%d-%b-%Y'),
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
                legacy_hrs = row.iloc[7]
                
                # Check if it's a valid row
                if component != 'nan' and pd.notnull(last_oh_date):
                    # Coerce the overhaul date safely
                    safe_oh_date = pd.to_datetime(last_oh_date, errors='coerce')
                    
                    if pd.notnull(safe_oh_date):
                        legacy_hrs_val = float(pd.to_numeric(legacy_hrs, errors='coerce')) if pd.notnull(legacy_hrs) else 0.0
                        
                        # Math: Sum Daily Hours where Date >= Overhaul Date
                        mask = daily_df[date_col_name] >= safe_oh_date
                        verified_hrs = daily_df.loc[mask, me_col_name].sum()
                        
                        delta = verified_hrs - legacy_hrs_val
                        
                        audit_results.append({
                            "Component": component,
                            "Overhaul Date": safe_oh_date.strftime('%d-%b-%Y'),
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

            st.subheader("2. Audit Results")
            c1, c2, c3 = st.columns(3)
            c1.metric("Components Audited", len(res_df))
            c2.metric("Discrepancies Corrected", errors_corrected, delta=f"{errors_corrected} Found", delta_color="inverse")
            c3.metric("SHA-256 Digital Seal", f"{digital_seal[:10]}...{digital_seal[-4:]}")

            st.divider()

            if physics_violations:
                st.subheader("⚠️ Phase 1: Daily Logs Physics Violations")
                st.error("The following entries in the Daily Logs violate the 24-hour limit.")
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
        st.info("Ensure you are dropping the PMS file in Box 1, and the Daily Logs in Box 2. The system could not read the expected columns.")
else:
    st.info("Engine Ready. Upload BOTH files to commence the audit.")
