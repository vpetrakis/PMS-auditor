import streamlit as st
import pandas as pd
from datetime import datetime
import hashlib

# --- 1. Page Configuration ---
st.set_page_config(
    page_title="Minoan Falcon | PMS Auditor",
    page_icon="🚢",
    layout="wide"
)

# --- 2. Custom CSS for Corporate Aesthetic ---
# FIXED: Changed from unsafe_allow_all_headers to unsafe_allow_html
st.markdown("""
    <style>
    .main { background-color: #f8f9fa; }
    .stMetric { background-color: #ffffff; padding: 15px; border-radius: 10px; border: 1px solid #e9ecef; }
    </style>
    """, unsafe_allow_html=True)

# --- 3. Title & Header ---
st.title("🚢 Vessel Reconciliation Suite")
st.caption("Enterprise Cloud Auditor | Minoan Falcon")
st.divider()

# --- 4. File Ingestion & Defensive Checks ---
uploaded_file = st.file_uploader("Drop TEC-001 PMS Excel File Here", type=["xlsx"])

if uploaded_file:
    try:
        # DEFENSIVE CHECK 1: Read the Excel structure before parsing data
        xls = pd.ExcelFile(uploaded_file)
        sheet_names = xls.sheet_names

        required_sheets = ['PMS', 'DAILY OPERATING HOURS']
        missing_sheets = [sheet for sheet in required_sheets if sheet not in sheet_names]

        if missing_sheets:
            # If sheets are missing, fail gracefully without crashing
            st.error(f"❌ Invalid File Format: Missing required sheets: {', '.join(missing_sheets)}")
            st.info("Please ensure the uploaded file contains the exact tabs from the TEC-001 standard.")
        else:
            # Load Sheets safely
            pms_df = pd.read_excel(xls, sheet_name='PMS', skiprows=7)
            daily_df = pd.read_excel(xls, sheet_name='DAILY OPERATING HOURS', skiprows=10)

            # Cleanup Daily Logs (Force invalid dates to NaT, then drop them)
            daily_df = daily_df.dropna(subset=['DATE'])
            daily_df['DATE'] = pd.to_datetime(daily_df['DATE'], errors='coerce')
            daily_df = daily_df.dropna(subset=['DATE'])

            # --- 5. Forensic Engine Logic ---
            results = []
            
            with st.spinner("Running Forensic Audit..."):
                for index, row in pms_df.iterrows():
                    try:
                        # DEFENSIVE CHECK 2: Ensure the row has enough columns
                        if len(row) < 8:
                            continue
                            
                        component = str(row.iloc[1]) # Column B (Component Name)
                        last_oh_date = row.iloc[5]   # Column F (Last Inspection Date)
                        claimed_hrs = row.iloc[7]    # Column H (Current Operating Hours)
                        
                        # Proceed only if we have a valid Date object in the overhaul column
                        if pd.notnull(last_oh_date) and isinstance(last_oh_date, datetime):
                            
                            # Filter logs to only include days ON or AFTER the overhaul date
                            mask = (daily_df['DATE'] >= last_oh_date)
                            
                            # Safely sum the MAIN ENGINE hours
                            if 'MAIN ENGINE' in daily_df.columns:
                                verified_hrs = daily_df.loc[mask, 'MAIN ENGINE'].fillna(0).sum()
                            else:
                                # Fallback if the column name has spaces/typos: use the 2nd column
                                verified_hrs = daily_df.loc[mask, daily_df.columns[1]].fillna(0).sum()
                            
                            # Handle missing claimed hours safely
                            claimed_hrs_val = float(claimed_hrs) if pd.notnull(claimed_hrs) else 0.0
                            delta = verified_hrs - claimed_hrs_val
                            
                            results.append({
                                "Component": component,
                                "Last Overhaul": last_oh_date.strftime('%d-%b-%Y'),
                                "Legacy Claim": int(claimed_hrs_val),
                                "Verified Math": int(verified_hrs),
                                "Discrepancy": int(delta),
                                "Status": "✅ Verified" if int(delta) == 0 else "❌ Discrepancy"
                            })
                    except Exception as row_err:
                        # If a specific row is totally corrupted, skip it instead of crashing the app
                        pass

            # --- 6. Display Triage Dashboard ---
            if results:
                res_df = pd.DataFrame(results)
                
                # Top Metrics
                total_errors = len(res_df[res_df['Discrepancy'] != 0])
                col1, col2, col3 = st.columns(3)
                col1.metric("Total Components Audited", len(res_df))
                col2.metric("Discrepancies Found", total_errors, delta=total_errors, delta_color="inverse")
                
                # Digital Seal Generation (SHA-256)
                seal_source = res_df.to_json()
                digital_seal = hashlib.sha256(seal_source.encode()).hexdigest()
                col3.write("**Cryptographic Seal**")
                col3.code(f"{digital_seal[:16]}...", language=None)

                st.subheader("Forensic Audit Results")
                
                # Highlight Rows with Errors in Red
                def highlight_errors(row):
                    color = '#ffebee' if row['Discrepancy'] != 0 else ''
                    return [f'background-color: {color}'] * len(row)

                st.dataframe(
                    res_df.style.apply(highlight_errors, axis=1),
                    use_container_width=True
                )

                # --- 7. Secure Export ---
                st.divider()
                st.download_button(
                    label="Download Verified Audit Report (CSV)",
                    data=res_df.to_csv(index=False).encode('utf-8'),
                    file_name=f"Verified_Audit_Report_{datetime.now().strftime('%Y%m%d')}.csv",
                    mime='text/csv',
                )
            else:
                st.warning("No valid component dates found to audit. Please check the 'PMS' sheet formatting.")

    except Exception as e:
        # DEFENSIVE CHECK 3: The ultimate fail-safe. If Pandas completely crashes, show this.
        st.error(f"🚨 Critical Pipeline Error: {str(e)}")
        st.info("The application caught a fatal error while reading the file. Please ensure it is a valid, uncorrupted Excel document.")

else:
    st.info("Awaiting TEC-001 Upload. Engine ready.")
