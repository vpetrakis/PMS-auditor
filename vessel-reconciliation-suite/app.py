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

# Custom CSS for Corporate Aesthetic
st.markdown("""
    <style>
    .main { background-color: #f8f9fa; }
    .stMetric { background-color: #ffffff; padding: 15px; border-radius: 10px; border: 1px solid #e9ecef; }
    </style>
    """, unsafe_allow_all_headers=True)

# --- 2. Title & Header ---
st.title("🚢 Vessel Reconciliation Suite")
st.caption("Enterprise Client-Side Zero-Trust Auditor | Minoan Falcon")
st.divider()

# --- 3. File Ingestion ---
uploaded_file = st.file_uploader("Drop TEC-001 PMS Excel File Here", type=["xlsx"])

if uploaded_file:
    try:
        # Load Sheets
        # Note: Names must match exactly: 'PMS' and 'DAILY OPERATING HOURS'
        pms_df = pd.read_excel(uploaded_file, sheet_name='PMS', skiprows=7)
        daily_df = pd.read_excel(uploaded_file, sheet_name='DAILY OPERATING HOURS', skiprows=10)

        # Cleanup Daily Logs
        daily_df = daily_df.dropna(subset=['DATE'])
        daily_df['DATE'] = pd.to_datetime(daily_df['DATE'])

        # --- 4. Forensic Engine Logic ---
        results = []
        
        # We look for rows that have a component name and an overhaul date
        # Assuming Columns: B (Items), F (Last Insp), H (Current Hours)
        # We filter for rows that likely contain equipment data
        for index, row in pms_df.iterrows():
            component = str(row.iloc[1]) # Column B
            last_oh_date = row.iloc[5]   # Column F
            claimed_hrs = row.iloc[7]    # Column H
            
            if pd.notnull(last_oh_date) and isinstance(last_oh_date, datetime):
                # Calculate True Hours from Daily Logs
                # Filter logs >= Overhaul Date
                mask = (daily_df['DATE'] >= last_oh_date)
                
                # Logic: Check if it's Main Engine or DG based on component name
                # (Simplified: Summing ME column for this example)
                verified_hrs = daily_df.loc[mask, 'MAIN ENGINE'].sum()
                
                delta = verified_hrs - claimed_hrs
                
                results.append({
                    "Component": component,
                    "Last Overhaul": last_oh_date.strftime('%d-%b-%Y'),
                    "Legacy Claim": int(claimed_hrs) if pd.notnull(claimed_hrs) else 0,
                    "Verified Math": int(verified_hrs),
                    "Discrepancy": int(delta),
                    "Status": "✅ Verified" if delta == 0 else "❌ Discrepancy"
                })

        # --- 5. Display Triage Dashboard ---
        if results:
            res_df = pd.DataFrame(results)
            
            # Metrics
            total_errors = len(res_df[res_df['Discrepancy'] != 0])
            col1, col2, col3 = st.columns(3)
            col1.metric("Total Components Audited", len(res_df))
            col2.metric("Discrepancies Found", total_errors, delta=total_errors, delta_color="inverse")
            
            # Digital Seal (SHA-256)
            seal_source = res_df.to_json()
            digital_seal = hashlib.sha256(seal_source.encode()).hexdigest()
            col3.write("**Cryptographic Seal**")
            col3.code(f"{digital_seal[:16]}...", language=None)

            st.subheader("Forensic Audit Results")
            
            # Highlight Rows with Errors
            def highlight_errors(val):
                color = '#ffebee' if val != 0 else ''
                return f'background-color: {color}'

            st.dataframe(
                res_df.style.applymap(highlight_errors, subset=['Discrepancy']),
                use_container_width=True
            )

            # --- 6. Export ---
            st.divider()
            st.download_button(
                label="Download Verified Audit Report",
                data=res_df.to_csv(index=False).encode('utf-8'),
                file_name=f"Audit_Report_{datetime.now().strftime('%Y%m%d')}.csv",
                mime='text/csv',
            )

    except Exception as e:
        st.error(f"Error parsing file: {e}")
        st.info("Ensure the sheet names are 'PMS' and 'DAILY OPERATING HOURS' and the format matches TEC-001.")

else:
    st.info("Please upload the TEC-001 Excel file to begin the forensic audit.")
