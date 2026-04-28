import streamlit as st
import pandas as pd
import numpy as np
import io
import re
import hashlib
from datetime import datetime
import warnings

warnings.filterwarnings("ignore")

# ═══════════════════════════════════════════════════════════════════════════════
# 1. ENTERPRISE PAGE CONFIGURATION
# ═══════════════════════════════════════════════════════════════════════════════
st.set_page_config(
    page_title="Vessel Reconciliation Suite",
    page_icon="⚓",
    layout="wide",
    initial_sidebar_state="collapsed"
)

st.markdown("""
    <style>
    .stApp { background-color: #f8fafc; font-family: 'Inter', sans-serif; }
    div[data-testid="metric-container"] {
        background-color: #ffffff; border: 1px solid #e2e8f0; padding: 20px; border-radius: 12px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.05);
    }
    h1, h2, h3 { color: #0f172a; }
    </style>
""", unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════════════════════
# 2. SEMANTIC PARSERS (Inspired by POSEIDON TITAN)
# ═══════════════════════════════════════════════════════════════════════════════
def semantic_logs_parse(file_bytes):
    """Hunts for the Date and Main Engine columns dynamically in the Daily Logs."""
    df_raw = pd.read_excel(io.BytesIO(file_bytes), header=None, engine='openpyxl', dtype=str)
    
    header_idx, date_idx, me_idx = -1, -1, -1
    
    # Scan the first 50 rows to find the headers
    for i in range(min(50, len(df_raw))):
        vals = [str(x).upper() for x in df_raw.iloc[i].values if pd.notna(x)]
        
        # Look for Date indicator AND Engine indicator in the same row
        if any('DATE' in v or 'DAY' in v for v in vals) and any('MAIN' in v or 'ME ' in v for v in vals):
            header_idx = i
            # Map the exact column indices
            for j, val in enumerate(df_raw.iloc[i].values):
                v_str = str(val).upper() if pd.notna(val) else ""
                if 'DATE' in v_str or 'DAY' in v_str: date_idx = j
                elif 'MAIN' in v_str or 'ME ' in v_str: me_idx = j
            break

    if header_idx == -1 or date_idx == -1 or me_idx == -1:
        raise ValueError("Semantic Lock Failed: Could not locate 'DATE' and 'MAIN ENGINE' columns in the Daily Logs.")

    # Slice data from the row below the header
    df = df_raw.iloc[header_idx + 1:].copy().reset_index(drop=True)
    
    # Clean the extracted vectors
    clean_df = pd.DataFrame()
    clean_df['Date'] = pd.to_datetime(df.iloc[:, date_idx], errors='coerce')
    clean_df['ME_Hours'] = pd.to_numeric(df.iloc[:, me_idx], errors='coerce').fillna(0)
    
    # Drop rows that don't have a valid calendar date
    clean_df = clean_df.dropna(subset=['Date']).reset_index(drop=True)
    return clean_df

def semantic_pms_parse(file_bytes):
    """Hunts for the Component, Overhaul Date, and Claimed Hours in the PMS file."""
    df_raw = pd.read_excel(io.BytesIO(file_bytes), header=None, engine='openpyxl', dtype=str)
    
    header_idx = -1
    comp_idx, date_idx, hrs_idx = -1, -1, -1

    for i in range(min(50, len(df_raw))):
        vals = [str(x).upper() for x in df_raw.iloc[i].values if pd.notna(x)]
        # Usually PMS headers contain ITEM, DATE, HOURS
        if any('ITEM' in v or 'COMPONENT' in v for v in vals) and any('DATE' in v or 'OVERHAUL' in v for v in vals):
            header_idx = i
            for j, val in enumerate(df_raw.iloc[i].values):
                v_str = str(val).upper() if pd.notna(val) else ""
                if 'ITEM' in v_str or 'COMPONENT' in v_str or 'NAME' in v_str: comp_idx = j
                elif 'DATE' in v_str or 'OVERHAUL' in v_str or 'INSP' in v_str: date_idx = j
                elif 'HOUR' in v_str or 'RUN' in v_str or 'CURRENT' in v_str: hrs_idx = j
            break

    # Fallback to absolute column indices if semantic mapping misses the exact phrasing
    if header_idx == -1:
        comp_idx, date_idx, hrs_idx = 1, 5, 7  # Columns B, F, H
        header_idx = 7 # skip 7 rows

    df = df_raw.iloc[header_idx + 1:].copy().reset_index(drop=True)
    
    clean_df = pd.DataFrame()
    clean_df['Component'] = df.iloc[:, comp_idx].astype(str).str.strip()
    clean_df['Last_Overhaul'] = pd.to_datetime(df.iloc[:, date_idx], errors='coerce')
    clean_df['Claimed_Hours'] = pd.to_numeric(df.iloc[:, hrs_idx], errors='coerce').fillna(0)
    
    # Filter out empty rows
    clean_df = clean_df[clean_df['Component'] != 'nan']
    return clean_df

# ═══════════════════════════════════════════════════════════════════════════════
# 3. TRIAGE UI & EXECUTION ROUTER
# ═══════════════════════════════════════════════════════════════════════════════
st.title("🚢 Vessel Reconciliation Suite")
st.markdown("**M/V Minoan Falcon | Forensic Baseline Auditor**")
st.divider()

st.subheader("1. Semantic Data Ingestion")
col1, col2 = st.columns(2)
with col1: pms_file = st.file_uploader("📥 Step 1: Upload TEC-001 (PMS Tab)", type=["xlsx", "xls"])
with col2: logs_file = st.file_uploader("📥 Step 2: Upload Daily Operating Hours", type=["xlsx", "xls"])

if pms_file and logs_file:
    try:
        with st.status("Executing Bi-Directional Semantic Audit...", expanded=True) as status:
            
            st.write("⚙️ Routing PMS through Semantic Parser...")
            pms_df = semantic_pms_parse(pms_file.getvalue())
            
            st.write("⚙️ Routing Daily Logs through Semantic Parser...")
            daily_df = semantic_logs_parse(logs_file.getvalue())

            st.write("🔎 Phase 1: Validating Physics Constraints...")
            physics_violations = []
            for _, row in daily_df.iterrows():
                if row['ME_Hours'] > 24 or row['ME_Hours'] < 0:
                    physics_violations.append({
                        "Date": row['Date'].strftime('%d-%b-%Y'),
                        "System": "MAIN ENGINE",
                        "Logged Hours": row['ME_Hours'],
                        "Reason": "Violation: Hours exceed 24h limit or are negative."
                    })

            st.write("🧮 Phase 2: Computing Baseline Drift...")
            audit_results = []
            for _, row in pms_df.iterrows():
                comp = row['Component']
                oh_date = row['Last_Overhaul']
                legacy_hrs = row['Claimed_Hours']
                
                if pd.notnull(oh_date):
                    # Pure Math Vector: Sum hours where Log Date >= Overhaul Date
                    mask = daily_df['Date'] >= oh_date
                    verified_hrs = daily_df.loc[mask, 'ME_Hours'].sum()
                    delta = verified_hrs - legacy_hrs
                    
                    audit_results.append({
                        "Component": comp,
                        "Overhaul Date": oh_date.strftime('%d-%b-%Y'),
                        "Legacy Claim": int(legacy_hrs),
                        "Verified Math": int(verified_hrs),
                        "Delta": int(delta),
                        "Status": "VERIFIED" if int(delta) == 0 else "OVERWRITTEN"
                    })

            status.update(label="Audit Complete. Sealing Baseline.", state="complete", expanded=False)

        # ═══════════════════════════════════════════════════════════════════════════════
        # 4. DASHBOARD & EXPORT
        # ═══════════════════════════════════════════════════════════════════════════════
        if audit_results:
            res_df = pd.DataFrame(audit_results)
            errors_corrected = len(res_df[res_df['Delta'] != 0])
            
            # Cryptographic Seal
            seal_source = res_df.to_json(orient='records')
            digital_seal = hashlib.sha256(seal_source.encode()).hexdigest()

            st.subheader("2. Audit Results")
            c1, c2, c3 = st.columns(3)
            c1.metric("Components Audited", len(res_df))
            c2.metric("Drift Anomalies Corrected", errors_corrected, delta=f"{errors_corrected} Found", delta_color="inverse")
            c3.metric("SHA-256 Digital Seal", f"{digital_seal[:10]}...{digital_seal[-4:]}")

            st.divider()

            if physics_violations:
                st.subheader("⚠️ Phase 1: Kinematic Violations")
                st.error("The following entries in the Daily Logs violate the 24-hour limit.")
                st.dataframe(pd.DataFrame(physics_violations), use_container_width=True, hide_index=True)
                st.divider()

            st.subheader("📑 Phase 2: PMS Baseline Reconciliation")
            def style_dataframe(row):
                return ['background-color: #fff1f2; color: #9f1239'] * len(row) if row['Delta'] != 0 else [''] * len(row)
            
            st.dataframe(res_df.style.apply(style_dataframe, axis=1), use_container_width=True, hide_index=True)

            st.markdown("### 💾 Export Immutable Baseline")
            st.download_button(
                label="⬇️ Download Verified Baseline (CSV)",
                data=res_df.to_csv(index=False).encode('utf-8'),
                file_name=f"Verified_Baseline_{datetime.now().strftime('%Y%m%d')}.csv",
                mime='text/csv',
                type="primary"
            )
        else:
            st.warning("No valid component dates were extracted. Verify the TEC-001 format.")

    except Exception as e:
        st.error(f"🚨 Pipeline Semantic Crash Prevented: {str(e)}")
        st.info("The semantic parser could not locate the required columns. Please check the Excel structure.")
