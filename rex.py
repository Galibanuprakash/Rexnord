import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Coupling Selector Tool", layout="wide")
st.title("🔧 Coupling Selector Tool")

uploaded_file = st.file_uploader("📤 Upload Excel File", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file, sheet_name="Main-Data", header=1)
    df.columns = df.columns.str.strip().str.lower()  # Normalize column names

    # Coupling type options
    coupling_types = [
        "Marine Type", "Marine Type with Hydraulic Hub", "REM design", "REM Hydraulic Hub",
        "Coplaner design with Hydraulic hub (REM)", "Coplanar design with Marine Hub",
        "Coplaner With Single adaptor", "Yoke Design", "Adaptor Design",
        "Double adaptor Design", "Double Adaptor With Coplanar", "DIRECT CONNECTION"
    ]

    st.sidebar.header("🔎 Filter Options")
    driver_input = st.sidebar.text_input("Driver")
    driven_input = st.sidebar.text_input("Driven")
    model_input = st.sidebar.text_input("Coupling Model")

    power_min = st.sidebar.number_input("Min Power (kW)")
    power_max = st.sidebar.number_input("Max Power (kW)")
    speed_min = st.sidebar.number_input("Min Speed (RPM)")
    speed_max = st.sidebar.number_input("Max Speed (RPM)")
    dbse_input = st.sidebar.number_input("DBSE /DBFF (mm)", help="Optional: Searches ±10mm if value provided")

    driver_coupling_type = st.sidebar.selectbox("Driver Coupling Type (Optional)", [""] + coupling_types)
    driven_coupling_type = st.sidebar.selectbox("Driven Coupling Type (Optional)", [""] + coupling_types)

    if st.sidebar.button("🔍 Search Couplings"):
        df_filtered = df.copy()

        # Case-insensitive filtering
        if driver_input:
            df_filtered = df_filtered[df_filtered["driver"].astype(str).str.lower() == driver_input.lower()]
        if driven_input:
            df_filtered = df_filtered[df_filtered["driven"].astype(str).str.lower() == driven_input.lower()]
        if model_input:
            model_col = [col for col in df.columns if "coupling" in col and "model" in col]
            if model_col:
                df_filtered = df_filtered[df_filtered[model_col[0]].astype(str).str.lower() == model_input.lower()]

        # Optional coupling type filters
        if driver_coupling_type:
            driver_type_col = [col for col in df.columns if "driver coupling" in col][0]
            df_filtered = df_filtered[df_filtered[driver_type_col].astype(str).str.strip().str.lower() == driver_coupling_type.lower()]
        if driven_coupling_type:
            driven_type_col = [col for col in df.columns if "driven coupling" in col][0]
            df_filtered = df_filtered[df_filtered[driven_type_col].astype(str).str.strip().str.lower() == driven_coupling_type.lower()]

        # Convert to numeric for filtering
        df_filtered["power (kw)"] = pd.to_numeric(df_filtered["power (kw)"], errors="coerce")
        df_filtered["speed (rpm)"] = pd.to_numeric(df_filtered["speed (rpm)"], errors="coerce")
        df_filtered["dbse /dbff (mm)"] = pd.to_numeric(df_filtered["dbse /dbff (mm)"], errors="coerce")

        df_filtered = df_filtered[
            df_filtered["power (kw)"].between(power_min, power_max, inclusive="both") &
            df_filtered["speed (rpm)"].between(speed_min, speed_max, inclusive="both")
        ]

        # Optional DBSE ±10mm range
        if dbse_input > 0:
            df_filtered = df_filtered[df_filtered["dbse /dbff (mm)"].between(dbse_input - 10, dbse_input + 10)]

        # Replace '-' and blanks with NA
        df_filtered.replace(["-", "", " "], pd.NA, inplace=True)

        # Output Columns (base)
        output_columns = [
            "sl # / couplig #", "oem (buyer)", "drawing no", "driver", "driven",
            "coupling \nmodel", "pcd-1", "pcd-2",
            "power (kw)", "speed (rpm)", "cyclic torque requirement (yes / no)", "sct (knm)",
            "torsional stiffness (mnm/rad)", "dbse /dbff (mm)", "total weight (kg)"
        ]

        # Add conditional columns based on driver coupling type
        if driver_coupling_type in coupling_types[:6]:
            output_columns += [
                "driver connection type", "driver - if keyed type, single / double/taper ratio",
                "driver end shaft dia", "driver end hub boss dia",
                "driver hub pull-up distance (mm)", "driver shaft juncture capacity (knm)"
            ]
        elif driver_coupling_type in coupling_types[6:]:
            output_columns += [
                "driver side flange size- od", "driver side flange size- pcd",
                "driver side  flange - location size"
            ]

        # Add conditional columns based on driven coupling type
        if driven_coupling_type in coupling_types[:6]:
            output_columns += [
                "driven connection type", "driven - if keyed type, single / double/taper ratio",
                "driven end shaft dia", "driven end hub boss dia",
                "driven hub pull-up distance (mm)", "driven shaft juncture capacity (knm)"
            ]
        elif driven_coupling_type in coupling_types[6:]:
            output_columns += [
                "driven side flange size- od", "driven side flange size- pcd",
                "driven side  flange - location size"
            ]

        # Filter and clean output
        output_columns = [col for col in output_columns if col in df_filtered.columns]
        df_result = df_filtered[output_columns].dropna(how='all', axis=1).dropna(how='all')

        if df_result.empty:
            st.warning("❌ No matching records found.")
        else:
            st.success(f"✅ {len(df_result)} records found.")
            st.dataframe(df_result, use_container_width=True)

            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_result.to_excel(writer, index=False, sheet_name="Filtered")
            st.download_button(
                label="📥 Download Excel",
                data=output.getvalue(),
                file_name="filtered_couplings.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

else:
    st.info("⬆️ Upload an Excel file to begin.") 
