import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Coupling Selector Tool", layout="wide")
st.title("🔧 Euroflex Coupling Selector Tool")

uploaded_file = st.file_uploader("📤 Upload Excel File", type=["xlsx"])

driver_coupling_options = [
    "Marine Type", "Marine Type with Hydraulic Hub", "REM design", "REM Hydraulic Hub",
    "Coplaner design with Hydraulic hub (REM)", "Coplanar design with Marine Hub"
]
driver_flange_options = [
    "Coplaner With Single adaptor", "Yoke Design", "Adaptor Design",
    "Double adaptor Design", "Double Adaptor With Coplanar", "DIRECT CONNECTION"
]
driven_coupling_options = driver_coupling_options + driver_flange_options

if uploaded_file:
    df = pd.read_excel(uploaded_file, sheet_name="Main-Data", header=1)
    df.columns = df.columns.str.strip()

    # Sidebar filters
    st.sidebar.header("🔍 Filter Inputs")
    driver = st.sidebar.text_input("Driver")
    driven = st.sidebar.text_input("Driven")
    model = st.sidebar.text_input("Coupling Model")

    power_min = st.sidebar.number_input("Min Power (kW)")
    power_max = st.sidebar.number_input("Max Power (kW)")
    speed_min = st.sidebar.number_input("Min Speed (RPM)")
    speed_max = st.sidebar.number_input("Max Speed (RPM)")

    dbse_center = st.sidebar.number_input("DBSE /DBFF (mm)", help="Optional. Matches ±10 mm range if provided.")

    driver_coupling_type = st.sidebar.selectbox("Driver Coupling Type", [""] + driver_coupling_options + driver_flange_options)
    driven_coupling_type = st.sidebar.selectbox("Driven Coupling Type", [""] + driven_coupling_options)

    if st.sidebar.button("🔍 Search Couplings"):
        df_filtered = df.copy()

        # Case-insensitive exact match filters
        if driver:
            df_filtered = df_filtered[df_filtered["Driver"].astype(str).str.lower() == driver.lower()]
        if driven:
            df_filtered = df_filtered[df_filtered["Driven"].astype(str).str.lower() == driven.lower()]
        if model:
            df_filtered = df_filtered[df_filtered["Coupling \nModel"].astype(str).str.lower() == model.lower()]

        # Numeric filtering
        for col in ["Power (kW)", "Speed (RPM)", "DBSE /DBFF (mm)"]:
            df_filtered[col] = pd.to_numeric(df_filtered[col], errors="coerce")

        df_filtered = df_filtered[
            df_filtered["Power (kW)"].between(power_min, power_max) &
            df_filtered["Speed (RPM)"].between(speed_min, speed_max)
        ]

        # Optional DBSE / DBFF filter ±10 mm range
        if dbse_center > 0:
            df_filtered = df_filtered[
                df_filtered["DBSE /DBFF (mm)"].between(dbse_center - 10, dbse_center + 10)
            ]

        # Replace '-' and empty with NaN
        df_filtered.replace("-", pd.NA, inplace=True)
        df_filtered.replace("", pd.NA, inplace=True)

        # Output columns always included
        output_columns = [
            "Sl # / Couplig #", "OEM (Buyer)", "Drawing \nno", "Driver", "Driven",
             "Coupling \nModel", "Power (kW)", "Speed (RPM)",
            "Cyclic Torque requirement (yes / No)", "SCT (kNm)", "Torsional Stiffness (MNm/rad)",
            "DBSE /DBFF (mm)", "Total Weight\n(Kg)", "PCD-1", "PCD-2"
        ]

        # Driver coupling logic
        if driver_coupling_type in driver_coupling_options:
            output_columns += [
                "Driver Connection Type \n(taper/keyed/Angled/ counterbore /stepped /other)",
                "Driver - If keyed type, Single / double/taper ratio",
                "Driver End shaft dia", "Driver End hub Boss dia",
                "Driver Hub Pull-up distance (mm)", "Driver Shaft Juncture Capacity (kNm)"
            ]
        elif driver_coupling_type in driver_flange_options:
            output_columns += [
                "Driver side Flange size- OD", "Driver side Flange size- PCD",
                "Driver side  Flange - Location size"
            ]

        # Driven coupling logic
        if driven_coupling_type in driver_coupling_options:
            output_columns += [
                "Driven Connection Type (taper/keyed/Angled/ counterbore /stepped /other)",
                "Driven - If keyed type, Single / double/taper ratio",
                "Driven End shaft dia", "Driven Hub boss diameter",
                "Driven Hub Pull-up distance (mm)", "Driven Shaft Juncture Capacity (kNm)"
            ]
        elif driven_coupling_type in driver_flange_options:
            output_columns += [
                "Driven side Flange size- OD", "Driven side Flange size- PCD",
                "Driven side  Flange - Location size"
            ]

        # Filter to only columns present in file
        output_columns = [col for col in output_columns if col in df_filtered.columns]

        # Filter out rows with no data
        df_result = df_filtered[output_columns].dropna(how="all")
        df_result.dropna(axis=1, how="all", inplace=True)

        if df_result.empty:
            st.warning("❌ No matching records found.")
        else:
            st.success(f"✅ {len(df_result)} records found.")
            st.dataframe(df_result, use_container_width=True)

            output = BytesIO()
            with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
                df_result.to_excel(writer, index=False, sheet_name="Filtered")

            st.download_button(
                label="📥 Download Filtered Couplings",
                data=output.getvalue(),
                file_name="filtered_couplings.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
else:
    st.info("⬆️ Please upload an Excel file to begin.")
