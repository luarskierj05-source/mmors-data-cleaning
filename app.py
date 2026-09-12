import pandas as pd
import numpy as np
import plotly.express as px
import streamlit as st
import os
import re

# ==========================================
# 1. GLOBAL CONSTANTS & PAGE SETUP
# ==========================================
st.set_page_config(
    page_title="MMORS Data Warehouse & Analytics",
    page_icon="🌊",
    layout="wide"
)

st.title("🌊 Marilao, Meycauayan & Obando River System (MMORS)")
st.caption("Data Warehousing ETL & Descriptive Analytics Dashboard")

EXCEL_FILE = "mmors_data.xlsx"

TARGET_HEADERS = [
    "Station_No",
    "Location_Barangay",
    "Latitude, North (degree)",
    "Longitude, East (degree)",
    "Period_Year",
    "Date",
    "Time",
    "Dissolved Oxygen, mg/L",
    "PH",
    "Temperature °C",
    "Biochemical Oxygen Demand, mg/L",
    "Total Suspended Solids, mg/L",
    "Color TCU",
    "Fecal Coliform, MPN/100mL",
    "Total Coliform, MPN/100mL",
    "Ammonia, mg/L",
    "Nitrates as Nitrogen, mg/L",
    "Phosphates as Phosphorous, mg/L",
    "Chlorides Cl - (mg/L)"
]

# Map for month ordering
MONTH_MAP = {
    'JANUARY': 1, 'FEB': 2, 'FEBRUARY': 2, 'MARCH': 3, 'MAR': 3,
    'APRIL': 4, 'APR': 4, 'MAY': 5, 'JUNE': 6, 'JUN': 6,
    'JULY': 7, 'JUL': 7, 'AUGUST': 8, 'AUG': 8, 'SEPTEMBER': 9, 'SEP': 9,
    'OCTOBER': 10, 'OCT': 10, 'NOVEMBER': 11, 'NOV': 11, 'DECEMBER': 12, 'DEC': 12
}

# ==========================================
# 2. ETL PIPELINE
# ==========================================
@st.cache_data
def load_and_transform_data(file_name):
    if not os.path.exists(file_name):
        return None, f"Hindi mahanap ang '{file_name}' sa kasalukuyang folder ({os.getcwd()})."

    try:
        xls = pd.ExcelFile(file_name)
        sheet_names = xls.sheet_names
    except Exception as e:
        return None, f"Error sa pagbasa ng Excel file: {e}"

    all_data = []

    for sheet in sheet_names:
        sheet_str = str(sheet).strip()
        if sheet_str.lower().startswith("table") or sheet_str.lower().startswith("external"):
            continue

        df_raw = pd.read_excel(file_name, sheet_name=sheet, header=None)
        
        current_period = "CY 2012 JUNE"
        records = []

        for row_idx, row in df_raw.iterrows():
            row_str = " ".join([str(val) for val in row.values if pd.notna(val)])
            
            period_match = re.search(r'(CY\s*\d{4}\s*[A-Z]+)', row_str, re.IGNORECASE)
            if period_match:
                current_period = period_match.group(1).upper()
                continue
            
            col0_val = str(row.iloc[0]).strip() if len(row) > 0 else ""
            col1_val = str(row.iloc[1]).strip() if len(row) > 1 else ""

            is_valid_station = col0_val in ['1', '2', '3', '4', '5'] or col1_val in ['1', '2', '3', '4', '5']

            if is_valid_station:
                row_list = list(row.values)
                
                if col0_val in ['1', '2', '3', '4', '5']:
                    stn_no = col0_val
                    loc_name = row_list[1] if len(row_list) > 1 else ""
                    lat_val = row_list[2] if len(row_list) > 2 else ""
                    long_val = row_list[3] if len(row_list) > 3 else ""
                    rest_vals = row_list[4:]
                else:
                    stn_no = col1_val
                    loc_name = row_list[2] if len(row_list) > 2 else ""
                    lat_val = row_list[3] if len(row_list) > 3 else ""
                    long_val = row_list[4] if len(row_list) > 4 else ""
                    rest_vals = row_list[5:]

                full_record = [stn_no, loc_name, lat_val, long_val, current_period] + rest_vals
                
                if len(full_record) < len(TARGET_HEADERS):
                    full_record += [np.nan] * (len(TARGET_HEADERS) - len(full_record))
                else:
                    full_record = full_record[:len(TARGET_HEADERS)]
                
                records.append(full_record)

        if records:
            sheet_df = pd.DataFrame(records, columns=TARGET_HEADERS)
            sheet_df['River_Tab'] = sheet_str
            all_data.append(sheet_df)

    if not all_data:
        return None, "Walang valid data na na-extract mula sa mga sheets."

    unified_df = pd.concat(all_data, ignore_index=True)

    if 'Station_No' in unified_df.columns:
        unified_df['Station_No'] = (
            unified_df['Station_No']
            .astype(str)
            .str.extract(r'([1-5])')[0]
        )
        unified_df = unified_df[unified_df['Station_No'].notna()].copy()

    # Extract Year and Month Index for Range Filtering
    unified_df['Year_Only'] = unified_df['Period_Year'].astype(str).str.extract(r'(\d{4})')[0]
    
    def get_month_num(period_str):
        for month, num in MONTH_MAP.items():
            if month in str(period_str).upper():
                return num
        return 1

    unified_df['Month_Num'] = unified_df['Period_Year'].apply(get_month_num)
    
    # Sortable Period Index (e.g., 201203 for MARCH 2012)
    unified_df['Period_Order'] = (
        pd.to_numeric(unified_df['Year_Only'], errors='coerce').fillna(0).astype(int) * 100 
        + unified_df['Month_Num']
    )
    
    unified_df = unified_df.sort_values(by=['Period_Order', 'Station_No'])

    numeric_cols = [
        "Latitude, North (degree)", "Longitude, East (degree)", 
        "Dissolved Oxygen, mg/L", "PH", "Temperature °C", 
        "Biochemical Oxygen Demand, mg/L", "Total Suspended Solids, mg/L",
        "Color TCU", "Fecal Coliform, MPN/100mL", "Total Coliform, MPN/100mL",
        "Ammonia, mg/L", "Nitrates as Nitrogen, mg/L", 
        "Phosphates as Phosphorous, mg/L", "Chlorides Cl - (mg/L)"
    ]

    for col in numeric_cols:
        if col in unified_df.columns:
            unified_df[col] = pd.to_numeric(unified_df[col], errors='coerce')

    if 'Dissolved Oxygen, mg/L' in unified_df.columns:
        unified_df['DO_Compliance'] = np.where(unified_df['Dissolved Oxygen, mg/L'] >= 5.0, 'Compliant (≥5.0)', 'Non-Compliant (<5.0)')
    else:
        unified_df['DO_Compliance'] = "N/A"

    if 'Biochemical Oxygen Demand, mg/L' in unified_df.columns:
        unified_df['BOD_Compliance'] = np.where(unified_df['Biochemical Oxygen Demand, mg/L'] <= 7.0, 'Compliant (≤7.0)', 'Non-Compliant (>7.0)')
    else:
        unified_df['BOD_Compliance'] = "N/A"

    FINAL_EXACT_HEADERS = [
        "River_Tab",
        "Location_Barangay",
        "Station_No",
        "Latitude, North (degree)",
        "Longitude, East (degree)",
        "Period_Year",
        "Year_Only",
        "Date",
        "Time",
        "Dissolved Oxygen, mg/L",
        "PH",
        "Temperature °C",
        "Biochemical Oxygen Demand, mg/L",
        "Total Suspended Solids, mg/L",
        "Color TCU",
        "Fecal Coliform, MPN/100mL",
        "Total Coliform, MPN/100mL",
        "Ammonia, mg/L",
        "Nitrates as Nitrogen, mg/L",
        "Phosphates as Phosphorous, mg/L",
        "Chlorides Cl - (mg/L)",
        "DO_Compliance",
        "BOD_Compliance",
        "Period_Order"
    ]

    unified_df = unified_df.reindex(columns=FINAL_EXACT_HEADERS)

    return unified_df, None

df, err = load_and_transform_data(EXCEL_FILE)

if df is None:
    st.error(f"⚠️ {err}")
    st.stop()

# ==========================================
# 3. SIDEBAR CONTROLS
# ==========================================
st.sidebar.header("🗂️ Data Warehouse Filters")

# River System Filter
river_tabs = ["All River Systems"] + list(df['River_Tab'].dropna().unique())
selected_tab = st.sidebar.selectbox("Select Excel River Tab:", river_tabs)

filtered_df = df[df['River_Tab'] == selected_tab].copy() if selected_tab != "All River Systems" else df.copy()

# Unique period list for chronological range filtering
available_periods_df = (
    filtered_df[['Period_Year', 'Period_Order']]
    .dropna()
    .drop_duplicates()
    .sort_values(by='Period_Order')
)
period_list = available_periods_df['Period_Year'].tolist()

st.sidebar.subheader("📅 Period Range Filter")

if period_list:
    start_period = st.sidebar.selectbox("From Period:", period_list, index=0)
    
    start_order = available_periods_df[available_periods_df['Period_Year'] == start_period]['Period_Order'].values[0]
    valid_end_periods = available_periods_df[available_periods_df['Period_Order'] >= start_order]['Period_Year'].tolist()
    
    end_period = st.sidebar.selectbox("To Period:", valid_end_periods, index=len(valid_end_periods)-1)
    
    end_order = available_periods_df[available_periods_df['Period_Year'] == end_period]['Period_Order'].values[0]

    view_df = filtered_df[
        (filtered_df['Period_Order'] >= start_order) & 
        (filtered_df['Period_Order'] <= end_order)
    ].copy()
else:
    start_period = "N/A"
    end_period = "N/A"
    view_df = filtered_df.copy()

app_mode = st.sidebar.radio(
    "Go to Module:",
    ["1. Data Warehouse & Pre-Processing (ETL)", 
     "2. Descriptive Analytics & Visualizations", 
     "3. Outcome Interpretation & DENR Compliance"]
)

# ==========================================
# 4. MODULE 1: DATA WAREHOUSE VIEW & CSV EXPORT
# ==========================================
if app_mode == "1. Data Warehouse & Pre-Processing (ETL)":
    st.subheader("🛠️ Data Warehouse Table")
    
    m1, m2, m3 = st.columns(3)
    m1.metric("Total Rows Displayed", len(view_df))
    m2.metric("Total Warehouse Attributes", len(view_df.columns))
    m3.metric("Selected Range", f"{start_period} ➔ {end_period}" if period_list else "All")

    st.dataframe(view_df.drop(columns=['Period_Order'], errors='ignore'), use_container_width=True, height=450)

    # 📥 DOWNLOAD CLEANED DATASET (.CSV)
    csv_data = view_df.drop(columns=['Period_Order'], errors='ignore').to_csv(index=False).encode('utf-8')
    st.download_button(
        label="📥 Download Cleaned MMORS Dataset (.CSV)",
        data=csv_data,
        file_name=f"MMORS_Cleaned_Data_{start_period}_to_{end_period}.csv",
        mime="text/csv",
        help="Click to download the preprocessed and clean analysis-ready dataset."
    )

# ==========================================
# 5. MODULE 2: VISUAL ANALYTICS
# ==========================================
elif app_mode == "2. Descriptive Analytics & Visualizations":
    st.subheader(f"📊 Descriptive Analytics ({selected_tab} | Range: {start_period} to {end_period})")

    k1, k2, k3, k4 = st.columns(4)
    
    # DO Metric
    do_series = view_df['Dissolved Oxygen, mg/L'].dropna()
    if not do_series.empty:
        avg_do = do_series.mean()
        do_delta = "Pass (≥ 5.0)" if avg_do >= 5.0 else "Fail (< 5.0)"
        do_color = "normal" if avg_do >= 5.0 else "inverse"
        k1.metric("Avg Dissolved Oxygen", f"{avg_do:.2f} mg/L", delta=f"{do_delta}", delta_color=do_color)
    else:
        k1.metric("Avg Dissolved Oxygen", "N/A")

    # BOD Metric
    bod_series = view_df['Biochemical Oxygen Demand, mg/L'].dropna()
    if not bod_series.empty:
        avg_bod = bod_series.mean()
        bod_delta = "Pass (≤ 7.0)" if avg_bod <= 7.0 else "Fail (> 7.0)"
        bod_color = "normal" if avg_bod <= 7.0 else "inverse"
        k2.metric("Avg BOD Level", f"{avg_bod:.2f} mg/L", delta=f"{bod_delta}", delta_color=bod_color)
    else:
        k2.metric("Avg BOD Level", "N/A")

    # Fecal Coliform Metric
    col_series = view_df['Fecal Coliform, MPN/100mL'].dropna()
    col_series = col_series[col_series > 0]
    if not col_series.empty:
        geomean_col = np.exp(np.log(col_series).mean())
        col_delta = "Pass (≤ 200)" if geomean_col <= 200 else "Fail (> 200)"
        col_color = "normal" if geomean_col <= 200 else "inverse"
        k3.metric("Geomean Fecal Coliform", f"{int(geomean_col):,} MPN", delta=f"{col_delta}", delta_color=col_color)
    else:
        k3.metric("Geomean Fecal Coliform", "N/A")

    # pH Metric
    ph_series = view_df['PH'].dropna()
    if not ph_series.empty:
        avg_ph = ph_series.mean()
        ph_ok = 6.5 <= avg_ph <= 9.0
        ph_delta = "Pass (6.5-9.0)" if ph_ok else "Out of Range"
        ph_color = "normal" if ph_ok else "inverse"
        k4.metric("Avg pH Level", f"{avg_ph:.2f}", delta=f"{ph_delta}", delta_color=ph_color)
    else:
        k4.metric("Avg pH Level", "N/A")

    st.markdown("---")

    EXCLUDE_FROM_PARAMS = [
        'Station_No', 'River_Tab', 'Location_Barangay', 'Period_Year', 'Year_Only',
        'Date', 'Time', 'DO_Compliance', 'BOD_Compliance', 'Period_Order',
        'Latitude, North (degree)', 'Longitude, East (degree)'
    ]
    
    water_parameters = [c for c in view_df.columns if c not in EXCLUDE_FROM_PARAMS]

    c1, c2 = st.columns(2)
    with c1:
        st.subheader("📈 Parameter Trend Across Selected Range")
        if water_parameters:
            selected_param = st.selectbox("Select Water Parameter:", water_parameters)
            fig_line = px.line(
                view_df, 
                x='Period_Year', 
                y=selected_param, 
                color='Location_Barangay', 
                markers=True, 
                title=f"{selected_param} Trend ({start_period} - {end_period})",
                template="plotly_dark"
            )
            fig_line.update_layout(xaxis_title="Period / Year", yaxis_title=selected_param, legend_title="Location")
            st.plotly_chart(fig_line, use_container_width=True)

    with c2:
        st.subheader("📊 Station Comparison")
        if water_parameters and selected_param in view_df.columns:
            fig_bar = px.bar(
                view_df, 
                x='Location_Barangay', 
                y=selected_param, 
                color='Station_No',
                barmode='group',
                title=f"{selected_param} by Station",
                template="plotly_dark"
            )
            fig_bar.update_layout(xaxis_title="Location Barangay", yaxis_title=selected_param, xaxis_tickangle=-30)
            st.plotly_chart(fig_bar, use_container_width=True)

# ==========================================
# 6. MODULE 3: OUTCOME INTERPRETATION & DENR COMPLIANCE
# ==========================================
elif app_mode == "3. Outcome Interpretation & DENR Compliance":
    st.subheader(f"💡 DENR Class C Water Quality Compliance Report ({start_period} to {end_period})")
    
    total_samples = len(view_df)
    
    do_non_compliant = len(view_df[view_df['DO_Compliance'] == 'Non-Compliant (<5.0)'])
    do_fail_rate = (do_non_compliant / total_samples * 100) if total_samples > 0 else 0
    
    bod_non_compliant = len(view_df[view_df['BOD_Compliance'] == 'Non-Compliant (>7.0)'])
    bod_fail_rate = (bod_non_compliant / total_samples * 100) if total_samples > 0 else 0

    if do_fail_rate > 50 or bod_fail_rate > 50:
        st.error(f"⚠️ **CRITICAL WATER QUALITY STATUS**: Over 50% non-compliance detected for DENR Class C Standards in the selected range.")
    else:
        st.success("✅ **MODERATE TO GOOD STATUS**: Water parameters mostly comply with DENR DAO 2016-08 Class C Standards.")

    st.markdown("---")

    col1, col2 = st.columns(2)
    with col1:
        st.markdown("### Dissolved Oxygen (DO) Compliance")
        st.caption("Standard: ≥ 5.0 mg/L (Higher is better)")
        fig_do = px.pie(
            view_df, 
            names='DO_Compliance', 
            color='DO_Compliance',
            color_discrete_map={'Compliant (≥5.0)': '#2ecc71', 'Non-Compliant (<5.0)': '#e74c3c'},
            hole=0.4
        )
        st.plotly_chart(fig_do, use_container_width=True)

    with col2:
        st.markdown("### Biochemical Oxygen Demand (BOD) Compliance")
        st.caption("Standard: ≤ 7.0 mg/L (Lower is better)")
        fig_bod = px.pie(
            view_df, 
            names='BOD_Compliance', 
            color='BOD_Compliance',
            color_discrete_map={'Compliant (≤7.0)': '#2ecc71', 'Non-Compliant (>7.0)': '#e74c3c'},
            hole=0.4
        )
        st.plotly_chart(fig_bod, use_container_width=True)

    st.markdown("---")

    st.markdown("### 📋 Executive Summary & Recommendations")

    with st.expander("📌 Detailed Outcome Interpretation", expanded=True):
        st.markdown(f"""
        * **Dissolved Oxygen (DO) Assessment:** 
          * Out of **{total_samples}** samples, **{do_fail_rate:.1f}%** failed to meet the required limit ($\ge 5.0\\text{{ mg/L}}$).
          * *Environmental Meaning:* Depleted oxygen levels threaten aquatic life and indicate high pollution loads in the river system.
        
        * **Biochemical Oxygen Demand (BOD) Assessment:** 
          * **{bod_fail_rate:.1f}%** of samples exceeded the allowable maximum limit ($\le 7.0\\text{{ mg/L}}$).
          * *Environmental Meaning:* High BOD signals excessive organic waste contamination, largely caused by untreated domestic wastewater and commercial runoffs.

        * **Policy & Action Items:**
          * Enforce strict inspection and monitoring of commercial and LGU wastewater discharge permits near non-compliant stations.
          * Accelerate the construction of community Septage Treatment Plants (STP) along the MMORS river basin.
        """)


##streamlit run app.py