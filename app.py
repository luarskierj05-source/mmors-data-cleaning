import pandas as pd
import numpy as np
import plotly.express as px
import streamlit as st
import streamlit.components.v1 as components
import os
import re

# ==========================================
# 1. PAGE SETUP & GLOBAL CONSTANTS
# ==========================================
st.set_page_config(page_title="MMORS Analytics & Presentation", page_icon="🌊", layout="wide")

EXCEL_FILE = "mmors_data.xlsx"

TARGET_HEADERS = [
    "Station_No", "Location_Barangay", "Latitude, North (degree)", "Longitude, East (degree)",
    "Period_Year", "Date", "Time", "Dissolved Oxygen, mg/L", "PH", "Temperature °C",
    "Biochemical Oxygen Demand, mg/L", "Total Suspended Solids, mg/L", "Color TCU",
    "Fecal Coliform, MPN/100mL", "Total Coliform, MPN/100mL", "Ammonia, mg/L",
    "Nitrates as Nitrogen, mg/L", "Phosphates as Phosphorous, mg/L", "Chlorides Cl - (mg/L)"
]

MONTH_MAP = {
    'JANUARY': 1, 'FEBRUARY': 2, 'MARCH': 3, 'APRIL': 4, 'MAY': 5, 'JUNE': 6,
    'JULY': 7, 'AUGUST': 8, 'SEPTEMBER': 9, 'OCTOBER': 10, 'NOVEMBER': 11, 'DECEMBER': 12
}

# Helper function to read raw sample sheet for "Before" view
@st.cache_data
def load_raw_sample(file_name):
    if not os.path.exists(file_name):
        return None
    try:
        xls = pd.ExcelFile(file_name)
        # Read the first sheet as-is (uncleaned raw data)
        raw_df = pd.read_excel(file_name, sheet_name=xls.sheet_names[0], header=None)
        return raw_df
    except Exception:
        return None

# ==========================================
# 2. ETL ENGINE
# ==========================================
@st.cache_data
def load_and_transform_data(file_name):
    if not os.path.exists(file_name):
        return None, f"File '{file_name}' not found."
    
    try:
        xls = pd.ExcelFile(file_name)
        all_data = []
        for sheet in xls.sheet_names:
            if not (sheet.lower().startswith("table") or sheet.lower().startswith("external")):
                df_raw = pd.read_excel(file_name, sheet_name=sheet, header=None)
                current_period = "CY 2012 JUNE"
                for _, row in df_raw.iterrows():
                    row_str = " ".join([str(v) for v in row.values if pd.notna(v)])
                    p_match = re.search(r'(CY\s*\d{4}\s*[A-Z]+)', row_str, re.IGNORECASE)
                    if p_match:
                        current_period = p_match.group(1).upper()
                        continue
                    
                    c0, c1 = str(row.iloc[0]).strip(), str(row.iloc[1]).strip()
                    if c0 in ['1','2','3','4','5'] or c1 in ['1','2','3','4','5']:
                        row_list = list(row.values)
                        if c0 in ['1','2','3','4','5']:
                            rec = [c0, row_list[1], row_list[2], row_list[3], current_period] + row_list[4:]
                        else:
                            rec = [c1, row_list[2], row_list[3], row_list[4], current_period] + row_list[5:]
                        all_data.append(rec[:len(TARGET_HEADERS)])

        unified_df = pd.DataFrame(all_data, columns=TARGET_HEADERS)
        
        unified_df['Year_Only'] = unified_df['Period_Year'].astype(str).str.extract(r'(\d{4})')[0]
        unified_df['Month_Num'] = unified_df['Period_Year'].apply(lambda x: next((v for k,v in MONTH_MAP.items() if k in str(x).upper()), 1))
        unified_df['Period_Order'] = pd.to_numeric(unified_df['Year_Only'], errors='coerce').fillna(0).astype(int)*100 + unified_df['Month_Num']
        
        num_cols = ["Dissolved Oxygen, mg/L", "Biochemical Oxygen Demand, mg/L", "PH", "Fecal Coliform, MPN/100mL"]
        for col in num_cols: unified_df[col] = pd.to_numeric(unified_df[col], errors='coerce')
        
        unified_df['DO_Compliance'] = np.where(unified_df['Dissolved Oxygen, mg/L'] >= 5.0, 'Compliant', 'Non-Compliant')
        unified_df['BOD_Compliance'] = np.where(unified_df['Biochemical Oxygen Demand, mg/L'] <= 7.0, 'Compliant', 'Non-Compliant')
        
        return unified_df.sort_values('Period_Order'), None
    except Exception as e:
        return None, str(e)

df, err = load_and_transform_data(EXCEL_FILE)
df_raw_sample = load_raw_sample(EXCEL_FILE)

# ==========================================
# 3. NAVIGATION & SIDEBAR
# ==========================================
st.sidebar.title("🌊 MMORS Portal")
app_mode = st.sidebar.radio("Navigate to:", [
    "0. Group Presentation", 
    "1. Data Warehouse & Pre-Processing", 
    "2. Descriptive Analytics", 
    "3. DENR Compliance Report"
])

if df is None:
    st.error(err)
    st.stop()

# ==========================================
# MODULE 0: HTML PRESENTATION
# ==========================================
if app_mode == "0. Group Presentation":
    st.header("📽️ Project Presentation")
    
    presentation_html = """
    <div style="background: #1e293b; color: white; padding: 30px; border-radius: 16px; font-family: 'Segoe UI', sans-serif; position: relative; min-height: 480px; box-shadow: 0 10px 25px rgba(0,0,0,0.3);">
        
        <!-- Navigation Buttons at Top Right -->
        <div style="position: absolute; top: 25px; right: 30px; z-index: 100;">
            <button onclick="change(-1)" style="padding: 10px 18px; border-radius: 8px; cursor: pointer; background: #334155; color: #f8fafc; border: 1px solid #475569; font-weight: 600; margin-right: 8px;">◀ Prev</button>
            <button onclick="change(1)" style="padding: 10px 18px; border-radius: 8px; cursor: pointer; background: #38bdf8; color: #0f172a; border: none; font-weight: 700;">Next ▶</button>
        </div>

        <div id="slides" style="padding-top: 10px;">
            <div class="slide" id="slide0">
                <h1 style="color: #38bdf8; font-size: 2.3rem; margin-bottom: 10px; margin-top: 0;">MMORS Data Preprocessing</h1>
                <p style="font-size: 1.2rem; color: #cbd5e1; margin-bottom: 25px;">Water Quality Analysis & Analytics (2012–2018)</p>
                <div style="background: #0f172a; padding: 20px; border-radius: 12px; border-left: 4px solid #38bdf8; max-width: 600px;">
                    <p style="margin-top: 0; margin-bottom: 12px; font-size: 1.1rem;"><b>👥 Group Members:</b></p>
                    <ul style="margin: 0; padding-left: 20px; color: #f8fafc; line-height: 1.8; font-size: 1.05rem;">
                        <li>Agustin V. Cabrera</li>
                        <li>Junralf Gedorio</li>
                        <li>Suzzette Castro</li>
                        <li>Roselyn Luar</li>
                    </ul>
                </div>
            </div>
            <div class="slide" id="slide1" style="display:none;">
                <h2 style="color: #38bdf8; margin-top: 0;">Why Python & Pandas?</h2>
                <div style="font-size: 1.1rem; line-height: 1.8;">
                    <p>1. <b>Automated ETL:</b> Handles 7 years of inconsistent multi-sheet data seamlessly.</p>
                    <p>2. <b>Regex Extraction:</b> Programmatically parses sampling dates from embedded text rows.</p>
                    <p>3. <b>Data Integrity:</b> Standardizes mixed data types into a clean warehouse repository.</p>
                </div>
            </div>
            <div class="slide" id="slide2" style="display:none;">
                <h2 style="color: #38bdf8; margin-top: 0;">Data Quality Issues Identified</h2>
                <ul style="font-size: 1.1rem; line-height: 1.8; padding-left: 20px;">
                    <li>Fragmented river worksheets and irrelevant metadata tabs.</li>
                    <li>Embedded period headers (e.g., "CY 2012 JUNE") hidden inside rows.</li>
                    <li>Non-numeric contamination in BOD and Coliform values.</li>
                </ul>
            </div>
            <div class="slide" id="slide3" style="display:none;">
                <h2 style="color: #38bdf8; margin-top: 0;">Live Dashboard Outcome</h2>
                <div style="font-size: 1.1rem; line-height: 1.8;">
                    <p>• <b>Module 1:</b> Cleaned Warehouse & CSV Export functionality.</p>
                    <p>• <b>Module 2:</b> Interactive Trends & Parameter Analytics.</p>
                    <p>• <b>Module 3:</b> DENR Class C Compliance Evaluation.</p>
                </div>
            </div>
        </div>
    </div>
    <script>
        let cur = 0; const s = document.querySelectorAll('.slide');
        function change(n) {
            s[cur].style.display = 'none';
            cur = (cur + n + s.length) % s.length;
            s[cur].style.display = 'block';
        }
    </script>
    """
    components.html(presentation_html, height=600)

# ==========================================
# MODULE 1: DATA WAREHOUSE & EXPORT (SIDE-BY-SIDE VIEW)
# ==========================================
elif app_mode == "1. Data Warehouse & Pre-Processing":
    st.header("🛠️ Data Pre-Processing Transformation")
    st.caption("Side-by-side comparison of Raw Unstructured Data vs. Structured Data Warehouse")

    col_before, col_after = st.columns(2)

    with col_before:
        st.subheader("🔴 BEFORE: Raw Excel Data")
        st.warning("Issues: Unstructured headers, embedded period rows, column shifts, missing schema.")
        if df_raw_sample is not None:
            st.dataframe(df_raw_sample.head(25), use_container_width=True, height=450)
        else:
            st.info("Raw data preview unavailable.")

    with col_after:
        st.subheader("🟢 AFTER: Cleaned Data Warehouse")
        st.success("Cleaned: Standardized columns, parsed period/year, numeric casting, ready for analytics.")
        st.dataframe(df.drop(columns=['Period_Order']), use_container_width=True, height=450)

    st.markdown("---")
    
    # Download Button
    csv = df.to_csv(index=False).encode('utf-8')
    st.download_button(
        label="📥 Download Cleaned Dataset (.CSV)",
        data=csv,
        file_name="MMORS_Cleaned_Data_2012_2018.csv",
        mime="text/csv"
    )

# ==========================================
# MODULE 2: ANALYTICS
# ==========================================
elif app_mode == "2. Descriptive Analytics":
    st.header("📊 Interactive Visualizations")
    param = st.selectbox("Select Parameter to Visualize:", ["Dissolved Oxygen, mg/L", "Biochemical Oxygen Demand, mg/L", "PH"])
    fig = px.line(df, x="Period_Year", y=param, color="Location_Barangay", markers=True, title=f"Trend of {param} (2012-2018)")
    st.plotly_chart(fig, use_container_width=True)

# ==========================================
# MODULE 3: COMPLIANCE
# ==========================================
else:
    st.header("💡 DENR Compliance Report")
    col1, col2 = st.columns(2)
    with col1:
        st.write("### Dissolved Oxygen (Standard: ≥ 5.0)")
        st.plotly_chart(px.pie(df, names="DO_Compliance", hole=0.4, color_discrete_sequence=['#2ecc71', '#e74c3c']), use_container_width=True)
    with col2:
        st.write("### BOD (Standard: ≤ 7.0)")
        st.plotly_chart(px.pie(df, names="BOD_Compliance", hole=0.4, color_discrete_sequence=['#2ecc71', '#e74c3c']), use_container_width=True)
