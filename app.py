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
RAW_IMAGE_FILE = "raw_excel_sample.png"

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
# MODULE 0: HTML PRESENTATION (8 SLIDES WITH SLIDE NUMBER)
# ==========================================
if app_mode == "0. Group Presentation":
    st.header("📽️ Project Presentation")
    
    presentation_html = """
    <div style="background: #1e293b; color: white; padding: 35px; border-radius: 16px; font-family: 'Segoe UI', sans-serif; position: relative; min-height: 500px; box-shadow: 0 10px 25px rgba(0,0,0,0.3);">
        
        <!-- Navigation Buttons & Slide Counter at Top Right -->
        <div style="position: absolute; top: 25px; right: 30px; z-index: 100; display: flex; align-items: center; gap: 12px;">
            <span id="slideCounter" style="background: #0f172a; padding: 8px 14px; border-radius: 20px; font-weight: 600; color: #38bdf8; font-size: 0.95rem; border: 1px solid #334155;">Slide 1 of 8</span>
            <button onclick="change(-1)" style="padding: 10px 18px; border-radius: 8px; cursor: pointer; background: #334155; color: #f8fafc; border: 1px solid #475569; font-weight: 600;">◀ Prev</button>
            <button onclick="change(1)" style="padding: 10px 18px; border-radius: 8px; cursor: pointer; background: #38bdf8; color: #0f172a; border: none; font-weight: 700;">Next ▶</button>
        </div>

        <div id="slides" style="padding-top: 10px;">
            <!-- Slide 1 -->
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

            <!-- Slide 2 -->
            <div class="slide" id="slide1" style="display:none;">
                <h2 style="color: #38bdf8; margin-top: 0;">Project Background & Context</h2>
                <div style="font-size: 1.1rem; line-height: 1.8; color: #e2e8f0;">
                    <p>• <b>Scope:</b> Meycauayan-Marilao-Obando River System (MMORS) Water Quality Data.</p>
                    <p>• <b>Time Horizon:</b> Multi-year monitoring dataset spanning from <b>2012 to 2018</b>.</p>
                    <p>• <b>Objective:</b> Clean, standardize, and build an automated analytics engine to assess river water safety against national DENR standards.</p>
                </div>
            </div>

            <!-- Slide 3 -->
            <div class="slide" id="slide2" style="display:none;">
                <h2 style="color: #38bdf8; margin-top: 0;">Data Quality & Structure Issues</h2>
                <ul style="font-size: 1.1rem; line-height: 1.8; padding-left: 20px; color: #e2e8f0;">
                    <li><b>Fragmented Worksheets:</b> Data separated across multiple Excel tabs and river sub-tables.</li>
                    <li><b>Embedded Row Headers:</b> Period indicators (e.g., "CY 2012 JUNE") hidden inside data rows.</li>
                    <li><b>Inconsistent Schema:</b> Varying column positions, merged cells, and non-standard field names.</li>
                    <li><b>Dirty Values:</b> Mixed text and non-numeric contamination in parameter columns.</li>
                </ul>
            </div>

            <!-- Slide 4 -->
            <div class="slide" id="slide3" style="display:none;">
                <h2 style="color: #38bdf8; margin-top: 0;">Why Python & Pandas ETL?</h2>
                <div style="font-size: 1.1rem; line-height: 1.8; color: #e2e8f0;">
                    <p>1. <b>Automated Ingestion:</b> Iterates through all sheets without manual copying.</p>
                    <p>2. <b>Regex Date Parsing:</b> Automatically extracts sampling periods from embedded text.</p>
                    <p>3. <b>Schema Standardization:</b> Maps messy columns to a clean 19-header target schema.</p>
                    <p>4. <b>Data Type Coercion:</b> Safely cleans non-numeric values for calculation.</p>
                </div>
            </div>

            <!-- Slide 5 -->
            <div class="slide" id="slide4" style="display:none;">
                <h2 style="color: #38bdf8; margin-top: 0;">Data Transformation Workflow</h2>
                <div style="background: #0f172a; padding: 20px; border-radius: 12px; font-size: 1.05rem; line-height: 1.8;">
                    <p style="margin: 0 0 10px 0;"><b>Step 1:</b> Scan and identify valid monitoring data sheets.</p>
                    <p style="margin: 0 0 10px 0;"><b>Step 2:</b> Parse and carry down sampling period context (e.g., CY 2012).</p>
                    <p style="margin: 0 0 10px 0;"><b>Step 3:</b> Align coordinates, barangay locations, and physical/chemical parameters.</p>
                    <p style="margin: 0;"><b>Step 4:</b> Export clean dataset to structured Data Warehouse (.CSV).</p>
                </div>
            </div>

            <!-- Slide 6 -->
            <div class="slide" id="slide5" style="display:none;">
                <h2 style="color: #38bdf8; margin-top: 0;">Key Environmental Parameters Analyzed</h2>
                <ul style="font-size: 1.1rem; line-height: 1.8; padding-left: 20px; color: #e2e8f0;">
                    <li><b>Dissolved Oxygen (DO):</b> Vital for aquatic life survival.</li>
                    <li><b>Biochemical Oxygen Demand (BOD):</b> Indicates organic pollution level.</li>
                    <li><b>pH Level:</b> Measures water acidity/alkalinity balance.</li>
                    <li><b>Coliform Count:</b> Indicates bacterial contamination levels.</li>
                </ul>
            </div>

            <!-- Slide 7 -->
            <div class="slide" id="slide6" style="display:none;">
                <h2 style="color: #38bdf8; margin-top: 0;">DENR Class C Compliance Standards</h2>
                <div style="font-size: 1.1rem; line-height: 1.8; color: #e2e8f0;">
                    <p>• <b>Class C Water Purpose:</b> Fishery, Recreation, & Industrial Water Supply Class II.</p>
                    <p>• <b>Dissolved Oxygen Standard:</b> Must be <b>≥ 5.0 mg/L</b> (Compliant if high).</p>
                    <p>• <b>BOD Standard:</b> Must be <b>≤ 7.0 mg/L</b> (Compliant if low).</p>
                </div>
            </div>

            <!-- Slide 8 -->
            <div class="slide" id="slide7" style="display:none;">
                <h2 style="color: #38bdf8; margin-top: 0;">Live Dashboard Modules</h2>
                <div style="font-size: 1.1rem; line-height: 1.8; color: #e2e8f0;">
                    <p>• <b>Module 1:</b> Data Warehouse (Before/After Side-by-Side & CSV Download)</p>
                    <p>• <b>Module 2:</b> Descriptive Analytics (Interactive Plotly Trends 2012–2018)</p>
                    <p>• <b>Module 3:</b> DENR Compliance Report (Interactive Compliance Pie Charts)</p>
                    <p style="color: #38bdf8; font-weight: bold; margin-top: 15px;">👉 Switch to Module 1 on the sidebar to view live demo!</p>
                </div>
            </div>
        </div>
    </div>

    <script>
        let cur = 0; 
        const s = document.querySelectorAll('.slide');
        const counter = document.getElementById('slideCounter');
        
        function change(n) {
            s[cur].style.display = 'none';
            cur = (cur + n + s.length) % s.length;
            s[cur].style.display = 'block';
            counter.innerText = 'Slide ' + (cur + 1) + ' of ' + s.length;
        }
    </script>
    """
    components.html(presentation_html, height=600)

# ==========================================
# MODULE 1: DATA WAREHOUSE & EXPORT (IMAGE SIDE-BY-SIDE VIEW)
# ==========================================
elif app_mode == "1. Data Warehouse & Pre-Processing":
    st.header("🛠️ Data Pre-Processing Transformation")
    st.caption("Side-by-side comparison of Raw Unstructured Data vs. Structured Data Warehouse")

    col_before, col_after = st.columns(2)

    with col_before:
        st.subheader("🔴 BEFORE: Raw Excel Sheet")
        st.warning("Issues: Unstructured headers, embedded period rows, merged cells, and multiple tabs.")
        
        if os.path.exists(RAW_IMAGE_FILE):
            st.image(RAW_IMAGE_FILE, caption="Uncleaned DENR EMB Water Quality Excel File", use_container_width=True)
        else:
            st.info(f"Pakisave ang screenshot bilang `{RAW_IMAGE_FILE}` sa iyong GitHub repository para lumabas ang larawan dito.")

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
