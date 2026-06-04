import streamlit as st
import pandas as pd
import numpy as np
import os
import plotly.graph_objects as go
from scipy.stats import gmean

# ตั้งค่าหน้าจอ Web-App
st.set_page_config(page_title="QCF Titer Dashboard 2026", layout="wide")
st.title("📊 QCF Titer & Serum Monitoring Dashboard (Professional Version)")

DB_EXCEL_PATH = "QCF Titer Data.xlsx"

if not os.path.exists(DB_EXCEL_PATH):
    st.error(f"❌ ไม่พบไฟล์ฐานข้อมูลหลัก '{DB_EXCEL_PATH}' ในระบบ กรุณาตรวจสอบว่ามีไฟล์นี้อยู่ใน GitHub Repository หรือยัง")
    st.stop()

# นิยามคอลัมน์บังคับสำหรับการตรวจสอบไฟล์ (Data Validation)
REQUIRED_ELISA_COLS = ['farm_name', 'House', 'Age (Wk)', 'elisa_test_kit', 'titer', 'Year']
REQUIRED_HI_COLS = ['farm_name', 'House', 'Age (Wk)', 'disease_name', 'GMT', 'Year', 'CV']

# ==========================================
# SIDEBAR: ระบบจัดการไฟล์และตรวจสอบข้อมูล
# ==========================================
st.sidebar.header("📁 ระบบจัดการข้อมูล (Data Management)")
upload_action = st.sidebar.radio(
    "เลือกรูปแบบการอัปโหลดข้อมูล:",
    ["1. เพิ่มข้อมูลต่อท้าย (Append)", "2. แทนที่ฐานข้อมูลใหม่ทั้งหมด (Overwrite)"]
)

if upload_action == "1. เพิ่มข้อมูลต่อท้าย (Append)":
    st.sidebar.markdown("---")
    data_type = st.sidebar.selectbox("เลือกประเภทข้อมูลที่ต้องการเพิ่ม", ["ELISA Data", "HI Data"])
    uploaded_file = st.sidebar.file_uploader(f"อัปโหลดไฟล์ Excel เพื่อดึงข้อมูลเข้าสู่ชีต '{data_type}'", type=["xlsx"])

    if uploaded_file is not None:
        try:
            new_data = pd.read_excel(uploaded_file, sheet_name=data_type)
            
            # ตรวจสอบโครงสร้างคอลัมน์ (Data Validation)
            required_cols = REQUIRED_ELISA_COLS if data_type == "ELISA Data" else REQUIRED_HI_COLS
            missing_cols = [col for col in required_cols if col not in new_data.columns]
            
            if missing_cols:
                st.sidebar.error(f"❌ โครงสร้างไฟล์ไม่ถูกต้อง! ขาดคอลัมน์: {', '.join(missing_cols)}")
            else:
                st.sidebar.success(f"📋 ตรวจสอบผ่าน! พบข้อมูลที่พร้อมอัปโหลดทั้งหมด {len(new_data)} แถ")

                if st.sidebar.button("💾 บันทึกเพิ่มต่อท้ายเข้าฐานข้อมูลหลัก"):
                    excel_file = pd.ExcelFile(DB_EXCEL_PATH)
                    all_sheets = {sheet: excel_file.parse(sheet) for sheet in excel_file.sheet_names}
                    
                    old_data = all_sheets.get(data_type, pd.DataFrame())
                    updated_data = pd.concat([old_data, new_data], ignore_index=True)
                    all_sheets[data_type] = updated_data
                    
                    with pd.ExcelWriter(DB_EXCEL_PATH, engine='openpyxl') as writer:
                        for sheet_name, df_sheet in all_sheets.items():
                            df_sheet.to_excel(writer, sheet_name=sheet_name, index=False)
                    
                    st.sidebar.balloons()
                    st.sidebar.success(f"✅ เพิ่มข้อมูลเข้าชีต '{data_type}' เรียบร้อยแล้ว!")
                    st.rerun()
        except Exception as e:
            st.sidebar.error(f"เกิดข้อผิดพลาด: {e}")
else:
    st.sidebar.markdown("---")
    st.sidebar.warning("⚠️ การทำ Overwrite จะแทนที่ไฟล์เดิมทั้งหมด")
    new_db_file = st.sidebar.file_uploader("อัปโหลดไฟล์ 'QCF Titer Data.xlsx' ตัวใหม่เข้าสู่ระบบ", type=["xlsx"])
    if new_db_file is not None:
        if st.sidebar.button("🚨 ยืนยันการบันทึกทับฐานข้อมูลเดิม"):
            try:
                with open(DB_EXCEL_PATH, "wb") as f:
                    f.write(new_db_file.getbuffer())
                st.sidebar.balloons()
                st.sidebar.success("✅ เปลี่ยนไฟล์ฐานข้อมูลหลักในระบบเรียบร้อยแล้ว!")
                st.rerun()
            except Exception as e:
                st.sidebar.error(f"เกิดข้อผิดพลาดในการบันทึกทับไฟล์: {e}")

# ==========================================
# LOAD DATA & PREPARATION
# ==========================================
try:
    excel_db = pd.ExcelFile(DB_EXCEL_PATH)
    elisa_raw = excel_db.parse("ELISA Data") if "ELISA Data" in excel_db.sheet_names else pd.DataFrame()
    hi_raw = excel_db.parse("HI Data") if "HI Data" in excel_db.sheet_names else pd.DataFrame()
except Exception as e:
    st.error(f"เกิดข้อผิดพลาดในการโหลดไฟล์ฐานข้อมูลหลัก: {e}")
    st.stop()

def clean_age(val):
    if pd.isna(val): return np.nan
    val_str = str(val).replace('สัปดาห์', '').replace('Wk', '').strip()
    try:
        return int(val_str)
    except:
        return val_str

# ฟังก์ชันคำนวณ Geometric Mean Titer (GMT) สำหรับ ELISA (Log10)
def elisa_gmt(series):
    valid = series.dropna()
    valid = valid[valid > 0]
    if len(valid) == 0: return 0
    return gmean(valid)

# ฟังก์ชันคำนวณ Geometric Mean Titer สำหรับ HI (Log2)
def hi_gmt_log2(series):
    valid = series.dropna()
    valid = valid[valid > 0]
    if len(valid) == 0: return 0
    # แปลงเป็นค่า Log2 หาค่าเฉลี่ยเลขคณิต แล้วค่อยแปลงกลับเป็น Antilog (ฐาน 2)
    log2_vals = np.log2(valid)
    return 2 ** log2_vals.mean()

tab1, tab2 = st.tabs(["🧪 ELISA Data Analysis", "🩸 HI Data Analysis"])

# ==========================================
# TAB 1: ELISA DATA ANALYSIS
# ==========================================
with tab1:
    st.header("สรุปผลการตรวจภูมิคุ้มกันด้วยวิธี ELISA")

    if elisa_raw.empty:
        st.info("ยังไม่มีข้อมูลสถิติในชีต 'ELISA Data'")
    else:
        # เตรียมข้อมูลเบื้องต้น
        elisa_raw['Age_Clean'] = elisa_raw['Age (Wk)'].apply(clean_age)
        elisa_raw['titer'] = pd.to_numeric(elisa_raw['titer'], errors='coerce')
        elisa_raw['Year'] = elisa_raw['Year'].astype(str).str.replace('.0', '', regex=False)

        # แยกข้อมูลฟาร์มจริง VS เกณฑ์มาตรฐาน (STD)
        is_std = elisa_raw['farm_name'].astype(str).str.contains('STD|Standard', case=False, na=False) | \
                 elisa_raw['House'].astype(str).str.contains('STD|Standard', case=False, na=False)
        
        elisa_actual = elisa_raw[~is_std]
        elisa_std_data = elisa_raw[is_std]

        # ตัวกรองข้อมูล (Filters)
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            years = st.multiselect("เลือกปี (Year - ELISA)", options=sorted(elisa_actual["Year"].dropna().unique()), default=sorted(elisa_actual["Year"].dropna().unique()))
        with col2:
            filtered_by_year = elisa_actual[elisa_actual["Year"].isin(years)]
            farms = st.multiselect("เลือกฟาร์ม (ELISA)", options=sorted(filtered_by_year["farm_name"].dropna().unique()), default=sorted(filtered_by_year["farm_name"].dropna().unique())[:1])
        with col3:
            kits = st.multiselect("เลือกชุดทดสอบ (Test Kit)", options=sorted(elisa_raw["elisa_test_kit"].dropna().unique()), default=sorted(elisa_raw["elisa_test_kit"].dropna().unique())[:1])
        with col4:
            filtered_by_farm = filtered_by_year[filtered_by_year["farm_name"].isin(farms)]
            available_houses = sorted(filtered_by_farm["House"].dropna().unique())
            houses = st.multiselect("เลือกโรงเรือน (House)", options=available_houses, default=available_houses)

        # กรองข้อมูลจริงตามเงื่อนไข
        f_elisa_actual = elisa_actual[
            (elisa_actual["Year"].isin(years)) &
            (elisa_actual["farm_name"].isin(farms)) & 
            (elisa_actual["elisa_test_kit"].isin(kits)) &
            (elisa_actual["House"].isin(houses))
        ].copy()

        if not f_elisa_actual.empty:
            # คำนวณ KPI โดยใช้ Geometric Mean Titer
            kpi1, kpi2, kpi3, kpi4 = st.columns(4)
            kpi1.metric("จำนวนตัวอย่างจริง", f"{len(f_elisa_actual)} ตัวอย่าง")
            
            gmt_val = elisa_gmt(f_elisa_actual['titer'])
            kpi2.metric("ค่าเฉลี่ยภูมิคุ้มกัน (GMT)", f"{gmt_val:.2f}" if gmt_val > 0 else "0.00")
            
            # คำนวณ %CV รวมกลุ่ม
            mean_arithmetic = f_elisa_actual['titer'].mean()
            cv_val = (f_elisa_actual['titer'].std() / mean_arithmetic * 100) if mean_arithmetic > 0 else 0
            kpi3.metric("ค่า %CV รวมกลุ่ม", f"{cv_val:.2f}%" if cv_val > 0 else "0.00%")

            if "result" in f_elisa_actual.columns:
                pos_count = len(f_elisa_actual[f_elisa_actual["result"].astype(str).str.lower() == "positive"])
                pos_rate = (pos_count / len(f_elisa_actual)) * 100
                kpi4.metric("เปอร์เซ็นต์ผลบวก (% Positive)", f"{pos_rate:.1f}%")

            # --- สร้างกราฟเส้นด้วย Plotly ---
            st.subheader("📈 ELISA Titer Profile (Geometric Mean) แยกตามฟาร์ม-โรงเรือน และเกณฑ์มาตรฐาน")
            
            f_elisa_actual['Legend_Name'] = f_elisa_actual['farm_name'].astype(str) + " (" + f_elisa_actual['House'].astype(str) + ")"
            chart_df = f_elisa_actual.groupby(['Age_Clean', 'Legend_Name'])['titer'].apply(elisa_gmt).unstack()
            chart_df['[Mean] Overall GMT'] = f_elisa_actual.groupby('Age_Clean')['titer'].apply(elisa_gmt)

            f_elisa_std = elisa_std_data[elisa_std_data["elisa_test_kit"].isin(kits)]
            if not f_elisa_std.empty:
                std_lines = f_elisa_std.groupby(['Age_Clean', 'House'])['titer'].mean().unstack()
                for col in std_lines.columns:
                    chart_df[f"[Standard] {col}"] = std_lines[col]

            chart_df = chart_df.sort_index()

            # สร้างองค์ประกอบกราฟ Plotly
            fig = go.Figure()
            for col in chart_df.columns:
                if "[Standard]" in col:
                    fig.add_trace(go.Scatter(x=chart_df.index, y=chart_df[col], name=col, line=dict(dash='dash', color='orange', width=1.5)))
                elif "[Mean]" in col:
                    fig.add_trace(go.Scatter(x=chart_df.index, y=chart_df[col], name=col, line=dict(color='black', width=4)))
                else:
                    fig.add_trace(go.Scatter(x=chart_df.index, y=chart_df[col], name=col, mode='lines+markers', line=dict(width=2)))
            
            fig.update_layout(xaxis_title="อายุ (สัปดาห์ - Age in Weeks)", yaxis_title="Geometric Mean Titer", hovermode="x unified", legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1))
            st.plotly_chart(fig, use_container_width=True)

            # ตารางสรุปรายสัปดาห์
            st.subheader("📋 ตารางวิเคราะห์ข้อมูลแยกตามรายอายุสัปดาห์ (%CV รายสัปดาห์)")
            summary_table = f_elisa_actual.groupby(['Year', 'farm_name', 'House', 'Age (Wk)']).agg(
                จำนวนตัวอย่าง=('titer', 'count'),
                ค่าเฉลี่ย_GMT=('titer', elisa_gmt),
                ความสม่ำเสมอ_CV_เปอร์เซ็นต์=('titer', lambda x: (x.std() / x.mean() * 100) if x.mean() > 0 else 0)
            ).reset_index()
            st.dataframe(summary_table.style.format({'ค่าเฉลี่ย_GMT': '{:.2f}', 'ความสม่ำเสมอ_CV_เปอร์เซ็นต์': '{:.2f}%'}), use_container_width=True)
        else:
            st.warning("ไม่พบข้อมูลตามเงื่อนไขที่เลือก")

# ==========================================
# TAB 2: HI DATA ANALYSIS
# ==========================================
with tab2:
    st.header("สรุปผลการตรวจด้วยวิธี Haemagglutination Inhibition (HI)")

    if hi_raw.empty:
        st.info("ยังไม่มีข้อมูลสถิติในชีต 'HI Data'")
    else:
        # เตรียมข้อมูลเบื้องต้น
        hi_raw['Age_Clean'] = hi_raw['Age (Wk)'].apply(clean_age)
        hi_raw['GMT'] = pd.to_numeric(hi_raw['GMT'], errors='coerce')
        hi_raw['CV'] = pd.to_numeric(hi_raw['CV'], errors='coerce')
        hi_raw['Year'] = hi_raw['Year'].astype(str).str.replace('.0', '', regex=False)

        # แยกข้อมูลฟาร์มจริง VS ค่ามาตรฐาน
        is_std_hi = hi_raw['farm_name'].astype(str).str.contains('STD|Standard', case=False, na=False) | \
                    hi_raw['House'].astype(str).str.contains('STD|Standard', case=False, na=False)
        
        hi_actual = hi_raw[~is_std_hi]
        hi_std_data = hi_raw[is_std_hi]

        # ตัวกรองข้อมูล (Filters)
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            years_hi = st.multiselect("เลือกปี (Year - HI)", options=sorted(hi_actual["Year"].dropna().unique()), default=sorted(hi_actual["Year"].dropna().unique()))
        with col2:
            filtered_by_year_hi = hi_actual[hi_actual["Year"].isin(years_hi)]
            farms_hi = st.multiselect("เลือกฟาร์ม (HI)", options=sorted(filtered_by_year_hi["farm_name"].dropna().unique()), default=sorted(filtered_by_year_hi["farm_name"].dropna().unique())[:1])
        with col3:
            diseases = st.multiselect("เลือกโรค (Disease)", options=sorted(hi_raw["disease_name"].dropna().unique()), default=sorted(hi_raw["disease_name"].dropna().unique())[:1])
        with col4:
            filtered_by_farm_hi = filtered_by_year_hi[filtered_by_year_hi["farm_name"].isin(farms_hi)]
            available_houses_hi = sorted(filtered_by_farm_hi["House"].dropna().unique())
            houses_hi = st.multiselect("เลือกโรงเรือน (House - HI)", options=available_houses_hi, default=available_houses_hi)

        # กรองข้อมูลจริงตามเงื่อนไข
        f_hi_actual = hi_actual[
            (hi_actual["Year"].isin(years_hi)) &
            (hi_actual["farm_name"].isin(farms_hi)) & 
            (hi_actual["disease_name"].isin(diseases)) &
            (hi_actual["House"].isin(houses_hi))
        ].copy()

        if not f_hi_actual.empty:
            # ส่วนคำนวณ KPI ของ HI โดยใช้สูตร Log2 GMT
            hkpi1, hkpi2, hkpi3 = st.columns(3)
            hkpi1.metric("จำนวนบันทึกจริง", f"{len(f_hi_actual)} รายการ")

            avg_gmt = hi_gmt_log2(f_hi_actual["GMT"])
            hkpi2.metric("ค่าเฉลี่ย GMT รวมกลุ่ม (Log2 Based)", f"{avg_gmt:.2f}" if avg_gmt > 0 else "0.00")

            avg_cv = f_hi_actual["CV"].mean()
            hkpi3.metric("ค่าเฉลี่ย %CV ของฝูง", f"{avg_cv:.2f}%" if avg_cv > 0 else "0.00%")

            # --- สร้างกราฟเส้น HI ด้วย Plotly ---
            st.subheader("📈 HI Titer Profile (Log2 GMT) แยกตามฟาร์ม-โรงเรือน และเกณฑ์มาตรฐาน")
            
            f_hi_actual['Legend_Name'] = f_hi_actual['farm_name'].astype(str) + " (" + f_hi_actual['House'].astype(str) + ")"
            chart_hi_df = f_hi_actual.groupby(['Age_Clean', 'Legend_Name'])['GMT'].apply(hi_gmt_log2).unstack()
            chart_hi_df['[Mean] Overall GMT'] = f_hi_actual.groupby('Age_Clean')['GMT'].apply(hi_gmt_log2)

            f_hi_std = hi_std_data[hi_std_data["disease_name"].isin(diseases)]
            if not f_hi_std.empty:
                std_lines_hi = f_hi_std.groupby(['Age_Clean', 'House'])['GMT'].mean().unstack()
                for col in std_lines_hi.columns:
                    chart_hi_df[f"[Standard] {col}"] = std_lines_hi[col]

            chart_hi_df = chart_hi_df.sort_index()

            fig_hi = go.Figure()
            for col in chart_hi_df.columns:
                if "[Standard]" in col:
                    fig_hi.add_trace(go.Scatter(x=chart_hi_df.index, y=chart_hi_df[col], name=col, line=dict(dash='dash', color='orange', width=1.5)))
                elif "[Mean]" in col:
                    fig_hi.add_trace(go.Scatter(x=chart_hi_df.index, y=chart_hi_df[col], name=col, line=dict(color='black', width=4)))
                else:
                    fig_hi.add_trace(go.Scatter(x=chart_hi_df.index, y=chart_hi_df[col], name=col, mode='lines+markers', line=dict(width=2)))
            
            fig_hi.update_layout(xaxis_title="อายุ (สัปดาห์ - Age in Weeks)", yaxis_title="Haemagglutination Inhibition GMT (Log2)", hovermode="x unified", legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1))
            st.plotly_chart(fig_hi, use_container_width=True)

            st.subheader("📋 ข้อมูลการทดสอบ HI จำแนกตามรายรุ่นและช่วงอายุ")
            st.dataframe(f_hi_actual.drop(columns=['Age_Clean', 'Legend_Name'], errors='ignore'), use_container_width=True)
        else:
            st.warning("ไม่พบข้อมูลตามเงื่อนไขที่เลือก")
