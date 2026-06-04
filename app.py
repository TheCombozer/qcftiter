import streamlit as st
import pandas as pd
import numpy as np
import os

# ตั้งค่าหน้าจอ Web-App
st.set_page_config(page_title="QCF Titer Dashboard 2026", layout="wide")
st.title("📊 QCF Titer & Serum Monitoring Dashboard")

DB_EXCEL_PATH = "QCF Titer Data.xlsx"

if not os.path.exists(DB_EXCEL_PATH):
    st.error(f"❌ ไม่พบไฟล์ฐานข้อมูลหลัก '{DB_EXCEL_PATH}' ในระบบ กรุณาตรวจสอบว่ามีไฟล์นี้อยู่ใน GitHub Repository หรือยัง")
    st.stop()

# ==========================================
# SIDEBAR: ระบบจัดการไฟล์และอัปโหลดข้อมูล
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
            st.sidebar.success(f"📋 โหลดข้อมูลสำเร็จ! พบข้อมูลทั้งหมด {len(new_data)} แถว")

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

# ฟังก์ชันจัดรูปแบบคำว่า อายุ (สัปดาห์) ให้เป็นตัวเลขอายุเพื่อให้เรียงลำดับบนแกนกราฟได้ถูกต้อง
def clean_age(val):
    if pd.isna(val): return np.nan
    val_str = str(val).replace('สัปดาห์', '').replace('Wk', '').strip()
    try:
        return int(val_str)
    except:
        return val_str

tab1, tab2 = st.tabs(["🧪 ELISA Data Analysis", "🩸 HI Data Analysis"])

# ==========================================
# TAB 1: ELISA DATA ANALYSIS
# ==========================================
with tab1:
    st.header("สรุปผลการตรวจภูมิคุ้มกันด้วยวิธี ELISA")

    if elisa_raw.empty:
        st.info("ยังไม่มีข้อมูลสถิติในชีต 'ELISA Data'")
    else:
        # เตรียมคอลัมน์อายุที่เป็นตัวเลข
        age_col = "Age (Wk)" if "Age (Wk)" in elisa_raw.columns else "age"
        elisa_raw['Age_Clean'] = elisa_raw[age_col].apply(clean_age)
        elisa_raw['titer'] = pd.to_numeric(elisa_raw['titer'], errors='coerce')

        # แยกข้อมูล: ข้อมูลฟาร์มจริง VS ข้อมูลค่ามาตรฐาน (STD)
        is_std = elisa_raw['farm_name'].astype(str).str.contains('STD|Standard', case=False, na=False) | \
                 elisa_raw['House'].astype(str).str.contains('STD|Standard', case=False, na=False)
        
        elisa_actual = elisa_raw[~is_std]  # ข้อมูลฟาร์มจริงเท่านั้นที่จะนำไปเฉลี่ยและคำนวณ KPI
        elisa_std_data = elisa_raw[is_std] # ค่าเกณฑ์มาตรฐานสำหรับทำกราฟ

        # ตัวกรองข้อมูล (Filters)
        col1, col2, col3 = st.columns(3)
        with col1:
            farms = st.multiselect(
                "เลือกฟาร์ม (ELISA)",
                options=sorted(elisa_actual["farm_name"].dropna().unique()),
                default=sorted(elisa_actual["farm_name"].dropna().unique())[:1] if len(elisa_actual) > 0 else []
            )
        with col2:
            kits = st.multiselect(
                "เลือกชุดทดสอบ (Test Kit)",
                options=sorted(elisa_raw["elisa_test_kit"].dropna().unique()),
                default=sorted(elisa_raw["elisa_test_kit"].dropna().unique())[:1] if len(elisa_raw) > 0 else []
            )
        with col3:
            # ดึงเฉพาะโรงเรือนของฟาร์มที่เลือกมาแสดงในตัวเลือก
            available_houses = sorted(elisa_actual[elisa_actual["farm_name"].isin(farms)]["House"].dropna().unique()) if "House" in elisa_actual.columns else []
            houses = st.multiselect("เลือกโรงเรือน (House)", options=available_houses, default=available_houses)

        # กรองข้อมูลจริงตามเงื่อนไขตัวเลือก
        f_elisa_actual = elisa_actual[
            (elisa_actual["farm_name"].isin(farms)) & 
            (elisa_actual["elisa_test_kit"].isin(kits)) &
            (elisa_actual["House"].isin(houses))
        ]

        if not f_elisa_actual.empty:
            # --- ส่วนคำนวณ KPI (ปลอดภัยจากการบิดเบือนของค่า STD) ---
            kpi1, kpi2, kpi3, kpi4 = st.columns(4)
            kpi1.metric("จำนวนตัวอย่างจริง", f"{len(f_elisa_actual)} ตัวอย่าง")
            
            mean_val = f_elisa_actual['titer'].mean()
            kpi2.metric("ค่าเฉลี่ย Titer (Mean)", f"{mean_val:.2f}" if mean_val > 0 else "0.00")
            
            cv_val = (f_elisa_actual['titer'].std() / mean_val * 100) if mean_val > 0 else 0
            kpi3.metric("ค่าความสม่ำเสมอ (%CV)", f"{cv_val:.2f}%" if cv_val > 0 else "0.00%")

            if "result" in f_elisa_actual.columns:
                pos_count = len(f_elisa_actual[f_elisa_actual["result"].astype(str).str.lower() == "positive"])
                pos_rate = (pos_count / len(f_elisa_actual)) * 100
                kpi4.metric("เปอร์เซ็นต์ผลบวก (% Positive)", f"{pos_rate:.1f}%")

            # --- ส่วนของการสร้างกราฟเส้นแยกโรงเรือน + เส้นเฉลี่ยรวม + เส้น STD ---
            st.subheader("📈 ELISA Titer Profile แยกตามโรงเรือนและเกณฑ์มาตรฐาน")
            
            # พล็อตแยกเส้นตามโรงเรือนจริง
            chart_df = f_elisa_actual.groupby(['Age_Clean', 'House'])['titer'].mean().unstack()
            
            # คำนวณเส้นเฉลี่ยรวมของทุกโรงเรือนที่ถูกเลือก (Overall Mean)
            chart_df['Overall Mean'] = f_elisa_actual.groupby('Age_Clean')['titer'].mean()

            # ดึงค่าเกณฑ์มาตรฐาน (STD) ที่แมตช์กับชุดตรวจที่เลือกมาพล็อตลงกราฟ
            f_elisa_std = elisa_std_data[elisa_std_data["elisa_test_kit"].isin(kits)]
            if not f_elisa_std.empty:
                std_lines = f_elisa_std.groupby(['Age_Clean', 'House'])['titer'].mean().unstack()
                for col in std_lines.columns:
                    chart_df[col] = std_lines[col] # เพิ่มเส้น STD เข้าไปใน DataFrame สำหรับพล็อต

            # เรียงอายุจากน้อยไปมากเพื่อให้เส้นกราฟลากอย่างถูกต้อง
            chart_df = chart_df.sort_index()
            
            # แสดงกราฟเส้น
            st.line_chart(chart_df)

            st.subheader("📋 ตารางข้อมูลสถิติเฉพาะฟาร์มจริง (ไม่รวม STD)")
            st.dataframe(f_elisa_actual.drop(columns=['Age_Clean'], errors='ignore'), use_container_width=True)
        else:
            st.warning("ไม่พบข้อมูลตามเงื่อนไขที่เลือก")

# ==========================================
# TAB 2: HI DATA ANALYSIS (ปรับปรุงเป็นกราฟเส้น)
# ==========================================
with tab2:
    st.header("สรุปผลการตรวจด้วยวิธี Haemagglutination Inhibition (HI)")

    if hi_raw.empty:
        st.info("ยังไม่มีข้อมูลสถิติในชีต 'HI Data'")
    else:
        # เตรียมคอลัมน์อายุและตัวเลขค่า GMT
        age_col_hi = "Age (Wk)" if "Age (Wk)" in hi_raw.columns else "age"
        hi_raw['Age_Clean'] = hi_raw[age_col_hi].apply(clean_age)
        hi_raw['GMT'] = pd.to_numeric(hi_raw['GMT'], errors='coerce')
        hi_raw['CV'] = pd.to_numeric(hi_raw['CV'], errors='coerce')

        # แยกข้อมูล: ฟาร์มจริง VS ค่ามาตรฐาน (STD min / STD max)
        is_std_hi = hi_raw['farm_name'].astype(str).str.contains('STD|Standard', case=False, na=False) | \
                    hi_raw['House'].astype(str).str.contains('STD|Standard', case=False, na=False)
        
        hi_actual = hi_raw[~is_std_hi]
        hi_std_data = hi_raw[is_std_hi]

        # ตัวกรองข้อมูล (Filters)
        col1, col2, col3 = st.columns(3)
        with col1:
            farms_hi = st.multiselect(
                "เลือกฟาร์ม (HI)",
                options=sorted(hi_actual["farm_name"].dropna().unique()),
                default=sorted(hi_actual["farm_name"].dropna().unique())[:1] if len(hi_actual) > 0 else []
            )
        with col2:
            diseases = st.multiselect(
                "เลือกโรค (Disease)",
                options=sorted(hi_raw["disease_name"].dropna().unique()),
                default=sorted(hi_raw["disease_name"].dropna().unique())[:1] if len(hi_raw) > 0 else []
            )
        with col3:
            available_houses_hi = sorted(hi_actual[hi_actual["farm_name"].isin(farms_hi)]["House"].dropna().unique()) if "House" in hi_actual.columns else []
            houses_hi = st.multiselect("เลือกโรงเรือน (House - HI)", options=available_houses_hi, default=available_houses_hi)

        # กรองข้อมูลจริงตามเงื่อนไข
        f_hi_actual = hi_actual[
            (hi_actual["farm_name"].isin(farms_hi)) & 
            (hi_actual["disease_name"].isin(diseases)) &
            (hi_actual["House"].isin(houses_hi))
        ]

        if not f_hi_actual.empty:
            # --- ส่วนคำนวณ KPI ของ HI ---
            hkpi1, hkpi2, hkpi3 = st.columns(3)
            hkpi1.metric("จำนวนบันทึกจริง", f"{len(f_hi_actual)} รายการ")

            avg_gmt = f_hi_actual["GMT"].mean()
            hkpi2.metric("ค่าเฉลี่ย GMT รวม", f"{avg_gmt:.2f}" if avg_gmt > 0 else "0.00")

            avg_cv = f_hi_actual["CV"].mean()
            hkpi3.metric("ค่าเฉลี่ย %CV", f"{avg_cv:.2f}%" if avg_cv > 0 else "0.00%")

            # --- ส่วนของการสร้างกราฟเส้นระบบ HI Titer Profile ---
            st.subheader("📈 HI Titer Profile (GMT) แยกตามโรงเรือนและเกณฑ์มาตรฐาน")
            
            # พล็อตแยกเส้นตามโรงเรือนจริง
            chart_hi_df = f_hi_actual.groupby(['Age_Clean', 'House'])['GMT'].mean().unstack()
            
            # คำนวณเส้นค่าเฉลี่ยของกลุ่มโรงเรือนที่ถูกเลือก (Overall Mean GMT)
            chart_hi_df['Overall Mean'] = f_hi_actual.groupby('Age_Clean')['GMT'].mean()

            # ดึงเกณฑ์มาตรฐานของโรคที่เลือกมาพล็อตลงกราฟเส้นประ
            f_hi_std = hi_std_data[hi_std_data["disease_name"].isin(diseases)]
            if not f_hi_std.empty:
                std_lines_hi = f_hi_std.groupby(['Age_Clean', 'House'])['GMT'].mean().unstack()
                for col in std_lines_hi.columns:
                    chart_hi_df[col] = std_lines_hi[col]

            # เรียงลำดับอายุตามแกน X
            chart_hi_df = chart_hi_df.sort_index()
            
            # พล็อตกราฟเส้นระบบ HI
            st.line_chart(chart_hi_df)

            st.subheader("📋 ข้อมูลการทดสอบ HI เฉพาะฟาร์มจริง (ไม่รวม STD)")
            st.dataframe(f_hi_actual.drop(columns=['Age_Clean'], errors='ignore'), use_container_width=True)
        else:
            st.warning("ไม่พบข้อมูลตามเงื่อนไขที่เลือก")
