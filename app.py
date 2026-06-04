import streamlit as st
import pandas as pd
import numpy as np
import os

# ตั้งค่าหน้าจอ Web-App
st.set_page_config(page_title="QCF Titer Dashboard 2026", layout="wide")
st.title("📊 QCF Titer & Serum Monitoring Dashboard")

# กำหนดชื่อไฟล์ฐานข้อมูล Excel หลักตามที่คุณระบุ
DB_EXCEL_PATH = "QCF Titer Data.xlsx"

# ตรวจสอบโครงสร้างไฟล์ Excel หลักในระบบ (ถ้าไม่มี ให้แจ้งเตือนผู้ใช้)
if not os.path.exists(DB_EXCEL_PATH):
    st.error(f"❌ ไม่พบไฟล์ฐานข้อมูลหลัก '{DB_EXCEL_PATH}' ในระบบ กรุณาตรวจสอบว่ามีไฟล์นี้อยู่ใน GitHub Repository หรือยัง")
    st.stop()

# ==========================================
# SIDEBAR: ระบบจัดการไฟล์และอัปโหลดข้อมูล
# ==========================================
st.sidebar.header("📁 ระบบจัดการข้อมูล (Data Management)")

# เลือกรูปแบบการอัปโหลดข้อมูล
upload_action = st.sidebar.radio(
    "เลือกรูปแบบการอัปโหลดข้อมูล:",
    ["1. เพิ่มข้อมูลต่อท้าย (Append)", "2. แทนที่ฐานข้อมูลใหม่ทั้งหมด (Overwrite)"]
)

if upload_action == "1. เพิ่มข้อมูลต่อท้าย (Append)":
    st.sidebar.markdown("---")
    data_type = st.sidebar.selectbox("เลือกประเภทข้อมูลที่ต้องการเพิ่ม", ["ELISA Data", "HI Data"])
    uploaded_file = st.sidebar.file_uploader(
        f"อัปโหลดไฟล์ Excel เพื่อดึงข้อมูลเข้าสู่ชีต '{data_type}'", type=["xlsx"]
    )

    if uploaded_file is not None:
        try:
            # ดึงข้อมูลจากไฟล์ที่ผู้ใช้อัปโหลด เฉพาะชีตที่ระบุ
            new_data = pd.read_excel(uploaded_file, sheet_name=data_type)
            st.sidebar.success(f"📋 โหลดข้อมูลสำเร็จ! พบข้อมูลทั้งหมด {len(new_data)} แถว")

            if st.sidebar.button("💾 บันทึกเพิ่มต่อท้ายเข้าฐานข้อมูลหลัก"):
                # อ่านไฟล์ Excel หลักทั้งหมดขึ้นมาก่อนเพื่อรักษารูปแบบชีตอื่นๆ ไว้
                excel_file = pd.ExcelFile(DB_EXCEL_PATH)
                all_sheets = {sheet: excel_file.parse(sheet) for sheet in excel_file.sheet_names}
                
                # นำข้อมูลใหม่ไป Append รวมเข้ากับข้อมูลเดิมในชีตนั้น
                old_data = all_sheets.get(data_type, pd.DataFrame())
                updated_data = pd.concat([old_data, new_data], ignore_index=True)
                all_sheets[data_type] = updated_data
                
                # เขียนทับลงไฟล์ Excel เดิม โดยเก็บชีตอื่นๆ ไว้ครบถ้วน
                with pd.ExcelWriter(DB_EXCEL_PATH, engine='openpyxl') as writer:
                    for sheet_name, df_sheet in all_sheets.items():
                        df_sheet.to_excel(writer, sheet_name=sheet_name, index=False)
                
                st.sidebar.balloons()
                st.sidebar.success(f"✅ เพิ่มข้อมูลเข้าชีต '{data_type}' เรียบร้อยแล้ว!")
                st.rerun()
                
        except ValueError:
            st.sidebar.error(f"❌ ไม่พบชื่อชีต '{data_type}' ในไฟล์ที่อัปโหลด กรุณาตรวจสอบชื่อชีตในไฟล์ Excel ของคุณ")
        except Exception as e:
            st.sidebar.error(f"เกิดข้อผิดพลาด: {e}")

else:  # กรณีต้องการ Overwrite แทนที่ไฟล์เดิมทั้งหมด
    st.sidebar.markdown("---")
    st.sidebar.warning("⚠️ การทำ Overwrite จะแทนที่ไฟล์ 'QCF Titer Data.xlsx' เดิมในระบบทั้งหมดด้วยไฟล์ใหม่ที่คุณอัปโหลด")
    new_db_file = st.sidebar.file_uploader(
        "อัปโหลดไฟล์ 'QCF Titer Data.xlsx' ตัวใหม่เข้าสู่ระบบ", type=["xlsx"]
    )
    
    if new_db_file is not None:
        if st.sidebar.button("🚨 ยืนยันการบันทึกทับฐานข้อมูลเดิม"):
            try:
                # บันทึกไฟล์ใหม่ทับไฟล์เดิมโดยตรง
                with open(DB_EXCEL_PATH, "wb") as f:
                    f.write(new_db_file.getbuffer())
                st.sidebar.balloons()
                st.sidebar.success("✅ เปลี่ยนไฟล์ฐานข้อมูลหลักในระบบเรียบร้อยแล้ว!")
                st.rerun()
            except Exception as e:
                st.sidebar.error(f"เกิดข้อผิดพลาดในการบันทึกทับไฟล์: {e}")


# ==========================================
# MAIN PAGE: โหลดฐานข้อมูลมาแสดงผลบน Dashboard
# ==========================================
try:
    excel_db = pd.ExcelFile(DB_EXCEL_PATH)
    elisa_df = excel_db.parse("ELISA Data") if "ELISA Data" in excel_db.sheet_names else pd.DataFrame()
    hi_df = excel_db.parse("HI Data") if "HI Data" in excel_db.sheet_names else pd.DataFrame()
except Exception as e:
    st.error(f"เกิดข้อผิดพลาดในการโหลดไฟล์ฐานข้อมูลหลัก: {e}")
    st.stop()

tab1, tab2 = st.tabs(["🧪 ELISA Data Analysis", "🩸 HI Data Analysis"])

# --- TAB 1: ELISA DATA ---
with tab1:
    st.header("สรุปผลการตรวจภูมิคุ้มกันด้วยวิธี ELISA")

    if elisa_df.empty or len(elisa_df) == 0:
        st.info("ยังไม่มีข้อมูลสถิติในชีต 'ELISA Data' กรุณาอัปโหลดเพิ่มข้อมูลที่แถบด้านซ้าย")
    else:
        # ตัวกรองข้อมูล (Filters)
        col1, col2, col3 = st.columns(3)
        with col1:
            farms = st.multiselect(
                "เลือกฟาร์ม (ELISA)",
                options=elisa_df["farm_name"].dropna().unique(),
                default=elisa_df["farm_name"].dropna().unique()[:1] if len(elisa_df["farm_name"].dropna().unique()) > 0 else []
            )
        with col2:
            kits = st.multiselect(
                "เลือกชุดทดสอบ (Test Kit)",
                options=elisa_df["elisa_test_kit"].dropna().unique(),
                default=elisa_df["elisa_test_kit"].dropna().unique() if len(elisa_df["elisa_test_kit"].dropna().unique()) > 0 else []
            )
        with col3:
            houses = st.multiselect(
                "เลือกโรงเรือน (House)",
                options=elisa_df["House"].dropna().unique() if "House" in elisa_df.columns else [],
                default=elisa_df["House"].dropna().unique() if "House" in elisa_df.columns else []
            )

        # กรองข้อมูลตามเงื่อนไขที่เลือก
        filtered_elisa = elisa_df[
            (elisa_df["farm_name"].isin(farms)) & 
            (elisa_df["elisa_test_kit"].isin(kits))
        ]
        if "House" in elisa_df.columns and len(houses) > 0:
            filtered_elisa = filtered_elisa[filtered_elisa["House"].isin(houses)]

        if not filtered_elisa.empty:
            # KPI Cards
            kpi1, kpi2, kpi3, kpi4 = st.columns(4)
            kpi1.metric("จำนวนตัวอย่างทั้งหมด", f"{len(filtered_elisa)} ตัวอย่าง")
            
            titer_col = pd.to_numeric(filtered_elisa['titer'], errors='coerce')
            mean_val = titer_col.mean() if titer_col.mean() > 0 else 0
            kpi2.metric("ค่าเฉลี่ย Titer (Mean)", f"{mean_val:.2f}")
            
            cv_val = (titer_col.std() / mean_val * 100) if mean_val > 0 else 0
            kpi3.metric("ค่า %CV", f"{cv_val:.2f}%")

            if "result" in filtered_elisa.columns:
                pos_count = len(filtered_elisa[filtered_elisa["result"].astype(str).str.lower() == "positive"])
                pos_rate = (pos_count / len(filtered_elisa)) * 100
                kpi4.metric("เปอร์เซ็นต์ผลบวก (% Positive)", f"{pos_rate:.1f}%")

            # กราฟแนวโน้มระดับภูมิคุ้มกัน (Titer Profile)
            st.subheader("📈 แนวโน้มระดับภูมิคุ้มกันเฉลี่ยตามช่วงอายุสัปดาห์")
            age_col = "Age (Wk)" if "Age (Wk)" in filtered_elisa.columns else "age"
            
            if age_col in filtered_elisa.columns:
                filtered_elisa['titer_numeric'] = pd.to_numeric(filtered_elisa['titer'], errors='coerce')
                chart_data = filtered_elisa.groupby([age_col, "elisa_test_kit"])["titer_numeric"].mean().unstack().fillna(0)
                st.line_chart(chart_data)

            # ตารางข้อมูล
            st.subheader("📋 ตารางสรุปข้อมูลจำแนกตามฝูงและอายุ")
            st.dataframe(filtered_elisa.drop(columns=['titer_numeric'], errors='ignore'), use_container_width=True)
        else:
            st.warning("ไม่พบข้อมูลตามเงื่อนไขที่เลือก")

# --- TAB 2: HI DATA ---
with tab2:
    st.header("สรุปผลการตรวจด้วยวิธี Haemagglutination Inhibition (HI)")

    if hi_df.empty or len(hi_df) == 0:
        st.info("ยังไม่มีข้อมูลสถิติในชีต 'HI Data' กรุณาอัปโหลดเพิ่มข้อมูลที่แถบด้านซ้าย")
    else:
        # ตัวกรองข้อมูล (Filters)
        col1, col2, col3 = st.columns(3)
        with col1:
            farms_hi = st.multiselect(
                "เลือกฟาร์ม (HI)",
                options=hi_df["farm_name"].dropna().unique(),
                default=hi_df["farm_name"].dropna().unique()[:1] if len(hi_df["farm_name"].dropna().unique()) > 0 else []
            )
        with col2:
            diseases = st.multiselect(
                "เลือกโรค (Disease)",
                options=hi_df["disease_name"].dropna().unique(),
                default=hi_df["disease_name"].dropna().unique() if len(hi_df["disease_name"].dropna().unique()) > 0 else []
            )
        with col3:
            houses_hi = st.multiselect(
                "เลือกโรงเรือน (House - HI)",
                options=hi_df["House"].dropna().unique() if "House" in hi_df.columns else [],
                default=hi_df["House"].dropna().unique() if "House" in hi_df.columns else []
            )

        # กรองข้อมูลตามเงื่อนไข
        filtered_hi = hi_df[
            (hi_df["farm_name"].isin(farms_hi)) & 
            (hi_df["disease_name"].isin(diseases))
        ]
        if "House" in hi_df.columns and len(houses_hi) > 0:
            filtered_hi = filtered_hi[filtered_hi["House"].isin(houses_hi)]

        if not filtered_hi.empty:
            # KPI Cards
            hkpi1, hkpi2, hkpi3 = st.columns(3)
            hkpi1.metric("จำนวนบันทึกฝูง (Records)", f"{len(filtered_hi)} รายการ")

            gmt_val = pd.to_numeric(filtered_hi["GMT"], errors="coerce").mean()
            hkpi2.metric("ค่าเฉลี่ย GMT", f"{gmt_val:.2f}" if gmt_val > 0 else "0.00")

            cv_hi_val = pd.to_numeric(filtered_hi["CV"], errors="coerce").mean()
            hkpi3.metric("ค่าเฉลี่ย %CV", f"{cv_hi_val:.2f}%" if cv_hi_val > 0 else "0.00%")

            # กราฟแท่ง HI GMT
            st.subheader("📈 ระดับไตเตอร์ HI (GMT) แยกตามอายุสัปดาห์")
            age_col_hi = "Age (Wk)" if "Age (Wk)" in filtered_hi.columns else "age"
            if age_col_hi in filtered_hi.columns:
                filtered_hi['gmt_numeric'] = pd.to_numeric(filtered_hi["GMT"], errors="coerce")
                hi_chart_data = filtered_hi.groupby([age_col_hi, "disease_name"])["gmt_numeric"].mean().unstack().fillna(0)
                st.bar_chart(hi_chart_data)

            # ตารางแสดงข้อมูล
            st.subheader("📋 ข้อมูลการทดสอบ HI ทั้งหมดที่ตรงตามเงื่อนไข")
            st.dataframe(filtered_hi.drop(columns=['gmt_numeric'], errors='ignore'), use_container_width=True)
        else:
            st.warning("ไม่พบข้อมูลตามเงื่อนไขที่เลือก")
