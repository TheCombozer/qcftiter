import streamlit as st
import pandas as pd
import numpy as np
import os

# ตั้งค่าหน้าจอ Web-App
st.set_page_config(page_title="QCF Titer Dashboard 2026", layout="wide")
st.title("📊 QCF Titer & Serum Monitoring Dashboard")

# กำหนดชื่อไฟล์ฐานข้อมูลหลัก (จำลองระบบฐานข้อมูลด้วย CSV)
ELISA_DB_PATH = "ELISA_Data.csv"
HI_DB_PATH = "HI_Data.csv"


# ฟังก์ชันสร้างไฟล์จำลองเริ่มต้น (หากยังไม่มีไฟล์)
def init_databases():
    if not os.path.exists(ELISA_DB_PATH):
        columns = [
            "farm_name",
            "method_name",
            "ani_group_name",
            "ani_no",
            "age",
            "elisa_test_kit",
            "ani_name",
            "sampling_date",
            "sampling_type_name",
            "O_D",
            "well",
            "result",
            "titer_group",
            "titer",
            "GMT",
            "CV",
            "Min",
            "Max",
            "Year",
            "Age (Wk)",
            "House",
        ]
        pd.DataFrame(columns=columns).to_csv(ELISA_DB_PATH, index=False)

    if not os.path.exists(HI_DB_PATH):
        columns = [
            "farm_name",
            "method_name",
            "disease_name",
            "ani_group_name",
            "ani_no",
            "age",
            "sampling_date",
            "sampling_type_name",
            "sample_count",
            "titer_group",
            "GMT",
            "CV",
            "STD",
            "Year",
            "Age (Wk)",
            "House",
        ]
        pd.DataFrame(columns=columns).to_csv(HI_DB_PATH, index=False)


init_databases()

# ==========================================
# SIDEBAR: ระบบอัพโหลดและเพิ่มข้อมูล (Data Ingestion)
# ==========================================
st.sidebar.header("📁 ระบบจัดการข้อมูล (Data Upload)")
data_type = st.sidebar.selectbox("เลือกประเภทข้อมูลที่ต้องการอัพโหลด", ["ELISA Data", "HI Data"])
uploaded_file = st.sidebar.file_uploader(
    "อัพโหลดไฟล์ CSV หรือ Excel เพื่อเพิ่มข้อมูล", type=["csv", "xlsx"]
)

if uploaded_file is not None:
    try:
        # อ่านไฟล์ที่อัพโหลด
        if uploaded_file.name.endswith(".csv"):
            new_data = pd.read_csv(uploaded_file)
        else:
            new_data = pd.read_excel(uploaded_file)

        st.sidebar.success(f"โหลดไฟล์ {uploaded_file.name} สำเร็จ! พบข้อมูล {len(new_data)} แถว")

        # ปุ่มกดยืนยันการบันทึกข้อมูลต่อท้าย (Append)
        if st.sidebar.button("💾 บันทึกเพิ่มเข้าสู่ระบบฐานข้อมูลหลัก"):
            db_path = ELISA_DB_PATH if data_type == "ELISA Data" else \
                HI_DB_PATH
            main_df = pd.read_csv(db_path)

            # ตรวจสอบความสอดคล้องของคอลัมน์
            missing_cols = [col for col in main_df.columns if col not in new_data.columns]

            if len(missing_cols) == 0 or st.sidebar.checkbox("ยอมรับการอัพโหลดแม้คอลัมน์ไม่ครบ"):
                # ดำเนินการต่อท้ายข้อมูล
                updated_df = pd.concat([main_df, new_data], ignore_index=True)
                updated_df.to_csv(db_path, index=False)
                st.sidebar.balloons()
                st.sidebar.success(f"เพิ่มข้อมูลเข้าสู่ {data_type} เรียบร้อยแล้ว!")
            else:
                st.sidebar.error(f"โครงสร้างไฟล์ไม่ตรงกัน ขาดคอลัมน์: {missing_cols}")
    except Exception as e:
        st.sidebar.error(f"เกิดข้อผิดพลาดในการอ่านไฟล์: {e}")

# ==========================================
# MAIN PAGE: ส่วนแสดงผล Dashboard
# ==========================================
tab1, tab2 = st.tabs(["🧪 ELISA Data Analysis", "🩸 HI Data Analysis"])

# --- TAB 1: ELISA DATA ---
with tab1:
    st.header("สรุปผลการตรวจภูมิคุ้มกันด้วยวิธี ELISA")
    elisa_df = pd.read_csv(ELISA_DB_PATH)

    if elisa_df.empty:
        st.info("ยังไม่มีข้อมูลในระบบ ELISA กรุณาอัพโหลดข้อมูลที่แถบด้านซ้าย")
    else:
        # ตัวกรองข้อมูล (Filters)
        col1, col2, col3 = st.columns(3)
        with col1:
            farms = st.multiselect(
                "เลือกฟาร์ม (ELISA)",
                options=elisa_df["farm_name"].dropna().unique(),
                default=elisa_df["farm_name"].dropna().unique()[:1],
            )
        with col2:
            kits = st.multiselect(
                "เลือกชุดทดสอบ (Test Kit)",
                options=elisa_df["elisa_test_kit"].dropna().unique(),
                default=elisa_df["elisa_test_kit"].dropna().unique(),
            )
        with col3:
            houses = st.multiselect(
                "เลือกโรงเรือน (House)",
                options=elisa_df["House"].dropna().unique(),
                default=elisa_df["House"].dropna().unique(),
            )

        # กรองข้อมูลตามเงื่อนไข
        filtered_elisa = elisa_df[
            (elisa_df["farm_name"].isin(farms))
            & (elisa_df["elisa_test_kit"].isin(kits))
            & (elisa_df["House"].isin(houses))
        ]

        if not filtered_elisa.empty:
            # KPI Cards
            kpi1, kpi2, kpi3, kpi4 = st.columns(4)
            kpi1.metric("จำนวนตัวอย่างทั้งหมด", f"{len(filtered_elisa)} ตัวอย่าง")
            kpi2.metric(
                "ค่าเฉลี่ย Titer (Mean)",
                f"{pd.to_numeric(filtered_elisa['titer'], errors='coerce').mean():.2f}",
            )
            # คำนวณ %CV รวมกลุ่ม
            titer_col = pd.to_numeric(filtered_elisa["titer"], errors="coerce")
            cv_val = (titer_col.std() / titer_col.mean() * 100) if titer_col.mean() > 0 else 0
            kpi3.metric("ค่า %CV", f"{cv_val:.2f}%")

            # อัตราการตรวจพบผลบวก (Positive Rate)
            if "result" in filtered_elisa.columns:
                pos_count = len(filtered_elisa[filtered_elisa["result"].str.lower() == "positive"])
                pos_rate = (pos_count / len(filtered_elisa)) * 100
                kpi4.metric("เปอร์เซ็นต์ผลบวก (% Positive)", f"{pos_rate:.1f}%")

            # กราฟแนวโน้มระดับภูมิคุ้มกัน (Titer Profile by Age)
            st.subheader("📈 แนวโน้มระดับภูมิคุ้มกันแบ่งตามอายุ (Wk)")
            chart_data = (
                filtered_elisa.groupby(["Age (Wk)", "elisa_test_kit"])["titer"]
                .mean()
                .unstack()
                .fillna(0)
            )
            st.line_chart(chart_data)

            # ตารางข้อมูลสถิติแยกตามฝูง/อายุ
            st.subheader("📋 ตารางสรุปข้อมูลจำแนกตามฝูงและอายุ")
            summary_elisa = (
                filtered_elisa.groupby(["farm_name", "House", "Age (Wk)", "elisa_test_kit"])
                .agg(
                    Total_Samples=("titer", "count"),
                    Mean_Titer=("titer", "mean"),
                    Min_Titer=("titer", "min"),
                    Max_Titer=("titer", "max"),
                )
                .reset_index()
            )
            st.dataframe(summary_elisa, use_container_width=True)
        else:
            st.warning("ไม่พบข้อมูลตามเงื่อนไขที่เลือก")

# --- TAB 2: HI DATA ---
with tab2:
    st.header("สรุปผลการตรวจด้วยวิธี Haemagglutination Inhibition (HI)")
    hi_df = pd.read_csv(HI_DB_PATH)

    if hi_df.empty:
        st.info("ยังไม่มีข้อมูลในระบบ HI กรุณาอัพโหลดข้อมูลที่แถบด้านซ้าย")
    else:
        # ตัวกรองข้อมูล (Filters)
        col1, col2, col3 = st.columns(3)
        with col1:
            farms_hi = st.multiselect(
                "เลือกฟาร์ม (HI)",
                options=hi_df["farm_name"].dropna().unique(),
                default=hi_df["farm_name"].dropna().unique()[:1],
            )
        with col2:
            diseases = st.multiselect(
                "เลือกโรค (Disease)",
                options=hi_df["disease_name"].dropna().unique(),
                default=hi_df["disease_name"].dropna().unique(),
            )
        with col3:
            houses_hi = st.multiselect(
                "เลือกโรงเรือน (House - HI)",
                options=hi_df["House"].dropna().unique(),
                default=hi_df["House"].dropna().unique(),
            )

        # กรองข้อมูลตามเงื่อนไข
        filtered_hi = hi_df[
            (hi_df["farm_name"].isin(farms_hi))
            & (hi_df["disease_name"].isin(diseases))
            & (hi_df["House"].isin(houses_hi))
        ]

        if not filtered_hi.empty:
            # KPI Cards
            hkpi1, hkpi2, hkpi3 = st.columns(3)
            hkpi1.metric("จำนวนฝูงที่ตรวจ (ทดสอบ)", f"{len(filtered_hi)} ฝูง/บันทึก")

            # สำหรับ HI มักจะดูค่าเรขาคณิตเฉลี่ย (Geometric Mean Titer: GMT)
            gmt_val = pd.to_numeric(filtered_hi["GMT"], errors="coerce").mean()
            hkpi2.metric("ค่าเฉลี่ย GMT", f"{gmt_val:.2f}")

            cv_hi_val = pd.to_numeric(filtered_hi["CV"], errors="coerce").mean()
            hkpi3.metric("ค่าเฉลี่ย %CV", f"{cv_hi_val:.2f}%")

            # กราฟเปรียบเทียบ GMT แต่ละช่วงอายุเทียบกับค่ามาตรฐาน
            st.subheader("📈 ระดับไตเตอร์ HI (GMT) แยกตามอายุสัปดาห์")
            hi_chart_data = (
                filtered_hi.groupby(["Age (Wk)", "disease_name"])["GMT"]
                .mean()
                .unstack()
                .fillna(0)
            )
            st.bar_chart(hi_chart_data)

            # ตารางแสดงข้อมูลดิบที่กรองแล้ว
            st.subheader("📋 ข้อมูลการทดสอบ HI")
            st.dataframe(
                filtered_hi[
                    ["farm_name", "House", "Age (Wk)", "disease_name", "GMT", "CV", "sample_count"]
                ],
                use_container_width=True,
            )
        else:
            st.warning("ไม่พบข้อมูลตามเงื่อนไขที่เลือก")
