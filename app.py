import io
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Lens Dispatcher - Batch Print", page_icon="🖨️", layout="wide"
)

st.title("🖨️ Lens Dispatcher - Batch Print Ready")
st.write(
    "อัปโหลดไฟล์ `TH_Consolidated_Sheet1.xlsx` เพื่อ คลีนข้อมูล + Sort แยกสาขา + เซ็ตหน้ากระดาษพร้อมพิมพ์แนวนอน/ขาวดำรวดเดียวจบ"
)

uploaded_file = st.file_uploader(
    "เลือกไฟล์ Excel ดั้งเดิม", type=["xlsx", "xls"]
)


def process_and_format_for_print(uploaded_file):
    # 1. อ่านไฟล์ Excel และ Sort / Clean ข้อมูลด้วย Pandas
    df = pd.read_excel(uploaded_file)

    # (ตรงนี้ใส่ Logic การคลีนและจัดกลุ่มข้อมูลตามสาขาเหมือนแอปแรก)
    # สมมติการจัดเรียงข้อมูลพื้นฐาน
    if "Branch" in df.columns:
        df = df.sort_values(by=["Branch"]).reset_index(drop=True)

    # 2. เขียนข้อมูลลงใน Memory Buffer ด้วย OpenPyXL เพื่อจัด Page Setup
    output_buffer = io.BytesIO()

    with pd.ExcelWriter(output_buffer, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Print_Ready", index=False)

    output_buffer.seek(0)

    # 3. โหลด Workbook เพื่อปรับแต่ง Page Setup สำหรับการพิมพ์
    wb = openpyxl.load_workbook(output_buffer)
    ws = wb.active

    # ตั้งค่าหน้ากระดาษ: แนวนอน / A4 / ขาว-ดำ
    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.blackAndWhite = True

    # บีบความกว้างให้พอดี 1 หน้ากระดาษพอดี (Fit to 1 Page Wide)
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0  # ปล่อยความยาวไหลตามจำนวนรายการ

    # หัวตาราง (แถว 1) ให้ขึ้นซ้ำทุกหน้าอัตโนมัติ
    ws.print_title_rows = "1:1"

    # จัดการเส้นขอบและฟอนต์ขนาด 10pt
    thin_border = Border(
        left=Side(style="thin", color="000000"),
        right=Side(style="thin", color="000000"),
        top=Side(style="thin", color="000000"),
        bottom=Side(style="thin", color="000000"),
    )

    for row in ws.iter_rows():
        for cell in row:
            cell.font = Font(name="Calibri", size=10)
            cell.border = thin_border

    # บันทึกไฟล์กลับลง Buffer
    final_buffer = io.BytesIO()
    wb.save(final_buffer)
    final_buffer.seek(0)

    return final_buffer


if uploaded_file is not None:
    st.success("อัปโหลดไฟล์เรียบร้อยแล้ว!")
    if st.button("🚀 ประมวลผลและสร้างไฟล์พร้อมพิมพ์", type="primary"):
        with st.spinner("กำลังจัดเรียงข้อมูลและตั้งค่าหน้ากระดาษ..."):
            final_excel = process_and_format_for_print(uploaded_file)

            st.download_button(
                label="📥 ดาวน์โหลดไฟล์สำหรับสั่งพิมพ์ (Print-Ready Excel)",
                data=final_excel,
                file_name="Lens_Picking_BatchPrint.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
