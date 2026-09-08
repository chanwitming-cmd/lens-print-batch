import io
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
import pandas as pd
import streamlit as st

# ตั้งค่าหน้าเว็บ
st.set_page_config(
    page_title="Optics Lens Dispatcher System",
    page_icon="🌐",
    layout="centered",
)

# Custom CSS เพื่อแต่งหน้าตา (UI Theme) ให้เหมือนรูปต้นฉบับ
st.markdown(
    """
    <style>
    /* พื้นหลังหลักของแอป */
    .stApp {
        background-color: #EBF1F6;
    }
    
    /* ซ่อน Header/Footer หลักของ Streamlit */
    header {visibility: hidden;}
    footer {visibility: hidden;}
    
    /* กล่อง Card สไตล์โค้งมน */
    .custom-card {
        background-color: #FFFFFF;
        padding: 30px;
        border-radius: 16px;
        box-shadow: 0 4px 12px rgba(0, 0, 0, 0.03);
        margin-bottom: 20px;
        text-align: center;
    }
    
    .instruction-card {
        background-color: #F8FAFC;
        padding: 24px;
        border-radius: 16px;
        border: 1px solid #E2E8F0;
        margin-bottom: 20px;
        text-align: left;
    }
    
    /* หัวข้อหลัก */
    .main-title {
        color: #1E3A8A;
        font-size: 26px;
        font-weight: 700;
        margin-top: 10px;
        margin-bottom: 8px;
    }
    
    .sub-title {
        color: #64748B;
        font-size: 15px;
        font-weight: 500;
    }
    
    /* ข้อความขั้นตอน */
    .instruction-title {
        color: #1E3A8A;
        font-size: 16px;
        font-weight: 700;
        margin-bottom: 12px;
    }
    
    .instruction-step {
        color: #334155;
        font-size: 14px;
        line-height: 1.8;
    }
    
    .code-span {
        background-color: #F1F5F9;
        color: #16A34A;
        padding: 2px 6px;
        border-radius: 4px;
        font-family: monospace;
        font-weight: 600;
    }
    
    /* ไอคอนไอคอนโลก */
    .icon-container {
        font-size: 48px;
        line-height: 1;
    }
    </style>
""",
    unsafe_allow_html=True,
)

# --- Header Section ---
st.markdown(
    """
    <div class="custom-card">
        <div class="icon-container">🌐</div>
        <div class="main-title">Optics Lens Dispatcher System</div>
        <div class="sub-title">ระบบจัดกลุ่มและสรุปรายการจัดส่งเลนส์แยกสาขาอัตโนมัติ (Print-Ready)</div>
    </div>
""",
    unsafe_allow_html=True,
)

# --- Instruction Section ---
st.markdown(
    """
    <div class="instruction-card">
        <div class="instruction-title">🔹 ขั้นตอนการทำงาน:</div>
        <div class="instruction-step">1. อัปโหลดไฟล์ <span class="code-span">TH_Consolidated_Sheet1.xlsx</span> ในช่องด้านล่าง</div>
        <div class="instruction-step">2. กดปุ่ม "ประมวลผลไฟล์" เพื่อจัดกลุ่มตามสาขาและเซ็ตหน้ากระดาษพิมพ์</div>
        <div class="instruction-step">3. ดาวน์โหลดไฟล์ Excel พร้อมนำไปสั่งพิมพ์รวดเดียวจบได้ทันที</div>
    </div>
""",
    unsafe_allow_html=True,
)

# --- File Uploader Section ---
uploaded_file = st.file_uploader(
    "เลือกไฟล์ Excel ต้นฉบับ (.xlsx)", type=["xlsx", "xls"]
)


def process_and_format_for_print(uploaded_file):
    df = pd.read_excel(uploaded_file)

    # จัดเรียงข้อมูลตามสาขา (ถ้ามีคอลัมน์ Branch)
    if "Branch" in df.columns:
        df = df.sort_values(by=["Branch"]).reset_index(drop=True)

    output_buffer = io.BytesIO()
    with pd.ExcelWriter(output_buffer, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Print_Ready", index=False)

    output_buffer.seek(0)
    wb = openpyxl.load_workbook(output_buffer)
    ws = wb.active

    # เซ็ตค่าการพิมพ์: แนวนอน / ขาว-ดำ / Fit to 1 Page Wide
    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.blackAndWhite = True

    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.print_title_rows = "1:1"

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

    final_buffer = io.BytesIO()
    wb.save(final_buffer)
    final_buffer.seek(0)
    return final_buffer


# --- Processing Button & Actions ---
if uploaded_file is not None:
    st.write("")
    if st.button("🚀 ประมวลผลไฟล์", type="primary", use_container_width=True):
        with st.spinner("กำลังจัดเรียงข้อมูลและตั้งค่าหน้ากระดาษ..."):
            final_excel = process_and_format_for_print(uploaded_file)

            st.success("ประมวลผลเรียบร้อยแล้ว!")
            st.download_button(
                label="📥 ดาวน์โหลดไฟล์ Excel พร้อมพิมพ์",
                data=final_excel,
                file_name="Lens_Picking_PrintReady.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
