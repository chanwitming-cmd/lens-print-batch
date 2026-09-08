import io
import math
import openpyxl
import pandas as pd
import streamlit as st
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# ==============================================================================
# 1. ตั้งค่าและตกแต่งด้วย CSS สไตล์ Soft 3D Light Theme
# ==============================================================================
st.set_page_config(
    page_title="Optics Lens Dispatcher System",
    page_icon="👓",
    layout="centered",
)

custom_css = """
<style>
    @import url('https://fonts.googleapis.com/css2?family=Kanit:wght@300;400;500;600&display=swap');
    
    html, body, [class*="css"] {
        font-family: 'Kanit', sans-serif;
    }

    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}

    .stApp {
        background: linear-gradient(180deg, #EBF1F7 0%, #E2E9F1 100%);
    }

    .header-box {
        background: #FFFFFF;
        padding: 30px 20px;
        border-radius: 20px;
        color: #1E293B;
        text-align: center;
        margin-bottom: 25px;
        box-shadow: 0 10px 25px rgba(148, 163, 184, 0.35),
                    inset 0 2px 3px rgba(255, 255, 255, 0.8);
        border: 1px solid rgba(255, 255, 255, 0.6);
    }
    
    .header-icon {
        font-size: 48px;
        margin-bottom: 10px;
    }

    .header-title {
        font-size: 24px;
        font-weight: 600;
        color: #1E3A8A;
        margin-bottom: 6px;
    }

    .header-subtitle {
        font-size: 14px;
        color: #64748B;
    }

    .step-box {
        background: rgba(255, 255, 255, 0.7);
        backdrop-filter: blur(10px);
        border-radius: 16px;
        padding: 20px;
        margin-bottom: 25px;
        color: #334155;
        box-shadow: 0 4px 15px rgba(148, 163, 184, 0.2);
        border: 1px solid #FFFFFF;
    }

    [data-testid="stFileUploader"] {
        background: #FFFFFF;
        border-radius: 16px;
        padding: 15px;
        box-shadow: 0 8px 20px rgba(148, 163, 184, 0.25);
        border: 2px dashed #93C5FD;
    }

    div.stButton > button:first-child {
        background: linear-gradient(135deg, #2563EB 0%, #1D4ED8 100%);
        color: white;
        font-size: 16px;
        font-weight: 500;
        border: none;
        border-radius: 14px;
        padding: 14px 28px;
        width: 100%;
        transition: all 0.2s ease-in-out;
        box-shadow: 0 6px 18px rgba(37, 99, 235, 0.35);
    }
    
    div.stButton > button:first-child:hover {
        transform: translateY(-2px);
        box-shadow: 0 10px 22px rgba(37, 99, 235, 0.45);
        background: linear-gradient(135deg, #1D4ED8 0%, #1E40AF 100%);
    }

    div.stDownloadButton > button:first-child {
        background: linear-gradient(135deg, #10B981 0%, #059669 100%);
        color: white;
        font-size: 16px;
        font-weight: 500;
        border: none;
        border-radius: 14px;
        padding: 14px 28px;
        width: 100%;
        transition: all 0.2s ease-in-out;
        box-shadow: 0 6px 18px rgba(16, 185, 129, 0.35);
    }
    
    div.stDownloadButton > button:first-child:hover {
        transform: translateY(-2px);
        box-shadow: 0 10px 22px rgba(16, 185, 129, 0.45);
        background: linear-gradient(135deg, #059669 0%, #047857 100%);
    }
</style>
"""
st.markdown(custom_css, unsafe_allow_html=True)


# ==============================================================================
# 2. ฟังก์ชันประมวลผล Excel
# ==============================================================================
def parse_num(val):
    if pd.isna(val) or val == "" or val is None:
        return 0
    try:
        return float(val) if "." in str(val) else int(val)
    except ValueError:
        return 0


def process_excel(uploaded_file):
    wb_raw = openpyxl.load_workbook(uploaded_file)
    ws_raw = wb_raw.active
    raw_data = [list(row) for row in ws_raw.iter_rows(values_only=True)]
    raw_df = pd.DataFrame(raw_data)

    header_row_idx = None
    col_mapping = {}

    for r in range(min(10, len(raw_df))):
        row_vals = [
            str(v).strip().lower() if pd.notna(v) else "" for v in raw_df.iloc[r]
        ]
        if "item pid" in row_vals:
            header_row_idx = r
            for c, val in enumerate(row_vals):
                if "item pid" in val:
                    col_mapping["pid"] = c
                elif "item name" in val:
                    col_mapping["name"] = c
                elif "sph" in val:
                    col_mapping["sph"] = c
                elif "cyl" in val:
                    col_mapping["cyl"] = c
            break

    store_id_row = header_row_idx - 2
    store_name_row = header_row_idx - 1
    do_number_row = header_row_idx
    item_start_row = header_row_idx + 1

    item_end_row = len(raw_df) - 1
    for r in range(item_start_row, len(raw_df)):
        first_val = str(raw_df.iloc[r, 0]).strip().lower()
        if "grand total" in first_val or "total" in first_val:
            item_end_row = r - 1
            break

    first_store_col = max(col_mapping.values()) + 1
    total_cols = raw_df.shape[1]

    store_cols_data = []
    for c_idx in range(first_store_col, total_cols):
        st_id = raw_df.iloc[store_id_row, c_idx]
        st_name = raw_df.iloc[store_name_row, c_idx]
        do_num = raw_df.iloc[do_number_row, c_idx]

        if pd.notna(st_id) and "total" not in str(st_id).strip().lower():
            store_cols_data.append(
                {
                    "col_idx": c_idx,
                    "store_id": str(st_id).strip(),
                    "store_name": (
                        str(st_name).strip() if pd.notna(st_name) else ""
                    ),
                    "do_number": (
                        str(do_num).strip() if pd.notna(do_num) else ""
                    ),
                }
            )

    stores_df = pd.DataFrame(store_cols_data)

    wb_out = openpyxl.Workbook()
    wb_out.remove(wb_out.active)

    NAVY_FILL = PatternFill(
        start_color="1F4E78", end_color="1F4E78", fill_type="solid"
    )
    STEEL_FILL = PatternFill(
        start_color="2F5597", end_color="2F5597", fill_type="solid"
    )
    HEADER_FILL = PatternFill(
        start_color="D9E1F2", end_color="D9E1F2", fill_type="solid"
    )
    ZEBRA_FILL = PatternFill(
        start_color="F9FBFD", end_color="F9FBFD", fill_type="solid"
    )

    thin_side = Side(border_style="thin", color="D9D9D9")
    box_border = Border(
        left=thin_side, right=thin_side, top=thin_side, bottom=thin_side
    )
    header_border = Border(
        left=thin_side,
        right=thin_side,
        top=Side(border_style="medium", color="1F4E78"),
        bottom=Side(border_style="medium", color="1F4E78"),
    )

    # --------------------------------------------------------------------------
    # Tab 1: Summary All Stores
    # --------------------------------------------------------------------------
    ws_summary = wb_out.create_sheet(title="Summary All Stores")
    ws_summary.views.sheetView[0].showGridLines = True

    ws_summary.page_setup.orientation = ws_summary.ORIENTATION_LANDSCAPE
    ws_summary.page_setup.paperSize = ws_summary.PAPERSIZE_A4
    ws_summary.page_setup.fitToWidth = 1
    ws_summary.page_setup.fitToHeight = 0
    ws_summary.sheet_properties.pageSetUpPr.fitToPage = True
    ws_summary.print_title_rows = "1:3"
    ws_summary.sheet_properties.pageSetUpPr.horizontalCentered = True
    ws_summary.page_setup.blackAndWhite = True

    ws_summary.merge_cells("A1:E1")
    ws_summary["A1"] = (
        "สรุปรายการจัดส่งเลนส์ประจำรอบ (Store Dispatch Summary)"
    )
    ws_summary["A1"].font = Font(
        name="Cordia New", size=18, bold=True, color="FFFFFF"
    )
    ws_summary["A1"].fill = NAVY_FILL
    ws_summary["A1"].alignment = Alignment(
        horizontal="center", vertical="center"
    )

    headers_summary = [
        "ลำดับ",
        "รหัสสาขา (Store ID)",
        "ชื่อสาขา (Store Name)",
        "จำนวน DO",
        "จำนวนเลนส์รวม (ชิ้น)",
    ]
    for col_num, h_text in enumerate(headers_summary, 1):
        cell = ws_summary.cell(row=3, column=col_num, value=h_text)
        cell.font = Font(name="Cordia New", size=13, bold=True, color="FFFFFF")
        cell.fill = STEEL_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center")

    store_groups = stores_df.groupby("store_id", sort=False)

    row_idx = 4
    idx = 1
    for store_id, group in store_groups:
        store_name = (
            str(group["store_name"].iloc[0])
            if pd.notna(group["store_name"].iloc[0])
            else str(store_id)
        )
        num_dos = len(group)
        st_cols = group["col_idx"].tolist()
        total_pcs = sum(
            int(parse_num(raw_df.iloc[r, c]))
            for r in range(item_start_row, item_end_row + 1)
            for c in st_cols
        )

        ws_summary.cell(row=row_idx, column=1, value=idx).alignment = Alignment(
            horizontal="center"
        )
        ws_summary.cell(row=row_idx, column=2, value=store_id).alignment = (
            Alignment(horizontal="center")
        )
        ws_summary.cell(row=row_idx, column=3, value=store_name)
        ws_summary.cell(
            row=row_idx, column=4, value=num_dos
        ).alignment = Alignment(horizontal="right")
        ws_summary.cell(
            row=row_idx, column=5, value=total_pcs
        ).alignment = Alignment(horizontal="right")

        for c in range(1, 6):
            cell = ws_summary.cell(row=row_idx, column=c)
            cell.font = Font(name="Cordia New", size=12)
            cell.border = box_border
            if idx % 2 == 0:
                cell.fill = ZEBRA_FILL
        row_idx += 1
        idx += 1

    ws_summary.cell(row=row_idx, column=1, value="Grand Total").font = Font(
        name="Cordia New", size=12, bold=True
    )
    ws_summary.cell(row=row_idx, column=1).alignment = Alignment(
        horizontal="center"
    )
    ws_summary.cell(
        row=row_idx, column=4, value=f"=SUM(D4:D{row_idx-1})"
    ).font = Font(name="Cordia New", size=12, bold=True)
    ws_summary.cell(row=row_idx, column=4).alignment = Alignment(
        horizontal="right"
    )
    ws_summary.cell(
        row=row_idx, column=5, value=f"=SUM(E4:E{row_idx-1})"
    ).font = Font(name="Cordia New", size=12, bold=True)
    ws_summary.cell(row=row_idx, column=5).alignment = Alignment(
        horizontal="right"
    )

    for c in range(1, 6):
        cell = ws_summary.cell(row=row_idx, column=c)
        cell.fill = HEADER_FILL
        cell.border = header_border

    ws_summary.column_dimensions["A"].width = 8
    ws_summary.column_dimensions["B"].width = 22
    ws_summary.column_dimensions["C"].width = 38
    ws_summary.column_dimensions["D"].width = 16
    ws_summary.column_dimensions["E"].width = 22

    # --------------------------------------------------------------------------
    # Tabs รายสาขา (ปรับปรุงตาม Feedback)
    # --------------------------------------------------------------------------
    MAX_ROWS_PER_PAGE = 28
    MAX_DO_PER_PAGE = 4

    for store_id, group in store_groups:
        store_name = (
            str(group["store_name"].iloc[0])
            if pd.notna(group["store_name"].iloc[0])
            else str(store_id)
        )
        base_title = f"{store_id} - {store_name}"[:28]

        store_cols_all = group["col_idx"].tolist()
        do_nums_all = group["do_number"].tolist()

        total_dos = len(do_nums_all)
        num_do_chunks = math.ceil(total_dos / MAX_DO_PER_PAGE)

        for chunk_idx in range(num_do_chunks):
            start_do = chunk_idx * MAX_DO_PER_PAGE
            end_do = min((chunk_idx + 1) * MAX_DO_PER_PAGE, total_dos)

            do_nums = do_nums_all[start_do:end_do]
            current_cols = store_cols_all[start_do:end_do]
            num_dos = len(do_nums)

            # --- แก้ไขจุดที่ 2: ตัดเลนส์ที่ไม่มีการสั่งซื้อ (ยอดสั่งซื้อใน DO ชุดนี้เป็น 0) ออกทั้งหมด ---
            store_items_sorted = []
            for r_idx in range(item_start_row, item_end_row + 1):
                qty_subset = [
                    parse_num(raw_df.iloc[r_idx, c]) for c in current_cols
                ]
                if sum(qty_subset) == 0:
                    continue  # ข้ามเลนส์ที่ไม่มียอดสั่งซื้อทันที

                first_do_idx = next(
                    (i for i, q in enumerate(qty_subset) if q > 0), 999
                )
                store_items_sorted.append(
                    {
                        "r_idx": r_idx,
                        "first_do_idx": first_do_idx,
                        "pid": raw_df.iloc[r_idx, col_mapping["pid"]],
                        "name": raw_df.iloc[r_idx, col_mapping["name"]],
                        "sph": raw_df.iloc[r_idx, col_mapping["sph"]],
                        "cyl": raw_df.iloc[r_idx, col_mapping["cyl"]],
                        "qty_subset": qty_subset,
                    }
                )

            store_items_sorted.sort(key=lambda x: (x["first_do_idx"], x["r_idx"]))

            sheet_title = (
                base_title
                if num_do_chunks == 1
                else f"{base_title[:25]}_{chunk_idx+1}"
            )
            ws = wb_out.create_sheet(title=sheet_title)
            ws.views.sheetView[0].showGridLines = True

            ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
            ws.page_setup.paperSize = ws.PAPERSIZE_A4
            ws.page_setup.fitToWidth = 1
            ws.page_setup.fitToHeight = 0
            ws.sheet_properties.pageSetUpPr.fitToPage = True
            ws.sheet_properties.pageSetUpPr.horizontalCentered = True
            ws.page_setup.blackAndWhite = True

            ws.page_margins.left = 0.25
            ws.page_margins.right = 0.25
            ws.page_margins.top = 0.4
            ws.page_margins.bottom = 0.4

            curr_row = 1
            total_items = len(store_items_sorted)
            num_row_chunks = (
                math.ceil(total_items / MAX_ROWS_PER_PAGE)
                if total_items > 0
                else 1
            )

            # เก็บตำแหน่งบรรทัดเริ่มต้นข้อมูลของแต่ละหน้าเพื่อคำนวณ Grand Total รวมหน้าสุดท้าย
            all_page_start_rows = []
            all_page_end_rows = []

            for r_chunk in range(num_row_chunks):
                is_last_page = r_chunk == (num_row_chunks - 1)

                start_item = r_chunk * MAX_ROWS_PER_PAGE
                end_item = min(
                    (r_chunk + 1) * MAX_ROWS_PER_PAGE, total_items
                )
                chunk_items = store_items_sorted[start_item:end_item]

                # --- 1. ส่วน Header ประจำหน้า ---
                ws.cell(
                    row=curr_row, column=1, value=f"Store ID: {store_id}"
                ).font = Font(name="Cordia New", size=11, bold=True)
                ws.cell(
                    row=curr_row, column=3, value=f"Store Name: {store_name}"
                ).font = Font(name="Cordia New", size=11, bold=True)
                curr_row += 1

                base_headers = ["Item PID", "Item Name", "SPH", "CYL"]
                for col_i, h_text in enumerate(base_headers, 1):
                    cell = ws.cell(row=curr_row, column=col_i, value=h_text)
                    cell.font = Font(
                        name="Cordia New", size=11, bold=True, color="FFFFFF"
                    )
                    cell.fill = STEEL_FILL
                    cell.alignment = Alignment(
                        horizontal="center", vertical="center"
                    )

                for idx_q, do_n in enumerate(do_nums):
                    c_i = 5 + idx_q
                    cell = ws.cell(row=curr_row, column=c_i, value=f"DO: {do_n}")
                    cell.font = Font(
                        name="Cordia New", size=11, bold=True, color="FFFFFF"
                    )
                    cell.fill = STEEL_FILL
                    cell.alignment = Alignment(
                        horizontal="center", vertical="center"
                    )

                max_col_idx = 4 + num_dos
                curr_row += 1

                # --- 2. ส่วนข้อมูลสินค้า ---
                start_data_row = curr_row
                for item in chunk_items:
                    ws.cell(
                        row=curr_row, column=1, value=item["pid"]
                    ).alignment = Alignment(horizontal="center")
                    ws.cell(
                        row=curr_row, column=2, value=item["name"]
                    ).alignment = Alignment(horizontal="left")
                    ws.cell(
                        row=curr_row, column=3, value=item["sph"]
                    ).alignment = Alignment(horizontal="right")
                    ws.cell(
                        row=curr_row, column=4, value=item["cyl"]
                    ).alignment = Alignment(horizontal="right")

                    for idx_q, q_val in enumerate(item["qty_subset"]):
                        c_i = 5 + idx_q
                        ws.cell(
                            row=curr_row,
                            column=c_i,
                            value=q_val if q_val > 0 else None,
                        ).alignment = Alignment(horizontal="right")

                    for c in range(1, max_col_idx + 1):
                        cell = ws.cell(row=curr_row, column=c)
                        cell.font = Font(name="Cordia New", size=11)
                        cell.border = box_border
                        if (curr_row - start_data_row) % 2 == 1:
                            cell.fill = ZEBRA_FILL

                    curr_row += 1

                end_data_row = curr_row - 1
                if start_data_row <= end_data_row:
                    all_page_start_rows.append(start_data_row)
                    all_page_end_rows.append(end_data_row)

                # --- แก้ไขจุดที่ 1: แสดง Grand Total เฉพาะหน้าสุดท้ายเท่านั้น ---
                if is_last_page:
                    ws.cell(
                        row=curr_row, column=1, value="Grand Total"
                    ).font = Font(name="Cordia New", size=11, bold=True)
                    ws.cell(row=curr_row, column=1).alignment = Alignment(
                        horizontal="center"
                    )

                    for idx_q in range(num_dos):
                        col_idx = 5 + idx_q
                        col_letter = get_column_letter(col_idx)

                        if all_page_start_rows:
                            # รวมสูตร SUM ทุกช่วงข้อมูลตั้งแต่หน้าแรกจนถึงหน้าสุดท้าย
                            sum_parts = [
                                f"{col_letter}{s}:{col_letter}{e}"
                                for s, e in zip(
                                    all_page_start_rows, all_page_end_rows
                                )
                            ]
                            sum_formula = f"=SUM({','.join(sum_parts)})"
                        else:
                            sum_formula = 0

                        ws.cell(
                            row=curr_row, column=col_idx, value=sum_formula
                        ).font = Font(name="Cordia New", size=11, bold=True)
                        ws.cell(row=curr_row, column=col_idx).alignment = (
                            Alignment(horizontal="right")
                        )

                    for c in range(1, max_col_idx + 1):
                        cell = ws.cell(row=curr_row, column=c)
                        cell.fill = HEADER_FILL
                        cell.border = header_border
                else:
                    # ถ้าไม่ใช่หน้าสุดท้าย ใส่จุดแบ่งหน้า (Page Break)
                    ws.row_breaks.append(
                        openpyxl.worksheet.pagebreak.Break(id=curr_row - 1)
                    )

            # ตั้งค่าความกว้างคอลัมน์มาตรฐาน
            ws.column_dimensions["A"].width = 14
            ws.column_dimensions["B"].width = 24
            ws.column_dimensions["C"].width = 9
            ws.column_dimensions["D"].width = 9
            for idx_q in range(num_dos):
                ws.column_dimensions[
                    get_column_letter(5 + idx_q)
                ].width = 14

    for sheet in wb_out.worksheets:
        sheet.sheet_view.tabSelected = True

    output = io.BytesIO()
    wb_out.save(output)
    output.seek(0)
    return output


# ==============================================================================
# 3. ส่วนการจัดวางหน้าตาเว็บ (UI Layout)
# ==============================================================================
st.markdown(
    """
    <div class="header-box">
        <div class="header-icon">🌐</div>
        <div class="header-title">Optics Lens Dispatcher System</div>
        <div class="header-subtitle">ระบบจัดกลุ่มและสรุปรายการจัดส่งเลนส์แยกสาขาอัตโนมัติ</div>
    </div>
""",
    unsafe_allow_html=True,
)

st.markdown(
    """
    <div class="step-box">
        <b>🔹 ขั้นตอนการทำงาน:</b><br>
        1. อัปโหลดไฟล์ <code>TH_Consolidated_Sheet1.xlsx</code> ในช่องด้านล่าง<br>
        2. กดปุ่ม <b>"ประมวลผลไฟล์"</b> เพื่อเริ่มจัดกลุ่มตามสาขา<br>
        3. ดาวน์โหลดไฟล์ Excel พร้อมนำไปใช้งานได้ทันที
    </div>
""",
    unsafe_allow_html=True,
)

uploaded_file = st.file_uploader(
    "เลือกไฟล์ Excel ต้นฉบับ (.xlsx)", type=["xlsx"]
)

if uploaded_file is not None:
    st.info(f"📄 **ไฟล์ที่เลือก:** `{uploaded_file.name}`")

    if st.button("🚀 ประมวลผลและแปลงไฟล์"):
        with st.spinner("⏳ กำลังจัดระเบียบตารางและคำนวณยอด..."):
            try:
                processed_data = process_excel(uploaded_file)
                st.success("✅ **ประมวลผลสำเร็จเรียบร้อย!**")
                st.markdown("<br>", unsafe_allow_html=True)

                st.download_button(
                    label="📥 ดาวน์โหลดไฟล์ Excel สรุปผล (Clean Print Layout)",
                    data=processed_data,
                    file_name="Consolidated_Picking_Lists_Clean.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            except Exception as e:
                st.error(
                    f"❌ เกิดข้อผิดพลาดในการประมวลผล โปรดตรวจสอบโครงสร้างไฟล์: {e}"
                )
