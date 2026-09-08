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
        return 0.0
    try:
        return float(val)
    except ValueError:
        return 0.0


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

    # นิยามรูปแบบเส้นตารางแบบคลีน (ไม่มีสีพื้นหลัง)
    thin_black = Side(border_style="thin", color="000000")
    double_black = Side(border_style="double", color="000000")

    box_border = Border(
        left=thin_black, right=thin_black, top=thin_black, bottom=thin_black
    )
    header_border = Border(
        left=thin_black, right=thin_black, top=thin_black, bottom=thin_black
    )
    total_border = Border(
        left=thin_black, right=thin_black, top=thin_black, bottom=double_black
    )

    # --------------------------------------------------------------------------
    # Tab 1: Summary All Stores
    # --------------------------------------------------------------------------
    ws_summary = wb_out.create_sheet(title="Summary All Stores")
    ws_summary.views.sheetView[0].showGridLines = True

    # ตั้งค่ากระดาษพิมพ์แบบคลีนแนวนอน + ล็อคพิมพ์ 1 หน้า
    ws_summary.page_setup.orientation = ws_summary.ORIENTATION_LANDSCAPE
    ws_summary.page_setup.paperSize = ws_summary.PAPERSIZE_A4
    ws_summary.page_setup.blackAndWhite = True
    ws_summary.sheet_properties.pageSetUpPr.fitToPage = True
    ws_summary.page_setup.fitToWidth = 1
    ws_summary.page_setup.fitToHeight = 1
    ws_summary.sheet_properties.pageSetUpPr.horizontalCentered = True

    ws_summary.merge_cells("A1:E1")
    ws_summary["A1"] = (
        "สรุปรายการจัดส่งเลนส์ประจำรอบ (Store Dispatch Summary)"
    )
    ws_summary["A1"].font = Font(
        name="Cordia New", size=18, bold=True, color="000000"
    )
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
        cell.font = Font(name="Cordia New", size=13, bold=True, color="000000")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = header_border

    store_groups = list(stores_df.groupby("store_id", sort=False))

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
        cell.border = total_border

    ws_summary.column_dimensions["A"].width = 8
    ws_summary.column_dimensions["B"].width = 22
    ws_summary.column_dimensions["C"].width = 38
    ws_summary.column_dimensions["D"].width = 16
    ws_summary.column_dimensions["E"].width = 22

    # --------------------------------------------------------------------------
    # 3. คัดแยกประเภทสาขา
    # --------------------------------------------------------------------------
    priority_stores = []
    normal_stores = []

    for store_id, group in store_groups:
        st_cols = group["col_idx"].tolist()
        num_dos = len(st_cols)

        active_items_count = sum(
            1
            for r in range(item_start_row, item_end_row + 1)
            if any(parse_num(raw_df.iloc[r, c]) > 0 for c in st_cols)
        )

        if active_items_count > 28 or num_dos > 8:
            priority_stores.append((store_id, group, True))
        else:
            normal_stores.append((store_id, group, False))

    sorted_store_groups = priority_stores + normal_stores

    # --------------------------------------------------------------------------
    # 4. สร้าง Sheet รายสาขา
    # --------------------------------------------------------------------------
    for store_id, group, is_heavy in sorted_store_groups:
        store_name = (
            str(group["store_name"].iloc[0])
            if pd.notna(group["store_name"].iloc[0])
            else str(store_id)
        )
        sheet_title = f"{store_id} - {store_name}"[:28]

        store_cols_all = group["col_idx"].tolist()
        do_nums_all = group["do_number"].tolist()

        normal_dos_info = []
        heavy_dos_info = []

        for c_idx, do_n in zip(store_cols_all, do_nums_all):
            do_item_count = sum(
                1
                for r in range(item_start_row, item_end_row + 1)
                if parse_num(raw_df.iloc[r, c_idx]) > 0
            )
            if do_item_count > 28:
                heavy_dos_info.append((c_idx, do_n))
            else:
                normal_dos_info.append((c_idx, do_n))

        ordered_dos = normal_dos_info + heavy_dos_info
        ordered_do_nums = [x[1] for x in ordered_dos]
        num_dos = len(ordered_do_nums)

        final_rows_list = []

        for idx_do, (c_idx, do_n) in enumerate(ordered_dos):
            do_items = []
            for r_idx in range(item_start_row, item_end_row + 1):
                q_val = parse_num(raw_df.iloc[r_idx, c_idx])
                if q_val > 0:
                    pid_val = raw_df.iloc[r_idx, col_mapping["pid"]]
                    name_val = raw_df.iloc[r_idx, col_mapping["name"]]
                    sph_val = parse_num(raw_df.iloc[r_idx, col_mapping["sph"]])
                    cyl_val = parse_num(raw_df.iloc[r_idx, col_mapping["cyl"]])

                    do_items.append(
                        {
                            "pid": str(pid_val) if pd.notna(pid_val) else "",
                            "name": (
                                str(name_val) if pd.notna(name_val) else ""
                            ),
                            "sph": sph_val,
                            "cyl": cyl_val,
                            "qty": int(q_val) if q_val == int(q_val) else q_val,
                            "do_col_index": idx_do,
                        }
                    )

            do_items_sorted = sorted(
                do_items,
                key=lambda x: (x["name"], -x["sph"], -x["cyl"], x["pid"]),
            )

            for item in do_items_sorted:
                qty_array = [None] * num_dos
                qty_array[item["do_col_index"]] = item["qty"]
                final_rows_list.append(
                    {
                        "pid": item["pid"],
                        "name": item["name"],
                        "sph": item["sph"],
                        "cyl": item["cyl"],
                        "qtys": qty_array,
                    }
                )

        ws = wb_out.create_sheet(title=sheet_title)
        ws.views.sheetView[0].showGridLines = True

        # ตั้งค่าการจัดหน้ากระดาษ: แนวนอน A4 + ขาวดำ + บังคับ 1 หน้าสำหรับทุก Sheet
        ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
        ws.page_setup.paperSize = ws.PAPERSIZE_A4
        ws.page_setup.blackAndWhite = True
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 1
        ws.sheet_properties.pageSetUpPr.horizontalCentered = True

        curr_row = 1

        # Header ข้อมูลสาขา
        ws.cell(
            row=curr_row, column=1, value=f"Store ID: {store_id}"
        ).font = Font(name="Cordia New", size=11, bold=True)
        ws.cell(
            row=curr_row, column=3, value=f"Store Name: {store_name}"
        ).font = Font(name="Cordia New", size=11, bold=True)
        curr_row += 1

        # Header ตาราง (ไม่มีสีพื้นหลัง ใช้ตัวหนา + ตีเส้นขอบ)
        base_headers = ["Item PID", "Item Name", "SPH", "CYL"]
        for col_i, h_text in enumerate(base_headers, 1):
            cell = ws.cell(row=curr_row, column=col_i, value=h_text)
            cell.font = Font(
                name="Cordia New", size=11, bold=True, color="000000"
            )
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = header_border

        for idx_q, do_n in enumerate(ordered_do_nums):
            c_i = 5 + idx_q
            cell = ws.cell(row=curr_row, column=c_i, value=f"DO: {do_n}")
            cell.font = Font(
                name="Cordia New", size=11, bold=True, color="000000"
            )
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = header_border

        max_col_idx = 4 + num_dos
        curr_row += 1
        start_data_row = curr_row

        # เขียนข้อมูลรายการเลนส์
        for row_data in final_rows_list:
            ws.cell(row=curr_row, column=1, value=row_data["pid"]).alignment = (
                Alignment(horizontal="center")
            )
            ws.cell(
                row=curr_row, column=2, value=row_data["name"]
            ).alignment = Alignment(horizontal="left")

            sph_fmt = (
                f"{row_data['sph']:+.2f}" if row_data["sph"] != 0 else "0.00"
            )
            cyl_fmt = (
                f"{row_data['cyl']:+.2f}" if row_data["cyl"] != 0 else "0.00"
            )

            ws.cell(row=curr_row, column=3, value=sph_fmt).alignment = Alignment(
                horizontal="right"
            )
            ws.cell(row=curr_row, column=4, value=cyl_fmt).alignment = Alignment(
                horizontal="right"
            )

            for idx_q, q_val in enumerate(row_data["qtys"]):
                c_i = 5 + idx_q
                ws.cell(row=curr_row, column=c_i, value=q_val).alignment = (
                    Alignment(horizontal="right")
                )

            for c in range(1, max_col_idx + 1):
                cell = ws.cell(row=curr_row, column=c)
                cell.font = Font(name="Cordia New", size=11)
                cell.border = box_border

            curr_row += 1

        end_data_row = curr_row - 1

        # แถว Grand Total ท้ายตาราง
        ws.cell(row=curr_row, column=1, value="Grand Total").font = Font(
            name="Cordia New", size=11, bold=True
        )
        ws.cell(row=curr_row, column=1).alignment = Alignment(
            horizontal="center"
        )

        for idx_q in range(num_dos):
            col_idx = 5 + idx_q
            col_letter = get_column_letter(col_idx)

            if start_data_row <= end_data_row:
                sum_formula = (
                    f"=SUM({col_letter}{start_data_row}:{col_letter}{end_data_row})"
                )
            else:
                sum_formula = 0

            ws.cell(row=curr_row, column=col_idx, value=sum_formula).font = Font(
                name="Cordia New", size=11, bold=True
            )
            ws.cell(row=curr_row, column=col_idx).alignment = Alignment(
                horizontal="right"
            )

        for c in range(1, max_col_idx + 1):
            cell = ws.cell(row=curr_row, column=c)
            cell.border = total_border

        ws.column_dimensions["A"].width = 14
        ws.column_dimensions["B"].width = 24
        ws.column_dimensions["C"].width = 9
        ws.column_dimensions["D"].width = 9
        for idx_q in range(num_dos):
            ws.column_dimensions[get_column_letter(5 + idx_q)].width = 13

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
        2. กดปุ่ม <b>"ประมวลผลไฟล์"</b> เพื่อจัดลำดับ Sheet + Sort สายตาแยกตาม DO<br>
        3. ดาวน์โหลดไฟล์ Excel สรุปผลพร้อมนำไปใช้งานได้ทันที
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
        with st.spinner("⏳ กำลังจัดลำดับ Sheet และประมวลผลตารางข้อมูล..."):
            try:
                processed_data = process_excel(uploaded_file)
                st.success("✅ **ประมวลผลสำเร็จเรียบร้อย!**")
                st.markdown("<br>", unsafe_allow_html=True)

                st.download_button(
                    label="📥 ดาวน์โหลดไฟล์ Excel สรุปผล (Consolidated Lists)",
                    data=processed_data,
                    file_name="Consolidated_Picking_Lists_Clean_Print.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            except Exception as e:
                st.error(
                    f"❌ เกิดข้อผิดพลาดในการประมวลผล โปรดตรวจสอบโครงสร้างไฟล์: {e}"
                )
