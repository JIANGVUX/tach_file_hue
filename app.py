import streamlit as st
import pandas as pd
from io import BytesIO
import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
import numpy as np
import math

st.set_page_config(page_title="Tách file chấm công", layout="wide")
st.title("Tách sheet từng nhân viên chuẩn header dòng 6 - Dữ liệu dòng 7 trở đi")

def safe_excel_value(val):
    if pd.isna(val) or val is None:
        return ""
    if isinstance(val, float):
        return round(val, 2)
    if isinstance(val, (np.integer, int)):
        return int(val)
    if hasattr(val, "strftime"):
        return val.strftime("%Y-%m-%d")
    return str(val)

def get_header_row_height(header, width=8):
    lines = []
    for cell in header:
        text = str(cell.value) if cell.value else ""
        line_len = max(1, int((width - 1) * 1.5))
        num_lines = math.ceil(len(text) / line_len)
        lines.append(num_lines)
    max_lines = max(lines)
    return max(24, max_lines * 15)

uploaded_file = st.file_uploader("Chọn file Excel chấm công gốc (.xlsx)", type=["xlsx"])
if uploaded_file is not None:
    # Đọc đúng dòng header: dòng 6 -> header=5
    df = pd.read_excel(uploaded_file, header=5)
    st.write("Tên các cột:", list(df.columns))
    
    # Tìm vị trí cột "Vào lần 1" và "Ra lần 2"
    vao_lan_1_col = next((col for col in df.columns if "Vào lần 1" in str(col)), None)
    ra_lan_2_col = next((col for col in df.columns if "Ra lần 2" in str(col)), None)
    if vao_lan_1_col is None or ra_lan_2_col is None:
        st.error('Không tìm thấy cột "Vào lần 1" hoặc "Ra lần 2" trong file!')
        st.stop()
    idx_vao_lan_1 = df.columns.get_loc(vao_lan_1_col)
    idx_ra_lan_2 = df.columns.get_loc(ra_lan_2_col)

    # Đảm bảo có đủ cột nhóm
    if "Mã NV" not in df.columns or "Họ tên" not in df.columns:
        st.error('Thiếu cột "Mã NV" hoặc "Họ tên" trong file!')
        st.stop()

    # Nhóm theo nhân viên
    groupby_obj = list(df.groupby(['Mã NV', 'Họ tên']))
    total_nv = len(groupby_obj)
    count_nv = 0
    status = st.empty()
    progress = st.progress(0)
    yellow_fill = PatternFill(start_color="FFFFFF99", end_color="FFFFFF99", fill_type="solid")  # Vàng nhạt

    output = BytesIO()
    wb_new = openpyxl.Workbook()
    default_sheet = wb_new.active
    wb_new.remove(default_sheet)

    for (ma_nv, ho_ten), group in groupby_obj:
        # Lấy vùng dữ liệu từ cột "Vào lần 1" trở đi
        region = group.iloc[:, idx_vao_lan_1:]
        arr = pd.Series(region.values.ravel()).astype(str).str.strip()
        arr = arr[~arr.isin(["", "nan", "NaT", "None"])]
        if arr.empty:
            continue  # BỎ QUA nhân viên không có dữ liệu thực

        count_nv += 1
        status.info(f"Đang xử lý nhân viên thứ {count_nv}/{total_nv}: **{ma_nv} - {ho_ten}**")
        progress.progress(count_nv / total_nv)

        # Ghi sheet NV với full dữ liệu
        sheet_name = f"{ma_nv}_{ho_ten}".replace(" ", "_").replace("/", "_")[:31]
        ws_nv = wb_new.create_sheet(title=sheet_name)
        ws_nv.append([safe_excel_value(col) for col in group.columns])
        for row in group.itertuples(index=False):
            ws_nv.append([safe_excel_value(cell) for cell in row])

        # Định dạng header
        header_row = ws_nv[1]
        header_fill = PatternFill(start_color="FF8C1A", end_color="FF8C1A", fill_type="solid")
        for cell in header_row:
            cell.alignment = Alignment(wrap_text=True, vertical='center', horizontal='center')
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = header_fill

        ws_nv.freeze_panes = "A2"

        # Đặt width từng cột: từ "Vào lần 1" trở đi width 8, cột trước đó auto
        for i, column_cells in enumerate(ws_nv.columns):
            if i >= idx_vao_lan_1:
                ws_nv.column_dimensions[column_cells[0].column_letter].width = 8
            else:
                length = max(len(str(cell.value) if cell.value else "") for cell in column_cells)
                ws_nv.column_dimensions[column_cells[0].column_letter].width = min(length + 2, 35)

        ws_nv.row_dimensions[1].height = get_header_row_height(header_row, width=8)

        # BÔI VÀNG những ô thiếu trong vùng 'Vào lần 1' đến 'Ra lần 2'
        for idx, row in enumerate(ws_nv.iter_rows(min_row=2), start=0):
            row_data = group.iloc[idx]
            region_row = row_data.iloc[idx_vao_lan_1:idx_ra_lan_2 + 1]
            for offset, value in enumerate(region_row):
                cell = ws_nv.cell(row=idx + 2, column=idx_vao_lan_1 + 1 + offset)
                if pd.isna(value) or str(value).strip() in ["", "nan", "NaT", "None"]:
                    cell.fill = yellow_fill
            for cell in row:
                cell.alignment = Alignment(wrap_text=True, vertical='center')

    wb_new.save(output)
    output.seek(0)
    status.success(f"✅ Đã xử lý xong {count_nv} nhân viên hợp lệ!")
    progress.empty()
    st.success(f"Đã tách xong! Tổng số nhân viên được tách sheet: **{count_nv}**")
    st.download_button(
        "Tải file Excel tổng hợp (bôi vàng ô thiếu vùng vào lần 1 - ra lần 2)",
        output,
        "output_tong_hop.xlsx",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
else:
    st.info("Đưa file lên để tách nhé!")
