import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
import math

st.set_page_config(page_title="Tách file chấm công", layout="wide")
st.title("Anh Jiang Đẹp Zai - Pro - toai kho")

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

def to_hhmm(val):
    if pd.isna(val) or val is None or str(val).strip() == "":
        return ""
    val_str = str(val).strip()
    if pd.Series([val_str]).str.match(r'^\d{1,2}:\d{1,2}$').bool():
        h, m = val_str.split(":")
        return f"{int(h):02}:{int(m):02}"
    if pd.Series([val_str]).str.match(r'^\d{1,2}:\d{1,2}:\d{1,2}$').bool():
        h, m, _ = val_str.split(":")
        return f"{int(h):02}:{int(m):02}"
    try:
        if isinstance(val, float) and 0 <= val < 1:
            total_minutes = int(round(val * 24 * 60))
            h, m = divmod(total_minutes, 60)
            return f"{int(h):02}:{int(m):02}"
    except: pass
    return val_str

def get_header_row_height(header, width=8, font_size=11):
    lines = []
    for cell in header:
        text = str(cell.value) if cell.value else ""
        line_len = max(1, int((width - 1) * 1.5))
        num_lines = math.ceil(len(text) / line_len)
        lines.append(num_lines)
    max_lines = max(lines)
    return max(24, max_lines * 15)

uploaded_file = st.file_uploader("Chọn file Excel gốc (.xlsx)", type=["xlsx"])
if uploaded_file is not None:
    df = pd.read_excel(uploaded_file, sheet_name=0, header=5)

    vao_lan_1_col = next((col for col in df.columns if "Vào lần 1" in str(col)), None)
    ra_lan_2_col = next((col for col in df.columns if "Ra lần 2" in str(col)), None)
    if vao_lan_1_col is None or ra_lan_2_col is None:
        st.error('Không tìm thấy cột "Vào lần 1" hoặc "Ra lần 2" trong file!')
        st.stop()
    idx_vao_lan_1 = df.columns.get_loc(vao_lan_1_col)
    idx_ra_lan_2 = df.columns.get_loc(ra_lan_2_col)

    if 'Ngày' not in df.columns:
        st.error("Không tìm thấy cột 'Ngày'!")
        st.stop()
    ngay_idx = list(df.columns).index('Ngày')
    def convert_day(date_val):
        try:
            d = pd.to_datetime(date_val, dayfirst=True)
            weekday_map = {
                0: 'Thứ 2', 1: 'Thứ 3', 2: 'Thứ 4', 3: 'Thứ 5', 4: 'Thứ 6', 5: 'Thứ 7', 6: 'Chủ nhật'
            }
            return weekday_map[d.weekday()]
        except: return ""
    df.insert(ngay_idx, "Thứ", df['Ngày'].apply(convert_day))

    cols_time = [col for col in df.columns if any(key in str(col) for key in ['Vào', 'Ra'])]
    for col in cols_time:
        df[col] = df[col].apply(to_hhmm)

    st.subheader("Dữ liệu đã lọc (giữ nguyên, thêm cột Thứ):")
    st.dataframe(df, use_container_width=True, height=350)

    if st.button("Tách & xuất Excel từng nhân viên (dòng trống vùng 'Vào lần 1' đến 'Ra lần 2' thì merge ghi 'Nghỉ')"):
        output = BytesIO()
        wb_new = openpyxl.Workbook()
        default_sheet = wb_new.active
        wb_new.remove(default_sheet)

        groupby_obj = list(df.groupby(['Mã NV', 'Họ tên']))
        total_nv = len(groupby_obj)
        count_nv = 0
        status = st.empty()
        progress = st.progress(0)
        black_fill = PatternFill(start_color="FF000000", end_color="FF000000", fill_type="solid")

        for (ma_nv, ho_ten), group in groupby_obj:
            # Kiểm tra nếu tất cả các dòng vùng 'Vào lần 1' trở đi đều trống thì bỏ qua nhân viên này
            check_region = group.iloc[:, idx_vao_lan_1:]
            if not check_region.notna().any(axis=None) and not (check_region != "").any(axis=None):
                continue
            count_nv += 1
            status.info(f"Đang xử lý nhân viên thứ {count_nv}/{total_nv}: **{ma_nv} - {ho_ten}**")
            progress.progress(count_nv / total_nv)

            group_with_total = group.copy()  # Không thêm dòng tổng vì không hợp với logic này

            # Ghi sheet NV với full dữ liệu
            sheet_name = f"{ma_nv}_{ho_ten}".replace(" ", "_").replace("/", "_")[:31]
            ws_nv = wb_new.create_sheet(title=sheet_name)
            ws_nv.append([safe_excel_value(col) for col in group_with_total.columns])
            for row in group_with_total.itertuples(index=False):
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

            # Định dạng các dòng còn lại + merge ghi "Nghỉ" nếu vùng "Vào lần 1" đến "Ra lần 2" đều trống
            for idx, row in enumerate(ws_nv.iter_rows(min_row=2), start=0):
                row_data = group_with_total.iloc[idx]
                region = row_data.iloc[idx_vao_lan_1:idx_ra_lan_2 + 1]
                if all((pd.isna(x) or str(x).strip() == "") for x in region):
                    # Merge các ô vùng này
                    start_col = ws_nv.cell(row=idx + 2, column=idx_vao_lan_1 + 1).column_letter
                    end_col = ws_nv.cell(row=idx + 2, column=idx_ra_lan_2 + 1).column_letter
                    ws_nv.merge_cells(f"{start_col}{idx + 2}:{end_col}{idx + 2}")
                    cell_ghi_nghi = ws_nv.cell(row=idx + 2, column=idx_vao_lan_1 + 1)
                    cell_ghi_nghi.value = "Nghỉ"
                    cell_ghi_nghi.fill = black_fill
                    cell_ghi_nghi.font = Font(bold=True, color="FFFFFF")
                    cell_ghi_nghi.alignment = Alignment(wrap_text=True, vertical='center', horizontal='center')
                    # Các ô còn lại vẫn căn giữa, không có giá trị
                    for c in row:
                        if c.column < idx_vao_lan_1 + 1 or c.column > idx_ra_lan_2 + 1:
                            c.alignment = Alignment(wrap_text=True, vertical='center')
                else:
                    for cell in row:
                        cell.alignment = Alignment(wrap_text=True, vertical='center')

        wb_new.save(output)
        output.seek(0)
        status.success(f"✅ Đã xử lý xong {count_nv} nhân viên hợp lệ!")
        progress.empty()
        st.success(f"Đã tách xong! Tổng số nhân viên được tách sheet: **{count_nv}**")
        st.download_button(
            "Tải file Excel tổng hợp (merge Nghỉ đúng chuẩn!)",
            output,
            "output_tong_hop.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

else:
    st.info("Đút file lên đi để anh Jiang xử lý")
