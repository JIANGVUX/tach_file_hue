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
    if vao_lan_1_col is None:
        st.error('Không tìm thấy cột "Vào lần 1" trong file!')
        st.stop()
    idx_vao_lan_1 = df.columns.get_loc(vao_lan_1_col)
    cols_check = df.columns[idx_vao_lan_1:]

    def dong_co_du_lieu_tu_vao_lan_1(row):
        return row[cols_check].notna().any() and (row[cols_check] != "").any()
    df_filtered = df[df.apply(dong_co_du_lieu_tu_vao_lan_1, axis=1)]
    df_filtered = df_filtered[df_filtered['Mã NV'].notna() & df_filtered['Họ tên'].notna()]

    if 'Ngày' not in df_filtered.columns:
        st.error("Không tìm thấy cột 'Ngày'!")
        st.stop()
    ngay_idx = list(df_filtered.columns).index('Ngày')
    def convert_day(date_val):
        try:
            d = pd.to_datetime(date_val, dayfirst=True)
            weekday_map = {
                0: 'Thứ 2', 1: 'Thứ 3', 2: 'Thứ 4', 3: 'Thứ 5', 4: 'Thứ 6', 5: 'Thứ 7', 6: 'Chủ nhật'
            }
            return weekday_map[d.weekday()]
        except: return ""
    df_filtered.insert(ngay_idx, "Thứ", df_filtered['Ngày'].apply(convert_day))

    cols_time = [col for col in df_filtered.columns if any(key in str(col) for key in ['Vào', 'Ra'])]
    for col in cols_time:
        df_filtered[col] = df_filtered[col].apply(to_hhmm)

    col_luong_gio_100 = next((col for col in df_filtered.columns if "Lương giờ 100%" in str(col)), None)
    if col_luong_gio_100 is None:
        st.error("Báo Anh Giang Pro toai kho xử lý ngay hép hép")
        st.stop()
    idx_luong_gio_100 = df_filtered.columns.get_loc(col_luong_gio_100)
    cols_sum = df_filtered.columns[idx_luong_gio_100:]

    st.subheader("Dữ liệu đã lọc (chuẩn giờ phút, giữ nguyên dữ liệu, thêm cột Thứ):")
    st.dataframe(df_filtered, use_container_width=True, height=350)

    if st.button("Tách & xuất Excel từng nhân viên 1 sheet (full dữ liệu, dòng toàn bộ vùng lương giờ 100% trống sẽ bôi xám!)"):
        output = BytesIO()
        wb_new = openpyxl.Workbook()
        default_sheet = wb_new.active
        wb_new.remove(default_sheet)

        groupby_obj = list(df_filtered.groupby(['Mã NV', 'Họ tên']))
        total_nv = len(groupby_obj)
        count_nv = 0
        status = st.empty()
        progress = st.progress(0)
        gray_fill = PatternFill(start_color="FFD9D9D9", end_color="FFD9D9D9", fill_type="solid")  # Màu xám nhạt

        for (ma_nv, ho_ten), group in groupby_obj:
            count_nv += 1
            status.info(f"Đang xử lý nhân viên thứ {count_nv}/{total_nv}: **{ma_nv} - {ho_ten}**")
            progress.progress(count_nv / total_nv)

            # ===== Thêm dòng tổng cuối (giữ lại logic cũ) =====
            total_row = {}
            for col in group.columns:
                if col in cols_sum and pd.api.types.is_numeric_dtype(group[col]):
                    if group[col].notna().any():
                        val = group[col].sum()
                        if isinstance(val, float):
                            total_row[col] = round(val, 2)
                        else:
                            total_row[col] = val
                    else:
                        total_row[col] = ""
                else:
                    total_row[col] = ""
            total_row['Ngày'] = "Tổng"
            total_row['Thứ'] = ""
            group_with_total = pd.concat([group, pd.DataFrame([total_row], columns=group.columns)], ignore_index=True)

            # ==== Ghi sheet NV với full dữ liệu ====
            sheet_name = f"{ma_nv}_{ho_ten}".replace(" ", "_").replace("/", "_")[:31]
            ws_nv = wb_new.create_sheet(title=sheet_name)
            ws_nv.append([safe_excel_value(col) for col in group_with_total.columns])
            for row in group_with_total.itertuples(index=False):
                ws_nv.append([safe_excel_value(cell) for cell in row])

            # ==== Định dạng header ====
            header_row = ws_nv[1]
            header_fill = PatternFill(start_color="FF8C1A", end_color="FF8C1A", fill_type="solid")
            for cell in header_row:
                cell.alignment = Alignment(wrap_text=True, vertical='center', horizontal='center')
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = header_fill

            ws_nv.freeze_panes = "A2"

            # Đặt width từng cột: từ "Vào lần 1" trở đi width 8, cột trước đó auto (giữ đẹp)
            for i, column_cells in enumerate(ws_nv.columns):
                if i >= idx_vao_lan_1:
                    ws_nv.column_dimensions[column_cells[0].column_letter].width = 8
                else:
                    length = max(len(str(cell.value) if cell.value else "") for cell in column_cells)
                    ws_nv.column_dimensions[column_cells[0].column_letter].width = min(length + 2, 35)

            ws_nv.row_dimensions[1].height = get_header_row_height(header_row, width=8)

            # ==== Định dạng các dòng còn lại + BÔI XÁM nếu dòng toàn bộ vùng lương giờ 100% trở đi là rỗng ====
            for idx, row in enumerate(ws_nv.iter_rows(min_row=2), start=0):
                # Dữ liệu gốc của dòng này ở group_with_total
                row_data = group_with_total.iloc[idx]
                region = row_data.iloc[idx_luong_gio_100:]
                if all((pd.isna(x) or str(x).strip() == "") for x in region):
                    # Toàn bộ vùng "Lương giờ 100%" trở đi là trống, bôi xám dòng
                    for cell in row:
                        cell.fill = gray_fill
                # Định dạng wrap text và căn giữa cho tất cả
                for cell in row:
                    cell.alignment = Alignment(wrap_text=True, vertical='center')

        wb_new.save(output)
        output.seek(0)
        status.success(f"✅ Đã xử lý xong {total_nv} nhân viên!")
        progress.empty()
        st.success(f"Đã tách xong! Tổng số nhân viên được tách sheet: **{total_nv}**")
        st.download_button(
            "Tải file Excel tổng hợp (full dữ liệu, dòng vùng lương giờ 100% trống bôi xám!)",
            output,
            "output_tong_hop.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

else:
    st.info("Đút file lên đi để anh Jiang xử lý")
