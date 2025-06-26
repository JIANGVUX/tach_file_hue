import streamlit as st
import pandas as pd
from io import BytesIO
import openpyxl
from openpyxl.styles import Alignment, Font
import numpy as np

st.set_page_config(page_title="Tách file chấm công", layout="wide")
st.title("Anh Jiang Đẹp Zai - Pro - toai kho")

def to_hhmm(val):
    # Chuẩn hóa mọi kiểu về hh:mm
    if pd.isna(val): return ""
    # time/datetime
    if hasattr(val, "hour") and hasattr(val, "minute"):
        return f"{int(val.hour):02}:{int(val.minute):02}"
    s = str(val).strip()
    # dạng 07:09:00 -> 07:09
    if len(s) == 8 and s[2] == ':' and s[5] == ':':
        return s[:5]
    # dạng 7:9 hoặc 07:09
    if ':' in s and len(s) <= 5:
        parts = s.split(":")
        if len(parts) == 2 and all(part.isdigit() for part in parts):
            h, m = parts
            return f"{int(h):02}:{int(m):02}"
    # float kiểu excel time
    try:
        f = float(s)
        if 0 <= f < 1:
            h = int(f * 24)
            m = int(round((f * 24 - h) * 60))
            return f"{h:02}:{m:02}"
    except:
        pass
    return s

uploaded_file = st.file_uploader("Chọn file Excel gốc (.xlsx)", type=["xlsx"])
if uploaded_file is not None:
    # --- B1: Mở workbook gốc giữ nguyên các sheet gốc ---
    wb_goc = openpyxl.load_workbook(uploaded_file)
    sheet_names_goc = wb_goc.sheetnames

    # --- B2: Đọc pandas từ sheet đầu để xử lý ---
    uploaded_file.seek(0)
    df = pd.read_excel(uploaded_file, sheet_name=0, header=5)

    # Tìm cột 'Vào lần 1'
    vao_lan_1_col = next((col for col in df.columns if "Vào lần 1" in str(col)), None)
    if vao_lan_1_col is None:
        st.error('Không tìm thấy cột "Vào lần 1" trong file!')
        st.stop()
    idx_vao_lan_1 = df.columns.get_loc(vao_lan_1_col)
    cols_check = df.columns[idx_vao_lan_1:]

    # Lọc dòng có dữ liệu từ "Vào lần 1" trở đi
    def dong_co_du_lieu_tu_vao_lan_1(row):
        return row[cols_check].notna().any() and (row[cols_check] != "").any()
    df_filtered = df[df.apply(dong_co_du_lieu_tu_vao_lan_1, axis=1)]
    df_filtered = df_filtered[df_filtered['Mã NV'].notna() & df_filtered['Họ tên'].notna()]

    # Xác định vị trí cột "Ngày" để thêm "Thứ" vào trước
    if 'Ngày' not in df_filtered.columns:
        st.error("Không tìm thấy cột 'Ngày'!")
        st.stop()
    ngay_idx = list(df_filtered.columns).index('Ngày')

    def convert_day(date_val):
        try:
            d = pd.to_datetime(date_val, dayfirst=True)
            weekday_map = {
                0: 'Thứ 2',
                1: 'Thứ 3',
                2: 'Thứ 4',
                3: 'Thứ 5',
                4: 'Thứ 6',
                5: 'Thứ 7',
                6: 'Chủ nhật'
            }
            return weekday_map[d.weekday()]
        except:
            return ""

    # Thêm cột "Thứ" (không chỉnh sửa dữ liệu khác)
    df_filtered.insert(ngay_idx, "Thứ", df_filtered['Ngày'].apply(convert_day))

    # Xác định cột bắt đầu tính tổng
    col_luong_gio_100 = next((col for col in df_filtered.columns if "Lương giờ 100%" in str(col)), None)
    if col_luong_gio_100 is None:
        st.error("Báo Anh Giang Pro toai kho xử lý ngay hép hép")
        st.stop()
    idx_luong_gio_100 = df_filtered.columns.get_loc(col_luong_gio_100)
    cols_sum = df_filtered.columns[idx_luong_gio_100:]

    # Tìm cột giờ (có "Vào", "Ra" trong tên, loại "Ngày", "Thứ")
    time_cols = [col for col in df_filtered.columns if ("Vào" in str(col) or "Ra" in str(col)) and "Ngày" not in str(col) and "Thứ" not in str(col)]
    # Tìm cột số liệu (numeric): ngoại trừ cột giờ và cột text
    numeric_cols = [col for col in cols_sum if pd.api.types.is_numeric_dtype(df_filtered[col])]

    st.subheader("Dữ liệu đã lọc (chỉ thêm cột Thứ, không đổi số liệu!):")
    st.dataframe(df_filtered, use_container_width=True, height=350)

    if st.button("Tách và xuất Excel tổng (giữ sheet gốc, thêm sheet nhân viên)"):
        # --- B3: Tạo workbook mới, copy toàn bộ sheet gốc sang ---
        output = BytesIO()
        wb_new = openpyxl.Workbook()
        # Xóa sheet mặc định ban đầu
        if "Sheet" in wb_new.sheetnames:
            del wb_new["Sheet"]
        # Copy tất cả sheet gốc
        for name in sheet_names_goc:
            ws_copy = wb_new.create_sheet(name)
            ws_goc = wb_goc[name]
            for row in ws_goc.iter_rows(values_only=False):
                ws_copy.append([cell.value for cell in row])
        # Xoá sheet trống đầu nếu cần
        if len(wb_new.sheetnames) > len(sheet_names_goc):
            del wb_new["Sheet"]

        # --- B4: Thêm sheet nhân viên ---
        groupby_obj = list(df_filtered.groupby(['Mã NV', 'Họ tên']))
        total_nv = len(groupby_obj)
        for (ma_nv, ho_ten), group in groupby_obj:
            group = group.copy()
            # Chuẩn hóa cột giờ về hh:mm
            for col in time_cols:
                group[col] = group[col].apply(to_hhmm)
            # Làm tròn các cột số liệu
            for col in numeric_cols:
                group[col] = group[col].round(0).astype('Int64')  # Làm tròn tới số nguyên
            # Tính tổng cho các cột có dữ liệu từ cột "Lương giờ 100%" trở đi
            total_row = {}
            for col in group.columns:
                if col in numeric_cols:
                    if group[col].notna().any():
                        total_row[col] = group[col].sum()
                    else:
                        total_row[col] = ""
                else:
                    total_row[col] = ""
            total_row['Ngày'] = "Tổng"
            total_row['Thứ'] = ""
            group_with_total = pd.concat([group, pd.DataFrame([total_row], columns=group.columns)], ignore_index=True)

            # Ghi vào sheet mới
            sheet_name = f"{ma_nv}_{ho_ten}".replace(" ", "_").replace("/", "_")[:31]
            ws_nv = wb_new.create_sheet(title=sheet_name)
            ws_nv.append(list(group_with_total.columns))
            for row in group_with_total.itertuples(index=False):
                ws_nv.append(list(row))
            # Định dạng tiêu đề cho đẹp
            for cell in ws_nv[1]:
                cell.alignment = Alignment(wrap_text=True, vertical='center', horizontal='center')
                cell.font = Font(bold=True)
            for row in ws_nv.iter_rows():
                for cell in row:
                    cell.alignment = Alignment(wrap_text=True, vertical='center')
            for column_cells in ws_nv.columns:
                length = max(len(str(cell.value) if cell.value else "") for cell in column_cells)
                ws_nv.column_dimensions[column_cells[0].column_letter].width = length + 2

        wb_new.save(output)
        output.seek(0)
        st.success(f"Đã tách xong! Tổng số nhân viên được tách sheet: **{total_nv}** (giữ nguyên sheet gốc và format gốc)")
        st.download_button("Tải file Excel tổng hợp (chuẩn 100% dữ liệu gốc!)", output, "output_tong_hop.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

else:
    st.info("Đút file lên đi để anh Jiang xử lý")
