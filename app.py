import streamlit as st
import pandas as pd
from io import BytesIO
import openpyxl
from openpyxl.styles import Alignment, Font
import datetime

st.set_page_config(page_title="Tách file chấm công", layout="wide")
st.title("Anh Jiang Đẹp Zai - Pro - toai kho")

uploaded_file = st.file_uploader("Chọn file Excel gốc (.xlsx)", type=["xlsx"])
if uploaded_file is not None:
    # Đọc file gốc bằng openpyxl để giữ nguyên sheet
    wb_goc = openpyxl.load_workbook(uploaded_file)
    ws_goc = wb_goc.active  # Sheet đầu tiên

    # Đọc bằng pandas để xử lý dữ liệu nhân viên (chỉ dùng để tách)
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

    st.subheader("Dữ liệu đã lọc (chỉ thêm cột Thứ, không chỉnh giờ, không đổi số liệu):")
    st.dataframe(df_filtered, use_container_width=True, height=350)

    if st.button("Tách và xuất Excel tổng (sheet 1 = gốc, sheet 2+ = từng nhân viên)"):
        # Tạo workbook mới, copy nguyên sheet gốc
        output = BytesIO()
        wb_new = openpyxl.Workbook()
        ws_new_goc = wb_new.active
        ws_new_goc.title = "Du_lieu_goc"
        # Copy từng cell giữ đúng dữ liệu gốc
        for row in ws_goc.iter_rows():
            ws_new_goc.append([cell.value for cell in row])

        # Các sheet nhân viên đã xử lý
        groupby_obj = list(df_filtered.groupby(['Mã NV', 'Họ tên']))
        total_nv = len(groupby_obj)
        for (ma_nv, ho_ten), group in groupby_obj:
            group = group.copy()
            # Tính tổng cho các cột có dữ liệu từ cột "Lương giờ 100%" trở đi
            total_row = {}
            for col in group.columns:
                if col in cols_sum and pd.api.types.is_numeric_dtype(group[col]):
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

        # Xóa sheet mặc định nếu còn (openpyxl luôn tạo 1 sheet "Sheet" khi WorkBook mới)
        if "Sheet" in wb_new.sheetnames and len(wb_new.sheetnames) > total_nv + 1:
            del wb_new["Sheet"]

        wb_new.save(output)
        output.seek(0)
        st.success(f"Đã tách xong! Tổng số nhân viên được tách sheet: **{total_nv}**")
        st.download_button("Tải file Excel tổng hợp (chuẩn 100% dữ liệu gốc!)", output, "output_tong_hop.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

else:
    st.info("Đút file lên đi để anh Jiang xử lý")
