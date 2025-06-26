import streamlit as st
import pandas as pd
from io import BytesIO
import openpyxl
from openpyxl.styles import Alignment, Font
import time

st.set_page_config(page_title="Tách file chấm công", layout="wide")
st.title("Anh Jiang Đẹp Zai - Pro - toai kho")

uploaded_file = st.file_uploader("Chọn file Excel gốc (.xlsx)", type=["xlsx"])
if uploaded_file is not None:
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

    # Thêm cột "Thứ" vào bên trái cột "Ngày"
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
    df_filtered.insert(ngay_idx, "Thứ", df_filtered['Ngày'].apply(convert_day))

    st.subheader("Dữ liệu đã lọc, thêm cột Thứ:")
    st.dataframe(df_filtered, use_container_width=True, height=350)

    # TÌM cột "Lương giờ 100%"
    col_luong_gio_100 = next((col for col in df_filtered.columns if "Lương giờ 100%" in str(col)), None)
    if col_luong_gio_100 is None:
        st.error("Báo Anh Giang Pro toai kho xử lý ngay hép hép")
        st.stop()
    idx_luong_gio_100 = df_filtered.columns.get_loc(col_luong_gio_100)
    cols_sum = df_filtered.columns[idx_luong_gio_100:]  # Tính tổng từ cột "Lương giờ 100%" trở đi

    if st.button("Tách và xuất Excel tổng (mỗi nhân viên 1 sheet)"):
        output = BytesIO()
        groupby_obj = list(df_filtered.groupby(['Mã NV', 'Họ tên']))
        total_nv = len(groupby_obj)
        count_nv = 0

        # Hiện trạng thái tổng số nhân viên và tiến trình đang xử lý
        status = st.empty()
        progress = st.progress(0)

        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            for (ma_nv, ho_ten), group in groupby_obj:
                count_nv += 1
                status.info(f"Đang xử lý nhân viên thứ {count_nv}/{total_nv}: **{ma_nv} - {ho_ten}**")
                progress.progress(count_nv / total_nv)

                # Tính tổng các cột từ "Lương giờ 100%" trở đi (chỉ cộng cột có dữ liệu)
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
                sheet_name = f"{ma_nv}_{ho_ten}".replace(" ", "_").replace("/", "_")[:31]
                group_with_total.to_excel(writer, sheet_name=sheet_name, index=False)

        output.seek(0)

        # Định dạng lại các sheet cho đẹp
        wb = openpyxl.load_workbook(output)
        for ws in wb.worksheets:
            # Định dạng tiêu đề
            for cell in ws[1]:
                cell.alignment = Alignment(wrap_text=True, vertical='center', horizontal='center')
                cell.font = Font(bold=True)
            # Căn giữa toàn bộ sheet
            for row in ws.iter_rows():
                for cell in row:
                    cell.alignment = Alignment(wrap_text=True, vertical='center')
            # Auto width từng cột
            for column_cells in ws.columns:
                length = max(len(str(cell.value) if cell.value else "") for cell in column_cells)
                ws.column_dimensions[column_cells[0].column_letter].width = length + 2
        output2 = BytesIO()
        wb.save(output2)
        output2.seek(0)

        status.success(f"✅ Đã xử lý xong {total_nv} nhân viên!")
        progress.empty()

        st.success(f"Đã tách xong! Tổng số nhân viên được tách sheet: **{total_nv}**")
        st.download_button("Cám ơn anh Giang đi rồi mà Tải file Excel về", output2, "output_tong_hop.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


else:
    st.info("Dí vào đây để anh JiangPro xử lý")
