import streamlit as st
import pandas as pd
from io import BytesIO
import openpyxl
from openpyxl.styles import Alignment, Font

st.set_page_config(page_title="Tách file chấm công", layout="wide")
st.title("Anh Jiang Đẹp Zai - Pro - toai kho")

uploaded_file = st.file_uploader("Chọn file Excel gốc (.xlsx)", type=["xlsx"])
if uploaded_file is not None:
    df = pd.read_excel(uploaded_file, sheet_name=0, header=5)

    # Xác định các cột giờ (có "Vào" hoặc "Ra" trong tên cột, trừ các cột "Ngày", "Thứ")
    time_cols = [col for col in df.columns if ("Vào" in str(col) or "Ra" in str(col)) and "Ngày" not in str(col) and "Thứ" not in str(col)]

    def to_hhmm(val):
        if pd.isna(val): return ""
        # Nếu là time hoặc datetime => chỉ lấy giờ phút
        if hasattr(val, 'hour') and hasattr(val, 'minute'):
            return f"{int(val.hour):02}:{int(val.minute):02}"
        s = str(val).strip()
        # Nếu là dạng 07:09:00 thì cắt lấy 5 ký tự đầu
        if len(s) == 8 and s[2] == ':' and s[5] == ':':
            return s[:5]
        # Nếu là dạng 7:9 hoặc 07:09
        if ':' in s and len(s) <= 5:
            h, m = s.split(":")
            return f"{int(h):02}:{int(m):02}"
        # Nếu là số kiểu Excel time (float < 1)
        try:
            f = float(s)
            if 0 <= f < 1:
                h = int(f * 24)
                m = int(round((f * 24 - h) * 60))
                return f"{h:02}:{m:02}"
        except:
            pass
        return s

    # Sheet gốc giữ nguyên (tuy nhiên, để nhất quán, chuẩn hóa các cột giờ về chuỗi HH:MM luôn)
    for col in time_cols:
        df[col] = df[col].apply(to_hhmm)

    # ---- Xử lý như cũ ----
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

    # Cũng chuẩn hóa các cột giờ ở sheet từng nhân viên (giống sheet gốc)
    for col in time_cols:
        df_filtered[col] = df_filtered[col].apply(to_hhmm)

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

    col_luong_gio_100 = next((col for col in df_filtered.columns if "Lương giờ 100%" in str(col)), None)
    if col_luong_gio_100 is None:
        st.error("Báo Anh Giang Pro toai kho xử lý ngay hép hép")
        st.stop()
    idx_luong_gio_100 = df_filtered.columns.get_loc(col_luong_gio_100)
    cols_sum = df_filtered.columns[idx_luong_gio_100:]

    if st.button("Tách và xuất Excel tổng (sheet 1 = gốc, sheet 2+ = từng nhân viên)"):
        output = BytesIO()
        groupby_obj = list(df_filtered.groupby(['Mã NV', 'Họ tên']))
        total_nv = len(groupby_obj)

        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            # Sheet đầu: Gốc (giữ đúng định dạng giờ)
            df.to_excel(writer, sheet_name="Du_lieu_goc", index=False)
            # Các sheet nhân viên đã xử lý
            for (ma_nv, ho_ten), group in groupby_obj:
                # Chuẩn hóa cột giờ từng nhân viên (đề phòng bị pandas tự đổi format)
                for col in time_cols:
                    group[col] = group[col].apply(to_hhmm)
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

        # Format lại các sheet cho đẹp
        wb = openpyxl.load_workbook(output)
        for ws in wb.worksheets:
            for cell in ws[1]:
                cell.alignment = Alignment(wrap_text=True, vertical='center', horizontal='center')
                cell.font = Font(bold=True)
            for row in ws.iter_rows():
                for cell in row:
                    cell.alignment = Alignment(wrap_text=True, vertical='center')
            for column_cells in ws.columns:
                length = max(len(str(cell.value) if cell.value else "") for cell in column_cells)
                ws.column_dimensions[column_cells[0].column_letter].width = length + 2
        output2 = BytesIO()
        wb.save(output2)
        output2.seek(0)

        st.success(f"Đã tách xong! Tổng số nhân viên được tách sheet: **{total_nv}**")
        st.download_button("Cám ơn anh Giang đi rồi mà Tải file Excel về", output2, "output_tong_hop.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

else:
    st.info("Đút file lên đi để anh Jiang xử lý")
