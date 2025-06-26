import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Alignment, Font
from io import BytesIO
import zipfile
import re

st.set_page_config(page_title="Tách file chấm công", layout="wide")
st.title("Tách file chấm công từng nhân viên (Online, dữ liệu giờ chuẩn)")

uploaded_file = st.file_uploader("Chọn file Excel gốc (.xlsx)", type=["xlsx"])
if uploaded_file is not None:
    # Đọc file gốc, header dòng 6 (index 5)
    df = pd.read_excel(uploaded_file, sheet_name=0, header=5)

    # Xác định cột 'Vào lần 1'
    vao_lan_1_col = next((col for col in df.columns if "Vào lần 1" in str(col)), None)
    if vao_lan_1_col is None:
        st.error('Không tìm thấy cột "Vào lần 1" trong file!')
        st.stop()

    # Tìm tất cả các cột có kiểu giờ (thường là các cột có "Vào", "Ra", hoặc từ vị trí 'Vào lần 1' trở đi, đến khi gặp cột số)
    time_cols = []
    start_idx = df.columns.get_loc(vao_lan_1_col)
    for col in df.columns[start_idx:]:
        # Nếu là cột lương, tổng, hoặc cột số => dừng lại
        if re.search(r'lương|tổng|sum|%', str(col), re.IGNORECASE):
            break
        # Nếu tên cột có "Vào" hoặc "Ra", xác định là cột giờ
        if "Vào" in str(col) or "Ra" in str(col):
            time_cols.append(col)

    # Hàm chuẩn hóa giờ sang định dạng HH:mm:ss
    def fix_time(val):
        if pd.isna(val):
            return ""
        if isinstance(val, pd.Timestamp):
            return val.strftime("%H:%M:%S")
        if isinstance(val, (float, int)):  # Excel time format số thập phân
            try:
                total_seconds = int(round(val * 24 * 3600))
                h = total_seconds // 3600
                m = (total_seconds % 3600) // 60
                s = total_seconds % 60
                return f"{h:02}:{m:02}:{s:02}"
            except:
                return ""
        val_str = str(val)
        if re.match(r'^\d{1,2}:\d{2}$', val_str):
            return val_str + ":00"
        if re.match(r'^\d{1,2}:\d{2}:\d{2}$', val_str):
            return val_str
        # Nếu là số nguyên 4 chữ số kiểu 715 => 07:15:00
        if re.match(r'^\d{3,4}$', val_str):
            h = int(val_str[:-2])
            m = int(val_str[-2:])
            return f"{h:02}:{m:02}:00"
        return val_str

    # Áp dụng chuẩn hóa giờ cho các cột giờ
    for col in time_cols:
        df[col] = df[col].apply(fix_time)

    # Hàm kiểm tra từ "Vào lần 1" trở đi có dữ liệu
    def co_du_lieu_tu_vao_lan_1(row):
        idx = df.columns.get_loc(vao_lan_1_col)
        return any([not pd.isna(cell) and str(cell).strip() != '' for cell in row[idx:]])

    # Lọc dòng hợp lệ
    df_filtered = df[df.apply(co_du_lieu_tu_vao_lan_1, axis=1)]
    df_filtered = df_filtered[df_filtered['Mã NV'].notna() & df_filtered['Họ tên'].notna()]

    # Xử lý thêm cột thứ trong tuần
    def convert_day(date_str):
        try:
            d = pd.to_datetime(date_str, dayfirst=True)
            week_day = d.strftime('%A')
            weekday_map = {
                'Monday': 'Thứ 2',
                'Tuesday': 'Thứ 3',
                'Wednesday': 'Thứ 4',
                'Thursday': 'Thứ 5',
                'Friday': 'Thứ 6',
                'Saturday': 'Thứ 7',
                'Sunday': 'Chủ nhật'
            }
            return weekday_map.get(week_day, '')
        except:
            return ''

    df_filtered['Thứ'] = df_filtered['Ngày'].apply(convert_day)

    # Đưa cột "Thứ" sau cột "Ngày"
    cols = list(df_filtered.columns)
    if 'Thứ' in cols and 'Ngày' in cols:
        cols.insert(cols.index('Ngày') + 1, cols.pop(cols.index('Thứ')))
        df_filtered = df_filtered[cols]

    # Cho user xem dữ liệu đã lọc và chuẩn hóa giờ
    st.subheader("Dữ liệu đã lọc & chuẩn hóa giờ")
    st.dataframe(df_filtered, use_container_width=True, height=300)

    # Hàm auto format Excel
    def auto_format_excel(file_bytes):
        wb = openpyxl.load_workbook(file_bytes)
        for ws in wb.worksheets:
            for cell in ws[1]:
                cell.alignment = Alignment(wrap_text=True, vertical='center', horizontal='center')
                cell.font = Font(bold=True)
            for row in ws.iter_rows():
                for cell in row:
                    cell.alignment = Alignment(wrap_text=True, vertical='center')
            for column_cells in ws.columns:
                max_length = max(len(str(cell.value) if cell.value else "") for cell in column_cells)
                ws.column_dimensions[column_cells[0].column_letter].width = max_length + 2
        buf = BytesIO()
        wb.save(buf)
        buf.seek(0)
        return buf

    # Xuất file từng người, nén zip cho tải về
    if st.button("Tách file và xuất kết quả ZIP"):
        zip_buffer = BytesIO()
        with zipfile.ZipFile(zip_buffer, "w") as zip_file:
            for (ma_nv, ho_ten), group in df_filtered.groupby(['Mã NV', 'Họ tên']):
                if len(group) == 0:
                    continue
                numeric_cols = group.select_dtypes(include='number').columns
                total_row = {col: group[col].sum() if col in numeric_cols else '' for col in group.columns}
                total_row['Ngày'] = 'Tổng'
                if 'Thứ' in total_row: total_row['Thứ'] = ''
                group_with_total = pd.concat([group, pd.DataFrame([total_row], columns=group.columns)], ignore_index=True)
                # Xuất vào memory
                excel_buffer = BytesIO()
                group_with_total.to_excel(excel_buffer, index=False)
                excel_buffer.seek(0)
                # Format lại
                formatted_buf = auto_format_excel(excel_buffer)
                file_name = f'{ma_nv}_{ho_ten}'.replace(" ", "_").replace("/", "_") + '.xlsx'
                zip_file.writestr(file_name, formatted_buf.getvalue())
        zip_buffer.seek(0)
        st.success("Đã tách xong! Bấm để tải file zip toàn bộ kết quả.")
        st.download_button("Tải file ZIP kết quả", zip_buffer, "ketqua_tach_file.xlsx.zip", "application/zip")

    st.caption("Tất cả các cột giờ đã chuẩn hóa về dạng HH:mm:ss. Nếu dữ liệu đầu vào lỗi hoặc trống, kiểm tra lại cấu trúc file gốc!")

else:
    st.info("Vui lòng upload file Excel để bắt đầu.")
