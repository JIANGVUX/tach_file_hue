import streamlit as st
import pandas as pd
from io import BytesIO
import zipfile
import re

st.set_page_config(page_title="Tách file chấm công", layout="wide")
st.title("Tách file chấm công từng nhân viên (Giờ chỉ HH:mm, không bao giờ có giây)")

uploaded_file = st.file_uploader("Chọn file Excel gốc (.xlsx)", type=["xlsx"])
if uploaded_file is not None:
    df = pd.read_excel(uploaded_file, sheet_name=0, header=5)

    # Thêm cột "Thứ" vào bên trái cột "Ngày"
    if 'Ngày' not in df.columns:
        st.error("Không tìm thấy cột 'Ngày'!")
        st.stop()
    ngay_idx = list(df.columns).index('Ngày')

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
    df.insert(ngay_idx, "Thứ", df['Ngày'].apply(convert_day))

    # CHUẨN HOÁ toàn bộ các cột giờ về đúng HH:mm, KHÔNG BAO GIỜ có giây!
    def to_hhmm_only(val):
        if pd.isna(val):
            return ""
        val_str = str(val).strip()
        if re.match(r'^\d{1,2}:\d{2}$', val_str):
            h, m = map(int, val_str.split(":"))
            return f"{h:02}:{m:02}"
        if re.match(r'^\d{1,2}:\d{2}:\d{2}(\.\d+)?$', val_str):
            h, m, *_ = val_str.split(":")
            return f"{int(h):02}:{int(m):02}"
        if re.match(r'^\d{3,4}$', val_str):
            h = int(val_str[:-2])
            m = int(val_str[-2:])
            return f"{h:02}:{m:02}"
        try:
            val_float = float(val)
            if 0 <= val_float < 1:
                total_seconds = int(round(val_float * 24 * 3600))
                h = total_seconds // 3600
                m = (total_seconds % 3600) // 60
                return f"{h:02}:{m:02}"
        except:
            pass
        return val_str

    # Áp dụng cho các cột giờ (chứa "Vào" hoặc "Ra" trong tên cột)
    time_cols = [col for col in df.columns if any(key in str(col) for key in ['Vào', 'Ra'])]
    for col in time_cols:
        df[col] = df[col].apply(to_hhmm_only)

    st.subheader("Dữ liệu đã chuẩn hóa giờ chỉ còn HH:mm:")
    st.dataframe(df, use_container_width=True, height=350)

    if st.button("Tách file và xuất ZIP"):
        zip_buffer = BytesIO()
        with zipfile.ZipFile(zip_buffer, "w") as zip_file:
            for (ma_nv, ho_ten), group in df.groupby(['Mã NV', 'Họ tên']):
                if len(group) == 0:
                    continue
                total_row = {col: group[col].sum() if pd.api.types.is_numeric_dtype(group[col]) else "" for col in group.columns}
                total_row['Ngày'] = "Tổng"
                total_row['Thứ'] = ""
                group_with_total = pd.concat([group, pd.DataFrame([total_row], columns=group.columns)], ignore_index=True)

                excel_buffer = BytesIO()
                group_with_total.to_excel(excel_buffer, index=False)
                excel_buffer.seek(0)

                file_name = f"{ma_nv}_{ho_ten}".replace(" ", "_").replace("/", "_") + ".xlsx"
                zip_file.writestr(file_name, excel_buffer.getvalue())
        zip_buffer.seek(0)
        st.success("Đã tách xong! Bấm để tải file zip toàn bộ kết quả.")
        st.download_button("Tải file ZIP kết quả", zip_buffer, "ketqua_tach_file.xlsx.zip", "application/zip")

    st.caption("Giờ/phút tất cả các cột chỉ còn HH:mm, không bao giờ có giây (00 hoặc bất kỳ số nào).")

else:
    st.info("Vui lòng upload file Excel để bắt đầu.")
