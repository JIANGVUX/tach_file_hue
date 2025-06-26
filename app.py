import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Alignment, Font
from io import BytesIO
import zipfile

st.set_page_config(page_title="Tách file chấm công", layout="wide")
st.title("Tách file chấm công từng nhân viên (GIỮ NGUYÊN DỮ LIỆU GỐC)")

uploaded_file = st.file_uploader("Chọn file Excel gốc (.xlsx)", type=["xlsx"])
if uploaded_file is not None:
    # Đọc file, header đúng dòng thực tế của bạn (thường là 5 hoặc 0)
    df = pd.read_excel(uploaded_file, sheet_name=0, header=5)

    # Tìm vị trí cột 'Ngày'
    if 'Ngày' not in df.columns:
        st.error("Không tìm thấy cột 'Ngày'!")
        st.stop()
    ngay_idx = list(df.columns).index('Ngày')

    # Thêm cột "Thứ" vào bên phải cột "Ngày"
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
    df.insert(ngay_idx + 1, "Thứ", df['Ngày'].apply(convert_day))

    # Cho user xem lại đúng dữ liệu gốc + cột Thứ vừa thêm
    st.subheader("Dữ liệu giữ nguyên như file gốc:")
    st.dataframe(df, use_container_width=True, height=350)

    # Tách từng nhân viên và xuất ZIP (giữ nguyên mọi giá trị)
    if st.button("Tách file và xuất ZIP"):
        zip_buffer = BytesIO()
        with zipfile.ZipFile(zip_buffer, "w") as zip_file:
            for (ma_nv, ho_ten), group in df.groupby(['Mã NV', 'Họ tên']):
                if len(group) == 0:
                    continue
                # Dòng tổng (nếu bạn muốn, hoặc bỏ nếu không cần)
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

    st.caption("Dữ liệu giờ/phút của nhân viên sẽ giữ nguyên như file gốc (nếu chỉ có HH:mm thì giữ HH:mm, không thêm :00).")

else:
    st.info("Vui lòng upload file Excel để bắt đầu.")
