import streamlit as st
import pandas as pd
from io import BytesIO
import datetime

st.set_page_config(page_title="Tách sheet nhân viên - Giữ nguyên dữ liệu", layout="wide")
st.title("Tách từng nhân viên ra mỗi sheet (giữ nguyên dữ liệu, thêm cột Thứ)")

uploaded_file = st.file_uploader("Chọn file Excel gốc (.xlsx)", type=["xlsx"])
if uploaded_file:
    # Đọc file với header, bạn chỉnh header=0 nếu dòng đầu là tiêu đề, hoặc header=5 nếu tiêu đề ở dòng 6
    df = pd.read_excel(uploaded_file, sheet_name=0)  # header tự động
    
    # Xác định vị trí cột "Ngày"
    if "Ngày" not in df.columns:
        st.error("Không tìm thấy cột 'Ngày' trong file Excel!")
        st.stop()
    ngay_idx = list(df.columns).index("Ngày")

    # Thêm cột "Thứ" vào bên trái "Ngày"
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
        except Exception:
            return ''
    df.insert(ngay_idx, "Thứ", df["Ngày"].apply(convert_day))

    # Xác định cột "Lương giờ 100%" để tính tổng các cột sau đó
    try:
        col_sum_idx = list(df.columns).index("Lương giờ 100%")
    except ValueError:
        st.error("Không tìm thấy cột 'Lương giờ 100%' trong file Excel!")
        st.stop()
    cols_sum = list(df.columns)[col_sum_idx:]  # Từ cột "Lương giờ 100%" trở đi

    # Nhóm theo Mã NV + Họ tên (bạn có thể thay đổi chỉ lấy theo Mã NV hoặc Họ tên nếu cần)
    group_cols = []
    if "Mã NV" in df.columns: group_cols.append("Mã NV")
    if "Họ tên" in df.columns: group_cols.append("Họ tên")
    if not group_cols:
        st.error("Không tìm thấy cột 'Mã NV' hoặc 'Họ tên' trong file Excel!")
        st.stop()

    df_preview = df.head(10)
    st.subheader("Dữ liệu gốc (xem trước):")
    st.dataframe(df_preview, use_container_width=True, height=300)

    # Lưu file Excel với mỗi nhân viên 1 sheet, dữ liệu giữ nguyên, chỉ thêm cột Thứ và dòng tổng
    out_buffer = BytesIO()
    with pd.ExcelWriter(out_buffer, engine='openpyxl') as writer:
        for keys, group in df.groupby(group_cols):
            sheet_name = "_".join([str(k) for k in keys])[:30]
            data_nv = group.copy()
            # Thêm dòng tổng cuối cùng
            total_row = {}
            for col in data_nv.columns:
                if col in cols_sum:
                    total_row[col] = data_nv[col].sum()
                else:
                    total_row[col] = ""
            total_row[group_cols[0]] = "Tổng"  # Hiện "Tổng" vào cột mã NV
            data_nv = pd.concat([data_nv, pd.DataFrame([total_row])], ignore_index=True)
            data_nv.to_excel(writer, index=False, sheet_name=sheet_name)
    out_buffer.seek(0)

    st.success("Đã tách xong!")
    st.download_button(
        label="Tải file kết quả (nhiều sheet, mỗi nhân viên 1 sheet)",
        data=out_buffer,
        file_name="tach_nhan_vien.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    st.caption("Dữ liệu mỗi sheet giữ nguyên, chỉ thêm cột Thứ và dòng tổng. Nếu lỗi hoặc thiếu dữ liệu, hãy kiểm tra lại cột 'Ngày', 'Lương giờ 100%', 'Mã NV', 'Họ tên'.")
else:
    st.info("Vui lòng upload file Excel để bắt đầu.")

