import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Tách sheet nhân viên - Giữ nguyên dữ liệu", layout="wide")
st.title("Tách từng nhân viên ra mỗi sheet (giữ nguyên dữ liệu, thêm cột Thứ)")

uploaded_file = st.file_uploader("Chọn file Excel gốc (.xlsx)", type=["xlsx"])
if uploaded_file:
    df = pd.read_excel(uploaded_file, sheet_name=0)

    st.write("**Các cột thực tế của file:**", df.columns.tolist())

    # Tìm cột có tên chứa 'ngày' (không phân biệt hoa thường, bỏ khoảng trắng đầu/cuối)
    ngay_col = next((col for col in df.columns if 'ngày' in str(col).lower().strip()), None)
    if ngay_col is None:
        st.error("Không tìm thấy cột nào chứa 'ngày'. Tên cột thực tế là: " + ", ".join(df.columns))
        st.stop()
    ngay_idx = list(df.columns).index(ngay_col)

    # Thêm cột "Thứ" vào bên trái cột ngày tìm được
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
    df.insert(ngay_idx, "Thứ", df[ngay_col].apply(convert_day))

    # Tìm cột chứa 'lương giờ 100%'
    col_sum_idx = next((i for i, col in enumerate(df.columns) if 'lương giờ 100%' in str(col).lower()), None)
    if col_sum_idx is None:
        st.error("Không tìm thấy cột nào chứa 'lương giờ 100%'. Tên cột thực tế là: " + ", ".join(df.columns))
        st.stop()
    cols_sum = list(df.columns)[col_sum_idx:]

    # Tìm cột mã NV/họ tên gần đúng
    ma_nv_col = next((col for col in df.columns if 'mã' in str(col).lower() and 'nv' in str(col).lower()), None)
    ho_ten_col = next((col for col in df.columns if 'họ' in str(col).lower() and 'tên' in str(col).lower()), None)
    group_cols = [col for col in [ma_nv_col, ho_ten_col] if col]
    if not group_cols:
        st.error("Không tìm thấy cột nào là 'Mã NV' hoặc 'Họ tên'. Tên cột thực tế là: " + ", ".join(df.columns))
        st.stop()

    st.subheader("Dữ liệu gốc (xem trước):")
    st.dataframe(df.head(10), use_container_width=True)

    # Xuất file
    out_buffer = BytesIO()
    with pd.ExcelWriter(out_buffer, engine='openpyxl') as writer:
        for keys, group in df.groupby(group_cols):
            sheet_name = "_".join([str(k) for k in keys])[:30]
            data_nv = group.copy()
            total_row = {}
            for col in data_nv.columns:
                if col in cols_sum:
                    total_row[col] = data_nv[col].sum()
                else:
                    total_row[col] = ""
            total_row[group_cols[0]] = "Tổng"
            data_nv = pd.concat([data_nv, pd.DataFrame([total_row])], ignore_index=True)
            data_nv.to_excel(writer, index=False, sheet_name=sheet_name)
    out_buffer.seek(0)

    st.success("Đã tách xong!")
    st.download_button(
        label="Tải file kết quả (mỗi nhân viên 1 sheet)",
        data=out_buffer,
        file_name="tach_nhan_vien.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    st.caption("Dữ liệu mỗi sheet giữ nguyên, chỉ thêm cột Thứ và dòng tổng. Nếu vẫn lỗi tên cột, hãy kiểm tra lại file gốc hoặc gửi mình sample file.")
else:
    st.info("Vui lòng upload file Excel để bắt đầu.")
