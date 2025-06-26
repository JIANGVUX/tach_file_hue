import streamlit as st
import pandas as pd
from io import BytesIO
import openpyxl
from openpyxl.styles import Alignment, Font

st.set_page_config(page_title="Tách file chấm công", layout="wide")
st.title("Tách file chấm công - Mỗi nhân viên một sheet, chỉ giữ dòng có dữ liệu từ 'Vào lần 1' trở đi")

uploaded_file = st.file_uploader("Chọn file Excel gốc (.xlsx)", type=["xlsx"])
if uploaded_file is not None:
    df = pd.read_excel(uploaded_file, sheet_name=0, header=5)

    # Xác định cột 'Vào lần 1'
    vao_lan_1_col = None
    for col in df.columns:
        if "Vào lần 1" in str(col):
            vao_lan_1_col = col
            break
    if vao_lan_1_col is None:
        st.error('Không tìm thấy cột "Vào lần 1" trong file!')
        st.stop()
    idx_vao_lan_1 = df.columns.get_loc(vao_lan_1_col)
    cols_check = df.columns[idx_vao_lan_1:]  # Từ "Vào lần 1" trở đi

    # Lọc tất cả dòng mà từ cột "Vào lần 1" trở đi có ít nhất 1 ô có dữ liệu
    def dong_co_du_lieu_tu_vao_lan_1(row):
        return row[cols_check].notna().any() and (row[cols_check] != "").any()

    df_filtered = df[df.apply(dong_co_du_lieu_tu_vao_lan_1, axis=1)]
    df_filtered = df_filtered[df_filtered['Mã NV'].notna() & df_filtered['Họ tên'].notna()]

    st.subheader("Dữ liệu đã lọc (chỉ giữ dòng có dữ liệu từ 'Vào lần 1' trở đi):")
    st.dataframe(df_filtered, use_container_width=True, height=350)

    if st.button("Tách và xuất Excel tổng (mỗi nhân viên 1 sheet)"):
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            for (ma_nv, ho_ten), group in df_filtered.groupby(['Mã NV', 'Họ tên']):
                if len(group) == 0:
                    continue
                sheet_name = f"{ma_nv}_{ho_ten}".replace(" ", "_").replace("/", "_")[:31]
                group.to_excel(writer, sheet_name=sheet_name, index=False)
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
        # Ghi lại vào buffer
        output2 = BytesIO()
        wb.save(output2)
        output2.seek(0)

        st.success("Đã tách xong! Bấm để tải file Excel tổng, mỗi nhân viên một sheet.")
        st.download_button("Tải file Excel", output2, "output_tong_hop.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    st.caption("Chỉ giữ các dòng mà từ cột 'Vào lần 1' trở đi có ít nhất 1 ô có dữ liệu. Mỗi nhân viên là 1 sheet.")

else:
    st.info("Vui lòng upload file Excel để bắt đầu.")
