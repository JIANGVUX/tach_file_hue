import pandas as pd

# Đường dẫn file gốc
file_path = '1.xlsx'
# Đọc file, header là dòng số 6 (index 5)
df = pd.read_excel(file_path, sheet_name=0, header=5)

# Xác định vị trí cột "Vào lần 1"
vao_lan_1_col = None
for col in df.columns:
    if "Vào lần 1" in str(col):
        vao_lan_1_col = col
        break

if vao_lan_1_col is None:
    raise Exception('Không tìm thấy cột "Vào lần 1" trong file!')

# Lọc chỉ lấy dòng mà từ "Vào lần 1" trở đi có dữ liệu (tức là không phải dòng trống)
def co_du_lieu_tu_vao_lan_1(row):
    idx = df.columns.get_loc(vao_lan_1_col)
    # Nếu bất kỳ ô nào từ "Vào lần 1" trở đi có dữ liệu, giữ lại
    return any([not pd.isna(cell) and str(cell).strip() != '' for cell in row[idx:]])

df_filtered = df[df.apply(co_du_lieu_tu_vao_lan_1, axis=1)]

# Chỉ lấy các cột chính (tránh lấy dòng trống tiêu đề phụ ở đầu file)
df_filtered = df_filtered[df_filtered['Mã NV'].notna() & df_filtered['Họ tên'].notna()]

# Tạo file cho từng nhân viên
for (ma_nv, ho_ten), group in df_filtered.groupby(['Mã NV', 'Họ tên']):
    # Bỏ qua nếu nhóm này không có dòng dữ liệu
    if len(group) == 0:
        continue
    # Tạo tên file đẹp, an toàn
    file_name = f'{ma_nv}_{ho_ten}'.replace(" ", "_").replace("/", "_")
    group.to_excel(f'{file_name}.xlsx', index=False)
    print(f'Đã xuất: {file_name}.xlsx')

print("Xuất hoàn tất!")
