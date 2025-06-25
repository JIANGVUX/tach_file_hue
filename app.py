from flask import Flask, request, send_from_directory
from flask_cors import CORS
import pandas as pd
import openpyxl
from openpyxl.styles import Alignment, Font
import os
import re

app = Flask(__name__)
CORS(app,origins=["https://qlldhue20.weebly.com", "http://qlldhue20.weebly.com"])

UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'output'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

def safe_filename(s):
    s = str(s)
    s = re.sub(r'[\\/*?:"<>|]', '_', s)
    s = re.sub(r'\s+', '_', s)
    return s.strip('_')

def auto_format_excel(file_path):
    wb = openpyxl.load_workbook(file_path)
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
    wb.save(file_path)

def weekday_vn(dt):
    if pd.isna(dt): return ""
    thu = ["Thứ 2", "Thứ 3", "Thứ 4", "Thứ 5", "Thứ 6", "Thứ 7", "Chủ nhật"]
    try:
        return thu[pd.to_datetime(dt).weekday()]
    except:
        return ""

@app.route("/", methods=["GET"])
def home():
    return "Server đã chạy OK!"

@app.route("/upload", methods=["POST"])
def upload():
    try:
        # Chỉ xóa file tổng hợp cũ nếu tồn tại, không xóa cả thư mục
        summary_file = 'Tong_hop_loc.xlsx'
        summary_path = os.path.join(OUTPUT_FOLDER, summary_file)
        if os.path.exists(summary_path):
            os.remove(summary_path)

        file = request.files['file']
        if not file:
            return "Vui lòng chọn file!", 400
        filename = safe_filename(file.filename)
        file_path = os.path.join(UPLOAD_FOLDER, filename)
        file.save(file_path)

        # Đọc sheet 0, header dòng 6
        df = pd.read_excel(file_path, sheet_name=0, header=5)

        # Xác định cột ngày
        col_ngay = next((col for col in df.columns if "ngày" in str(col).lower()), None)
        if not col_ngay:
            return "Không tìm thấy cột ngày!", 400

        # Thêm cột thứ vào bên trái cột ngày
        idx_ngay = df.columns.get_loc(col_ngay)
        df.insert(idx_ngay, 'Thứ', df[col_ngay].apply(weekday_vn))

        # Xác định cột Lương giờ 100%
        luong_col_idx = next((i for i, col in enumerate(df.columns) if "lương giờ 100%" in str(col).lower()), None)
        if luong_col_idx is None:
            return 'Không tìm thấy cột "Lương giờ 100%"!', 400

        # Xác định cột Vào lần 1
        vao_lan_1_col = next((col for col in df.columns if "Vào lần 1" in str(col)), None)
        if vao_lan_1_col is None:
            return 'Không tìm thấy cột "Vào lần 1"!', 400

        # Lọc: chỉ giữ dòng có dữ liệu từ cột "Vào lần 1" trở đi
        col_vl1_idx = df.columns.get_loc(vao_lan_1_col)
        mask = df.iloc[:, col_vl1_idx:].notna().any(axis=1)
        mask &= df['Mã NV'].notna() & df['Họ tên'].notna()
        df_filtered = df[mask].copy()

        # Tạo dòng tổng cộng cho các cột số từ Lương giờ 100% trở đi
        sum_row = {}
        for i, col in enumerate(df_filtered.columns):
            if i < luong_col_idx:
                sum_row[col] = "" if i != 0 else "TỔNG"
            else:
                sum_row[col] = pd.to_numeric(df_filtered[col], errors='coerce').sum(skipna=True)
        df_filtered.loc[len(df_filtered)] = sum_row

        # Ghi ra file, dọn dẹp
        df_filtered.to_excel(summary_path, index=False)
        auto_format_excel(summary_path)
        os.remove(file_path)
        return send_from_directory(OUTPUT_FOLDER, summary_file, as_attachment=True)
    except Exception as e:
        return f"Lỗi xử lý file: {e}", 500

if __name__ == "__main__":
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
