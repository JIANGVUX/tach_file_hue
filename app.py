from flask import Flask, request, send_from_directory
from flask_cors import CORS
import pandas as pd
import openpyxl
from openpyxl.styles import Alignment, Font
import os
import re
import shutil
from werkzeug.utils import secure_filename

app = Flask(__name__)
CORS(app)

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
    # Đổi thành tiếng Việt
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
        shutil.rmtree(OUTPUT_FOLDER)
        os.makedirs(OUTPUT_FOLDER, exist_ok=True)

        file = request.files['file']
        if not file:
            return "Vui lòng chọn file!", 400
        filename = secure_filename(file.filename)
        file_path = os.path.join(UPLOAD_FOLDER, filename)
        file.save(file_path)

        df = pd.read_excel(file_path, sheet_name=0, header=5)
        
        # Tìm vị trí cột "Ngày" (bạn đổi lại cho đúng tên cột ngày nhé)
        col_ngay = None
        for col in df.columns:
            if "ngày" in str(col).lower():
                col_ngay = col
                break
        if not col_ngay:
            return "Không tìm thấy cột ngày!", 400

        # Thêm cột Thứ vào bên trái cột ngày
        df.insert(df.columns.get_loc(col_ngay), 'Thứ', df[col_ngay].apply(weekday_vn))

        # Tìm vị trí cột "Lương giờ 100%"
        luong_col_idx = None
        for idx, col in enumerate(df.columns):
            if "lương giờ 100%" in str(col).lower():
                luong_col_idx = idx
                break
        if luong_col_idx is None:
            return 'Không tìm thấy cột "Lương giờ 100%"!', 400

        # Lọc theo yêu cầu (giữ nguyên logic bạn muốn)
        vao_lan_1_col = next((col for col in df.columns if "Vào lần 1" in str(col)), None)
        def co_du_lieu_tu_vao_lan_1(row):
            idx = df.columns.get_loc(vao_lan_1_col)
            return any([not pd.isna(cell) and str(cell).strip() != '' for cell in row[idx:]])
        df_filtered = df[df.apply(co_du_lieu_tu_vao_lan_1, axis=1)]
        df_filtered = df_filtered[df_filtered['Mã NV'].notna() & df_filtered['Họ tên'].notna()]

        # Tạo dòng tổng cộng cho các cột từ "Lương giờ 100%" trở đi
        sum_row = {col: "" for col in df_filtered.columns}
        for col in df_filtered.columns[luong_col_idx:]:
            sum_row[col] = df_filtered[col].apply(pd.to_numeric, errors='coerce').sum(skipna=True)
        sum_row[next(iter(df_filtered.columns))] = "TỔNG"
        df_filtered = pd.concat([df_filtered, pd.DataFrame([sum_row])], ignore_index=True)

        # Xuất file tổng hợp (1 sheet)
        summary_file = 'Tong_hop_loc.xlsx'
        summary_path = os.path.join(OUTPUT_FOLDER, summary_file)
        df_filtered.to_excel(summary_path, index=False)
        auto_format_excel(summary_path)
        return send_from_directory(OUTPUT_FOLDER, summary_file, as_attachment=True)
    except Exception as e:
        return f"Lỗi xử lý file: {e}", 500

if __name__ == "__main__":
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
