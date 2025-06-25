from flask import Flask, request, send_from_directory, make_response
from flask_cors import CORS
import pandas as pd
import openpyxl
from openpyxl.styles import Alignment, Font
import os
import re
import traceback

app = Flask(__name__)
CORS(app, resources={r"/*": {"origins": ["https://qlldhue20.weebly.com"]}}, supports_credentials=True)

UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'output'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

def safe_filename(s):
    s = str(s)
    s = re.sub(r'[\\/*?:"<>|]', '_', s)
    s = re.sub(r'\s+', '_', s)
    return s.strip('_')

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
        summary_file = 'Tong_hop_loc.xlsx'
        summary_path = os.path.join(OUTPUT_FOLDER, summary_file)
        if os.path.exists(summary_path):
            os.remove(summary_path)

        if 'file' not in request.files:
            return "Vui lòng chọn file!", 400

        file = request.files['file']
        filename = safe_filename(file.filename)
        file_path = os.path.join(UPLOAD_FOLDER, filename)
        file.save(file_path)

        df = pd.read_excel(file_path, sheet_name=0, header=5)
        col_ngay = next((col for col in df.columns if "ngày" in str(col).lower()), None)
        if not col_ngay:
            return "Không tìm thấy cột ngày!", 400

        idx_ngay = df.columns.get_loc(col_ngay)
        df.insert(idx_ngay, 'Thứ', df[col_ngay].apply(weekday_vn))

        luong_col_idx = next((i for i, col in enumerate(df.columns) if "lương giờ 100%" in str(col).lower()), None)
        if luong_col_idx is None:
            return 'Không tìm thấy cột "Lương giờ 100%"!', 400

        vao_lan_1_col = next((col for col in df.columns if "Vào lần 1" in str(col)), None)
        if vao_lan_1_col is None:
            return 'Không tìm thấy cột "Vào lần 1"!', 400

        col_vl1_idx = df.columns.get_loc(vao_lan_1_col)
        mask = df.iloc[:, col_vl1_idx:].notna().any(axis=1)
        mask &= df['Mã NV'].notna() & df['Họ tên'].notna()
        df_filtered = df[mask].copy()

        sum_row = {}
        for i, col in enumerate(df_filtered.columns):
            if i < luong_col_idx:
                sum_row[col] = "" if i != 0 else "TỔNG"
            else:
                sum_row[col] = pd.to_numeric(df_filtered[col], errors='coerce').sum(skipna=True)

        df_filtered.loc[len(df_filtered)] = sum_row
        df_filtered.to_excel(summary_path, index=False)
        # auto_format_excel(summary_path)  # Tạm thời bỏ để giảm tải RAM
        os.remove(file_path)

        response = make_response(send_from_directory(OUTPUT_FOLDER, summary_file, as_attachment=True))
        response.headers['Access-Control-Allow-Origin'] = 'https://qlldhue20.weebly.com'
        response.headers['Access-Control-Expose-Headers'] = 'Content-Disposition'
        return response

    except Exception as e:
        traceback.print_exc()
        return f"Lỗi xử lý file: {str(e)}", 500

if __name__ == "__main__":
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)