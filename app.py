from flask import Flask, request, send_from_directory, make_response
from flask_cors import CORS
import pandas as pd
import os
import re
import traceback
import shutil
from datetime import datetime

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
        return thu[pd.to_datetime(dt, dayfirst=True).weekday()]
    except:
        return ""

@app.route("/", methods=["GET"])
def home():
    return "Server đã chạy OK!"

@app.route("/upload", methods=["POST"])
def upload():
    try:
        shutil.rmtree(OUTPUT_FOLDER, ignore_errors=True)
        os.makedirs(OUTPUT_FOLDER, exist_ok=True)

        if 'file' not in request.files:
            return "Vui lòng chọn file!", 400

        file = request.files['file']
        file_path = os.path.join(UPLOAD_FOLDER, safe_filename(file.filename))
        file.save(file_path)

        df = pd.read_excel(file_path, sheet_name=0, header=5)
        df = df.dropna(how='all')

        col_ngay = next((col for col in df.columns if "ngày" in str(col).lower()), None)
        if not col_ngay:
            return "Không tìm thấy cột ngày!", 400

        idx_ngay = df.columns.get_loc(col_ngay)
        df.insert(idx_ngay, 'Thứ', df[col_ngay].apply(weekday_vn))

        col_ma_nv = 'Mã NV'
        col_ho_ten = 'Họ tên'

        if col_ma_nv not in df.columns or col_ho_ten not in df.columns:
            return "Không tìm thấy cột Mã NV hoặc Họ tên", 400

        df = df[df[col_ma_nv].notna() & df[col_ho_ten].notna()].copy()

        output_file = os.path.join(OUTPUT_FOLDER, "Tong_hop.xlsx")

        with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
            workbook = writer.book
            worksheet = workbook.add_worksheet("Tổng hợp")
            writer.sheets["Tổng hợp"] = worksheet

            bold_format = workbook.add_format({'bold': True})
            wrap_format = workbook.add_format({'text_wrap': True})

            row_idx = 0
            for ma_nv, group in df.groupby(col_ma_nv):
                ho_ten = group.iloc[0][col_ho_ten]
                print(f"⏳ Đang xử lý: {ma_nv} - {ho_ten}")

                group = group.dropna(how='all')
                group.reset_index(drop=True, inplace=True)

                # Ghi header
                if row_idx == 0:
                    for col_idx, col_name in enumerate(group.columns):
                        worksheet.write(row_idx, col_idx, col_name, bold_format)
                    row_idx += 1

                for i in range(len(group)):
                    for j in range(len(group.columns)):
                        worksheet.write(row_idx, j, group.iat[i, j])
                    row_idx += 1

                # Ghi dòng TỔNG sau mỗi nhóm
                worksheet.write(row_idx, 0, "TỔNG", bold_format)
                for j in range(len(group.columns)):
                    col_data = group.iloc[:, j]
                    if pd.api.types.is_numeric_dtype(col_data):
                        worksheet.write(row_idx, j, col_data.sum(), bold_format)
                row_idx += 2  # chỉnh giá trị 2 dể cách

        os.remove(file_path)

        response = make_response(send_from_directory(OUTPUT_FOLDER, "Tong_hop.xlsx", as_attachment=True))
        response.headers['Access-Control-Allow-Origin'] = 'https://qlldhue20.weebly.com'
        response.headers['Access-Control-Expose-Headers'] = 'Content-Disposition'
        return response

    except Exception as e:
        traceback.print_exc()
        return f"Lỗi xử lý file: {str(e)}", 500

if __name__ == "__main__":
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)