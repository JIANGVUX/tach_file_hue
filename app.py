from flask import Flask, request, send_from_directory, make_response
from flask_cors import CORS
import pandas as pd
import os
import re
import traceback
import zipfile
import shutil

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
        # Xoá dữ liệu cũ
        if os.path.exists(OUTPUT_FOLDER):
            for f in os.listdir(OUTPUT_FOLDER):
                path = os.path.join(OUTPUT_FOLDER, f)
                if os.path.isfile(path):
                    os.remove(path)
                elif os.path.isdir(path):
                    shutil.rmtree(path)

        if 'file' not in request.files:
            return "Vui lòng chọn file!", 400

        file = request.files['file']
        file_path = os.path.join(UPLOAD_FOLDER, safe_filename(file.filename))
        file.save(file_path)

        df = pd.read_excel(file_path, sheet_name=0, header=5)
        df = df.dropna(how='all')  # Xoá dòng rỗng hoàn toàn

        col_ngay = next((col for col in df.columns if "ngày" in str(col).lower()), None)
        if not col_ngay:
            return "Không tìm thấy cột ngày!", 400

        idx_ngay = df.columns.get_loc(col_ngay)
        df.insert(idx_ngay, 'Thứ', df[col_ngay].apply(weekday_vn))

        vao_lan_1_col = next((col for col in df.columns if "Vào lần 1" in str(col)), None)
        col_vl1_idx = df.columns.get_loc(vao_lan_1_col)

        mask = df.iloc[:, col_vl1_idx:].notna().any(axis=1)
        mask &= df['Mã NV'].notna() & df['Họ tên'].notna()
        df_filtered = df[mask].copy()

        output_file = os.path.join(OUTPUT_FOLDER, "Tong_hop.xlsx")
        with pd.ExcelWriter(output_file, engine="xlsxwriter") as writer:
            workbook = writer.book
            worksheet = workbook.add_worksheet("Tổng hợp")
            writer.sheets["Tổng hợp"] = worksheet

            header_format = workbook.add_format({'bold': True, 'text_wrap': True, 'valign': 'center'})

            columns = df_filtered.columns.tolist()
            for col_idx, col_name in enumerate(columns):
                worksheet.write(0, col_idx, col_name, header_format)

            row_idx = 1
            for ma_nv, group in df_filtered.groupby("Mã NV"):
                ten = group.iloc[0]["Họ tên"]
                print(f"⏳ Đang xử lý: {ma_nv} - {ten}")

                for i in range(len(group)):
                    for j in range(len(columns)):
                        value = group.iat[i, j]
                        if pd.isna(value) or value in [float("inf"), float("-inf")]:
                            value = ""
                        worksheet.write(row_idx, j, value)
                    row_idx += 1

                # Tính tổng cho cột số học
                worksheet.write(row_idx, 0, "Tổng")
                for j in range(len(columns)):
                    col_data = pd.to_numeric(group.iloc[:, j], errors='coerce')
                    if col_data.notna().any():
                        worksheet.write(row_idx, j, col_data.sum(skipna=True))
                row_idx += 1

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