from flask import Flask, request, send_from_directory, make_response
from flask_cors import CORS
import pandas as pd
import os
import re
import traceback
import shutil
import xlsxwriter

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
        if os.path.exists(OUTPUT_FOLDER):
            shutil.rmtree(OUTPUT_FOLDER)
        os.makedirs(OUTPUT_FOLDER, exist_ok=True)

        if 'file' not in request.files:
            return "Vui lòng chọn file!", 400

        file = request.files['file']
        file_path = os.path.join(UPLOAD_FOLDER, safe_filename(file.filename))
        file.save(file_path)

        df = pd.read_excel(file_path, sheet_name=0, header=5).dropna(how='all')
        col_ngay = next((col for col in df.columns if "ngày" in str(col).lower()), None)
        if not col_ngay:
            return "Không tìm thấy cột ngày!", 400

        idx_ngay = df.columns.get_loc(col_ngay)
        df.insert(idx_ngay, 'Đứ', df[col_ngay].apply(weekday_vn))

        vao_lan_1_col = next((col for col in df.columns if "Vào lần 1" in str(col)), None)
        col_vl1_idx = df.columns.get_loc(vao_lan_1_col)

        mask = df.iloc[:, col_vl1_idx:].notna().any(axis=1)
        mask &= df['Mã NV'].notna() & df['Họ tên'].notna()
        df_filtered = df[mask].copy()

        summary_file = os.path.join(OUTPUT_FOLDER, "Tong_hop.xlsx")
        with pd.ExcelWriter(summary_file, engine="xlsxwriter", engine_kwargs={"options": {"nan_inf_to_errors": True}}) as writer:
            workbook = writer.book
            worksheet = workbook.add_worksheet("Tong")
            header_format = workbook.add_format({'text_wrap': True, 'bold': True})
            df_columns = df_filtered.columns
            for j, col in enumerate(df_columns):
                worksheet.write(0, j, col, header_format)

            row_idx = 1
            employee_groups = list(df_filtered.groupby("Mã NV"))
            print(f"Tổng nhân viên: {len(employee_groups)}")
            for idx, (ma_nv, group) in enumerate(employee_groups, start=1):
                ten = group.iloc[0]["Họ tên"]
                print(f"⏳ Đang xử lý {idx}/{len(employee_groups)}: {ma_nv} - {ten}")
                for i in range(len(group)):
                    for j in range(len(df_columns)):
                        val = group.iat[i, j]
                        worksheet.write(row_idx, j, "" if pd.isna(val) else val)
                    row_idx += 1
                worksheet.write(row_idx, 0, "TỔNG")
                for j in range(len(df_columns)):
                    if pd.api.types.is_numeric_dtype(group.iloc[:, j]):
                        col_sum = pd.to_numeric(group.iloc[:, j], errors='coerce').sum(skipna=True)
                        worksheet.write(row_idx, j, col_sum)
                row_idx += 2

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
