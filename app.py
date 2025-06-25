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
        df = df.dropna(how='all')

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

        folder_path = os.path.join(OUTPUT_FOLDER, "files")
        os.makedirs(folder_path, exist_ok=True)

        MAX_EMPLOYEES = 100
        employee_groups = list(df_filtered.groupby("Mã NV"))
        if len(employee_groups) > MAX_EMPLOYEES:
            return f"Quá nhiều nhân viên ({len(employee_groups)}), giới hạn là {MAX_EMPLOYEES}. Hãy chia nhỏ file gốc.", 400

        for ma_nv, group in employee_groups:
            ten = group.iloc[0]["Họ tên"]
            print(f"⏳ Đang xử lý: {ma_nv} - {ten}")
            filename = f"{safe_filename(ma_nv)}_{safe_filename(ten)}.xlsx"
            file_out = os.path.join(folder_path, filename)
            group_clean = group.dropna(how='all')

            with pd.ExcelWriter(file_out, engine="xlsxwriter") as writer:
                group_clean.to_excel(writer, index=False, sheet_name="Sheet1")
                workbook  = writer.book
                worksheet = writer.sheets["Sheet1"]

                header_format = workbook.add_format({
                    'bold': True,
                    'text_wrap': True,
                    'valign': 'center',
                    'align': 'center',
                    'border': 1
                })

                for col_num, value in enumerate(group_clean.columns.values):
                    worksheet.write(0, col_num, value, header_format)
                    max_len = max(group_clean[value].astype(str).map(len).max(), len(str(value)))
                    worksheet.set_column(col_num, col_num, max_len + 2)

        zip_path = os.path.join(OUTPUT_FOLDER, "Tong_hop.zip")
        with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for f in os.listdir(folder_path):
                zipf.write(os.path.join(folder_path, f), arcname=f)

        os.remove(file_path)

        response = make_response(send_from_directory(OUTPUT_FOLDER, "Tong_hop.zip", as_attachment=True))
        response.headers['Access-Control-Allow-Origin'] = 'https://qlldhue20.weebly.com'
        response.headers['Access-Control-Expose-Headers'] = 'Content-Disposition'
        return response

    except Exception as e:
        traceback.print_exc()
        return f"Lỗi xử lý file: {str(e)}", 500

if __name__ == "__main__":
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
