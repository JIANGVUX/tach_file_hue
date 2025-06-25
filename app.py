from flask import Flask, render_template, request, send_from_directory, redirect, url_for, flash
from flask_cors import CORS     # <--- DÒNG NÀY!
import pandas as pd
import openpyxl
from openpyxl.styles import Alignment, Font
import os
import re
import shutil
from werkzeug.utils import secure_filename

app = Flask(__name__)
CORS(app)                       # <--- DÒNG NÀY!
app.secret_key = 'supersecret'
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
        vao_lan_1_col = next((col for col in df.columns if "Vào lần 1" in str(col)), None)
        if vao_lan_1_col is None:
            return 'Không tìm thấy cột "Vào lần 1" trong file!', 400
        def co_du_lieu_tu_vao_lan_1(row):
            idx = df.columns.get_loc(vao_lan_1_col)
            return any(pd.notna(cell) and str(cell).strip() != '' for cell in row[idx:])
        df_filtered = df[df.apply(co_du_lieu_tu_vao_lan_1, axis=1)]
        df_filtered = df_filtered[df_filtered['Mã NV'].notna() & df_filtered['Họ tên'].notna()]
        if df_filtered.empty:
            return "Không có dữ liệu nhân viên hợp lệ sau khi lọc!", 400

        summary_file = 'Tong_hop_loc.xlsx'
        df_filtered.to_excel(os.path.join(OUTPUT_FOLDER, summary_file), index=False)
        auto_format_excel(os.path.join(OUTPUT_FOLDER, summary_file))

        # Trả file tổng hợp về cho người dùng tải luôn
        return send_from_directory(OUTPUT_FOLDER, summary_file, as_attachment=True)
    except Exception as e:
        return f"Lỗi xử lý file: {e}", 500


@app.route("/download/<filename>")
def download_file(filename):
    return send_from_directory(OUTPUT_FOLDER, filename, as_attachment=True)

if __name__ == "__main__":
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
