from flask import Flask, request, send_file
from flask_cors import CORS
import pandas as pd, os, re, shutil, traceback
from io import BytesIO

app = Flask(__name__)
CORS(app, resources={r"/*": {"origins": ["https://qlldhue20.weebly.com"]}}, supports_credentials=True)

def safe_filename(s):
    return re.sub(r'[\\/*?:"<>|\s]+', '_', str(s)).strip('_')

def weekday_vn(dt):
    try:
        return ["Thứ 2","Thứ 3","Thứ 4","Thứ 5","Thứ 6","Thứ 7","Chủ nhật"][pd.to_datetime(dt, dayfirst=True).weekday()]
    except:
        return ""

@app.route("/upload", methods=["POST"])
def upload():
    try:
        f = request.files.get("file")
        if not f:
            return "Vui lòng chọn file!", 400

        df = pd.read_excel(f, header=5).dropna(how="all")
        date_col = next((c for c in df.columns if "ngày" in c.lower()), None)
        if not date_col:
            return "Không tìm thấy cột 'Ngày'!", 400

        df.insert(df.columns.get_loc(date_col), "Thứ", df[date_col].apply(weekday_vn))
        col_vl1 = next((c for c in df.columns if "vào lần 1" in c.lower()), None)
        if not col_vl1:
            return "Không tìm thấy cột 'Vào lần 1'!", 400

        df = df[df[col_vl1].notna() & df["Mã NV"].notna() & df["Họ tên"].notna()]
        groups = df.groupby("Mã NV")
        total_groups = len(groups)

        output = BytesIO()
        writer = pd.ExcelWriter(output, engine="xlsxwriter", options={'nan_inf_to_errors': True})
        workbook = writer.book
        worksheet = workbook.add_worksheet("TongHop")

        headers = df.columns.tolist()
        for j, h in enumerate(headers):
            worksheet.write(0, j, h)

        row_pos = 1
        for idx, (code, grp) in enumerate(groups, start=1):
            name = grp["Họ tên"].iat[0]
            print(f"⏳ Đang xử lý {idx}/{total_groups}: {code} - {name}")
            grp = grp.reset_index(drop=True)

            numeric_cols = grp.select_dtypes(include='number').columns

            for i in range(len(grp)):
                for j, col in enumerate(headers):
                    val = grp.iat[i, j]
                    if pd.notna(val):
                        worksheet.write(row_pos, j, val)
                row_pos += 1

            worksheet.write(row_pos, 0, "Tổng")
            for col in numeric_cols:
                j = headers.index(col)
                col_letter = chr(65 + j)
                start_row = row_pos - len(grp) + 1
                end_row = row_pos
                worksheet.write_formula(row_pos, j, f"=SUM({col_letter}{start_row+1}:{col_letter}{end_row})")
            row_pos += 1

        writer.close()
        output.seek(0)

        return send_file(output, as_attachment=True, download_name="Tong_hop.xlsx",
                         mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    except Exception as e:
        traceback.print_exc()
        return f"Lỗi xử lý: {e}", 500

@app.route("/", methods=["GET"])
def home():
    return "✅ Ready!"
