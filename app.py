from flask import Flask, render_template, request, send_file
import pandas as pd
import os
from datetime import datetime
from werkzeug.utils import secure_filename


app = Flask(__name__)

# Folder untuk menyimpan file yang diunggah
UPLOAD_FOLDER = 'uploads'
RESULT_FOLDER = 'results'

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['RESULT_FOLDER'] = RESULT_FOLDER

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(RESULT_FOLDER, exist_ok=True)

# Fungsi untuk memproses file Excel
def process_excel(file_path):
    try:
        raw_data = pd.read_excel(file_path, sheet_name='DATA CONTROL', header=None).values.tolist()
    except ValueError:
        raise ValueError("Sheet 'DATA CONTROL' tidak ditemukan dalam file Excel.")

    header_row_index = -1
    for i, row in enumerate(raw_data):
        if any("nomor id jaminan" in str(cell).lower() for cell in row):
            header_row_index = i
            break

    if header_row_index == -1:
        raise ValueError("Header yang mengandung 'Nomor ID Jaminan' tidak ditemukan.")

    headers = raw_data[header_row_index]
    data = pd.DataFrame(raw_data[header_row_index + 1:], columns=headers)
    headers_lower = data.columns.str.lower()

    try:
        idx_kantor = headers_lower.get_loc("kantor")
        idx_status_pembayaran = headers_lower.get_loc("status pembayaran")
        idx_status_verifikasi = headers_lower.get_loc("status verifikasi")
        idx_gl_status = headers_lower.get_loc("gl status")
    except KeyError as e:
        raise ValueError(f"Kolom tidak ditemukan: {e}")

    filtered_data = data[
        (data.iloc[:, idx_gl_status].str.lower() == "active") &
        (data.iloc[:, idx_status_pembayaran].str.lower() == "unpaid")
    ]
    result = {}

    for _, row in filtered_data.iterrows():
        kantor = row[idx_kantor]
        status_verifikasi = str(row[idx_status_verifikasi]).lower()

        if kantor not in result:
            result[kantor] = {
                "Done": 0,
                "Revision": 0,
                "New": 0,
                "Waiting First Layer Verification": 0,
                "Lain-lain": 0,
                "Total": 0
            }

        if "done" in status_verifikasi or "resend" in status_verifikasi:
            result[kantor]["Done"] += 1
        elif status_verifikasi == "revision":
            result[kantor]["Revision"] += 1
        elif status_verifikasi in ["new", "draft"]:
            result[kantor]["New"] += 1
        elif status_verifikasi == "waiting first layer verification":
            result[kantor]["Waiting First Layer Verification"] += 1
        else:
            result[kantor]["Lain-lain"] += 1

        result[kantor]["Total"] += 1

    summary_data = [
        [kantor, *counts.values()] for kantor, counts in result.items()
    ]
    summary_df = pd.DataFrame(summary_data, columns=["Kantor", "Done", "Revision", "New",
                                                     "Waiting First Layer Verification", "Lain-lain", "Total"])
    summary_df = summary_df.sort_values(by="Kantor", ascending=True)

    grand_total = summary_df.iloc[:, 1:].sum(axis=0)
    grand_total["Kantor"] = "Total"
    summary_df = pd.concat([summary_df, pd.DataFrame([grand_total])], ignore_index=True)

    return summary_df

def process_and_save(file_path, result_file):
    summary_df = process_excel(file_path)
    
    
    with pd.ExcelWriter(result_file, engine="openpyxl") as writer:
        summary_df.to_excel(writer, index=False, sheet_name="Summary")
        

import shutil  # Tambahkan modul shutil untuk menghapus folder

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        file = request.files.get("file")
        if not file:
            return render_template("index.html", error="Harap unggah file.")

        file_path = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
        file.save(file_path)

        try:
            result_file = os.path.join(app.config['RESULT_FOLDER'], "summaryloket.xlsx")
            process_and_save(file_path, result_file)
            os.remove(file_path)
            return send_file(result_file, as_attachment=True)
        except Exception as e:
            if os.path.exists(file_path):
                os.remove(file_path)
            return render_template("index.html", error=str(e))
    
    return render_template("index.html")


if __name__ == '__main__':
    app.run(host="0.0.0.0", port=5550, debug=True)
