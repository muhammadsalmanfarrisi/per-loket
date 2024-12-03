from flask import Flask, render_template, request, redirect, url_for
import pandas as pd
import os

app = Flask(__name__)

# Folder untuk menyimpan file yang diunggah
UPLOAD_FOLDER = 'uploads'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# Memastikan folder uploads ada
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

# Fungsi untuk memproses file Excel
def process_excel(file_path):
    # Membaca data dari sheet 'DATA CONTROL' pada file Excel
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
    except KeyError as e:
        raise ValueError(f"Kolom tidak ditemukan: {e}")

    filtered_data = data[data.iloc[:, idx_status_pembayaran].str.lower() == "unpaid"]
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

    grand_total = summary_df.iloc[:, 1:].sum(axis=0)
    grand_total["Kantor"] = "Total"
    summary_df = pd.concat([summary_df, pd.DataFrame([grand_total])], ignore_index=True)

    return summary_df

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return redirect(request.url)
    
    file = request.files['file']
    if file.filename == '':
        return redirect(request.url)
    
    if file:
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
        file.save(file_path)
        
        try:
            summary_df = process_excel(file_path)
            return render_template('index.html', tables=[summary_df.to_html(classes='data', header=True)], title="Summary")
        except Exception as e:
            return f"An error occurred: {e}"

if __name__ == '__main__':
    app.run(host="0.0.0.0", port=5000)
