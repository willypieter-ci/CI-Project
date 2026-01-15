# app.py
import os
import shutil
import time
from flask import Flask, request, send_file, render_template_string

# Folder output tetap (bukan temp)
OUTPUT_FOLDER = "outputs"
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 10 * 1024 * 1024  # Max 10 MB

@app.route('/')
def home():
    return render_template_string('''
    <!DOCTYPE html>
    <html>
    <head>
        <title>Absensi HRD - PT Canang Indah</title>
        <style>
            body { font-family: Arial, sans-serif; background: #f9f9f9; text-align: center; padding: 50px; }
            .container { max-width: 600px; margin: auto; background: white; padding: 30px; border-radius: 10px; box-shadow: 0 0 10px rgba(0,0,0,0.1); }
            h2 { color: #2c3e50; }
            .upload-btn { background: #3498db; color: white; padding: 12px 24px; border: none; border-radius: 5px; cursor: pointer; font-size: 16px; margin-top: 10px; }
            .upload-btn:hover { background: #2980b9; }
            input[type="file"] { margin-top: 10px; }
        </style>
    </head>
    <body>
        <div class="container">
            <h2>📁 Upload File Absensi (CSV)</h2>
            <p>Format: file dari mesin fingerprint (ekspor sebagai CSV)</p>
            <form action="/upload" method="post" enctype="multipart/form-data">
                <input type="file" name="file" accept=".csv" required><br>
                <button type="submit" class="upload-btn">Proses & Download Hasil</button>
            </form>
        </div>
    </body>
    </html>
    ''')

@app.route('/upload', methods=['POST'])
def upload():
    file = request.files.get('file')
    if not file or not file.filename.endswith('.csv'):
        return "<h3>❌ Error: Harap upload file CSV!</h3><a href='/'>Kembali</a>", 400

    # Bersihkan folder output lama
    for f in os.listdir(OUTPUT_FOLDER):
        fp = os.path.join(OUTPUT_FOLDER, f)
        try:
            if os.path.isfile(fp):
                os.remove(fp)
            elif os.path.isdir(fp):
                shutil.rmtree(fp)
        except Exception as e:
            print(f"Gagal hapus {fp}: {e}")

    # Simpan file upload
    csv_path = os.path.join(OUTPUT_FOLDER, "absensi_mentah.csv")
    file.save(csv_path)

    try:
        from processor import process_attendance
        hasil_folder = process_attendance(csv_path, os.path.join(OUTPUT_FOLDER, "hasil"))

        # Buat ZIP
        zip_path = os.path.join(OUTPUT_FOLDER, "hasil_absensi.zip")
        if os.path.exists(zip_path):
            os.remove(zip_path)
        shutil.make_archive(zip_path.replace('.zip', ''), 'zip', hasil_folder)

        # Kirim file
        return send_file(zip_path, as_attachment=True, download_name="hasil_absensi.zip")

    except Exception as e:
        return f"<h3>❌ Gagal memproses: {str(e)}</h3><a href='/'>Coba lagi</a>", 500

if __name__ == '__main__':
    print("🚀 Server berjalan di http://localhost:5000")
    app.run(host='127.0.0.1', port=5000, debug=False)