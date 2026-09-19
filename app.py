import os
import time
import pythoncom
import win32com.client
from flask import Flask, request, render_template_string, flash, redirect, send_from_directory, url_for
from werkzeug.utils import secure_filename
import win32api
import win32print

app = Flask(__name__)
app.secret_key = "kunci_rahasia_aplikasi"
UPLOAD_FOLDER = 'uploads'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# Memastikan folder uploads tersedia
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

PRINTER_NAME = win32print.GetDefaultPrinter()

def convert_to_pdf(input_path, output_path, ext):
    pythoncom.CoInitialize()
    try:
        if ext in ['doc', 'docx', 'rtf']:
            word = win32com.client.Dispatch('Word.Application')
            word.Visible = False
            doc = word.Documents.Open(os.path.abspath(input_path))
            doc.SaveAs(os.path.abspath(output_path), FileFormat=17)
            doc.Close()
            word.Quit()
        elif ext in ['xls', 'xlsx']:
            excel = win32com.client.Dispatch('Excel.Application')
            excel.Visible = False
            wb = excel.Workbooks.Open(os.path.abspath(input_path))
            wb.ExportAsFixedFormat(0, os.path.abspath(output_path))
            wb.Close()
            excel.Quit()
    finally:
        pythoncom.CoUninitialize()

# CSS untuk kedua halaman
COMMON_STYLE = '''
    <style>
        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; text-align: center; margin-top: 50px; background-color: #f4f7f6; }
        .container { background: white; padding: 40px; border-radius: 12px; box-shadow: 0 4px 15px rgba(0,0,0,0.05); display: inline-block; max-width: 600px; width: 100%; }
        h2 { color: #333; margin-top: 10px; }
        input[type=file], select { margin: 15px 0; display: block; width: 100%; padding: 10px; border-radius: 5px; border: 1px solid #ccc; box-sizing: border-box; }
        button { background-color: #0078D7; color: white; padding: 12px 25px; border: none; border-radius: 6px; cursor: pointer; font-size: 16px; width: 100%; font-weight: bold; margin-top: 15px; }
        button:hover { background-color: #005A9E; }
        button.cancel { background-color: #dc3545; margin-top: 10px; }
        button.cancel:hover { background-color: #c82333; }
        .alert { color: #155724; background-color: #d4edda; border-color: #c3e6cb; padding: 10px; border-radius: 5px; margin-bottom: 20px; text-align: left; }
        .alert-error { color: #721c24; background-color: #f8d7da; border-color: #f5c6cb; }
        .form-group { text-align: left; margin-bottom: 15px; }
        .form-group label { font-weight: bold; display: block; margin-bottom: 5px; color: #555; }
        .loader-overlay { position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: rgba(255,255,255,0.8); display: none; justify-content: center; align-items: center; z-index: 9999; flex-direction: column; }
        .spinner { border: 8px solid #f3f3f3; border-top: 8px solid #0078D7; border-radius: 50%; width: 50px; height: 50px; animation: spin 1s linear infinite; }
        @keyframes spin { 0% { transform: rotate(0deg); } 100% { transform: rotate(360deg); } }
    </style>
    <script>
        function showLoader() {
            document.getElementById('loader').style.display = 'flex';
        }
    </script>
'''

HTML_TEMPLATE = f'''
<!DOCTYPE html>
<html lang="id">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Printer Sharing - Canon G1010</title>
    <link rel="icon" type="image/jpeg" href="/bagiprinter/static/logo.jpg">
    {COMMON_STYLE}
</head>
<body>
    <div id="loader" class="loader-overlay">
        <div class="spinner"></div>
        <h3 style="margin-top: 20px; color: #333;">Memproses Dokumen...</h3>
        <p>Mohon tunggu sebentar.</p>
    </div>
    <div class="container">
        <img src="/bagiprinter/static/logo.jpg" alt="Logo BagiPrinter" style="max-width: 120px; border-radius: 10px;">
        <h2>🖨️ Cetak Dokumen Jaringan</h2>
        <p style="color: #666;">Printer Target: <b>{{{{ printer_name }}}}</b></p>
        
        {{% with messages = get_flashed_messages(with_categories=true) %}}
          {{% if messages %}}
            {{% for category, message in messages %}}
                <div class="alert {{% if category == 'error' %}}alert-error{{% endif %}}">
                    {{{{ message }}}}
                </div>
            {{% endfor %}}
          {{% endif %}}
        {{% endwith %}}
        
        <form action="/bagiprinter/upload" method="post" enctype="multipart/form-data" onsubmit="showLoader()">
            <input type="file" name="file" accept=".pdf,.rtf,.doc,.docx,.xls,.xlsx,.jpg,.jpeg,.png" required>
            <button type="submit">Unggah & Preview</button>
        </form>
    </div>
</body>
</html>
'''

PREVIEW_TEMPLATE = f'''
<!DOCTYPE html>
<html lang="id">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Preview & Setting - BagiPrinter</title>
    <link rel="icon" type="image/jpeg" href="/bagiprinter/static/logo.jpg">
    {COMMON_STYLE}
</head>
<body>
    <div id="loader" class="loader-overlay">
        <div class="spinner"></div>
        <h3 style="margin-top: 20px; color: #333;">Mengirim ke Printer...</h3>
    </div>
    <div class="container" style="max-width: 800px;">
        <h2>Preview Dokumen</h2>
        <p style="color: #666;">File: <b>{{{{ filename }}}}</b></p>
        
        {{% if ext in ['jpg', 'jpeg', 'png'] %}}
            <img src="/bagiprinter/uploads/{{{{ filename }}}}" style="max-width: 100%; max-height: 400px; border: 1px solid #ddd; border-radius: 5px; margin-bottom: 20px;">
        {{% elif ext == 'pdf' %}}
            <iframe src="/bagiprinter/uploads/{{{{ filename }}}}" width="100%" height="400px" style="border: 1px solid #ddd; border-radius: 5px; margin-bottom: 20px;"></iframe>
        {{% else %}}
            <div style="width: 100%; height: 200px; background-color: #f8f9fa; border: 1px dashed #ccc; border-radius: 5px; margin-bottom: 20px; display: flex; align-items: center; justify-content: center; color: #666;">
                <i>Preview tidak tersedia untuk format .{{{{ ext }}}}. Dokumen tetap bisa dicetak.</i>
            </div>
        {{% endif %}}
        
        <p style="font-size: 13px; color: #888; margin-top: -10px; margin-bottom: 20px; font-style: italic;">
            * Tampilan di atas adalah dokumen asli Anda. Pengaturan ukuran kertas dan orientasi di bawah ini tidak mengubah gambar preview, namun akan diterapkan langsung ke mesin printer saat proses pencetakan.
        </p>
        
        <form action="/bagiprinter/print/{{{{ filename }}}}" method="post" onsubmit="showLoader()">
            <div class="form-group">
                <label>Ukuran Kertas:</label>
                <select name="paper_size">
                    <option value="9">A4</option>
                    <option value="14">F4 / Folio</option>
                    <option value="1">Letter</option>
                    <option value="5">Legal</option>
                    <option value="11">A5</option>
                </select>
            </div>
            <div class="form-group">
                <label>Orientasi:</label>
                <select name="orientation">
                    <option value="1">Portrait</option>
                    <option value="2">Landscape</option>
                </select>
            </div>
            <button type="submit">Cetak Sekarang</button>
            <a href="/bagiprinter/" style="text-decoration: none;"><button type="button" class="cancel">Batal</button></a>
        </form>
    </div>
</body>
</html>
'''

@app.route('/bagiprinter/')
@app.route('/bagiprinter')
def index():
    return render_template_string(HTML_TEMPLATE, printer_name=PRINTER_NAME)

@app.route('/bagiprinter/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        flash('Tidak ada file yang dikirim.', 'error')
        return redirect('/bagiprinter/')
    file = request.files['file']
    if file.filename == '':
        flash('Belum ada file yang dipilih.', 'error')
        return redirect('/bagiprinter/')
    
    ext = file.filename.split('.')[-1].lower() if '.' in file.filename else ''
    allowed_pdf = {'pdf', 'jpg', 'jpeg', 'png'}
    allowed_word = {'doc', 'docx', 'rtf'}
    allowed_excel = {'xls', 'xlsx'}
    
    if file and ext in (allowed_pdf | allowed_word | allowed_excel):
        filename = secure_filename(file.filename)
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(filepath)
        
        preview_filename = filename
        
        if ext in allowed_word or ext in allowed_excel:
            # Generate unique name to avoid overwriting original pdfs if names match
            pdf_filename = filename.rsplit('.', 1)[0] + f"_{int(time.time())}.pdf"
            pdf_filepath = os.path.join(app.config['UPLOAD_FOLDER'], pdf_filename)
            try:
                convert_to_pdf(filepath, pdf_filepath, ext)
                preview_filename = pdf_filename
            except Exception as e:
                flash(f'Gagal memproses dokumen (Word/Excel mungkin tidak terinstall di server): {str(e)}', 'error')
                return redirect('/bagiprinter/')
                
        return redirect(f'/bagiprinter/preview/{preview_filename}')
    else:
        flash(f'Format {ext} tidak didukung. Silakan gunakan format yang diizinkan.', 'error')
        return redirect('/bagiprinter/')

@app.route('/bagiprinter/uploads/<filename>')
def serve_upload(filename):
    return send_from_directory(app.config['UPLOAD_FOLDER'], filename)

@app.route('/bagiprinter/preview/<filename>')
def preview(filename):
    ext = filename.split('.')[-1].lower() if '.' in filename else ''
    return render_template_string(PREVIEW_TEMPLATE, filename=filename, ext=ext)

@app.route('/bagiprinter/print/<filename>', methods=['POST'])
def print_file(filename):
    filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    if not os.path.exists(filepath):
        flash('File tidak ditemukan.', 'error')
        return redirect('/bagiprinter/')
        
    paper_size = int(request.form.get('paper_size', 9))
    orientation = int(request.form.get('orientation', 1))
    
    try:
        # Buka koneksi printer
        hprinter = win32print.OpenPrinter(PRINTER_NAME, {"DesiredAccess": win32print.PRINTER_ALL_ACCESS})
        
        # Ambil devmode saat ini
        pinfo = win32print.GetPrinter(hprinter, 2)
        devmode = pinfo['pDevMode']
        
        # Simpan state lama
        old_paper_size = devmode.PaperSize
        old_orientation = devmode.Orientation
        
        # Ubah devmode sesuai settingan
        devmode.PaperSize = paper_size
        devmode.Orientation = orientation
        
        # Terapkan perubahan sementara
        win32print.SetPrinter(hprinter, 2, pinfo, 0)
        
        # Eksekusi Print
        win32api.ShellExecute(0, "print", filepath, f'/d:"{PRINTER_NAME}"', ".", 0)
        
        # Beri waktu sebentar agar driver menangkap devmode yang baru sebelum dikembalikan
        time.sleep(3)
        
        # Kembalikan settingan awal (opsional tapi disarankan)
        devmode.PaperSize = old_paper_size
        devmode.Orientation = old_orientation
        win32print.SetPrinter(hprinter, 2, pinfo, 0)
        
        win32print.ClosePrinter(hprinter)
        
        flash(f'Sukses! Dokumen {filename} telah dikirim ke printer dengan setting yang dipilih.', 'success')
    except Exception as e:
        flash(f'Terjadi kesalahan saat mencetak: {str(e)}', 'error')
        
    return redirect('/bagiprinter/')

if __name__ == '__main__':
    print(f"Server berjalan! Menggunakan printer default: {PRINTER_NAME}")
    app.run(host='0.0.0.0', port=8181)
