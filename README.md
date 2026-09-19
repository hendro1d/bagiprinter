# BagiPrinter

BagiPrinter adalah sebuah aplikasi berbasis web lokal (Plug-and-Play) yang memungkinkan Anda untuk berbagi printer (seperti Canon G1010) dengan komputer dan perangkat genggam lain yang berada dalam satu jaringan WiFi.

Aplikasi ini dibuat menggunakan Python dengan framework Flask, serta memanfaatkan Windows API (`win32print`, `win32api`, `win32com`) untuk mengatur orientasi kertas dan mengonversi format dokumen secara otomatis.

## Fitur Utama
- **Cetak Langsung dari Browser**: Pengguna di jaringan yang sama dapat mengunggah dokumen dan mencetaknya langsung melalui web browser.
- **Konversi Dokumen Otomatis**: Secara otomatis di belakang layar mengonversi dokumen Word (`.doc`, `.docx`, `.rtf`) dan Excel (`.xls`, `.xlsx`) menjadi PDF agar tidak memunculkan dialog aplikasi di komputer server.
- **Preview Dokumen**: Pengguna dapat meninjau (preview) dokumen PDF maupun gambar (`.jpg`, `.png`) langsung di dalam antarmuka web.
- **Pengaturan Kertas & Orientasi**: Mendukung berbagai ukuran kertas (A4, F4/Folio, Letter, Legal, A5) dan pengaturan orientasi (Portrait & Landscape).

## Prasyarat Server
1. Sistem Operasi **Windows** (karena menggunakan `win32api`).
2. **Python 3.x** terpasang.
3. Microsoft Office (Word & Excel) terpasang di komputer server untuk mendukung fitur konversi dokumen otomatis.
4. Pastikan *default printer* di Windows sudah diatur ke printer yang ingin Anda bagikan.

## Instalasi

1. Clone repositori ini.
2. Buka terminal atau *Command Prompt* dan masuk ke direktori proyek.
3. Instal semua dependensi yang dibutuhkan dengan perintah:
   ```bash
   pip install flask werkzeug pywin32 zeroconf
   ```
4. Jalankan aplikasi menggunakan perintah:
   ```bash
   python app.py
   ```
5. Akses aplikasi melalui alamat `http://localhost:8181` atau `http://<IP-Address-Komputer>:8181` di browser.

## Catatan
- Jika perangkat lain dalam jaringan yang sama gagal mengakses aplikasi, pastikan *Windows Firewall* tidak memblokir port `8181` pada komputer server.
