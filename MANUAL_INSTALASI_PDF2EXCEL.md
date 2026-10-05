# Manual Instalasi PDF2Excel pada Windows

## Kondisi Sistem

Manual ini ditujukan untuk komputer yang menggunakan **Python 3.14 sebagai Python utama**, sementara PDF2Excel dijalankan menggunakan **Python 3.11 pada virtual environment terpisah**.

Python 3.14 tidak perlu dihapus. Kedua versi dapat terpasang berdampingan:

```text
Python 3.14  -> tetap menjadi Python utama
Python 3.11  -> khusus virtual environment PDF2Excel
```

Alasan penggunaan Python 3.11 adalah kompatibilitas dependency OCR. Konfigurasi PDF2Excel menggunakan PaddleOCR/PaddlePaddle, dan environment Python 3.14 belum cocok untuk stack tersebut. Selain itu, `numpy==1.26.4` tidak memiliki wheel untuk Python 3.14 sehingga pip mencoba mengompilasi package dari source.

---

## 1. Tutup Virtual Environment Lama

Jika Command Prompt masih menampilkan `(.venv)`, jalankan:

```cmd
deactivate
```

Jika perintah tersebut tidak berhasil, tutup Command Prompt dan buka Command Prompt baru.

Jangan membuat ulang `.venv` saat environment lama masih aktif.

---

## 2. Instal Python 3.11 64-bit

Instal Python 3.11 64-bit sebagai instalasi per-user. Python 3.14 tetap dipertahankan.

Saat menjalankan installer Python 3.11:

1. Aktifkan **Add python.exe to PATH**.
2. Gunakan instalasi per-user.
3. Jangan memilih lokasi seperti `C:\Program Files` jika memerlukan Administrator.
4. Selesaikan instalasi, lalu tutup dan buka kembali Command Prompt.

Periksa seluruh instalasi Python:

```cmd
py -0p
```

Contoh hasil yang benar:

```text
-V:3.14  C:\Users\NamaUser\AppData\Local\Python\pythoncore-3.14-64\python.exe
-V:3.11  C:\Users\NamaUser\AppData\Local\Programs\Python\Python311\python.exe
```

Periksa Python 3.11 secara langsung:

```cmd
py -3.11 --version
```

Hasil yang diharapkan:

```text
Python 3.11.x
```

Jika `py -3.11` tidak ditemukan, Python 3.11 belum terpasang dengan benar atau Windows Python Launcher belum mengenal instalasinya.

---

## 3. Masuk ke Folder PDF2Excel

```cmd
cd C:\Users\Desyanto.kurniawan\Documents\PDF2Excel
```

Periksa isi folder:

```cmd
dir
```

Pastikan minimal terdapat:

```text
pdf_to_excel.py
requirements.txt
README.md
input
output
models
temp
```

---

## 4. Hapus Virtual Environment yang Gagal

Pastikan prompt tidak lagi diawali `(.venv)`, kemudian jalankan:

```cmd
rmdir /s /q .venv
```

Jika folder sedang digunakan:

1. Tutup semua Command Prompt atau terminal yang masih mengaktifkan `.venv`.
2. Pastikan tidak ada proses Python PDF2Excel yang berjalan.
3. Buka Command Prompt baru.
4. Ulangi perintah penghapusan.

---

## 5. Buat Virtual Environment dengan Python 3.11

Gunakan Python Launcher secara eksplisit. Jangan gunakan `python -m venv .venv`, karena perintah `python` Anda mengarah ke Python 3.14.

```cmd
py -3.11 -m venv .venv
```

Aktifkan environment:

```cmd
.venv\Scripts\activate
```

Prompt seharusnya berubah menjadi:

```text
(.venv) C:\Users\Desyanto.kurniawan\Documents\PDF2Excel>
```

Verifikasi interpreter:

```cmd
python --version
where python
```

Hasil yang diharapkan:

```text
Python 3.11.x
C:\Users\Desyanto.kurniawan\Documents\PDF2Excel\.venv\Scripts\python.exe
```

Jika masih menunjukkan Python 3.14, hapus `.venv` dan ulangi pembuatan dengan `py -3.11 -m venv .venv`.

---

## 6. Upgrade Peralatan Instalasi

Setelah `.venv` Python 3.11 aktif:

```cmd
python -m pip install --upgrade pip setuptools wheel
```

Periksa bahwa pip berasal dari `.venv`:

```cmd
python -m pip --version
```

Path yang ditampilkan harus mengarah ke:

```text
...\PDF2Excel\.venv\Lib\site-packages\pip
```

---

## 7. Instal Dependency PDF2Excel

Jalankan:

```cmd
python -m pip install -r requirements.txt
```

Gunakan `python -m pip`, bukan hanya `pip`, agar package dipasang ke interpreter aktif yang benar.

Instalasi akan memasang komponen seperti:

- PyMuPDF
- OpenCV
- NumPy
- OpenPyXL
- Pillow
- PaddlePaddle
- PaddleOCR

Model OCR mungkin diunduh pada penggunaan OCR pertama. Setelah tersimpan dalam cache lokal, model tidak perlu diunduh setiap kali program dijalankan.

---

## 8. Jalankan Diagnostic

```cmd
python pdf_to_excel.py --diagnose
```

Hasil yang diharapkan:

```text
PDF2Excel diagnostic
Python: 3.11.x
Writable output: True
Missing: none
```

Jika masih ada dependency yang hilang, pastikan:

```cmd
python --version
where python
python -m pip --version
```

Ketiga pemeriksaan harus menunjuk ke Python 3.11 dan folder `.venv` milik PDF2Excel.

---

## 9. Menjalankan GUI

```cmd
python pdf_to_excel.py
```

Alur penggunaan:

1. Klik **SELECT PDF**.
2. Pilih satu atau beberapa file PDF.
3. Biarkan mode **Auto** untuk penggunaan umum.
4. Klik **CONVERT TO EXCEL**.
5. Hasil disimpan ke folder `output`.

---

## 10. Menjalankan CLI

Konversi dengan nama output otomatis:

```cmd
python pdf_to_excel.py input\test.pdf
```

Menentukan nama output:

```cmd
python pdf_to_excel.py input\test.pdf -o output\hasil.xlsx
```

Mode tabel:

```cmd
python pdf_to_excel.py input\test.pdf --mode table
```

Memaksa OCR:

```cmd
python pdf_to_excel.py input\test.pdf --mode ocr --dpi 300
```

Mempertahankan layout relatif:

```cmd
python pdf_to_excel.py input\test.pdf --mode layout
```

Menampilkan bantuan:

```cmd
python pdf_to_excel.py --help
```

---

## 11. Menonaktifkan Virtual Environment

Setelah selesai:

```cmd
deactivate
```

Setelah `deactivate`, perintah berikut dapat kembali menunjukkan Python utama 3.14:

```cmd
python --version
```

Hal tersebut normal. Python 3.11 hanya digunakan ketika `.venv` PDF2Excel aktif.

---

## 12. Urutan Perintah Lengkap

Setelah Python 3.11 terpasang, jalankan dari Command Prompt baru:

```cmd
cd C:\Users\Desyanto.kurniawan\Documents\PDF2Excel

rmdir /s /q .venv

py -3.11 --version
py -3.11 -m venv .venv

.venv\Scripts\activate

python --version
where python

python -m pip install --upgrade pip setuptools wheel
python -m pip install -r requirements.txt

python pdf_to_excel.py --diagnose
python pdf_to_excel.py
```

---

## Troubleshooting

### A. `Unable to copy venvlauncher.exe`

Penyebab yang paling umum adalah `.venv` lama masih aktif atau sedang digunakan.

Solusi:

```cmd
deactivate
```

Tutup terminal, buka Command Prompt baru, kemudian:

```cmd
cd C:\Users\Desyanto.kurniawan\Documents\PDF2Excel
rmdir /s /q .venv
py -3.11 -m venv .venv
```

### B. NumPy Mengunduh File `.tar.gz`

Jika pip mengunduh `numpy-1.26.4.tar.gz`, interpreter kemungkinan masih Python 3.14. Pada Python 3.11 Windows 64-bit, pip seharusnya mengambil wheel yang sesuai dan tidak mencoba build NumPy dari source.

Periksa:

```cmd
python --version
where python
```

### C. `No matching distribution found for paddlepaddle`

Periksa apakah environment benar-benar menggunakan Python 3.11 64-bit:

```cmd
python --version
python -c "import platform; print(platform.architecture()); print(platform.machine())"
```

Hasil yang diharapkan mencantumkan `64bit` dan arsitektur seperti `AMD64`.

### D. `Missing dependency`

Instal ulang requirements menggunakan interpreter aktif:

```cmd
python -m pip install -r requirements.txt
python pdf_to_excel.py --diagnose
```

### E. File Excel Tidak Bisa Ditulis

Pastikan:

- File Excel hasil sebelumnya tidak sedang dibuka.
- Folder `output` dapat ditulis.
- Project tidak berada di `C:\Program Files` atau folder sistem.
- Nama output tidak mengandung karakter Windows yang tidak valid.

### F. Download Model OCR Gagal

Pastikan koneksi internet tersedia pada penggunaan OCR pertama. Setelah model berhasil tersimpan dalam cache, penggunaan berikutnya dapat berjalan secara lokal. Periksa `pdf2excel.log` untuk detail teknis.

---

## Catatan Penting

- Jangan hapus Python 3.14 hanya untuk menjalankan PDF2Excel.
- Jangan membuat `.venv` PDF2Excel dengan Python 3.14.
- Selalu aktifkan `.venv` sebelum menjalankan PDF2Excel.
- Gunakan `python -m pip` untuk menghindari instalasi ke interpreter yang salah.
- Review hasil OCR, terutama angka finansial, tanggal, kode, ID, dan scan berkualitas rendah.
- Quality indicator adalah estimasi internal, bukan jaminan akurasi ilmiah.
