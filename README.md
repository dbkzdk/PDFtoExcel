# PDF2Excel

Aplikasi desktop Windows dan CLI untuk mengubah PDF digital maupun hasil scan menjadi Excel `.xlsx` yang editable. Satu processing engine dipakai oleh GUI dan CLI.

## Persyaratan

- Windows 10/11 64-bit
- Python 3.11 64-bit direkomendasikan
- RAM 8 GB disarankan untuk scan besar
- Internet hanya diperlukan saat instalasi dan unduhan model OCR pertama kali. Model berikutnya menggunakan cache lokal PaddleOCR.
- Tidak memerlukan hak Administrator. Simpan project di folder yang dapat ditulis user, bukan `C:\Program Files`.

## Instalasi

```cmd
cd PDF2Excel
python --version
python -m venv .venv
.venv\Scripts\activate
python -m pip install --upgrade pip setuptools wheel
pip install -r requirements.txt
python pdf_to_excel.py --diagnose
```

Pada sebagian mesin, PaddlePaddle menyediakan wheel CPU sesuai versi Python/Windows. Jika instalasi `paddlepaddle` gagal, buka dokumentasi instalasi PaddlePaddle dan pilih perintah CPU Windows untuk Python 3.11, lalu ulangi `pip install -r requirements.txt`.

## GUI

```cmd
python pdf_to_excel.py
```

1. Klik **SELECT PDF** dan pilih satu atau beberapa PDF.
2. Biarkan mode **Auto** untuk penggunaan umum.
3. Klik **CONVERT TO EXCEL**.

Output default disimpan di `output\`. Tombol Cancel menghentikan proses di batas halaman berikutnya. GUI memakai worker thread sehingga jendela tetap responsif.

## CLI

```cmd
python pdf_to_excel.py input\test.pdf
python pdf_to_excel.py input\test.pdf -o output\hasil.xlsx
python pdf_to_excel.py input\test.pdf --mode table
python pdf_to_excel.py input\test.pdf --mode ocr --dpi 300
python pdf_to_excel.py input\test.pdf --mode layout
python pdf_to_excel.py --help
```

## Mode

- `auto`: memakai text layer PDF bila memadai, lalu OCR untuk halaman scan.
- `table`: fokus clustering baris/kolom dari posisi teks.
- `ocr`: memaksa OCR untuk semua halaman.
- `layout`: menempatkan teks pada grid Excel berdasarkan koordinat relatif halaman.

## Cara Kerja

1. PDF dibuka secara lokal dengan PyMuPDF, tanpa mengeksekusi macro atau script dokumen.
2. Halaman diproses satu per satu untuk membatasi RAM.
3. Halaman scan dirender, ditingkatkan kontrasnya, di-denoise, dan di-deskew dengan OpenCV.
4. PaddleOCR menghasilkan teks, bounding polygon, dan confidence.
5. Koordinat dinormalisasi 0–1 agar pergeseran scan tidak langsung menjadi koordinat absolut.
6. Baris dikelompokkan secara adaptif dari median tinggi teks. Kolom diperkirakan dari anchor posisi X.
7. OpenPyXL membuat workbook editable, format header, border, wrap, freeze pane, autofilter, tipe data, dan komentar untuk OCR confidence rendah.

## Folder

- `input\`: lokasi input opsional.
- `output\`: output default. Nama baru `_1`, `_2`, dst. dipakai agar file tidak tertimpa.
- `models\`: disediakan untuk model portable/manual.
- `temp\`: workspace lokal; implementasi saat ini memproses gambar di memori per halaman sehingga tidak meninggalkan gambar sementara.
- `pdf2excel.log`: log proses dan error teknis.

## Catatan Akurasi

Indikator quality adalah heuristik internal dari confidence OCR dan konsistensi tabel, bukan akurasi ilmiah. Cell ber-confidence rendah diberi warna kuning dan komentar. Selalu lakukan review, khususnya untuk nilai finansial, kode, ID, tanggal, tabel tanpa garis, dan scan buram.

Bahasa model default menggunakan model Latin/English PaddleOCR, yang sesuai untuk alfabet Indonesia dan Inggris. Parameter `language` disediakan pada konfigurasi internal untuk perluasan selanjutnya.

## Batasan Versi Pertama

- Rekonstruksi merged cell dan tabel sangat kompleks bersifat heuristik.
- Rotasi 90/180/270 bergantung pada document orientation classifier PaddleOCR.
- Tabel lintas halaman digabung bila skor struktur halaman relatif serupa; header berulang pertama dilewati.
- PDF terenkripsi harus dibuka proteksinya terlebih dahulu.
- Tidak ada klaim 100% akurat. Scan sulit harus ditinjau.

## Pengujian

Jalankan diagnostic:

```cmd
python pdf_to_excel.py --diagnose
```

Checklist penerimaan manual:

1. PDF digital sederhana
2. scan biasa
3. scan miring
4. scan tidak center
5. tabel dengan border
6. tabel tanpa border
7. multi-page
8. blank page
9. scan kualitas rendah
10. mixed text/table

Buka hasil dengan Microsoft Excel atau LibreOffice Calc dan validasi cell, angka, tanggal, header, serta warning OCR.
