# Aplikasi Laporan Akuisisi (Flask)

## Jalankan Project

1. Masuk ke folder project:
   - `cd C:\Users\Lenovo\Documents\akuisisi-app`
2. Install dependency:
   - `pip install -r requirements.txt`
3. Jalankan app:
   - `python app.py`
4. Buka browser:
   - `http://127.0.0.1:5000`

## Fitur yang sudah ada

- Input nomor laporan manual
- Validasi nomor laporan unik
- Form data inti laporan akuisisi
- Bagian A. Dasar dinamis (tambah/hapus item)
- Upload multi foto barang bukti
- Simpan data ke SQLite
- Generate DOCX dari template Word

## Cara pakai template Word

1. Simpan file template di folder project dengan nama:
   - `template_laporan.docx`
2. Isi placeholder dalam template, contoh:
   - `{{nomor_laporan}}`
   - `{{tanggal_laporan}}`
   - `{{nama_instansi}}`
   - `{{alamat_instansi}}`
   - `{{dasar_items}}`
   - `{{maksud_tujuan}}`
   - `{{prosedur_akuisisi}}`
   - `{{kesimpulan}}`
3. Klik tombol `Generate DOCX` di halaman daftar laporan.
