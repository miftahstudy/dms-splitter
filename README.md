# DMS Data Import Splitter

Aplikasi Streamlit untuk:
1. Menggabungkan data PO dari satu/banyak file (xlsx/csv).
2. Memetakan otomatis ke 12 kolom template resmi DMS.
3. **Pemrosesan 100% otomatis (tanpa edit manual):**
   - Baris dengan **Qty = 0** — **satu-satunya kondisi yang dihapus**.
   - Baris dengan field wajib kosong, Qty tidak valid (bukan angka/bukan bulat), atau
     PO Date tidak terbaca — **tetap dicantumkan apa adanya**, tidak dihapus, hanya
     diberi keterangan di kolom tambahan **"Catatan Validasi"** (kolom setelah Qty).
   - PO Date yang formatnya berantakan tapi masih bisa dibaca (`dd/mm/yyyy`, nama bulan
     Indonesia/Inggris, angka serial Excel, dll) dirapikan otomatis ke `yyyy-mm-dd`.
4. Menyediakan **ringkasan & rincian setiap aksi** yang dilakukan (berapa baris dihapus,
   berapa yang ditandai per kategori, baris mana saja, dan tanggal mana saja yang
   formatnya diubah beserta nilai asli → nilai baru) — bisa dilihat di layar maupun
   diunduh sebagai laporan Excel.
5. Membagi data menjadi beberapa file `.xlsx` (default 6.000 baris/file, bisa diubah),
   tetap memakai header, catatan instruksi, dan formatting resmi dari template.
6. **Dua versi hasil**, masing-masing bisa didownload satuan per file maupun sekaligus ZIP:
   - **Versi Split Saja** — data asli 100% apa adanya langsung dibagi ke template, tanpa
     ada yang dihapus/diubah/ditambah kolom apa pun.
   - **Versi Sudah Diproses** — Qty=0 sudah dihapus, PO Date dirapikan, dan ada kolom
     Catatan Validasi untuk baris lain yang masih bermasalah.

## Menjalankan di komputer sendiri

```bash
pip install -r requirements.txt
streamlit run app.py
```

Buka `http://localhost:8501` di browser.

## Deploy gratis & publik (Streamlit Community Cloud)

1. Buat repository baru di GitHub (boleh publik atau privat), lalu upload seluruh isi
   folder ini (`app.py`, `requirements.txt`, folder `assets/`).
2. Buka https://share.streamlit.io, login dengan akun GitHub.
3. Klik **"New app"** → pilih repository & branch tadi → Main file path isi `app.py`.
4. Klik **Deploy**. Setelah build selesai, aplikasi akan punya URL publik
   (contoh: `https://nama-app.streamlit.app`) yang bisa diakses siapa saja tanpa login,
   kecuali Anda mengatur app tersebut sebagai privat di pengaturan Streamlit Cloud.
5. Setiap kali kode di GitHub diupdate, tinggal *push* lagi — Streamlit Cloud akan
   otomatis redeploy.

> Alternatif lain: Render.com, Railway, atau server internal perusahaan (asal Python
> bisa jalan) — cukup jalankan `streamlit run app.py --server.port $PORT --server.address 0.0.0.0`.

## Template

Folder `assets/default_template.xlsx` berisi template resmi bawaan (dipakai otomatis
jika pengguna tidak mengupload template sendiri di sidebar). Jika suatu saat DMS
mengubah format template, cukup upload template baru lewat sidebar — aplikasi akan
membaca ulang posisi kolom & baris header secara otomatis, tidak perlu ubah kode.

## Catatan

- Baris kosong otomatis diabaikan (tidak dihitung sebagai data).
- Satu-satunya kondisi yang benar-benar **menghapus** baris adalah **Qty = 0**.
- Qty negatif **tidak** dihapus dan tidak ditandai (aturan template hanya melarang
  Qty = 0 / bukan bilangan bulat) — kalau ternyata Qty negatif juga harus dianggap
  tidak valid di proses bisnis Anda, beri tahu agar aturannya ditambahkan.
- Kolom **"Catatan Validasi (otomatis)"** hanya muncul di **Versi Sudah Diproses**,
  ditambahkan tepat setelah kolom Qty (kolom ke-13). Versi Split Saja tetap murni
  12 kolom asli tanpa tambahan apa pun.
- Laporan Excel (`..._Laporan_Pembersihan.xlsx`) berisi sheet Ringkasan + rincian per
  kategori (baris dihapus, baris ditandai per jenis masalah, dan tanggal mana saja
  yang formatnya diperbaiki beserta nilai asli → nilai baru) sebagai jejak audit.
