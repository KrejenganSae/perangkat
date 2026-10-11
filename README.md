# Perangkat Desa

Web statis dari `perangkat.xlsx`, dibangun otomatis oleh GitHub Action dan dipublikasikan lewat GitHub Pages.

## Alur operator
1. Edit `perangkat.xlsx` 
2. Taruh foto di `foto/NAMA_DESA/NAMA.jpg`. Nama berkas harus sama dengan kolom NAMA (huruf besar/kecil dan tanda baca diabaikan).
3. `git add . && git commit -m "update" && git push`.
4. Cek tab Actions. Merah berarti ada ERROR di Excel; rinciannya ada di ringkasan job.

## Lambang
Taruh lambang resmi Pemkab Probolinggo sebagai `logo.png` (latar transparan, tinggi minimal 200 px) di akar repositori. Tanpa berkas ini, header tampil tanpa lambang.

## Aktifkan sekali saja
Settings > Pages > Source: **GitHub Actions**.

## Uji lokal
`pip install -r requirements.txt && python build.py`, hasilnya di `dist/`.

## Pengaturan
`config.json`: nama kecamatan, sheet yang diabaikan, usia pensiun (60; pensiun jatuh pada tanggal 1 bulan berikutnya setelah berusia 60), dan `bulan_peringatan` (lencana "Pensiun …" muncul sekian bulan sebelumnya).
