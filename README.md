# Perangkat

Web statis untuk data perangkat .xlsx

## Alur operator
1. Perubahan selanjutnya cukup edit `perangkat.xlsx` 
2. Taruh foto di `foto/NAMA_DESA/NAMA.jpg`. Nama berkas harus sama dengan kolom NAMA (huruf besar/kecil dan tanda baca diabaikan).
3. `git add . && git commit -m "update" && git push`.
4. Cek tab Actions. Merah berarti ada ERROR di Excel; rinciannya ada di ringkasan job.

## Pengaturan Lain
`config.json`: nama kecamatan, sheet yang diabaikan, usia pensiun (60; pensiun jatuh pada tanggal 1 bulan berikutnya setelah berusia 60), dan `bulan_peringatan` (lencana "Pensiun …" muncul sekian bulan sebelumnya).
