# Sales Contest - Cloudflare Pages

Web bonus per toko, AM, dan RSM. Jumlah kategori mengikuti template yang dipilih. Versi utama berjalan sepenuhnya di browser dan bisa dihosting pada Cloudflare Pages gratis, tanpa server Python berbayar.

## Lokal

Node.js 22+: jalankan `npm ci`, lalu `npm run dev`. Buka http://127.0.0.1:5173. Build production: `npm run build`; preview: `npm run preview`.

## Pemakaian

1. Download template kriteria dan mapping di panel kiri. Template distribusi berisi contoh fiktif; ganti semua baris contoh.
2. Pilih satu atau beberapa template Criteria/Result, Revenue, serta mapping AM/RSM. File AM dan RSM terpisah bisa digabung berdasarkan Store ID.
3. Periksa periode, filter produk/status/output type, ranking, target, dan reward. Nama kategori harus unik.
4. Hitung dan download Resume Toko, Resume AM, Resume RSM, Result/Aturan/Criteria setiap kategori, serta Catatan.

File dibaca dan dihitung oleh Web Worker di browser. Isi file tidak dikirim ke server. Data/pengaturan tidak disimpan permanen; download hasil sebelum refresh atau menutup halaman. Batas pemilihan 80 MB per file; kapasitas praktis tergantung RAM perangkat.

## Perhitungan

- SO dari Revenue atau Quantity dengan filter periode dan produk.
- Toko: SO/target, minimum pencapaian, tier ranking. Seri: peringkat sama atau urutan Store ID.
- AM: target tier 90%/100%, reward berdasarkan jumlah toko dari Criteria atau mapping.
- RSM: target/reward tier 90%/100% setiap penerima. Satu toko boleh memiliki beberapa penerima regional.
- Cakupan manager: toko peserta kategori atau seluruh toko dalam mapping.
- Target/mapping/aturan belum lengkap menghasilkan catatan dan reward kosong, bukan nol final.

Skema khusus rasio pembiayaan Kredivo, growth antarperiode, dan persyaratan gabungan AC TCL belum didukung. Saran filter dari teks Scheme harus diperiksa sebelum menghitung.

## GitHub dan data

`.gitignore` mengecualikan `.env`, `.env.*`, `.evv`, private key, credential, secret, cache, serta Excel sumber/hasil. Hanya `assets/Template Kriteria Sales Contest.xlsx` (data fiktif) yang diizinkan masuk Git. `npm run audit:repo` memeriksa nama dan isi file dalam index Git tanpa mencetak nilai credential.

Aplikasi tidak memerlukan `.env`. Jangan menaruh token dalam JavaScript atau variabel frontend; kode frontend bersifat publik. Secret deployment opsional hanya disimpan di GitHub Actions Secrets.

## Verifikasi

Jalankan `npm test`, `npm run build`, dan `npm run audit:repo` (setelah staging). Tes mencakup filter, ranking seri, mapping terpisah, AM/RSM, kategori dinamis, parsing/ekspor Excel, dan data tidak lengkap. Rekonsiliasi Agustus dijalankan hanya bila file contoh asli tersedia lokal; data tersebut tidak masuk repository.

`web/` berisi aplikasi browser. `tests-js/` berisi tes. File Python/Streamlit tetap tersedia sebagai versi lokal lama dan referensi perhitungan; tidak masuk build Pages.

Panduan deployment: [DEPLOYMENT.md](DEPLOYMENT.md). Gaya visual dan logo mengacu pada halaman publik https://erabluebp.id/.
