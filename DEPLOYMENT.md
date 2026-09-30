# Deploy Cloudflare Pages gratis

Repository: https://github.com/mikoadi4312/Sales-Contest

Versi sekarang berupa web statis. Browser membaca Excel, menghitung bonus, dan menulis hasil. Tidak memakai Containers, Durable Objects, atau paket Workers Paid.

## Integrasi GitHub di Cloudflare

1. Workers & Pages > Create application > Pages > Connect to Git.
2. Pilih mikoadi4312/Sales-Contest, branch main.
3. Framework preset: None (atau Vite).
4. Build command: npm run build.
5. Build output directory: dist.
6. Root directory: kosong/root repository. NODE_VERSION: 22.
7. Save and Deploy. Push berikutnya ke main akan memperbarui situs.

Tidak perlu API key/token dalam environment aplikasi. Vite hanya memasukkan source web, logo, dan template distribusi fiktif ke build.

## Deployment langsung

```sh
npm ci
npm test
npm run build
npm run cf:login
npm run cf:whoami
npx wrangler pages project create sales-contest --production-branch main
npm run deploy
```

Lewati project create bila project sudah ada. Project direct upload menggunakan workflow Actions untuk deployment dari GitHub. Jika menginginkan integrasi Git bawaan Cloudflare, buat project melalui Connect to Git sejak awal.

## GitHub Actions

Check Pages application menjalankan audit, tes, dan build pada push. Workflow manual Deploy Cloudflare Pages menggunakan project sales-contest yang sudah dibuat.

Simpan CLOUDFLARE_ACCOUNT_ID dan CLOUDFLARE_API_TOKEN dalam Repository Settings > Secrets and variables > Actions. Token memerlukan Account > Cloudflare Pages > Edit untuk akun tujuan. Jangan menaruh token dalam source atau chat. Jalankan Actions > Deploy Cloudflare Pages > Run workflow.

## Batas operasional

- Static Pages mengikuti kuota Cloudflare Free. Tidak ada server perhitungan berbayar.
- Isi Revenue, template, dan mapping pengguna diproses lokal; tidak diunggah ke Cloudflare atau GitHub.
- Batas file 80 MB; kapasitas tergantung RAM browser. Contoh lokal 216.169 baris sudah diuji.
- Data tidak disimpan antar sesi/perangkat. Unduh hasil sebelum refresh atau menutup halaman.
- Situs tidak memiliki login. Cloudflare Access dapat ditambahkan pada domain jika diperlukan untuk akses internal.
- .env, .evv, secret, private key, dan data Excel asli tidak masuk Git. Jalankan npm run audit:repo setelah staging.

Dokumentasi: https://developers.cloudflare.com/pages/framework-guides/deploy-a-vite3-project/ dan https://developers.cloudflare.com/pages/get-started/direct-upload/.
