# Agents.md — Panduan Pengguna AI

File ini berisi catatan dan catatan untuk membantu AI agents bekerja pada proyek JX Meal Checklist Generator.

## Struktur Proyek

```
C:\xampp\htdocs\jxgenerator\
├─ index.html       : Halaman frontend (PHP tidak diperlukan di sini, HTML statis)
├─ api/
│   ├─ detect.php   : Endpoint deteksi tanggal menggunakan PhpSpreadsheet
│   └─ generate.php : Endpoint generate checklist dengan template base64
├─ assets/
│   └─ img/
│       ├─ aerofood.png   : Logo di atas judul halaman
│       └─ favicon.ico    : Ikon tab browser
├─ vercel.json      : Konfigurasi runtime Vercel (vercel-php@0.8.0)
├─ composer.json    : Dependensi PhpSpreadsheet 2.4.8
└─ composer.lock
```

## Alur Kerja Umum

1. **Deteksi Tanggal** — Pengguna upload file acuan Meal Charge → `api/detect.php`→ Daftar tanggal muncul di dropdown
2. **Generate** — Pengguna memilih tanggal → `api/generate.php`→ File .xlsx di-generate dan di-download
3. **Modal** — Klik tombol `?` atau link "Panduan Penggunaan" → modal muncul dengan animasi fade-in/slide-up
4. **Tutup Modal** — Hanya via tombol `×` di kanan atas atau tombol "Tutup" di bagian bawah; klik di luar tidak menutup

## Catatan Penting untuk AI

- **Template di-embed** — Template checklist ada sebagai base64 di `api/generate.php`, tidak ada file .xlsx terpisah yang perlu di-upload ke server
- **Favicone & Logo** — `assets/img/aerofood.png` dan `assets/img/favicon.ico` harus ada di tempat yang sama
- **PHP Extensions** — `extension=zip` harus aktif di `php.ini` untuk PhpSpreadsheet bekerja
- **Vercel Runtime** — Gunakan `vercel-php@0.8.0`, pastikan `vercel.json` ada di root
- **Tidak ada Python** — Versi desktop di `D:\JX Generator v1.2\`, versi web hanya PHP + PhpSpreadsheet
- **Section 7 (Dilarang)** — Hanya 2 poin: jangan menutup file acuan selama generate, jangan rubah nama/pindah sheet di file acuan
- **Dua poin template sudah dihapus** — "Jangan ubah nama file template atau hapus file template." dan "Jangan edit isi file template." sudah dihapus dari panduan

## Troubleshooting

| Masalah | Solusi |
|---------|--------|
| Tanggal tidak muncul | Pastikan file acuan valid .xlsx dan extension=zip aktif |
| Generate gagal | Cek konsol browser, pastikan PHP tidak error saat decode base64 template |
| Modal tidak tutup | Pastikan klick via tombol X atau Tutup, klik luar tidak berfungsi |
| Scrollbar keluar posisi | Cek CSS `.modal-body::-webkit-scrollbar` di `index.html` |

## Versi

- Web: PHP + PhpSpreadsheet 2.4.8
- Deploy: Vercel dengan vercel-php@0.8.0
- Bahasa: Indonesia (ID)