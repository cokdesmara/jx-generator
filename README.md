# JX Meal Checklist Generator

Web-based generator for creating Meal Number Check List files from Meal Charge input files. Built with PHP + PhpSpreadsheet, deployable to Vercel.

## Fitur

- **Drag & drop file upload** — upload file acuan Meal Charge (.xlsx)
- **Tanggal penerbangan** — dropdown otomatis dari file acuan
- **Progress bar** — menunjukkan persentase proses generate
- **Modal panduan** — tutup hanya via tombol X atau Tutup
- **Nama file otomatis** — include timestamp: `v2 Meal Number Check List JX 2026D - DD-MM-YYYY - YYYYMMDDHHMMSS.xlsx`
- **Format warna terpreservasi** — kuning (FFFF00), biru (FFC6D9F0), bold, wrap text, merges
- **SPML splitting** —pecah ke sel kru dengan format: `19 MBF1, 11 MCK1, 1 DBML`
- **Output — tidak tulis 0** — kotak kosong tetap kosong, bukan angka 0

## Arsitektur

```
/
├─ index.html       : Frontend single-page (drag-drop, date select, progress bar, modal)
├─ api/
│   ├─ detect.php   : Fast date detection using PhpSpreadsheet read filters
│   └─ generate.php : Full generate & download using embedded template base64
├─ assets/
│   └─ img/
│       ├─ aerofood.png   : Logo di atas judul
│       └─ favicon.ico    : Favicon
├─ vercel.json      : Vercel runtime config (vercel-php@0.8.0)
├─ composer.json    : PhpSpreadsheet 2.4.8 dependency
└─ composer.lock
```

## Endpoint API

| Endpoint | Method | Deskripsi |
|----------|--------|-----------|
| `/api/detect.php` | POST | Mendeteksi tanggal penerbangan dari file acuan |
| `/api/generate.php` | POST | Generate file checklist dan download |

## Deploy ke Vercel

1. Push repo ke GitHub
2. Import di Vercel
3. Vercel akan otomatis mendeteksi `vercel-php@0.8.0` dari `vercel.json`
4. Pastikan `extension=zip` aktif di PHP konfigurasi

## Catatan

- File acuan boleh tetap terbuka saat generate
- Output nama include timestamp — file lama tidak tertimpa
- Hasil format: pemisah koma tanpa spasi di B4/C4, spasi + koma di B6/B9/F5/F7/F9
- Jika input kosong: jangan tulis 0 (kosongkan saja)