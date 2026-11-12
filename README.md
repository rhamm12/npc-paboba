# Next Page Challenge (NPC) by PABOBA

Frontend statis untuk Vercel. Data tersimpan di Google Sheets lewat Google Apps Script.

## Isi folder
- `index.html` : aplikasi utama
- `config.js` : tempat menempel URL Web App Apps Script
- `Code.gs` : cadangan kode backend (yang dipasang di Apps Script, bukan dijalankan Vercel)
- `vercel.json` : pengaturan Vercel

## Langkah singkat
1. Tempel `Code.gs` ke proyek Apps Script, lalu Deploy sebagai Web App (Execute as: Me, Who has access: Anyone).
2. Salin URL `/exec`, tempel ke `config.js`.
3. Push folder ini ke GitHub, impor ke Vercel (Framework Preset: Other), lalu Deploy.
