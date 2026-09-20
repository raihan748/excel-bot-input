# 📊 ExcelBot AI — Bot Otonom Pengubah Foto ke Excel

Aplikasi web modern dan otonom untuk mengubah puluhan foto dokumen, struk belanja, nota, tabel fisik, dan catatan keuangan menjadi file Excel (`.xlsx`) siap pakai secara otomatis menggunakan Vision AI.

Dilengkapi fitur **Histori Pembuatan File**, **In-Browser Spreadsheet Editor**, dan **Fitur Merge File Fleksibel (Append Baris atau Multi-Sheet)**.

---

## 🚀 Cara Deploy ke Vercel

Aplikasi ini sudah 100% kompatibel dengan arsitektur **Vercel Serverless Functions** (`/api/*.js`) dan frontend statis (`/public`).

### Opsi 1: Deploy via Vercel CLI (Paling Cepat)
1. Buka terminal di folder project ini.
2. Jalankan perintah:
   ```bash
   npx vercel
   ```
3. Ikuti instruksi di layar (pilih default).
4. Selesai! Web langsung online dan bisa diakses dari HP atau komputer bapak.

### Opsi 2: Deploy via GitHub + Dashboard Vercel
1. Upload folder project ini ke repository GitHub pribadi lu.
2. Buka [vercel.com](https://vercel.com) dan klik **Add New Project**.
3. Import repo GitHub tadi.
4. Di bagian **Environment Variables** (opsional):
   - `API_KEY` : `sk-ts-VB0BNV245K445QF7ZCRVCN6B7ADS`
   - `AI_MODEL` : `thirty/qwen3.7-max`
5. Klik **Deploy**. Selesai!

---

## 💻 Cara Menjalankan di Komputer Lokal

### Untuk Bapak (Paling Mudah — Tanpa Koding):
Cukup **klik 2x file `jalankan.bat`**. 
Aplikasi akan otomatis menyalakan server dan membuka browser di `http://localhost:3000`.

### Manual via Terminal:
```bash
npm install
npm start
```
Buka browser di `http://localhost:3000`.

---

## 🎯 Panduan Fitur untuk Bapak

| Fitur | Cara Penggunaan |
|---|---|
| **1. Pilih Tipe Dokumen** | Klik salah satu tombol preset: *🛒 Nota Belanja*, *💰 Buku Kas*, *📦 Stok Barang*, *📋 Absensi*, atau *🔍 Deteksi Bebas*. |
| **2. Upload Puluhan Foto** | Tarik dan lepaskan (drag & drop) 1 sampai puluhan foto sekaligus ke kotak upload. |
| **3. Mulai Proses Otomatis** | Klik tombol hijau **"Mulai Proses Otomatis"**. Bot AI akan membaca tiap foto satu per satu dan menyusun baris tabelnya. |
| **4. Edit Langsung di Web** | Jika ada tulisan nota yang kurang pas, klik langsung pada sel tabel di web untuk mengedit teks atau harganya. Total otomatis terkalkulasi ulang. |
| **5. Unduh Excel (.xlsx)** | Klik tombol **"Unduh File Excel (.xlsx)"** untuk menyimpan tabel rapi bergaris, berlatar warna, dan berumus `=SUM()`. |
| **6. Riwayat & Histori File** | Buka tab **"2. Riwayat & Histori File"**. Semua file yang pernah dibuat tersimpan permanen di sini. Bisa dibuka kembali atau didownload ulang kapan saja. |
| **7. Gabungkan File (Merge)** | Di tab Riwayat, centang 2 atau lebih file yang mau digabung, lalu klik **"Gabungkan File Terpilih"**. Pilih apakah mau digabung jadi 1 lembar bersambung (*Append*) atau dipisah per-sheet (*Multi-Sheet*). |

---

## 🛠️ Pengaturan Kunci API & Model
Secara default, aplikasi sudah terkonfigurasi dengan:
- **API Key**: `sk-ts-VB0BNV245K445QF7ZCRVCN6B7ADS`
- **Endpoint**: `https://api.thirtystore.com/v1/chat/completions`
- **Model Vision Teruji**: `thirty/qwen3.7-max` (Didukung juga: `thirty/qwen3.7-plus`, `thirty/qwen3.8-max`)

Pengaturan ini bisa diubah kapan saja melalui ikon **Gerigi (Settings)** di pojok kanan atas website.
