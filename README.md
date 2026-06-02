# Convert Mania — Panduan Penggunaan & Pengembangan

Convert Mania adalah aplikasi web berbasis client-side (offline-first) untuk mengonversi kartu soal pilihan ganda (dari teks langsung, Word, atau Excel) ke template soal Moodle (.docx) secara otomatis.

Dokumen ini berisi panduan untuk menjalankan aplikasi serta cara melanjutkan pengembangan di komputer lain (Windows maupun macOS).

---

## 🚀 Cara Menjalankan Aplikasi

Aplikasi ini bersifat **serverless/client-side**, artinya berjalan sepenuhnya di dalam browser Anda tanpa memerlukan server backend.

### Cara 1: Buka Langsung (Paling Mudah)
1. Buka folder proyek **Convert Mania**.
2. Klik ganda (atau klik kanan -> Open With browser) pada file [index.html](file:///d:/MyLearn/covert-soal/index.html).
3. Aplikasi akan langsung berjalan di browser Anda (Chrome, Firefox, Safari, atau Edge).

### Cara 2: Menggunakan Local Server (Sangat Direkomendasikan)
Beberapa browser memiliki batasan keamanan ketat ketika membuka file lokal melalui protokol `file://` (terutama macOS Safari/Chrome). Untuk performa terbaik dan menghindari pemblokiran skrip, jalankan dengan server lokal:

#### Di macOS / OS X (Terminal):
1. Buka aplikasi **Terminal**.
2. Masuk ke direktori proyek menggunakan perintah `cd`:
   ```bash
   cd /path/to/folder/covert-soal
   ```
3. Jalankan server Python bawaan macOS:
   ```bash
   python3 -m http.server 8000
   ```
4. Buka browser dan akses alamat: `http://localhost:8000`

#### Di Windows (PowerShell / Command Prompt):
* **Menggunakan VS Code (Paling Praktis):**
  1. Buka folder proyek di VS Code.
  2. Install ekstensi **Live Server**.
  3. Klik tombol **Go Live** di pojok kanan bawah VS Code.
* **Menggunakan Python (jika terinstall):**
  ```powershell
  python -m http.server 8000
  ```
  Lalu buka browser ke `http://localhost:8000`.

---

## 💻 Melanjutkan Pengembangan di Komputer Lain

Jika Anda ingin memindahkan proyek ini ke komputer baru (Windows atau macOS) untuk melanjutkan coding, ikuti langkah-langkah berikut:

### 1. Prasyarat (Prerequisites)
Pastikan komputer baru sudah terinstall **Node.js** (versi 16 atau lebih baru). Unduh di [nodejs.org](https://nodejs.org/).

### 2. Setup Proyek

#### A. Di Komputer Windows:
1. Buka **Command Prompt** atau **PowerShell**.
2. Arahkan ke folder proyek:
   ```powershell
   cd C:\path\to\covert-soal
   ```
3. Install dependencies (jika ingin memperbarui/mengembangkan library docx):
   ```powershell
   npm install
   ```

#### B. Di Komputer macOS:
1. Buka aplikasi **Terminal**.
2. Arahkan ke folder proyek:
   ```bash
   cd /path/to/covert-soal
   ```
3. Install dependencies:
   ```bash
   npm install
   ```

*(Catatan: Library `docx.js` versi browser telah di-bundle secara lokal di folder `js/docx.min.js` agar aplikasi tetap dapat berjalan secara offline/tanpa internet).*

---

## 📁 Struktur Direktori Proyek

* `index.html` : Halaman antarmuka utama (UI).
* `css/style.css` : File styling (Premium Dark Theme).
* `js/`
  * `app.js` : Mengontrol alur aplikasi (Wizard flow, interaksi modal, upload, dan tombol download).
  * `parsers.js` : Logika parsing teks biasa, Excel (.xlsx), dan Word (.docx).
  * `generator.js` : Logika pembuatan file `.docx` baru menggunakan pustaka `docx.js`.
  * `docx.min.js` : Bundle pustaka docx v7 lokal (untuk mode offline).
* `package.json` : Konfigurasi npm package dependencies.
* `template-input-soal-pg-sederhana.docx` : File referensi template soal Moodle.

---

## ⚠️ Catatan Penting
* **Koneksi Internet:** Aplikasi membutuhkan internet pada saat pertama kali dibuka karena pustaka **SheetJS (xlsx)** dan **Mammoth (docx parser)** dimuat melalui CDN publik. Jika ingin benar-benar offline, Anda dapat men-download skrip tersebut dan menyimpannya di folder `js/` lokal lalu memperbarui tag `<script>` di `index.html`.
* **Kunci Jawaban:** Pastikan input soal Anda memiliki kunci jawaban yang valid (A, B, C, D, atau E). Jika tidak ada kunci jawaban, sistem akan menampilkan indikator peringatan (⚠️).
