# 04 — Panduan Pengembangan

Panduan untuk menjalankan, mengembangkan, dan memperluas fitur aplikasi **SoalMoodle Converter**.

Proyek ini juga diterbitkan di **Zenodo**. Lihat [.zenodo.json](/.zenodo.json) untuk metadata publikasi, [abstract-zenodo.md](abstract-zenodo.md) untuk abstrak, dan [paper-scientific.md](paper-scientific.md) untuk artikel ilmiah lengkap (Bahasa Indonesia).

---

## 1. Menjalankan Aplikasi

### 1.1 Tanpa Server (buka langsung)
1. Buka folder proyek.
2. Klik ganda `index.html` (atau klik kanan → Open With → browser).

> ⚠️ Beberapa browser (terutama Safari/Chrome di macOS) membatasi fungsi lewat protokol `file://`. Local server sangat disarankan.

### 1.2 Dengan Local Server (disarankan)

**macOS / Linux (Terminal):**
```bash
cd /path/to/soalmoodle-converter
python3 -m http.server 8000
# buka http://localhost:8000
```

**Windows (PowerShell):**
```powershell
cd C:\path\to\soalmoodle-converter
python -m http.server 8000
```

**Alternatif VS Code:** install ekstensi **Live Server**, klik **Go Live**.

---

## 2. Prasyarat & Setup

- **Node.js** versi 16 atau lebih baru ([nodejs.org](https://nodejs.org/)).

```bash
# Di dalam folder proyek
npm install        # hanya dependency 'docx'
```

> `js/docx.min.js` adalah **bundle lokal docx v7.8.2**, sehingga aplikasi tetap jalan tanpa internet untuk fungsi generate. Jika ingin memperbarui versi docx, jalankan `npm install docx@<versi>` lalu rebundle.

---

## 3. Struktur Direktori

```
soalmoodle-converter/
├── index.html                  # UI utama (wizard 3 langkah + modal edit)
├── package.json                # dependency: docx ^7.8.2
├── css/style.css               # Premium Dark Theme
├── js/
│   ├── app.js                  # Controller utama (alur, state, event)
│   ├── parsers.js              # Parsing teks/Excel/Word → SoalObject
│   ├── generator.js            # SoalObject → Blob .docx (docx library)
│   └── docx.min.js             # Library docx v7.8.2 (offline bundle)
├── extracted_template/         # Unzip template (referensi XML format)
└── docs/                       # Dokumentasi ini
    ├── template-input-soal-pg-sederhana.docx  # Template resmi Moodle
    ├── template-input-soal-pg-sederhana.zip
    └── sample/                 # Contoh dokumen hasil format
```

---

## 4. Cara Menambahkan Fitur

### 4.1 Menambahkan type field metadata baru (mis. "Guru")

**Parser teks** (`js/parsers.js`):
- Tambahkan key pada regex `metaRe` di `_parseSingleBlock()` (mis. `|Guru`).
- Tambahkan cabang `else if` untuk menyimpannya (mis. `soal.guru = val`).

**Parser Excel** (`js/parsers.js`):
- Tambahkan kolom pada objek `cols` (mis. `guru: findCol('guru', 'pengampu')`).
- Baca nilainya: `soal.guru = get(cols.guru);`

**Struktur `SoalObject`** (`createSoal`): tambahkan field baru dengan nilai default.

**Preview** (`js/app.js` — `_soalCardHTML`): tambahkan `<div class="field-item">` untuk menampilkan field.

**Generator** (`js/generator.js` — `metaTable`): tambahkan baris baru `metaRow('Guru', soal.guru || '')`.

> Ingat: perubahan pada 4 file ini (parser, createSoal, preview, generator) secara konsisten.

### 4.2 Menambahkan tipe input baru (mis. Upload CSV)
1. Tambahkan tombol tab baru di `index.html` (`data-tab="csv"`).
2. Tambahkan section `tab-csv` dengan dropzone/input.
3. Di `App._handleParse()` (`app.js`), tambahkan cabang `else if (this.currentTab === 'csv')` yang membaca file dan memanggil parser baru.
4. Tambahkan parser di `parsers.js` yang mengembalikan array `SoalObject`.

### 4.3 Mengubah urutan/langkah wizard
- Langkah diatur via section `id="step-1|2|3"`.
- Navigasi lewat `App._goToStep(n)` dan indikator `.step`.
- Untuk menambah langkah: duplikasi pola section + indikator, sesuaikan logika `_goToStep`.

### 4.4 Mengubah tema/styling
- Semua warna dan variabel dipusatkan di `css/style.css` (gunakan variabel seperti `--accent`, `--text-muted`).
- Edit di file tersebut, tidak perlu sentuh HTML.

---

## 5. Strategi Offline (Menghapus Ketergantungan CDN)

Saat ini `xlsx` dan `mammoth` dimuat via CDN, sehingga aplikasi **membutuhkan internet saat pertama kali dibuka** untuk tab Excel/Word.

Untuk membuat **benar-benar offline**:
1. Unduh `xlsx.full.min.js` dan `mammoth.browser.min.js`.
2. Simpan di `js/` (mis. `js/xlsx.min.js`, `js/mammoth.min.js`).
3. Ubah tag `<script>` di `index.html` untuk menggunakan file lokal:

```html
<script src="js/xlsx.min.js"></script>
<script src="js/mammoth.min.js"></script>
```

---

## 6. Menambah/Memperbarui Library `docx`

`js/docx.min.js` adalah bundle UMD dari package `docx`. Untuk memperbarui:
```bash
npm install docx@<versi-baru>
```
Kemudian buat ulang bundle lokal (mis. dengan bundler/script build) dan pastikan `window.docx` masih terekspos dengan export yang dibutuhkan `generator.js`:
- `Document`, `Paragraph`, `TextRun`, `Table`, `TableRow`, `TableCell`, `Packer`
- Enums: `AlignmentType`, `WidthType`, `BorderStyle`, `ShadingType`

`generator.js` sudah menggunakan fallback string-literal untuk enum, sehingga cukup toleran terhadap perubahan versi.

---

## 7. Tips Pengujian

- **Contoh cepat**: klik **"🔖 Muat 3 contoh soal →"** di tab teks untuk mengisi contoh langsung.
- **Uji ketiga tab**: teks, Excel, dan Word — pastikan hasil parse konsisten.
- **Periksa status**: cek statistik Valid/Warning/Error di preview; pastikan soal error sulit lolos generate.
- **Uji file generate**: buka `.docx` hasil dan bandingkan dengan `docs/template-input-soal-pg-sederhana.docx` (kolom, warna, border).
- **Uji offline**: buka halaman di mode pesawat / tanpa internet untuk memastikan `docx` bundle lokal berfungsi.

---

## 8. Catatan & Kendala Umum

- **Brute "output-soal.docx"** nama default; generate menimpa dengan nama dinamis `[mapel]_[kelas]_ASAT_2526.docx`.
- **Toleransi format teks**: parser fleksibel terhadap berbagai variasi penulisan (lihat [02-input-format.md](02-input-format.md)), sehingga input yang sedikit tidak rapi masih bisa diproses.
- **Kunci jawaban wajib**: soal tanpa kunci valid (A–E) ber-status `error` dan dilewati pada generate.

---

Dokumentasi lain: [README.md](README.md) · [01-arsitektur.md](01-arsitektur.md) · [02-input-format.md](02-input-format.md) · [03-modul-js.md](03-modul-js.md)
