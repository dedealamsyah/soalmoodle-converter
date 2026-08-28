# 03 — Dokumentasi Modul JavaScript

Dokumentasi teknis tiga modul utama aplikasi: `app.js`, `parsers.js`, dan `generator.js`.

---

## 1. `js/app.js` — Kontrol Utama UI & Alur

**Tujuan:** Mengontrol seluruh alur wizard, interaksi pengguna, upload file, render preview, modal edit, serta generate/download.

### 1.1 State utama (`const App`)

```js
const App = {
  soalList: [],          // Array SoalObject — state inti aplikasi
  generatedBlob: null,   // Blob .docx hasil generate
  downloadFilename: 'output-soal.docx',
  currentTab: 'text',    // 'text' | 'excel' | 'word'
  editingIndex: -1,      // index soal yang sedang diedit
  excelFile: null,       // File excel terpilih
  wordFile: null         // File word terpilih
}
```

### 1.2 Metode Publik

| Metode | Fungsi |
|--------|--------|
| `init()` | Boot — memanggil `_setupTabs`, `_setupDropzones`, `_setupButtons`, `_setupModal`. Dipanggil saat `DOMContentLoaded`. |
| `_handleParse()` | Mendapatkan defaults, membaca input sesuai tab aktif, memanggil parser yang tepat, lalu menampilkan preview (langkah 2). |
| `_renderPreview()` | Menghitung statistik valid/warning/error, merender kartu soal, dan mengikat event edit/delete/expand. |
| `_handleGenerate()` | Menyaring soal `status !== 'error'`, memanggil `DocxGenerator.generate()`, membuat nama file dinamis, lalu pindah ke langkah 3. |
| `_handleDownload()` | Membuat URL objek dari blob dan memicu unduhan browser. |
| `_handleReset()` | Me-reset semua state dan kembali ke langkah 1. |

### 1.3 Metode Pendukung (private)

| Metode | Fungsi |
|--------|--------|
| `_setupTabs()` / `_switchTab(tab)` | Mengelola tab input (text/excel/word). |
| `_setupDropzones()` | Drag & drop + klik untuk memilih file Excel/Word. |
| `_handleFile(type, file)` | Menyimpan file terpilih & menampilkan info nama/ukuran. |
| `_getDefaults()` | Membaca nilai default dari form (mapel, kelas, dsb.). |
| `_readFileBuffer(file)` | Mem-baca file sebagai `ArrayBuffer` (Promise). |
| `_soalCardHTML(soal, index)` | Membuat HTML kartu soal untuk preview. |
| `_setupModal()` / `_openEditModal()` / `_saveEdit()` / `_closeModal()` | Modal edit per soal. |
| `_loadExample()` | Mengisi textarea dengan 3 contoh soal. |
| `_goToStep(n)` | Navigasi antar langkah wizard. |
| `_showLoading()` / `_hideLoading()` / `_showError()` | Utilitas UI (loading overlay & pesan error). |

### 1.4 Catatan Penting
- **Nama file unduhan dinamis**: `[mapel]_[kelas]_ASAT_2526.docx`, dibangun dari soal pertama atau nilai default. Karakter ilegal diganti `_`.
- **Soal ber-status `error` dilewati** saat generate; muncul konfirmasi jika ada yang dilewati.
- **Klik pada header kartu** toggles expand; tombol edit/delete dikecualikan.

---

## 2. `js/parsers.js` — Parsing & Validasi

**Tujuan:** Mengonversi berbagai format input menjadi array `SoalObject` valid.

### 2.1 Struktur

```js
const SoalParser = {
  normalizeLevel(val),
  validateSoal(soal),
  createSoal(num, defaults),
  parseText(rawText, defaults),
  _splitTextIntoBlocks(text),
  _parseSingleBlock(block, defaultNum, defaults),
  parseExcel(rows, defaults),
  parseWordText(extractedText, defaults)
}
```

### 2.2 `validateSoal(soal)`
Memberikan `status` dan mengisi `errors` / `warnings`. Lihat aturan lengkap di [02-input-format.md](02-input-format.md#4-aturan-validasi-status-soal).

### 2.3 `parseText(rawText, defaults)`
1. Menormalkan `\r\n` / `\r` → `\n`.
2. `_splitTextIntoBlocks()` memecah menjadi blok (prioritas: `SOAL N` → `---` → soal bernomor).
3. `_parseSingleBlock()` mem-parsing tiap blok dengan **state machine** (`meta` → `pertanyaan` / `stimulus` / `pilihan` / `pembahasan` / `post`).
4. Setiap hasil divalidasi.

### 2.4 `parseExcel(rows, defaults)`
- Baris `0` = header, dinormalisasi untuk deteksi kolom.
- `findCol(...candidates)` mencocokkan nama kolom (normalisasi: huruf kecil + hapus non-alfanumerik).
- Setiap baris data diubah jadi `SoalObject` + `validateSoal()`.

### 2.5 `parseWordText(extractedText, defaults)`
- Membersihkan artefak Mammoth (tab → spasi, spasi berlebih).
- Delegasi ke `parseText()`.

---

## 3. `js/generator.js` — Pembuatan .docx

**Tujuan:** Membangun file `.docx` **identik** dengan `template-input-soal-pg-sederhana.docx` (disimpan di `docs/`) menggunakan library `docx` v7.8.2 (bundle lokal `js/docx.min.js`).

### 3.1 Spesifikasi Layout (dari XML template asli)

| Elemen | Nilai |
|--------|-------|
| Judul dokumen | `#0B2545`, bold, 18pt (size 36) |
| Heading `SOAL N` | `#1F4D78`, bold, 14pt (size 28) |
| Label metadata bg | `#F2F4F7` |
| Header opsi bg | `#E8EEF5` |
| Border | `#D9E2EC`, single, size 6 |
| Total lebar tabel | 9360 twips |

Lebar kolom (twips):
- Metadata label : 2232 | value : 7128
- Opsi A–E : 792 | jawaban : 8568
- Kunci/Pembahasan : 2232 | value : 7128

### 3.2 Metode

| Metode | Fungsi |
|--------|--------|
| `generate(soalList)` | Sanity-check library `docx`, lalu memanggil `_buildDocument`. Mengembalikan `Promise<Blob>`. |
| `_buildDocument(soalList, lib)` | Merakit elemen dokumen: header, tabel metadata per soal, tabel opsi, dst. |

### 3.3 Alur `_buildDocument`
1. Destructure class `Document`, `Paragraph`, `TextRun`, `Table`, `TableRow`, `TableCell`, `Packer`.
2. Resolve enum (`AlignmentType`, `WidthType`, `BorderStyle`, `ShadingType`) dengan **fallback string literal** (kompatibilitas versi library).
3. Definisi helper: `mkBorder()`, `w(size)`, `shade(fill)`, `cell()`, `metaRow()`, `metaTable()`, `optTable()`, `para()`, `emptyPara()`.
4. Susun `children`:
   - **Header dokumen** (judul + subjudul + catatan).
   - **Per soal**: heading `SOAL N` (dengan `pageBreakBefore` untuk soal ke-2 dst.), tabel metadata, blok stimulus, blok pertanyaan, dan tabel pilihan jawaban.
5. Bangun `Document` dengan satu `section`, lalu `Packer.toBlob(doc)`.

### 3.4 Pengecualian & Penanganan Error
- Jika `window.docx` tidak ada → throw error *"Library docx belum termuat…"*.
- Jika export penting (`Document`, `Paragraph`, dst.) tidak ditemukan → throw error spesifik.
- Semua pembuatan dibungkus try/catch dengan pesan error yang informatif.

---

## 4. Urutan Load Script (index.html)

```
1. xlsx (CDN)          → untuk upload Excel
2. mammoth (CDN)       → untuk upload Word
3. js/docx.min.js (lokal) → untuk generate .docx
4. js/parsers.js
5. js/generator.js
6. js/app.js           → terakhir (menggunakan parser & generator)
```

`App.init()` dipanggil pada `DOMContentLoaded`, jadi aman meskipun skrip dimuat di akhir `body`.

---

Lanjut ke: [04-pengembangan.md](04-pengembangan.md) untuk panduan pengembangan.
