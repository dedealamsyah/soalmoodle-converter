# 01 — Arsitektur Aplikasi

Dokumen ini menjelaskan arsitektur **SoalMoodle Converter**, termasuk arsitektur umum, alur kerja pengguna, alur data, dan prinsip desain.

---

## 1. Gambaran Arsitektur

Aplikasi ini adalah **Single Page Application (SPA) client-side yang sepenuhnya statis**:

```
┌────────────────────────────────────────────────────────────┐
│                        BROWSER                             │
│                                                            │
│   index.html  ──▶  UI (markup & styling)                   │
│     │                      css/style.css                    │
│     │                                                       │
│     ├──▶ js/parsers.js    Mem-parse input → SoalObject[]    │
│     ├──▶ js/app.js        Controller (wizard, event, state) │
│     └──▶ js/generator.js  SoalObject[] → Blob .docx          │
│              │                                               │
│              └──▶ js/docx.min.js  (library docx v7.8.2)     │
│                                                              │
│   CDN (butuh internet):                                      │
│     • xlsx.full.min.js    (membaca file Excel)               │
│     • mammoth.browser.js  (ekstrak teks dari Word)           │
└────────────────────────────────────────────────────────────┘
```

Tidak ada server, basis data, atau API eksternal. Semua pemrosesan terjadi di mesin pengguna.

---

## 2. Alur Kerja Pengguna (Wizard 3 Langkah)

UI disusun sebagai wizard dengan indikator langkah di bagian atas (`step-indicator-1..3`).

### Langkah 1 — Input Soal
Pengguna memilih salah satu dari 3 tab sumber input:

| Tab | Input | Parser yang dipakai |
|-----|-------|---------------------|
| ✏️ **Ketik Langsung** | Teks di `textarea` | `SoalParser.parseText()` |
| 📊 **Upload Excel** | File `.xlsx` / `.xls` | `XLSX.read()` → `SoalParser.parseExcel()` |
| 📄 **Upload Word** | File `.docx` | `mammoth.extractRawText()` → `SoalParser.parseWordText()` |

Di halaman ini juga pengguna mengisi **Nilai Default** (mapel, kelas, kompetensi, level, skor) yang akan diterapkan ke semua soal yang kekurangan field tersebut.

Tombol **"Parse & Preview Soal"** memicu fungsi `App._handleParse()`.

### Langkah 2 — Preview
Menampilkan hasil parsing setiap soal dalam kartu yang bisa di-expand:

- **Statistik bar**: jumlah Valid ✅ / Peringatan ⚠️ / Error ❌.
- **Kartu soal**: menampilkan semua metadata, stimulus, pertanyaan, opsi A–E (kunci ditandai hijau), pembahasan, dan daftar isu (error/warning).
- **Aksi per soal**: tombol **Edit ✏️** (membuka modal edit) dan **Hapus 🗑️**.
- Tombol **Generate .docx** memicu `App._handleGenerate()`.

### Langkah 3 — Download
Setelah generate berhasil, pengguna bisa:
- **Download .docx** → mengunduh file hasil (nama file dinamis: `[mapel]_[kelas]_ASAT_2526.docx`).
- **Konversi Lagi** → reset dan kembali ke langkah 1 (`App._handleReset()`).

### Flow antar langkah
```
App._goToStep(n)  →  toggle class 'active' pada step-section & step indicator
Window scroll-to-top otomatis setiap pindah langkah.
```

---

## 3. Alur Data

### 3.1 Objek `SoalObject` (struktur sentral)
Semua parser menghasilkan objek dengan bentuk yang sama, dipakai bersama oleh preview dan generator:

```js
{
  nomorUrut: 1,                  // Nomor urut / SOAL N
  kodeSoal: 'VDG-001',           // Kode soal (opsional)
  mapel: 'Videografi',           // Mata pelajaran
  kelasJurusan: 'XI DKV 4',      // Kelas / jurusan
  kompetensiMateri: 'Teknik',    // Kompetensi / materi
  level: 'Sedang',               // Mudah | Sedang | Sulit
  skor: 1,                       // Skor/bobot
  stimulus: '',                  // Teks stimulus (opsional)
  pertanyaan: 'Teks soal...',    // Pertanyaan utama
  pilihanJawaban: { A:'', B:'', C:'', D:'', E:'' },  // Opsi jawaban
  kunci: 'D',                    // Kunci jawaban A–E
  pembahasan: '',                // Pembahasan (opsional)
  // Status & validasi:
  status: 'valid',               // 'valid' | 'warning' | 'error'
  errors: [],                    // Daftar isu fatal
  warnings: []                   // Daftar isu non-fatal
}
```

### 3.2 Diagram alur data

```
[Input]                       [Parsing]                    [Validasi]
Teks langsung ──┐
Excel (.xlsx) ──┼────▶ SoalParser ──▶ SoalObject[] ──▶ validateSoal()
Word (.docx) ──┘                                          → set status
                                                              │
            [Preview]           [Generate]                    │
        renderPreview()     DocxGenerator.generate() ◀────────┤
        (edit/delete)            │  (skip status 'error')     │
                                 ▼                             │
                           Blob .docx ──▶ Download             │
                                 ▲                             │
                                 └── 'error' soal dilewati ────┘
```

---

## 4. Komponen Kunci

### 4.1 `App` object (`js/app.js`)
Kontainer utama yang menyimpan state aplikasi **secara global**:

```js
const App = {
  soalList: [],          // Array SoalObject hasil parsing (state inti)
  generatedBlob: null,   // Blob .docx hasil generate
  downloadFilename,      // Nama file unduhan dinamis
  currentTab,            // 'text' | 'excel' | 'word'
  editingIndex,          // index soal yang sedang diedit (-1 = none)
  excelFile, wordFile    // File yang dipilih user
}
```

### 4.2 `SoalParser` object (`js/parsers.js`)
Berisi seluruh logika parsing & validasi. Terbagi menjadi:
- **Text Parser** — `parseText()`, `_splitTextIntoBlocks()`, `_parseSingleBlock()`.
- **Excel Parser** — `parseExcel()` dengan deteksi kolom otomatis.
- **Word Parser** — `parseWordText()` (delegasi ke text parser).
- **Bantuan** — `normalizeLevel()`, `validateSoal()`, `createSoal()`.

### 4.3 `DocxGenerator` object (`js/generator.js`)
Membangun dokumen `.docx` memakai library `docx`. Struktur output sengaja dibuat **identik** dengan `template-input-soal-pg-sederhana.docx` (dianalisis dari `extracted_template/word/document.xml`). Template tersimpan di `docs/template-input-soal-pg-sederhana.docx`.

---

## 5. Prinsip Desain

1. **Offline-first**: inti (parsing + generate) bekerja tanpa internet; hanya library pembantu (xlsx, mammoth) yang via CDN.
2. **Tanpa dependency berat**: hanya `docx` sebagai dependency npm; sisanya di-bundle lokal / CDN.
3. **Validasi per-soal**: setiap soal di-validasi dan diberi status, sehingga pengguna bisa lihat mana yang perlu diperbaiki sebelum generate.
4. **Nilai default fleksibel**: metadata default diterapkan hanya jika soal tidak memilikinya — memudahkan pengisian massal.
5. **UI wizard**: memecah alur kompleks menjadi langkah-langkah yang jelas (Input → Preview → Download).

---

Lanjut ke: [02-input-format.md](02-input-format.md) untuk format input lengkap.
