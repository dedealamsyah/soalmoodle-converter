# 02 — Format Input Soal

Panduan lengkap format input yang didukung aplikasi: **teks langsung**, **Excel (.xlsx)**, dan **Word (.docx)**. Semua input pada akhirnya diproses menjadi `SoalObject` (lihat [01-arsitektur.md](01-arsitektur.md)).

> **Prinsip nilai default:** Field metadata yang tidak disertakan pada suatu soal otomatis diisi dari **Nilai Default** yang diatur di langkah 1 (mapel, kelas, kompetensi, level, skor).

---

## 1. Input Teks Langsung

Parser utama: `SoalParser.parseText()` → `_splitTextIntoBlocks()` → `_parseSingleBlock()`.

### 1.1 Menentukan batas antar soal

Paraser email memisahkan teks menjadi blok-blok soal dengan **3 prioritas** (urutan pengecekan):

1. **Header `SOAL N`** — jika terdapat baris yang diawali `SOAL` + angka.
2. **Separator `---`** — baris berisi minimal 3 tanda hubung.
3. **Soal bernomor** — baris diawali angka + titik (`1.`), dengan syarat minimal salah satu bagian memiliki opsi A–E.

```text
SOAL 1
Kode: VDG-001
Mapel: Videografi
Kelas/Jurusan: XI DKV 4
Level: Sedang
Skor: 1

Pertanyaan:
Teks pertanyaan di sini...

A. Opsi A
B. Opsi B
C. Opsi C
D. Opsi D
E. Opsi E

Kunci: A
Pembahasan: Penjelasan...

---

SOAL 2
...
```

### 1.2 Field metadata yang dikenali

Parser membaca baris berformat `Key: value` (urutan bebas, case-insensitive):

| Key yang dikenali | Disimpan ke |
|-------------------|-------------|
| `Kode`, `Kode Soal` | `kodeSoal` |
| `Mapel`, `Mata Pelajaran` | `mapel` |
| `Kelas`, `Kelas/Jurusan`, `Jurusan` | `kelasJurusan` |
| `Kompetensi`, `Materi`, `Kompetensi/Materi`, `Topik` | `kompetensiMateri` |
| `Level`, `Kesulitan` | `level` (dinormalisasi) |
| `Skor`, `Bobot`, `Poin` | `skor` |

### 1.3 Header bagian

- `Pertanyaan:` → teks setelahnya menjadi isi `pertanyaan` (sampai baris kosong sebelum opsi).
- `Stimulus:` → teks setelahnya menjadi `stimulus`.
- `Pilihan Jawaban:` → menandai awal mode opsi.
- `Kunci:` / `Kunci Jawaban:` → diikuti huruf `A–E`.
- `Pembahasan:` / `Penjelasan:` / `Solusi:` → teks setelahnya menjadi `pembahasan`.

### 1.4 Opsi jawaban

Dikenali format awal baris:
- `A. teks`
- `A) teks`
- `(A) teks`

Baris lanjutan (continuation) opsi multi-baris akan digabung dengan spasi ke opsi aktif.

### 1.5 Normalisasi Level

| Input | Hasil |
|-------|-------|
| `Mudah`, `easy`, `L1`, `1` | `Mudah` |
| `Sedang`, `medium`, `L2`, `2` | `Sedang` |
| `Sulit`, `hard`, `L3`, `3` | `Sulit` |

---

## 2. Input Excel (.xlsx)

Parser utama: `SoalParser.parseExcel(rows, defaults)`.

- Baris **pertama** = header kolom (urutan bebas).
- Deteksi kolom otomatis berdasarkan kemiripan nama (normalisasi: huruf kecil, tanpa spasi/tanda baca).
- Baris kosong dilewati; `nomorUrut` diambil dari kolom `No` jika ada, jika tidak memakai indeks baris.

### 2.1 Alias nama kolom yang dikenali

| Kategori | Alias yang dikenali |
|----------|---------------------|
| No | `No`, `Nomor`, `Number`, `Urutan` |
| Kode | `Kode`, `KodeSoal`, `Code` |
| Mapel | `Mapel`, `MataPelajaran`, `Subject`, `Pelajaran`, `Matkul` |
| Kelas | `Kelas`, `KelasJurusan`, `Jurusan`, `Class` |
| Kompetensi | `Kompetensi`, `KompetensiMateri`, `Materi`, `Topik`, `Topic` |
| Level | `Level`, `Kesulitan`, `Difficulty`, `Tingkat` |
| Skor | `Skor`, `Bobot`, `Score`, `Nilai`, `Poin` |
| Stimulus | `Stimulus`, `Wacana`, `Bacaan`, `Teks` |
| Pertanyaan | `Pertanyaan`, `Soal`, `Question` |
| Opsi A–E | `OpsiA`, `PilihanA`, `OptionA`, `A` (dst. untuk B–E) |
| Kunci | `Kunci`, `KunciJawaban`, `Jawaban`, `Answer`, `Key`, `Jawab` |
| Pembahasan | `Pembahasan`, `Penjelasan`, `Explanation`, `Solusi`, `Alasan` |

> **⚠️ Kolom wajib** (untuk status `valid`): **Pertanyaan**, **A–E**, dan **Kunci**. Kolom lain opsional (fallback ke nilai default).

### 2.2 Contoh tabel Excel

| No | Kode | Mapel | Kelas | Pertanyaan | A | B | ... | Kunci | Pembahasan |
|----|------|-------|-------|------------|---|---|-----|-------|------------|
| 1 | VDG-001 | Videografi | XI DKV 4 | Sudut pengambilan... | Bird eye | High angle | ... | D | Low angle... |

---

## 3. Input Word (.docx)

Parser utama: `SoalParser.parseWordText(extractedText, defaults)`.

Alur:
1. Teks diekstrak dari `.docx` menggunakan library **Mammoth** (`mammoth.extractRawText`).
2. Artefak umum (tabs, spasi berlebih) dibersihkan.
3. Teks yang dihasilkan **delegasikan ke text parser** (`parseText`).

Artinya **format penulisan soal di Word mengikuti aturan yang sama** dengan input teks langsung (header `SOAL N`, `Kunci:`, opsi `A.`, dsb). Contoh pola yang dikenali:

```
1. Pertanyaan pertama...
A. Opsi A
B. Opsi B
...
Kunci: A

2. Pertanyaan kedua...
...
```

---

## 4. Aturan Validasi (Status Soal)

Fungsi `SoalParser.validateSoal()` menetapkan status:

| Kondisi | Status | Isu |
|---------|--------|-----|
| Pertanyaan kosong | `error` | "Pertanyaan kosong" |
| Ada opsi A–E kosong | `error` | "Opsi X kosong" |
| Kunci tidak valid (bukan A–E) | `error` | "Kunci jawaban tidak valid" |
| Mapel / Kelas tidak diisi | `warning` | "Mapel tidak diisi" / "Kelas/Jurusan tidak diisi" |
| Tidak ada isu di atas | `valid` | — |

---

Lanjut ke: [03-modul-js.md](03-modul-js.md) untuk detail teknis tiap modul.
