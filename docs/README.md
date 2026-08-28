# 📚 Documentation — SoalMoodle Converter

Dokumentasi lengkap aplikasi **SoalMoodle Converter**, tools berbasis web untuk mengonversi kartu soal pilihan ganda menjadi template soal Moodle (`.docx`) secara otomatis.

## Isi Dokumentasi

| Dokumen | Deskripsi |
|---------|-----------|
| [**01-arsitektur.md**](01-arsitektur.md) | Arsitektur aplikasi, alur kerja (wizard 3 langkah), dan alur data. |
| [**02-input-format.md**](02-input-format.md) | Panduan lengkap format input: teks langsung, Excel (.xlsx), dan Word (.docx). |
| [**03-modul-js.md**](03-modul-js.md) | Dokumentasi teknis tiap modul JavaScript (`app.js`, `parsers.js`, `generator.js`). |
| [**04-pengembangan.md**](04-pengembangan.md) | Panduan pengembangan: menjalankan, struktur direktori, cara menambahkan fitur, dan strategi offline. |
| [**paper-scientific.md**](paper-scientific.md) | Artikel ilmiah lengkap (Bahasa Indonesia) tentang aplikasi, siap sunting untuk publikasi. |
| [**abstract-zenodo.md**](abstract-zenodo.md) | Abstrak (ID & EN) untuk metadata publikasi di Zenodo. |

---

## Ringkasan Singkat

- **Jenis:** Aplikasi web client-side / serverless (berjalan penuh di browser, tanpa backend).
- **Fungsi utama:** Mengubah input soal PG (teks / Excel / Word) menjadi file `.docx` yang mengikuti format *Template Soal Moodle*.
- **Input didukung:** Teks langsung, file Excel (`.xlsx`), file Word (`.docx`).
- **Output:** File `.docx` dengan layout tabel metadata + opsi A–E + kunci + pembahasan, identik dengan template resmi.
- **Dependensi library:**
  - `docx` v7.8.2 → dibundle lokal di `js/docx.min.js` (offline).
  - `SheetJS (xlsx)` → dimuat via CDN (butuh internet saat pertama kali).
  - `Mammoth` → dimuat via CDN (butuh internet saat pertama kali).

---

## 📁 Struktur Proyek

```
soalmoodle-converter/
├── index.html                  # Halaman UI utama (3 langkah: Input → Preview → Download)
├── package.json                # Konfigurasi npm (dependensi docx v7.8.2)
├── package-lock.json
├── README.md                   # Panduan penggunaan & pengembangan (root)
├── css/
│   └── style.css               # Styling Premium Dark Theme
├── js/
│   ├── app.js                  # Kontrol alur wizard, upload, preview, edit, download
│   ├── parsers.js              # Logika parsing teks / Excel / Word → array SoalObject
│   ├── generator.js            # Logika pembuatan .docx dengan library docx.js
│   └── docx.min.js             # Bundle docx v7.8.2 lokal (mode offline)
├── extracted_template/         # Hasil unzip template (analisis XML asli)
│   ├── word/                   # document.xml, styles.xml, dll. — referensi format
│   ├── [Content_Types].xml
│   └── docProps/
└── docs/
    ├── README.md               # Dokumentasi ini
    ├── 01-arsitektur.md
    ├── 02-input-format.md
    ├── 03-modul-js.md
    ├── 04-pengembangan.md
    ├── template-input-soal-pg-sederhana.docx   # File referensi template soal Moodle
    ├── template-input-soal-pg-sederhana.zip    # Versi zip dari template
    └── sample/
        └── format soal videografi awal_1.docx   # Contoh dokumen hasil format
```

> **Catatan:** Subfolder `docs/` ini berisi dokumentasi lengkap. Mulailah dari [01-arsitektur.md](01-arsitektur.md) untuk gambaran besar, atau [04-pengembangan.md](04-pengembangan.md) jika ingin berkontribusi kode.
