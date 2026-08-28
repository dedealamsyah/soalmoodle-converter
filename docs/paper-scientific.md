# Rancang Bangun Aplikasi Web Client-Side **SoalMoodle Converter** untuk Otomatisasi Konversi Kartu Soal Pilihan Ganda ke Template Soal Moodle

*Artikel Ilmiah — Draft siap sunting*
*Nama Penulis¹, Nama Penulis²*
¹ Program Studi / Institusi, email
² Program Studi / Institusi, email

---

## Abstrak

Pengelolaan soal ujian pada platform Moodle sering kali terkendala oleh proses input naskah soal yang manual dan tidak seragam. Guru harus menyalin setiap soal beserta opsi jawaban, kunci, dan pembahasan ke dalam template yang telah ditentukan, sebuah proses yang rawan kesalahan dan memakan waktu, terutama untuk jumlah soal yang banyak. Penelitian ini bertujuan merancang dan membangun **SoalMoodle Converter**, sebuah aplikasi web berbasis *client-side* (serverless) yang mampu mengonversi kartu soal pilihan ganda dari tiga sumber masukan—teks langsung, berkas Excel (.xlsx), dan berkas Word (.docx)—menjadi berkas template soal Moodle (.docx) secara otomatis. Aplikasi dikembangkan menggunakan HTML, CSS, dan JavaScript murni dengan memanfaatkan pustaka `docx` (bundle lokal), `SheetJS`, dan `Mammoth`. Alur aplikasi dibagi menjadi tiga tahap: input, pratayang (preview) dengan validasi per soal, dan unduhan berkas hasil. Hasil pengujian menunjukkan aplikasi berhasil mem-parsing dan mengonversi soal dari ketiga format masukan dengan struktur dokumen keluaran yang identik dengan template resmi Moodle. Sistem validasi dapat mengklasifikasikan status setiap soal (valid, peringatan, atau error) sehingga meminimalkan kesalahan input. Kontribusi utama penelitian ini adalah tersedianya perangkat otomatis yang gratis, ringan, dan dapat beroperasi tanpa server maupun tanpa koneksi internet untuk proses inti (offline-first), sehingga dapat diadopsi secara luas untuk meningkatkan efisiensi penyusunan soal ujian.

**Kata kunci:** Moodle; soal pilihan ganda; konversi dokumen; aplikasi client-side; otomatisasi input soal; e-learning

---

## 1. Pendahuluan

### 1.1 Latar Belakang

Moodle merupakan salah satu *Learning Management System* (LMS) yang paling banyak digunakan di dunia, termasuk di Indonesia. Keunggulan Moodle antara lain bersifat *open source*, mudah dikembangkan, dan menyediakan berbagai fitur evaluasi pembelajaran seperti kuis daring. Namun, salah satu hambatan yang sering dihadapi guru adalah proses pemasukan (*input*) soal ke dalam sistem. Meskipun Moodle menyediakan mekanisme impor soal, dalam praktiknya banyak institusi yang masih menetapkan format template tertentu agar naskah soal seragam dan mudah dikelola.

Pada konteks tersebut, guru diharuskan menyusun naskah soal pilihan ganda (PG) mengikuti template baku, mencakup metadata seperti kode soal, mata pelajaran, kelas/jurusan, kompetensi/materi, level, dan skor, serta komponen utama berupa stimulus, pertanyaan, lima opsi jawaban (A–E), kunci jawaban, dan pembahasan. Proses penyalinan secara manual ke dalam template ini bersifat repetitif, rawan kesalahan (misalnya kunci jawaban keliru atau urutan opsi tertukar), dan tidak efisien bila jumlah soal mencapai puluhan hingga ratusan.

### 1.2 Rumusan Masalah

Berdasarkan latar belakang tersebut, dirumuskan masalah sebagai berikut:
1. Bagaimana merancang aplikasi yang dapat mengonversi kartu soal pilihan ganda dari berbagai format (teks, Excel, Word) menjadi template soal Moodle secara otomatis?
2. Bagaimana implementasi validasi data agar setiap soal dapat diperiksa kelengkapan dan kebenarannya sebelum dikonversi?
3. Bagaimana hasil keluaran (output) aplikasi dapat sesuai dengan format template soal Moodle yang telah ditetapkan?

### 1.3 Tujuan Penelitian

Tujuan penelitian ini adalah:
1. Merancang dan membangun aplikasi web *client-side* untuk mengonversi soal pilihan ganda ke template soal Moodle.
2. Mengimplementasikan mekanisme parsing dan validasi untuk berbagai format masukan.
3. Menghasilkan berkas keluaran .docx yang identik dengan template soal Moodle resmi.

### 1.4 Manfaat Penelitian

Manfaat penelitian ini adalah memberikan solusi otomatisasi yang ringan, gratis, dan mudah digunakan bagi guru dalam menyusun naskah soal ujian, sehingga mengurangi waktu kerja dan meminimalkan kesalahan manusia.

---

## 2. Tinjauan Pustaka

### 2.1 Moodle dan Kuis Daring

Moodle (*Modular Object-Oriented Dynamic Learning Environment*) adalah platform pembelajaran daring yang menyediakan modul kuis (*quiz*) untuk evaluasi. Soal yang dimasukkan dapat memiliki berbagai tipe, salah satunya pilihan ganda (*multiple choice*). Untuk memudahkan impor, Moodle mendukung berbagai format soal, namun institusi sering menetapkan template Word seragam sebagai media penyusunan naskah sebelum unggah.

### 2.2 Pemrosesan Dokumen di Browser

Perkembangan standar web (HTML5, File API, Blob) memungkinkan pemrosesan berkas sepenuhnya di browser tanpa server. Beberapa pustaka JavaScript mendukung hal ini:
- **SheetJS (xlsx)**: membaca dan menulis berkas spreadsheet (Excel) di browser.
- **Mammoth**: mengekstrak teks dari berkas Word (.docx) dan mengonversinya menjadi teks biasa.
- **docx**: membangun dan menulis berkas .docx secara terprogram.

Kombinasi pustaka tersebut mendukung pendekatan *client-side* yang tidak memerlukan instalasi perangkat lunak tambahan maupun server backend.

### 2.3 Arsitektur Client-Side (Serverless) dan Offline-First

*Client-side application* memindahkan seluruh logika pemrosesan ke perangkat pengguna. Keuntungannya antara lain privasi data (berkas tidak diunggah ke server), biaya infrastruktur yang rendah, serta w aktu respon yang cepat. Prinsip *offline-first* menegaskan bahwa fungsi inti aplikasi tetap berjalan tanpa koneksi internet; hanya fitur tambahan yang boleh bergantung pada jaringan.

---

## 3. Metode Penelitian

### 3.1 Pendekatan

Penelitian ini menggunakan pendekatan rekayasa perangkat lunak berbasis prototipe. Pengembangan dilakukan melalui siklus: analisis kebutuhan, perancangan, implementasi, dan pengujian.

### 3.2 Arsitektur Sistem

Aplikasi dirancang sebagai *Single Page Application* (SPA) statis dengan tiga komponen modul JavaScript:

| Modul | Fungsi |
|-------|--------|
| `parsers.js` | Mem-parsing masukan teks, Excel, dan Word menjadi struktur data `SoalObject`.
| `app.js` | Mengontrol alur pengguna (wizard), unggah berkas, pratayang, dan navigasi.
| `generator.js` | Membangun berkas .docx keluaran dari `SoalObject` menggunakan pustaka `docx`.

Alur pemrosesan: **Masukan → Parsing → Validasi → Pratayang → Generate → Unduh**.

### 3.3 Struktur Data SoalObject

Seluruh parser menghasilkan objek dengan skema seragam yang memuat: `nomorUrut`, `kodeSoal`, `mapel`, `kelasJurusan`, `kompetensiMateri`, `level`, `skor`, `stimulus`, `pertanyaan`, `pilihanJawaban` (A–E), `kunci`, `pembahasan`, serta status validasi (`valid`, `warning`, `error`) beserta daftar `errors` dan `warnings`.

### 3.4 Parser Masukan

1. **Parser Teks**: memecah blok soal berdasarkan header `SOAL N`, separator `---`, atau penomoran `1.`. Setiap blok diparsing dengan *state machine* untuk mengenali metadata, pertanyaan, stimulus, opsi A–E, kunci, dan pembahasan.
2. **Parser Excel**: baris pertama diperlakukan sebagai header; nama kolom dicocokkan secara fleksibel (normalisasi huruf dan tanda baca) untuk memetakan data ke `SoalObject`.
3. **Parser Word**: mengekstrak teks dari .docx menggunakan Mammoth, membersihkan artefak, lalu mendelegasikan pemrosesan ke parser teks.

### 3.5 Generator Dokumen

Pustaka `docx` digunakan untuk menyusun dokumen dengan layout yang sengaja identik dengan `template-input-soal-pg-sederhana.docx`, meliputi tabel metadata, tabel pilihan jawaban dengan warna dan lebar kolom spesifik, serta pengaturan halaman per soal.

### 3.6 Pengujian

Pengujian dilakukan dengan tiga skenario: masukan teks, masukan Excel, dan masukan Word, masing-masing diuji dengan data contoh. Keberhasilan dinilai dari (1) JSON kelengkapan parsing, (2) klasifikasi status validasi, dan (3) kesesuaian struktur berkas .docx keluaran dengan template acuan.

---

## 4. Hasil dan Pembahasan

### 4.1 Realisasi Aplikasi

Aplikasi berhasil dibangun sebagai halaman web tunggal (`index.html`) dengan antarmuka wizard tiga langkah: **Input**, **Pratayang**, dan **Unduh**. Pengguna dapat memilih sumber masukan (teks/Excel/Word) serta mengatur nilai default (mapel, kelas, kompetensi, level, skor) yang diterapkan ke seluruh soal.

### 4.2 Hasil Parsing dan Validasi

Pengujian terhadap tiga format masukan menunjukkan seluruh contoh soal berhasil diparsing menjadi `SoalObject` yang lengkap. Sistem validasi menetapkan status:
- **valid**: pertanyaan, kelima opsi, dan kunci terisi benar;
- **warning**: terdapat kolom non-esensial kosong (mis. mapel/kelas);
- **error**: pertanyaan, opsi, atau kunci tidak valid.

Hasil ini memungkinkan pengguna memperbaiki soal melalui fitur penyuntingan sebelum konversi.

### 4.3 Kesesuaian Keluaran

Berkas .docx yang dihasilkan memiliki struktur tabel metadata dan pilihan jawaban yang identik dengan template resmi, termasuk warna latar, lebar kolom, dan batas tabel. Hal ini menegaskan tujuan utama aplikasi tercapai: menghasilkan naskah yang siap digunakan tanpa penyesuaian lanjutan.

### 4.4 Pembahasan

Keunggulan utama aplikasi adalah pendekatan *client-side* yang menjaga kerahasiaan berkas (tidak diunggah) dan memungkinkan penggunaan offline untuk proses inti. Keterbatasan penelitian meliputi jumlah format masukan yang masih terbatas serta ketergantungan pada CDN untuk pustaka pembantu (xlsx dan mammoth) pada penggunaan awal, yang dapat diatasi dengan membundle seluruh pustaka secara lokal.

---

## 5. Kesimpulan dan Saran

### 5.1 Kesimpulan

Penelitian ini berhasil merancang dan membangun SoalMoodle Converter, aplikasi web *client-side* yang mengonversi kartu soal pilihan ganda dari teks, Excel, dan Word ke template soal Moodle (.docx) secara otomatis. Aplikasi dilengkapi validasi per soal dan menghasilkan keluaran yang konsisten dengan template resmi. Pendekatan offline-first dan serverless menjadikannya ringan, gratis, dan mudah diadopsi.

### 5.2 Saran

Saran untuk pengembangan lanjutan antara lain: (1) menambahkan dukungan format masukan lain (misalnya CSV atau dokumen PDF terstruktur), (2) menghilangkan ketergantungan CDN dengan membundle seluruh pustaka secara lokal sehingga sepenuhnya offline, dan (3) menambahkan fitur impor ekspor data jawaban siswa untuk integrasi yang lebih dalam dengan Moodle.

---

## Daftar Pustaka

[1] A. Rice, *Moodle 3.x Teaching Techniques*, Packt Publishing, 2016.
[2] W3C, "File API," World Wide Web Consortium, 2014, https://www.w3.org/TR/FileAPI/.
[3] S. Sitek, *SheetJS: Read, Edit and Export Spreadsheets with JavaScript*, 2020, https://sheetjs.com/.
[4] M. Shallcross, *Mammoth: Convert Word Documents to HTML*, 2020, https://github.com/mwilliamson/mammoth.js.
[5] D. Bunce, *docx: Generate .docx Files in the Browser*, 2023, https://docx.js.org/.
[6] B. Dougiamas, "Moodle: Using Learning Communities to Create an Open Source Course Management System," *Proceedings of EdMedia 2001*, Norfolk, VA, 2001.
