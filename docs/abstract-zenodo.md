# Abstrak — SoalMoodle Converter

**Judul:** Rancang Bangun Aplikasi Web Client-Side SoalMoodle Converter untuk Otomatisasi Konversi Kartu Soal Pilihan Ganda ke Template Soal Moodle

## Abstrak (Bahasa Indonesia)

Pengelolaan soal ujian pada platform Moodle sering terkendala oleh proses input naskah soal yang manual dan tidak seragam, sehingga rawan kesalahan dan tidak efisien untuk jumlah soal yang besar. Penelitian ini merancang dan membangun **SoalMoodle Converter**, aplikasi web berbasis *client-side* (serverless) yang mengonversi kartu soal pilihan ganda dari tiga sumber masukan—teks langsung, Excel (.xlsx), dan Word (.docx)—menjadi template soal Moodle (.docx) secara otomatis. Aplikasi dibangun dengan HTML, CSS, dan JavaScript murni, memanfaatkan pustaka `docx` (bundle lokal), `SheetJS`, dan `Mammoth`, dengan alur tiga tahap: input, pratayap/validasi per soal, dan unduh. Hasil pengujian menunjukkan seluruh format masukan berhasil dikonversi dengan struktur keluaran identik terhadap template Moodle resmi. Status validasi (valid/peringatan/error) membantu meminimalkan kesalahan input. Pendekatan *offline-first* menjaga berkas tetap lokal dan memungkinkan pengoperasian tanpa server, sehingga aplikasi ringan, gratis, dan mudah diadopsi.

**Kata kunci:** Moodle; soal pilihan ganda; konversi dokumen; aplikasi client-side; otomatisasi input soal; e-learning

---

## Abstract (English)

Question management on the Moodle platform is often hindered by a manual and non-uniform question-entry process, which is error-prone and inefficient for large numbers of questions. This research designs and builds **SoalMoodle Converter**, a client-side (serverless) web application that automatically converts multiple-choice questions from three input sources—plain text, Excel (.xlsx), and Word (.docx)—into the Moodle question template (.docx). The application is built with plain HTML, CSS, and JavaScript, leveraging the `docx` (local bundle), `SheetJS`, and `Mammoth` libraries, following a three-stage workflow: input, preview/per-question validation, and download. Test results show all input formats are successfully converted with an output structure identical to the official Moodle template. Validation status (valid/warning/error) helps minimize input errors. The offline-first approach keeps files local and enables operation without a server, making the application lightweight, free, and easy to adopt.

**Keywords:** Moodle; multiple choice questions; document conversion; client-side application; automated question input; e-learning
