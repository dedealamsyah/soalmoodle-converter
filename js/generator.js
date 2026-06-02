/**
 * GENERATOR.JS — docx v7.8.2 compatible
 *
 * Menghasilkan file .docx yang IDENTIK dengan
 * template-input-soal-pg-sederhana.docx (Template Soal Moodle).
 *
 * Spesifikasi warna dari XML asli:
 *   Judul dokumen  : #0B2545
 *   Heading soal   : #1F4D78  (bold, 14pt)
 *   Label bg       : #F2F4F7
 *   Header opsi bg : #E8EEF5
 *   Border         : #D9E2EC  (single, size 6)
 *
 * Lebar kolom (twips, total 9360):
 *   Metadata label : 2232  |  value : 7128
 *   Opsi A-E       :  792  |  jawaban : 8568
 *   Kunci/Pembahasan: 2232 |  value : 7128
 */

const DocxGenerator = {

  /* =========================================================
     PUBLIC: generate(soalList) → Promise<Blob>
     ========================================================= */
  async generate(soalList) {
    /* --- sanity check ------------------------------------ */
    if (!window.docx) {
      throw new Error(
        'Library docx belum termuat. Pastikan koneksi internet aktif dan refresh halaman.'
      );
    }

    const lib = window.docx;

    // Verify essential exports exist
    const needed = ['Document', 'Paragraph', 'TextRun', 'Table',
                    'TableRow', 'TableCell', 'Packer'];
    for (const name of needed) {
      if (typeof lib[name] !== 'function') {
        throw new Error(
          `docx.${name} tidak ditemukan. Library mungkin gagal dimuat — refresh halaman.`
        );
      }
    }

    console.log('[Generator] docx library OK, mulai generate…');

    try {
      const blob = await this._buildDocument(soalList, lib);
      console.log('[Generator] Berhasil! Blob size:', blob.size);
      return blob;
    } catch (err) {
      console.error('[Generator] ERROR:', err);
      throw new Error('Gagal membuat dokumen: ' + err.message);
    }
  },

  /* =========================================================
     PRIVATE: _buildDocument
     ========================================================= */
  async _buildDocument(soalList, lib) {
    const {
      Document, Paragraph, TextRun, Table, TableRow, TableCell, Packer
    } = lib;

    /* --- resolve enum values with string-literal fallback -- */
    const ALIGN  = lib.AlignmentType  || {};
    const WIDTH  = lib.WidthType      || {};
    const BORDER = lib.BorderStyle    || {};
    const SHADE  = lib.ShadingType    || {};

    const AL_CENTER  = ALIGN.CENTER   || 'center';
    const WT_DXA     = WIDTH.DXA      || 'dxa';
    const BS_SINGLE  = BORDER.SINGLE  || 'single';
    const ST_CLEAR   = SHADE.CLEAR    || 'clear';

    /* --- shared border definition (one per cell, no reuse) - */
    const mkBorder = () => ({
      top:    { style: BS_SINGLE, size: 6, color: 'D9E2EC' },
      bottom: { style: BS_SINGLE, size: 6, color: 'D9E2EC' },
      left:   { style: BS_SINGLE, size: 6, color: 'D9E2EC' },
      right:  { style: BS_SINGLE, size: 6, color: 'D9E2EC' },
    });

    /* --- width helpers ------------------------------------ */
    const w = (size) => ({ size, type: WT_DXA });

    /* --- shading helpers ---------------------------------- */
    const shade = (fill) => ({ fill, type: ST_CLEAR, color: 'auto' });

    /* -------------------------------------------------------
       CELL FACTORY
       ------------------------------------------------------- */
    const cell = (widthTwips, text, opts = {}) => {
      const cellOpts = {
        width: w(widthTwips),
        borders: mkBorder(),
        children: [
          new Paragraph({
            alignment: opts.center ? AL_CENTER : undefined,
            children: [
              new TextRun({
                text: String(text || ''),
                bold: !!opts.bold,
                italics: !!opts.italics,
                color: opts.color || undefined,
              })
            ]
          })
        ]
      };
      if (opts.fill) cellOpts.shading = shade(opts.fill);
      return new TableCell(cellOpts);
    };

    /* -------------------------------------------------------
       TABLE FACTORY — metadata (label | value), 9360 total
       ------------------------------------------------------- */
    const metaRow = (label, value) => new TableRow({
      children: [
        cell(2232, label, { bold: true, fill: 'F2F4F7' }),
        cell(7128, value),
      ]
    });

    const metaTable = (soal) => new Table({
      width: w(9360),
      rows: [
        metaRow('Kode Soal',         soal.kodeSoal         || ''),
        metaRow('Mapel',             soal.mapel             || ''),
        metaRow('Kelas/Jurusan',     soal.kelasJurusan      || ''),
        metaRow('Kompetensi/Materi', soal.kompetensiMateri  || ''),
        metaRow('Level',             soal.level             || 'Mudah / Sedang / Sulit'),
        metaRow('Skor',              String(soal.skor != null ? soal.skor : 1)),
      ]
    });

    /* -------------------------------------------------------
       TABLE FACTORY — options (Opsi A-E + Kunci + Pembahasan)
       ------------------------------------------------------- */
    const optTable = (soal) => {
      const rows = [];

      // Header row
      rows.push(new TableRow({
        children: [
          cell(792,  'Opsi',    { bold: true, fill: 'E8EEF5' }),
          cell(8568, 'Jawaban', { bold: true, fill: 'E8EEF5' }),
        ]
      }));

      // Option rows A–E
      for (const opt of ['A', 'B', 'C', 'D', 'E']) {
        rows.push(new TableRow({
          children: [
            cell(792,  opt,                              { bold: true, center: true }),
            cell(8568, soal.pilihanJawaban[opt] || ''),
          ]
        }));
      }

      // Kunci
      rows.push(new TableRow({
        children: [
          cell(2232, 'Kunci',      { bold: true, fill: 'F2F4F7' }),
          cell(7128, soal.kunci    || 'A / B / C / D / E'),
        ]
      }));

      // Pembahasan
      rows.push(new TableRow({
        children: [
          cell(2232, 'Pembahasan', { bold: true, fill: 'F2F4F7' }),
          cell(7128, soal.pembahasan || ''),
        ]
      }));

      return new Table({ width: w(9360), rows });
    };

    /* -------------------------------------------------------
       PARAGRAPH HELPERS
       ------------------------------------------------------- */
    const para = (text, opts = {}) => new Paragraph({
      alignment: opts.center ? AL_CENTER : undefined,
      pageBreakBefore: opts.pageBreak || false,
      children: [
        new TextRun({
          text: String(text || ''),
          bold:    opts.bold    || false,
          italics: opts.italic  || false,
          color:   opts.color   || undefined,
          size:    opts.size    || undefined,
        })
      ]
    });

    const emptyPara = () => new Paragraph({ children: [] });

    /* -------------------------------------------------------
       ASSEMBLE DOCUMENT CHILDREN
       ------------------------------------------------------- */
    const children = [];

    /* Document header */
    children.push(
      para('TEMPLATE INPUT SOAL PILIHAN GANDA',
           { center: true, bold: true, color: '0B2545', size: 36 }),
      para('Template Soal Moodle - Guru cukup mengisi format di bawah ini',
           { center: true, italic: true }),
      para('Catatan singkat: gambar boleh langsung ditempel di dokumen. Untuk rumus, gunakan Equation Word jika rumus cukup kompleks.')
    );

    /* Soal blocks */
    soalList.forEach((soal, idx) => {
      const num = idx + 1;
      const addBreak = idx > 0;  // page break before soal 2, 3, ...

      console.log(`[Generator] Building SOAL ${num}…`);

      // Heading SOAL N
      children.push(
        para(`SOAL ${num}`, { pageBreak: addBreak, bold: true, color: '1F4D78', size: 28 })
      );

      // Metadata table
      children.push(metaTable(soal));

      // Stimulus
      children.push(para('Stimulus / Gambar / Tabel:'));
      if (soal.stimulus && soal.stimulus.trim()) {
        children.push(para(soal.stimulus));
      } else {
        children.push(para(
          '[Tempel stimulus, gambar, tabel, atau grafik di sini. Gambar cukup ditempel langsung di dokumen.]',
          { italic: true, color: '666666' }
        ));
      }

      // Pertanyaan
      children.push(para('Pertanyaan:'));
      children.push(para(soal.pertanyaan || ''));
      children.push(emptyPara());

      // Pilihan Jawaban
      children.push(para('Pilihan Jawaban:'));
      children.push(optTable(soal));
    });

    /* Build Document */
    const doc = new Document({
      sections: [{ children }]
    });

    return Packer.toBlob(doc);
  }
};
