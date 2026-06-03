/**
 * PARSERS.JS
 * Parser untuk format teks, Excel (.xlsx), dan Word (.docx)
 * Mengonversi berbagai format input menjadi array SoalObject
 */

const SoalParser = {

  // ── Normalize level string ──────────────────────────────
  normalizeLevel(val) {
    if (!val) return null;
    const v = String(val).toLowerCase().trim();
    if (v.includes('mudah') || v.includes('easy') || v === 'l1' || v === '1') return 'Mudah';
    if (v.includes('sulit') || v.includes('hard') || v === 'l3' || v === '3') return 'Sulit';
    if (v.includes('sedan') || v.includes('medium') || v === 'l2' || v === '2') return 'Sedang';
    return null;
  },

  // ── Validate soal object ────────────────────────────────
  validateSoal(soal) {
    const errors = [];
    const warnings = [];

    if (!soal.pertanyaan || !soal.pertanyaan.trim()) {
      errors.push('Pertanyaan kosong');
    }

    const opts = ['A', 'B', 'C', 'D', 'E'];
    const missingOpts = opts.filter(o => !soal.pilihanJawaban[o] || !soal.pilihanJawaban[o].trim());
    if (missingOpts.length > 0) {
      errors.push(`Opsi ${missingOpts.join(', ')} kosong`);
    }

    if (!soal.kunci || !opts.includes(soal.kunci)) {
      errors.push('Kunci jawaban tidak valid (harus A/B/C/D/E)');
    }

    if (!soal.mapel) warnings.push('Mapel tidak diisi');
    if (!soal.kelasJurusan) warnings.push('Kelas/Jurusan tidak diisi');

    if (errors.length > 0) soal.status = 'error';
    else if (warnings.length > 0) soal.status = 'warning';
    else soal.status = 'valid';

    soal.errors = errors;
    soal.warnings = warnings;
    return soal;
  },

  // ── Create empty soal with defaults ────────────────────
  createSoal(num, defaults = {}) {
    return {
      nomorUrut: num,
      kodeSoal: '',
      mapel: defaults.mapel || '',
      kelasJurusan: defaults.kelas || '',
      kompetensiMateri: defaults.kompetensi || '',
      level: defaults.level || 'Sedang',
      skor: defaults.skor || 1,
      stimulus: '',
      pertanyaan: '',
      pilihanJawaban: { A: '', B: '', C: '', D: '', E: '' },
      kunci: '',
      pembahasan: '',
      status: 'valid',
      errors: [],
      warnings: []
    };
  },

  // ════════════════════════════════════════════════════════
  // TEXT PARSER
  // ════════════════════════════════════════════════════════
  parseText(rawText, defaults = {}) {
    if (!rawText || !rawText.trim()) return [];
    const blocks = this._splitTextIntoBlocks(rawText.replace(/\r\n/g, '\n').replace(/\r/g, '\n'));
    const result = [];
    blocks.forEach((block, i) => {
      if (!block.trim()) return;
      const soal = this._parseSingleBlock(block, i + 1, defaults);
      if (soal) result.push(soal);
    });
    return result;
  },

  _splitTextIntoBlocks(text) {
    // Priority 1: "SOAL N" headers
    if (/^SOAL\s+\d+/im.test(text)) {
      return text.split(/(?=^SOAL\s+\d+)/im).filter(b => b.trim());
    }

    // Priority 2: "--- " separator lines
    if (/^-{3,}$/m.test(text)) {
      return text.split(/^-{3,}$/m).filter(b => b.trim());
    }

    // Priority 3: Numbered questions "1." pattern where options A-E follow
    const numberedParts = text.split(/(?=^\d+\.\s+)/m).filter(b => b.trim());
    if (numberedParts.length > 1) {
      // Verify at least one part has options
      const hasOptions = numberedParts.some(p => /^[A-Ea-e][.\)]\s/m.test(p));
      if (hasOptions) return numberedParts;
    }

    return [text];
  },

  _parseSingleBlock(block, defaultNum, defaults) {
    const soal = this.createSoal(defaultNum, defaults);
    const lines = block.split('\n');

    let mode = 'meta';
    const pertLines = [], stimLines = [], pembahasanLines = [];
    // Track current option being extended (multiline)
    let currentOpt = null;

    for (let i = 0; i < lines.length; i++) {
      const raw = lines[i];
      const line = raw.trim();
      if (!line) {
        // Empty line may end pertanyaan mode
        if (mode === 'pertanyaan' && pertLines.length > 0 && i < lines.length - 1) {
          // Check if next non-empty line is an option
          let j = i + 1;
          while (j < lines.length && !lines[j].trim()) j++;
          if (j < lines.length && /^[(]?[A-Ea-e][.)]\s/.test(lines[j].trim())) {
            mode = 'pilihan';
          }
        }
        continue;
      }

      // ── SOAL N header ──
      if (/^SOAL\s+(\d+)/i.test(line)) {
        const m = line.match(/^SOAL\s+(\d+)/i);
        if (m) soal.nomorUrut = parseInt(m[1]);
        mode = 'meta';
        continue;
      }

      // ── Numbered question "1. text..." ──
      if (/^\d+\.\s+\S/.test(line) && mode === 'meta') {
        const m = line.match(/^(\d+)\.\s+(.*)/);
        if (m) {
          soal.nomorUrut = parseInt(m[1]);
          if (m[2].trim()) pertLines.push(m[2].trim());
          mode = 'pertanyaan';
          continue;
        }
      }

      // ── Explicit Kompetensi/Materi check ──
      if (/^Kompetensi\s*[/]\s*Materi\s*:/i.test(line)) {
        const val = line.replace(/^Kompetensi\s*[/]\s*Materi\s*:\s*/i, '').trim();
        if (val) soal.kompetensiMateri = val;
        continue;
      }

      // ── Metadata key:value ──
      const metaRe = /^(Kode\s*Soal|Kode|Mapel|Mata\s*Pelajaran|Kelas[/]Jurusan|Kelas|Jurusan|Kompetensi[/]Materi|Kompetensi|Materi|Topik|Level|Kesulitan|Skor|Bobot|Poin)\s*:\s*(.*)/i;
      const metaM = line.match(metaRe);
      if (metaM) {
        const key = metaM[1].toLowerCase().replace(/[\s\/]/g, '');
        const val = metaM[2].trim();
        if (key === 'kode' || key === 'kodesoal') soal.kodeSoal = val;
        else if (key === 'mapel' || key === 'matapelajaran') { if (val) soal.mapel = val; }
        else if (key === 'kelas' || key === 'kelasjurusan' || key === 'jurusan') { if (val) soal.kelasJurusan = val; }
        else if (key === 'kompetensimater' || key === 'kompetensi' || key === 'materi' || key === 'topik') { if (val) soal.kompetensiMateri = val; }
        else if (key === 'level' || key === 'kesulitan') { const lv = this.normalizeLevel(val); if (lv) soal.level = lv; }
        else if (key === 'skor' || key === 'bobot' || key === 'poin') soal.skor = parseInt(val) || 1;
        continue;
      }

      // ── Section headers ──
      if (/^Stimulus\b.*:/i.test(line)) {
        const after = line.replace(/^Stimulus[^:]*:\s*/i, '').trim();
        if (after) stimLines.push(after);
        mode = 'stimulus'; currentOpt = null; continue;
      }
      if (/^Pertanyaan\s*:/i.test(line)) {
        const after = line.replace(/^Pertanyaan\s*:\s*/i, '').trim();
        if (after) pertLines.push(after);
        mode = 'pertanyaan'; currentOpt = null; continue;
      }
      if (/^Pilihan\s*(Jawaban)?\s*:/i.test(line)) {
        mode = 'pilihan'; currentOpt = null; continue;
      }

      // ── Answer key ──
      const kunciRe = /^(Kunci\s*Jawaban|Kunci|Jawaban\s*Benar|Jawaban|Jawab|Answer\s*Key|Answer)\s*:\s*([A-Ea-e])?/i;
      const kunciM = line.match(kunciRe);
      if (kunciM) {
        if (kunciM[2]) {
          soal.kunci = kunciM[2].toUpperCase();
        } else {
          // Try to find letter in remaining text
          const rest = line.replace(/^[^:]+:\s*/, '');
          const lm = rest.match(/^([A-Ea-e])/);
          if (lm) soal.kunci = lm[1].toUpperCase();
        }
        mode = 'post'; currentOpt = null; continue;
      }

      // ── Pembahasan ──
      if (/^(Pembahasan|Penjelasan|Solusi|Explanation)\s*:/i.test(line)) {
        const after = line.replace(/^[^:]+:\s*/i, '').trim();
        if (after) pembahasanLines.push(after);
        mode = 'pembahasan'; currentOpt = null; continue;
      }

      // ── Options: A. / A) / (A) / A  text ──
      const optRe = /^[(]?([A-Ea-e])[.\)]\s+(.*)/;
      const optM = line.match(optRe);
      if (optM) {
        currentOpt = optM[1].toUpperCase();
        soal.pilihanJawaban[currentOpt] = optM[2].trim();
        mode = 'pilihan'; continue;
      }

      // ── Continuation / free text ──
      switch (mode) {
        case 'stimulus':
          stimLines.push(line); break;
        case 'pertanyaan':
          pertLines.push(line); break;
        case 'pilihan':
          // Multi-line option continuation
          if (currentOpt && soal.pilihanJawaban[currentOpt] !== undefined) {
            soal.pilihanJawaban[currentOpt] += ' ' + line;
          }
          break;
        case 'pembahasan':
          pembahasanLines.push(line); break;
        case 'meta':
          // If starts looking like a question text (not a metadata field)
          // Allow word characters, spaces, slashes, and other common separators in metadata keys
          if (!/^[\w\s\/\-]+\s*:/.test(line)) {
            pertLines.push(line); mode = 'pertanyaan';
          }
          break;
      }
    }

    soal.stimulus = stimLines.join('\n').trim();
    soal.pertanyaan = pertLines.join('\n').trim();
    soal.pembahasan = pembahasanLines.join('\n').trim();

    return this.validateSoal(soal);
  },

  // ════════════════════════════════════════════════════════
  // EXCEL PARSER
  // ════════════════════════════════════════════════════════
  parseExcel(rows, defaults = {}) {
    if (!rows || rows.length < 2) return [];

    // Normalize headers
    const rawHeaders = rows[0].map(h => String(h || '').trim());
    const headers = rawHeaders.map(h => h.toLowerCase()
      .replace(/\s+/g, '')
      .replace(/[\/\\]/g, '')
      .replace(/[^a-z0-9]/g, '')
    );

    const findCol = (...candidates) => {
      for (const c of candidates) {
        const normalized = c.toLowerCase().replace(/[^a-z0-9]/g, '');
        const idx = headers.findIndex(h => h === normalized || h.startsWith(normalized));
        if (idx >= 0) return idx;
      }
      return -1;
    };

    const cols = {
      no:          findCol('no', 'nomor', 'number', 'urutan'),
      kode:        findCol('kode', 'kodesoal', 'code'),
      mapel:       findCol('mapel', 'matapelajaran', 'subject', 'pelajaran', 'matkul'),
      kelas:       findCol('kelas', 'kelasjurusan', 'jurusan', 'class'),
      kompetensi:  findCol('kompetensi', 'kompetensimater', 'materi', 'topik', 'topic'),
      level:       findCol('level', 'kesulitan', 'difficulty', 'tingkat'),
      skor:        findCol('skor', 'bobot', 'score', 'nilai', 'poin'),
      stimulus:    findCol('stimulus', 'wacana', 'bacaan', 'teks'),
      pertanyaan:  findCol('pertanyaan', 'soal', 'question', 'pertanyaan'),
      a:           findCol('opsia', 'piliha', 'optiona', 'a'),
      b:           findCol('opsib', 'pilihb', 'optionb', 'b'),
      c:           findCol('opsic', 'pilihc', 'optionc', 'c'),
      d:           findCol('opsid', 'pilihd', 'optiond', 'd'),
      e:           findCol('opsie', 'pilihe', 'optione', 'e'),
      kunci:       findCol('kunci', 'kuncijawaban', 'jawaban', 'answer', 'key', 'jawab'),
      pembahasan:  findCol('pembahasan', 'penjelasan', 'explanation', 'solusi', 'alasan'),
    };

    const result = [];

    for (let i = 1; i < rows.length; i++) {
      const row = rows[i];
      if (!row || row.length === 0) continue;
      const isEmpty = row.every(c => c === '' || c === null || c === undefined);
      if (isEmpty) continue;

      const get = (col) => col >= 0 && row[col] !== undefined ? String(row[col]).trim() : '';

      const soal = this.createSoal(i, defaults);

      if (cols.no >= 0) { const n = parseInt(get(cols.no)); if (!isNaN(n)) soal.nomorUrut = n; }
      soal.kodeSoal        = get(cols.kode);
      if (get(cols.mapel))       soal.mapel          = get(cols.mapel);
      if (get(cols.kelas))       soal.kelasJurusan   = get(cols.kelas);
      soal.kompetensiMateri = get(cols.kompetensi);
      const lv = this.normalizeLevel(get(cols.level));
      if (lv) soal.level = lv;
      const sk = parseInt(get(cols.skor));
      if (!isNaN(sk)) soal.skor = sk;
      soal.stimulus         = get(cols.stimulus);
      soal.pertanyaan       = get(cols.pertanyaan);
      soal.pilihanJawaban.A = get(cols.a);
      soal.pilihanJawaban.B = get(cols.b);
      soal.pilihanJawaban.C = get(cols.c);
      soal.pilihanJawaban.D = get(cols.d);
      soal.pilihanJawaban.E = get(cols.e);
      const rawKunci = get(cols.kunci).toUpperCase();
      soal.kunci = rawKunci.charAt(0);
      soal.pembahasan = get(cols.pembahasan);

      result.push(this.validateSoal(soal));
    }

    return result;
  },

  // ════════════════════════════════════════════════════════
  // WORD PARSER — delegates to text parser after extraction
  // ════════════════════════════════════════════════════════
  parseWordText(extractedText, defaults = {}) {
    // Clean up common mammoth artifacts
    const cleaned = extractedText
      .replace(/\t/g, ' ')
      .replace(/ {3,}/g, '  ')
      .trim();
    return this.parseText(cleaned, defaults);
  }

};
