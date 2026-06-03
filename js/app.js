/**
 * APP.JS
 * Main UI controller — manages wizard flow, file uploads,
 * preview rendering, edit modal, and generate/download actions.
 */

const App = {
  soalList: [],
  generatedBlob: null,
  downloadFilename: 'output-soal.docx',
  currentTab: 'text',
  editingIndex: -1,
  excelFile: null,
  wordFile: null,

  // ── Init ────────────────────────────────────────────────
  init() {
    this._setupTabs();
    this._setupDropzones();
    this._setupButtons();
    this._setupModal();
  },

  // ── Tabs ────────────────────────────────────────────────
  _setupTabs() {
    document.querySelectorAll('.tab-btn').forEach(btn => {
      btn.addEventListener('click', () => this._switchTab(btn.dataset.tab));
    });
  },

  _switchTab(tab) {
    this.currentTab = tab;
    document.querySelectorAll('.tab-btn').forEach(b => b.classList.toggle('active', b.dataset.tab === tab));
    document.querySelectorAll('.tab-content').forEach(c => c.classList.toggle('active', c.id === `tab-${tab}`));
    document.getElementById('error-message').style.display = 'none';
  },

  // ── File Dropzones ──────────────────────────────────────
  _setupDropzones() {
    ['excel', 'word'].forEach(type => {
      const zone  = document.getElementById(`${type}-dropzone`);
      const input = document.getElementById(`${type}-input`);

      zone.addEventListener('dragover',  e => { e.preventDefault(); zone.classList.add('dragover'); });
      zone.addEventListener('dragleave', () => zone.classList.remove('dragover'));
      zone.addEventListener('drop',      e => {
        e.preventDefault(); zone.classList.remove('dragover');
        const file = e.dataTransfer.files[0];
        if (file) this._handleFile(type, file);
      });

      input.addEventListener('change', () => {
        if (input.files[0]) this._handleFile(type, input.files[0]);
      });

      // Also make whole dropzone clickable
      zone.addEventListener('click', e => {
        if (!e.target.closest('.btn-upload')) input.click();
      });
    });
  },

  _handleFile(type, file) {
    const el = document.getElementById(`${type}-file-info`);
    el.innerHTML = `📎 <strong>${file.name}</strong> (${this._fmtSize(file.size)})`;
    this[`${type}File`] = file;
  },

  _fmtSize(bytes) {
    if (bytes < 1024) return bytes + ' B';
    if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + ' KB';
    return (bytes / (1024 * 1024)).toFixed(1) + ' MB';
  },

  // ── Buttons ─────────────────────────────────────────────
  _setupButtons() {
    document.getElementById('btn-parse').addEventListener('click',    () => this._handleParse());
    document.getElementById('btn-back').addEventListener('click',     () => this._goToStep(1));
    document.getElementById('btn-generate').addEventListener('click', () => this._handleGenerate());
    document.getElementById('btn-download').addEventListener('click', () => this._handleDownload());
    document.getElementById('btn-reset').addEventListener('click',    () => this._handleReset());
    document.getElementById('btn-load-example').addEventListener('click', () => this._loadExample());
  },

  // ── Defaults ────────────────────────────────────────────
  _getDefaults() {
    return {
      mapel: document.getElementById('default-mapel').value.trim(),
      kelas: document.getElementById('default-kelas').value.trim(),
      kompetensi: document.getElementById('default-kompetensi').value.trim(),
      level: document.getElementById('default-level').value,
      skor:  parseInt(document.getElementById('default-skor').value) || 1
    };
  },

  // ── PARSE ────────────────────────────────────────────────
  async _handleParse() {
    const defaults = this._getDefaults();
    let soalList = [];

    try {
      this._showLoading('Memproses soal...');

      if (this.currentTab === 'text') {
        const text = document.getElementById('text-input').value;
        if (!text.trim()) throw new Error('Input teks kosong. Masukkan atau paste soal terlebih dahulu.');
        soalList = SoalParser.parseText(text, defaults);

      } else if (this.currentTab === 'excel') {
        if (!this.excelFile) throw new Error('Belum ada file Excel dipilih. Drag & drop atau pilih file .xlsx terlebih dahulu.');
        if (typeof XLSX === 'undefined') throw new Error('Library XLSX belum termuat. Coba refresh halaman.');
        const buf = await this._readFileBuffer(this.excelFile);
        const wb  = XLSX.read(new Uint8Array(buf), { type: 'array' });
        const ws  = wb.Sheets[wb.SheetNames[0]];
        const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
        soalList = SoalParser.parseExcel(data, defaults);

      } else if (this.currentTab === 'word') {
        if (!this.wordFile) throw new Error('Belum ada file Word dipilih. Drag & drop atau pilih file .docx terlebih dahulu.');
        if (typeof mammoth === 'undefined') throw new Error('Library Mammoth belum termuat. Coba refresh halaman.');
        const buf = await this._readFileBuffer(this.wordFile);
        const res = await mammoth.extractRawText({ arrayBuffer: buf });
        if (!res.value.trim()) throw new Error('File Word tidak mengandung teks yang dapat dibaca.');
        soalList = SoalParser.parseWordText(res.value, defaults);
      }

      if (soalList.length === 0) {
        throw new Error('Tidak ada soal yang berhasil diparsing. Pastikan format sesuai panduan.');
      }

      this.soalList = soalList;
      this._renderPreview();
      this._goToStep(2);

    } catch (err) {
      this._showError(err.message);
    } finally {
      this._hideLoading();
    }
  },

  _readFileBuffer(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload  = e => resolve(e.target.result);
      reader.onerror = () => reject(new Error('Gagal membaca file'));
      reader.readAsArrayBuffer(file);
    });
  },

  // ── PREVIEW ─────────────────────────────────────────────
  _renderPreview() {
    const valid   = this.soalList.filter(s => s.status === 'valid').length;
    const warning = this.soalList.filter(s => s.status === 'warning').length;
    const error   = this.soalList.filter(s => s.status === 'error').length;

    document.getElementById('preview-count').textContent =
      `${this.soalList.length} soal berhasil diparsing`;
    document.getElementById('stat-valid').textContent   = valid;
    document.getElementById('stat-warning').textContent = warning;
    document.getElementById('stat-error').textContent   = error;

    const list = document.getElementById('preview-list');
    list.innerHTML = this.soalList.map((s, i) => this._soalCardHTML(s, i)).join('');

    // Events
    list.querySelectorAll('.soal-card-header').forEach(h => {
      h.addEventListener('click', e => {
        if (e.target.closest('.icon-btn')) return;
        h.closest('.soal-card').classList.toggle('expanded');
      });
    });

    list.querySelectorAll('.btn-edit').forEach(b => {
      b.addEventListener('click', () => this._openEditModal(+b.dataset.index));
    });

    list.querySelectorAll('.btn-delete').forEach(b => {
      b.addEventListener('click', () => {
        if (confirm(`Hapus Soal ${this.soalList[+b.dataset.index].nomorUrut}?`)) {
          this.soalList.splice(+b.dataset.index, 1);
          this._renderPreview();
        }
      });
    });
  },

  _soalCardHTML(soal, index) {
    const icons = { valid: '✅', warning: '⚠️', error: '❌' };
    const labels = { valid: 'Valid', warning: 'Peringatan', error: 'Error' };
    const preview = soal.pertanyaan
      ? soal.pertanyaan.replace(/\n/g, ' ').substring(0, 90) + (soal.pertanyaan.length > 90 ? '…' : '')
      : '<em style="color:var(--text-muted)">Pertanyaan kosong</em>';
    const issues = [...(soal.errors || []), ...(soal.warnings || [])];

    const optRow = opt => {
      const isCorrect = soal.kunci === opt;
      const val = soal.pilihanJawaban[opt] || '<em style="color:var(--text-muted)">-</em>';
      return `
        <div class="pilihan-item ${isCorrect ? 'correct' : ''}">
          <span class="pilihan-opt">${opt}</span>
          <span class="pilihan-val">${val}</span>
        </div>`;
    };

    return `
      <div class="soal-card ${soal.status}">
        <div class="soal-card-header">
          <div class="soal-card-title">
            <span class="soal-num">Soal ${soal.nomorUrut}</span>
            <span class="soal-preview">${preview}</span>
          </div>
          <div class="soal-card-meta">
            <span class="status-badge ${soal.status}">${icons[soal.status]} ${labels[soal.status]}</span>
            <button class="btn-edit icon-btn" data-index="${index}" title="Edit soal">✏️</button>
            <button class="btn-delete icon-btn" data-index="${index}" title="Hapus soal">🗑️</button>
            <span class="chevron">›</span>
          </div>
        </div>
        <div class="soal-card-body">
          ${issues.length ? `
            <div class="issues-list ${soal.status}" style="margin-top:16px">
              ${issues.map(i => `<div class="issue-item">${i}</div>`).join('')}
            </div>` : ''}
          <div class="soal-fields-grid">
            <div class="field-item"><span class="field-label">Kode Soal</span><span class="field-value">${soal.kodeSoal || '–'}</span></div>
            <div class="field-item"><span class="field-label">Mapel</span><span class="field-value">${soal.mapel || '–'}</span></div>
            <div class="field-item"><span class="field-label">Kelas/Jurusan</span><span class="field-value">${soal.kelasJurusan || '–'}</span></div>
            <div class="field-item"><span class="field-label">Kompetensi/Materi</span><span class="field-value">${soal.kompetensiMateri || '–'}</span></div>
            <div class="field-item"><span class="field-label">Level</span><span class="field-value">${soal.level || '–'}</span></div>
            <div class="field-item"><span class="field-label">Skor</span><span class="field-value">${soal.skor}</span></div>
            <div class="field-item"><span class="field-label">Kunci</span><span class="field-value"><span class="kunci-badge">${soal.kunci || '?'}</span></span></div>
          </div>
          ${soal.stimulus ? `<div class="field-stimulus"><span class="field-label">Stimulus</span>${soal.stimulus.replace(/\n/g,'<br>')}</div>` : ''}
          <div class="field-pertanyaan"><span class="field-label">Pertanyaan</span>${(soal.pertanyaan || '–').replace(/\n/g,'<br>')}</div>
          <div class="pilihan-grid">
            ${'ABCDE'.split('').map(optRow).join('')}
          </div>
          ${soal.pembahasan ? `<div class="field-pembahasan"><span class="field-label">Pembahasan</span>${soal.pembahasan.replace(/\n/g,'<br>')}</div>` : ''}
        </div>
      </div>`;
  },

  // ── EDIT MODAL ──────────────────────────────────────────
  _setupModal() {
    document.getElementById('btn-modal-save').addEventListener('click',   () => this._saveEdit());
    document.getElementById('btn-modal-cancel').addEventListener('click', () => this._closeModal());
    document.getElementById('btn-modal-close').addEventListener('click',  () => this._closeModal());
    document.getElementById('edit-modal').addEventListener('click', e => {
      if (e.target === e.currentTarget) this._closeModal();
    });
    document.addEventListener('keydown', e => {
      if (e.key === 'Escape') this._closeModal();
    });
  },

  _openEditModal(index) {
    this.editingIndex = index;
    const s = this.soalList[index];
    const $ = id => document.getElementById(id);

    $('edit-kode').value       = s.kodeSoal || '';
    $('edit-mapel').value      = s.mapel || '';
    $('edit-kelas').value      = s.kelasJurusan || '';
    $('edit-kompetensi').value = s.kompetensiMateri || '';
    $('edit-level').value      = s.level || 'Sedang';
    $('edit-skor').value       = s.skor || 1;
    $('edit-stimulus').value   = s.stimulus || '';
    $('edit-pertanyaan').value = s.pertanyaan || '';
    $('edit-a').value          = s.pilihanJawaban.A || '';
    $('edit-b').value          = s.pilihanJawaban.B || '';
    $('edit-c').value          = s.pilihanJawaban.C || '';
    $('edit-d').value          = s.pilihanJawaban.D || '';
    $('edit-e').value          = s.pilihanJawaban.E || '';
    $('edit-kunci').value      = s.kunci || '';
    $('edit-pembahasan').value = s.pembahasan || '';

    $('edit-modal').classList.add('open');
  },

  _saveEdit() {
    if (this.editingIndex < 0) return;
    const s = this.soalList[this.editingIndex];
    const $ = id => document.getElementById(id);

    s.kodeSoal         = $('edit-kode').value.trim();
    s.mapel            = $('edit-mapel').value.trim();
    s.kelasJurusan     = $('edit-kelas').value.trim();
    s.kompetensiMateri = $('edit-kompetensi').value.trim();
    s.level            = $('edit-level').value;
    s.skor             = parseInt($('edit-skor').value) || 1;
    s.stimulus         = $('edit-stimulus').value.trim();
    s.pertanyaan       = $('edit-pertanyaan').value.trim();
    s.pilihanJawaban.A = $('edit-a').value.trim();
    s.pilihanJawaban.B = $('edit-b').value.trim();
    s.pilihanJawaban.C = $('edit-c').value.trim();
    s.pilihanJawaban.D = $('edit-d').value.trim();
    s.pilihanJawaban.E = $('edit-e').value.trim();
    s.kunci            = $('edit-kunci').value;
    s.pembahasan       = $('edit-pembahasan').value.trim();

    SoalParser.validateSoal(s);
    this._closeModal();
    this._renderPreview();
  },

  _closeModal() {
    document.getElementById('edit-modal').classList.remove('open');
    this.editingIndex = -1;
  },

  // ── GENERATE ────────────────────────────────────────────
  async _handleGenerate() {
    const toGenerate = this.soalList.filter(s => s.status !== 'error');

    if (toGenerate.length === 0) {
      alert('Tidak ada soal yang valid untuk digenerate.\nPerbaiki error pada soal terlebih dahulu, atau gunakan tombol edit (✏️).');
      return;
    }

    const errorCount = this.soalList.length - toGenerate.length;
    if (errorCount > 0) {
      const ok = confirm(
        `${errorCount} soal memiliki error dan akan DILEWATI.\n` +
        `Hanya ${toGenerate.length} soal valid yang akan digenerate.\n\nLanjutkan?`
      );
      if (!ok) return;
    }

    try {
      this._showLoading(`Generating ${toGenerate.length} soal ke .docx...`);
      const blob = await DocxGenerator.generate(toGenerate);
      this.generatedBlob = blob;

      // Generate dynamic filename: [mapel]_[kelas]_ASAT_2526.docx
      const firstSoal = toGenerate[0] || {};
      const defaults = this._getDefaults();
      const mapelRaw = (firstSoal.mapel || defaults.mapel || '').trim();
      const kelasRaw = (firstSoal.kelasJurusan || defaults.kelas || '').trim();

      const cleanForFilename = (str) => str.replace(/[\/\\?%*:|"<>\s]+/g, '_');
      const cleanMapel = cleanForFilename(mapelRaw);
      const cleanKelas = cleanForFilename(kelasRaw);

      let nameParts = [];
      if (cleanMapel) nameParts.push(cleanMapel);
      if (cleanKelas) nameParts.push(cleanKelas);
      nameParts.push('ASAT_2526');

      this.downloadFilename = `${nameParts.join('_')}.docx`;

      document.getElementById('success-count').textContent =
        `${toGenerate.length} soal berhasil dikonversi ke template soal Moodle`;
      document.getElementById('download-filename').textContent = this.downloadFilename;
      document.getElementById('download-count').textContent = `${toGenerate.length} soal`;

      this._goToStep(3);
    } catch (err) {
      console.error(err);
      this._showError('Gagal generate dokumen: ' + err.message);
    } finally {
      this._hideLoading();
    }
  },

  // ── DOWNLOAD ────────────────────────────────────────────
  _handleDownload() {
    if (!this.generatedBlob) return;
    const url = URL.createObjectURL(this.generatedBlob);
    const a   = document.createElement('a');
    a.href     = url;
    a.download = this.downloadFilename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  },

  // ── RESET ────────────────────────────────────────────────
  _handleReset() {
    this.soalList      = [];
    this.generatedBlob = null;
    this.excelFile     = null;
    this.wordFile      = null;

    document.getElementById('text-input').value = '';
    document.getElementById('excel-file-info').innerHTML = '';
    document.getElementById('word-file-info').innerHTML  = '';
    document.getElementById('error-message').style.display = 'none';

    const excelInput = document.getElementById('excel-input');
    const wordInput  = document.getElementById('word-input');
    excelInput.value = '';
    wordInput.value  = '';

    this._goToStep(1);
    this._switchTab('text');
  },

  // ── Example Soal ────────────────────────────────────────
  _loadExample() {
    document.getElementById('text-input').value = `SOAL 1
Kode: VDG-001
Mapel: Videografi
Kelas/Jurusan: XI DKV 4
Kompetensi/Materi: Teknik Pengambilan Gambar
Level: Sedang
Skor: 1

Pertanyaan:
Sudut pengambilan gambar di mana kamera diletakkan lebih rendah dari ketinggian objek, sehingga objek tampak lebih dominan, besar, dan berwibawa disebut...

A. Bird eye view
B. High angle
C. Eye level
D. Low angle
E. Frog eye view

Kunci: D
Pembahasan: Low angle adalah teknik pengambilan gambar dengan memosisikan kamera lebih rendah dari objek sehingga memberikan kesan objek terlihat megah, berwibawa, atau dominan.

---

SOAL 2
Kode: VDG-002
Mapel: Videografi
Kelas/Jurusan: XI DKV 4
Kompetensi/Materi: Pengaturan Eksposur Kamera
Level: Mudah
Skor: 1

Pertanyaan:
Komponen segitiga eksposur (exposure triangle) yang berfungsi mengontrol durasi sensor kamera dalam menerima cahaya adalah...

A. Aperture
B. Shutter speed
C. ISO
D. White balance
E. Focus ring

Kunci: B
Pembahasan: Shutter speed (kecepatan rana) menentukan lamanya waktu sensor kamera terbuka untuk menerima cahaya yang masuk.

---

SOAL 3
Kode: VDG-003
Mapel: Videografi
Kelas/Jurusan: XI DKV 4
Kompetensi/Materi: Pergerakan Kamera (Camera Movement)
Level: Sulit
Skor: 2

Stimulus:
Dalam memproduksi sebuah film pendek tentang kepanikan tokoh utama di tengah keramaian kota, sutradara menginginkan kamera bergerak mengikuti langkah kaki tokoh utama secara stabil tanpa adanya guncangan.

Pertanyaan:
Alat bantu pergerakan kamera dan teknik kamera yang paling tepat digunakan untuk mencapai visualisasi tersebut adalah...

A. Tripod dengan teknik panning
B. Jib arm dengan teknik tilting
C. Gimbal / Stabilizer dengan teknik tracking
D. Dolly track dengan teknik pedestal
E. Handheld dengan teknik rolling

Kunci: C
Pembahasan: Gimbal/stabilizer berfungsi meredam guncangan saat kamera dibawa bergerak dinamis mengikuti objek (tracking/following shot).`;

    this._switchTab('text');
    document.getElementById('text-input').scrollTop = 0;
  },

  // ── Step Navigation ─────────────────────────────────────
  _goToStep(n) {
    document.querySelectorAll('.step-section').forEach((el, i) => {
      el.classList.toggle('active', i + 1 === n);
    });
    document.querySelectorAll('.step').forEach((el, i) => {
      el.classList.remove('active', 'completed');
      if (i + 1 < n)  el.classList.add('completed');
      if (i + 1 === n) el.classList.add('active');
    });
    window.scrollTo({ top: 0, behavior: 'smooth' });
  },

  // ── Loading / Error UI ──────────────────────────────────
  _showLoading(msg) {
    document.getElementById('loading-msg').textContent = msg || 'Memproses...';
    document.getElementById('loading-overlay').style.display = 'flex';
  },

  _hideLoading() {
    document.getElementById('loading-overlay').style.display = 'none';
  },

  _showError(msg) {
    const el = document.getElementById('error-message');
    el.textContent = '⚠️ ' + msg;
    el.style.display = 'block';
    el.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
    // Auto-hide after 8s
    clearTimeout(this._errorTimeout);
    this._errorTimeout = setTimeout(() => { el.style.display = 'none'; }, 8000);
  }
};

// Boot
document.addEventListener('DOMContentLoaded', () => App.init());
