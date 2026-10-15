/*
 * METT-4 Web — Model data penilaian
 * ------------------------------------------------------------------
 * - Menyimpan isian pengguna sebagai "sel virtual" dengan alamat yang SAMA
 *   dengan file Excel (mis. "1.!B15", "Atribut Kawasan!B5",
 *   "Question_data_dump!H2" untuk checkbox yang tertaut ke sel tsb.).
 * - Mereplikasi logika VBA file asli (skor option button A–D, skor checkbox
 *   tambahan, "tidak relevan", isian wajib, multi-pilih dropdown). Aturan VBA
 *   diurai otomatis dari kode makro oleh tools/extract.py; di sini hanya
 *   dijalankan.
 * - Validasi data (dropdown, angka, email, bilangan bulat) mengikuti Data
 *   Validation pada file Excel.
 */
(function (root) {
  'use strict';
  const E = root.MettEngine || (typeof require !== 'undefined' ? require('./engine.js') : null);

  const APP_ID = 'METT4-KSDAE-web';
  const FILE_SCHEMA = 1;
  const DAFTAR = 'Daftar Pertanyaan & Skor';
  const DUMP = 'Question_data_dump';

  // ------------------------------------------------------------------ util alamat
  function splitKey(key) {
    const i = key.lastIndexOf('!');
    return { sheet: key.slice(0, i), addr: key.slice(i + 1) };
  }
  function parseRef(ref) {
    // "'Daftar Pertanyaan & Skor'!$A$6" -> {sheet, addr}
    const m = /^'?(.*?)'?!\$?([A-Z]+)\$?(\d+)$/.exec(ref.replace(/&amp;/g, '&'));
    return m ? { sheet: m[1].replace(/''/g, "'"), addr: m[2] + m[3] } : null;
  }
  function inSqref(sqref, addr) {
    const p = E.parseA1(addr);
    for (const part of String(sqref).split(/\s+/)) {
      if (!part) continue;
      const [a, b] = part.split(':');
      const p1 = E.parseA1(a), p2 = b ? E.parseA1(b) : p1;
      if (!p1 || !p2) continue;
      if (p.c >= Math.min(p1.c, p2.c) && p.c <= Math.max(p1.c, p2.c) && p.r >= Math.min(p1.r, p2.r) && p.r <= Math.max(p1.r, p2.r)) return true;
    }
    return false;
  }
  const decodeX = (s) => (s == null ? s : String(s).replace(/_x000a_/gi, '\n').replace(/_x000d_/gi, ''));

  // ------------------------------------------------------------------ skema pertanyaan
  function buildSchema(wb) {
    const byName = new Map(wb.sheets.map((s) => [s.name, s]));
    const daftar = byName.get(DAFTAR);
    const dump = byName.get(DUMP);
    // peta sel checkbox di Question_data_dump -> sel teks keterangan di kolom berikutnya
    const dumpTextFor = (linkAddr) => {
      const p = E.parseA1(linkAddr);
      const next = E.toA1(p.c + 1, p.r);
      const c = dump && dump.cells[next];
      if (!c || !c.f) return null;
      const r = parseRef(c.f.replace(/^=/, ''));
      return r ? r.sheet + '!' + r.addr : null;
    };
    const questions = [];
    for (const s of wb.sheets) {
      const rule = s.vba;
      if (!rule || !rule.kind) continue;
      const ctl = (name) => s.controls.find((c) => c.name === name);
      const naCtl = rule.na ? ctl(rule.na) : null;
      const naLink = naCtl && naCtl.link ? parseRef(naCtl.link) : null;
      const naMacro = naCtl && naCtl.macro ? wb.naMacros[naCtl.macro] : null;
      const q = {
        sheet: s.name,
        kind: rule.kind,
        title: cellText(s, 'B1'),
        naKey: naLink ? naLink.sheet + '!' + naLink.addr : null,
        naCtl: rule.na,
        naLabel: cellText(s, 'B2'),
        naReasonKey: s.name + '!' + (rule.naReasonCell || 'B3'),
        scoreKey: s.name + '!' + (rule.scoreCell || 'C6'),
        daftarRow: naLink ? E.parseA1(naLink.addr).r : null,
        naMacro: naMacro || null,
        required: rule.required.map((r) => Object.assign({}, r, { key: s.name + '!' + r.cell })),
        options: [], boxes: [], sources: [], vbaNote: null,
      };
      const reqAnswered = rule.required.filter((r) => r.when === 'answered');
      q.justKey = reqAnswered[0] ? s.name + '!' + reqAnswered[0].cell : null;
      q.actionKey = reqAnswered[1] ? s.name + '!' + reqAnswered[1].cell : null;
      const labelAbove = (cell) => { if (!cell) return ''; const p = E.parseA1(cell); return cellText(s, 'B' + (p.r - 2)); };
      q.justLabel = labelAbove(reqAnswered[0] && reqAnswered[0].cell);
      q.actionLabel = labelAbove(reqAnswered[1] && reqAnswered[1].cell);
      if (rule.kind === 'radio') {
        q.question = cellText(s, 'B5');
        q.guidance = cellText(s, 'B6');
        q.optionsHeader = cellText(s, 'B7');
        for (const o of rule.options) {
          const c = ctl(o.name);
          const row = c ? c.row : null;
          q.options.push({ name: o.name, score: o.score, row, letter: row ? cellText(s, 'A' + row).trim() : '', text: row ? cellText(s, 'B' + row) : '' });
        }
        q.options.sort((a, b) => a.row - b.row);
      } else {
        q.guidance = cellText(s, 'B5');
        q.boxesHeader = cellText(s, 'B6');
        let names = rule.scoreBoxes.filter((n) => ctl(n));
        if (names.length === 0 && naMacro) {
          // VBA merujuk checkbox yang tidak ada di sheet (lihat docs/TEMUAN.md) -> pakai checkbox
          // yang dikelola makro "tidak relevan" pada sheet ini (maksud perancang).
          names = naMacro.clearWhenOn.filter((n) => ctl(n) && /Check Box/.test(n));
          q.vbaNote = 'VBA asli menghitung "' + rule.scoreBoxes.join('", "') + '" yang tidak ada pada sheet ini; ' +
            'versi web menghitung checkbox ' + names.join(', ') + ' (lihat TEMUAN.md).';
        }
        const sorted = names.map((n) => ctl(n)).sort((a, b) => (a.anchor[2] - b.anchor[2]) || (a.anchor[3] - b.anchor[3]));
        // teks pernyataan: sel kolom B setelah judul (B6) sampai sebelum label "Bukti"
        const firstReqRow = reqAnswered[0] ? E.parseA1(reqAnswered[0].cell).r - 2 : 11;
        const texts = [];
        for (let r = 7; r < firstReqRow; r++) { const t = cellText(s, 'B' + r); if (t) texts.push({ row: r, text: t }); }
        sorted.forEach((c, i) => q.boxes.push({ name: c.name, text: texts[i] ? texts[i].text : '', row: texts[i] ? texts[i].row : null }));
      }
      // sumber informasi: checkbox yang tertaut ke Question_data_dump
      const srcCtls = s.controls.filter((c) => c.link && /Question_data_dump/i.test(c.link))
        .sort((a, b) => (a.anchor[2] - b.anchor[2]) || (a.anchor[3] - b.anchor[3]) || (a.anchor[0] - b.anchor[0]));
      // sel teks keterangan sumber = sel isian (tidak terkunci) kolom B di bawah sel "Aksi",
      // dipasangkan berurutan dengan checkbox (sel yang benar-benar diketik pengguna di Excel)
      const actRow = reqAnswered[1] ? E.parseA1(reqAnswered[1].cell).r : 19;
      const srcCells = (s.unlocked || []).filter((a) => /^B\d+$/.test(a) && E.parseA1(a).r > actRow)
        .sort((a, b) => E.parseA1(a).r - E.parseA1(b).r);
      srcCtls.forEach((c, i) => {
        const l = parseRef(c.link);
        q.sources.push({
          label: (c.label || '').replace(/\s+/g, ' ').trim(), key: l.sheet + '!' + l.addr,
          textKey: srcCells[i] ? s.name + '!' + srcCells[i] : dumpTextFor(l.addr),
          exportTextKey: dumpTextFor(l.addr), // sel yang dibaca Question_data_dump (bisa berbeda, lihat TEMUAN.md)
        });
      });
      q.sourcesLabel = cellText(s, 'B' + (actRow + 2));
      // baris di "Daftar Pertanyaan & Skor"
      if (q.daftarRow && daftar) {
        q.no = cellText(daftar, 'B' + q.daftarRow);
        q.shortQuestion = cellText(daftar, 'C' + q.daftarRow);
        q.element = cellText(daftar, 'G' + q.daftarRow);
      }
      questions.push(q);
    }
    return { questions, byName };
  }
  function cellText(sheet, addr) {
    const c = sheet.cells[addr];
    if (!c || c.v === undefined || c.v === null) return '';
    return String(c.v);
  }

  // ------------------------------------------------------------------ validasi data
  function findValidation(wb, sheetName, addr) {
    const s = wb.sheets.find((x) => x.name === sheetName);
    if (!s) return null;
    // validasi x14/standard dengan tipe lebih diprioritaskan daripada yang hanya berisi prompt
    let found = null;
    for (const v of s.validations) {
      if (!inSqref(v.sqref, addr)) continue;
      if (!found || (!found.type && v.type)) found = v;
      else if (found && found.type && v.type && !found.prompt && v.prompt) found = Object.assign({}, found, { prompt: v.prompt });
    }
    if (found && !found.prompt) {
      const withPrompt = s.validations.find((v) => inSqref(v.sqref, addr) && v.prompt);
      if (withPrompt) found = Object.assign({}, found, { prompt: withPrompt.prompt });
    }
    return found ? Object.assign({}, found, { prompt: decodeX(found.prompt), error: decodeX(found.error) }) : null;
  }
  function listOptions(engine, validation, ctxSheet) {
    if (!validation || validation.type !== 'list' || !validation.formula1) return null;
    let f = validation.formula1.trim();
    if (/^".*"$/.test(f)) return f.slice(1, -1).split(',').map((x) => x.trim());
    const def = engine.wb.definedNames && engine.wb.definedNames[f];
    if (def) f = def;
    const r = /^'?(.*?)'?!\$?([A-Z]+)\$?(\d+)(?::\$?([A-Z]+)\$?(\d+))?$/.exec(f);
    let sheet = ctxSheet, a, b;
    if (r) { sheet = r[1].replace(/''/g, "'"); a = r[2] + r[3]; b = r[4] ? r[4] + r[5] : a; }
    else { const m = /^\$?([A-Z]+)\$?(\d+)(?::\$?([A-Z]+)\$?(\d+))?$/.exec(f); if (!m) return null; a = m[1] + m[2]; b = m[3] ? m[3] + m[4] : a; }
    const p1 = E.parseA1(a), p2 = E.parseA1(b);
    const out = [];
    for (let rr = p1.r; rr <= p2.r; rr++) for (let cc = p1.c; cc <= p2.c; cc++) {
      const v = engine.get(sheet, E.toA1(cc, rr));
      if (E.isBlank(v) || E.isErr(v) || v === '' || v === 0) continue; // sel kosong tidak ditampilkan
      out.push(E.toText(v));
    }
    return out;
  }
  /** Periksa nilai sesuai Data Validation Excel. Mengembalikan {ok, level:'stop'|'warning', msg} */
  function checkValidation(v, value) {
    if (!v || value === null || value === undefined || value === '') return { ok: true };
    const fail = (fallback) => ({ ok: false, level: v.errorStyle === 'warning' ? 'warning' : (v.errorStyle === 'information' ? 'information' : 'stop'), msg: v.error || fallback });
    if (v.type === 'whole' || v.type === 'decimal') {
      const n = typeof value === 'number' ? value : E.toNumber(String(value).replace(',', '.'));
      if (E.isErr(n) || (v.type === 'whole' && !Number.isInteger(n))) return fail('Nilai yang dimasukkan tidak valid.');
      const lo = v.formula1 != null ? parseFloat(v.formula1) : -Infinity;
      const hi = v.formula2 != null ? parseFloat(v.formula2) : Infinity;
      const op = v.operator || 'between';
      const ok = op === 'between' ? (n >= lo && n <= hi) : op === 'notBetween' ? !(n >= lo && n <= hi)
        : op === 'equal' ? n === lo : op === 'notEqual' ? n !== lo : op === 'greaterThan' ? n > lo
        : op === 'lessThan' ? n < lo : op === 'greaterThanOrEqual' ? n >= lo : op === 'lessThanOrEqual' ? n <= lo : true;
      return ok ? { ok: true } : fail('Nilai yang dimasukkan tidak valid.');
    }
    if (v.type === 'custom' && v.formula1) {
      if (/^ISNUMBER\(/i.test(v.formula1)) return typeof value === 'number' ? { ok: true } : fail('Masukan angka saja');
      if (/FIND\("@"/i.test(v.formula1)) {
        // =AND(FIND("@",x), FIND(".",x,FIND("@",x)+2)-FIND("@",x)>2)  (indeks 1-based seperti Excel)
        const s = String(value);
        const at = s.indexOf('@') + 1;
        if (at === 0) return fail('Email tidak valid');
        const dot = s.indexOf('.', at + 2 - 1) + 1;
        if (dot === 0) return fail('Email tidak valid');
        return dot - at > 2 ? { ok: true } : fail('Email tidak valid');
      }
    }
    if (v.type === 'list' && v._options && v.showError !== false) {
      if (!v._options.includes(String(value))) return fail('Nilai tidak ada dalam daftar.');
    }
    return { ok: true };
  }

  // ------------------------------------------------------------------ pengurai isian ala Excel
  /**
   * Excel mengubah ketikan menjadi angka/tanggal bila cocok (kecuali sel berformat Teks "@").
   * Direplikasi agar rumus yang memakai sel tsb. (mis. format Rp, tanggal) berperilaku sama.
   */
  function parseTyped(text, nf) {
    if (text === null || text === undefined) return null;
    if (typeof text !== 'string') return text;
    const s = text.trim();
    if (s === '') return null;
    if (nf === '@') return text;
    if (/^[+-]?\d{1,3}(\.\d{3})+(,\d+)?$/.test(s)) return parseFloat(s.replace(/\./g, '').replace(',', '.'));
    if (/^[+-]?\d+(,\d+)?$/.test(s) && !/^0\d/.test(s)) return parseFloat(s.replace(',', '.'));
    if (/^[+-]?\d+\.\d+$/.test(s) && !/^0\d/.test(s)) return parseFloat(s);
    const d = /^(\d{1,2})[/-](\d{1,2})[/-](\d{4})$/.exec(s);
    if (d) {
      const day = +d[1], mon = +d[2], yr = +d[3];
      if (mon >= 1 && mon <= 12 && day >= 1 && day <= 31) return E.dateToSerial(yr, mon, day);
    }
    return text;
  }

  // ------------------------------------------------------------------ penilaian (state)
  function newAssessment() {
    const now = new Date().toISOString();
    return { app: APP_ID, schema: FILE_SCHEMA, id: 'a' + Date.now().toString(36) + Math.random().toString(36).slice(2, 6), created: now, updated: now, inputs: {}, controls: {}, finalized: false };
  }

  class Model {
    constructor(wb) {
      this.wb = wb;
      this.schema = buildSchema(wb);
      this.state = newAssessment();
      this.engine = new E.Engine(wb, this.state.inputs);
      this.listeners = [];
    }
    load(state) {
      const s = Object.assign(newAssessment(), state || {});
      s.inputs = Object.assign({}, s.inputs || {});
      s.controls = Object.assign({}, s.controls || {});
      this.state = s;
      this.applyVba();
      this.engine.setInputs(this.state.inputs);
      this.emit();
    }
    onChange(fn) { this.listeners.push(fn); }
    emit(info) { this.state.updated = new Date().toISOString(); for (const fn of this.listeners) fn(info || {}); }

    // ---- data bersama (sinkronisasi per sel)
    /** Kunci turunan (skor C6 hasil VBA) tidak dikirim — selalu dihitung ulang dari pilihan jawaban. */
    isDerivedKey(k) {
      if (!this._derived) this._derived = new Set(this.schema.questions.map((q) => q.scoreKey));
      return this._derived.has(k);
    }
    /** Semua isian yang perlu dibagikan: {kunci: nilai}. Kontrol (radio/checkbox) memakai kunci "Sheet!#Nama". */
    syncMap() {
      const out = {};
      for (const [k, v] of Object.entries(this.state.inputs)) if (!this.isDerivedKey(k)) out[k] = v === undefined ? null : v;
      for (const [k, v] of Object.entries(this.state.controls)) out[k] = !!v;
      out['__app!finalized'] = this.state.finalized ? (this.state.finalizedAt || true) : null;
      return out;
    }
    /** Kelompok kunci yang harus dikirim/diterima bersama agar satu butir tetap konsisten. */
    expandSyncGroups(keys) {
      if (!this._groups) {
        this._groups = new Map();
        for (const q of this.schema.questions) {
          const g = [q.naKey].concat(q.options.map((o) => q.sheet + '!#' + o.name), q.boxes.map((b) => q.sheet + '!#' + b.name)).filter(Boolean);
          for (const k of g) this._groups.set(k, g);
        }
      }
      const out = new Set();
      for (const k of keys) { const g = this._groups.get(k); if (g) g.forEach((x) => out.add(x)); else out.add(k); }
      return Array.from(out);
    }
    inSyncGroup(k) { this.expandSyncGroups([]); return this._groups.has(k); }
    /** Terapkan perubahan dari server (orang lain) */
    applyRemote(changes) {
      const keys = [];
      for (const { k, v } of changes) {
        keys.push(k);
        if (k === '__app!finalized') {
          this.state.finalized = !!v; this.state.finalizedAt = typeof v === 'string' ? v : undefined;
        } else if (k.includes('!#')) this.state.controls[k] = !!v;
        else this.state.inputs[k] = v === undefined ? null : v;
      }
      this.applyVba();
      this.engine.invalidate();
      this.emit({ remote: true, keys });
    }

    // ---- nilai sel
    get(key) { const { sheet, addr } = splitKey(key); return this.engine.get(sheet, addr); }
    display(key) { const { sheet, addr } = splitKey(key); return this.engine.display(sheet, addr); }
    templateValue(key) {
      const { sheet, addr } = splitKey(key);
      const s = this.schema.byName.get(sheet);
      const c = s && s.cells[addr];
      return c && c.v !== undefined ? c.v : null;
    }
    inputText(key) {
      const v = this.get(key);
      if (E.isBlank(v) || E.isErr(v)) return '';
      if (typeof v === 'number') {
        const { sheet, addr } = splitKey(key);
        const s = this.schema.byName.get(sheet);
        const nf = s && s.cells[addr] && s.cells[addr].nf;
        if (nf && /d{1,2}[/.-]m{1,2}|mm-dd-yy/i.test(nf)) return E.formatDate(v, 'dd/mm/yyyy');
        return E.toText(v);
      }
      return String(v);
    }
    nf(key) { const { sheet, addr } = splitKey(key); const s = this.schema.byName.get(sheet); return s && s.cells[addr] && s.cells[addr].nf; }
    set(key, value, opts) {
      opts = opts || {};
      if (value === '' || value === undefined) value = null;
      this.state.inputs[key] = value;
      if (!opts.silent) this.afterChange({ key });
    }
    setTyped(key, text) { this.set(key, parseTyped(text, this.nf(key))); }
    control(sheet, name) { return !!this.state.controls[sheet + '!#' + name]; }
    setControl(sheet, name, val) { this.state.controls[sheet + '!#' + name] = !!val; }
    afterChange(info) {
      this.applyVba();
      this.engine.invalidate();
      this.emit(info);
    }

    // ---- logika VBA pertanyaan
    question(sheet) { return this.schema.questions.find((q) => q.sheet === sheet); }
    isNA(q) { return q.naKey ? this.get(q.naKey) === true : false; }
    setNA(q, on) {
      if (q.naKey) this.state.inputs[q.naKey] = !!on;
      if (on && q.naMacro) {
        // makro "...ReleventCB1_Click": kosongkan pilihan & set skor 0
        for (const n of q.naMacro.clearWhenOn) this.setControl(q.sheet, n, false);
        for (const s of q.naMacro.setWhenOn) this.state.inputs[q.sheet + '!' + s.cell] = s.value;
      }
      this.afterChange({ key: q.naKey });
    }
    selectOption(q, name) {
      if (this.isNA(q)) return;
      for (const o of q.options) this.setControl(q.sheet, o.name, o.name === name);
      this.afterChange({ key: q.scoreKey });
    }
    toggleBox(q, name, on) {
      if (this.isNA(q)) return;
      this.setControl(q.sheet, name, on);
      this.afterChange({ key: q.scoreKey });
    }
    selectedOption(q) { return q.options.find((o) => this.control(q.sheet, o.name)) || null; }
    /** Worksheet_Deactivate: tulis skor ke C6 (di Excel terjadi saat meninggalkan sheet) */
    applyVba() {
      for (const q of this.schema.questions) {
        let score;
        if (q.kind === 'radio') {
          const sel = this.selectedOption(q);
          if (sel) score = sel.score;
        } else {
          score = q.boxes.reduce((n, b) => n + (this.control(q.sheet, b.name) ? 1 : 0), 0);
        }
        if (this.isNA(q)) score = 0;
        if (score === undefined) score = this.state.inputs[q.scoreKey] !== undefined ? this.state.inputs[q.scoreKey] : this.templateValue(q.scoreKey);
        this.state.inputs[q.scoreKey] = score === null ? null : score;
      }
    }
    isAnswered(q) {
      if (q.kind === 'radio') return !!this.selectedOption(q);
      return q.boxes.some((b) => this.control(q.sheet, b.name));
    }
    /** Pesan isian wajib sesuai MsgBox VBA */
    messages(q) {
      const out = [];
      const empty = (key) => { const v = this.get(key); return E.isBlank(v) || v === ''; };
      for (const r of q.required) {
        if (r.when === 'na' && this.isNA(q) && empty(r.key)) out.push({ key: r.key, msg: r.msg });
      }
      if (this.isAnswered(q)) {
        // VBA memakai If ... ElseIf: hanya pesan pertama yang muncul; di sini semua ditampilkan berurutan
        for (const r of q.required) if (r.when === 'answered' && empty(r.key)) out.push({ key: r.key, msg: r.msg });
      }
      return out;
    }
    status(q) {
      if (this.isNA(q)) return this.messages(q).length ? 'warn' : 'na';
      if (!this.isAnswered(q)) return 'empty';
      return this.messages(q).length ? 'warn' : 'done';
    }
    progress() {
      const qs = this.schema.questions;
      let done = 0, warn = 0;
      for (const q of qs) { const s = this.status(q); if (s === 'done' || s === 'na') done++; if (s === 'warn') warn++; }
      return { total: qs.length, done, warn };
    }
    allMessages() {
      const out = [];
      for (const q of this.schema.questions) for (const m of this.messages(q)) out.push(Object.assign({ q }, m));
      return out;
    }

    // ---- multi-pilih dropdown (Worksheet_Change "Oldvalue & ", " & Newvalue")
    isMultiSelect(sheet, addr) {
      const s = this.schema.byName.get(sheet);
      return !!(s && s.multiSelect && s.multiSelect.some((r) => inSqref(r, addr)));
    }
    multiAdd(key, newValue) {
      if (!newValue) return;
      const old = this.inputText(key);
      let v;
      if (old === '') v = newValue;
      else if (old.indexOf(newValue) === -1) v = old + ', ' + newValue; // InStr(1, Oldvalue, Newvalue) = 0
      else v = old;
      this.set(key, v);
    }

    // ---- validasi
    validation(key) {
      const { sheet, addr } = splitKey(key);
      const v = findValidation(this.wb, sheet, addr);
      if (v && v.type === 'list') v._options = listOptions(this.engine, v, sheet) || [];
      return v;
    }
    options(key) { const v = this.validation(key); return v && v.type === 'list' ? v._options : null; }

    // ---- ekspor (setara makro Copy_Paste_Below_Last_Cell_*)
    exportTables() {
      const defs = [
        { name: 'PA attributes', src: 'PA attributes data sheet', lastCol: 'CZ' }, // makro: Range("A2:CZ")
        { name: 'Threat data', src: 'Threat data for export', lastCol: 'PT' },    // makro: Range("A2:PT")
        { name: 'Assessment data', src: 'Question_data_dump', lastCol: 'AAN' },   // makro: Range("A2:AAN")
      ];
      return defs.map((d) => {
        const last = E.colToNum(d.lastCol);
        const header = [], values = [];
        for (let c = 1; c <= last; c++) {
          const a = E.numToCol(c);
          const h = this.engine.get(d.src, a + '1');
          header.push(E.isBlank(h) ? null : h);
          const v = this.engine.get(d.src, a + '2');
          values.push(E.isBlank(v) ? null : v);
        }
        while (header.length && header[header.length - 1] === null && values[values.length - 1] === null) { header.pop(); values.pop(); }
        return { name: d.name, src: d.src, header, values };
      });
    }

    // ---- ringkasan untuk rekap
    summary() {
      const g = (s, a) => this.engine.get(s, a);
      const num = (v) => (typeof v === 'number' ? v : null);
      const els = [];
      for (let r = 6; r <= 11; r++) els.push({ element: E.toText(g('Dashboard', 'U' + r)), score: num(g('Dashboard', 'V' + r)), max: num(g('Dashboard', 'X' + r)), pct: num(g('Dashboard', 'Y' + r)) });
      return {
        name: E.toText(g('Atribut Kawasan', 'B5')),
        status: E.toText(g('Atribut Kawasan', 'B8')),
        authority: E.toText(g('Atribut Kawasan', 'B19')),
        year: g('Atribut Kawasan', 'B63'),
        date: g('Atribut Kawasan', 'B64'),
        total: num(g(DAFTAR, 'E51')), max: num(g(DAFTAR, 'D51')),
        elements: els,
      };
    }

    // ---- berkas
    toFile() {
      return JSON.stringify({
        app: APP_ID, schema: FILE_SCHEMA, source: this.wb.meta.source, id: this.state.id,
        created: this.state.created, updated: this.state.updated, savedAt: new Date().toISOString(),
        finalized: !!this.state.finalized, inputs: this.state.inputs, controls: this.state.controls,
      }, null, 1);
    }
    static parseFile(text) {
      const o = JSON.parse(text);
      if (!o || o.app !== APP_ID || !o.inputs) throw new Error('Berkas bukan file penilaian METT-4 web.');
      return o;
    }
    fileName() {
      let n = E.toText(this.engine.get('Tentang', 'D29'));
      if (!n || /^Contoh:/i.test(n)) {
        const nm = E.toText(this.engine.get('Atribut Kawasan', 'B5'));
        const yr = E.toText(this.engine.get('Atribut Kawasan', 'B63'));
        n = (nm && nm !== '0' ? nm : 'METT4') + (yr && yr !== '0' ? '_' + yr : '');
      }
      return n.replace(/[\\/:*?"<>|]+/g, '_').trim() || 'METT4';
    }
  }

  const api = { Model, buildSchema, findValidation, listOptions, checkValidation, parseTyped, inSqref, splitKey, parseRef, decodeX, newAssessment, APP_ID, DAFTAR };
  root.MettModel = api;
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
})(typeof window !== 'undefined' ? window : globalThis);
