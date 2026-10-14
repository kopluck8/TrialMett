/*
 * METT-4 Web — Mesin Rumus (formula engine)
 * ------------------------------------------------------------------
 * Mengevaluasi RUMUS EXCEL ASLI (teks rumus disalin verbatim dari file .xlsm
 * ke data/workbook.js) di atas "sel virtual". Tidak ada rumus yang ditulis
 * ulang secara manual: aplikasi membaca string rumus seperti
 *     =COUNTIF(' Penilaian Ancaman'!E3:E5,"*")-(COUNTIF(...,"tidak ada ancaman"))
 * lalu menghitungnya dengan aturan Excel.
 *
 * Fungsi Excel yang dipakai file METT-4 KSDAE v4.4 dan didukung di sini:
 *   SUM, IF, COUNTIF, CONCAT (_xlfn.CONCAT), TEXT
 * Operator: + - * / ^ & % = <> < > <= >=, referensi sel/rentang lintas sheet,
 * dan structured reference Tabel Excel (Table2[[#This Row],[Kolom]]).
 *
 * Semantik Excel yang direplikasi:
 *   - sel kosong = 0 dalam aritmetika, "" dalam penggabungan teks,
 *     dan rumus yang hasil akhirnya merujuk sel kosong menghasilkan 0;
 *   - SUM atas rentang mengabaikan teks/boolean;
 *   - COUNTIF dengan wildcard (* ?), tidak peka huruf besar/kecil;
 *   - pembagian dengan nol -> #DIV/0!, teks non-angka dalam aritmetika -> #VALUE!;
 *   - galat (error) merambat ke rumus yang memakainya.
 */
(function (root) {
  'use strict';

  // ------------------------------------------------------------------ nilai khusus
  const BLANK = Object.freeze({ blank: true, toString() { return ''; } });
  class XlError { constructor(code) { this.error = code; } toString() { return this.error; } }
  const ERR = {
    DIV0: new XlError('#DIV/0!'), VALUE: new XlError('#VALUE!'), REF: new XlError('#REF!'),
    NAME: new XlError('#NAME?'), NA: new XlError('#N/A'), CIRC: new XlError('#CIRC!'), NUM: new XlError('#NUM!'),
  };
  const isErr = (v) => v instanceof XlError;
  const isBlank = (v) => v === BLANK || v === undefined || v === null;

  // ------------------------------------------------------------------ alamat sel
  function colToNum(col) {
    let n = 0;
    for (const ch of col.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
    return n;
  }
  function numToCol(n) {
    let s = '';
    while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
    return s;
  }
  function parseA1(a1) {
    const m = /^\$?([A-Za-z]{1,3})\$?(\d+)$/.exec(a1);
    if (!m) return null;
    return { c: colToNum(m[1]), r: parseInt(m[2], 10) };
  }
  const toA1 = (c, r) => numToCol(c) + r;

  // ------------------------------------------------------------------ tokenizer
  const RE_A1 = /^\$?[A-Za-z]{1,3}\$?\d+/;
  function tokenize(src) {
    let s = src.trim();
    if (s.startsWith('=')) s = s.slice(1);
    const toks = [];
    let i = 0;
    const peek = (k = 0) => s[i + k];
    while (i < s.length) {
      const ch = s[i];
      if (ch === ' ' || ch === '\n' || ch === '\r' || ch === '\t') { i++; continue; }
      if (ch === '"') {
        let j = i + 1, out = '';
        while (j < s.length) {
          if (s[j] === '"') { if (s[j + 1] === '"') { out += '"'; j += 2; continue; } break; }
          out += s[j++];
        }
        toks.push({ t: 'str', v: out }); i = j + 1; continue;
      }
      if (ch === '#') {
        const m = /^#(DIV\/0!|VALUE!|REF!|NAME\?|N\/A|NUM!|NULL!)/.exec(s.slice(i));
        if (m) { toks.push({ t: 'err', v: m[0] }); i += m[0].length; continue; }
      }
      if (/[0-9.]/.test(ch) && !(RE_A1.test(s.slice(i)))) {
        const m = /^(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?/.exec(s.slice(i));
        if (m) { toks.push({ t: 'num', v: parseFloat(m[0]) }); i += m[0].length; continue; }
      }
      if (ch === "'" || /[A-Za-z_$\\]/.test(ch)) {
        // referensi bersheet: 'Nama Sheet'!A1  atau  Sheet!A1
        let sheet = null, j = i;
        if (ch === "'") {
          let k = i + 1, name = '';
          while (k < s.length) {
            if (s[k] === "'") { if (s[k + 1] === "'") { name += "'"; k += 2; continue; } break; }
            name += s[k++];
          }
          if (s[k + 1] === '!') { sheet = name; j = k + 2; }
        } else {
          const m = /^([A-Za-z0-9_.\\]+)!/.exec(s.slice(i));
          if (m) { sheet = m[1]; j = i + m[0].length; }
        }
        if (sheet !== null) {
          const rest = s.slice(j);
          if (rest.startsWith('#REF!')) { toks.push({ t: 'err', v: '#REF!' }); i = j + 5; continue; }
          const m1 = RE_A1.exec(rest);
          if (!m1) throw new Error('Referensi tidak dikenali: ' + s.slice(i));
          let a1 = m1[0], a2 = null, len = m1[0].length;
          if (rest[len] === ':') {
            const m2 = RE_A1.exec(rest.slice(len + 1));
            if (m2) { a2 = m2[0]; len += 1 + m2[0].length; }
          }
          toks.push({ t: 'ref', sheet, a1, a2 }); i = j + len; continue;
        }
        // nama: fungsi, boolean, structured ref, atau referensi A1 tanpa sheet
        const m = /^[A-Za-z_\\][A-Za-z0-9_.\\]*/.exec(s.slice(i)) || /^\$?[A-Za-z]{1,3}\$?\d+/.exec(s.slice(i));
        const word = m[0];
        const after = s[i + word.length];
        if (after === '(') { toks.push({ t: 'func', v: word.toUpperCase().replace(/^_XLFN\./, '').replace(/^_XLWS\./, '') }); i += word.length; continue; }
        if (after === '[') {
          let depth = 0, k = i + word.length;
          for (; k < s.length; k++) { if (s[k] === '[') depth++; else if (s[k] === ']') { depth--; if (depth === 0) break; } }
          toks.push({ t: 'sref', table: word, spec: s.slice(i + word.length, k + 1) }); i = k + 1; continue;
        }
        const ma1 = RE_A1.exec(s.slice(i));
        if (ma1) {
          let a1 = ma1[0], a2 = null, len = a1.length;
          if (s[i + len] === ':') {
            const m2 = RE_A1.exec(s.slice(i + len + 1));
            if (m2) { a2 = m2[0]; len += 1 + m2[0].length; }
          }
          toks.push({ t: 'ref', sheet: null, a1, a2 }); i += len; continue;
        }
        const up = word.toUpperCase();
        if (up === 'TRUE' || up === 'FALSE') { toks.push({ t: 'bool', v: up === 'TRUE' }); i += word.length; continue; }
        toks.push({ t: 'name', v: word }); i += word.length; continue;
      }
      const two = s.slice(i, i + 2);
      if (two === '<>' || two === '<=' || two === '>=') { toks.push({ t: 'op', v: two }); i += 2; continue; }
      if ('+-*/^&=<>%'.includes(ch)) { toks.push({ t: 'op', v: ch }); i++; continue; }
      if (ch === '(' || ch === ')' || ch === ',' || ch === ';') { toks.push({ t: ch === ';' ? ',' : ch }); i++; continue; }
      throw new Error('Karakter tidak dikenali "' + ch + '" pada rumus: ' + src);
    }
    return toks;
  }

  // ------------------------------------------------------------------ parser
  function parse(src) {
    const toks = tokenize(src);
    let p = 0;
    const peek = () => toks[p];
    const next = () => toks[p++];
    const isOp = (v) => peek() && peek().t === 'op' && (Array.isArray(v) ? v.includes(peek().v) : peek().v === v);

    function comparison() {
      let left = concat();
      while (isOp(['=', '<>', '<', '>', '<=', '>='])) { const op = next().v; left = { k: 'bin', op, a: left, b: concat() }; }
      return left;
    }
    function concat() {
      let left = additive();
      while (isOp('&')) { next(); left = { k: 'bin', op: '&', a: left, b: additive() }; }
      return left;
    }
    function additive() {
      let left = multiplicative();
      while (isOp(['+', '-'])) { const op = next().v; left = { k: 'bin', op, a: left, b: multiplicative() }; }
      return left;
    }
    function multiplicative() {
      let left = power();
      while (isOp(['*', '/'])) { const op = next().v; left = { k: 'bin', op, a: left, b: power() }; }
      return left;
    }
    function power() {
      let left = unary();
      while (isOp('^')) { next(); left = { k: 'bin', op: '^', a: left, b: unary() }; }
      return left;
    }
    function unary() {
      if (isOp('-')) { next(); return { k: 'neg', a: unary() }; }
      if (isOp('+')) { next(); return unary(); }
      return postfix();
    }
    function postfix() {
      let e = primary();
      while (isOp('%')) { next(); e = { k: 'pct', a: e }; }
      return e;
    }
    function primary() {
      const t = next();
      if (!t) throw new Error('Rumus terpotong: ' + src);
      switch (t.t) {
        case 'num': return { k: 'lit', v: t.v };
        case 'str': return { k: 'lit', v: t.v };
        case 'bool': return { k: 'lit', v: t.v };
        case 'err': return { k: 'lit', v: new XlError(t.v) };
        case 'ref': return { k: 'ref', sheet: t.sheet, a1: t.a1, a2: t.a2 };
        case 'sref': return { k: 'sref', table: t.table, spec: t.spec };
        case 'name': return { k: 'name', v: t.v };
        case '(': { const e = comparison(); if (!next() || toks[p - 1].t !== ')') throw new Error('Kurung tidak seimbang: ' + src); return e; }
        case 'func': {
          next(); // (
          const args = [];
          if (peek() && peek().t === ')') { next(); return { k: 'call', name: t.v, args }; }
          for (;;) {
            if (peek() && (peek().t === ',' || peek().t === ')')) args.push({ k: 'missing' });
            else args.push(comparison());
            const sep = next();
            if (!sep) throw new Error('Rumus terpotong: ' + src);
            if (sep.t === ')') break;
            if (sep.t !== ',') throw new Error('Diharapkan koma: ' + src);
          }
          return { k: 'call', name: t.v, args };
        }
        default: throw new Error('Token tak terduga ' + JSON.stringify(t) + ' pada: ' + src);
      }
    }
    const ast = comparison();
    if (p < toks.length) throw new Error('Sisa token tak terduga pada: ' + src);
    return ast;
  }

  // ------------------------------------------------------------------ koersi
  const NUM_RE = /^[+-]?(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?$/;
  function toNumber(v) {
    if (isErr(v)) return v;
    if (isBlank(v)) return 0;
    if (typeof v === 'number') return v;
    if (typeof v === 'boolean') return v ? 1 : 0;
    if (typeof v === 'string') {
      const s = v.trim();
      if (NUM_RE.test(s)) return parseFloat(s);
      if (/^[+-]?(\d+\.?\d*|\.\d+)%$/.test(s)) return parseFloat(s) / 100;
      return ERR.VALUE;
    }
    return ERR.VALUE;
  }
  function numberToText(n) {
    // Format "General" Excel (maks. 15 digit signifikan), desimal koma (lokal Indonesia)
    if (!isFinite(n)) return ERR.NUM;
    if (Number.isInteger(n) && Math.abs(n) < 1e15) return String(n);
    let s = String(parseFloat(n.toPrecision(15)));
    if (/e/.test(s)) s = n.toExponential(9).replace(/\.?0+e/, 'E').replace('e', 'E');
    return s.replace('.', ',');
  }
  function toText(v) {
    if (isErr(v)) return v;
    if (isBlank(v)) return '';
    if (typeof v === 'number') return numberToText(v);
    if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
    return String(v);
  }
  function toBool(v) {
    if (isErr(v)) return v;
    if (isBlank(v)) return false;
    if (typeof v === 'boolean') return v;
    if (typeof v === 'number') return v !== 0;
    if (typeof v === 'string') {
      const u = v.trim().toUpperCase();
      if (u === 'TRUE') return true;
      if (u === 'FALSE') return false;
      return ERR.VALUE;
    }
    return ERR.VALUE;
  }
  const typeRank = (v) => (typeof v === 'number' ? 1 : typeof v === 'string' ? 2 : typeof v === 'boolean' ? 3 : 0);
  function compare(a, b) {
    if (isBlank(a) && isBlank(b)) return 0;
    if (isBlank(a)) a = typeof b === 'string' ? '' : typeof b === 'boolean' ? false : 0;
    if (isBlank(b)) b = typeof a === 'string' ? '' : typeof a === 'boolean' ? false : 0;
    const ra = typeRank(a), rb = typeRank(b);
    if (ra !== rb) return ra < rb ? -1 : 1;
    if (ra === 2) { const x = a.toLowerCase(), y = b.toLowerCase(); return x < y ? -1 : x > y ? 1 : 0; }
    if (ra === 3) return (a === b) ? 0 : (a ? 1 : -1);
    return a < b ? -1 : a > b ? 1 : 0;
  }

  // ------------------------------------------------------------------ tanggal (serial Excel)
  const HARI = ['Minggu', 'Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat', 'Sabtu'];
  const BULAN = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'];
  function serialToParts(serial) {
    const s = Math.floor(serial);
    const weekday = ((s + 6) % 7 + 7) % 7; // 0 = Minggu (sama dengan WEEKDAY Excel, termasuk serial 0 = Sabtu)
    if (s === 0) return { y: 1900, m: 1, d: 0, weekday };
    if (s === 60) return { y: 1900, m: 2, d: 29, weekday }; // tanggal fiktif Excel (bug 1900)
    const base = s < 60 ? Date.UTC(1899, 11, 31) : Date.UTC(1899, 11, 30);
    const dt = new Date(base + s * 86400000);
    return { y: dt.getUTCFullYear(), m: dt.getUTCMonth() + 1, d: dt.getUTCDate(), weekday };
  }
  function dateToSerial(y, m, d) {
    const ms = Date.UTC(y, m - 1, d) - Date.UTC(1899, 11, 30);
    let s = Math.round(ms / 86400000);
    if (s < 61) s -= 1; // sebelum 1 Mar 1900 (bug tahun kabisat 1900 di Excel)
    return s;
  }
  const pad = (n, w = 2) => String(n).padStart(w, '0');
  function formatDate(serial, fmt) {
    const p = serialToParts(serial);
    return fmt.replace(/dddd|ddd|dd|d|mmmm|mmm|mm|m|yyyy|yy/gi, (tok) => {
      switch (tok.toLowerCase()) {
        case 'dddd': return HARI[p.weekday];
        case 'ddd': return HARI[p.weekday].slice(0, 3);
        case 'dd': return pad(p.d);
        case 'd': return String(p.d);
        case 'mmmm': return BULAN[p.m - 1];
        case 'mmm': return BULAN[p.m - 1].slice(0, 3);
        case 'mm': return pad(p.m);
        case 'm': return String(p.m);
        case 'yyyy': return String(p.y);
        case 'yy': return pad(p.y % 100);
        default: return tok;
      }
    });
  }

  // ------------------------------------------------------------------ format angka (tampilan)
  function groupId(intStr) { return intStr.replace(/\B(?=(\d{3})+(?!\d))/g, '.'); }
  function fixedId(n, dec) {
    const s = Math.abs(n).toFixed(dec);
    const [i, f] = s.split('.');
    return groupId(i) + (f ? ',' + f : '');
  }
  function formatValue(v, nf) {
    if (isErr(v)) return v.error;
    if (isBlank(v)) return '';
    if (nf === ';;;') return '';
    if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
    if (typeof v === 'string') return v;
    if (typeof v !== 'number') return String(v);
    if (!nf || nf === 'General' || nf === '@') return numberToText(v);
    if (nf === '0') return String(Math.round(v));
    if (nf === '0%') return Math.round(v * 100) + '%';
    if (nf === '0.00%') return (v < 0 ? '-' : '') + fixedId(v * 100, 2) + '%';
    if (/Rp/.test(nf) && /"-"/.test(nf)) { // format akuntansi Rupiah
      if (Math.abs(v) < 0.005) return 'Rp -';
      return (v < 0 ? '-' : '') + 'Rp ' + fixedId(v, 2);
    }
    if (/Rp/.test(nf)) return (v < 0 ? '-' : '') + 'Rp' + fixedId(v, 0);
    if (nf === 'mm-dd-yy' || /d{1,2}[/.-]m{1,2}[/.-]y{2,4}/i.test(nf) || /m{1,2}[/.-]d{1,2}[/.-]y{2,4}/i.test(nf)) {
      return formatDate(v, 'dd/mm/yyyy');
    }
    if (/^0\.0+$/.test(nf)) return fixedId(v, nf.split('.')[1].length);
    if (/#,##0/.test(nf)) return (v < 0 ? '-' : '') + fixedId(v, (nf.split('.')[1] || '').replace(/[^0]/g, '').length);
    return numberToText(v);
  }

  // ------------------------------------------------------------------ COUNTIF
  function wildcardToRegex(s) {
    let out = '';
    for (let i = 0; i < s.length; i++) {
      const ch = s[i];
      if (ch === '~' && (s[i + 1] === '*' || s[i + 1] === '?' || s[i + 1] === '~')) { out += '\\' + s[++i]; continue; }
      if (ch === '*') out += '[\\s\\S]*';
      else if (ch === '?') out += '[\\s\\S]';
      else out += ch.replace(/[.+^${}()|[\]\\]/g, '\\$&');
    }
    return new RegExp('^' + out + '$', 'i');
  }
  function makeCriteria(crit) {
    if (isErr(crit)) return () => false;
    if (typeof crit === 'number') return (v) => (typeof v === 'number' && v === crit) || (typeof v === 'string' && NUM_RE.test(v.trim()) && parseFloat(v) === crit);
    if (typeof crit === 'boolean') return (v) => v === crit;
    if (isBlank(crit)) crit = '';
    const s = String(crit);
    const m = /^(<=|>=|<>|<|>|=)?([\s\S]*)$/.exec(s);
    const op = m[1] || '', rest = m[2];
    const numeric = NUM_RE.test(rest.trim()) && rest.trim() !== '';
    const hasWild = /[*?]/.test(rest);
    const eq = (v) => {
      if (rest === '') return isBlank(v) || v === '';
      if (numeric && !hasWild) {
        const n = parseFloat(rest);
        return (typeof v === 'number' && v === n) || (typeof v === 'string' && v.trim() !== '' && NUM_RE.test(v.trim()) && parseFloat(v) === n);
      }
      if (/^(TRUE|FALSE)$/i.test(rest) && typeof v === 'boolean') return v === (rest.toUpperCase() === 'TRUE');
      if (typeof v !== 'string') return false;
      return hasWild ? wildcardToRegex(rest).test(v) : v.toLowerCase() === rest.toLowerCase();
    };
    if (op === '' || op === '=') return eq;
    if (op === '<>') return (v) => !eq(v);
    return (v) => {
      if (isBlank(v)) return false;
      let c;
      if (numeric) { if (typeof v !== 'number') return false; c = v - parseFloat(rest); }
      else { if (typeof v !== 'string') return false; c = compare(v, rest); }
      return op === '<' ? c < 0 : op === '>' ? c > 0 : op === '<=' ? c <= 0 : c >= 0;
    };
  }

  // ------------------------------------------------------------------ Engine
  class Range {
    constructor(sheet, c1, r1, c2, r2) { this.sheet = sheet; this.c1 = Math.min(c1, c2); this.r1 = Math.min(r1, r2); this.c2 = Math.max(c1, c2); this.r2 = Math.max(r1, r2); }
    get single() { return this.c1 === this.c2 && this.r1 === this.r2; }
  }

  class Engine {
    /**
     * @param {object} wb   data workbook hasil extract.py (window.METT_WB)
     * @param {object} inputs  peta "Sheet!A1" -> nilai isian pengguna (menimpa nilai bawaan sel)
     */
    constructor(wb, inputs) {
      this.wb = wb;
      this.sheetMap = new Map();
      this.sheetMapLower = new Map();
      for (const s of wb.sheets) { this.sheetMap.set(s.name, s); this.sheetMapLower.set(s.name.toLowerCase(), s); }
      this.inputs = inputs || {};
      this.cache = new Map();
      this.inProgress = new Set();
      this.astCache = new Map();
    }
    setInputs(inputs) { this.inputs = inputs; this.invalidate(); }
    invalidate() { this.cache.clear(); }
    sheet(name) {
      return this.sheetMap.get(name) || this.sheetMapLower.get(String(name).toLowerCase()) || null;
    }
    /** Nilai sel (tipe asli: number/string/boolean/BLANK/XlError) */
    get(sheetName, addr) {
      addr = addr.replace(/\$/g, '').toUpperCase();
      const key = sheetName + '!' + addr;
      if (Object.prototype.hasOwnProperty.call(this.inputs, key)) {
        const v = this.inputs[key];
        return (v === null || v === undefined || v === '') ? BLANK : v;
      }
      const sh = this.sheet(sheetName);
      if (!sh) return ERR.REF;
      const cell = sh.cells[addr];
      if (!cell) return BLANK;
      if (cell.f === undefined) return cell.v === undefined || cell.v === null ? BLANK : cell.v;
      const ck = sh.name + '!' + addr;
      if (this.cache.has(ck)) return this.cache.get(ck);
      if (this.inProgress.has(ck)) return ERR.CIRC;
      this.inProgress.add(ck);
      let val;
      try {
        val = this.evalFormula(sh.name, addr, cell.f);
      } catch (e) {
        if (typeof console !== 'undefined') console.error('Gagal menghitung', ck, cell.f, e);
        val = ERR.NAME;
      }
      this.inProgress.delete(ck);
      this.cache.set(ck, val);
      return val;
    }
    /** Teks tampilan sesuai format angka sel di Excel */
    display(sheetName, addr) {
      const sh = this.sheet(sheetName);
      const cell = sh && sh.cells[addr.replace(/\$/g, '')];
      return formatValue(this.get(sheetName, addr), cell && cell.nf);
    }
    formula(sheetName, addr) {
      const sh = this.sheet(sheetName);
      const c = sh && sh.cells[addr];
      return c && c.f;
    }
    ast(f) {
      let a = this.astCache.get(f);
      if (!a) { a = parse(f); this.astCache.set(f, a); }
      return a;
    }
    evalFormula(sheetName, addr, f) {
      const pos = parseA1(addr);
      const ctx = { sheet: sheetName, row: pos.r, col: pos.c };
      let v = this.eval(this.ast(f), ctx);
      v = this.scalar(v, ctx);
      if (isBlank(v)) v = 0; // Excel: =A1 dengan A1 kosong -> 0
      return v;
    }
    scalar(v, ctx) {
      if (v instanceof Range) {
        if (v.single) return this.get(v.sheet, toA1(v.c1, v.r1));
        // irisan implisit
        if (v.c1 === v.c2 && ctx.row >= v.r1 && ctx.row <= v.r2) return this.get(v.sheet, toA1(v.c1, ctx.row));
        if (v.r1 === v.r2 && ctx.col >= v.c1 && ctx.col <= v.c2) return this.get(v.sheet, toA1(ctx.col, v.r1));
        return ERR.VALUE;
      }
      return v;
    }
    *cells(range) {
      for (let r = range.r1; r <= range.r2; r++)
        for (let c = range.c1; c <= range.c2; c++) yield this.get(range.sheet, toA1(c, r));
    }
    resolveSheet(name, ctx) {
      if (name === null || name === undefined) return ctx.sheet;
      const sh = this.sheet(name);
      return sh ? sh.name : null;
    }
    eval(n, ctx) {
      switch (n.k) {
        case 'lit': return n.v;
        case 'missing': return BLANK;
        case 'ref': {
          const sh = this.resolveSheet(n.sheet, ctx);
          if (!sh) return ERR.REF;
          const p1 = parseA1(n.a1), p2 = n.a2 ? parseA1(n.a2) : p1;
          return new Range(sh, p1.c, p1.r, p2.c, p2.r);
        }
        case 'sref': return this.structured(n, ctx);
        case 'name': {
          const dn = this.wb.definedNames && this.wb.definedNames[n.v];
          if (dn) return this.eval(this.ast('=' + dn), ctx);
          return ERR.NAME;
        }
        case 'neg': { const v = toNumber(this.scalar(this.eval(n.a, ctx), ctx)); return isErr(v) ? v : -v; }
        case 'pct': { const v = toNumber(this.scalar(this.eval(n.a, ctx), ctx)); return isErr(v) ? v : v / 100; }
        case 'bin': return this.binary(n, ctx);
        case 'call': return this.call(n, ctx);
        default: throw new Error('Node tidak dikenal ' + n.k);
      }
    }
    binary(n, ctx) {
      const a = this.scalar(this.eval(n.a, ctx), ctx);
      const b = this.scalar(this.eval(n.b, ctx), ctx);
      if (n.op === '&') {
        const x = toText(a); if (isErr(x)) return x;
        const y = toText(b); if (isErr(y)) return y;
        return x + y;
      }
      if (['=', '<>', '<', '>', '<=', '>='].includes(n.op)) {
        if (isErr(a)) return a;
        if (isErr(b)) return b;
        const c = compare(a, b);
        switch (n.op) {
          case '=': return c === 0; case '<>': return c !== 0; case '<': return c < 0;
          case '>': return c > 0; case '<=': return c <= 0; default: return c >= 0;
        }
      }
      const x = toNumber(a); if (isErr(x)) return x;
      const y = toNumber(b); if (isErr(y)) return y;
      switch (n.op) {
        case '+': return x + y;
        case '-': return x - y;
        case '*': return x * y;
        case '/': return y === 0 ? ERR.DIV0 : x / y;
        case '^': return Math.pow(x, y);
        default: throw new Error('Operator tidak dikenal ' + n.op);
      }
    }
    structured(n, ctx) {
      const t = this.wb.tables && this.wb.tables[n.table];
      if (!t) return ERR.REF;
      const [a, b] = t.ref.split(':');
      const p1 = parseA1(a), p2 = parseA1(b);
      const parts = [];
      const re = /\[([^\[\]]*)\]/g;
      let m;
      const inner = n.spec.replace(/^\[/, '').replace(/\]$/, '');
      if (!/\[/.test(inner)) parts.push(inner); else while ((m = re.exec(inner))) parts.push(m[1]);
      let thisRow = false, colName = null;
      for (let p of parts) {
        p = p.trim();
        if (p.toLowerCase() === '#this row') thisRow = true;
        else if (p.startsWith('@')) { thisRow = true; colName = p.slice(1); }
        else colName = p;
      }
      const idx = t.columns.findIndex((c) => c.trim().toLowerCase() === String(colName).trim().toLowerCase());
      if (idx < 0) return ERR.REF;
      const col = p1.c + idx;
      const sh = this.resolveSheet(t.sheet, ctx);
      if (thisRow) return new Range(sh, col, ctx.row, col, ctx.row);
      return new Range(sh, col, p1.r + 1, col, p2.r);
    }
    call(n, ctx) {
      const name = n.name;
      const args = n.args;
      switch (name) {
        case 'SUM': {
          let total = 0;
          for (const a of args) {
            const v = this.eval(a, ctx);
            if (v instanceof Range) {
              for (const c of this.cells(v)) { if (isErr(c)) return c; if (typeof c === 'number') total += c; }
            } else {
              const x = toNumber(v); if (isErr(x)) return x; total += x;
            }
          }
          return total;
        }
        case 'IF': {
          const c = toBool(this.scalar(this.eval(args[0], ctx), ctx));
          if (isErr(c)) return c;
          if (c) return args.length > 1 ? (args[1].k === 'missing' ? 0 : this.eval(args[1], ctx)) : true;
          return args.length > 2 ? (args[2].k === 'missing' ? 0 : this.eval(args[2], ctx)) : false;
        }
        case 'COUNTIF': {
          const rng = this.eval(args[0], ctx);
          if (!(rng instanceof Range)) return ERR.VALUE;
          const crit = this.scalar(this.eval(args[1], ctx), ctx);
          const test = makeCriteria(crit);
          let count = 0;
          for (const c of this.cells(rng)) if (!isErr(c) && test(c)) count++;
          return count;
        }
        case 'CONCAT':
        case 'CONCATENATE': {
          let out = '';
          for (const a of args) {
            const v = this.eval(a, ctx);
            if (v instanceof Range) { for (const c of this.cells(v)) { const t = toText(c); if (isErr(t)) return t; out += t; } }
            else { const t = toText(v); if (isErr(t)) return t; out += t; }
          }
          return out;
        }
        case 'TEXT': {
          const v = this.scalar(this.eval(args[0], ctx), ctx);
          const fmt = toText(this.scalar(this.eval(args[1], ctx), ctx));
          if (isErr(v)) return v;
          if (isErr(fmt)) return fmt;
          const num = toNumber(v);
          if (isErr(num)) return typeof v === 'string' ? v : num;
          if (/[dmy]/i.test(fmt) && !/[#0]/.test(fmt)) return formatDate(num, fmt);
          return formatValue(num, fmt);
        }
        case 'ISNUMBER': return typeof this.scalar(this.eval(args[0], ctx), ctx) === 'number';
        case 'AND': {
          let r = true;
          for (const a of args) { const v = toBool(this.scalar(this.eval(a, ctx), ctx)); if (isErr(v)) return v; r = r && v; }
          return r;
        }
        default:
          return ERR.NAME;
      }
    }
  }

  const api = {
    Engine, BLANK, XlError, ERR, isErr, isBlank, parse, tokenize, formatValue, formatDate,
    serialToParts, dateToSerial, colToNum, numToCol, parseA1, toA1, toNumber, toText, makeCriteria,
  };
  root.MettEngine = api;
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
})(typeof window !== 'undefined' ? window : globalThis);
