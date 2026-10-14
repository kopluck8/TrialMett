/*
 * Membuat skenario isian uji + nilai harapan dari mesin rumus web.
 *   node tests/make_scenarios.js tests/out
 * Menghasilkan:
 *   tests/out/scenarios.json  -> isian per skenario (alamat sel Excel -> nilai)
 *   tests/out/expected.json   -> hasil semua rumus menurut mesin web
 * Lalu jalankan tests/compare_with_excel.py untuk menghitung skenario yang sama
 * dengan LibreOffice (file .xlsm asli) dan membandingkan hasilnya.
 */
const fs = require('fs');
const path = require('path');
global.window = global;
require('../data/workbook.js');
const E = require('../js/engine.js');
const M = require('../js/model.js');
const WB = global.METT_WB;

const outDir = process.argv[2] || path.join(__dirname, 'out');
fs.mkdirSync(outDir, { recursive: true });

let seed = 20261014;
const rnd = () => { seed = (seed * 1103515245 + 12345) % 2147483648; return seed / 2147483648; };
const pick = (arr) => arr[Math.floor(rnd() * arr.length)];

const ATTR = 'Atribut Kawasan', THREAT = ' Penilaian Ancaman', DAFTAR = 'Daftar Pertanyaan & Skor';
const LEVELS = ['Sangat Tinggi', 'Tinggi', 'Sedang', 'Rendah atau tidak ada informasi', 'Tidak ada ancaman'];

function scenario(name, mode) {
  const m = new M.Model(WB);
  m.load(M.newAssessment());
  const set = (k, v) => m.set(k, v, { silent: true });
  if (mode !== 'empty') {
    set(ATTR + '!B5', 'TN Uji ' + name);
    set(ATTR + '!B19', 'Balai Taman Nasional Uji');
    set(ATTR + '!B20', 12500.5);
    set(ATTR + '!B26', 1500000000);
    set(ATTR + '!B27', 250000000);
    set(ATTR + '!B59', 'Penanggung Jawab ' + name);
    set(ATTR + '!B60', 'Balai Uji');
    set(ATTR + '!B63', 2025);
    set(ATTR + '!B64', E.dateToSerial(2025, 8, 17));
    for (let r = 48; r <= 56; r++) set(ATTR + '!B' + r, Math.floor(rnd() * 6));
    for (let i = 0; i < 5; i++) set(ATTR + '!B' + (31 + i), 'Nilai penting ' + (i + 1) + ' ' + name);
    set(ATTR + '!B8', 'Taman Nasional');
    set('Aksi meningkatkan Pengelolaan!F5', E.dateToSerial(2026, 3, 1));
    set('Aksi meningkatkan Pengelolaan!I5', 75000000);
    set('Aksi meningkatkan Pengelolaan!G6', 'Kepala Seksi');
  }
  // pertanyaan
  for (const q of m.schema.questions) {
    if (mode === 'empty') break;
    const na = mode === 'allna' ? true : mode === 'max' ? false : rnd() < 0.12;
    if (na) { m.setNA(q, true); set(q.naReasonKey, 'Tidak relevan karena uji'); continue; }
    if (q.kind === 'radio') {
      const o = mode === 'max' ? q.options[q.options.length - 1] : pick(q.options);
      m.selectOption(q, o.name);
    } else {
      for (const b of q.boxes) if (mode === 'max' || rnd() < 0.5) m.toggleBox(q, b.name, true);
    }
    set(q.justKey, 'Bukti ' + q.sheet);
    set(q.actionKey, 'Aksi ' + q.sheet);
    for (const s of q.sources) if (rnd() < 0.4) { set(s.key, true); if (s.textKey) set(s.textKey, 'Sumber ' + s.label); }
    if (rnd() < 0.7) set(DAFTAR + '!F' + q.daftarRow, Math.floor(rnd() * 4));
  }
  // ancaman
  if (mode !== 'empty') {
    for (let r = 3; r <= 64; r++) {
      if (mode !== 'max' && rnd() < 0.45) continue;
      set(THREAT + '!E' + r, pick(LEVELS));
      if (rnd() < 0.85) set(THREAT + '!F' + r, pick(LEVELS));
      set(THREAT + '!C' + r, 'Deskripsi ancaman baris ' + r);
      if (r <= 62) set(THREAT + '!D' + r, 'Nilai penting 1 ' + name + (rnd() < 0.3 ? ', Nilai penting 2 ' + name : ''));
    }
    // detail
    const cond = ['Sangat baik', 'Baik', 'Cukup', 'Kurang', 'Tidak Tahu'], tren = ['Meningkat', 'Stabil', 'Menurun', 'Tidak tahu'];
    for (let r = 19; r <= 23; r++) { set('Detail penilaian nilai utama!B' + r, pick(cond)); set('Detail penilaian nilai utama!C' + r, pick(tren)); }
    for (let r = 12; r <= 16; r++) {
      set('Detail Penilaian Spesies!A' + r, 'Spesies ' + r);
      for (const c of 'BCDEFG') set('Detail Penilaian Spesies!' + c + r, pick(['Meningkat', 'Stabil', 'Menurun', 'Membaik']));
    }
    for (let r = 11; r <= 15; r++) {
      set('Detail Penilaian Habitat!A' + r, 'Habitat ' + r);
      for (const c of 'BCDE') set('Detail Penilaian Habitat!' + c + r, pick(['Meningkat', 'Stabil', 'Menurun', 'Membaik']));
    }
  }
  m.afterChange({});
  // nilai harapan semua rumus
  const expected = {};
  for (const s of WB.sheets) for (const [a, c] of Object.entries(s.cells)) {
    if (!c.f) continue;
    const v = m.engine.get(s.name, a);
    (expected[s.name] = expected[s.name] || {})[a] = v instanceof E.XlError ? { error: v.error } : v;
  }
  return { name, inputs: m.state.inputs, expected, summary: m.summary() };
}

const list = [scenario('kosong', 'empty'), scenario('semua-na', 'allna'), scenario('maksimal', 'max'),
  scenario('acak-1', 'random'), scenario('acak-2', 'random'), scenario('acak-3', 'random')];
fs.writeFileSync(path.join(outDir, 'scenarios.json'), JSON.stringify(list.map((s) => ({ name: s.name, inputs: s.inputs })), null, 0));
fs.writeFileSync(path.join(outDir, 'expected.json'), JSON.stringify(Object.fromEntries(list.map((s) => [s.name, s.expected]))));
for (const s of list) console.log(s.name.padEnd(10), 'total', s.summary.total, '/', s.summary.max, s.summary.elements.map((e) => e.element + '=' + (e.pct == null ? '-' : (e.pct * 100).toFixed(2) + '%')).join(' '));
