/*
 * Uji 1: hasil mesin rumus web pada file kosong (template) == nilai yang terakhir
 * dihitung oleh Microsoft Excel sendiri (cache nilai di dalam file .xlsm).
 *   node tests/check_excel_cache.js
 */
global.window = global;
require('../data/workbook.js');
const E = require('../js/engine.js');
const cached = require('./fixtures/excel_cached_values.json');
const WB = global.METT_WB;
const eng = new E.Engine(WB, {});
const EN_ID = { Saturday: 'Sabtu', January: 'Januari' }; // file disimpan dengan Excel berlokal Inggris
let n = 0, bad = 0;
for (const s of WB.sheets) for (const [a, c] of Object.entries(s.cells)) {
  if (!c.f) continue;
  n++;
  let v = eng.get(s.name, a);
  v = v instanceof E.XlError ? v.error : v;
  let x = cached[s.name] ? cached[s.name][a] : undefined;
  if (typeof x === 'string') for (const [en, id] of Object.entries(EN_ID)) x = x.split(en).join(id);
  const ok = (x === null || x === undefined) ? (v === 0 || v === '' || v === false)
    : (typeof x === 'number' && typeof v === 'number') ? Math.abs(x - v) < 1e-9 : String(x) === String(v);
  if (!ok) { bad++; console.log('BEDA', s.name + '!' + a, 'web=', JSON.stringify(v), 'excel=', JSON.stringify(x)); }
}
console.log(`${n} rumus dibandingkan dengan nilai cache Excel: ${bad} berbeda`);
process.exit(bad ? 1 : 0);
