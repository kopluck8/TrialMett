/*
 * Uji 3 (opsional): uji antarmuka di browser (Playwright + Chromium).
 *   python3 -m http.server 8765   (di folder repo)
 *   node tests/e2e.js http://localhost:8765/ tests/out
 */
const { chromium } = require(process.env.PLAYWRIGHT || 'playwright');
const fs = require('fs'), path = require('path');
(async () => {
  const base = (process.argv[2] || 'http://localhost:8765/') + 'index.html#/';
  const out = process.argv[3] || path.join(__dirname, 'out');
  fs.mkdirSync(out, { recursive: true });
  const b = await chromium.launch();
  const ctx = await b.newContext({ acceptDownloads: true, viewport: { width: 1300, height: 900 } });
  const p = await ctx.newPage();
  const errs = []; p.on('pageerror', (e) => errs.push(e.message));
  const ok = (name, c) => { console.log((c ? 'LULUS ' : 'GAGAL ') + name); if (!c) process.exitCode = 1; };
  await p.goto(base + 's/Atribut%20Kawasan');
  await p.fill('input[data-key="Atribut Kawasan!B5"]', 'TN Uji Browser'); await p.press('input[data-key="Atribut Kawasan!B5"]', 'Tab');
  await p.fill('input[data-key="Atribut Kawasan!B63"]', '2025'); await p.press('input[data-key="Atribut Kawasan!B63"]', 'Tab');
  await p.selectOption('select[data-key="Atribut Kawasan!B8"]', { index: 1 });
  await p.fill('input[data-key="Atribut Kawasan!B62"]', 'salah@email'); await p.press('input[data-key="Atribut Kawasan!B62"]', 'Tab');
  ok('validasi email menolak isian tidak valid', await p.isVisible('.field-err'));
  await p.goto(base + 's/3.');
  await p.locator('input[data-opt]').nth(3).check(); await p.waitForTimeout(200);
  ok('skor butir 3 = 3', (await p.textContent('.score-pill strong')).trim() === '3');
  await p.goto(base + 's/%20Penilaian%20Ancaman');
  await p.click('[data-act="expand-all"]');
  await p.selectOption('select[data-key=" Penilaian Ancaman!E3"]', 'Tinggi');
  await p.goto(base + 's/Dashboard'); await p.waitForTimeout(300);
  ok('dashboard menampilkan nama kawasan', (await p.textContent('h1.page-h')).includes('TN Uji Browser'));
  ok('dashboard total skor 3', (await p.textContent('[data-f="Dashboard!V11"]')).trim() === '3');
  await p.goto(base + 'data');
  const [d1] = await Promise.all([p.waitForEvent('download'), p.click('[data-act="export-xlsx"]')]);
  const xf = path.join(out, 'ekspor_ui.xlsx'); await d1.saveAs(xf);
  const [d2] = await Promise.all([p.waitForEvent('download'), p.click('#btnSave')]);
  const jf = path.join(out, 'penilaian_ui.json'); await d2.saveAs(jf);
  ok('file JSON tersimpan', JSON.parse(fs.readFileSync(jf, 'utf8')).inputs['Atribut Kawasan!B5'] === 'TN Uji Browser');
  await p.goto(base + 'rekap');
  await p.setInputFiles('#rekapFiles', [jf]); await p.waitForTimeout(300);
  ok('rekap memuat file', (await p.textContent('#main table')).includes('TN Uji Browser'));
  await p.goto(base + 's/Print%20BA'); await p.waitForTimeout(300);
  await p.pdf({ path: path.join(out, 'berita_acara.pdf'), format: 'A4' }).catch(() => {});
  ok('tidak ada error JavaScript', errs.length === 0);
  if (errs.length) console.log(errs);
  console.log('XLSX:', xf);
  await b.close();
})();
