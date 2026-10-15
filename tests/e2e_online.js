/*
 * Uji 5: mode online dengan 3 "kelompok" (3 browser terpisah) mengisi bersamaan.
 *   node tests/sync_server.js 8770 8771 &
 *   node tests/e2e_online.js http://localhost:8770/ http://localhost:8771/
 */
const { chromium } = require(process.env.PLAYWRIGHT || 'playwright');
const APP = process.argv[2] || 'http://localhost:8770/';
const API = process.argv[3] || 'http://localhost:8771/';
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
let failed = 0;
const ok = (n, c) => { console.log((c ? 'LULUS ' : 'GAGAL ') + n); if (!c) failed++; };

(async () => {
  const b = await chromium.launch();
  const errs = [];
  const open = async (name) => {
    const ctx = await b.newContext({ viewport: { width: 1280, height: 860 } });
    const p = await ctx.newPage();
    p.on('pageerror', (e) => errs.push(name + ': ' + e.message));
    p.on('dialog', (d) => d.accept(d.type() === 'prompt' ? name : undefined));
    await p.goto(APP + 'index.html#/s/Tentang');
    await p.waitForFunction(() => window.METT_APP && document.querySelector('#saveState').textContent.includes('Tersinkron'), null, { timeout: 15000 });
    return p;
  };
  const go = (p, r) => p.goto(APP + 'index.html#/' + r).then(() => p.waitForTimeout(150));
  const val = (p, k) => p.evaluate((k) => { const v = window.METT_APP.model.get(k); return v && v.error ? v.error : (v && v.blank ? null : v); }, k);
  const settle = async (...ps) => { for (const p of ps) await p.evaluate(() => window.METT_APP.sync && 0); await sleep(5200); };

  const k1 = await open('Ani – Kelompok 1');
  const k2 = await open('Budi – Kelompok 2');
  const k3 = await open('Citra – Kelompok 3');
  ok('ketiga kelompok tersambung', true);

  // --- isian paralel pada bagian berbeda
  await go(k1, 's/Atribut%20Kawasan');
  await k1.fill('input[data-key="Atribut Kawasan!B5"]', 'TN Uji Bersama'); await k1.press('input[data-key="Atribut Kawasan!B5"]', 'Tab');
  await go(k1, 's/1.'); await k1.locator('input[data-opt]').nth(2).check();
  await go(k2, 's/2.'); await k2.locator('input[data-opt]').nth(3).check();
  await k2.fill('textarea[data-key="2.!B15"]', 'Bukti dari Kelompok 2');
  await go(k3, 's/%20Penilaian%20Ancaman'); await k3.click('[data-act="expand-all"]');
  await k3.selectOption('select[data-key=" Penilaian Ancaman!E3"]', 'Tinggi');
  await go(k3, 's/7A.'); await k3.locator('input[data-box]').nth(0).check(); await k3.locator('input[data-box]').nth(1).check();
  await settle(k1, k2, k3);
  for (const [n, p] of [['K1', k1], ['K2', k2], ['K3', k3]]) {
    ok(`${n} melihat nama kawasan`, await val(p, 'Atribut Kawasan!B5') === 'TN Uji Bersama');
    ok(`${n} total skor = 2 + 3 + 2 = 7`, await val(p, 'Daftar Pertanyaan & Skor!E51') === 7);
    ok(`${n} melihat ancaman E3 Tinggi`, await val(p, ' Penilaian Ancaman!E3') === 'Tinggi');
    ok(`${n} melihat bukti K2`, await val(p, '2.!B15') === 'Bukti dari Kelompok 2');
  }
  await go(k1, 's/Dashboard'); await k1.waitForTimeout(300);
  ok('Dashboard K1 menampilkan total 7', (await k1.textContent('[data-f="Dashboard!V11"]')).trim() === '7');

  // --- dua kelompok memilih jawaban berbeda pada butir yang sama hampir bersamaan
  await go(k1, 's/5.'); await go(k2, 's/5.');
  await Promise.all([k1.locator('input[data-opt]').nth(0).check(), k2.locator('input[data-opt]').nth(3).check()]);
  await settle(k1, k2, k3); await sleep(2500);
  const sel = async (p) => p.evaluate(() => { const m = window.METT_APP.model; const q = m.question('5.'); return q.options.filter((o) => m.control('5.', o.name)).map((o) => o.letter); });
  const s1 = await sel(k1), s2 = await sel(k2), s3 = await sel(k3);
  ok('butir 5: tepat satu jawaban di tiap kelompok', s1.length === 1 && s2.length === 1 && s3.length === 1);
  ok('butir 5: ketiga kelompok sepakat (konsisten)', s1[0] === s2[0] && s2[0] === s3[0]);
  console.log('     (jawaban akhir butir 5:', s1[0] + ')');

  // --- sedang mengetik tidak tertimpa perubahan orang lain
  await go(k1, 's/1.'); await go(k2, 's/1.');
  await k1.click('textarea[data-key="1.!B15"]');
  await k1.keyboard.type('Ketikan K1 yang belum selesai');
  await k2.fill('textarea[data-key="1.!B19"]', 'Aksi dari K2'); await k2.press('textarea[data-key="1.!B19"]', 'Tab');
  await sleep(3500);
  ok('ketikan K1 tidak hilang saat data K2 masuk', (await k1.inputValue('textarea[data-key="1.!B15"]')) === 'Ketikan K1 yang belum selesai');
  ok('kolom aksi di layar K1 ikut terisi dari K2', (await k1.inputValue('textarea[data-key="1.!B19"]')) === 'Aksi dari K2');
  await k1.press('textarea[data-key="1.!B15"]', 'Tab');
  await settle(k1, k2);
  ok('ketikan K1 sampai ke K2', await val(k2, '1.!B15') === 'Ketikan K1 yang belum selesai');

  // --- putus koneksi lalu tersambung kembali
  await k3.evaluate((u) => fetch(u + '_offline?v=1').catch(() => {}), API);
  await go(k3, 's/3.'); await k3.locator('input[data-opt]').nth(1).check();
  await sleep(3000);
  const st = await k3.textContent('#saveState');
  ok('K3 menampilkan status gagal/offline + antrean', /Gagal|Offline|menunggu/.test(st));
  await (await b.newContext()).newPage().then((p) => p.goto(API + '_offline?v=0'));
  await sleep(6000);
  ok('antrean K3 terkirim setelah tersambung', await val(k1, 'Daftar Pertanyaan & Skor!E8') === 1);

  // --- muat ulang halaman: data tetap ada
  await k2.reload(); await k2.waitForFunction(() => document.querySelector('#saveState').textContent.includes('Tersinkron'), null, { timeout: 15000 });
  ok('K2 setelah muat ulang: total skor sama', await val(k2, 'Daftar Pertanyaan & Skor!E51') === await val(k1, 'Daftar Pertanyaan & Skor!E51'));

  // --- riwayat & kembalikan
  await go(k1, 'riwayat'); await k1.waitForSelector('[data-restore]', { timeout: 10000 });
  const txt = await k1.textContent('#main table');
  ok('riwayat mencatat nama pengisi', txt.includes('Budi – Kelompok 2') && txt.includes('Citra – Kelompok 3'));
  // hapus nama kawasan lalu kembalikan dari riwayat
  await go(k2, 's/Atribut%20Kawasan');
  await k2.fill('input[data-key="Atribut Kawasan!B5"]', ''); await k2.press('input[data-key="Atribut Kawasan!B5"]', 'Tab');
  await settle(k1, k2);
  ok('nama kawasan terhapus (simulasi salah hapus)', await val(k1, 'Atribut Kawasan!B5') === 0 || await val(k1, 'Atribut Kawasan!B5') === null);
  await go(k1, 'riwayat'); await k1.click('[data-act="reload-log"]'); await k1.waitForTimeout(800);
  const rows = await k1.$$eval('#main tbody tr', (trs) => trs.map((t) => t.textContent));
  const idx = rows.findIndex((r) => r.includes('Nama Kawasan') && r.includes('TN Uji Bersama'));
  await k1.locator('[data-restore]').nth(idx).click();
  await settle(k1, k2);
  ok('nama kawasan dikembalikan dari riwayat', await val(k2, 'Atribut Kawasan!B5') === 'TN Uji Bersama');

  ok('tidak ada error JavaScript', errs.length === 0);
  if (errs.length) console.log(errs);
  await b.close();
  console.log(failed ? `${failed} GAGAL` : 'SEMUA LULUS');
  process.exit(failed ? 1 : 0);
})().catch((e) => { console.error(e); process.exit(1); });
