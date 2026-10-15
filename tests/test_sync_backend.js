/*
 * Uji 4: backend/Code.gs (dijalankan apa adanya di atas tiruan layanan Google).
 *   node tests/test_sync_backend.js
 */
const { createGas } = require('./fake_gas.js');
let pass = 0, fail = 0;
const t = (n, c) => { if (c) pass++; else { fail++; console.log('GAGAL:', n); } };

const g = createGas();
t('setup', /Siap/.test(g.ctx.setup()));
t('info awal rev 0', g.get({ action: 'info' }).rev === 0);

let r = g.post({ action: 'push', by: 'Kelompok 1 - Ani', client: 'c1', changes: [
  { k: 'Atribut Kawasan!B5', v: 'TN Baluran' }, { k: '1.!#Option Button 28', v: true }, { k: 'Atribut Kawasan!B20', v: 25000 },
  { k: 'Question_data_dump!H2', v: false }, { k: '1.!B15', v: null },
] });
t('push diterima', r.ok && r.applied === 5 && r.rev === 5);
let p = g.get({ action: 'pull', since: '0' });
const m = Object.fromEntries(p.changes.map((c) => [c.k, c.v]));
t('tipe teks terjaga', m['Atribut Kawasan!B5'] === 'TN Baluran');
t('tipe angka terjaga', m['Atribut Kawasan!B20'] === 25000);
t('tipe boolean terjaga', m['1.!#Option Button 28'] === true && m['Question_data_dump!H2'] === false);
t('null terjaga', m['1.!B15'] === null);
t('pencatat "by"', p.changes[0].by === 'Kelompok 1 - Ani');

r = g.post({ action: 'push', by: 'Kelompok 2', changes: [{ k: 'Atribut Kawasan!B5', v: 'TN Baluran' }] });
t('nilai sama tidak menaikkan revisi', r.applied === 0 && r.rev === 5);
r = g.post({ action: 'push', by: 'Kelompok 2 - Budi', changes: [{ k: 'Atribut Kawasan!B5', v: 'Taman Nasional Baluran' }, { k: 'Atribut Kawasan!B19', v: 'Balai TN Baluran' }] });
t('perubahan kedua', r.applied === 2 && r.rev === 7);
p = g.get({ action: 'pull', since: '5' });
t('pull sejak rev 5 hanya 2 perubahan', p.changes.length === 2 && p.rev === 7);
t('tidak ada baris duplikat', g.sheets.data.getLastRow() === 1 + 6);

const lg = g.get({ action: 'log', limit: '10' });
t('log terbaru di atas', lg.log[0].r === 7);
const ch = lg.log.find((x) => x.k === 'Atribut Kawasan!B5' && x.r === 6);
t('log mencatat nilai lama & baru', ch && ch.old === 'TN Baluran' && ch.v === 'Taman Nasional Baluran' && ch.by === 'Kelompok 2 - Budi');

// banyak perubahan sekaligus (impor file) → jalur tulis massal
const many = [];
for (let i = 0; i < 1500; i++) many.push({ k: ' Penilaian Ancaman!C' + i, v: 'teks ' + i });
r = g.post({ action: 'push', by: 'Admin', changes: many });
t('impor 1500 sel', r.applied === 1500);
many.length = 0;
for (let i = 0; i < 40; i++) many.push({ k: ' Penilaian Ancaman!C' + i, v: 'ubah ' + i });
r = g.post({ action: 'push', by: 'Admin', changes: many });
t('ubah 40 sel lama (tulis massal)', r.applied === 40);
p = g.get({ action: 'pull', since: '0' });
const mm = Object.fromEntries(p.changes.map((c) => [c.k, c.v]));
t('data konsisten setelah tulis massal', mm[' Penilaian Ancaman!C3'] === 'ubah 3' && mm[' Penilaian Ancaman!C100'] === 'teks 100' && mm['Atribut Kawasan!B5'] === 'Taman Nasional Baluran');
t('teks dengan tanda kutip & baris baru', (() => { g.post({ action: 'push', changes: [{ k: '2.!B15', v: 'Bukti "SK 123"\nbaris 2' }] }); return Object.fromEntries(g.get({ action: 'pull', since: '0' }).changes.map((c) => [c.k, c.v]))['2.!B15'] === 'Bukti "SK 123"\nbaris 2'; })());
t('aksi tidak dikenal ditolak', g.post({ action: 'hapus' }).ok === false);
t('JSON rusak ditolak dengan pesan', g.post('{bukan json').ok === false);

console.log(`${pass} lulus, ${fail} gagal`);
process.exit(fail ? 1 : 0);
