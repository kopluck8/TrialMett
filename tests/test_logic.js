/*
 * Uji 2: logika makro VBA, validasi, dan ekspor.
 *   node tests/test_logic.js
 */
global.window = global;
require('../data/workbook.js');
const E = require('../js/engine.js');
const M = require('../js/model.js');
const X = require('../js/xlsx.js');
const WB = global.METT_WB;
let pass = 0, fail = 0;
const t = (name, cond) => { if (cond) pass++; else { fail++; console.log('GAGAL:', name); } };

const m = new M.Model(WB);
m.load(M.newAssessment());
const D = 'Daftar Pertanyaan & Skor';
const qs = m.schema.questions;

t('42 butir penilaian', qs.length === 42);
t('38 butir pilihan A-D', qs.filter((q) => q.kind === 'radio').length === 38);
t('4 butir tambahan (checkbox)', qs.filter((q) => q.kind === 'additive').length === 4);
t('skor maks total 126', m.engine.get(D, 'D51') === 126);

// skor pilihan A-D = 0..3 (sesuai VBA Worksheet_Deactivate)
for (const q of qs.filter((x) => x.kind === 'radio')) {
  t(q.sheet + ' punya 4 opsi', q.options.length === 4);
  q.options.forEach((o, i) => {
    m.selectOption(q, o.name);
    t(`${q.sheet} opsi ${o.letter} -> skor ${i}`, m.engine.get(D, 'E' + q.daftarRow) === i && o.letter === 'ABCD'[i]);
  });
}
// tidak relevan -> skor 0, maks 0, opsi dikosongkan
const q1 = m.question('1.');
m.selectOption(q1, q1.options[3].name);
m.setNA(q1, true);
t('NA: skor 0', m.engine.get(D, 'E6') === 0);
t('NA: maks 0', m.engine.get(D, 'D6') === 0);
t('NA: opsi dikosongkan (makro Q1ReleventCB1_Click)', m.selectedOption(q1) === null);
t('NA: pesan penjelasan wajib muncul bila B3 kosong', (() => { m.set('1.!B3', null); return m.messages(q1).some((x) => /Penjelasan harus diisi/.test(x.msg)); })());
m.setNA(q1, false);
t('lepas NA: maks kembali 3', m.engine.get(D, 'D6') === 3);

// checkbox tambahan: skor = jumlah centang
for (const q of qs.filter((x) => x.kind === 'additive')) {
  t(q.sheet + ' punya 3 pernyataan', q.boxes.length === 3 && q.boxes.every((b) => b.text));
  q.boxes.forEach((b) => m.toggleBox(q, b.name, false));
  t(q.sheet + ' skor 0', m.engine.get(D, 'E' + q.daftarRow) === 0);
  m.toggleBox(q, q.boxes[0].name, true); m.toggleBox(q, q.boxes[2].name, true);
  t(q.sheet + ' skor 2', m.engine.get(D, 'E' + q.daftarRow) === 2);
  m.toggleBox(q, q.boxes[1].name, true);
  t(q.sheet + ' skor 3', m.engine.get(D, 'E' + q.daftarRow) === 3);
}
t('7A ditandai catatan VBA (checkbox 29-31 tidak ada)', !!m.question('7A.').vbaNote);

// pesan isian wajib (teks MsgBox VBA)
const q2 = m.question('2.');
m.selectOption(q2, q2.options[1].name);
const msgs = m.messages(q2).map((x) => x.msg);
t('pesan bukti wajib', msgs.includes('Bukti dan justifikasi harus diisi.'));
t('pesan aksi wajib', msgs.includes('Aksi untuk meningkatkan efektifitas pengelolaan harus diisi'));
m.set('2.!B15', 'bukti'); m.set('2.!B19', 'aksi');
t('pesan hilang setelah diisi', m.messages(q2).length === 0);
t('aksi muncul di lembar Aksi (E6)', m.engine.get('Aksi meningkatkan Pengelolaan', 'E6') === 'aksi');
t('aksi muncul di Print_2 (E9)', m.engine.get('Print_2', 'E9') === 'aksi');

// multi-pilih dropdown (VBA Worksheet_Change)
const K = ' Penilaian Ancaman!D3';
m.set(K, null);
m.multiAdd(K, 'Harimau'); m.multiAdd(K, 'Gajah'); m.multiAdd(K, 'Harimau');
t('multi-pilih: "Harimau, Gajah" tanpa duplikat', m.inputText(K) === 'Harimau, Gajah');
t('Atribut B17 multi-pilih', m.isMultiSelect('Atribut Kawasan', 'B17'));

// validasi data
const v = (k, val) => M.checkValidation(m.validation(k), val);
t('F6 menolak 4 (whole 0-3)', !v(D + '!F6', 4).ok);
t('F6 menerima 2', v(D + '!F6', 2).ok);
t('B48 peringatan (warning) di luar 0-500', v('Atribut Kawasan!B48', 600).level === 'warning');
t('email valid', v('Atribut Kawasan!B62', 'nama@kemenhut.go.id').ok);
t('email tidak valid', !v('Atribut Kawasan!B62', 'nama@kemenhut').ok);
t('B20 angka saja', !v('Atribut Kawasan!B20', 'seribu').ok && v('Atribut Kawasan!B20', 1000).ok);

// pengurai ketikan ala Excel
t('"1.500.000" -> 1500000', M.parseTyped('1.500.000') === 1500000);
t('"17/08/2025" -> serial tanggal', M.parseTyped('17/08/2025') === E.dateToSerial(2025, 8, 17));
t('"0812345" tetap teks (nomor HP)', M.parseTyped('0812345') === '0812345');
t('serial 45886 = 17/08/2025', E.formatDate(45886, 'dd/mm/yyyy') === '17/08/2025');

// TEXT hari/bulan (lokal Indonesia)
m.set('Atribut Kawasan!B64', E.dateToSerial(2025, 8, 17));
t('Berita Acara: "Minggu, 17 Agustus 2025"', /Minggu , tanggal 17 , bulan Agustus , tahun 2025/.test(m.engine.get('Print BA', 'B14')));

// ekspor
const ex = m.exportTables();
t('ekspor 3 tabel', ex.length === 3);
t('PA attributes s.d. kolom CZ (rentang makro)', ex[0].header.length === E.colToNum('CZ'));
t('Threat data 436 kolom', ex[1].header.length === 436);
t('Assessment data 716 kolom', ex[2].header.length === 716);
const bytes = X.writeXlsx(ex.map((x) => ({ name: x.name, rows: [x.header, x.values.map((y) => (y instanceof E.XlError ? { error: y.error } : y))] })));
t('XLSX tertulis (ZIP)', bytes[0] === 0x50 && bytes[1] === 0x4b && bytes.length > 5000);

// simpan/buka file
const file = m.toFile();
const m2 = new M.Model(WB); m2.load(M.Model.parseFile(file));
t('buka file: skor total sama', m2.engine.get(D, 'E51') === m.engine.get(D, 'E51'));

console.log(`${pass} lulus, ${fail} gagal`);
process.exit(fail ? 1 : 0);
