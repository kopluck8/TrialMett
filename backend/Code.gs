/**
 * METT-4 KSDAE Web — Backend data bersama (Google Apps Script + Google Sheets)
 * =========================================================================
 * Menyimpan isian penilaian SATU kawasan sehingga bisa diisi bersama oleh banyak
 * orang/kelompok dari link aplikasi web (GitHub Pages).
 *
 * Cara pasang (ringkas — lengkapnya di docs/PANDUAN-ONLINE.md):
 *   1. Buat Google Sheet baru → Ekstensi → Apps Script → tempel seluruh file ini.
 *   2. Jalankan fungsi `setup` sekali (izinkan akses).
 *   3. Terapkan → Deployment baru → Aplikasi web:
 *        Jalankan sebagai: Saya   |   Yang memiliki akses: Siapa saja
 *   4. Salin URL aplikasi web (…/exec) ke js/config.js pada repository.
 *
 * Model data: satu baris per "sel" penilaian, memakai alamat sel Excel yang sama
 * dengan file METT-4 asli (mis. "1.!B15", "Atribut Kawasan!B5"). Setiap
 * perubahan menaikkan nomor revisi dan dicatat di sheet "log" (siapa, kapan,
 * nilai lama → baru) sehingga isian yang terhapus/tertimpa dapat dikembalikan.
 *
 * API (dipanggil oleh js/sync.js):
 *   GET  ?action=info                    → {ok, rev, rows}
 *   GET  ?action=pull&since=<rev>        → {ok, rev, changes:[{k,v,r,t,by}]}
 *   GET  ?action=log&limit=<n>           → {ok, log:[{t,by,k,old,v,r}]}
 *   POST {action:'push', by, client, changes:[{k,v}]}  (Content-Type text/plain)
 *                                        → {ok, rev, applied}
 */

var SHEET_DATA = 'data';
var SHEET_LOG = 'log';
var DATA_HEADER = ['key', 'value', 'rev', 'updatedAt', 'by'];
var LOG_HEADER = ['time', 'by', 'key', 'oldValue', 'newValue', 'rev', 'client'];
var MAX_BATCH = 5000;
var MAX_VALUE_LEN = 45000; // batas aman isi sel Google Sheets (50.000 karakter)

/** Jalankan sekali dari editor Apps Script untuk menyiapkan sheet. */
function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var data = ss.getSheetByName(SHEET_DATA) || ss.insertSheet(SHEET_DATA);
  var log = ss.getSheetByName(SHEET_LOG) || ss.insertSheet(SHEET_LOG);
  data.getRange(1, 1, 1, DATA_HEADER.length).setValues([DATA_HEADER]).setFontWeight('bold');
  log.getRange(1, 1, 1, LOG_HEADER.length).setValues([LOG_HEADER]).setFontWeight('bold');
  // kolom nilai disimpan sebagai teks polos agar tidak diubah otomatis oleh Sheets
  data.getRange('A:B').setNumberFormat('@');
  log.getRange('C:E').setNumberFormat('@');
  data.setFrozenRows(1);
  log.setFrozenRows(1);
  var props = PropertiesService.getScriptProperties();
  if (!props.getProperty('rev')) props.setProperty('rev', '0');
  return 'Siap. Revisi saat ini: ' + props.getProperty('rev');
}

function doGet(e) {
  var p = (e && e.parameter) || {};
  try {
    if (p.action === 'pull') return json_(pull_(Number(p.since) || 0));
    if (p.action === 'log') return json_(readLog_(Math.min(Number(p.limit) || 300, 2000)));
    return json_({ ok: true, app: 'METT4-KSDAE-sync', rev: getRev_(), rows: Math.max(0, sheet_(SHEET_DATA).getLastRow() - 1) });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}

function doPost(e) {
  try {
    var body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (body.action === 'push') return json_(push_(body));
    return json_({ ok: false, error: 'Aksi tidak dikenal' });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}

// ---------------------------------------------------------------- inti

function pull_(since) {
  var sh = sheet_(SHEET_DATA);
  var n = sh.getLastRow() - 1;
  var changes = [];
  if (n > 0) {
    var rows = sh.getRange(2, 1, n, 5).getValues();
    for (var i = 0; i < rows.length; i++) {
      var r = Number(rows[i][2]) || 0;
      if (r > since && rows[i][0] !== '') {
        changes.push({ k: String(rows[i][0]), v: decode_(rows[i][1]), r: r, t: iso_(rows[i][3]), by: String(rows[i][4] || '') });
      }
    }
  }
  return { ok: true, rev: getRev_(), changes: changes };
}

function push_(body) {
  var changes = body.changes || [];
  if (!changes.length) return { ok: true, rev: getRev_(), applied: 0 };
  if (changes.length > MAX_BATCH) throw new Error('Terlalu banyak perubahan sekaligus');
  var by = String(body.by || '').slice(0, 120);
  var client = String(body.client || '').slice(0, 60);
  var lock = LockService.getScriptLock();
  lock.waitLock(25000); // antrekan penulisan bersamaan dari banyak pengguna
  try {
    var sh = sheet_(SHEET_DATA);
    var n = sh.getLastRow() - 1;
    var rows = n > 0 ? sh.getRange(2, 1, n, 5).getValues() : [];
    var index = {};
    for (var i = 0; i < rows.length; i++) index[String(rows[i][0])] = i;
    var props = PropertiesService.getScriptProperties();
    var rev = Number(props.getProperty('rev')) || 0;
    var now = new Date();
    var logRows = [];
    var applied = 0;
    var touched = {};
    var firstNew = rows.length;
    for (var j = 0; j < changes.length; j++) {
      var k = String(changes[j].k || '');
      if (!k || k.length > 200) continue;
      var enc = encode_(changes[j].v);
      var at = index[k];
      var old = at === undefined ? '' : String(rows[at][1]);
      if (old === enc) continue; // tidak berubah
      rev++;
      applied++;
      if (at === undefined) { at = rows.length; index[k] = at; rows.push([k, enc, rev, now, by]); }
      else rows[at] = [k, enc, rev, now, by];
      touched[at] = true;
      logRows.push([now, by, k, old, enc, rev, client]);
    }
    if (applied) {
      var idx = Object.keys(touched).map(Number).filter(function (x) { return x < firstNew; });
      if (idx.length > 25) {
        sh.getRange(2, 1, rows.length, 5).setValues(rows); // banyak perubahan: tulis sekaligus
      } else {
        for (var t = 0; t < idx.length; t++) sh.getRange(idx[t] + 2, 1, 1, 5).setValues([rows[idx[t]]]);
        if (rows.length > firstNew) sh.getRange(firstNew + 2, 1, rows.length - firstNew, 5).setValues(rows.slice(firstNew));
      }
      var lg = sheet_(SHEET_LOG);
      lg.getRange(lg.getLastRow() + 1, 1, logRows.length, LOG_HEADER.length).setValues(logRows);
      props.setProperty('rev', String(rev));
    }
    return { ok: true, rev: rev, applied: applied };
  } finally {
    lock.releaseLock();
  }
}

function readLog_(limit) {
  var lg = sheet_(SHEET_LOG);
  var last = lg.getLastRow();
  var out = [];
  if (last > 1) {
    var start = Math.max(2, last - limit + 1);
    var rows = lg.getRange(start, 1, last - start + 1, LOG_HEADER.length).getValues();
    for (var i = rows.length - 1; i >= 0; i--) {
      out.push({ t: iso_(rows[i][0]), by: String(rows[i][1] || ''), k: String(rows[i][2]), old: decode_(rows[i][3]), v: decode_(rows[i][4]), r: Number(rows[i][5]) || 0 });
    }
  }
  return { ok: true, log: out };
}

// ---------------------------------------------------------------- util

function getRev_() { return Number(PropertiesService.getScriptProperties().getProperty('rev')) || 0; }

function sheet_(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (!sh) { setup(); sh = ss.getSheetByName(name); }
  return sh;
}

/** Nilai disimpan sebagai JSON (mempertahankan tipe: teks/angka/benar-salah/kosong). */
function encode_(v) {
  if (v === undefined) v = null;
  var s = JSON.stringify(v);
  if (s.length > MAX_VALUE_LEN) throw new Error('Isian terlalu panjang (maks. ' + MAX_VALUE_LEN + ' karakter)');
  return s;
}
function decode_(s) {
  if (s === '' || s === null || s === undefined) return null;
  if (typeof s !== 'string') return s; // jaga-jaga bila sel bukan teks
  try { return JSON.parse(s); } catch (e) { return s; }
}
function iso_(d) { return d instanceof Date ? d.toISOString() : String(d || ''); }

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
