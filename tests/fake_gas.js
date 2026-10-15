/*
 * Tiruan minimal layanan Google Apps Script (SpreadsheetApp, PropertiesService,
 * LockService, ContentService) untuk menjalankan backend/Code.gs ASLI di Node.js.
 * Dipakai oleh tests/test_sync_backend.js dan tests/sync_server.js.
 */
const fs = require('fs');
const path = require('path');
const vm = require('vm');

function makeSheet(name) {
  const cells = []; // cells[r][c], 1-based diisi lewat helper
  const sheet = {
    name,
    _cells: cells,
    getLastRow() { for (let r = cells.length - 1; r >= 1; r--) if (cells[r] && cells[r].some((v) => v !== '' && v !== undefined && v !== null)) return r; return 0; },
    getRange(a, b, c, d) {
      if (typeof a === 'string') { // 'A:B' → hanya untuk setNumberFormat
        return { setNumberFormat() { return this; } };
      }
      const row = a, col = b, nr = c || 1, nc = d || 1;
      return {
        setValues(vals) {
          if (vals.length !== nr || vals.some((x) => x.length !== nc)) throw new Error(`Ukuran data tidak cocok ${vals.length}x${vals[0] && vals[0].length} vs ${nr}x${nc}`);
          for (let i = 0; i < nr; i++) { cells[row + i] = cells[row + i] || []; for (let j = 0; j < nc; j++) cells[row + i][col + j] = vals[i][j]; }
          return this;
        },
        getValues() {
          const out = [];
          for (let i = 0; i < nr; i++) { const r = []; for (let j = 0; j < nc; j++) { const v = cells[row + i] && cells[row + i][col + j]; r.push(v === undefined || v === null ? '' : v); } out.push(r); }
          return out;
        },
        setFontWeight() { return this; },
        setNumberFormat() { return this; },
      };
    },
    setFrozenRows() {},
  };
  return sheet;
}

function createGas() {
  const sheets = {};
  const props = {};
  let locked = false;
  const ss = {
    getSheetByName: (n) => sheets[n] || null,
    insertSheet: (n) => (sheets[n] = makeSheet(n)),
  };
  const ctx = {
    SpreadsheetApp: { getActiveSpreadsheet: () => ss },
    PropertiesService: { getScriptProperties: () => ({ getProperty: (k) => (k in props ? props[k] : null), setProperty: (k, v) => { props[k] = String(v); } }) },
    LockService: { getScriptLock: () => ({ waitLock() { if (locked) throw new Error('lock busy (tes sinkron)'); locked = true; }, releaseLock() { locked = false; } }) },
    ContentService: {
      MimeType: { JSON: 'application/json' },
      createTextOutput: (s) => ({ _text: s, setMimeType() { return this; }, getContent() { return this._text; } }),
    },
    console, JSON, Date, Math, Number, String, Object, Error,
  };
  vm.createContext(ctx);
  vm.runInContext(fs.readFileSync(path.join(__dirname, '..', 'backend', 'Code.gs'), 'utf8'), ctx, { filename: 'Code.gs' });
  return {
    ctx, sheets, props,
    get: (params) => JSON.parse(ctx.doGet({ parameter: params || {} }).getContent()),
    post: (body) => JSON.parse(ctx.doPost({ postData: { contents: typeof body === 'string' ? body : JSON.stringify(body) } }).getContent()),
  };
}

module.exports = { createGas };
