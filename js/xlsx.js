/*
 * METT-4 Web — Penulis & pembaca XLSX minimal (tanpa pustaka eksternal)
 * - writeXlsx(sheets): membuat file .xlsx (ZIP tanpa kompresi) berisi beberapa sheet
 * - readXlsx(arrayBuffer): membaca .xlsx (mendukung ZIP terkompresi via DecompressionStream)
 */
(function (root) {
  'use strict';

  // ------------------------------------------------------------- CRC32 & ZIP
  const CRC_TABLE = (() => {
    const t = new Uint32Array(256);
    for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1; t[n] = c >>> 0; }
    return t;
  })();
  function crc32(bytes) {
    let c = 0xFFFFFFFF;
    for (let i = 0; i < bytes.length; i++) c = CRC_TABLE[(c ^ bytes[i]) & 0xFF] ^ (c >>> 8);
    return (c ^ 0xFFFFFFFF) >>> 0;
  }
  const enc = (s) => new TextEncoder().encode(s);

  function zipStore(files) {
    // files: [{name, data: Uint8Array}]
    const chunks = [], central = [];
    let offset = 0;
    const now = new Date();
    const dosTime = (now.getHours() << 11) | (now.getMinutes() << 5) | (now.getSeconds() >> 1);
    const dosDate = ((now.getFullYear() - 1980) << 9) | ((now.getMonth() + 1) << 5) | now.getDate();
    for (const f of files) {
      const name = enc(f.name), data = f.data, crc = crc32(data);
      const lh = new DataView(new ArrayBuffer(30));
      lh.setUint32(0, 0x04034b50, true); lh.setUint16(4, 20, true); lh.setUint16(6, 0x0800, true); lh.setUint16(8, 0, true);
      lh.setUint16(10, dosTime, true); lh.setUint16(12, dosDate, true); lh.setUint32(14, crc, true);
      lh.setUint32(18, data.length, true); lh.setUint32(22, data.length, true); lh.setUint16(26, name.length, true); lh.setUint16(28, 0, true);
      chunks.push(new Uint8Array(lh.buffer), name, data);
      const ch = new DataView(new ArrayBuffer(46));
      ch.setUint32(0, 0x02014b50, true); ch.setUint16(4, 20, true); ch.setUint16(6, 20, true); ch.setUint16(8, 0x0800, true); ch.setUint16(10, 0, true);
      ch.setUint16(12, dosTime, true); ch.setUint16(14, dosDate, true); ch.setUint32(16, crc, true);
      ch.setUint32(20, data.length, true); ch.setUint32(24, data.length, true); ch.setUint16(28, name.length, true);
      ch.setUint32(42, offset, true);
      central.push(new Uint8Array(ch.buffer), name);
      offset += 30 + name.length + data.length;
    }
    const cdSize = central.reduce((a, b) => a + b.length, 0);
    const end = new DataView(new ArrayBuffer(22));
    end.setUint32(0, 0x06054b50, true); end.setUint16(8, files.length, true); end.setUint16(10, files.length, true);
    end.setUint32(12, cdSize, true); end.setUint32(16, offset, true);
    const parts = [...chunks, ...central, new Uint8Array(end.buffer)];
    const total = parts.reduce((a, b) => a + b.length, 0);
    const out = new Uint8Array(total);
    let p = 0; for (const c of parts) { out.set(c, p); p += c.length; }
    return out;
  }

  async function unzip(buf) {
    const u8 = new Uint8Array(buf), dv = new DataView(buf);
    let eocd = -1;
    for (let i = u8.length - 22; i >= Math.max(0, u8.length - 70000); i--) if (dv.getUint32(i, true) === 0x06054b50) { eocd = i; break; }
    if (eocd < 0) throw new Error('Bukan berkas ZIP/XLSX yang valid');
    const count = dv.getUint16(eocd + 10, true);
    let p = dv.getUint32(eocd + 16, true);
    const files = {};
    for (let n = 0; n < count; n++) {
      const method = dv.getUint16(p + 10, true), csize = dv.getUint32(p + 20, true);
      const nlen = dv.getUint16(p + 28, true), xlen = dv.getUint16(p + 30, true), clen = dv.getUint16(p + 32, true);
      const lho = dv.getUint32(p + 42, true);
      const name = new TextDecoder().decode(u8.slice(p + 46, p + 46 + nlen));
      const lnlen = dv.getUint16(lho + 26, true), lxlen = dv.getUint16(lho + 28, true);
      const start = lho + 30 + lnlen + lxlen;
      const raw = u8.slice(start, start + csize);
      files[name] = { method, raw };
      p += 46 + nlen + xlen + clen;
    }
    const get = async (name) => {
      const f = files[name];
      if (!f) return null;
      if (f.method === 0) return new TextDecoder().decode(f.raw);
      if (f.method === 8) {
        if (typeof DecompressionStream === 'undefined') throw new Error('Browser tidak mendukung dekompresi ZIP');
        const ds = new DecompressionStream('deflate-raw');
        const stream = new Blob([f.raw]).stream().pipeThrough(ds);
        return await new Response(stream).text();
      }
      throw new Error('Metode kompresi ZIP tidak didukung: ' + f.method);
    };
    return { names: Object.keys(files), get };
  }

  // ------------------------------------------------------------- XML util
  const xmlEsc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')
    // karakter kontrol tidak valid di XML
    .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '');
  const colName = (n) => { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };

  /**
   * sheets: [{name, rows: [[cell, ...], ...]}], cell = string | number | boolean | null | {error:'#REF!'}
   */
  function writeXlsx(sheets) {
    const files = [];
    const sheetXml = (rows) => {
      let x = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>';
      rows.forEach((row, ri) => {
        x += `<row r="${ri + 1}">`;
        row.forEach((v, ci) => {
          if (v === null || v === undefined || v === '') return;
          const ref = colName(ci + 1) + (ri + 1);
          const style = ri === 0 ? ' s="1"' : '';
          if (typeof v === 'number' && isFinite(v)) x += `<c r="${ref}"${style}><v>${v}</v></c>`;
          else if (typeof v === 'boolean') x += `<c r="${ref}" t="b"${style}><v>${v ? 1 : 0}</v></c>`;
          else if (v && typeof v === 'object' && v.error) x += `<c r="${ref}" t="e"${style}><v>${xmlEsc(v.error)}</v></c>`;
          else x += `<c r="${ref}" t="inlineStr"${style}><is><t xml:space="preserve">${xmlEsc(v)}</t></is></c>`;
        });
        x += '</row>';
      });
      return x + '</sheetData></worksheet>';
    };
    const safeName = (n) => n.replace(/[\\/?*[\]:]/g, '_').slice(0, 31);
    files.push({ name: '[Content_Types].xml', data: enc('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
      sheets.map((_, i) => `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join('') + '</Types>') });
    files.push({ name: '_rels/.rels', data: enc('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>') });
    files.push({ name: 'xl/workbook.xml', data: enc('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>' +
      sheets.map((s, i) => `<sheet name="${xmlEsc(safeName(s.name))}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join('') + '</sheets></workbook>') });
    files.push({ name: 'xl/_rels/workbook.xml.rels', data: enc('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      sheets.map((_, i) => `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`).join('') +
      `<Relationship Id="rId${sheets.length + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`) });
    files.push({ name: 'xl/styles.xml', data: enc('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/></cellXfs></styleSheet>') });
    sheets.forEach((s, i) => files.push({ name: `xl/worksheets/sheet${i + 1}.xml`, data: enc(sheetXml(s.rows)) }));
    return zipStore(files);
  }

  function decodeXml(s) {
    return s.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'")
      .replace(/&#(\d+);/g, (_, n) => String.fromCharCode(+n)).replace(/&#x([0-9a-f]+);/gi, (_, n) => String.fromCharCode(parseInt(n, 16))).replace(/&amp;/g, '&');
  }
  async function readXlsx(buf) {
    const z = await unzip(buf);
    const wb = await z.get('xl/workbook.xml');
    const rels = await z.get('xl/_rels/workbook.xml.rels');
    const ssXml = await z.get('xl/sharedStrings.xml');
    const shared = [];
    if (ssXml) for (const m of ssXml.matchAll(/<si>([\s\S]*?)<\/si>/g)) shared.push(decodeXml([...m[1].matchAll(/<t[^>]*>([\s\S]*?)<\/t>/g)].map((x) => x[1]).join('')));
    const relMap = {};
    for (const m of rels.matchAll(/<Relationship\b[^>]*>/g)) {
      const id = /Id="([^"]+)"/.exec(m[0])[1], t = /Target="([^"]+)"/.exec(m[0])[1];
      relMap[id] = t.replace(/^\/?xl\//, '').replace(/^\//, '');
    }
    const out = [];
    for (const m of wb.matchAll(/<sheet\b[^>]*>/g)) {
      const name = decodeXml(/name="([^"]*)"/.exec(m[0])[1]);
      const rid = /r:id="([^"]+)"/.exec(m[0])[1];
      const xml = await z.get('xl/' + relMap[rid]);
      const rows = [];
      for (const c of xml.matchAll(/<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g)) {
        const attrs = c[1], body = c[2] || '';
        const ref = /r="([A-Z]+)(\d+)"/.exec(attrs);
        if (!ref) continue;
        let col = 0; for (const ch of ref[1]) col = col * 26 + ch.charCodeAt(0) - 64;
        const row = +ref[2];
        const t = (/t="([^"]+)"/.exec(attrs) || [])[1];
        const vm = /<v>([\s\S]*?)<\/v>/.exec(body);
        let v = null;
        if (t === 's') v = shared[+vm[1]];
        else if (t === 'inlineStr') v = decodeXml([...body.matchAll(/<t[^>]*>([\s\S]*?)<\/t>/g)].map((x) => x[1]).join(''));
        else if (t === 'b') v = vm ? vm[1] === '1' : null;
        else if (t === 'e') v = vm ? { error: decodeXml(vm[1]) } : null;
        else if (t === 'str') v = vm ? decodeXml(vm[1]) : null;
        else v = vm ? parseFloat(vm[1]) : null;
        (rows[row - 1] = rows[row - 1] || [])[col - 1] = v;
      }
      out.push({ name, rows });
    }
    return out;
  }

  function toCsv(rows) {
    const q = (v) => {
      if (v === null || v === undefined) return '';
      if (typeof v === 'object' && v.error) v = v.error;
      const s = typeof v === 'boolean' ? (v ? 'TRUE' : 'FALSE') : String(v);
      return /[",;\n\r]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
    };
    return '﻿' + rows.map((r) => r.map(q).join(',')).join('\r\n');
  }

  const api = { writeXlsx, readXlsx, toCsv, crc32, unzip };
  root.MettXlsx = api;
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
})(typeof window !== 'undefined' ? window : globalThis);
