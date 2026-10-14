/*
 * METT-4 Web — Antarmuka
 * Semua teks yang tampil (pertanyaan, opsi, panduan, label) dibaca langsung dari
 * sel Excel asli (data/workbook.js). Semua angka hasil dihitung oleh mesin rumus
 * dari rumus Excel asli. File ini hanya mengatur tata letak & interaksi.
 */
(function () {
  'use strict';
  const WB = window.METT_WB, E = window.MettEngine, M = window.MettModel, C = window.MettCharts, X = window.MettXlsx;
  const model = new M.Model(WB);
  const DAFTAR = M.DAFTAR;
  const ATTR = 'Atribut Kawasan', THREAT = ' Penilaian Ancaman';
  const $ = (s, el) => (el || document).querySelector(s);
  const $$ = (s, el) => Array.from((el || document).querySelectorAll(s));
  const esc = (s) => String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  const sheetOf = (name) => WB.sheets.find((s) => s.name === name);
  const raw = (sheet, addr) => { const s = sheetOf(sheet); const c = s && s.cells[addr]; return c && c.v != null ? String(c.v) : ''; };
  const disp = (sheet, addr) => model.display(sheet + '!' + addr);
  const key = (sheet, addr) => sheet + '!' + addr;
  const clean = (s) => String(s || '').replace(/[ \t]{2,}/g, ' ').trim();

  // ================================================================ penyimpanan browser
  const LS = {
    get(k) { try { return localStorage.getItem(k); } catch (e) { return null; } },
    set(k, v) { try { localStorage.setItem(k, v); return true; } catch (e) { return false; } },
    del(k) { try { localStorage.removeItem(k); } catch (e) { /* abaikan */ } },
  };
  const LS_CUR = 'mett4.current', LS_LIST = 'mett4.list', LS_A = 'mett4.a.';
  function listSaved() { try { return JSON.parse(LS.get(LS_LIST) || '[]'); } catch (e) { return []; } }
  let saveTimer = null;
  function persistSoon() {
    $('#saveState').textContent = 'menyimpan…';
    clearTimeout(saveTimer);
    saveTimer = setTimeout(persistNow, 400);
  }
  function persistNow() {
    const s = model.state;
    const ok = LS.set(LS_A + s.id, model.toFile());
    const sum = model.summary();
    const list = listSaved().filter((x) => x.id !== s.id);
    list.unshift({ id: s.id, name: sum.name && sum.name !== '0' ? sum.name : '(tanpa nama kawasan)', year: E.toText(sum.year), updated: s.updated, total: sum.total, max: sum.max });
    LS.set(LS_LIST, JSON.stringify(list));
    LS.set(LS_CUR, s.id);
    $('#saveState').textContent = ok ? 'tersimpan di browser ' + new Date().toLocaleTimeString('id-ID', { hour: '2-digit', minute: '2-digit' }) : 'gagal menyimpan di browser — simpan ke file!';
    updateHeader();
  }
  function loadSaved(id) {
    const txt = LS.get(LS_A + id);
    if (!txt) return false;
    try { model.load(M.Model.parseFile(txt)); LS.set(LS_CUR, id); return true; } catch (e) { console.error(e); return false; }
  }

  function toast(msg, ms) {
    const t = $('#toast');
    t.textContent = msg; t.classList.add('show');
    clearTimeout(t._h); t._h = setTimeout(() => t.classList.remove('show'), ms || 2600);
  }
  function download(name, data, type) {
    const blob = data instanceof Blob ? data : new Blob([data], { type: type || 'application/octet-stream' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob); a.download = name;
    document.body.appendChild(a); a.click();
    setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 500);
  }

  // ================================================================ daftar halaman
  const DETAIL = { 'Detail penilaian nilai utama': 'detail-nilai', 'Detail Penilaian Spesies': 'detail-spesies', 'Detail Penilaian Habitat': 'detail-habitat' };
  const TYPE_BY_SHEET = {
    'Tentang': 'tentang', 'Pengenalan METT 4': 'pengenalan', 'Dashboard': 'dashboard', [ATTR]: 'atribut', [THREAT]: 'ancaman',
    [DAFTAR]: 'daftar', 'Aksi meningkatkan Pengelolaan': 'aksi', 'Print BA': 'ba', 'Print_2': 'rencana', 'Bantuan mengidentifikasi nilai': 'bantuan',
  };
  const ALIAS = { 'Print BA': 'Print BA — Berita Acara', 'Print_2': 'Print_2 — Rencana Aksi', [THREAT]: 'Penilaian Ancaman' };
  const pages = [];
  for (const s of WB.sheets) {
    if (s.state !== 'visible') continue; // sheet tersembunyi di Excel tetap tersembunyi
    let type = TYPE_BY_SHEET[s.name] || DETAIL[s.name] || (model.question(s.name) ? 'question' : null);
    if (!type) continue;
    pages.push({ id: 's/' + encodeURIComponent(s.name), sheet: s.name, type, title: ALIAS[s.name] || s.name.trim() });
  }
  const appPages = [
    { id: 'data', type: 'data', title: 'Kelola Data & Ekspor' },
    { id: 'rekap', type: 'rekap', title: 'Rekap Multi-Kawasan' },
    { id: 'info', type: 'info', title: 'Tentang Aplikasi Ini' },
  ];
  const seqPages = pages.filter((p) => p.type === 'question' || p.type.startsWith('detail-'));
  const pageBySheet = (name) => pages.find((p) => p.sheet === name);
  const href = (p) => '#/' + p.id;
  const hrefSheet = (name) => { const p = pageBySheet(name); return p ? href(p) : '#/s/Tentang'; };

  // tautan Excel (hyperlink / tombol) -> rute. Nama sheet lama pada file asli dipetakan ke sheet yang ada.
  const LEGACY = { 'METT 4 questions+scores': DAFTAR, 'PENGENALAN METT 4': 'Pengenalan METT 4', 'METT 4 introduction': 'Pengenalan METT 4', 'Protected area attributes': ATTR };
  function linkToHref(target) {
    if (!target) return null;
    const m = /^#?'?(.*?)'?!/.exec(target);
    if (!m) return null;
    let name = m[1].replace(/''/g, "'");
    name = LEGACY[name] || name;
    const s = WB.sheets.find((x) => x.name === name) || WB.sheets.find((x) => x.name.toLowerCase() === name.toLowerCase());
    return s ? hrefSheet(s.name) : null;
  }

  // ================================================================ komponen isian
  function fieldText(k, opts) {
    opts = opts || {};
    const v = model.inputText(k);
    const ph = opts.placeholder ? ` placeholder="${esc(opts.placeholder)}"` : '';
    const lbl = opts.label ? ` aria-label="${esc(opts.label)}"` : '';
    if (opts.multiline) return `<textarea data-key="${esc(k)}" data-kind="text" class="${opts.cls || ''}"${ph}${lbl} rows="${opts.rows || 3}">${esc(v)}</textarea>`;
    const t = opts.kind || 'text';
    const im = t === 'number' || t === 'int' ? ' inputmode="decimal"' : '';
    return `<input type="text" data-key="${esc(k)}" data-kind="${t}" value="${esc(v)}"${im}${ph}${lbl}>`;
  }
  function fieldDate(k, label) {
    const v = model.get(k);
    let iso = '';
    if (typeof v === 'number' && v > 0) { const p = E.serialToParts(v); iso = p.y + '-' + String(p.m).padStart(2, '0') + '-' + String(p.d).padStart(2, '0'); }
    return `<input type="date" data-key="${esc(k)}" data-kind="date" value="${iso}" aria-label="${esc(label || '')}">`;
  }
  function fieldSelect(k, options, label) {
    const cur = model.inputText(k);
    let html = `<select data-key="${esc(k)}" data-kind="select" aria-label="${esc(label || '')}"><option value=""></option>`;
    let found = false;
    for (const o of options) { const sel = o === cur; if (sel) found = true; html += `<option value="${esc(o)}"${sel ? ' selected' : ''}>${esc(clean(o))}</option>`; }
    if (cur && !found) html += `<option value="${esc(cur)}" selected>${esc(cur)}</option>`;
    return html + '</select>';
  }
  function fieldMulti(k, options, label) {
    const cur = model.inputText(k);
    const parts = cur ? cur.split(', ') : [];
    const chips = parts.map((p, i) => `<span class="chip">${esc(clean(p))}<button type="button" data-chip-del="${i}" aria-label="Hapus ${esc(clean(p))}">×</button></span>`).join('');
    const opts = options.map((o) => `<option value="${esc(o)}">${esc(clean(o))}</option>`).join('');
    return `<div class="multi-wrap"><div class="multi" data-key="${esc(k)}" data-kind="multi"><div class="chips">${chips || '<span class="muted small">belum ada pilihan</span>'}</div>` +
      `<select data-multi-add aria-label="${esc(label || 'Tambah pilihan')}"><option value="">+ tambah pilihan…</option>${opts}</select></div></div>`;
  }
  function fieldBool(k, labelHtml) {
    const on = model.get(k) === true;
    return `<label class="check"><input type="checkbox" data-key="${esc(k)}" data-kind="bool"${on ? ' checked' : ''}><span>${labelHtml || ''}</span></label>`;
  }
  /** Pilih komponen berdasarkan Data Validation & format sel pada Excel */
  function autoField(sheet, addr, opts) {
    opts = opts || {};
    const k = key(sheet, addr);
    const s = sheetOf(sheet);
    const cell = s.cells[addr];
    if (cell && cell.f) return `<div class="computed" data-f="${esc(k)}">${esc(model.display(k))}</div>`;
    if (model.isMultiSelect(sheet, addr)) {
      const o = model.options(k);
      if (o) return fieldMulti(k, o, opts.label);
    }
    const v = model.validation(k);
    if (v && v.type === 'list') return fieldSelect(k, v._options, opts.label);
    if (cell && cell.nf && /dd\/mm\/yyyy/i.test(cell.nf)) return fieldDate(k, opts.label);
    if (v && v.type === 'whole') return fieldText(k, Object.assign({ kind: 'int' }, opts));
    if (v && v.type === 'custom' && /ISNUMBER/i.test(v.formula1 || '')) return fieldText(k, Object.assign({ kind: 'number' }, opts));
    if (v && v.type === 'custom' && /FIND\("@"/i.test(v.formula1 || '')) return fieldText(k, Object.assign({ kind: 'email' }, opts));
    return fieldText(k, opts);
  }
  const computed = (sheet, addr, cls) => `<span class="${cls || ''}" data-f="${esc(key(sheet, addr))}">${esc(disp(sheet, addr))}</span>`;

  function showFieldMsg(el, msg, level) {
    let m = el.parentElement.querySelector(':scope > .field-err, :scope > .field-warn');
    if (!msg) { if (m) m.remove(); el.classList.remove('invalid'); return; }
    if (!m) { m = document.createElement('div'); el.insertAdjacentElement('afterend', m); }
    m.className = level === 'stop' ? 'field-err' : 'field-warn';
    m.textContent = msg;
    el.classList.toggle('invalid', level === 'stop');
  }

  function bindFields(root) {
    // teks / angka / email
    $$('input[data-kind="text"], textarea[data-kind="text"], input[data-kind="number"], input[data-kind="int"], input[data-kind="email"]', root).forEach((el) => {
      const k = el.dataset.key;
      const hasTemplate = model.templateValue(k) != null && !(k in model.state.inputs);
      if (hasTemplate) el.addEventListener('focus', () => { if (!(k in model.state.inputs)) el.select(); }, { once: true });
      const commit = () => {
        const kind = el.dataset.kind;
        let val = M.parseTyped(el.value, kind === 'text' && el.tagName === 'TEXTAREA' ? '@' : model.nf(k));
        if (kind === 'email' || el.tagName === 'TEXTAREA') val = el.value === '' ? null : el.value;
        const v = model.validation(k);
        const res = M.checkValidation(v, val);
        if (!res.ok && res.level === 'stop') { showFieldMsg(el, res.msg, 'stop'); return false; }
        showFieldMsg(el, res.ok ? null : res.msg, res.level);
        model.set(k, val);
        return true;
      };
      if (el.dataset.kind === 'text') {
        let t; el.addEventListener('input', () => { clearTimeout(t); t = setTimeout(commit, 300); });
        el.addEventListener('change', () => { clearTimeout(t); commit(); });
      } else {
        el.addEventListener('change', commit);
        el.addEventListener('input', () => showFieldMsg(el, null));
      }
    });
    $$('input[data-kind="date"]', root).forEach((el) => {
      el.addEventListener('change', () => {
        const k = el.dataset.key;
        if (!el.value) { model.set(k, null); return; }
        const [y, m, d] = el.value.split('-').map(Number);
        model.set(k, E.dateToSerial(y, m, d));
      });
    });
    $$('select[data-kind="select"]', root).forEach((el) => {
      el.addEventListener('change', () => { model.set(el.dataset.key, el.value === '' ? null : el.value); refreshLevels(root); });
    });
    $$('.multi[data-kind="multi"]', root).forEach((box) => {
      const k = box.dataset.key;
      const rerender = () => {
        const wrap = box.parentElement; // .multi-wrap
        const tmp = document.createElement('div');
        tmp.innerHTML = fieldMulti(k, model.options(k) || [], $('select', box).getAttribute('aria-label'));
        wrap.innerHTML = tmp.firstChild.innerHTML;
        bindFields(wrap);
      };
      $('select[data-multi-add]', box).addEventListener('change', (ev) => { const v = ev.target.value; if (!v) return; model.multiAdd(k, v); rerender(); });
      $$('[data-chip-del]', box).forEach((b) => b.addEventListener('click', () => {
        const parts = model.inputText(k).split(', ');
        parts.splice(+b.dataset.chipDel, 1);
        model.set(k, parts.length ? parts.join(', ') : null); rerender();
      }));
    });
    $$('input[data-kind="bool"]', root).forEach((el) => el.addEventListener('change', () => model.set(el.dataset.key, el.checked)));
  }

  function refreshComputed() {
    $$('[data-f]').forEach((el) => { const v = model.display(el.dataset.f); if (el.textContent !== v) el.textContent = v; });
    $$('[data-chart]').forEach((el) => drawChart(el));
    renderNav();
    updateHeader();
  }
  function refreshLevels(root) {
    $$('.lvl[data-lvl]', root || document).forEach((el) => {
      const v = model.inputText(el.dataset.lvl);
      el.textContent = v || '—';
      el.className = 'lvl' + (v ? ' lvl-' + v.toLowerCase().replace(/[^a-z]+/g, '-') : '');
    });
  }
  function drawChart(el) {
    const def = WB.charts.find((c) => c.file === el.dataset.chart);
    if (!def) return;
    el.innerHTML = C.render(model.engine, def, { width: +el.dataset.w || 520, height: +el.dataset.h || 300, title: el.dataset.title });
  }
  const chartBox = (file, w, h, title) => `<div data-chart="${file}" data-w="${w || 520}" data-h="${h || 300}" data-title="${esc(title || '')}"></div>`;

  // ================================================================ header & navigasi
  function updateHeader() {
    const nm = E.toText(model.get(key(ATTR, 'B5')));
    const yr = E.toText(model.get(key(ATTR, 'B63')));
    $('#currentName').textContent = (nm ? nm : 'Penilaian baru (nama kawasan belum diisi)') + (yr ? ' — ' + yr : '') + (model.state.finalized ? ' · FINAL' : '');
  }
  function renderNav() {
    const cur = location.hash.replace(/^#\//, '') || 's/Tentang';
    const link = (p, text) => `<a class="nav-link${cur === p.id ? ' active' : ''}" href="${href(p)}"><span class="nl-text">${esc(text || p.title)}</span></a>`;
    const byType = (t) => pages.filter((p) => p.type === t);
    const prog = model.progress();
    let html = '';
    html += `<div class="nav-group"><div class="nav-title">Mulai</div>${byType('tentang').map((p) => link(p)).join('')}${byType('pengenalan').map((p) => link(p)).join('')}</div>`;
    html += `<div class="nav-group"><div class="nav-title">Data Kawasan</div>${byType('atribut').map((p) => link(p)).join('')}${byType('ancaman').map((p) => link(p)).join('')}</div>`;
    html += `<div class="nav-group"><div class="nav-title">Penilaian</div>${byType('daftar').map((p) => link(p)).join('')}`;
    html += `<div class="progress">${prog.done} dari ${prog.total} butir lengkap${prog.warn ? ` · <span style="color:var(--warn)">${prog.warn} perlu dilengkapi</span>` : ''}<div class="bar"><span style="width:${(prog.done / prog.total * 100).toFixed(1)}%"></span></div></div>`;
    html += '<div class="nav-q">';
    for (const p of seqPages) {
      if (p.type === 'question') {
        const q = model.question(p.sheet);
        const st = model.status(q);
        const label = p.sheet.replace(/\.$/, '');
        html += `<a href="${href(p)}" class="${st}${cur === p.id ? ' active' : ''}" title="${esc(q.title)} — ${st === 'done' ? 'lengkap' : st === 'na' ? 'tidak relevan' : st === 'warn' ? 'perlu dilengkapi' : 'belum dijawab'}">${esc(label)}</a>`;
      } else {
        const short = p.type === 'detail-nilai' ? 'Nilai' : p.type === 'detail-spesies' ? 'Spesies' : 'Habitat';
        html += `<a href="${href(p)}" class="detail${cur === p.id ? ' active' : ''}" title="${esc(p.title)}">${short}</a>`;
      }
    }
    html += '</div></div>';
    html += `<div class="nav-group"><div class="nav-title">Hasil & Tindak Lanjut</div>${byType('dashboard').map((p) => link(p)).join('')}${byType('aksi').map((p) => link(p)).join('')}${byType('ba').map((p) => link(p)).join('')}${byType('rencana').map((p) => link(p)).join('')}</div>`;
    html += `<div class="nav-group"><div class="nav-title">Bantuan</div>${byType('bantuan').map((p) => link(p)).join('')}</div>`;
    html += `<div class="nav-group"><div class="nav-title">Data</div>${appPages.map((p) => link(p)).join('')}</div>`;
    $('#sidebar').innerHTML = html;
  }
  function shapeButtons(sheet, skip) {
    const s = sheetOf(sheet);
    if (!s || !s.shapes) return '';
    const out = [];
    for (const sh of s.shapes) {
      const t = clean(sh.text);
      if (skip && skip.test(t)) continue;
      const h = linkToHref(sh.link);
      if (h) out.push(`<a class="btn btn-sm" href="${h}">${esc(t)}</a>`);
      else if (/^SAVE$/i.test(t)) out.push(`<button class="btn btn-sm" data-act="save-file">${esc(t)}</button>`);
      else if (/eksport|export/i.test(t)) out.push(`<a class="btn btn-sm" href="#/data">${esc(t)}</a>`);
    }
    return out.join('');
  }

  // ================================================================ halaman
  const R = {};

  R.tentang = (p) => {
    const s = p.sheet;
    const rows = [28, 29, 30].map((r) => `<div class="attr-row" style="grid-template-columns:minmax(160px,.8fr) 2fr;padding:8px 0;border:0">
      <div class="lbl">${esc(raw(s, 'B' + r))}</div><div>${fieldText(key(s, 'D' + r), { label: raw(s, 'B' + r) })}</div></div>`).join('');
    return `<div class="page">
      <div class="cover card">
        <div class="card-body"><img class="sampul" src="assets/img/sampul-mett4.jpg" alt="Sampul METT-4 Kementerian Kehutanan Republik Indonesia"></div>
        <div class="card-body">
          <h1 class="page-h">Management Effectiveness Tracking Tool (METT-4)</h1>
          <p class="page-sub">Versi web METT-4 KSDAE v4.4 — seluruh pertanyaan, pilihan jawaban, rumus skor, dan dashboard mengikuti file Excel asli.</p>
          ${rows}
          <div class="toolbar" style="margin-top:12px">${shapeButtons(s)}</div>
        </div>
      </div>
      <div class="quick">
        <a href="${hrefSheet(ATTR)}">1. Atribut Kawasan<span>Informasi latar belakang kawasan</span></a>
        <a href="${hrefSheet(THREAT)}">2. Penilaian Ancaman<span>62 sub-ancaman dalam 15 kategori</span></a>
        <a href="${hrefSheet(DAFTAR)}">3. Pertanyaan &amp; Skor<span>42 butir penilaian</span></a>
        <a href="${hrefSheet('Dashboard')}">4. Dashboard<span>Skor per elemen &amp; grafik</span></a>
        <a href="${hrefSheet('Print BA')}">5. Berita Acara<span>Cetak / simpan PDF</span></a>
      </div>
      <div class="card" style="margin-top:18px"><div class="card-body note">${esc(raw(s, 'B36'))}</div></div>
      <p class="small muted">Data penilaian tersimpan otomatis di browser ini. Gunakan <strong>Simpan file</strong> untuk membuat salinan (.json) yang dapat dibuka di perangkat lain atau dikirim ke pusat.</p>
    </div>`;
  };

  R.pengenalan = (p) => {
    const s = sheetOf(p.sheet);
    const merged = new Set(s.merged.filter((m) => /^A\d+:B\d+$/.test(m)).map((m) => m.split(':')[0]));
    let body = '';
    const maxRow = Math.max(...Object.keys(s.cells).map((a) => E.parseA1(a).r));
    for (let r = 3; r <= maxRow; r++) {
      const a = raw(p.sheet, 'A' + r), b = raw(p.sheet, 'B' + r);
      if (!a && !b) continue;
      if (a && !b && merged.has('A' + r)) body += `<div class="bar-title">${esc(a)}</div>`;
      else body += `<div class="intro-item"><div class="ii-h">${esc(a)}</div><div class="ii-b">${esc(b)}</div></div>`;
    }
    return `<div class="page"><div class="card"><h1 class="sheet-title">${esc(raw(p.sheet, 'A1'))}</h1>
      <div class="card-body note" style="border-radius:0">${esc(raw(p.sheet, 'A2'))}</div>${body}</div>
      <div class="toolbar">${shapeButtons(p.sheet)}</div></div>`;
  };

  // ---- Atribut Kawasan
  const ATTR_SECTIONS = [30, 36, 41, 58];
  const ATTR_SUB = [12, 13, 14, 15, 16, 22, 23, 25, 26, 27, 31, 32, 33, 34, 35, 37, 38, 39, 40, 42, 43, 44, 45, 46, 48, 49, 50, 51, 52, 53, 54, 55, 56, 59, 60, 61, 62];
  R.atribut = (p) => {
    const s = p.sheet;
    let html = `<div class="attr-row" style="background:var(--bar);color:#fff;font-weight:600">
      <div>${esc(raw(s, 'A2'))}</div><div>${esc(raw(s, 'B2'))}</div><div>${esc(raw(s, 'C2'))}</div></div>`;
    for (let r = 3; r <= 64; r++) {
      const a = raw(s, 'A' + r);
      if (!a) continue;
      const help = raw(s, 'D' + r);
      if (ATTR_SECTIONS.includes(r)) {
        const link = sheetOf(s).links['B' + r];
        const b = raw(s, 'B' + r);
        html += `<div class="attr-section">${esc(a)}${b ? ` <a href="${linkToHref(link) || '#'}" style="color:#ffd9bf;font-weight:400">${esc(b)}</a>` : ''}${help ? `<span class="muted">${esc(help)}</span>` : ''}</div>`;
        continue;
      }
      const v = model.validation(key(s, 'B' + r));
      const prompt = v && v.prompt ? v.prompt : '';
      html += `<div class="attr-row${ATTR_SUB.includes(r) ? ' sub' : ''}">
        <div class="lbl">${esc(a.trim())}</div>
        <div><span class="col-lbl">${esc(raw(s, 'B2').trim())}</span>${autoField(s, 'B' + r, { label: a })}</div>
        <div><span class="col-lbl">${esc(raw(s, 'C2'))}</span>${fieldText(key(s, 'C' + r), { label: 'Keterangan ' + a })}</div>
        <div class="hint">${esc([help, prompt && prompt !== help ? '› ' + prompt : ''].filter(Boolean).join('\n'))}</div>
      </div>`;
    }
    return `<div class="page"><div class="toolbar">${shapeButtons(s)}</div>
      <div class="card"><h1 class="sheet-title">${esc(raw(s, 'A1').trim())}</h1>${html}</div></div>`;
  };

  // ---- Penilaian Ancaman
  R.ancaman = (p) => {
    const s = p.sheet, sh = sheetOf(s);
    const cats = [];
    const mergedA = sh.merged.filter((m) => /^A\d+:A\d+$/.test(m)).map((m) => { const [a, b] = m.split(':'); return [E.parseA1(a).r, E.parseA1(b).r]; });
    for (let r = 3; r <= 64; r++) {
      const a = raw(s, 'A' + r);
      if (!a) continue;
      const mg = mergedA.find((x) => x[0] === r);
      cats.push({ r1: r, r2: mg ? mg[1] : r, text: a });
    }
    const hdr = (col) => raw(s, col + '2');
    const short = (col) => hdr(col).split('\n')[0].trim();
    const lv = (r) => { const e = model.inputText(key(s, 'E' + r)); return e; };
    let html = '';
    for (const c of cats) {
      const [title, ...rest] = c.text.split('\n');
      html += `<section class="threat-cat"><div class="cat-h"><strong>${esc(title.trim())}</strong><span>${esc(rest.join('\n').trim())}</span></div>`;
      for (let r = c.r1; r <= c.r2; r++) {
        const b = raw(s, 'B' + r);
        const filled = ['C', 'D', 'E', 'F', 'G', 'H', 'I'].some((col) => model.inputText(key(s, col + r)));
        html += `<div class="threat${filled ? '' : ' collapsed'}" data-row="${r}">
          <div class="t-name"><span>${b ? esc(b.trim()) : `<span class="muted">Ancaman lainnya:</span>`}</span>
            <span style="display:flex;gap:6px;align-items:center"><span class="lvl" data-lvl="${esc(key(s, 'E' + r))}"></span>
            <button class="btn btn-sm" data-toggle-threat>${filled ? 'Tutup' : 'Isi'}</button></span></div>
          <div class="t-grid">
            ${b ? '' : `<div class="span2" style="grid-column:1/-1"><label>${esc(short('B'))}</label>${fieldText(key(s, 'B' + r), { label: 'Nama ancaman' })}</div>`}
            <div class="span2"><label title="${esc(hdr('C'))}">${esc(short('C'))}</label>${fieldText(key(s, 'C' + r), { multiline: true, rows: 2, label: short('C') })}</div>
            <div class="span2"><label title="${esc(hdr('D'))}">${esc(short('D'))}</label>${autoField(s, 'D' + r, { label: short('D') })}</div>
            <div><label title="${esc(hdr('E'))}">${esc(short('E'))}</label>${autoField(s, 'E' + r, { label: short('E') })}</div>
            <div><label title="${esc(hdr('F'))}">${esc(short('F'))}</label>${autoField(s, 'F' + r, { label: short('F') })}</div>
            <div class="span2"><label title="${esc(hdr('G'))}">${esc(short('G'))}</label>${fieldText(key(s, 'G' + r), { label: short('G') })}</div>
            <div class="span2"><label title="${esc(hdr('H'))}">${esc(short('H'))}</label>${fieldText(key(s, 'H' + r), { multiline: true, rows: 2, label: short('H') })}</div>
            <div class="span2"><label title="${esc(hdr('I'))}">${esc(short('I'))}</label>${fieldText(key(s, 'I' + r), { multiline: true, rows: 2, label: short('I') })}</div>
          </div></div>`;
      }
      html += '</section>';
    }
    const vE = model.validation(key(s, 'E3')), vF = model.validation(key(s, 'F3'));
    const nilai = [31, 32, 33, 34, 35].map((r) => model.inputText(key(ATTR, 'B' + r))).filter(Boolean);
    return `<div class="page"><div class="toolbar">${shapeButtons(s)}<span class="spacer"></span>
        <button class="btn btn-sm" data-act="expand-all">Buka semua</button><button class="btn btn-sm" data-act="collapse-all">Tutup yang kosong</button></div>
      <div class="card"><h1 class="sheet-title">${esc(raw(s, 'A1'))}</h1>
        <div class="card-body" style="display:grid;grid-template-columns:repeat(auto-fit,minmax(260px,1fr));gap:12px">
          <div class="help"><span class="help-h">${esc(short('E'))}</span>${esc(vE && vE.prompt)}</div>
          <div class="help"><span class="help-h">${esc(short('F'))}</span>${esc(vF && vF.prompt)}</div>
          <div class="note"><strong>${esc(short('D'))}</strong>\n${nilai.length ? nilai.map((n, i) => (i + 1) + '. ' + esc(n)).join('\n') : 'Nilai penting belum diisi pada lembar Atribut Kawasan (baris "Nilai penting 1–5").'}</div>
        </div></div>
      ${html}</div>`;
  };

  // ---- Daftar Pertanyaan & Skor
  R.daftar = (p) => {
    const s = p.sheet;
    let rows = '';
    for (let r = 6; r <= 51; r++) {
      const b = raw(s, 'B' + r), c = raw(s, 'C' + r), g = raw(s, 'G' + r);
      if (!b && !c) continue;
      const link = sheetOf(s).links['C' + r];
      const h = linkToHref(link);
      if (r === 51) {
        rows += `<tr class="total"><td></td><td>${esc(c)}</td><td class="num">${computed(s, 'D' + r)}</td><td class="num">${computed(s, 'E' + r)}</td><td class="num">${computed(s, 'F' + r)}</td><td></td></tr>`;
        continue;
      }
      if (!b) { rows += `<tr class="section"><td></td><td colspan="5"><a href="${h || '#'}">${esc(c)}</a></td></tr>`; continue; }
      const q = model.schema.questions.find((x) => x.daftarRow === r);
      const st = q ? model.status(q) : '';
      const badge = st === 'done' ? '<span class="badge ok">lengkap</span>' : st === 'na' ? '<span class="badge na">tidak relevan</span>' : st === 'warn' ? '<span class="badge warn">perlu dilengkapi</span>' : '<span class="badge">belum</span>';
      rows += `<tr><td class="center"><strong>${esc(b)}</strong></td><td><a href="${h || '#'}">${esc(c)}</a> ${badge}</td>
        <td class="num">${computed(s, 'D' + r)}</td><td class="num"><strong>${computed(s, 'E' + r)}</strong></td>
        <td class="num" style="width:110px">${fieldText(key(s, 'F' + r), { kind: 'int', label: 'Nilai sebelumnya ' + b })}</td><td>${esc(g)}</td></tr>`;
    }
    return `<div class="page wide">
      <div class="toolbar">${shapeButtons(s)}<span class="spacer"></span>
        <button class="btn btn-sm" data-act="import-prev">Isi kolom "sebelumnya" dari file tahun lalu…</button></div>
      <div class="card"><h1 class="sheet-title">${esc(raw(s, 'B2'))} <span style="float:right;font-size:15px;font-weight:400">${esc(raw(s, 'D2'))}: ${computed(s, 'E3')}</span></h1>
        <details class="card-body"><summary class="small">Catatan lembar ini</summary><div class="note" style="margin-top:8px">${esc(raw(s, 'B1'))}</div></details>
        <div class="tbl-wrap"><table class="grid"><thead><tr><th class="center">${esc(raw(s, 'B5'))}</th><th>${esc(raw(s, 'C5'))}</th><th class="num">${esc(raw(s, 'D5'))}</th><th class="num">${esc(raw(s, 'E5'))}</th><th class="num">${esc(raw(s, 'F5'))}</th><th>${esc(raw(s, 'G5'))}</th></tr></thead>
        <tbody>${rows}</tbody></table></div></div></div>`;
  };

  // ---- halaman pertanyaan
  function seqNav(p) {
    const i = seqPages.findIndex((x) => x.id === p.id);
    const prev = seqPages[i - 1], next = seqPages[i + 1];
    return `<div class="q-nav">
      <a class="btn btn-sm" href="${hrefSheet(DAFTAR)}">← ${esc("Daftar Pertanyaan & Skor")}</a>
      ${prev ? `<a class="btn btn-sm" href="${href(prev)}">‹ ${esc(prev.title.replace(/\.$/, ''))}</a>` : ''}
      ${next ? `<a class="btn btn-sm btn-teal" href="${href(next)}">Indikator selanjutnya: ${esc(next.title.replace(/\.$/, ''))} ›</a>` : `<a class="btn btn-sm btn-teal" href="${hrefSheet('Aksi meningkatkan Pengelolaan')}">Selesai → Aksi meningkatkan Pengelolaan</a>`}
    </div>`;
  }
  R.question = (p) => {
    const q = model.question(p.sheet), s = p.sheet;
    const na = model.isNA(q);
    const [naMain, naHint] = q.naLabel.split(/\s{3,}/);
    const extra = shapeButtons(s, /kembali|selanjutnya/i);
    let answer = '';
    if (q.kind === 'radio') {
      answer = `<div class="q-section"><div class="help"><span class="help-h">${esc((q.guidance.split('\n')[0] || '').trim())}</span>${esc(q.guidance.split('\n').slice(1).join('\n').trim())}</div></div>
        <div class="bar-title">${esc(q.optionsHeader)}</div><div class="q-section" role="radiogroup" aria-label="${esc(q.optionsHeader)}">` +
        q.options.map((o) => {
          const sel = model.control(s, o.name);
          return `<label class="opt${sel ? ' selected' : ''}${na ? ' disabled' : ''}"><input type="radio" name="opt-${esc(s)}" data-opt="${esc(o.name)}"${sel ? ' checked' : ''}${na ? ' disabled' : ''}>
            <span class="letter">${esc(o.letter)}</span><span>${esc(o.text)}</span></label>`;
        }).join('') + '</div>';
    } else {
      const g = q.guidance;
      const m = /^(Bantuan untuk menjawab:?)([\s\S]*)$/.exec(g);
      answer = `<div class="q-section"><div class="help">${m ? `<span class="help-h">${esc(m[1])}</span>${esc(m[2].trim())}` : esc(g)}</div></div>
        <div class="bar-title">${esc(q.boxesHeader)}</div><div class="q-section">` +
        q.boxes.map((b) => {
          const on = model.control(s, b.name);
          return `<label class="opt box-opt${on ? ' selected' : ''}${na ? ' disabled' : ''}"><input type="checkbox" data-box="${esc(b.name)}"${on ? ' checked' : ''}${na ? ' disabled' : ''}><span>${esc(b.text)}</span></label>`;
        }).join('') + (q.vbaNote ? `<div class="alert alert-warn small">Catatan teknis: ${esc(q.vbaNote)}</div>` : '') + '</div>';
    }
    const sources = q.sources.map((src) => `<label class="check"><input type="checkbox" data-key="${esc(src.key)}" data-kind="bool"${model.get(src.key) === true ? ' checked' : ''}><span>${esc(src.label)}</span></label>
      <div>${src.textKey ? fieldText(src.textKey, { label: 'Rincian: ' + src.label }) : ''}</div>`).join('');
    const thanks = raw(s, 'D6');
    return `<div class="page">
      <div class="q-head">${seqNav(p)}<span class="score-pill">Nilai tercatat <strong data-f="${esc(key(DAFTAR, 'E' + q.daftarRow))}">${esc(disp(DAFTAR, 'E' + q.daftarRow))}</strong> / <span data-f="${esc(key(DAFTAR, 'D' + q.daftarRow))}">${esc(disp(DAFTAR, 'D' + q.daftarRow))}</span></span></div>
      ${extra ? `<div class="toolbar">${extra}</div>` : ''}
      <article class="card q-card">
        <h1 class="sheet-title">${esc(q.title)}</h1>
        <div class="na-row"><label class="check"><input type="checkbox" data-na="1"${na ? ' checked' : ''}><span><strong>${esc((naMain || '').trim())}</strong></span></label><span class="muted">${esc((naHint || '').trim())}</span></div>
        <div class="q-section" ${na ? '' : 'hidden'} id="naReason"><label class="small"><strong>${esc(raw(s, 'A3'))}</strong></label>${fieldText(q.naReasonKey, { multiline: true, rows: 2, label: 'Penjelasan tidak relevan' })}</div>
        ${q.kind === 'radio' ? `<div class="q-question">${esc(q.question)}</div>` : ''}
        ${answer}
        <div class="bar-title">${esc(q.justLabel)}</div><div class="q-section">${fieldText(q.justKey, { multiline: true, cls: 'tall', label: q.justLabel })}</div>
        <div class="bar-title">${esc(q.actionLabel)}</div><div class="q-section">${fieldText(q.actionKey, { multiline: true, label: q.actionLabel })}</div>
        <div class="bar-title">${esc(q.sourcesLabel)}</div><div class="q-section"><div class="src-grid">${sources}</div></div>
        <div class="q-section" id="qMsgs" aria-live="polite"></div>
      </article>
      ${thanks ? `<div class="card"><div class="thanks">${esc(thanks)}</div></div>` : ''}
      <div class="q-head">${seqNav(p)}</div>
    </div>`;
  };
  function bindQuestion(p, root) {
    const q = model.question(p.sheet);
    const rerender = () => render();
    const na = $('input[data-na]', root);
    na.addEventListener('change', () => { model.setNA(q, na.checked); rerender(); });
    $$('input[data-opt]', root).forEach((el) => el.addEventListener('change', () => { model.selectOption(q, el.dataset.opt); rerender(); }));
    $$('input[data-box]', root).forEach((el) => el.addEventListener('change', () => { model.toggleBox(q, el.dataset.box, el.checked); rerender(); }));
    const showMsgs = () => {
      const msgs = model.messages(q);
      $('#qMsgs', root).innerHTML = msgs.length ? msgs.map((m) => `<div class="alert alert-warn">${esc(m.msg)}</div>`).join('') : (model.isAnswered(q) || model.isNA(q) ? '<div class="alert alert-ok">Isian butir ini sudah lengkap.</div>' : '<div class="muted small">Pilih salah satu jawaban, atau centang “tidak relevan”.</div>');
    };
    showMsgs();
    model.onceChange = showMsgs;
  }

  // ---- detail
  function detailTable(s, headerRow, rows, cols, opts) {
    opts = opts || {};
    const head = cols.map((c) => `<th>${esc(raw(s, c + headerRow).trim())}</th>`).join('');
    let body = '';
    for (const r of rows) {
      body += '<tr>' + cols.map((c) => {
        const cell = sheetOf(s).cells[c + r];
        if (cell && cell.f) return `<td>${computed(s, c + r)}</td>`;
        const k = key(s, c + r);
        const o = model.options(k);
        if (o) return `<td>${fieldSelect(k, o, raw(s, c + headerRow))}</td>`;
        const long = (opts.long || []).includes(c);
        return `<td>${fieldText(k, { multiline: long, rows: 2, label: raw(s, c + headerRow) })}</td>`;
      }).join('') + '</tr>';
    }
    return `<div class="tbl-wrap"><table class="grid compact"><thead><tr>${head}</tr></thead><tbody>${body}</tbody></table></div>`;
  }
  function detailPage(p, cfg) {
    const s = p.sheet, sh = sheetOf(s);
    const insuff = sh.controls.find((c) => c.type === 'CheckBox' && !c.link);
    let help = '';
    for (let r = cfg.helpFrom; r <= cfg.helpTo; r++) {
      const t = raw(s, 'A' + r);
      if (!t || r === cfg.insuffRow) continue;
      help += `<p>${esc(t)}</p>`;
    }
    const insuffHtml = insuff ? `<label class="check"><input type="checkbox" data-ctl="${esc(insuff.name)}"${model.control(s, insuff.name) ? ' checked' : ''}><span>${esc(raw(s, 'A' + cfg.insuffRow))}</span></label>` : '';
    return `<div class="page wide"><div class="q-head">${seqNav(p)}</div>
      <div class="card"><h1 class="sheet-title">${esc(raw(s, 'A1'))}</h1>
        <div class="card-body"><div class="note">${esc(raw(s, 'A2'))}</div><div style="margin:12px 0">${insuffHtml}</div>
          <details><summary>Panduan pengisian</summary><div class="prose">${help}</div></details></div>
        ${detailTable(s, cfg.headerRow, cfg.rows, cfg.cols, { long: cfg.long })}
        ${cfg.footRow ? `<div class="card-body small muted">${esc(raw(s, 'A' + cfg.footRow))}</div>` : ''}
      </div><div class="q-head">${seqNav(p)}</div></div>`;
  }
  R['detail-nilai'] = (p) => detailPage(p, { helpFrom: 4, helpTo: 17, insuffRow: 3, headerRow: 18, rows: [19, 20, 21, 22, 23], cols: ['A', 'B', 'C', 'D', 'E'], long: ['D', 'E'] });
  R['detail-spesies'] = (p) => detailPage(p, { helpFrom: 4, helpTo: 9, insuffRow: 3, headerRow: 11, rows: [12, 13, 14, 15, 16], cols: ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I'], long: ['H', 'I'], footRow: 18 });
  R['detail-habitat'] = (p) => detailPage(p, { helpFrom: 4, helpTo: 8, insuffRow: 3, headerRow: 10, rows: [11, 12, 13, 14, 15], cols: ['A', 'B', 'C', 'D', 'E', 'F', 'G'], long: ['F', 'G'], footRow: 17 });

  // ---- Dashboard
  function rangeTable(sheet, ref, opts) {
    opts = opts || {};
    const [a, b] = ref.split(':');
    const p1 = E.parseA1(a), p2 = E.parseA1(b);
    let html = '<div class="tbl-wrap"><table class="grid compact"><thead><tr>';
    for (let c = p1.c; c <= p2.c; c++) html += `<th${c > p1.c ? ' class="num"' : ''}>${esc(disp(sheet, E.toA1(c, p1.r)))}</th>`;
    html += '</tr></thead><tbody>';
    for (let r = p1.r + 1; r <= p2.r; r++) {
      const first = E.toText(model.engine.get(sheet, E.toA1(p1.c, r)));
      html += `<tr class="${opts.elementColors ? 'el-' + esc(first) : ''}${first === 'Total' ? ' total' : ''}">`;
      for (let c = p1.c; c <= p2.c; c++) html += `<td${c > p1.c ? ' class="num"' : ''}>${computed(sheet, E.toA1(c, r))}</td>`;
      html += '</tr>';
    }
    return html + '</tbody></table></div>';
  }
  R.dashboard = (p) => {
    const s = p.sheet;
    const tile = (label, addr, sub) => `<div class="stat"><div class="s-l">${esc(label)}</div><div class="s-v">${computed(s, addr)}</div><div class="s-s">${sub || ''}</div></div>`;
    const link = (addr) => { const t = sheetOf(s).links[addr]; return `<a href="${linkToHref(t) || '#'}">${esc(raw(s, addr))}</a>`; };
    return `<div class="page wide">
      <div class="toolbar">${shapeButtons(s)}</div>
      <div class="card"><div class="card-body" style="display:flex;flex-wrap:wrap;gap:10px 30px;align-items:baseline">
        <h1 class="page-h" style="margin:0">${computed(s, 'B1')}</h1><span class="muted">Tanggal penilaian: ${computed(s, 'G1')}</span></div></div>
      <div class="stat-row">
        ${tile('Total skor METT', 'V11', 'dari maks. ' + `<span data-f="${esc(key(s, 'X11'))}">${esc(disp(s, 'X11'))}</span>`)}
        ${tile('Nilai total (%)', 'Y11')}
        ${['V6', 'V7', 'V8', 'V9', 'V10'].map((a, i) => tile(raw(s, 'U' + (6 + i)) + ' (%)', 'Y' + (6 + i))).join('')}
      </div>
      <div class="dash-grid">
        <div class="card chart-card"><div class="teal-title">${esc(raw(s, 'B3'))}</div><div class="chart-body">${chartBox('chart1.xml', 520, 340, raw(s, 'B3'))}</div></div>
        <div class="card chart-card"><div class="teal-title">${esc(raw(s, 'L3'))}</div><div class="chart-body">${chartBox('chart2.xml', 520, 340, raw(s, 'L3'))}</div></div>
        <div class="card chart-card" style="grid-row:span 2"><div class="teal-title">${esc(raw(s, 'B28'))}</div><div class="chart-body">${chartBox('chart3.xml', 560, 640, raw(s, 'B28'))}</div></div>
        <div class="card chart-card"><div class="teal-title">${esc(raw(s, 'N28'))}</div><div class="chart-body">${chartBox('chart4.xml', 480, 280, raw(s, 'N28'))}</div></div>
        <div class="card chart-card"><div class="teal-title">${esc(raw(s, 'N43'))}</div><div class="chart-body">${chartBox('chart5.xml', 480, 280, raw(s, 'N43'))}</div></div>
      </div>
      <div class="card"><div class="teal-title">${esc(raw(s, 'U3'))}</div>${rangeTable(s, 'U5:AA11', { elementColors: true })}</div>
      <div class="dash-grid">
        <div class="card"><div class="teal-title">${esc(raw(s, 'U15'))}</div>${rangeTable(s, 'U17:W22')}<div class="card-body small">${link('X23')}</div></div>
        <div class="card"><div class="teal-title">${esc(raw(s, 'U43'))}</div>${rangeTable(s, 'U45:Y50')}<div class="card-body small">${link('Y52')}</div></div>
      </div>
      <div class="card"><div class="teal-title">${esc(raw(s, 'U28'))}</div>${rangeTable(s, 'U30:AA36')}<div class="card-body small">${link('AA38')}</div></div>
    </div>`;
  };

  // ---- Aksi meningkatkan Pengelolaan
  R.aksi = (p) => {
    const s = p.sheet;
    let rows = '';
    for (let r = 5; r <= 46; r++) {
      if (!raw(s, 'A' + r)) continue;
      rows += `<tr><td class="center"><strong>${esc(raw(s, 'A' + r))}</strong></td><td>${computed(s, 'B' + r)}</td><td class="num">${computed(s, 'C' + r)}</td><td class="num">${computed(s, 'D' + r)}</td>
        <td style="min-width:220px" class="pre">${computed(s, 'E' + r)}</td>` +
        ['F', 'G', 'H', 'I', 'J'].map((c) => `<td>${fieldText(key(s, c + r), { label: raw(s, c + '4') + ' ' + raw(s, 'A' + r) })}</td>`).join('') + '</tr>';
    }
    return `<div class="page wide"><div class="toolbar">${shapeButtons(s)}</div>
      <div class="card"><h1 class="sheet-title">${esc(raw(s, 'A2'))}</h1>
        <div class="card-body note" style="border-radius:0">${esc(raw(s, 'A1'))}</div>
        <div class="tbl-wrap"><table class="grid compact"><thead><tr>${['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J'].map((c) => `<th>${esc(raw(s, c + '4'))}</th>`).join('')}</tr></thead><tbody>${rows}</tbody></table></div>
      </div></div>`;
  };

  // ---- Print BA (Berita Acara)
  function baCell(addr, opts) {
    opts = opts || {};
    const s = 'Print BA', cell = sheetOf(s).cells[addr];
    if (cell && cell.f) return computed(s, addr, opts.cls);
    const k = key(s, addr);
    const v = model.inputText(k);
    return `<span class="editable ${opts.cls || ''}" contenteditable="true" data-edit="${esc(k)}" spellcheck="false">${esc(v)}</span>`;
  }
  R.ba = (p) => {
    const s = p.sheet;
    const tr = (r) => `<tr><td class="center">${computed(s, 'B' + r)}</td><td>${esc(raw(s, 'C' + r))}</td><td class="center">${computed(s, 'E' + r)}</td><td class="center">${computed(s, 'F' + r)}</td><td class="center">${computed(s, 'G' + r)}</td><td>${baCell('H' + r)}</td></tr>`;
    const sig = (n, ni, ins) => `<div class="s-block"><div></div><div><strong>${baCell(n)}</strong><br>${baCell(ni)}<br>${baCell(ins)}</div></div>`;
    const chartTitle = (file) => { const c = WB.charts.find((x) => x.file === file); return c && c.groupTitle ? c.groupTitle : ''; };
    const smallTable = (ref) => rangeTable('Print BA', ref, { elementColors: ref === 'V6:Z12' });
    return `<div class="page wide">
      <div class="print-actions no-print"><button class="btn btn-primary" data-act="print">Cetak / Simpan PDF</button>${shapeButtons(s)}
        <span class="small muted" style="align-self:center">Teks bergaris titik (No. BA, kota/tanggal, nama, NIP, instansi, keterangan) dapat diketik langsung.</span></div>
      <div class="paper">
        <div class="kop"><img src="assets/img/logo-kemenhut.png" alt="Logo"><div class="k-txt"><div class="k1">${esc(raw(s, 'C4'))}</div><div class="k2">${esc(raw(s, 'C5'))}</div><div class="k3">${computed(s, 'C6')}</div></div></div>
        <div class="ba-title"><div class="t1">${esc(raw(s, 'A8'))}</div><div>${baCell('A9')}</div><div class="t1">${esc(raw(s, 'A10'))}</div><div>${computed(s, 'A11')}</div><div>${computed(s, 'A12')}</div></div>
        <p class="justify">${computed(s, 'B14')}</p>
        <div class="sec-h"><span>${esc(raw(s, 'B21'))}</span><span>${esc(raw(s, 'C21'))}</span></div>
        <table class="pt"><thead><tr><th>${esc(raw(s, 'B22'))}</th><th>${esc(raw(s, 'C22'))}</th><th>${esc(raw(s, 'E22'))}</th><th>${esc(raw(s, 'F22'))}</th><th>${esc(raw(s, 'G22'))}</th><th>${esc(raw(s, 'H22'))}</th></tr></thead>
          <tbody>${[23, 24, 25, 26, 27].map(tr).join('')}<tr><td colspan="2" class="center"><strong>${esc(raw(s, 'B28'))}</strong></td><td class="center"><strong>${computed(s, 'E28')}</strong></td><td class="center"><strong>${computed(s, 'F28')}</strong></td><td></td><td></td></tr></tbody></table>
        <div class="sec-h"><span>${esc(raw(s, 'B31'))}</span><span>${esc(raw(s, 'C31'))}</span></div>
        <div class="chart-pair">
          <div class="chart-box"><div class="cb-h">${esc(chartTitle('chart7.xml'))}</div>${chartBox('chart7.xml', 420, 300)}</div>
          <div class="chart-box"><div class="cb-h">${esc(chartTitle('chart6.xml'))}</div>${chartBox('chart6.xml', 420, 300)}</div>
        </div>
        <div class="page-break"></div>
        <div class="chart-box" style="margin-top:14px"><div class="cb-h">${esc(chartTitle('chart8.xml'))}</div>${chartBox('chart8.xml', 640, 420)}</div>
        <div class="chart-pair" style="margin-top:12px">
          <div class="chart-box orange"><div class="cb-h">${esc(chartTitle('chart9.xml'))}</div>${chartBox('chart9.xml', 420, 240)}</div>
          <div class="chart-box orange"><div class="cb-h">${esc(chartTitle('chart10.xml'))}</div>${chartBox('chart10.xml', 420, 240)}</div>
        </div>
        <div class="page-break"></div>
        <div style="margin-top:14px"><div class="r-title">${esc(raw(s, 'V4'))}</div>${smallTable('V6:Z12')}</div>
        <div style="margin-top:12px"><div class="r-title">${esc(raw(s, 'V14'))}</div>${smallTable('V16:X21')}</div>
        <div style="margin-top:12px"><div class="r-title">${esc(raw(s, 'V24'))}</div>${smallTable('V26:Z31')}</div>
        <div style="margin-top:12px"><div class="r-title">${esc(raw(s, 'V34'))}</div>${smallTable('V36:Z41')}</div>
        <div class="page-break"></div>
        <p style="margin-top:14px">${esc(raw(s, 'AB4'))}</p>
        <div class="sign">
          <div></div><div>${baCell('AG7')}</div>
          <div><strong>${esc(raw(s, 'AB9'))}</strong></div><div><strong>${esc(raw(s, 'AG9'))}</strong></div>
          ${sig('AB14', 'AB15', 'AB16')}${sig('AG14', 'AG15', 'AG16')}
          ${sig('AB20', 'AB21', 'AB22')}${sig('AG20', 'AG21', 'AG22')}
          ${sig('AB26', 'AB27', 'AB28')}<div></div>
          ${sig('AB33', 'AB34', 'AB35')}<div></div>
        </div>
        <div class="sign" style="grid-template-columns:1fr"><div class="s-block" style="min-height:140px"><div>${esc(raw(s, 'AA37'))}<br>${esc(raw(s, 'AA38'))}</div>
          <div><strong>${baCell('AA43')}</strong><br>${baCell('AA44')}<br>${baCell('AA45')}</div></div></div>
      </div></div>`;
  };

  // ---- Print_2 (Rencana Aksi)
  R.rencana = (p) => {
    const s = p.sheet;
    const cols = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J'];
    let rows = '';
    for (let r = 8; r <= 53; r++) {
      if (!raw(s, 'A' + r) && !sheetOf(s).cells['C' + r]) continue;
      rows += '<tr>' + cols.map((c) => {
        const cell = sheetOf(s).cells[c + r];
        return `<td class="${['A', 'C', 'D'].includes(c) ? 'center' : ''} pre">${cell && cell.f ? computed(s, c + r) : esc(raw(s, c + r))}</td>`;
      }).join('') + '</tr>';
    }
    return `<div class="page wide">
      <div class="print-actions no-print"><button class="btn btn-primary" data-act="print">Cetak / Simpan PDF (lanskap)</button>${shapeButtons(s)}
        <span class="small muted" style="align-self:center">Isi kolom Waktu, Penanggung Jawab, dst. pada lembar “Aksi meningkatkan Pengelolaan”.</span></div>
      <div class="paper landscape"><div class="r-title">${computed(s, 'A5')}</div>
        <table class="pt rencana"><thead><tr>${cols.map((c) => `<th>${esc(raw(s, c + '7'))}</th>`).join('')}</tr></thead><tbody>${rows}</tbody></table></div></div>`;
  };

  // ---- Bantuan mengidentifikasi nilai
  R.bantuan = (p) => {
    const s = p.sheet;
    let html = '';
    for (let r = 2; r <= 40; r++) {
      const t = raw(s, 'A' + r);
      if (!t) continue;
      if (t === t.toUpperCase() && t.length < 40) html += `<h2>${esc(t)}</h2>`;
      else html += `<p>${esc(t)}</p>`;
    }
    return `<div class="page"><div class="toolbar">${shapeButtons(s)}</div><div class="card"><h1 class="sheet-title">${esc(raw(s, 'A1'))}</h1><div class="card-body prose">${html}</div></div></div>`;
  };

  // ---- Kelola Data & Ekspor
  R.data = () => {
    const list = listSaved();
    const msgs = model.allMessages();
    const prog = model.progress();
    const items = list.map((x) => `<tr><td>${esc(x.name)}${x.id === model.state.id ? ' <span class="badge ok">sedang dibuka</span>' : ''}</td><td>${esc(x.year || '')}</td>
      <td class="num">${x.total != null ? esc(E.formatValue(x.total)) + ' / ' + esc(E.formatValue(x.max)) : ''}</td><td>${esc(new Date(x.updated).toLocaleString('id-ID'))}</td>
      <td style="white-space:nowrap">${x.id === model.state.id ? '' : `<button class="btn btn-sm" data-open-saved="${esc(x.id)}">Buka</button> `}<button class="btn btn-sm btn-danger" data-del-saved="${esc(x.id)}">Hapus</button></td></tr>`).join('');
    return `<div class="page">
      <h1 class="page-h">Kelola Data &amp; Ekspor</h1>
      <p class="page-sub">Data tidak dikirim ke server mana pun. Penilaian tersimpan di browser ini dan dapat disimpan sebagai file.</p>
      <div class="card"><div class="bar-title">Penilaian yang sedang dibuka</div><div class="card-body">
        <p><strong>${esc($('#currentName').textContent)}</strong><br><span class="small muted">Progres: ${prog.done} dari ${prog.total} butir lengkap${prog.warn ? `, ${prog.warn} perlu dilengkapi` : ''}.</span></p>
        <div class="toolbar">
          <button class="btn btn-primary" data-act="save-file">Simpan ke file (.json)</button>
          <button class="btn" data-act="open-file">Buka file (.json)</button>
          <button class="btn" data-act="new">Penilaian baru</button>
          <button class="btn" data-act="duplicate">Duplikat sebagai penilaian baru</button>
        </div></div></div>
      <div class="card"><div class="bar-title">Finalisasi</div><div class="card-body">
        ${msgs.length ? `<div class="alert alert-warn"><strong>${msgs.length} isian wajib belum lengkap</strong> (aturan sama dengan pesan pada file Excel):</div>
          <ul class="small">${msgs.slice(0, 60).map((m) => `<li><a href="${hrefSheet(m.q.sheet)}">${esc(m.q.title)}</a> — ${esc(m.msg)}</li>`).join('')}</ul>`
        : '<div class="alert alert-ok">Semua isian wajib pada butir yang dijawab sudah lengkap.</div>'}
        <p class="small muted">${prog.total - prog.done - prog.warn} butir belum dijawab.</p>
        ${model.state.finalized ? `<div class="alert alert-ok">Penilaian ditandai FINAL pada ${esc(new Date(model.state.finalizedAt || model.state.updated).toLocaleString('id-ID'))}.</div><button class="btn" data-act="unfinal">Batalkan status final</button>`
        : `<button class="btn btn-teal" data-act="final"${msgs.length || prog.done < prog.total ? ' disabled title="Lengkapi semua butir terlebih dahulu"' : ''}>Tandai sebagai FINAL</button>`}
      </div></div>
      <div class="card"><div class="bar-title">Ekspor (pengganti makro “Eksport Data …”)</div><div class="card-body">
        <p class="small">Berisi tiga tabel yang sama dengan sheet tersembunyi pada Excel: <span class="kbd">PA attributes</span>, <span class="kbd">Threat data</span>, <span class="kbd">Assessment data</span> — baris judul + satu baris nilai, dengan kolom yang sama seperti yang disalin makro ke <span class="kbd">METT4_exported_data.xlsm</span>.</p>
        <div class="toolbar"><button class="btn btn-primary" data-act="export-xlsx">Unduh XLSX (3 sheet)</button>
          <button class="btn" data-act="export-csv" data-i="0">CSV: PA attributes</button><button class="btn" data-act="export-csv" data-i="1">CSV: Threat data</button><button class="btn" data-act="export-csv" data-i="2">CSV: Assessment data</button></div>
      </div></div>
      <div class="card"><div class="bar-title">Penilaian yang tersimpan di browser ini</div>
        ${items ? `<div class="tbl-wrap"><table class="grid compact"><thead><tr><th>Kawasan</th><th>Tahun</th><th class="num">Skor</th><th>Terakhir diubah</th><th></th></tr></thead><tbody>${items}</tbody></table></div>` : '<div class="card-body muted">Belum ada.</div>'}
      </div></div>`;
  };

  // ---- Rekap
  let rekapItems = [];
  R.rekap = () => {
    const rows = rekapItems.map((it, i) => `<tr><td>${esc(it.sum.name)}</td><td>${esc(E.toText(it.sum.year))}</td><td>${esc(it.sum.authority)}</td>
      <td class="num"><strong>${esc(E.formatValue(it.sum.total))}</strong> / ${esc(E.formatValue(it.sum.max))}</td>
      ${it.sum.elements.map((e) => `<td class="num">${e.pct == null ? '' : esc(E.formatValue(e.pct, '0.00%'))}</td>`).join('')}
      <td class="small muted">${esc(it.file)}</td><td><button class="btn btn-sm btn-danger" data-rekap-del="${i}">×</button></td></tr>`).join('');
    const els = ['Planning', 'Inputs', 'Process', 'Outputs', 'Outcomes', 'Total'];
    return `<div class="page wide"><h1 class="page-h">Rekap Multi-Kawasan</h1>
      <p class="page-sub">Gabungkan banyak file penilaian (.json dari aplikasi ini) untuk rekap tingkat UPT/Direktorat. Pengganti file <span class="kbd">METT4_exported_data.xlsm</span>: tiap penilaian menjadi satu baris pada setiap tabel ekspor.</p>
      <div class="toolbar"><label class="btn btn-primary">Tambah file penilaian…<input type="file" id="rekapFiles" accept=".json,application/json" multiple hidden></label>
        <button class="btn" data-act="rekap-local">Tambahkan semua yang tersimpan di browser</button>
        <span class="spacer"></span>
        <button class="btn" data-act="rekap-xlsx"${rekapItems.length ? '' : ' disabled'}>Unduh XLSX gabungan</button>
        <button class="btn" data-act="rekap-clear"${rekapItems.length ? '' : ' disabled'}>Kosongkan</button></div>
      <div class="card"><div class="tbl-wrap"><table class="grid compact"><thead><tr><th>Kawasan</th><th>Tahun</th><th>Otoritas</th><th class="num">Skor</th>${els.map((e) => `<th class="num">${e} %</th>`).join('')}<th>File</th><th></th></tr></thead>
      <tbody>${rows || `<tr><td colspan="13" class="muted">Belum ada file. Klik “Tambah file penilaian…”.</td></tr>`}</tbody></table></div></div></div>`;
  };
  function addRekap(state, file) {
    const m = new M.Model(WB);
    m.load(state);
    rekapItems.push({ file, state, sum: m.summary(), tables: m.exportTables() });
  }

  // ---- Info
  R.info = () => `<div class="page"><div class="card"><h1 class="sheet-title">Tentang aplikasi ini</h1><div class="card-body prose">
    <p>Aplikasi ini adalah versi web dari file <span class="kbd">${esc(WB.meta.source)}</span> (METT-4 KSDAE v4.4). File Excel tersebut adalah acuan utama: teks pertanyaan, pilihan jawaban, panduan, daftar dropdown, rumus, dan aturan skor diambil langsung darinya oleh skrip <span class="kbd">tools/extract.py</span> (ekstraksi ${esc(WB.meta.extractedAt)}).</p>
    <h2>Bagaimana skor dihitung</h2>
    <p>Rumus Excel asli (${WB.sheets.reduce((n, s) => n + Object.values(s.cells).filter((c) => c.f).length, 0)} rumus) dievaluasi oleh mesin rumus di browser. Logika makro VBA (skor pilihan A–D = 0–3, skor tambahan = jumlah centang, “tidak relevan” = skor & maks 0, isian wajib) dijalankan ulang sesuai kode makro.</p>
    <h2>Perbedaan yang tidak terhindarkan</h2>
    <p>• Di Excel skor baru tertulis saat Anda pindah sheet; di sini skor langsung diperbarui.\n• Tombol makro ekspor diganti tombol unduh XLSX/CSV.\n• Tata letak cetak dibuat semirip mungkin, bukan identik piksel.\n• Nama hari/bulan pada Berita Acara memakai bahasa Indonesia (seperti Excel berlokal Indonesia); angka memakai koma desimal.</p>
    <h2>Temuan pada file asli</h2>
    <p>Beberapa rumus/makro pada file asli mengandung kejanggalan. Semuanya <strong>dipertahankan apa adanya</strong> kecuali satu kesalahan makro yang membuat skor 7a-c tidak pernah terhitung. Daftar lengkap: <a href="docs/TEMUAN.md" target="_blank" rel="noopener">docs/TEMUAN.md</a>.</p>
    <h2>Privasi</h2><p>Tidak ada data yang dikirim ke server. Semua isian tersimpan di browser (localStorage) dan di file yang Anda simpan sendiri.</p>
  </div></div></div>`;

  // ================================================================ router & render
  let currentPage = null;
  function findPage() {
    const h = decodeURIComponent(location.hash.replace(/^#\//, ''));
    return pages.find((p) => decodeURIComponent(p.id) === h) || appPages.find((p) => p.id === h) || pages[0];
  }
  function render() {
    const p = findPage();
    const changed = currentPage !== p;
    currentPage = p;
    model.onceChange = null;
    const main = $('#main');
    const scroll = changed ? 0 : window.scrollY;
    main.innerHTML = (R[p.type] || (() => '<div class="page">Halaman tidak ditemukan.</div>'))(p);
    document.title = p.title + ' — METT-4 KSDAE';
    bindFields(main);
    if (p.type === 'question') bindQuestion(p, main);
    $$('[data-chart]', main).forEach(drawChart);
    refreshLevels(main);
    $$('[data-ctl]', main).forEach((el) => el.addEventListener('change', () => { model.setControl(p.sheet, el.dataset.ctl, el.checked); model.afterChange({}); }));
    $$('[data-edit]', main).forEach((el) => el.addEventListener('blur', () => { const v = el.textContent.trim(); model.set(el.dataset.edit, v === '' ? null : v); }));
    $$('[data-toggle-threat]', main).forEach((b) => b.addEventListener('click', () => { const t = b.closest('.threat'); t.classList.toggle('collapsed'); b.textContent = t.classList.contains('collapsed') ? 'Isi' : 'Tutup'; }));
    document.getElementById('pageSize') || document.head.insertAdjacentHTML('beforeend', '<style id="pageSize"></style>');
    $('#pageSize').textContent = p.type === 'rencana' ? '@media print { @page { size: A4 landscape; margin: 10mm; } }' : '';
    renderNav();
    updateHeader();
    window.scrollTo(0, scroll);
    if (changed) { document.body.classList.remove('nav-open'); $('#navToggle').setAttribute('aria-expanded', 'false'); }
  }

  // ================================================================ aksi global
  function saveFile() {
    download(model.fileName() + '.json', model.toFile(), 'application/json');
    toast('File disimpan: ' + model.fileName() + '.json');
  }
  function openFile() { $('#fileInput').click(); }
  function startNew(fromState) {
    persistNow();
    const st = M.newAssessment();
    if (fromState) { st.inputs = JSON.parse(JSON.stringify(fromState.inputs)); st.controls = JSON.parse(JSON.stringify(fromState.controls)); }
    model.load(st);
    persistNow();
  }
  function exportXlsx() {
    const t = model.exportTables();
    const bytes = X.writeXlsx(t.map((x) => ({ name: x.name, rows: [x.header, x.values.map(cellOut)] })));
    download(model.fileName() + '_exported_data.xlsx', new Blob([bytes], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
  }
  const cellOut = (v) => (v instanceof E.XlError ? { error: v.error } : v);

  document.addEventListener('click', async (ev) => {
    const b = ev.target.closest('[data-act],[data-open-saved],[data-del-saved],[data-rekap-del]');
    if (!b) return;
    if (b.dataset.openSaved) { persistNow(); if (loadSaved(b.dataset.openSaved)) { toast('Penilaian dibuka'); render(); } return; }
    if (b.dataset.delSaved) {
      if (!confirm('Hapus penilaian ini dari browser? (File yang sudah disimpan tidak terpengaruh)')) return;
      const id = b.dataset.delSaved;
      LS.del(LS_A + id); LS.set(LS_LIST, JSON.stringify(listSaved().filter((x) => x.id !== id)));
      if (id === model.state.id) startNew();
      render(); return;
    }
    if (b.dataset.rekapDel) { rekapItems.splice(+b.dataset.rekapDel, 1); render(); return; }
    switch (b.dataset.act) {
      case 'save-file': saveFile(); break;
      case 'open-file': openFile(); break;
      case 'new': if (confirm('Mulai penilaian baru? Penilaian saat ini tetap tersimpan di browser.')) { startNew(); location.hash = '#/s/Tentang'; render(); toast('Penilaian baru dibuat'); } break;
      case 'duplicate': startNew(model.state); render(); toast('Salinan dibuat sebagai penilaian baru'); break;
      case 'final': model.state.finalized = true; model.state.finalizedAt = new Date().toISOString(); model.emit(); render(); break;
      case 'unfinal': model.state.finalized = false; model.emit(); render(); break;
      case 'export-xlsx': exportXlsx(); break;
      case 'export-csv': { const t = model.exportTables()[+b.dataset.i]; download(model.fileName() + '_' + t.name.replace(/\s+/g, '_') + '.csv', X.toCsv([t.header, t.values.map(cellOut)]), 'text/csv;charset=utf-8'); break; }
      case 'print': window.print(); break;
      case 'expand-all': $$('.threat').forEach((t) => { t.classList.remove('collapsed'); const x = $('[data-toggle-threat]', t); if (x) x.textContent = 'Tutup'; }); break;
      case 'collapse-all': $$('.threat').forEach((t) => { const filled = $$('input,textarea,select', t).some((i) => i.value); if (!filled) { t.classList.add('collapsed'); const x = $('[data-toggle-threat]', t); if (x) x.textContent = 'Isi'; } }); break;
      case 'import-prev': {
        const inp = document.createElement('input'); inp.type = 'file'; inp.accept = '.json,application/json';
        inp.onchange = async () => {
          try {
            const st = M.Model.parseFile(await inp.files[0].text());
            const m = new M.Model(WB); m.load(st);
            let n = 0;
            for (let r = 6; r <= 50; r++) {
              const c = sheetOf(DAFTAR).cells['E' + r];
              if (!c || !c.f) continue;
              const v = m.engine.get(DAFTAR, 'E' + r);
              if (typeof v === 'number') { model.set(key(DAFTAR, 'F' + r), v, { silent: true }); n++; }
            }
            model.afterChange({}); render(); toast(n + ' skor sebelumnya diisi dari ' + inp.files[0].name);
          } catch (e) { alert('Gagal membaca file: ' + e.message); }
        };
        inp.click(); break;
      }
      case 'rekap-local': {
        for (const x of listSaved()) { const t = LS.get(LS_A + x.id); if (t) { try { addRekap(M.Model.parseFile(t), '(browser) ' + x.name); } catch (e) { /* lewati */ } } }
        render(); break;
      }
      case 'rekap-clear': rekapItems = []; render(); break;
      case 'rekap-xlsx': {
        const names = rekapItems[0].tables.map((t) => t.name);
        const sheets = names.map((nm, i) => ({ name: nm, rows: [rekapItems[0].tables[i].header].concat(rekapItems.map((it) => it.tables[i].values.map(cellOut))) }));
        const sumRows = [['Kawasan', 'Tahun', 'Otoritas', 'Skor', 'Maks', 'Planning %', 'Inputs %', 'Process %', 'Outputs %', 'Outcomes %', 'Total %', 'File']]
          .concat(rekapItems.map((it) => [it.sum.name, it.sum.year, it.sum.authority, it.sum.total, it.sum.max].concat(it.sum.elements.map((e) => e.pct), [it.file])));
        sheets.unshift({ name: 'Ringkasan', rows: sumRows });
        download('METT4_exported_data_rekap.xlsx', new Blob([X.writeXlsx(sheets)], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
        break;
      }
      default: break;
    }
  });
  document.addEventListener('change', async (ev) => {
    if (ev.target.id === 'rekapFiles') {
      for (const f of ev.target.files) {
        try { addRekap(M.Model.parseFile(await f.text()), f.name); } catch (e) { alert(f.name + ': ' + e.message); }
      }
      render();
    }
  });
  $('#fileInput').addEventListener('change', async (ev) => {
    const f = ev.target.files[0];
    ev.target.value = '';
    if (!f) return;
    try {
      const st = M.Model.parseFile(await f.text());
      persistNow();
      model.load(st);
      persistNow();
      render();
      toast('File dibuka: ' + f.name);
    } catch (e) { alert('Gagal membuka file: ' + e.message); }
  });
  $('#btnSave').addEventListener('click', saveFile);
  $('#btnOpen').addEventListener('click', openFile);
  $('#navToggle').addEventListener('click', () => {
    const open = document.body.classList.toggle('nav-open');
    $('#navToggle').setAttribute('aria-expanded', String(open));
  });
  window.addEventListener('hashchange', render);
  window.addEventListener('beforeunload', () => { clearTimeout(saveTimer); persistNow(); });

  model.onChange(() => {
    persistSoon();
    refreshComputed();
    if (typeof model.onceChange === 'function') model.onceChange();
  });

  // ================================================================ mulai
  const curId = LS.get(LS_CUR);
  if (!(curId && loadSaved(curId))) model.load(M.newAssessment());
  render();
  if ('serviceWorker' in navigator && location.protocol.startsWith('http')) {
    navigator.serviceWorker.register('sw.js').catch(() => { /* offline tidak tersedia */ });
  }
  window.METT_APP = { model, render }; // untuk pengujian/diagnosa
})();
