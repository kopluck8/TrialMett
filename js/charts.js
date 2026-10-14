/*
 * METT-4 Web — Grafik SVG
 * Menggambar ulang grafik Excel (radar, batang vertikal/horizontal, donat) dari
 * definisi grafik yang diekstrak dari file asli (jenis, seri, rentang data, warna).
 * Data diambil dari sel yang sama dengan grafik Excel melalui mesin rumus.
 */
(function (root) {
  'use strict';
  const E = root.MettEngine;
  const THEME = { accent1: '#4472C4', accent2: '#ED7D31', accent3: '#A5A5A5', accent4: '#FFC000', accent5: '#5B9BD5', accent6: '#70AD47' };
  const color = (c, fallback) => {
    if (!c) return fallback || '#888';
    if (c.startsWith('scheme:')) return THEME[c.slice(7)] || fallback || '#888';
    return '#' + c;
  };
  const esc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

  function readRange(engine, ref) {
    if (!ref) return [];
    const m = /^'?(.*?)'?!\$?([A-Z]+)\$?(\d+)(?::\$?([A-Z]+)\$?(\d+))?$/.exec(ref);
    if (!m) return [];
    const sheet = m[1].replace(/''/g, "'");
    const p1 = E.parseA1(m[2] + m[3]), p2 = m[4] ? E.parseA1(m[4] + m[5]) : p1;
    const out = [];
    for (let r = p1.r; r <= p2.r; r++) for (let c = p1.c; c <= p2.c; c++) out.push(engine.get(sheet, E.toA1(c, r)));
    return out;
  }
  const num = (v) => (typeof v === 'number' && isFinite(v) ? v : 0); // galat/teks digambar sebagai 0 (seperti Excel)
  const label = (v) => (E.isBlank(v) ? '' : E.isErr(v) ? v.error : typeof v === 'number' ? E.formatValue(v) : String(v));

  function series(engine, def) {
    return def.series.map((s, i) => ({
      name: s.name || (s.nameRef ? label(readRange(engine, s.nameRef)[0]) : 'Seri ' + (i + 1)),
      cats: readRange(engine, s.catRef).map(label),
      vals: readRange(engine, s.valRef).map(num),
      raw: readRange(engine, s.valRef),
      color: color(s.color, ['#398787', '#ED7D31', '#A5A5A5'][i % 3]),
      pointColors: (s.pointColors || []).map((p) => color(p.color)),
      showVal: s.showVal, showPercent: s.showPercent,
    }));
  }
  function fmtAxis(v, nf) {
    if (nf && /%/.test(nf)) return Math.round(v * 100) + '%';
    return String(Math.round(v * 100) / 100).replace('.', ',');
  }
  function niceMax(m) {
    if (m <= 0) return 1;
    const p = Math.pow(10, Math.floor(Math.log10(m)));
    const f = m / p;
    const n = f <= 1 ? 1 : f <= 2 ? 2 : f <= 2.5 ? 2.5 : f <= 5 ? 5 : 10;
    return n * p;
  }
  function legend(items, x, y, w) {
    let out = '', cx = x;
    const widths = items.map((it) => 18 + it.name.length * 6.2 + 14);
    const total = widths.reduce((a, b) => a + b, 0);
    cx = x + Math.max(0, (w - total) / 2);
    items.forEach((it, i) => {
      out += `<rect x="${cx}" y="${y - 8}" width="10" height="10" fill="${it.color}" rx="1"/>` +
        `<text x="${cx + 14}" y="${y + 1}" class="ch-leg">${esc(it.name)}</text>`;
      cx += widths[i];
    });
    return out;
  }

  // ---------------------------------------------------------------- radar
  function radar(engine, def, W, H) {
    const ss = series(engine, def).filter((s) => s.vals.length);
    const cats = (ss.find((s) => s.cats.length) || { cats: [] }).cats;
    const n = Math.max(cats.length, ...ss.map((s) => s.vals.length));
    const cx = W / 2, cy = H / 2 - 8, R = Math.min(W, H) / 2 - 46;
    const max = niceMax(Math.max(1, ...ss.flatMap((s) => s.vals)));
    let svg = '';
    for (let k = 1; k <= 2; k++) {
      const rr = R * k / 2;
      const pts = []; for (let i = 0; i < n; i++) { const a = -Math.PI / 2 + i * 2 * Math.PI / n; pts.push((cx + rr * Math.cos(a)).toFixed(1) + ',' + (cy + rr * Math.sin(a)).toFixed(1)); }
      svg += `<polygon points="${pts.join(' ')}" class="ch-grid"/>`;
      svg += `<text x="${cx + 3}" y="${cy - rr + 4}" class="ch-tick">${fmtAxis(max * k / 2, def.valAxis && def.valAxis.numFmt)}</text>`;
    }
    for (let i = 0; i < n; i++) {
      const a = -Math.PI / 2 + i * 2 * Math.PI / n;
      svg += `<line x1="${cx}" y1="${cy}" x2="${(cx + R * Math.cos(a)).toFixed(1)}" y2="${(cy + R * Math.sin(a)).toFixed(1)}" class="ch-grid"/>`;
      const lx = cx + (R + 16) * Math.cos(a), ly = cy + (R + 16) * Math.sin(a);
      const anchor = Math.abs(Math.cos(a)) < 0.2 ? 'middle' : Math.cos(a) > 0 ? 'start' : 'end';
      svg += `<text x="${lx.toFixed(1)}" y="${(ly + 4).toFixed(1)}" text-anchor="${anchor}" class="ch-cat">${esc(cats[i] || '')}</text>`;
    }
    for (const s of ss) {
      const pts = s.vals.map((v, i) => { const a = -Math.PI / 2 + i * 2 * Math.PI / n; const rr = R * Math.max(0, v) / max; return [cx + rr * Math.cos(a), cy + rr * Math.sin(a)]; });
      svg += `<polygon points="${pts.map((p) => p[0].toFixed(1) + ',' + p[1].toFixed(1)).join(' ')}" fill="none" stroke="${s.color}" stroke-width="2.2"/>`;
      svg += pts.map((p, i) => `<circle cx="${p[0].toFixed(1)}" cy="${p[1].toFixed(1)}" r="3" fill="${s.color}"><title>${esc(s.name)} — ${esc(cats[i] || '')}: ${esc(label(s.raw[i]))}</title></circle>`).join('');
    }
    svg += legend(ss, 0, H - 10, W);
    return svg;
  }

  // ---------------------------------------------------------------- batang
  function bars(engine, def, W, H) {
    const t = def.types[0] || {};
    const horizontal = t.barDir === 'bar';
    const ss = series(engine, def);
    const cats = (ss.find((s) => s.cats.length) || { cats: [] }).cats;
    const n = Math.max(cats.length, ...ss.map((s) => s.vals.length));
    const nf = def.valAxis && def.valAxis.numFmt;
    const dataMax = Math.max(0, ...ss.flatMap((s) => s.vals));
    const max = def.valAxis && def.valAxis.max != null ? def.valAxis.max : niceMax(dataMax * 1.1 || 1);
    const ticks = 5;
    let svg = '';
    const fmtVal = (s, i) => {
      const v = s.raw[i];
      if (E.isErr(v)) return v.error;
      return nf && /%/.test(nf) ? Math.round(num(v) * 100) + '%' : label(v);
    };
    if (!horizontal) {
      const L = 34, Rm = 8, T = 14, B = 46;
      const pw = W - L - Rm, ph = H - T - B;
      for (let k = 0; k <= ticks; k++) {
        const y = T + ph - ph * k / ticks;
        svg += `<line x1="${L}" x2="${W - Rm}" y1="${y}" y2="${y}" class="ch-grid"/><text x="${L - 4}" y="${y + 3}" text-anchor="end" class="ch-tick">${fmtAxis(max * k / ticks, nf)}</text>`;
      }
      const gw = pw / n, bw = gw * 0.75 / Math.max(1, ss.length);
      for (let i = 0; i < n; i++) {
        ss.forEach((s, j) => {
          const v = Math.max(0, s.vals[i] || 0), h = ph * Math.min(v, max) / max;
          const x = L + gw * i + gw * 0.125 + bw * j, y = T + ph - h;
          svg += `<rect x="${x.toFixed(1)}" y="${y.toFixed(1)}" width="${(bw * 0.9).toFixed(1)}" height="${h.toFixed(1)}" fill="${s.color}"><title>${esc(s.name)} — ${esc(cats[i] || '')}: ${esc(fmtVal(s, i))}</title></rect>`;
          if (s.showVal) svg += `<text x="${(x + bw * 0.45).toFixed(1)}" y="${(y - 3).toFixed(1)}" text-anchor="middle" class="ch-val">${esc(fmtVal(s, i))}</text>`;
        });
        svg += `<text x="${(L + gw * i + gw / 2).toFixed(1)}" y="${T + ph + 14}" text-anchor="middle" class="ch-cat">${esc(cats[i] || '')}</text>`;
      }
      svg += legend(ss, 0, H - 8, W);
    } else {
      const L = Math.min(W * 0.45, 250), Rm = 16, T = 8, B = 40;
      const pw = W - L - Rm, ph = H - T - B;
      for (let k = 0; k <= ticks; k++) {
        const x = L + pw * k / ticks;
        svg += `<line x1="${x}" x2="${x}" y1="${T}" y2="${T + ph}" class="ch-grid"/><text x="${x}" y="${T + ph + 13}" text-anchor="middle" class="ch-tick">${fmtAxis(max * k / ticks, nf)}</text>`;
      }
      const gh = ph / n, bh = gh * 0.7 / Math.max(1, ss.length);
      for (let i = 0; i < n; i++) {
        // kategori pertama di bawah (seperti Excel)
        const gy = T + ph - gh * (i + 1);
        ss.forEach((s, j) => {
          const v = Math.max(0, s.vals[i] || 0), w = pw * Math.min(v, max) / max;
          const y = gy + gh * 0.15 + bh * (ss.length - 1 - j);
          svg += `<rect x="${L}" y="${y.toFixed(1)}" width="${w.toFixed(1)}" height="${(bh * 0.9).toFixed(1)}" fill="${s.color}"><title>${esc(s.name)} — ${esc(cats[i] || '')}: ${esc(fmtVal(s, i))}</title></rect>`;
          if (s.showVal && v > 0) svg += `<text x="${(L + w + 3).toFixed(1)}" y="${(y + bh * 0.7).toFixed(1)}" class="ch-val">${esc(fmtVal(s, i))}</text>`;
        });
        const c = cats[i] || '';
        svg += `<text x="${L - 6}" y="${(gy + gh / 2 + 3).toFixed(1)}" text-anchor="end" class="ch-cat ch-cat-sm">${esc(c.length > 48 ? c.slice(0, 46) + '…' : c)}<title>${esc(c)}</title></text>`;
      }
      svg += legend(ss.slice().reverse(), 0, H - 8, W);
    }
    return svg;
  }

  // ---------------------------------------------------------------- donat
  function doughnut(engine, def, W, H) {
    const s = series(engine, def)[0];
    if (!s) return '';
    const t = def.types[0] || {};
    const hole = (parseFloat(t.holeSize || '50') / 100);
    const total = s.vals.reduce((a, b) => a + Math.max(0, b), 0);
    const legendW = Math.min(150, W * 0.42);
    const cx = (W - legendW) / 2, cy = H / 2, R = Math.min(W - legendW, H) / 2 - 10, r = R * hole;
    let svg = '';
    const colors = s.vals.map((_, i) => s.pointColors[i] || ['#ED7D31', '#D97566', '#D8C77F', '#FFC000', '#398787'][i % 5]);
    if (total <= 0) {
      svg += `<circle cx="${cx}" cy="${cy}" r="${(R + r) / 2}" fill="none" stroke="#e3e6e8" stroke-width="${R - r}"/>`;
      svg += `<text x="${cx}" y="${cy + 4}" text-anchor="middle" class="ch-tick">tidak ada data</text>`;
    } else {
      let a0 = -Math.PI / 2;
      s.vals.forEach((v, i) => {
        if (v <= 0) return;
        const a1 = a0 + 2 * Math.PI * v / total;
        const large = a1 - a0 > Math.PI ? 1 : 0;
        const p = (rad, ang) => [cx + rad * Math.cos(ang), cy + rad * Math.sin(ang)];
        let d;
        if (v === total) {
          d = `M ${cx - R} ${cy} A ${R} ${R} 0 1 1 ${cx + R} ${cy} A ${R} ${R} 0 1 1 ${cx - R} ${cy} M ${cx - r} ${cy} A ${r} ${r} 0 1 0 ${cx + r} ${cy} A ${r} ${r} 0 1 0 ${cx - r} ${cy} Z`;
        } else {
          const [x1, y1] = p(R, a0), [x2, y2] = p(R, a1), [x3, y3] = p(r, a1), [x4, y4] = p(r, a0);
          d = `M ${x1} ${y1} A ${R} ${R} 0 ${large} 1 ${x2} ${y2} L ${x3} ${y3} A ${r} ${r} 0 ${large} 0 ${x4} ${y4} Z`;
        }
        const pct = Math.round(v / total * 100) + '%';
        svg += `<path d="${d}" fill="${colors[i]}" fill-rule="evenodd" stroke="#fff" stroke-width="1"><title>${esc(s.cats[i] || '')}: ${v} (${pct})</title></path>`;
        if (s.showPercent) { const [lx, ly] = p((R + r) / 2, (a0 + a1) / 2); svg += `<text x="${lx.toFixed(1)}" y="${(ly + 4).toFixed(1)}" text-anchor="middle" class="ch-pct">${pct}</text>`; }
        a0 = a1;
      });
    }
    const lx = W - legendW + 4;
    s.cats.forEach((c, i) => {
      const y = cy - (s.cats.length * 16) / 2 + i * 16 + 10;
      svg += `<rect x="${lx}" y="${y - 8}" width="10" height="10" fill="${colors[i]}"/><text x="${lx + 14}" y="${y + 1}" class="ch-leg">${esc(c)}</text>`;
    });
    return svg;
  }

  function render(engine, def, opts) {
    opts = opts || {};
    const W = opts.width || 520, H = opts.height || 300;
    const t = (def.types[0] || {}).type;
    let inner = '';
    if (t === 'radarChart') inner = radar(engine, def, W, H);
    else if (t === 'barChart') inner = bars(engine, def, W, H);
    else if (t === 'doughnutChart' || t === 'pieChart') inner = doughnut(engine, def, W, H);
    return `<svg class="chart-svg" viewBox="0 0 ${W} ${H}" preserveAspectRatio="xMidYMid meet" role="img" aria-label="${esc(opts.title || 'grafik')}">${inner}</svg>`;
  }

  root.MettCharts = { render, readRange };
})(typeof window !== 'undefined' ? window : globalThis);
