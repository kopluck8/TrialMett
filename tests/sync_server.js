/*
 * Server uji lokal untuk mode online:
 *   - port A: menyajikan aplikasi (file statis), dengan js/config.js diarahkan ke port B
 *   - port B: menjalankan backend/Code.gs ASLI (lewat tiruan layanan Google) sebagai .../exec
 * Pemakaian:  node tests/sync_server.js 8770 8771
 */
const http = require('http');
const fs = require('fs');
const path = require('path');
const { createGas } = require('./fake_gas.js');

const ROOT = path.join(__dirname, '..');
const PA = +(process.argv[2] || 8770), PB = +(process.argv[3] || 8771);
const LATENCY = +(process.env.SYNC_LATENCY || 150);
const gas = createGas();
gas.ctx.setup();
const TYPES = { '.html': 'text/html', '.js': 'text/javascript', '.css': 'text/css', '.png': 'image/png', '.jpg': 'image/jpeg', '.json': 'application/json', '.webmanifest': 'application/manifest+json', '.md': 'text/markdown' };

http.createServer((req, res) => {
  let p = decodeURIComponent(req.url.split('?')[0]);
  if (p === '/') p = '/index.html';
  if (p === '/js/config.js') {
    res.writeHead(200, { 'Content-Type': 'text/javascript' });
    return res.end(`window.METT_CONFIG = { SYNC_URL: 'http://localhost:${PB}/exec', POLL_SECONDS: 2 };`);
  }
  const f = path.join(ROOT, p);
  if (!f.startsWith(ROOT) || !fs.existsSync(f) || fs.statSync(f).isDirectory()) { res.writeHead(404); return res.end('404'); }
  res.writeHead(200, { 'Content-Type': TYPES[path.extname(f)] || 'application/octet-stream' });
  fs.createReadStream(f).pipe(res);
}).listen(PA);

let offline = false;
http.createServer((req, res) => {
  const url = new URL(req.url, 'http://x');
  const cors = { 'Access-Control-Allow-Origin': '*', 'Content-Type': 'application/json' };
  if (url.pathname === '/_offline') { offline = url.searchParams.get('v') === '1'; res.writeHead(200, cors); return res.end('{}'); }
  if (offline) { req.destroy(); return; }
  if (url.pathname !== '/exec') { res.writeHead(404, cors); return res.end('{}'); }
  if (req.method === 'OPTIONS') { res.writeHead(405, cors); return res.end(); } // Apps Script tidak mendukung preflight
  let body = '';
  req.on('data', (c) => (body += c));
  req.on('end', () => setTimeout(() => {
    let out;
    try {
      out = req.method === 'POST' ? gas.post(body) : gas.get(Object.fromEntries(url.searchParams));
    } catch (e) { out = { ok: false, error: String(e.message) }; }
    res.writeHead(200, cors);
    res.end(JSON.stringify(out));
  }, LATENCY));
}).listen(PB);
console.log(`aplikasi: http://localhost:${PA}/   backend: http://localhost:${PB}/exec`);
