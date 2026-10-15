/* METT-4 Web — service worker: aplikasi tetap bisa dibuka tanpa internet setelah kunjungan pertama. */
const CACHE = 'mett4-ksdae-v1.1.0';
const FILES = [
  './', 'index.html', 'manifest.webmanifest', 'css/app.css',
  'data/workbook.js', 'js/engine.js', 'js/model.js', 'js/charts.js', 'js/xlsx.js', 'js/sync.js', 'js/config.js', 'js/app.js',
  'assets/img/logo-kemenhut.png', 'assets/img/sampul-mett4.jpg',
  'assets/img/icon-192.png', 'assets/img/icon-512.png',
];
self.addEventListener('install', (e) => {
  e.waitUntil(caches.open(CACHE).then((c) => c.addAll(FILES)).then(() => self.skipWaiting()));
});
self.addEventListener('activate', (e) => {
  e.waitUntil(caches.keys().then((keys) => Promise.all(keys.filter((k) => k !== CACHE).map((k) => caches.delete(k)))).then(() => self.clients.claim()));
});
// network-first agar pembaruan di GitHub langsung terpakai; cache dipakai saat offline
self.addEventListener('fetch', (e) => {
  if (e.request.method !== 'GET' || new URL(e.request.url).origin !== location.origin) return;
  e.respondWith(
    fetch(e.request).then((res) => {
      const copy = res.clone();
      caches.open(CACHE).then((c) => c.put(e.request, copy)).catch(() => {});
      return res;
    }).catch(() => caches.match(e.request, { ignoreSearch: true }).then((r) => r || caches.match('index.html')))
  );
});
