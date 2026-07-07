const CACHE_NAME = 'jan-scanner-v13';
const ASSETS = [
  './',
  './index.html',
  './app.js',
  './manifest.json',
  './lib/xlsx.full.min.js',
  './lib/html5-qrcode.min.js',
  './icons/icon-192.png',
  './icons/icon-512.png',
  './icons/apple-touch-icon.png',
];

self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => cache.addAll(ASSETS))
  );
  self.skipWaiting();
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(keys.filter((k) => k !== CACHE_NAME).map((k) => caches.delete(k)))
    )
  );
  self.clients.claim();
});

// オンライン時は常に最新を取りに行き（GitHub PagesのHTTPキャッシュも無視）、
// 取得できない時だけキャッシュを使う「ネットワーク優先」に変更。
// 以前の「キャッシュ優先」だと、更新後もずっと古い版が表示され続けてしまっていた。
self.addEventListener('fetch', (event) => {
  if (event.request.method !== 'GET') return;
  event.respondWith(
    fetch(event.request, { cache: 'no-store' })
      .then((res) => {
        const copy = res.clone();
        caches.open(CACHE_NAME).then((cache) => cache.put(event.request, copy));
        return res;
      })
      .catch(() => caches.match(event.request))
  );
});
