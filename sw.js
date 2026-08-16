/* Cash Tray Manager service worker.
 * Strategy:
 *  - App shell (index.html, app.js, manifest, icons): pre-cached on install.
 *  - Same-origin assets: stale-while-revalidate.
 *  - CDN assets (fonts, pdf.js, jspdf, html2canvas, chart.js, font-awesome):
 *    cache-first with background update, so the app works fully offline after
 *    first use.
 *  - Google Apps Script API: network only (never cached, POST and GET pass through).
 */

const CACHE = 'ctm-v9.9';
const APP_SHELL = [
  './',
  './index.html',
  './app.js',
  './manifest.webmanifest',
  './icons/icon-192.png',
  './icons/icon-512.png'
];

const API_HOSTS = ['script.google.com'];

function isApiRequest(url) {
  return API_HOSTS.some(h => url.hostname === h || url.hostname.endsWith('.' + h));
}

self.addEventListener('install', (e) => {
  e.waitUntil(
    caches.open(CACHE)
      .then((c) => c.addAll(APP_SHELL))
      .then(() => self.skipWaiting())
  );
});

self.addEventListener('activate', (e) => {
  e.waitUntil(
    caches.keys()
      .then((keys) => Promise.all(keys.filter((k) => k !== CACHE).map((k) => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

async function staleWhileRevalidate(request) {
  const cached = await caches.match(request);
  const fetchPromise = fetch(request)
    .then((res) => {
      if (res && res.ok) {
        const clone = res.clone();
        caches.open(CACHE).then((c) => c.put(request, clone));
      }
      return res;
    })
    .catch(() => cached);
  return cached || fetchPromise;
}

self.addEventListener('fetch', (e) => {
  const url = new URL(e.request.url);

  // Apps Script API: network only.
  if (isApiRequest(url)) return;

  // Navigations: network first, fall back to cached shell when offline.
  if (e.request.mode === 'navigate') {
    e.respondWith(
      fetch(e.request)
        .then((res) => {
          const clone = res.clone();
          caches.open(CACHE).then((c) => c.put(e.request, clone));
          return res;
        })
        .catch(() => caches.match('./index.html'))
    );
    return;
  }

  // Everything else: stale-while-revalidate (caches CDN assets after first use).
  e.respondWith(staleWhileRevalidate(e.request));
});
