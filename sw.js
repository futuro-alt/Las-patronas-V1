/* Las Patronas · service worker
   Para publicar una versión nueva: cambia VERSION aquí y VERSION_APP en app.js. */
const VERSION = '1.0.0';
const CACHE = 'las-patronas-' + VERSION;
const ARCHIVOS = [
  "./",
  "app.js",
  "fonts/cabin-condensed-latin-500-normal.woff2",
  "fonts/cabin-condensed-latin-600-normal.woff2",
  "fonts/cabin-condensed-latin-700-normal.woff2",
  "fonts/cabin-latin-400-normal.woff2",
  "fonts/cabin-latin-500-normal.woff2",
  "fonts/cabin-latin-600-normal.woff2",
  "fonts/cabin-latin-700-normal.woff2",
  "fonts/norican-latin-400-normal.woff2",
  "icons/apple-touch-icon.png",
  "icons/favicon-32.png",
  "icons/icon-192.png",
  "icons/icon-512.png",
  "icons/icon-maskable-512.png",
  "img/gallina-blanca.jpg",
  "img/gallina-negra.jpg",
  "img/gallina-roja.jpg",
  "img/gallinas.jpg",
  "index.html",
  "manifest.webmanifest",
  "styles.css",
  "vendor/xlsx.mini.min.js"
];

self.addEventListener('install', (e) => {
  e.waitUntil((async () => {
    const cache = await caches.open(CACHE);
    await Promise.all(ARCHIVOS.map(async (f) => {
      const r = await fetch(new Request(f + '?v=' + VERSION, { cache: 'reload' }));
      if (!r.ok) throw new Error('No se pudo descargar ' + f);
      const limpia = r.redirected ? new Response(await r.blob(), { status: r.status, statusText: r.statusText, headers: r.headers }) : r;
      await cache.put(new Request(f), limpia);
    }));
  })());
});

self.addEventListener('activate', (e) => {
  e.waitUntil((async () => {
    for (const k of await caches.keys()) {
      if (k.startsWith('las-patronas-') && k !== CACHE) await caches.delete(k);
    }
    await self.clients.claim();
  })());
});

self.addEventListener('message', (e) => {
  if (e.data === 'SKIP_WAITING') self.skipWaiting();
});

self.addEventListener('fetch', (e) => {
  const req = e.request;
  if (req.method !== 'GET') return;
  const url = new URL(req.url);
  if (url.origin !== self.location.origin) return;
  if (req.mode === 'navigate') {
    e.respondWith((async () => {
      const cache = await caches.open(CACHE);
      return (await cache.match('./')) || (await cache.match('index.html')) || fetch(req);
    })());
    return;
  }
  e.respondWith((async () => {
    const cache = await caches.open(CACHE);
    return (await cache.match(req, { ignoreSearch: true })) || fetch(req);
  })());
});
