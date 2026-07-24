/* ================================================================
   SERVICE WORKER — REPORTES C.C.R.M
   Ubicación física: /CONFIG APP MOVIL/sw.js
   Scope registrado: '/'  (controla toda la aplicación)
   Estrategia: Stale-While-Revalidate
================================================================ */

const CACHE_NAME = 'metricas-v22';

/* Archivos pre-cacheados en la instalación.
   Las rutas son absolutas desde la raíz del sitio porque el scope es '/'. */
const PRECACHE_URLS = [
  '/index.html',
  '/style.css',
  '/app.js',
  '/CONFIG APP MOVIL/manifest.json',
  '/CONFIG APP MOVIL/icon-192.png',
  '/CONFIG APP MOVIL/icon-512.png'
];

/* ── INSTALL ── */
self.addEventListener('install', event => {
  event.waitUntil(
    caches.open(CACHE_NAME).then(cache => cache.addAll(PRECACHE_URLS))
  );
  self.skipWaiting();
});

/* ── ACTIVATE: elimina cachés de versiones anteriores ── */
self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys().then(keys =>
      Promise.all(
        keys.filter(key => key !== CACHE_NAME).map(key => caches.delete(key))
      )
    )
  );
  self.clients.claim();
});

/* ── FETCH: Stale-While-Revalidate ──
   Responde con caché si existe (rápido) y en paralelo actualiza el caché.
   APIs externas (GitHub, Fonts, CDNs) se sirven directo desde la red.
*/
self.addEventListener('fetch', event => {
  const url = new URL(event.request.url);

  const isExternal =
    url.hostname.includes('api.github.com')       ||
    url.hostname.includes('fonts.googleapis.com') ||
    url.hostname.includes('fonts.gstatic.com')    ||
    url.hostname.includes('cdnjs.cloudflare.com') ||
    url.hostname.includes('sheets.googleapis.com');

  if (isExternal || event.request.method !== 'GET') {
    event.respondWith(fetch(event.request));
    return;
  }

  event.respondWith(
    caches.open(CACHE_NAME).then(async cache => {
      const cached = await cache.match(event.request);

      const networkFetch = fetch(event.request)
        .then(res => {
          if (res && res.status === 200) {
            cache.put(event.request, res.clone());
          }
          return res;
        })
        .catch(() => undefined);

      return cached || networkFetch;
    })
  );
});
