// Service worker de PPT Anim. Fichier généré par scripts/build.mjs : ne pas modifier dist/sw.js à la main.
// Les chemins sont relatifs à la portée d'enregistrement (/pptanim/ sur GitHub Pages).

const VERSION = '__VERSION__';
const CACHE = `pptanim-${VERSION}`;
const ASSETS = __ASSETS__;
const KNOWN = new Set(ASSETS);
const scope = () => new URL(self.registration.scope);

self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE).then((cache) =>
      cache.addAll(ASSETS.map((path) => new Request(new URL(path, scope()), { cache: 'reload' })))),
  );
});

self.addEventListener('activate', (event) => {
  event.waitUntil((async () => {
    for (const key of await caches.keys()) {
      if (key.startsWith('pptanim-') && key !== CACHE) await caches.delete(key);
    }
    await self.clients.claim();
  })());
});

// La page demande l'activation immédiate quand l'utilisateur accepte la mise à jour.
self.addEventListener('message', (event) => {
  if (event.data === 'skipWaiting') self.skipWaiting();
});

self.addEventListener('fetch', (event) => {
  const req = event.request;
  if (req.method !== 'GET') return;
  const url = new URL(req.url);
  const base = scope();
  if (url.origin !== base.origin || !url.pathname.startsWith(base.pathname)) return;
  const path = url.pathname.slice(base.pathname.length) || 'index.html';
  if (!KNOWN.has(path)) return; // tout le reste (dont la page du complément) passe par le réseau
  event.respondWith((async () => {
    const cache = await caches.open(CACHE);
    return (await cache.match(new URL(path, base))) || fetch(req);
  })());
});
