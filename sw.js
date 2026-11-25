// Offline cache: the app must open at the takeoff field without reception.
const CACHE = 'ppg-logbook-v2';
const FILES = ['./', 'index.html', 'app.js', 'drive.js', 'styles.css', 'manifest.json', 'icon.svg', 'icon-192.png', 'icon-512.png', 'vendor/exceljs.min.js'];
self.addEventListener('install', e => e.waitUntil(caches.open(CACHE).then(c => c.addAll(FILES)).then(() => self.skipWaiting())));
self.addEventListener('activate', e => e.waitUntil(
  caches.keys().then(keys => Promise.all(keys.filter(k => k !== CACHE && k !== 'ppg-share').map(k => caches.delete(k)))).then(() => self.clients.claim())
));
// network first (so updates arrive), cache as fallback when offline
self.addEventListener('fetch', e => {
  const url = new URL(e.request.url);
  // Web Share Target: Gaggle -> Share -> this app (Android)
  if (e.request.method === 'POST' && url.pathname.endsWith('/share-target')) {
    e.respondWith((async () => {
      const files = (await e.request.formData()).getAll('track').filter(f => f && typeof f !== 'string');
      const cache = await caches.open('ppg-share');
      for (const k of await cache.keys()) await cache.delete(k);
      await Promise.all(files.map((f, i) => cache.put(`shared-${i}`, new Response(f, { headers: { 'x-name': encodeURIComponent(f.name || `track-${i}.igc`) } }))));
      return Response.redirect(new URL('./?shared=' + files.length, self.registration.scope).href, 303);
    })());
    return;
  }
  if (e.request.method !== 'GET' || new URL(e.request.url).origin !== location.origin) return;
  e.respondWith(
    fetch(e.request).then(r => { const copy = r.clone(); caches.open(CACHE).then(c => c.put(e.request, copy)); return r; })
      .catch(() => caches.match(e.request, { ignoreSearch: true }).then(r => r || caches.match('index.html')))
  );
});
