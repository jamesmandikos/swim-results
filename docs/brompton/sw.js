const CACHE = 'brompton-v6';
const STATIC = ['./icon-192.png', './icon-512.png', './manifest.json'];
const HTML_NAMES = ['brompton_home.html', 'brompton_rankings.html', 'bsc_u14_girls_rankings.html', 'cwsc_u14_girls_rankings.html', 'index.html'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(STATIC)));
  self.skipWaiting();
});

self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys().then(keys =>
      Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k)))
    )
  );
  self.clients.claim();
});

self.addEventListener('fetch', e => {
  const url = new URL(e.request.url);
  // Network-first for HTML pages — always serve fresh content
  if (HTML_NAMES.some(n => url.pathname.endsWith('/' + n) || url.pathname.endsWith('/' + n.replace('.html','')))) {
    e.respondWith(
      fetch(e.request).then(r => {
        const clone = r.clone();
        caches.open(CACHE).then(c => c.put(e.request, clone));
        return r;
      }).catch(() => caches.match(e.request))
    );
    return;
  }
  // Cache-first for static assets
  e.respondWith(caches.match(e.request).then(cached => cached || fetch(e.request)));
});
