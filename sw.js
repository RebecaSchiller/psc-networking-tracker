const CACHE_NAME = 'network-momentum-v1';
const ASSETS = [
  '/psc-networking-tracker/app.html',
  '/psc-networking-tracker/favicon.png',
  '/psc-networking-tracker/favicon.svg',
  '/psc-networking-tracker/manifest.json'
];

self.addEventListener('install', event => {
  event.waitUntil(
    caches.open(CACHE_NAME).then(cache => cache.addAll(ASSETS))
  );
  self.skipWaiting();
});

self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys().then(keys =>
      Promise.all(keys.filter(k => k !== CACHE_NAME).map(k => caches.delete(k)))
    )
  );
  self.clients.claim();
});

self.addEventListener('fetch', event => {
  // Network-first for API calls, cache-first for static assets
  if (event.request.url.includes('script.google.com') || 
      event.request.url.includes('anthropic.com')) {
    event.respondWith(fetch(event.request));
    return;
  }
  event.respondWith(
    caches.match(event.request).then(cached => cached || fetch(event.request))
  );
});
