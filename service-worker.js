const CACHE_NAME = "os-fogoes-v1";
const APP_SHELL = [
  "./",
  "./index.html",
  "./manifest.json",
  "./icon-192.png",
  "./icon-512.png"
];

self.addEventListener('install', (event) => {
  self.skipWaiting();
  event.waitUntil(
    caches.open(CACHE_NAME).then(cache => cache.addAll(APP_SHELL))
  );
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then(keys =>
      Promise.all(keys.filter(k => k !== CACHE_NAME).map(k => caches.delete(k)))
    )
  );
  self.clients.claim();
});

// App shell: cache-first. Everything else (API calls to Apps Script): network-first, no caching of data.
self.addEventListener('fetch', (event) => {
  const url = new URL(event.request.url);

  if (APP_SHELL.some(p => url.pathname.endsWith(p.replace('./', '')) ) || url.pathname === '/' ) {
    event.respondWith(
      caches.match(event.request).then(cached => cached || fetch(event.request))
    );
    return;
  }

  // Let API/network requests pass straight through (handled by app's own offline queue)
});
