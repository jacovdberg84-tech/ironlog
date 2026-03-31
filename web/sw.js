const CACHE_NAME = "ironlog-v1";

const STATIC_ASSETS = [
  "/web/index.html",
  "/web/styles.css",
  "/web/app.js",
  "/web/offline.html"
];

self.addEventListener("install", event => {
  event.waitUntil(
    caches.open(CACHE_NAME).then(cache => cache.addAll(STATIC_ASSETS))
  );
  self.skipWaiting();
});

self.addEventListener("activate", event => {
  event.waitUntil(self.clients.claim());
});

self.addEventListener("fetch", event => {
  const { request } = event;

  // Only handle GET requests
  if (request.method !== "GET") return;

  event.respondWith(
    fetch(request)
      .then(response => {
        const copy = response.clone();
        caches.open(CACHE_NAME).then(cache => cache.put(request, copy));
        return response;
      })
      .catch(() =>
        caches.match(request).then(res => res || caches.match("/web/offline.html"))
      )
  );
});