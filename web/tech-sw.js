// IRONLOG technician portal service worker.
// Keeps only the portal's own page, scripts and styles so it opens without
// signal. API data is never cached here: the portal keeps its own last-seen
// copy and its unsent updates in the phone's storage.
const CACHE = "ironlog-tech-shell-v7";
const SHELL = [
  "./technician-terminal.html",
  "./technician-terminal.css?v=3",
  "./tech-portal.css?v=6",
  "./auth-shared.js?v=5",
  "./tech-portal-pt.js?v=5",
  "./tech-portal.js?v=6",
  "./technician-terminal.js?v=7",
  "./tech-manifest.json",
];

self.addEventListener("install", (event) => {
  event.waitUntil(caches.open(CACHE).then((c) => c.addAll(SHELL)).then(() => self.skipWaiting()));
});

self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches.keys()
      .then((keys) => Promise.all(keys.filter((k) => k.startsWith("ironlog-tech-shell-") && k !== CACHE).map((k) => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

function isShell(url) {
  if (url.origin !== self.location.origin || url.pathname.includes("/api/")) return false;
  const file = url.pathname.split("/").pop();
  return SHELL.some((s) => s.replace("./", "").split("?")[0] === file);
}

self.addEventListener("fetch", (event) => {
  const req = event.request;
  if (req.method !== "GET") return;
  const url = new URL(req.url);
  if (!isShell(url)) return;
  // Network first so updates arrive at once; the cached copy only when offline.
  event.respondWith(
    fetch(req)
      .then((res) => {
        if (res.ok) {
          const copy = res.clone();
          caches.open(CACHE).then((c) => c.put(req, copy));
        }
        return res;
      })
      .catch(() => caches.match(req, { ignoreSearch: req.mode === "navigate" }))
  );
});
