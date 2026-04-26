// ==========================================
// Service Worker – Learning App
// ==========================================

const APP_CACHE = "learning-app-shell-v2";
const CSV_CACHE = "learning-app-csv-v1";

// ?? Statische Dateien (App-Shell)
const STATIC_ASSETS = [
  "index.html",
  "assets/css/style.css",
  "assets/js/app.js",
  "manifest.json",
  "icons/icon-192.png",
  "icons/icon-512.png"
];

// ==========================================
// INSTALL
// ==========================================
self.addEventListener("install", event => {
  self.skipWaiting();
  event.waitUntil(
    caches.open(APP_CACHE).then(cache => cache.addAll(STATIC_ASSETS))
  );
});

// ==========================================
// ACTIVATE
// ==========================================
self.addEventListener("activate", event => {
  event.waitUntil(
    caches.keys().then(keys =>
      Promise.all(
        keys.map(key => {
          if (![APP_CACHE, CSV_CACHE].includes(key)) {
            return caches.delete(key);
          }
        })
      )
    )
  );
  self.clients.claim();
});

// ==========================================
// FETCH
// ==========================================
self.addEventListener("fetch", event => {
  const req = event.request;
  const url = req.url;

  if (req.method !== "GET") return;

  // ? CSV: Network First (Content aktualisieren)
  if (url.endsWith(".csv")) {
    event.respondWith(networkFirstCSV(req));
    return;
  }

  // ? App-Shell: Cache First (stabil)
  event.respondWith(
    caches.match(req).then(cached => {
      if (cached) return cached;
      return fetch(req).then(response => {
        return caches.open(APP_CACHE).then(cache => {
          cache.put(req, response.clone());
          return response;
        });
      });
    })
  );
});

// ==========================================
// STRATEGIEN
// ==========================================
async function networkFirstCSV(request) {
  const cache = await caches.open(CSV_CACHE);
  try {
    const response = await fetch(request);
    cache.put(request, response.clone());
    return response;
  } catch {
    const cached = await cache.match(request);
    if (cached) return cached;
    throw new Error("CSV offline nicht verfügbar");
  }
}