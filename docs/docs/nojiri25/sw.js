var LEGACY_CACHE_NAME = "kaseki25-legacy-redirect-v6";
var LEGACY_INDEX = "./index.html";

self.addEventListener("install", function (event) {
  event.waitUntil(
    caches.open(LEGACY_CACHE_NAME).then(function (cache) {
      return cache.add(LEGACY_INDEX);
    })
  );
  self.skipWaiting();
});

self.addEventListener("activate", function (event) {
  event.waitUntil(
    caches.keys().then(function (keys) {
      return Promise.all(keys.map(function (key) {
        if (/^kaseki25-pwa-v[1-8]$/.test(key) || /^kaseki25-legacy-redirect-v[1-5]$/.test(key)) {
          return caches.delete(key);
        }
        return null;
      }));
    })
  );
  self.clients.claim();
});

self.addEventListener("fetch", function (event) {
  if (event.request.method !== "GET") {
    return;
  }
  if (event.request.mode === "navigate") {
    event.respondWith(
      fetch(LEGACY_INDEX, { cache: "no-store" }).catch(function () {
        return caches.match(LEGACY_INDEX);
      })
    );
  }
});
