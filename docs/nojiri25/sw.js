var CACHE_NAME = "kaseki25-pwa-v1";
var ASSETS = [
  "./",
  "./index.html",
  "./styles.css",
  "./app.js",
  "./team_roster_data.js",
  "./manifest.webmanifest",
  "./icon.svg",
  "./shapes/inline_data.js",
  "./shapes/palmate_antler.png",
  "./shapes/triangle.png",
  "./shapes/incisor.png",
  "./shapes/diamond_hira.png",
  "./shapes/manifest.json",
  "./shapes/constricted_shape.png",
  "./shapes/rib_curved.png",
  "./shapes/c_shape.png",
  "./assets/large-shapes/elephant_upper_molar.png",
  "./assets/large-shapes/palmate_antler.png",
  "./assets/large-shapes/elephant_lower_molar.png",
  "./assets/large-shapes/triangle.png",
  "./assets/large-shapes/incisor.png",
  "./assets/large-shapes/constricted_bone.png",
  "./assets/large-shapes/diamond_hira.png",
  "./assets/large-shapes/diamond.png",
  "./assets/large-shapes/rib.png",
  "./assets/large-shapes/constricted_shape.png",
  "./assets/large-shapes/rib_curved.png",
  "./assets/large-shapes/c_shape.png"
];

self.addEventListener("install", function (event) {
  event.waitUntil(
    caches.open(CACHE_NAME).then(function (cache) {
      return cache.addAll(ASSETS);
    })
  );
  self.skipWaiting();
});

self.addEventListener("activate", function (event) {
  event.waitUntil(
    caches.keys().then(function (keys) {
      return Promise.all(
        keys.map(function (key) {
          if (key !== CACHE_NAME) {
            return caches.delete(key);
          }
          return null;
        })
      );
    })
  );
  self.clients.claim();
});

self.addEventListener("fetch", function (event) {
  if (event.request.method !== "GET") {
    return;
  }
  event.respondWith(
    caches.match(event.request).then(function (cached) {
      if (cached) {
        return cached;
      }
      return fetch(event.request).then(function (response) {
        var copy = response.clone();
        caches.open(CACHE_NAME).then(function (cache) {
          cache.put(event.request, copy);
        });
        return response;
      });
    })
  );
});
