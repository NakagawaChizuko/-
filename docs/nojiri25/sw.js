const CACHE_NAME = "kaseki25-pwa-v35-ts-multipoint";
const REQUIRED_ASSETS = [
  "./", "./index.html", "./styles.css", "./app.js",
  "./vendor/three.min.js", "./vendor/OrbitControls.js", "./grid_reference_data.js",
  "./team_roster_data.js", "./manifest.webmanifest", "./icon.svg", "./icon-512.png"
];
const OPTIONAL_ASSETS = [
  "./shapes/inline_data.js", "./shapes/palmate_antler.png", "./shapes/triangle.png",
  "./shapes/incisor.png", "./shapes/diamond_hira.png", "./shapes/manifest.json",
  "./shapes/constricted_shape.png", "./shapes/rib_curved.png", "./shapes/c_shape.png",
  "./assets/large-shapes/elephant_upper_molar.png", "./assets/large-shapes/palmate_antler.png",
  "./assets/large-shapes/elephant_lower_molar.png", "./assets/large-shapes/triangle.png",
  "./assets/large-shapes/incisor.png", "./assets/large-shapes/constricted_bone.png",
  "./assets/large-shapes/diamond_hira.png", "./assets/large-shapes/diamond.png",
  "./assets/large-shapes/rib.png", "./assets/large-shapes/constricted_shape.png",
  "./assets/large-shapes/rib_curved.png", "./assets/large-shapes/c_shape.png"
];

function cacheOptional(cache, url) {
  return fetch(url, { cache: "no-cache" })
    .then((response) => (response?.ok ? cache.put(url, response) : null))
    .catch(() => null);
}

self.addEventListener("install", (event) => {
  event.waitUntil(caches.open(CACHE_NAME).then((cache) =>
    cache.addAll(REQUIRED_ASSETS).then(() => Promise.all(OPTIONAL_ASSETS.map((url) => cacheOptional(cache, url))))
  ));
  self.skipWaiting();
});

self.addEventListener("activate", (event) => {
  event.waitUntil(caches.keys().then((keys) => Promise.all(keys.map((key) => key !== CACHE_NAME ? caches.delete(key) : null))));
  self.clients.claim();
});

self.addEventListener("fetch", (event) => {
  if (event.request.method !== "GET") return;
  if (event.request.mode === "navigate") {
    event.respondWith(fetch(event.request, { cache: "no-store" }).then((response) => {
      if (response?.ok) caches.open(CACHE_NAME).then((cache) => cache.put("./index.html", response.clone()));
      return response;
    }).catch(() => caches.match("./index.html")));
    return;
  }
  event.respondWith(caches.match(event.request).then((cached) => cached || fetch(event.request).then((response) => {
    if (response?.ok) caches.open(CACHE_NAME).then((cache) => cache.put(event.request, response.clone()));
    return response;
  }).catch(() => caches.match("./index.html"))));
});
