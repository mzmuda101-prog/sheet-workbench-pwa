const CACHE_VERSION = "20260703-01";
const APP_CACHE = `excel-wb-shell-${CACHE_VERSION}`;
const HEAVY_CACHE = `excel-wb-heavy-${CACHE_VERSION}`;
const RUNTIME_CACHE = `excel-wb-runtime-${CACHE_VERSION}`;
const ASSET_V = "20260701-01";

const SHELL_ASSETS = [
  "./",
  "./index.html",
  "./manifest.json",
  `./styles/app.css?v=${ASSET_V}`,
  "./assets/fonts/space-grotesk-latin.woff2",
  "./assets/fonts/space-grotesk-latin-ext.woff2",
  `./app/core.js?v=${ASSET_V}`,
  `./app/language.js?v=${ASSET_V}`,
  `./app/analysis.js?v=${ASSET_V}`,
  `./app/workbook.js?v=${ASSET_V}`,
  `./app/xlsx-patch.js?v=${ASSET_V}`,
  `./app/table.js?v=${ASSET_V}`,
  `./app/conditional-formatting.js?v=${ASSET_V}`,
  `./app/data-validation.js?v=${ASSET_V}`,
  `./app/derived-columns.js?v=${ASSET_V}`,
  `./app/formulas.js?v=${ASSET_V}`,
  `./app/edit-tools.js?v=${ASSET_V}`,
  `./app/ui-controls.js?v=${ASSET_V}`,
  `./app/scroll-diagnostics.js?v=${ASSET_V}`,
  `./app/debug-loaders.js?v=${ASSET_V}`,
  `./app/build-rows-core.js?v=${ASSET_V}`,
  `./app/build-rows-worker.js?v=${ASSET_V}`,
  `./app/bootstrap.js?v=${ASSET_V}`,
  "./assets/images/favicon.png?v=20260610-05",
  "./assets/images/apple-touch-icon.png?v=20260429-01",
  "./assets/images/icon-512.png",
  "./assets/images/logo-mateusz-transparent.webp",
  "./assets/images/logo-mateusz-orange.webp",
  "./assets/images/logo-refresh.webp",
];

// [EN] Large libs + media — separate bucket; still precached so offline stays intact after install
const HEAVY_ASSETS = [
  "./lib/xlsx.full.min.js",
  "./lib/jszip.min.js",
  "./assets/media/mateusz-intro.mp4",
];

function isStaticAsset(url) {
  return /\.(?:css|js|png|svg|jpg|jpeg|gif|webp|ico|woff2?|mp4)$/i.test(url.pathname);
}

function isHeavyAsset(url) {
  return /\/lib\/(?:xlsx\.full\.min|jszip\.min)\.js$/i.test(url.pathname)
    || /\/assets\/media\/mateusz-intro\.mp4$/i.test(url.pathname);
}

function cacheNameForUrl(url) {
  if (isHeavyAsset(url)) return HEAVY_CACHE;
  return RUNTIME_CACHE;
}

self.addEventListener("install", (event) => {
  event.waitUntil(
    Promise.all([
      caches.open(APP_CACHE).then((cache) => cache.addAll(SHELL_ASSETS)),
      caches.open(HEAVY_CACHE).then((cache) => cache.addAll(HEAVY_ASSETS)),
    ]).catch(() => {})
  );
});

self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(
        keys
          .filter((key) => key !== APP_CACHE && key !== HEAVY_CACHE && key !== RUNTIME_CACHE)
          .map((key) => caches.delete(key))
      )
    )
  );
  self.clients.claim();
});

self.addEventListener("message", (event) => {
  if (event.data && event.data.type === "SKIP_WAITING") {
    self.skipWaiting();
  }
});

self.addEventListener("fetch", (event) => {
  const { request } = event;
  if (request.method !== "GET") return;

  const reqUrl = new URL(request.url);
  const sameOrigin = reqUrl.origin === self.location.origin;

  if (request.mode === "navigate") {
    event.respondWith(
      fetch(request)
        .then((response) => {
          const copy = response.clone();
          caches.open(RUNTIME_CACHE).then((cache) => cache.put(request, copy)).catch(() => {});
          return response;
        })
        .catch(async () => {
          const cachedPage = await caches.match(request);
          return cachedPage || caches.match("./index.html");
        })
    );
    return;
  }

  if (sameOrigin && isStaticAsset(reqUrl)) {
    event.respondWith(
      caches.match(request).then((cached) => {
        const network = fetch(request)
          .then((response) => {
            if (response && response.ok) {
              const copy = response.clone();
              const bucket = isHeavyAsset(reqUrl) ? HEAVY_CACHE : cacheNameForUrl(reqUrl);
              caches.open(bucket).then((cache) => cache.put(request, copy)).catch(() => {});
            }
            return response;
          })
          .catch(() => cached);
        return cached || network;
      })
    );
  }
});
