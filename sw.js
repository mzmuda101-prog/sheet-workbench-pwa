const CACHE_VERSION = "20260921-15";
const APP_CACHE = `excel-wb-shell-${CACHE_VERSION}`;
const HEAVY_CACHE = `excel-wb-heavy-${CACHE_VERSION}`;
const RUNTIME_CACHE = `excel-wb-runtime-${CACHE_VERSION}`;
// Spięte z CACHE_VERSION (nie osobna stała) — inaczej npm run release bumpuje tylko
// CACHE_VERSION i literalne ?v= w index.html, a ASSET_V zostaje w tyle: precache
// instalacyjny celuje wtedy w URL-e, których strona już nie prosi (cache miss na
// starcie, offline-first dla JS/CSS realnie nie działa do pierwszego online-visit).
const ASSET_V = CACHE_VERSION;

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
  `./app/smart-filters.js?v=${ASSET_V}`,
  `./app/workbook.js?v=${ASSET_V}`,
  `./app/xlsx-patch.js?v=${ASSET_V}`,
  `./app/table.js?v=${ASSET_V}`,
  `./app/conditional-formatting.js?v=${ASSET_V}`,
  `./app/data-validation.js?v=${ASSET_V}`,
  `./app/derived-columns.js?v=${ASSET_V}`,
  `./app/formulas.js?v=${ASSET_V}`,
  `./app/edit-tools.js?v=${ASSET_V}`,
  `./app/ui-controls.js?v=${ASSET_V}`,
  `./app/transcribe.js?v=${ASSET_V}`,
  `./app/scroll-diagnostics.js?v=${ASSET_V}`,
  `./app/debug-loaders.js?v=${ASSET_V}`,
  `./app/build-rows-core.js?v=${ASSET_V}`,
  `./app/build-rows-worker.js?v=${ASSET_V}`,
  `./app/bootstrap.js?v=${ASSET_V}`,
  "./assets/images/favicon.png?v=20260610-05",
  "./assets/images/apple-touch-icon.png?v=20260429-01",
  "./assets/images/icon-192.png",
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

// Instalacja pobiera TYLKO lekką powłokę. Ciężkie zasoby (xlsx ~900 KB, jszip, film
// z intro) szły wcześniej w tej samej paczce — czyli zaraz po każdej aktualizacji
// telefon ściągał i zapisywał kilka megabajtów dokładnie wtedy, gdy użytkownik
// wczytuje arkusz i zaczyna nim przewijać. Teraz dogrywamy je po aktywacji, z opóźnieniem,
// a gdyby service worker został w międzyczasie uśpiony — i tak trafią do cache przy
// pierwszym użyciu, bo handler fetch zapisuje je do HEAVY_CACHE.
self.addEventListener("install", (event) => {
  event.waitUntil(
    caches.open(APP_CACHE).then((cache) => cache.addAll(SHELL_ASSETS)).catch(() => {})
  );
});

function precacheHeavyAssetsLater(delayMs = 8000) {
  return new Promise((resolve) => {
    setTimeout(async () => {
      try {
        const cache = await caches.open(HEAVY_CACHE);
        for (const asset of HEAVY_ASSETS) {
          // Pojedynczo i sekwencyjnie — równoległe addAll potrafi zapchać łącze telefonu.
          if (await cache.match(asset)) continue;
          try { await cache.add(asset); } catch (_) { /* dogramy przy pierwszym użyciu */ }
        }
      } catch (_) {
        // brak miejsca / prywatny tryb — zostaje ścieżka „cache przy pierwszym użyciu"
      }
      resolve();
    }, delayMs);
  });
}

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
  // Ciężkie zasoby dogrywamy po chwili, już poza ścieżką krytyczną startu.
  event.waitUntil(precacheHeavyAssetsLater());
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
