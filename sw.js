const CACHE_VERSION = "20260628-04";
const APP_CACHE = `excel-wb-shell-${CACHE_VERSION}`;
const RUNTIME_CACHE = `excel-wb-runtime-${CACHE_VERSION}`;

const APP_ASSETS = [
  "./",
  "./index.html",
  "./manifest.json",
  "./styles/app.css?v=20260628-04",
  "./assets/fonts/space-grotesk-latin.woff2",
  "./assets/fonts/space-grotesk-latin-ext.woff2",
  "./lib/xlsx-js-style.bundle.min.js",
  "./lib/jszip.min.js",
  "./app/core.js?v=20260628-04",
  "./app/language.js?v=20260628-04",
  "./app/analysis.js?v=20260628-04",
  "./app/workbook.js?v=20260628-04",
  "./app/xlsx-patch.js?v=20260628-04",
  "./app/table.js?v=20260628-04",
  "./app/conditional-formatting.js?v=20260628-04",
  "./app/data-validation.js?v=20260628-04",
  "./app/derived-columns.js?v=20260628-04",
  "./app/formulas.js?v=20260628-04",
  "./app/edit-tools.js?v=20260628-04",
  "./app/ui-controls.js?v=20260628-04",
  "./app/scroll-diagnostics.js?v=20260628-04",
  "./app/bootstrap.js?v=20260628-04",
  "./assets/images/favicon.png?v=20260610-05",
  "./assets/images/apple-touch-icon.png?v=20260429-01",
  "./assets/images/icon-512.png",
  "./assets/images/logo-mateusz-transparent.webp",
  "./assets/images/logo-mateusz-orange.webp",
  "./assets/images/logo-refresh.webp",
  "./assets/media/mateusz-intro.mp4",
];

function isStaticAsset(url) {
  return /\.(?:css|js|png|svg|jpg|jpeg|gif|webp|ico|woff2?|mp4)$/i.test(url.pathname);
}

self.addEventListener("install", (event) => {
  event.waitUntil(
    caches
      .open(APP_CACHE)
      .then((cache) => cache.addAll(APP_ASSETS))
      .catch(() => {})
  );
});

self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(
        keys
          .filter((key) => key !== APP_CACHE && key !== RUNTIME_CACHE)
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
              caches.open(RUNTIME_CACHE).then((cache) => cache.put(request, copy)).catch(() => {});
            }
            return response;
          })
          .catch(() => cached);
        return cached || network;
      })
    );
  }
});
