// [EN] Dedicated worker — importScripts xlsx + shared buildRowsCore
// Wersja z adresu workera (?v=…) — ten sam URL co w precache SW, więc działa też
// offline, zanim worker zostanie użyty pierwszy raz online.
importScripts("../lib/xlsx.full.min.js", "build-rows-core.js" + self.location.search);

self.onmessage = (event) => {
  const msg = event.data || {};
  if (msg.type !== "buildRows") return;
  try {
    const result = buildRowsCore(msg.sheet, msg.headerRow, msg.workbook, msg.options || {});
    self.postMessage({ id: msg.id, ok: true, result });
  } catch (err) {
    self.postMessage({ id: msg.id, ok: false, error: String(err && err.message ? err.message : err) });
  }
};
