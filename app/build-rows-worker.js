// [EN] Dedicated worker — importScripts xlsx + shared buildRowsCore
importScripts("../lib/xlsx.full.min.js", "build-rows-core.js");

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
