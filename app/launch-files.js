// launch-files.js — otwieranie pliku „z zewnątrz”: Otwórz za pomocą / Udostępnij
// (ten sam mechanizm co w Documents Workbench, 2026-10-31).
//
// Windows / macOS / ChromeOS (zainstalowana aplikacja w Chrome / Edge): manifest.json
// `file_handlers` wpisuje aplikację do „Otwórz za pomocą” dla .xlsx/.xlsm. Plik przychodzi
// przez launchQueue jako UCHWYT — więc „Zapisz” zapisuje prosto do oryginału (po zgodzie
// przeglądarki), tak jak po „Otwórz”. `launch_handler: navigate-new` = każdy plik we własnym
// oknie, więc niezapisane zmiany w innym oknie nie przepadają.
//
// Android (zainstalowana aplikacja w Chrome): manifest `share_target` = pozycja w menu
// „Udostępnij”. System wysyła plik POST-em na ./share-target; sw.js odkłada go do cache
// i przekierowuje na ./?open=shared — tu go odbieramy. To KOPIA (bez zapisu do oryginału).
//
// iPhone / iPad i Safari: nie obsługują żadnego z tych mechanizmów — kod nic tam nie robi.

const SWB_SHARE_CACHE = "excel-wb-share";
const SWB_SHARE_KEY = "./__shared-file";

// window.handleFile = wersja z drafts.js: sama pyta o porzucenie niezapisanych zmian
async function swbOpenExternal(file, handle = null) {
  if (await window.handleFile(file, handle)) openDefaultSheet();
}

function swbInitFileLaunch() {
  if ("launchQueue" in window && window.LaunchParams && "files" in window.LaunchParams.prototype) {
    window.launchQueue.setConsumer(async (params) => {
      // jedno okno = jeden skoroszyt; przy kilku plikach Chrome i tak otwiera okno na każdy
      const handle = (params?.files || []).find((h) => h.kind === "file");
      if (!handle) return;
      try {
        await swbOpenExternal(await handle.getFile(), handle);
      } catch (e) {
        log(String(e?.message || e), "error");
        toast(t("openExternalFailed"), "error");
      }
    });
  }
  const open = new URLSearchParams(location.search).get("open");
  if (open) history.replaceState(null, "", location.pathname); // odświeżenie strony nie otwiera pliku drugi raz
  if (open === "shared") swbOpenSharedFile();
}

async function swbOpenSharedFile() {
  if (!("caches" in window)) return;
  try {
    const cache = await caches.open(SWB_SHARE_CACHE);
    const res = await cache.match(SWB_SHARE_KEY);
    if (!res) return;
    const name = decodeURIComponent(res.headers.get("X-File-Name") || "") || "arkusz.xlsx";
    const blob = await res.blob();
    await caches.delete(SWB_SHARE_CACHE);
    await swbOpenExternal(new File([blob], name, { type: blob.type || XLSX_MIME }));
  } catch (e) {
    log(String(e?.message || e), "error");
    toast(t("openExternalFailed"), "error");
  }
}

// Po DOMContentLoaded: drafts.js podmienia wtedy handleFile (pytanie o niezapisane zmiany, szkic).
if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", swbInitFileLaunch);
else swbInitFileLaunch();
