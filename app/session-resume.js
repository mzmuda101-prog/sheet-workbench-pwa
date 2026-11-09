// session-resume.js — „Wrócić do B234?” po ponownym otwarciu tego samego pliku (audyt QoL S11).
//
// Szkic (drafts.js) chroni NIEZAPISANE zmiany, ale po czystym zamknięciu kontekst czytania
// przepadał: przy kolejnym otwarciu pliku start był zawsze od góry pierwszego arkusza.
// Tu lekka, lokalna pamięć per plik (nazwa + rozmiar + data modyfikacji): arkusz i aktywna
// komórka. Treści pliku NIE przechowujemy. Po otwarciu — dymek z akcją „Wróć”, bez
// samowolnego przeskakiwania (czasem chcesz zacząć od początku).

const SWB_SESSIONS_KEY = "swb-sessions-v1";
const SWB_SESSIONS_MAX = 30;

const swbSessionResume = (() => {
  let fileKey = "";
  let timer = 0;

  const keyOf = (file) => (file && file.name ? `${file.name}|${file.size}|${file.lastModified}` : "");
  function load() {
    try { const m = JSON.parse(localStorage.getItem(SWB_SESSIONS_KEY) || "{}"); return m && typeof m === "object" ? m : {}; } catch { return {}; }
  }
  function store(map) {
    const keys = Object.keys(map).sort((a, b) => (map[b].at || 0) - (map[a].at || 0));
    keys.slice(SWB_SESSIONS_MAX).forEach((k) => delete map[k]);
    try { localStorage.setItem(SWB_SESSIONS_KEY, JSON.stringify(map)); } catch { /* bez pamięci */ }
  }
  function currentRef() {
    const a = typeof cfmtActiveCell === "function" ? cfmtActiveCell() : null;
    return a && a.ref ? a.ref : "";
  }
  function saveNow() {
    clearTimeout(timer);
    if (!fileKey || !workbook || !currentSheetName) return;
    const ref = currentRef();
    const map = load();
    map[fileKey] = { sheet: currentSheetName, ref, at: Date.now() };
    store(map);
  }
  const saveSoon = () => { if (!fileKey) return; clearTimeout(timer); timer = setTimeout(saveNow, 1500); };

  const waitFor = (cond, ms = 6000) => new Promise((resolve) => {
    const t0 = performance.now();
    const tick = () => (cond() ? resolve(true) : performance.now() - t0 > ms ? resolve(false) : setTimeout(tick, 80));
    tick();
  });
  const tableReady = () => typeof currentDisplayModel !== "undefined" && currentDisplayModel && currentDisplayModel.rows && !!currentSheetName;

  async function jumpTo(rec) {
    if (rec.sheet && rec.sheet !== currentSheetName && workbook.Sheets[rec.sheet] && typeof sheetSelect !== "undefined" && sheetSelect) {
      sheetSelect.value = rec.sheet;
      loadBtn.click();
      await waitFor(() => currentSheetName === rec.sheet && tableReady());
    }
    if (rec.ref && typeof sheetEditUi !== "undefined" && sheetEditUi.goToRef) sheetEditUi.goToRef(rec.ref);
  }

  async function offer(key) {
    const rec = load()[key];
    if (!rec || !rec.sheet) return;
    if (!(await waitFor(tableReady))) return;
    if (fileKey !== key) return; // w międzyczasie inny plik
    if (!workbook.Sheets[rec.sheet]) return;
    const sameSheet = rec.sheet === currentSheetName;
    if (sameSheet && (!rec.ref || rec.ref === currentRef())) return; // i tak jesteśmy w tym miejscu
    const multi = (workbook.SheetNames || []).length > 1;
    const where = rec.ref ? (multi ? `${rec.sheet} · ${rec.ref}` : rec.ref) : rec.sheet;
    toast(t("resumeAsk", { where }), "info", { action: t("resumeGo"), duration: 9000, onAction: () => jumpTo(rec) });
  }

  function init() {
    if (typeof window.handleFile !== "function") return;
    const orig = window.handleFile;
    window.handleFile = async function handleFileResume(file, ...rest) {
      saveNow(); // pozycja w poprzednim pliku, zanim go podmienimy
      const r = await orig.call(this, file, ...rest);
      if (workbook && currentFileName === (file && file.name)) {
        const key = keyOf(file);
        fileKey = key;
        if (key) offer(key);
      }
      return r;
    };
    document.addEventListener("keyup", saveSoon, true);
    document.addEventListener("pointerup", saveSoon, true);
    document.addEventListener("visibilitychange", () => { if (document.visibilityState === "hidden") saveNow(); });
    window.addEventListener("pagehide", saveNow);
  }

  return { init, saveNow, _load: load, _key: () => fileKey };
})();

// po drafts.js (ten też owija handleFile przy DOMContentLoaded) — kolejność skryptów to gwarantuje
document.addEventListener("DOMContentLoaded", () => swbSessionResume.init());
