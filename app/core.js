// Core runtime: DOM refs, shared state, and base UI helpers.

// Pamięć przeglądarki zablokowana (Safari „Blokuj wszystkie pliki cookie”, część trybów
// prywatnych): samo sięgnięcie po localStorage rzuca SecurityError, a odczyt ustawień
// (język, motyw…) na starcie modułów wywracał cały interfejs. Wtedy podstawiamy pamięć
// tylko na czas sesji (ustawienia nie przetrwają zamknięcia, ale aplikacja działa).
(() => {
  let ok = false;
  try { window.localStorage.getItem("__wb_probe__"); ok = true; } catch (_) {}
  if (!ok) {
    const mem = new Map();
    const shim = {
      getItem: (k) => (mem.has(String(k)) ? mem.get(String(k)) : null),
      setItem: (k, v) => { mem.set(String(k), String(v)); },
      removeItem: (k) => { mem.delete(String(k)); },
      clear: () => mem.clear(),
      key: (i) => Array.from(mem.keys())[i] ?? null,
      get length() { return mem.size; },
    };
    try { Object.defineProperty(window, "localStorage", { value: shim, configurable: true }); } catch (_) {}
  }
})();

// Heurystyka „słabszego urządzenia" — używana do optymalizacji, które wolno
// odpalać TYLKO tam, gdzie się opłacają (np. tańszy render zamrożonej kolumny
// na dotykowych tabletach). Na mocnych urządzeniach nic się nie zmienia.
const IS_LOW_POWER = (() => {
  try {
    const cores = navigator.hardwareConcurrency || 8;
    const mem = navigator.deviceMemory || 8; // Safari nie wspiera → traktuj jak 8
    const coarse = !!(window.matchMedia && window.matchMedia("(pointer: coarse)").matches);
    return cores <= 4 || mem <= 4 || (coarse && cores <= 6);
  } catch (_) {
    return false;
  }
})();
document.documentElement.classList.toggle("low-power", IS_LOW_POWER);

// Safari: sticky + backdrop-filter podczas poziomego scrollu = smearing (znany bug).
// Wykrywamy raz i dodajemy klasę — CSS zdejmuje blur ze sticky komórek tabeli.
const IS_SAFARI = (() => {
  try { return /^((?!chrome|android).)*safari/i.test(navigator.userAgent); } catch { return false; }
})();
document.documentElement.classList.toggle("browser-safari", IS_SAFARI);

// Szkło (backdrop-filter) jest na dotyku WYŁĄCZONE — rozstrzygnięte pomiarem A/B
// na fizycznym iPhonie (patrz komentarz w app.css). `?glass` przywraca je z powrotem,
// gdyby trzeba było powtórzyć pomiar albo sprawdzić tańszy wariant rozmycia.
(() => {
  try {
    const force = /(\?|&)glass\b/.test(location.search)
      || /(^|#)glass/.test(location.hash)
      || localStorage.getItem("glass") === "1";
    document.documentElement.classList.toggle("force-glass", force);
  } catch (_) {
    document.documentElement.classList.toggle("force-glass", /[?&#]glass\b/.test(location.search + location.hash));
  }
})();

// iPadOS 13+ podaje często „MacIntel" + dotyk — osobna klasa pod poprawki scrolla
// (słaby native fling na overflow:auto; telefon zostaje bez zmian).
const IS_IPAD = (() => {
  try {
    const ua = navigator.userAgent || "";
    return /iPad/.test(ua) ||
      (navigator.platform === "MacIntel" && (navigator.maxTouchPoints || 0) > 1);
  } catch (_) {
    return false;
  }
})();
document.documentElement.classList.toggle("device-ipad", IS_IPAD);

const rootEl = document.documentElement;
const appShellEl = document.querySelector(".app");
const logEl = document.getElementById("log");
const statusEl = document.getElementById("status");
const tableEl = document.getElementById("dataTable");
const theadEl = tableEl.querySelector("thead");
const tbodyEl = tableEl.querySelector("tbody");
const tableWrapEl = document.getElementById("tableWrap");
const tableScrollbarEl = document.getElementById("tableScrollbar");
const tableScrollbarInnerEl = document.getElementById("tableScrollbarInner");
const cellStatsBarEl = document.getElementById("cellStatsBar");
const scrollTopFabEl = document.getElementById("scrollTopFab");
const clearSelectionFabEl = document.getElementById("clearSelectionFab");
const emptyStateEl = document.getElementById("emptyState");
const emptyTitleEl = document.getElementById("emptyTitle");
const emptySubEl = document.getElementById("emptySub");
const DEFAULT_EMPTY_TITLE = emptyTitleEl.textContent;
const tablePanelEl = document.querySelector(".table-panel");

const fileInput = document.getElementById("fileInput");
const dropZone = document.getElementById("dropZone");
const fileNameEl = document.getElementById("fileName");
const fileNameTextEl = document.getElementById("fileNameText");
const sheetSelect = document.getElementById("sheetSelect");
const sheetTabsEl = document.getElementById("sheetTabs");
const sheetPickerFabEl = document.getElementById("sheetPickerFab");
const sheetPickerFabNameEl = document.getElementById("sheetPickerFabName");
const sheetPickerOverlayEl = document.getElementById("sheetPickerOverlay");
const sheetPickerListEl = document.getElementById("sheetPickerList");
const headerRowEl = document.getElementById("headerRow");
const autoHeaderRowEl = document.getElementById("autoHeaderRow");
const displayModeEl = document.getElementById("displayMode");
const maxRowsEl = document.getElementById("maxRows");
const zoomLevelEl = document.getElementById("zoomLevel");
const freezeHeadersEl = document.getElementById("freezeHeaders");
const showFontColorsEl = document.getElementById("showFontColors");
const showCellFillsEl = document.getElementById("showCellFills");
const showCellFontsEl = document.getElementById("showCellFonts");
const showCellBordersEl = document.getElementById("showCellBorders");
const wrapCellsEl = document.getElementById("wrapCells");
const virtRowsEl = document.getElementById("virtRows");
const showConditionalFormattingEl = document.getElementById("showConditionalFormatting");
const showSubheadersEl = document.getElementById("showSubheaders");
const recalcDatesEl = document.getElementById("recalcDates");
const rowHeightAllEl = document.getElementById("rowHeightAll");
const colWidthAllEl = document.getElementById("colWidthAll");
const freezeFirstColEl = document.getElementById("freezeFirstCol");
const smartColWidthsEl = document.getElementById("smartColWidths");
const excelLayoutToggleEl = document.getElementById("excelLayoutToggle");
const loadBtn = document.getElementById("loadBtn");
const loadSampleBtn = document.getElementById("loadSampleBtn");
const toolbarToggleEl = document.getElementById("toolbarToggle");

const searchQueryEl = document.getElementById("searchQuery");
const searchQuery2El = document.getElementById("searchQuery2");
const filterModeEl = document.getElementById("filterMode");
const filterMode2El = document.getElementById("filterMode2");
const filter1ColumnsEl = document.getElementById("filter1Columns");
const filter2ColumnsEl = document.getElementById("filter2Columns");
const filter1PickBtn = document.getElementById("filter1Pick");
const filter2PickBtn = document.getElementById("filter2Pick");
const filterEmptyModeEl = document.getElementById("filterEmptyMode");
const filterNegateEl = document.getElementById("filterNegate");
const filterOperatorsEl = document.getElementById("filterOperators");
const filterEmptyMode2El = document.getElementById("filterEmptyMode2");
const filterNegate2El = document.getElementById("filterNegate2");
const filterOperators2El = document.getElementById("filterOperators2");
const onlyNonEmptyEl = document.getElementById("onlyNonEmpty");
const highlightMatchCellsEl = document.getElementById("highlightMatchCells");
const highlightMatchCellsDateEl = document.getElementById("highlightMatchCellsDate");
const applyFilterBtn = document.getElementById("applyFilterBtn");
const filterBadgeEl = document.getElementById("filterBadge");
const sortColumnSelectEl = document.getElementById("sortColumnSelect");
const sortDirectionSelectEl = document.getElementById("sortDirectionSelect");
const addSortRuleBtn = document.getElementById("addSortRuleBtn");
const sortRulesListEl = document.getElementById("sortRulesList");
const sortPresetSelectEl = document.getElementById("sortPresetSelect");
const saveSortPresetBtn = document.getElementById("saveSortPresetBtn");
const applySortPresetBtn = document.getElementById("applySortPresetBtn");
const deleteSortPresetBtn = document.getElementById("deleteSortPresetBtn");

const dateModeEl = document.getElementById("dateMode");
const dateFromEl = document.getElementById("dateFrom");
const dateToEl = document.getElementById("dateTo");
const lastDaysEl = document.getElementById("lastDays");
const dateEmptyModeEl = document.getElementById("dateEmptyMode");
const dateNegateEl = document.getElementById("dateNegate");
const dateColumnsEl = document.getElementById("dateColumns");
const datePickBtn = document.getElementById("datePick");

const resetFiltersBtn = document.getElementById("resetFiltersBtn");
const exportCsvBtn = document.getElementById("exportCsvBtn");
const saveBtn = document.getElementById("saveBtn");
const saveAsBtn = document.getElementById("saveAsBtn");
const resetWidthsBtn = document.getElementById("resetWidthsBtn");
const resetSortBtn = document.getElementById("resetSortBtn");
const readingToggle = document.getElementById("readingToggle");
const quickSearchWrap = document.getElementById("quickSearchWrap");
const quickSearchEl = document.getElementById("quickSearch");
const quickSearchColumnsBtn = document.getElementById("quickSearchColumnsBtn");
const quickSearchBtn = document.getElementById("quickSearchBtn");
const quickSearchActionEl = document.getElementById("quickSearchAction");
const wideLongToggleEl = document.getElementById("wideLongToggle");

const columnPickerEl = document.getElementById("columnPicker");
const columnPickerTitleEl = document.getElementById("columnPickerTitle");
const columnListEl = document.getElementById("columnList");
const columnSearchEl = document.getElementById("columnSearch");
const selectAllBtn = document.getElementById("selectAllBtn");
const clearAllBtn = document.getElementById("clearAllBtn");
const applyPickBtn = document.getElementById("applyPickBtn");
const closePickerBtn = document.getElementById("closePicker");

const themeToggle = document.getElementById("themeToggle");
const langSwitchEl = document.getElementById("langSwitch");
const langSwitchIndicatorEl = langSwitchEl?.querySelector(".lang-switch-indicator") || null;
const langButtons = Array.from(document.querySelectorAll(".lang-button"));
const panelToggle = document.getElementById("panelToggle");
const panelHandle = document.getElementById("panelHandle");
const sidebarEl = document.querySelector(".sidebar");
const sidebarScrim = document.getElementById("sidebarScrim");
const brandRefreshBtn = document.getElementById("brandRefresh");
const appUpdateBtn = document.getElementById("appUpdateBtn");
const networkBadgeEl = document.getElementById("networkBadge");
const heroRightEl = document.getElementById("heroRight");
const loadingOverlayEl = document.getElementById("loadingOverlay");
const loadingTextEl = document.getElementById("loadingText");
const tableSkeletonEl = document.getElementById("tableSkeleton");
let _loadingToken = 0;
let _loadingShowTimer = 0;
let _loadingMode = "overlay"; // [EN] overlay | skeleton | file — soft UI per workload
const toastContainerEl = document.getElementById("toastContainer");
const cellTooltipEl = document.getElementById("cellTooltip");
const quickSearchPopupEl = document.getElementById("quickSearchPopup");
const quickSearchPopupInput = document.getElementById("quickSearchPopupInput");
const quickSearchPopupColumnsBtn = document.getElementById("quickSearchPopupColumnsBtn");
const quickSearchPopupBtn = document.getElementById("quickSearchPopupBtn");
const quickSearchPopupActionEl = document.getElementById("quickSearchPopupAction");
const quickSearchOperatorsEl = document.getElementById("quickSearchOperators");
const quickSearchPopupOperatorsEl = document.getElementById("quickSearchPopupOperators");
const workbookInsightsEl = document.getElementById("workbookInsights");
const sheetInsightsEl = document.getElementById("sheetInsights");
const insightFlagsEl = document.getElementById("insightFlags");
const kpiSummaryEl = document.getElementById("kpiSummary");
const kpiListEl = document.getElementById("kpiList");
const sheetInspectorSummaryEl = document.getElementById("sheetInspectorSummary");
const columnProfilerEl = document.getElementById("columnProfiler");
const sectionNavigatorEl = document.getElementById("sectionNavigator");
const repeatBlockDetectorEl = document.getElementById("repeatBlockDetector");
const durationAnalysisSummaryEl = document.getElementById("durationAnalysisSummary");
const durationAnalysisListEl = document.getElementById("durationAnalysisList");
const aggregationWorkbenchSummaryEl = document.getElementById("aggregationWorkbenchSummary");
const monthlySummaryEl = document.getElementById("monthlySummary");
const aggregationWorkbenchListEl = document.getElementById("aggregationWorkbenchList");
const formulaSearchEl = document.getElementById("formulaSearch");
const formulaFilterEl = document.getElementById("formulaFilter");
const formulaFunctionFilterEl = document.getElementById("formulaFunctionFilter");
const formulaWorkbenchSummaryEl = document.getElementById("formulaWorkbenchSummary");
const formulaWorkbenchListEl = document.getElementById("formulaWorkbenchList");

const quickRangeButtons = Array.from(document.querySelectorAll(".chip[data-range]"));
const prefersReducedMotion = window.matchMedia("(prefers-reduced-motion: reduce)").matches;

/* ───────────────────────────────────────────────────────────────────────────
   Animacje „scene change" — wyłącznie transform/opacity (szczyt listy
   kompozytora), nigdy width/height/top/left. Dwa narzędzia:
     • withSceneTransition() — View Transitions API: morf starego widoku w nowy
       (webowy odpowiednik ReplacementTransform). Feature-detect + fallback.
     • flipRows()            — technika FLIP: wiersze dojeżdżają translateY na
       nowe pozycje po sortowaniu; tożsamość wiersza = data-row-key.
   Oba respektują prefers-reduced-motion i degradują się do natychmiastowej
   zmiany — nie blokują ani nie psują UI, gdy API niedostępne. ───────────── */

// Morf całej sceny (wczytanie arkusza, tryb wide↔long): cross-fade root.
let sceneTransitionInFlight = false;
let activeSceneTransition = null;
function withSceneTransition(mutate) {
  if (
    prefersReducedMotion ||
    typeof document.startViewTransition !== "function" ||
    sceneTransitionInFlight
  ) {
    mutate();
    return;
  }
  sceneTransitionInFlight = true;
  document.documentElement.classList.add("vt-scene");
  const vt = document.startViewTransition(() => mutate());
  activeSceneTransition = vt;
  // skipTransition() odrzuca także vt.ready — bez tego catcha to unhandled rejection.
  if (vt.ready) vt.ready.catch(() => {});
  const cleanup = () => {
    if (activeSceneTransition === vt) activeSceneTransition = null;
    sceneTransitionInFlight = false;
    document.documentElement.classList.remove("vt-scene");
  };
  // .catch przed .finally: gdy skipTransition() dobija animację, vt.finished
  // ODRZUCA się („Transition was skipped") — bez tego leci unhandled rejection.
  vt.finished.catch(() => {}).finally(cleanup);
  // Bezpiecznik: gdyby .finished nigdy się nie rozwiązało (iOS usypia webview
  // standalone W TRAKCIE animacji → zamrożona nakładka ::view-transition łapie
  // dotyk i „nic nie działa" aż do ubicia procesu). Watchdog domyka stan; po
  // wznowieniu z tła timer odpala i czyści zawieszenie.
  setTimeout(() => {
    if (activeSceneTransition === vt) {
      try { vt.skipTransition(); } catch (_) {}
      cleanup();
    }
  }, 1500);
}

// Gdy aplikacja schodzi w tło — NATYCHMIAST domknij animację sceny (skipTransition
// dobija do stanu końcowego i zdejmuje nakładkę View Transitions), żeby po
// wznowieniu nie zostać z zamrożonym, nieklikalnym widokiem (iPad PWA standalone).
function abortSceneTransitionForBackground() {
  if (!activeSceneTransition) return;
  try { activeSceneTransition.skipTransition(); } catch (_) {}
  activeSceneTransition = null;
  sceneTransitionInFlight = false;
  document.documentElement.classList.remove("vt-scene");
}
document.addEventListener("visibilitychange", () => {
  if (document.visibilityState === "hidden") abortSceneTransitionForBackground();
});
window.addEventListener("pagehide", abortSceneTransitionForBackground);

// Sygnał „następny render tabeli ma użyć FLIP" — ustawiany przy sortowaniu,
// resecie sortowania i filtrowaniu (wszystko, co przestawia/odsłania wiersze).
let flipNextRender = false;
// Maks. liczba JEDNOCZEŚNIE animowanych wierszy (movers+entering). Gradient wg
// mocy CPU/RAM i wysokości ekranu — żeby nie zalać kompozytora na słabszych
// urządzeniach, ale i nie okrajać sprawnych. To NIE limit wyświetlania: dzięki
// animacji tylko w paśmie viewportu (flipRows) liczba kandydatów i tak jest mała,
// a ten cap jest bezpiecznikiem. Liczony przy każdym FLIP (obrót iPada zmienia ekran).
// Wiersze, którym FLIP nałożył inline transform/transition — tylko je trzeba
// wyczyścić przed kolejnym pomiarem (zamiast całej tabeli).
let _flipDirtyRows = [];

// Wiersze w okolicy widoku, wybrane wyszukiwaniem binarnym po offsetTop.
// Tabela układa wiersze w rosnącym offsetTop, więc binsearch jest poprawny,
// a kosztuje ~log2(n) odczytów zamiast n.
function flipViewportSlice(rows, buffer) {
  const list = Array.isArray(rows) ? rows : Array.from(rows);
  if (!tableWrapEl || list.length <= 60) return list; // mało wierszy → taniej zmierzyć wszystkie
  const top = tableWrapEl.scrollTop - buffer;
  const bottom = tableWrapEl.scrollTop + tableWrapEl.clientHeight + buffer;
  const offsetAt = (i) => list[i].offsetTop;
  const firstAtLeast = (value) => {
    let lo = 0;
    let hi = list.length - 1;
    let found = list.length;
    while (lo <= hi) {
      const mid = (lo + hi) >> 1;
      if (offsetAt(mid) >= value) { found = mid; hi = mid - 1; } else { lo = mid + 1; }
    }
    return found;
  };
  const start = Math.max(0, firstAtLeast(top) - 1);
  const end = Math.min(list.length, firstAtLeast(bottom) + 1);
  return list.slice(start, end);
}

function computeFlipRowCap() {
  const cores = navigator.hardwareConcurrency || 8;
  const mem = navigator.deviceMemory || 8; // Safari nie wspiera → traktuj jak 8
  let cap = 80;                              // sprawne urządzenia (baza)
  if (cores <= 2 || mem <= 2) cap = 30;      // bardzo słabe
  else if (cores <= 4 || mem <= 4) cap = 45; // słabe (sporo iPadów/telefonów)
  else if (cores <= 6) cap = 65;             // średnie
  const vh = window.innerHeight || 800;      // mniejszy ekran = mniej widać naraz
  if (vh < 680) cap = Math.min(cap, 40);
  else if (vh < 880) cap = Math.min(cap, 65);
  return cap;
}

// ── FLIP przez „duszki” (2026-10-29) ────────────────────────────────────────────
// Chromium animuje transform/opacity na <tr> na WĄTKU GŁÓWNYM: co klatkę przeliczał styl
// każdego animowanego wiersza z jego komórkami (~465 elementów) — ok. 0,5 s pracy na jedno
// sortowanie przy CPU ×4, animacja cięła się na słabszym sprzęcie, a kolejny klik czekał.
// Teraz poruszające się wiersze płyną jako KOPIE w warstwie nad tabelą (zwykłe <div>-y
// z małą tabelą w środku: te same klasy, szerokości kolumn i wysokość), a prawdziwe wiersze
// stoją już na miejscu, przezroczyste. Animacja to wyłącznie transform/opacity kopii, więc
// liczy ją kompozytor (zmierzone: wątek główny 451 → 18 ms na 30 wierszy). Zapas mocy idzie
// w wygląd: wiersz lecący daleko „unosi się” — cień narasta i gaśnie (też sama opacity).
// Koniec animacji = kopie znikają w tej samej klatce, w której wiersze wracają (piksel
// w piksel — strażnik scripts/flip-ghost-playwright.js). Każdy klik, klawisz, kółko,
// dotyk albo przebudowa tabeli kończy animację od razu (nic nie zostaje „w powietrzu”).
let _flipGhost = null;

function flipGhostFinish() {
  const g = _flipGhost;
  if (!g) return;
  _flipGhost = null;
  clearTimeout(g.timer);
  g.mo.disconnect();
  g.stopEvents.forEach((type) => document.removeEventListener(type, flipGhostFinish, true));
  document.removeEventListener("visibilitychange", flipGhostFinish);
  g.anims.forEach((a) => { try { a.cancel(); } catch (_) {} });
  g.rows.forEach((tr) => tr.classList.remove("flip-hidden"));
  g.layer.remove();
}

function flipGhostClone(tr) {
  const c = tr.cloneNode(true);
  c.removeAttribute("id");
  c.removeAttribute("tabindex");
  c.querySelectorAll("[id]").forEach((el) => el.removeAttribute("id"));
  c.querySelectorAll("[tabindex]").forEach((el) => el.removeAttribute("tabindex"));
  c.classList.remove("flip-hidden");
  c.style.transform = "";
  c.style.transition = "";
  c.style.opacity = "";
  return c;
}

function flipGhostAnimate(movers, entering, delayFor, enterOffset) {
  if (typeof Element.prototype.animate !== "function" || !tableWrapEl || !tableEl) return false;
  // nowy silnik wierszy dorysowuje DOM w trakcie przewijania, scalenia i otwarty edytor
  // nie dają się wiernie skopiować jednym wierszem — tam zostaje stary FLIP
  if (tbodyEl.classList.contains("is-virtual") || tbodyEl.querySelector("input.cell-editor, td[colspan], td[rowspan]")) return false;
  // odczyty (układ jest świeży po pomiarze FLIP) — zapisów jeszcze nie było
  const wrapRect = tableWrapEl.getBoundingClientRect();
  const tableRect = tableEl.getBoundingClientRect();
  const originX = tableRect.left - wrapRect.left + tableWrapEl.scrollLeft - tableWrapEl.clientLeft;
  const originY = tableWrapEl.scrollTop - wrapRect.top - tableWrapEl.clientTop;
  const lastRow = tbodyEl.rows[tbodyEl.rows.length - 1];
  const colgroup = tableEl.querySelector("colgroup");
  // Kopia tylko dla wierszy, które w trakcie ruchu da się zobaczyć: start i koniec poza widokiem
  // po TEJ SAMEJ stronie = nic nie widać, wiersz po prostu stoi na miejscu (mniej warstw GPU
  // w pierwszej klatce). Wiersz przelatujący PRZEZ widok (z góry na dół) dostaje kopię.
  const visible = (a, b, h) => Math.max(a, b) + h > wrapRect.top && Math.min(a, b) < wrapRect.bottom;
  const entries = [];
  movers.forEach((m) => { const h = m.tr.getBoundingClientRect().height; if (visible(m.top, m.top + m.dy, h)) entries.push({ ...m, h, kind: "move" }); });
  entering.forEach((e) => { const h = e.tr.getBoundingClientRect().height; if (visible(e.top, e.top, h)) entries.push({ ...e, h, kind: "enter" }); });
  if (!entries.length) return false;
  // „Uniesienie” (cień) dostają wiersze, które lecą daleko — co najmniej dwie wysokości, maks.
  // 10 najdalszych: widać, co przeskoczyło najwięcej, a drobne przesunięcia nie migają cieniem.
  const lifted = new Set(entries
    .filter((e) => e.kind === "move" && Math.abs(e.dy) >= e.h * 2)
    .sort((a, b) => Math.abs(b.dy) - Math.abs(a.dy))
    .slice(0, 10)
    .map((e) => e.tr));

  const layer = document.createElement("div");
  layer.className = "flip-ghosts";
  layer.setAttribute("aria-hidden", "true");
  const tableClass = `${tableEl.className} flip-ghost-table`.trim();
  const ghosts = entries.map((e) => {
    const ghost = document.createElement("div");
    ghost.className = `flip-ghost flip-ghost-${e.kind}`;
    ghost.style.left = `${originX}px`;
    ghost.style.top = `${e.top + originY}px`;
    ghost.style.width = `${tableRect.width}px`;
    ghost.style.height = `${e.h}px`;
    let shadow = null;
    if (lifted.has(e.tr)) {
      shadow = document.createElement("div");
      shadow.className = "flip-ghost-shadow";
      ghost.appendChild(shadow);
    }
    const t = document.createElement("table");
    t.className = tableClass;
    t.style.width = `${tableRect.width}px`;
    if (colgroup) t.appendChild(colgroup.cloneNode(true));
    const body = document.createElement("tbody");
    body.className = tbodyEl.className;
    const row = flipGhostClone(e.tr);
    row.style.height = `${e.h}px`;
    body.appendChild(row);
    // „tbody tr:last-child td” zdejmuje dolną kreskę — kopia nie może być „ostatnia”,
    // jeśli prawdziwy wiersz nie jest ostatni (wypełniacz nie zajmuje miejsca)
    if (e.tr !== lastRow) {
      const pad = document.createElement("tr");
      pad.className = "flip-ghost-pad";
      body.appendChild(pad);
    }
    t.appendChild(body);
    ghost.appendChild(t);
    layer.appendChild(ghost);
    return { e, ghost, shadow };
  });

  // zapisy: warstwa + ukrycie prawdziwych wierszy w tej samej klatce
  tableWrapEl.appendChild(layer);
  entries.forEach((e) => e.tr.classList.add("flip-hidden"));

  const ease = "cubic-bezier(0.22, 1, 0.36, 1)";
  const anims = [];
  let longest = 0;
  ghosts.forEach(({ e, ghost, shadow }) => {
    const delay = delayFor.get(e.tr) || 0;
    if (e.kind === "move") {
      const duration = 380;
      anims.push(ghost.animate(
        [{ transform: `translateY(${e.dy}px)` }, { transform: "none" }],
        { duration, delay, easing: ease, fill: "backwards" },
      ));
      // uniesienie: cień szybko narasta, a gaśnie, zanim wiersz osiądzie
      if (shadow) anims.push(shadow.animate(
        [{ opacity: 0 }, { opacity: 1, offset: 0.3 }, { opacity: 0 }],
        { duration, delay, easing: "ease-out", fill: "both" },
      ));
      longest = Math.max(longest, delay + duration);
    } else {
      const duration = 340;
      anims.push(ghost.animate(
        [{ opacity: 0, transform: `translateY(${enterOffset}px)` }, { opacity: 1, transform: "none" }],
        { duration, delay, easing: ease, fill: "backwards" },
      ));
      longest = Math.max(longest, delay + duration);
    }
  });

  const stopEvents = ["pointerdown", "keydown", "wheel", "touchstart"];
  const g = {
    layer,
    rows: entries.map((e) => e.tr),
    anims,
    stopEvents,
    // przebudowa tabeli w trakcie (filtr, edycja, nowy plik) — kopie byłyby nieaktualne
    mo: new MutationObserver(flipGhostFinish),
    // bezpiecznik: iOS potrafi uśpić stronę w trakcie animacji (patrz withSceneTransition)
    timer: setTimeout(flipGhostFinish, longest + 120),
  };
  g.mo.observe(tbodyEl, { childList: true });
  stopEvents.forEach((type) => document.addEventListener(type, flipGhostFinish, { capture: true, passive: true }));
  document.addEventListener("visibilitychange", flipGhostFinish);
  _flipGhost = g;
  return true;
}

// Sygnał „następny render ma delikatnie rozjaśnić trafione wiersze" — tryb
// „zaznacz" w szybkim szukaniu nie przestawia wierszy (FLIP nic by nie pokazał),
// więc zamiast tego trafienia łagodnie się rozświetlają (czyste opacity, raz).
let animateMatchPulseNextRender = false;

// FLIP: zmierz pozycje wierszy po kluczu → wykonaj mutację DOM → „cofnij"
// transformem → puść do zera. Wiersze obecne przed i po płynnie zjeżdżają na
// nowe miejsca (sort/reset), a wiersze NOWE (np. po poluzowaniu filtra) wchodzą
// delikatnym fade+slide. Pomiary i zapisy stylów rozdzielone na dwa przebiegi,
// żeby nie thrashować layoutu.
function flipRows(mutate) {
  flipGhostFinish(); // poprzednia animacja (szybkie drugie sortowanie) — wiersze na miejsce
  if (prefersReducedMotion || !tbodyEl) { mutate(); return; }
  const oldRows = tbodyEl.querySelectorAll("tr[data-row-key]");
  // Sufit bezpieczeństwa: przy absurdalnie wielkim DOM pomijamy FLIP (czysty render).
  if (!oldRows.length || oldRows.length > 1500) { mutate(); return; }

  // Pasmo viewportu tabeli + zapas — animujemy TYLKO wiersze blisko widoku. Off-screen
  // i tak nie widać, więc nie marnujemy mocy kompozytora (kluczowe na słabszym iPadzie),
  // a widoczne wiersze zawsze glide'ują niezależnie od rozmiaru wyniku.
  const wrapRect = tableWrapEl
    ? tableWrapEl.getBoundingClientRect()
    : { top: 0, bottom: window.innerHeight || 800 };
  const buffer = Math.max(160, (wrapRect.bottom - wrapRect.top) * 0.6);
  const bandTop = wrapRect.top - buffer;
  const bandBottom = wrapRect.bottom + buffer;
  const inBand = (y) => y != null && y >= bandTop && y <= bandBottom;
  const cap = computeFlipRowCap();

  // Wyzeruj resztkowe transformy z POPRZEDNIEJ animacji. Wcześniej szło to pętlą po
  // wszystkich wierszach (200 zapisów stylu), choć inline-style ma tylko garstka
  // animowanych — trzymamy więc ich listę i czyścimy dokładnie je.
  _flipDirtyRows.forEach((tr) => { tr.style.transition = "none"; tr.style.transform = ""; });
  _flipDirtyRows = [];

  // Mierzymy TYLKO wiersze w okolicy widoku. Wcześniej getBoundingClientRect()
  // leciało po wszystkich wierszach dwa razy (przed i po przebudowie) — przy 200
  // wierszach to 400 odczytów layoutu na każde sortowanie/filtrowanie, czyli lwia
  // część kosztu (zmierzone: sortowanie 524 ms z animacją vs 166 ms bez, CPU ×6).
  // Zakres znajdujemy wyszukiwaniem binarnym po offsetTop (~8 odczytów zamiast 200).
  // Świadomy kompromis: wiersz, który WYJEŻDŻA daleko poza widok, nie dostaje już
  // animacji odlotu — i tak nie widać go po ruchu.
  const before = new Map();
  flipViewportSlice(oldRows, buffer).forEach((tr) => {
    const top = tr.getBoundingClientRect().top;
    if (inBand(top)) before.set(tr.dataset.rowKey, top);
  });

  mutate();

  // Przebieg 1 (odczyt): nowe pozycje + dopasowanie do „before" po kluczu.
  const newRows = Array.from(tbodyEl.querySelectorAll("tr[data-row-key]"));
  const measured = flipViewportSlice(newRows, buffer).map((tr) => ({
    tr,
    top: tr.getBoundingClientRect().top,
    prev: before.get(tr.dataset.rowKey),
  }));

  // Przebieg 2 (zapis): ustaw stany startowe (transition: none). Animujemy tylko
  // wiersze w paśmie viewportu i nie więcej niż cap (gradientowy bezpiecznik).
  // Wjazd nowych wierszy (entering) musi być WIDOCZNY: po selektywnym filtrze
  // większość wierszy w widoku to świeże wjazdy z dołu, a nie movers — gdy ich
  // ruch to ledwie 6px, oko rejestruje animację tylko na kilku moverach, które
  // przejechały duży dystans. Większy zjazd + kaskada (niżej) rozwiązują to.
  const ENTER_OFFSET = 18; // px — czytelny wjazd (było 6 → ginęło obok moverów)
  const movers = [];
  const entering = [];
  for (const { tr, top, prev } of measured) {
    if (movers.length + entering.length >= cap) break;
    if (prev == null) {
      if (!inBand(top)) continue; // nowy/wjeżdżający z daleka, ale ląduje poza widokiem
      entering.push({ tr, top });
      continue;
    }
    const dy = prev - top;
    if (Math.abs(dy) < 1) continue;
    if (!inBand(top) && !inBand(prev)) continue; // ruch w całości poza widokiem → snap
    movers.push({ tr, top, dy });
  }
  if (!movers.length && !entering.length) return;

  // Kaskada góra→dół: wiersze bliżej szczytu widoku ruszają pierwsze, kolejne z
  // małym opóźnieniem. Dzięki temu animacja „przelatuje" przez CAŁY widoczny
  // zakres (movers + entering czytane jako jeden ruch), zamiast wyglądać jak
  // skok kilku wierszy. Stagger ma dodawać czytelności, nie spowalniać — krok
  // mały i z sufitem; na słabszych urządzeniach zerowany (mniej pracy kompozytora).
  const lowPower = (navigator.hardwareConcurrency || 8) <= 4 || (navigator.deviceMemory || 8) <= 4;
  const staggerStep = lowPower ? 0 : 14; // ms na wiersz w kolejności od góry
  const maxStagger = 168;                // sufit całej kaskady (≈12 wierszy)
  const delayFor = new Map();
  [...movers, ...entering]
    .sort((a, b) => a.top - b.top)
    .forEach((entry, i) => delayFor.set(entry.tr, Math.min(i * staggerStep, maxStagger)));

  if (flipGhostAnimate(movers, entering, delayFor, ENTER_OFFSET)) return;

  // Zapas (kopia wiersza niemożliwa / brak Web Animations): stary FLIP na samych <tr>.
  movers.forEach(({ tr, dy }) => { tr.style.transition = "none"; tr.style.transform = `translateY(${dy}px)`; });
  entering.forEach(({ tr }) => { tr.style.transition = "none"; tr.style.opacity = "0"; tr.style.transform = `translateY(${ENTER_OFFSET}px)`; });
  _flipDirtyRows = [...movers, ...entering].map((entry) => entry.tr);

  requestAnimationFrame(() =>
    requestAnimationFrame(() => {
      movers.forEach(({ tr }) => {
        const d = delayFor.get(tr) || 0;
        tr.style.transition = `transform 380ms cubic-bezier(0.22, 1, 0.36, 1) ${d}ms`;
        tr.style.transform = "";
      });
      entering.forEach(({ tr }) => {
        const d = delayFor.get(tr) || 0;
        tr.style.transition = `transform 340ms cubic-bezier(0.22, 1, 0.36, 1) ${d}ms, opacity 300ms ease ${d}ms`;
        tr.style.opacity = "";
        tr.style.transform = "";
      });
    })
  );
  // Sprzątanie inline styli po animacji, żeby nie kolidowały z kolejnym renderem.
  // Margines uwzględnia maksymalny stagger + czas trwania najdłuższego przejścia.
  setTimeout(() => {
    movers.forEach(({ tr }) => { tr.style.transition = ""; tr.style.transform = ""; });
    entering.forEach(({ tr }) => { tr.style.transition = ""; tr.style.transform = ""; tr.style.opacity = ""; });
    _flipDirtyRows = [];
  }, maxStagger + 420);
}

let workbook = null;
// Znacznik pokolenia danych arkusza. Rośnie przy wczytaniu pliku i przy każdej
// edycji komórki — cache'e zbudowane na zawartości arkusza (np. wyniki buildRows
// dla autodetekcji agregacji) porównują go i unieważniają się same.
let sheetDataStamp = 0;
// Przy wczytaniu pliku ten sam ZIP był rozpakowywany trzy razy niezależnie
// (mapa stylów, formatowanie warunkowe, data validation), a każda z nich czytała
// jeszcze te same wpisy xl/workbook.xml i xl/_rels/workbook.xml.rels.
// Tu trzymamy jedną instancję na plik + cache rozpakowanych wpisów tekstowych.
let _sharedZipBytes = null;
let _sharedZipPromise = null;
// Cache odczytów wiązany z KONKRETNĄ instancją zip — ta sama funkcja bywa wołana
// także z innym obiektem zip (np. przy ponownym otwarciu pliku), więc klucz musi
// obejmować instancję, nie samą ścieżkę.
const _sharedZipTextCache = new WeakMap();

function getSharedWorkbookZip(bytes) {
  if (typeof JSZip === "undefined" || !bytes) return Promise.resolve(null);
  if (_sharedZipBytes === bytes && _sharedZipPromise) return _sharedZipPromise;
  _sharedZipBytes = bytes;
  _sharedZipPromise = JSZip.loadAsync(bytes).catch((err) => {
    // nieudane rozpakowanie nie może zostać zapamiętane jako wynik
    _sharedZipBytes = null;
    _sharedZipPromise = null;
    throw err;
  });
  return _sharedZipPromise;
}

// Rozpakowany XML arkuszy potrafi ważyć wielokrotnie więcej niż sam plik, a jest
// potrzebny tylko na czas wczytywania — po nim zwalniamy pamięć.
function releaseSharedZipTextCache() {
  if (_sharedZipPromise) _sharedZipPromise.then((zip) => { if (zip) _sharedZipTextCache.delete(zip); }).catch(() => {});
}

// Odczyt wpisu ZIP jako tekst, z pamięcią w obrębie jednego pliku.
async function readSharedZipText(zip, entryPath) {
  if (!zip || !entryPath) return undefined;
  let perZip = _sharedZipTextCache.get(zip);
  if (!perZip) {
    perZip = new Map();
    _sharedZipTextCache.set(zip, perZip);
  }
  if (perZip.has(entryPath)) return perZip.get(entryPath);
  const file = zip.file(entryPath);
  const text = file ? await file.async("string") : undefined;
  perZip.set(entryPath, text);
  return text;
}
// XML arkusza BEZ <sheetData> — do DOMParsera w czytnikach formatowania warunkowego
// i walidacji danych. Obie sekcje (i extLst z x14:*) leżą ZA sheetData (kolejność
// wymusza schemat OOXML), a sheetData to prawie cały plik: parsowanie go do DOM tylko
// po to, żeby go pominąć, kosztowało ~1 s na realnym arkuszu przy CPU ×6.
// Zostaje znacznik otwierający korzeń (wszystkie xmlns) + ogon po sheetData, więc
// dokument jest poprawny, a prefiksy x14:/xm: rozwiązują się jak wcześniej.
// Gdy struktura jest nietypowa — zwracamy pełny XML (dawne zachowanie).
function sheetXmlWithoutData(xml) {
  if (typeof xml !== "string") return xml;
  const rootMatch = /<([A-Za-z_][\w.-]*:)?worksheet\b[^>]*>/.exec(xml);
  if (!rootMatch) return xml;
  const prefix = rootMatch[1] || "";
  const closeTag = `</${prefix}sheetData>`;
  let tailStart = xml.indexOf(closeTag, rootMatch.index);
  if (tailStart !== -1) {
    tailStart += closeTag.length;
  } else {
    const selfClosing = new RegExp(`<${prefix}sheetData\\b[^>]*/>`).exec(xml);
    if (!selfClosing) return xml;
    tailStart = selfClosing.index + selfClosing[0].length;
  }
  return rootMatch[0] + xml.slice(tailStart);
}

function bumpSheetDataStamp() { sheetDataStamp += 1; }

// <Relationship …/> → Map(Id → Target) BEZ względu na kolejność atrybutów. Excel pisze
// Id przed Target, openpyxl/LibreOffice odwrotnie — regex „Id…Target” nie widział wtedy
// żadnego arkusza: bez stylów (czcionki, ramki, format pustych komórek), formatowania
// warunkowego i list wyboru z takich plików.
function parseRelationshipTargets(relsXml) {
  const map = new Map();
  const tagRe = /<Relationship\b[^>]*>/g;
  let m;
  while ((m = tagRe.exec(relsXml || ""))) {
    const id = (m[0].match(/\bId="([^"]+)"/) || [])[1];
    const target = (m[0].match(/\bTarget="([^"]+)"/) || [])[1];
    if (id && target) map.set(id, target);
  }
  return map;
}
let currentHeaders = [];
let baseRows = [];
let viewRows = [];
let matchedRowIndexes = new Set(); // wiersze pasujące do quick search w trybie "zaznacz"
let quickSearchHighlightMode = false; // true = zaznacz (pokaż wszystkie, wyróżnij WIERSZE)
let quickSearchCellsMode = false;     // true = pokaż wszystkie, podświetl pasujące KOMÓRKI (bez wierszy)
let quickSearchFilterCellsMode = false; // true = filtruj wiersze ORAZ podświetl pasujące komórki w pozostałych
// Mapa rowIndex0 -> Set(colIndex) komórek, które pozytywnie dopasowały filtr.
// Wypełniana tylko gdy highlightMatchedCells === true. Służy do subtelnego
// podświetlenia w siatce komórek „dzięki którym" wiersz przeszedł filtr.
let matchedCellsByRow = new Map();
let highlightMatchedCells = false; // true = podświetl pasujące komórki po filtrowaniu
let filtersCommitted = false; // true dopiero po jawnym „Filtruj" — wtedy applyFilters ukrywa wiersze
// Walidacja listą referencyjną: gdy showOnly === true, applyFilters zostawia w widoku
// tylko wiersze z wartością w `colIdx` SPOZA `allowed` (kompozycja z filtrami tekstu/dat).
let validationState = { colIdx: -1, allowed: null, ignoreEmpty: true, caseInsensitive: true, showOnly: false };
// Map<sheetName, Map<cellRef, xfIndex>> odzyskana z surowego .xlsx (JSZip) — bo ten
// build xlsx-js-style gubi font/border/alignment z cell.s; indeks pozwala je odtworzyć
// z wb.Styles.CellXf/Fonts/Fills. null = brak (sample, nie-xlsx, błąd) → fallback do cell.s.
let currentStyleIndexMap = null;
// Opcjonalne pokazywanie stylów komórek z pliku (panel „Widok"). Domyślnie wszystko
// włączone — odwzorowanie Excela. Czytane w applyCellStyle przy każdym renderze.
let cellStyleShowFontColors = true;
let cellStyleShowFills = true;
let cellStyleShowFonts = true; // rodzina + rozmiar
let cellStyleShowBorders = true;
// Formatowanie warunkowe (CF): reguły + dxf parsowane z surowego .xlsx przy wczytaniu,
// ewaluowane leniwie per arkusz (cache). Pokazują kolory/tła zmienione przez CF w Excelu.
let cellStyleShowConditionalFormatting = true;
let cellStyleShowSubheaders = true; // jasnozielone podświetlenie wykrytych podnagłówków
let recalcDateFormulas = true; // przeliczaj podgląd formuł zależnych od TODAY()/NOW()
let cellStyleSmartWidths = true; // true = dopasuj do większości (p90, przycina skrajnie długie); false = zmieść wszystko
let currentDxfs = [];        // [{fontColor, fillColor}] z xl/styles.xml <dxfs>
let currentCFRules = null;   // Map<sheetName, Array<block>> z <conditionalFormatting>
let cfEvalCache = new Map(); // Map<sheetName, Map<cellRef, {fontColor?, fillColor?}>>
// Tabele Excela: nazwa(lower) -> { columns: { nazwaKol(lower): absColIndex } }. Z xl/tables/*.xml.
// Do rozwijania odwołań strukturalnych w formułach (np. Tabela[[#This Row],[od]]) przy
// odświeżaniu formuł zależnych od TODAY().
let currentTables = {};
let quickSearchOperatorsEnabled = false; // true = &&/|| traktowane jako operatory
let currentFileName = "";
// Oryginalne bajty wczytanego pliku (Uint8Array). Służą do zapisu metodą ZIP-patch:
// podmieniamy tylko wartości edytowanych komórek w sheet.xml, zachowując tabele,
// wykresy, style i formuły. null = brak (np. nieobsługiwany przypadek) → fallback XLSX.write.
let originalFileBytes = null;
// Zarejestrowane edycje komórek do naniesienia przy zapisie:
// { [sheetName]: { [cellRef]: { v, t } | null } }  (null = usunięcie komórki)
let pendingEdits = {};
// Uchwyt pliku z File System Access API (jeśli plik otwarto przez showOpenFilePicker
// lub zapisano przez showSaveFilePicker). Pozwala nadpisać oryginał w miejscu.
// null = plik wczytany przez <input>/drag-drop (brak uchwytu) lub przeglądarka bez FSA.
let currentFileHandle = null;
// Czy działamy w osadzonym (cross-origin) iframe, np. podgląd webu w VS Code.
// File System Access API jest tam wystawione w window, ale jego wywołanie rzuca
// SecurityError ("Cross origin sub frames aren't allowed to show a file picker"),
// więc musimy je tam wyłączyć i wrócić do natywnego <input> / pobierania.
const isEmbeddedFrame = (() => {
  try {
    return window.self !== window.top;
  } catch {
    return true; // brak dostępu do window.top = cross-origin iframe
  }
})();
// Wykrycie File System Access API. Zapis w miejscu i picker tylko gdy dostępne
// i NIE w osadzonym iframe; inaczej fallback do pobrania pliku (iOS Safari, Firefox, VS Code preview).
const canFSA = typeof window !== "undefined" && "showSaveFilePicker" in window && !isEmbeddedFrame;
const canOpenFSA = typeof window !== "undefined" && "showOpenFilePicker" in window && !isEmbeddedFrame;
let currentSheetName = "";
let currentHeaderRow = 1;
let currentStartCol = 0;
let currentMerges = [];
let currentHeaderStyles = [];
let currentSheetColWidths = [];
let currentSheetRowHeights = {};
let currentWorkbookStats = null;
let currentSheetStats = null;
let currentKpiEntries = [];
let currentKpiAnchorRow = 1;
let currentColumnProfiles = [];
let currentSections = [];
let currentRepeatingBlocks = [];
let currentFormulaEntries = [];
let currentDisplayModel = null;
let tableViewMode = "wide";

const columnSelections = {
  filter1: new Set(),
  filter2: new Set(),
  date: new Set(),
};
let activePickerKey = null;
let lastPickerTriggerEl = null;
let sortState = { col: "", dir: "asc" };
let multiSortState = [];
let manualColumnWidths = {};
let manualRowHeights = {};   // rowIndex0 -> px (ręczne przeciąganie pojedynczych wierszy)
let manualRowHeightAll = 0;  // px > 0 = jednolita wysokość wszystkich wierszy (pole w „Widok")
let manualColWidthAll = 0;   // px > 0 = jednolita szerokość wszystkich kolumn (pole w „Widok")
let hasUnsavedChanges = false;
let focusedCellState = null;
// Wynik ostatniego applyFilters (ile widać / z ilu) — czyta go pasek aktywnych filtrów.
let lastAppliedFilters = null;
let selectedCellState = null;
let syncingHorizontalScroll = false;
let tooltipHideTimer = null;
let durationAnalysisState = {
  statusFilter: "all",
  sortMetric: "avg",
  showCount: 14,
};
let currentAggregationMeasureCandidates = [];
let currentAggregationGroupOptions = [];
let aggregationWorkbenchState = {
  sourceMode: "auto",
  scopeMode: "filtered",
  headerRowChoice: "auto",
  customHeaderRow: 1,
  groupBy: "",
  groupBy2: "",
  groupBy3: "",
  measures: ["count_rows"],
  aggregation: "count",
  matchMode: "contains",
  groupMode: "exact",
  groupPattern: "=*",
  similarPairs: [], // „Podobne nazwiska”: ręcznie zatwierdzone literówki [[a, b]] (similar-names.js)
  showCount: 20,
  havingMode: "all",
  havingValue: 10,
  measureFilterMode: "all",
  measureFilterValue: "",
  resultSearch: "",
  resultSearchOperators: false, // operatory (&&, ||, !, {}, >>, <<) w szukajce wyników
};
const APP_BUILD_VERSION = "20261031-04";

// Coalesced view refresh — jedna klatka zamiast kaskady render*() w handlerze.
let _viewRefreshRaf = 0;
const _viewRefreshPending = { table: false, analyses: false, formula: false, filterBadge: false };

function flushViewRefresh() {
  const p = { ..._viewRefreshPending };
  Object.keys(_viewRefreshPending).forEach((k) => { _viewRefreshPending[k] = false; });
  if (p.table && typeof renderActiveTable === "function") renderActiveTable();
  if (p.analyses) {
    if (typeof renderInsights === "function") renderInsights();
    if (typeof renderKpiExtractor === "function") renderKpiExtractor();
    if (typeof renderSheetInspectorSummary === "function") renderSheetInspectorSummary();
    if (typeof renderColumnProfiles === "function") renderColumnProfiles();
    if (typeof renderSections === "function") renderSections();
    if (typeof renderRepeatingBlocks === "function") renderRepeatingBlocks();
    if (typeof renderSmartFilterUi === "function") renderSmartFilterUi();
    if (typeof renderDurationAnalysis === "function") renderDurationAnalysis();
    if (typeof renderAggregationWorkbench === "function") renderAggregationWorkbench();
  }
  if (p.formula && typeof renderFormulaWorkbench === "function") renderFormulaWorkbench();
  if (p.filterBadge && typeof updateFilterBadge === "function") updateFilterBadge();
}

function scheduleViewRefresh(opts = {}) {
  const o = opts === true ? { table: true, analyses: true, formula: true } : opts;
  if (o.table || o.all) _viewRefreshPending.table = true;
  if (o.analyses || o.all) _viewRefreshPending.analyses = true;
  if (o.formula || o.all) _viewRefreshPending.formula = true;
  if (o.filterBadge) _viewRefreshPending.filterBadge = true;
  if (o.sync || o.immediate) { flushViewRefresh(); return; }
  if (_viewRefreshRaf) return;
  _viewRefreshRaf = requestAnimationFrame(() => {
    _viewRefreshRaf = 0;
    flushViewRefresh();
  });
}

// Leniwe cursor-hint (~29KB parse) — idle lub pierwsza interakcja z data-hint.
let _cursorHintPromise = null;
function ensureCursorHint() {
  if (window.MateuszCursorHint) return Promise.resolve(window.MateuszCursorHint);
  if (_cursorHintPromise) return _cursorHintPromise;
  _cursorHintPromise = new Promise((resolve, reject) => {
    const s = document.createElement("script");
    s.src = `app/cursor-hint.js?v=${APP_BUILD_VERSION}`;
    s.async = false;
    s.onload = () => resolve(window.MateuszCursorHint);
    s.onerror = () => { _cursorHintPromise = null; reject(new Error("cursor-hint")); };
    document.head.appendChild(s);
  });
  return _cursorHintPromise;
}

function initCursorHintsWhenReady() {
  if (window.__cursorHintsInited) return;
  if (!window.MateuszCursorHint || typeof t !== "function") return;
  window.MateuszCursorHint.initCursorHints({ fallbackHint: t("hintDefault") });
  window.__cursorHintsInited = true;
}

function scheduleCursorHintInit() {
  const boot = () => ensureCursorHint().then(initCursorHintsWhenReady).catch(() => {});
  const ric = window.requestIdleCallback;
  if (ric) ric(() => boot(), { timeout: 2500 });
  else window.setTimeout(boot, 400);
  const early = (e) => {
    const el = e.target && e.target.closest && e.target.closest("[data-hint], [data-hint-pl], [data-hint-en]");
    if (!el) return;
    document.removeEventListener("pointerover", early, true);
    document.removeEventListener("focusin", early, true);
    boot();
  };
  document.addEventListener("pointerover", early, true);
  document.addEventListener("focusin", early, true);
}

const THEME_KEY = "excel-workbench-theme";
const MAX_ROWS_KEY = "excel-workbench-max-rows";
const EXCEL_LAYOUT_KEY = "excel-workbench-excel-layout";
const CELL_STYLE_PREFS_KEY = "excel-workbench-cell-style-prefs";
const ROW_HEIGHT_KEY = "excel-workbench-row-height-all";
const COL_WIDTH_KEY = "excel-workbench-col-width-all";
const SORT_PRESETS_KEY = "excel-workbench-sort-presets";
// v2: stary klucz zapisywał "0" przy KAŻDYM starcie (nawet bez decyzji użytkownika),
// więc nowy domyślny stan na telefonie nigdy by nie zadziałał. Nowy klucz zapisujemy
// wyłącznie po świadomym kliknięciu w przełącznik paska.
const TOOLBAR_COLLAPSED_KEY = "excel-workbench-toolbar-collapsed-v2";
// Tryb zaznaczania wynika z GESTU, nie z ustawienia:
//   zwykly klik / wejscie Tabem  -> caly wiersz (selectionKind = "row")
//   Shift+klik, Shift+strzalka   -> poziom komorki (selectionKind = "cell")
//   Shift+Spacja                 -> przelacza tam i z powrotem (jedyne wejscie bez myszy)
// Strzalki ruszaja tym, co jest aktualnie zaznaczone: w trybie wiersza gora/dol
// przesuwaja wiersz, a lewo/prawo przewijaja tabele w bok (wiersz nie ma "boku").
let selectionKind = "row";
const INTRO_PLAYED_KEY = "introPlayed";

// Dziennik ma limit linii: wcześniej rósł bez końca (każde wczytanie, zapis, edycja
// dopisywały węzeł DOM na całą sesję). toLocaleTimeString z locale budował przy każdym
// wpisie nowy formater Intl — drogi w budowie, tani w użyciu — więc trzymamy jeden.
const LOG_MAX_LINES = 200;
let _logTimeFmt = null;
let _logTimeFmtLocale = "";
function log(msg, type = "info") {
  const locale = I18N[currentLang].locale;
  if (!_logTimeFmt || _logTimeFmtLocale !== locale) {
    _logTimeFmt = new Intl.DateTimeFormat(locale, { timeStyle: "medium" }); // = toLocaleTimeString(locale)
    _logTimeFmtLocale = locale;
  }
  const line = document.createElement("div");
  line.className = `log-line log-${type}`;
  line.textContent = `${_logTimeFmt.format(new Date())} ${msg}`;
  logEl.prepend(line);
  while (logEl.childElementCount > LOG_MAX_LINES) logEl.lastElementChild.remove();
}

// ── WYŁĄCZONE 2026-09-21 (decyzja Mateusza, po testach na iPhonie) ─────────────
// Własne przyciski kopiuj/wklej na dotyku: pasek „Akcje" nad siatką ORAZ przycisk
// „Wklej" w pustej komórce. Powód wyłączenia: w edytorze działa systemowe menu
// (iOS/Android), a schowek jest wspólny — „Kopiuj" z apki widać w natywnym „Wklej",
// więc własne przyciski w większości przypadków dublują system i zabierają miejsce,
// którego na telefonie jest mało.
// KOD ZOSTAJE (razem z testami) — to jest przełącznik, nie usunięcie. Czego brakuje
// przy wyłączonym pasku: kopiowania/wklejania ZAKRESU komórek i „Wypełnij w dół/w prawo"
// bez klawiatury — systemowe menu działa tylko wewnątrz otwartego pola tekstowego.
// Włączenie: `?cellactions=1` w adresie albo
// localStorage.setItem("excel-workbench-cell-actions", "1") i przeładowanie.
const CELL_ACTIONS_ENABLED = (() => {
  try {
    if (/[?&]cellactions=1/.test(location.search)) return true;
    return localStorage.getItem("excel-workbench-cell-actions") === "1";
  } catch {
    return false;
  }
})();

// Dwa powiadomienia tego samego rodzaju tuż po sobie (np. „Plik wczytany”, a zaraz
// „Wczytano arkusz”, albo kilka „Cofnięto zmianę” pod rząd) nie piętrzą się jeden na
// drugim — kolejne PODMIENIA treść poprzedniego. Na telefonie stos toastów u góry zasłaniał
// szukanie i przyciski (film z iPhone'a 2026-10-31: trzy „Cofnięto zmianę (1 kom.)”).
// Ostrzeżenia i błędy zostają osobno — każde niesie inną informację, której nie wolno zgubić.
const TOAST_MERGE_TYPES = new Set(["success", "info"]);
let _lastMergeToast = null; // { el, label, type, timer, at }
const TOAST_MERGE_MS = 2000;

function toast(msg, type = "info") {
  const last = _lastMergeToast;
  if (TOAST_MERGE_TYPES.has(type) && last && last.type === type && last.el.isConnected && !last.el.classList.contains("out")
    && performance.now() - last.at < TOAST_MERGE_MS) {
    last.label.textContent = msg;
    last.at = performance.now();
    clearTimeout(last.timer);
    last.timer = scheduleToastOut(last.el);
    return;
  }
  const toastEl = document.createElement("div");
  toastEl.className = `toast ${type}`;

  const icon = document.createElement("div");
  icon.className = "toast-icon";
  icon.textContent = type === "success" ? "✓" : type === "error" ? "!" : type === "warning" ? "!" : "i";

  const label = document.createElement("div");
  label.textContent = msg;

  toastEl.appendChild(icon);
  toastEl.appendChild(label);
  toastContainerEl.appendChild(toastEl);

  const timer = scheduleToastOut(toastEl);
  if (TOAST_MERGE_TYPES.has(type)) _lastMergeToast = { el: toastEl, label, type, timer, at: performance.now() };
}

function scheduleToastOut(toastEl) {
  return setTimeout(() => {
    toastEl.classList.add("out");
    setTimeout(() => toastEl.remove(), 200);
  }, 2800);
}

function isQuickSearchPopupOpen() {
  return !!(quickSearchPopupEl && !quickSearchPopupEl.classList.contains("hidden"));
}

// [EN] Hide skeleton/overlay; restore table or empty state without waiting for rAF render
function hideLoadingUI() {
  if (loadingOverlayEl) {
    loadingOverlayEl.classList.add("hidden");
    loadingOverlayEl.classList.remove("is-skeleton", "is-file");
  }
  const skeletonWasOn = tableSkeletonEl && !tableSkeletonEl.classList.contains("hidden");
  if (tableSkeletonEl) {
    tableSkeletonEl.classList.add("hidden");
    tableSkeletonEl.setAttribute("aria-hidden", "true");
  }
  if (document.body) document.body.removeAttribute("aria-busy");
  if (!skeletonWasOn) return;
  // Przywróć widok po skeletone — renderTable i tak dociągnie treść w następnym rAF.
  if (typeof currentHeaders !== "undefined" && currentHeaders.length && typeof showTable === "function") {
    showTable();
  } else if (emptyStateEl) {
    emptyStateEl.classList.remove("hidden");
    if (tableWrapEl) tableWrapEl.classList.add("hidden");
    if (tableScrollbarEl) tableScrollbarEl.classList.add("hidden");
  }
}

function revealLoadingUI(mode) {
  if (!loadingOverlayEl) return;
  loadingOverlayEl.classList.remove("hidden", "is-skeleton", "is-file");
  if (mode === "skeleton" && tableSkeletonEl) {
    loadingOverlayEl.classList.add("is-skeleton");
    if (emptyStateEl) emptyStateEl.classList.add("hidden");
    if (tableWrapEl) tableWrapEl.classList.add("hidden");
    if (tableScrollbarEl) tableScrollbarEl.classList.add("hidden");
    tableSkeletonEl.classList.remove("hidden");
    tableSkeletonEl.setAttribute("aria-hidden", "false");
  } else if (mode === "file") {
    loadingOverlayEl.classList.add("is-file");
    if (tableSkeletonEl) {
      tableSkeletonEl.classList.add("hidden");
      tableSkeletonEl.setAttribute("aria-hidden", "true");
    }
  } else if (tableSkeletonEl) {
    tableSkeletonEl.classList.add("hidden");
    tableSkeletonEl.setAttribute("aria-hidden", "true");
  }
  if (document.body) document.body.setAttribute("aria-busy", "true");
}

// Smart loading: delay show (~90ms) żeby szybkie operacje nie migały UI;
// mode "skeleton" = kości w miejscu tabeli (tańsze niż pełny blur), "file" = miękki overlay.
function setLoading(isLoading, text, options = {}) {
  if (loadingTextEl) loadingTextEl.textContent = text || t("loadingGeneric");
  if (isLoading) {
    _loadingToken += 1;
    const token = _loadingToken;
    _loadingMode = options.mode || "overlay";
    if (document.body) document.body.setAttribute("aria-busy", "true");
    window.clearTimeout(_loadingShowTimer);
    const delay = options.delay != null ? options.delay : 90;
    _loadingShowTimer = window.setTimeout(() => {
      if (token !== _loadingToken) return;
      revealLoadingUI(_loadingMode);
    }, delay);
  } else {
    _loadingToken += 1;
    window.clearTimeout(_loadingShowTimer);
    _loadingShowTimer = 0;
    hideLoadingUI();
  }
}

let _statusAnimRaf = 0;
let _statusPulseKey = null;
// opts.pulseKey: część komunikatu, której zmiana zasługuje na puls (licznik wierszy) —
// sam dopisek „wiersz N” przy chodzeniu strzałkami tylko podmienia tekst. Puls animuje
// kolor (malowanie), więc przy każdej strzałce dokładał ~17 klatek malowania.
function setStatus(msg, opts) {
  // Animuj tylko gdy treść faktycznie się zmienia (np. licznik wierszy po filtrze) —
  // powtórne wywołania z tym samym tekstem (częste przy re-renderach) nie mają migać.
  const changed = statusEl.textContent !== msg;
  statusEl.textContent = msg;
  const key = opts && opts.pulseKey !== undefined ? opts.pulseKey : msg;
  const pulse = changed && key !== _statusPulseKey;
  _statusPulseKey = key;
  if (!pulse) return;
  // Restart animacji BEZ `void statusEl.offsetWidth`. Ten odczyt wymuszał
  // synchroniczne przeliczenie układu CAŁEJ strony — a przy tabeli na tysiące
  // komórek i kilku wywołaniach w trakcie wczytywania zebrało się z tego 374 ms
  // (11% profilu CPU wczytywania arkusza). Klasa wraca w następnej klatce, co dla
  // oka jest nie do odróżnienia, a nie dotyka layoutu.
  statusEl.classList.remove("status-updated");
  if (_statusAnimRaf) cancelAnimationFrame(_statusAnimRaf);
  _statusAnimRaf = requestAnimationFrame(() => {
    _statusAnimRaf = 0;
    statusEl.classList.add("status-updated");
  });
}

// Min/maks tablicy pętlą: Math.min(...arr) w Chromium rzuca RangeError od ~125 tys.
// elementów (kolumna długiego arkusza), w Safari zależy od stosu.
function arrMin(a) { let m = Infinity; for (let i = 0; i < a.length; i++) if (a[i] < m) m = a[i]; return m; }
function arrMax(a) { let m = -Infinity; for (let i = 0; i < a.length; i++) if (a[i] > m) m = a[i]; return m; }

// Tytuł okna/karty: nazwa pliku (z kropką, gdy są niezapisane zmiany), a bez pliku — nazwa
// aplikacji. Zainstalowana aplikacja (okno PWA) sama dokleja swoją nazwę przed tytułem,
// gdy się od niej różni — „* Sheet Workbench” dawało „Sheet Workbench - * Sheet Workbench”
// (film 2026-10-31). Dlatego w oknie PWA bez nazwy aplikacji, w karcie przeglądarki — z nią.
const swbStandaloneMQ = typeof matchMedia === "function"
  ? matchMedia("(display-mode: standalone), (display-mode: window-controls-overlay), (display-mode: minimal-ui)") : null;
function syncDocumentTitle() {
  if (typeof document === "undefined") return;
  if (document.body && document.body.classList.contains("rp-printing")) return; // raport ustawia własny
  const name = typeof currentFileName === "string" && currentFileName ? currentFileName : "";
  if (!name) { document.title = BASE_TITLE; return; }
  const mark = hasUnsavedChanges ? "• " : "";
  const standalone = !!(swbStandaloneMQ && swbStandaloneMQ.matches);
  document.title = standalone ? `${mark}${name}` : `${mark}${name} — ${BASE_TITLE}`;
}

// Licznik edycji: zapis trwa (budowa ZIP + zapis na dysk), a w tym czasie da się
// dalej edytować. Bez niego „Zapisano” gasiło też zmiany wpisane PO zbudowaniu pliku
// (brak ostrzeżenia przy zamknięciu, a drafts.js kasował szkic z tymi zmianami).
let dirtyEpoch = 0;
let savedBuildEpoch = 0;
function setDirtyState(isDirty) {
  if (isDirty) dirtyEpoch++;
  hasUnsavedChanges = !!isDirty;
  statusEl.classList.toggle("unsaved", hasUnsavedChanges);
  syncDocumentTitle();
  if (typeof appFrame !== "undefined") appFrame.syncSave(); // „Zapisz” w nagłówku + liczba zmian
}

// Po udanym zapisie: czysto tylko, jeśli od zbudowania pliku nic się nie zmieniło.
function markSavedAfterBuild() {
  setDirtyState(dirtyEpoch !== savedBuildEpoch);
}

function updateExcelLayoutButtonLabel() {
  if (!excelLayoutToggleEl) return;
  const next = isExcelLayoutEnabled();
  excelLayoutToggleEl.textContent = next ? t("excelViewOn") : t("excelView");
}

function applyZoom() {
  if (!tableEl || !zoomLevelEl) return;
  const zoom = parseFloat(zoomLevelEl.value) || 1;
  tableEl.style.setProperty("--table-zoom", String(zoom));
  if (tableWrapEl) {
    tableWrapEl.style.setProperty("--table-zoom", String(zoom));
  }
  syncFrozenHeaderMetrics();
  syncTableViewportHeight();
}

// Te dwie funkcje CZYTAJĄ layout (getBoundingClientRect) i ustawiają CSS-zmienne.
// Wołane są kilka razy podczas jednego renderu/zmiany (applyZoom + applyFreezeHeaders
// + render tail) — gdy odpalały się synchronicznie tuż po zapisach klas/stylów,
// każda wymuszała pełny re-layout tabeli (forced reflow / layout thrashing, ~270ms
// na słabszym iPadzie przy 500 wierszach). Żaden wołacz nie czyta tych zmiennych
// z powrotem w JS, więc bezpiecznie zlewamy je do JEDNEGO odczytu w najbliższym
// rAF — po naturalnym layoucie przeglądarki, bez wymuszania synchronicznego.
let _frozenMetricsScheduled = false;
let _viewportHeightScheduled = false;

function readFrozenHeaderMetrics() {
  if (!tableWrapEl || !theadEl) return;
  const guideRow = theadEl.querySelector(".guide-row");
  const guideHeight = guideRow ? Math.ceil(guideRow.getBoundingClientRect().height) : 0;
  tableWrapEl.style.setProperty("--frozen-guide-height", `${guideHeight}px`);
}

function syncFrozenHeaderMetrics() {
  if (!tableWrapEl || !theadEl || _frozenMetricsScheduled) return;
  _frozenMetricsScheduled = true;
  requestAnimationFrame(() => {
    _frozenMetricsScheduled = false;
    readFrozenHeaderMetrics();
  });
}

function readTableViewportHeight() {
  if (!tablePanelEl) return;
  const rect = tablePanelEl.getBoundingClientRect();
  const viewportHeight = window.innerHeight || document.documentElement.clientHeight || 720;
  // Odstęp pod tabelą = dolny padding .app (jedno źródło: CSS). Stała w JS rozjeżdżała się
  // z CSS i strona wystawała: na desktopie o 24 px (scrollIntoView komórki przesuwał CAŁĄ
  // stronę, tabela skakała), na telefonie o 18 px (32 px paddingu vs 14 px tu — cały ekran
  // „jechał” pod palcem, 2026-10-25).
  const appEl = tablePanelEl.closest(".app");
  const pad = appEl ? parseFloat(getComputedStyle(appEl).paddingBottom) || 0 : 0;
  const bottomGap = Math.max(12, Math.ceil(pad));
  const available = Math.floor(viewportHeight - rect.top - bottomGap);
  const minHeight = window.matchMedia("(max-width: 768px)").matches ? 320 : 420;
  tablePanelEl.style.setProperty("--table-panel-height", `${Math.max(minHeight, available)}px`);
}

function syncTableViewportHeight() {
  if (!tablePanelEl || _viewportHeightScheduled) return;
  _viewportHeightScheduled = true;
  requestAnimationFrame(() => {
    _viewportHeightScheduled = false;
    readTableViewportHeight();
  });
}

function applyFreezeHeaders() {
  if (!tableWrapEl) return;
  const enabled = !freezeHeadersEl || freezeHeadersEl.checked;
  tableWrapEl.classList.toggle("freeze-headers", enabled);
  tableWrapEl.classList.toggle("headers-unlocked", !enabled);
  syncFrozenHeaderMetrics();
}

function applyFreezeFirstColumn() {
  if (!tableWrapEl) return;
  tableWrapEl.classList.toggle("freeze-first-col", !!(freezeFirstColEl && freezeFirstColEl.checked));
  // uchwyt wypełniania w zamrożonej kolumnie siedzi w komórce, poza nią — w #tableWrap (autofill.js)
  if (typeof swbFillHandleSync === "function") swbFillHandleSync();
}


// ── Wskaźnik trybu wprowadzania (klawiatura vs wskaźnik) ─────────────────────
// WebKit (Safari, w tym iPad z klawiaturą sprzętową) DEKLARUJE wsparcie dla
// :focus-visible, ale nie dopasowuje go przy nawigacji klawiaturą — zmierzone:
// td.matches(":focus-visible") === false po Tabie i strzałkach, mimo CSS.supports() === true.
// Cała widoczność fokusu stała na tej pseudoklasie, więc na iPadzie po prostu nie było
// widać, gdzie się stoi. Klasa na <html> pozwala CSS-owi dorysować obwódkę awaryjnie
// (patrz reguły `.kbd-focus :focus:not(:focus-visible)` w app.css) — w Chromium ta
// warstwa nigdy się nie odpala, bo tam :focus-visible działa i wyklucza ją sam selektor.
(function trackInputModality() {
  const root = document.documentElement;
  document.addEventListener("keydown", () => root.classList.add("kbd-focus"), true);
  const pointerUsed = () => root.classList.remove("kbd-focus");
  document.addEventListener("pointerdown", pointerUsed, true);
  document.addEventListener("mousedown", pointerUsed, true);
  document.addEventListener("touchstart", pointerUsed, { capture: true, passive: true });
})();

// ── Zasięg klawiatury w Safari/WebKit (iPad z klawiaturą) ────────────────────
// Safari domyślnie prowadzi Tab TYLKO po polach tekstowych i listach wyboru —
// przyciski i pola wyboru są pomijane (ustawienie „Press Tab to highlight each item"
// jest wyłączone). Zmierzone w WebKicie: obieg Tab po oknie szukania to
// „Tryb → Akcja → ‹BODY› → td", czyli fokus wypadał z okna, a „Operatory
// wyszukiwania" i „Szukaj" były NIEOSIĄGALNE. W sidebarze Tab widział wyłącznie
// nagłówki paneli i trzy pola liczbowe.
// Jawny tabindex="0" wciąga te elementy z powrotem do nawigacji sekwencyjnej.
// Stosujemy to we WSZYSTKICH przeglądarkach świadomie: w Chromium jest to no-op
// (przyciski i tak są w kolejności Tab, a tabindex="0" zachowuje kolejność DOM),
// więc nie ma tu czego bramkować — inaczej niż przy warstwie CSS, która realnie
// mogłaby dublować obwódki i dlatego wyłącza się sama przez :not(:focus-visible).
// Elementy z WŁASNYM tabindex (np. pozycje listy wyników = -1) zostają nietknięte.
function ensureKeyboardReachable(root) {
  const scope = root || document;
  const sel = 'button:not([tabindex]), input[type="checkbox"]:not([tabindex]), input[type="radio"]:not([tabindex])';
  const apply = (el) => { el.setAttribute("tabindex", "0"); };
  if (scope.matches && scope.matches(sel)) apply(scope);
  scope.querySelectorAll(sel).forEach(apply);
}

// Panele analiz, agregacje i podobne dosypują przyciski długo po starcie, więc
// samo wywołanie przy inicjalizacji nie wystarcza (zostawało ~31 nieosiągalnych).
// Wyliczanie miejsc po jednym było kruche — obserwator łapie wszystko, także kod
// dopisany w przyszłości.
// KOSZT: świadomie omijamy #dataTable jednym `closest()` na dodany węzeł, więc
// przebudowa ~1500 komórek nie płaci tu nic poza tym testem. Właściwe skanowanie
// (querySelectorAll) idzie do bezczynności, nigdy w trakcie renderu.
let pendingKbdRoots = null;
let kbdScanScheduled = false;

function flushKeyboardReachability() {
  kbdScanScheduled = false;
  const roots = pendingKbdRoots;
  pendingKbdRoots = null;
  if (!roots) return;
  roots.forEach((node) => { if (node.isConnected) ensureKeyboardReachable(node); });
}

let kbdReachObserver = null;

function observeKeyboardReachability() {
  if (typeof MutationObserver !== "function" || !document.body || kbdReachObserver) return;
  kbdReachObserver = new MutationObserver((records) => {
    for (const rec of records) {
      for (const node of rec.addedNodes) {
        if (node.nodeType !== 1) continue;
        if (node.closest && node.closest("#dataTable")) continue; // siatka nie ma przycisków
        (pendingKbdRoots || (pendingKbdRoots = [])).push(node);
      }
    }
    if (pendingKbdRoots && !kbdScanScheduled) {
      kbdScanScheduled = true;
      if (typeof requestIdleCallback === "function") requestIdleCallback(flushKeyboardReachability, { timeout: 500 });
      else setTimeout(flushKeyboardReachability, 100);
    }
  });
  kbdReachObserver.observe(document.body, { childList: true, subtree: true });
}
