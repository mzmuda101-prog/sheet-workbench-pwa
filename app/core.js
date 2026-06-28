// Core runtime: DOM refs, shared state, and base UI helpers.

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
const DEFAULT_EMPTY_SUB = emptySubEl.textContent;
const tablePanelEl = document.querySelector(".table-panel");

const fileInput = document.getElementById("fileInput");
const dropZone = document.getElementById("dropZone");
const fileNameEl = document.getElementById("fileName");
const fileNameTextEl = document.getElementById("fileNameText");
const sheetSelect = document.getElementById("sheetSelect");
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
const quickSearchModeEl = document.getElementById("quickSearchMode");
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
const toastContainerEl = document.getElementById("toastContainer");
const cellTooltipEl = document.getElementById("cellTooltip");
const quickSearchPopupEl = document.getElementById("quickSearchPopup");
const quickSearchPopupInput = document.getElementById("quickSearchPopupInput");
const quickSearchPopupModeEl = document.getElementById("quickSearchPopupMode");
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
  vt.finished.finally(() => {
    sceneTransitionInFlight = false;
    document.documentElement.classList.remove("vt-scene");
  });
}

// Sygnał „następny render tabeli ma użyć FLIP" — ustawiany przy sortowaniu,
// resecie sortowania i filtrowaniu (wszystko, co przestawia/odsłania wiersze).
let flipNextRender = false;
// Maks. liczba JEDNOCZEŚNIE animowanych wierszy (movers+entering). Gradient wg
// mocy CPU/RAM i wysokości ekranu — żeby nie zalać kompozytora na słabszych
// urządzeniach, ale i nie okrajać sprawnych. To NIE limit wyświetlania: dzięki
// animacji tylko w paśmie viewportu (flipRows) liczba kandydatów i tak jest mała,
// a ten cap jest bezpiecznikiem. Liczony przy każdym FLIP (obrót iPada zmienia ekran).
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

  // Wyzeruj resztkowe transformy z poprzedniej animacji, by mierzyć layout.
  oldRows.forEach((tr) => { tr.style.transition = "none"; tr.style.transform = ""; });
  const before = new Map();
  oldRows.forEach((tr) => {
    const top = tr.getBoundingClientRect().top;
    if (inBand(top)) before.set(tr.dataset.rowKey, top); // tylko widoczne pasmo
  });

  mutate();

  // Przebieg 1 (odczyt): nowe pozycje + dopasowanie do „before" po kluczu.
  const newRows = Array.from(tbodyEl.querySelectorAll("tr[data-row-key]"));
  const measured = newRows.map((tr) => ({
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
      tr.style.transition = "none";
      tr.style.opacity = "0";
      tr.style.transform = `translateY(${ENTER_OFFSET}px)`;
      entering.push({ tr, top });
      continue;
    }
    const dy = prev - top;
    if (Math.abs(dy) < 1) continue;
    if (!inBand(top) && !inBand(prev)) continue; // ruch w całości poza widokiem → snap
    tr.style.transition = "none";
    tr.style.transform = `translateY(${dy}px)`;
    movers.push({ tr, top });
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
  }, maxStagger + 420);
}

let workbook = null;
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
  showCount: 20,
  havingMode: "all",
  havingValue: 10,
  measureFilterMode: "all",
  measureFilterValue: "",
  resultSearch: "",
};
const APP_BUILD_VERSION = "20260628-03";

const THEME_KEY = "excel-workbench-theme";
const MAX_ROWS_KEY = "excel-workbench-max-rows";
const EXCEL_LAYOUT_KEY = "excel-workbench-excel-layout";
const CELL_STYLE_PREFS_KEY = "excel-workbench-cell-style-prefs";
const ROW_HEIGHT_KEY = "excel-workbench-row-height-all";
const COL_WIDTH_KEY = "excel-workbench-col-width-all";
const SORT_PRESETS_KEY = "excel-workbench-sort-presets";
const TOOLBAR_COLLAPSED_KEY = "excel-workbench-toolbar-collapsed";
const INTRO_PLAYED_KEY = "introPlayed";

function log(msg, type = "info") {
  const line = document.createElement("div");
  line.className = `log-line log-${type}`;
  line.textContent = `${new Date().toLocaleTimeString(I18N[currentLang].locale)} ${msg}`;
  logEl.prepend(line);
}

function toast(msg, type = "info") {
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

  setTimeout(() => {
    toastEl.classList.add("out");
    setTimeout(() => toastEl.remove(), 200);
  }, 2800);
}

function setLoading(isLoading, text) {
  loadingTextEl.textContent = text || t("loadingGeneric");
  loadingOverlayEl.classList.toggle("hidden", !isLoading);
}

function setStatus(msg) {
  statusEl.textContent = msg;
}

function setDirtyState(isDirty) {
  hasUnsavedChanges = !!isDirty;
  statusEl.classList.toggle("unsaved", hasUnsavedChanges);
  document.title = hasUnsavedChanges ? `* ${BASE_TITLE}` : BASE_TITLE;
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
  const bottomGap = window.matchMedia("(max-width: 768px)").matches ? 14 : 24;
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
}

