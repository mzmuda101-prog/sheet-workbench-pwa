// Core runtime: DOM refs, shared state, and base UI helpers.

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
const excelLayoutToggleEl = document.getElementById("excelLayoutToggle");
const loadBtn = document.getElementById("loadBtn");

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
const aggregationWorkbenchListEl = document.getElementById("aggregationWorkbenchList");
const formulaSearchEl = document.getElementById("formulaSearch");
const formulaFilterEl = document.getElementById("formulaFilter");
const formulaFunctionFilterEl = document.getElementById("formulaFunctionFilter");
const formulaWorkbenchSummaryEl = document.getElementById("formulaWorkbenchSummary");
const formulaWorkbenchListEl = document.getElementById("formulaWorkbenchList");

const quickRangeButtons = Array.from(document.querySelectorAll(".chip[data-range]"));
const prefersReducedMotion = window.matchMedia("(prefers-reduced-motion: reduce)").matches;

let workbook = null;
let currentHeaders = [];
let baseRows = [];
let viewRows = [];
let matchedRowIndexes = new Set(); // wiersze pasujące do quick search w trybie "zaznacz"
let quickSearchHighlightMode = false; // true = zaznacz zamiast filtruj
let quickSearchOperatorsEnabled = false; // true = &&/|| traktowane jako operatory
let currentFileName = "";
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
let hasUnsavedChanges = false;
let focusedCellState = null;
let selectedCellState = null;
let syncingHorizontalScroll = false;
let tableTouchAxisLock = null;
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
  showCount: 20,
  havingMode: "all",
  havingValue: 10,
  measureFilterMode: "all",
  measureFilterValue: "",
  resultSearch: "",
};
const APP_BUILD_VERSION = "20260529-03";

const THEME_KEY = "excel-workbench-theme";
const MAX_ROWS_KEY = "excel-workbench-max-rows";
const EXCEL_LAYOUT_KEY = "excel-workbench-excel-layout";
const SORT_PRESETS_KEY = "excel-workbench-sort-presets";
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

function syncFrozenHeaderMetrics() {
  if (!tableWrapEl || !theadEl) return;
  const guideRow = theadEl.querySelector(".guide-row");
  const guideHeight = guideRow ? Math.ceil(guideRow.getBoundingClientRect().height) : 0;
  tableWrapEl.style.setProperty("--frozen-guide-height", `${guideHeight}px`);
}

function syncTableViewportHeight() {
  if (!tablePanelEl) return;
  const rect = tablePanelEl.getBoundingClientRect();
  const viewportHeight = window.innerHeight || document.documentElement.clientHeight || 720;
  const bottomGap = window.matchMedia("(max-width: 768px)").matches ? 14 : 24;
  const available = Math.floor(viewportHeight - rect.top - bottomGap);
  const minHeight = window.matchMedia("(max-width: 768px)").matches ? 320 : 420;
  tablePanelEl.style.setProperty("--table-panel-height", `${Math.max(minHeight, available)}px`);
}

function applyFreezeHeaders() {
  if (!tableWrapEl) return;
  const enabled = !freezeHeadersEl || freezeHeadersEl.checked;
  tableWrapEl.classList.toggle("freeze-headers", enabled);
  tableWrapEl.classList.toggle("headers-unlocked", !enabled);
  syncFrozenHeaderMetrics();
}

function startTableTouchAxisLock(event) {
  if (!tableWrapEl || !event.touches || event.touches.length !== 1) {
    tableTouchAxisLock = null;
    return;
  }
  const touch = event.touches[0];
  tableTouchAxisLock = {
    mode: "",
    startX: touch.clientX,
    startY: touch.clientY,
    startScrollLeft: tableWrapEl.scrollLeft,
  };
}

function updateTableTouchAxisLock(event) {
  if (!tableWrapEl || !tableTouchAxisLock || !event.touches || event.touches.length !== 1) return;
  const touch = event.touches[0];
  const dx = touch.clientX - tableTouchAxisLock.startX;
  const dy = touch.clientY - tableTouchAxisLock.startY;
  const absX = Math.abs(dx);
  const absY = Math.abs(dy);

  if (!tableTouchAxisLock.mode && Math.max(absX, absY) > 8) {
    if (absY > absX * 1.35) {
      tableTouchAxisLock.mode = "vertical";
    } else if (absX > absY * 1.15) {
      tableTouchAxisLock.mode = "horizontal";
    }
  }

  if (tableTouchAxisLock.mode === "vertical") {
    tableWrapEl.scrollLeft = tableTouchAxisLock.startScrollLeft;
    requestAnimationFrame(() => {
      if (tableTouchAxisLock?.mode === "vertical") {
        tableWrapEl.scrollLeft = tableTouchAxisLock.startScrollLeft;
      }
    });
  }
}

function endTableTouchAxisLock() {
  tableTouchAxisLock = null;
}
