// UI controls, theming, picker UX, and auxiliary interactions.

function hexToRgba(hex, alpha = 0.35) {
  if (!hex || typeof hex !== "string") return null;
  const m = hex.replace(/^#/, "").replace(/^([A-Fa-f0-9]{6})[A-Fa-f0-9]*$/, "$1").match(/([A-Fa-f0-9]{2})([A-Fa-f0-9]{2})([A-Fa-f0-9]{2})/);
  if (!m) return null;
  return `rgba(${parseInt(m[1], 16)}, ${parseInt(m[2], 16)}, ${parseInt(m[3], 16)}, ${alpha})`;
}

// [EN] Mark rows that look like subheaders (e.g. second header row or important info) in first N data rows
function markSubheaderRows(rows, maxCheck = 10) {
  const toCheck = Math.min(maxCheck, rows.length);
  for (let i = 0; i < toCheck; i++) {
    const row = rows[i];
    let nonEmpty = 0;
    let textLike = 0;
    let numericLike = 0;
    row.values.forEach((v) => {
      if (v != null && String(v).trim() !== "") {
        nonEmpty += 1;
        if (typeof v === "string") textLike += 1;
        else if (typeof v === "number" || v instanceof Date) numericLike += 1;
        else if (!(v instanceof Date) && typeof v !== "number") textLike += 1;
      }
    });
    const n = row.values.length;
    if (n === 0) continue;
    const textRatio = nonEmpty ? textLike / nonEmpty : 0;
    const numericRatio = nonEmpty ? numericLike / nonEmpty : 0;
    if (
      nonEmpty >= 2
      && textRatio >= 0.8
      && numericRatio === 0
      && nonEmpty <= Math.max(6, Math.ceil(n * 0.75))
    ) {
      row.isSubheader = true;
    }
  }
  return rows;
}

function detectHeaderRowSimple(sheet) {
  const range = computeEffectiveSheetRange(sheet, 1);
  const maxRow = Math.min(range.e.r, range.s.r + 100);
  let bestRow = range.s.r;
  let bestScore = -Infinity;
  for (let r = range.s.r; r <= maxRow; r++) {
    let filled = 0;
    let stringCount = 0;
    let numericCount = 0;
    let formulaCount = 0;
    for (let c = range.s.c; c <= range.e.c; c++) {
      const cell = sheet[XLSX.utils.encode_cell({ r, c })];
      if (!cell) continue;
      const v = cell.v;
      if (v === null || v === "") continue;
      filled += 1;
      if (typeof v === "string") stringCount += 1;
      if (typeof v === "number" || v instanceof Date) numericCount += 1;
      if (cell.f) formulaCount += 1;
    }
    if (!filled) continue;
    const textRatio = stringCount / filled;
    const numericRatio = numericCount / filled;
    let score = (filled * 5) + (stringCount * 4) - (numericCount * 3) - formulaCount;
    if (filled >= 4) score += 10;
    if (textRatio >= 0.7) score += 10;
    if (numericRatio === 0) score += 4;
    if (r > range.s.r) score += Math.min(4, r - range.s.r);
    if (score > bestScore) {
      bestScore = score;
      bestRow = r;
    }
  }
  return bestRow + 1;
}

function applyAutoHeaderRowIfEnabled() {
  if (!autoHeaderRowEl || !autoHeaderRowEl.checked) return false;
  if (!workbook) return false;
  const sheetName = sheetSelect.value;
  const sheet = workbook.Sheets[sheetName];
  if (!sheet) return false;
  const detected = detectHeaderRowSimple(sheet);
  headerRowEl.value = String(detected);
  return true;
}

function columnSummary(set) {
  if (!set.size) return t("allColumns");
  if (set.size === 1) return Array.from(set)[0];
  return t("selectedColumnsCount", { count: set.size });
}

function updateColumnSummary() {
  filter1ColumnsEl.value = columnSummary(columnSelections.filter1);
  filter2ColumnsEl.value = columnSummary(columnSelections.filter2);
  dateColumnsEl.value = columnSummary(columnSelections.date);
  updateQuickSearchColumnButtons();
}

function updateFilterBadge() {
  let count = 0;
  if (searchQueryEl.value.trim()) count += 1;
  if (searchQuery2El.value.trim()) count += 1;
  if (getNormalizedSelectValue(filterEmptyModeEl) !== "all") count += 1;
  if (getNormalizedSelectValue(filterEmptyMode2El) !== "all") count += 1;
  if (filterNegateEl.checked) count += 1;
  if (filterNegate2El.checked) count += 1;
  if (filterOperatorsEl?.checked) count += 1;
  if (filterOperators2El?.checked) count += 1;
  if (onlyNonEmptyEl.checked) count += 1;
  if (getNormalizedSelectValue(dateModeEl) === "last_n_days") count += 1;
  if (dateFromEl.value.trim() || dateToEl.value.trim()) count += 1;
  if (getNormalizedSelectValue(dateEmptyModeEl) !== "all") count += 1;
  if (dateNegateEl.checked) count += 1;
  if (columnSelections.filter1.size) count += 1;
  if (columnSelections.filter2.size) count += 1;
  if (columnSelections.date.size) count += 1;

  filterBadgeEl.textContent = String(count);
  filterBadgeEl.classList.toggle("hidden", count === 0);
}

function updateDateChipsActive() {
  const isLastN = getNormalizedSelectValue(dateModeEl) === "last_n_days";
  const days = lastDaysEl.value.trim() ? String(lastDaysEl.value) : "30";
  quickRangeButtons.forEach((btn) => {
    const active = isLastN && btn.dataset.range === days;
    btn.classList.toggle("active", !!active);
  });
}

function isSidebarOpen() {
  return rootEl.classList.contains("sidebar-open");
}

function syncQuickSearchInputs() {
  if (quickSearchEl) quickSearchEl.value = searchQueryEl.value;
  if (quickSearchPopupInput) quickSearchPopupInput.value = searchQueryEl.value;
}

function getQuickSearchModeValue() {
  return filterModeEl && getNormalizedSelectValue(filterModeEl) === "equals" ? "exact" : "contains";
}

function syncQuickSearchModeControls() {
  const mode = getQuickSearchModeValue();
  if (quickSearchModeEl) quickSearchModeEl.value = mode;
  if (quickSearchPopupModeEl) quickSearchPopupModeEl.value = mode;
}

function syncQuickSearchOperatorsControls() {
  const enabled = !!filterOperatorsEl?.checked;
  quickSearchOperatorsEnabled = enabled;
  if (quickSearchOperatorsEl) quickSearchOperatorsEl.checked = enabled;
  if (quickSearchPopupOperatorsEl) quickSearchPopupOperatorsEl.checked = enabled;
}

function applyQuickSearchMode(mode) {
  const normalizedQuickMode = normalizeSelectValue("quickSearchMode", mode);
  const normalized = normalizedQuickMode === "exact" ? "equals" : "contains";
  if (filterModeEl) filterModeEl.value = normalized;
  syncQuickSearchModeControls();
}

function updateQuickSearchColumnButtons() {
  const summary = columnSummary(columnSelections.filter1);
  const count = columnSelections.filter1.size;
  const label = count ? `${t("quickSearchColumns")} (${count})` : t("quickSearchColumns");
  [quickSearchColumnsBtn, quickSearchPopupColumnsBtn].forEach((btn) => {
    if (!btn) return;
    btn.textContent = label;
    btn.title = `${t("quickSearchColumns")}: ${summary}`;
    btn.setAttribute("aria-label", `${t("quickSearchColumns")} - ${summary}.`);
  });
}

function resetFilterInputs() {
  searchQueryEl.value = "";
  searchQuery2El.value = "";
  filterModeEl.value = "contains";
  filterMode2El.value = "contains";
  filterEmptyModeEl.value = "all";
  filterEmptyMode2El.value = "all";
  filterNegateEl.checked = false;
  filterNegate2El.checked = false;
  if (filterOperatorsEl) filterOperatorsEl.checked = false;
  if (filterOperators2El) filterOperators2El.checked = false;
  onlyNonEmptyEl.checked = false;
  dateModeEl.value = "between";
  dateFromEl.value = "";
  dateToEl.value = "";
  lastDaysEl.value = "";
  dateEmptyModeEl.value = "all";
  dateNegateEl.checked = false;
  columnSelections.filter1.clear();
  columnSelections.filter2.clear();
  columnSelections.date.clear();
  quickSearchHighlightMode = false;
  matchedRowIndexes = new Set();
  quickSearchOperatorsEnabled = false;
  if (quickSearchActionEl) quickSearchActionEl.value = "filter";
  if (quickSearchPopupActionEl) quickSearchPopupActionEl.value = "filter";
  if (quickSearchOperatorsEl) quickSearchOperatorsEl.checked = false;
  if (quickSearchPopupOperatorsEl) quickSearchPopupOperatorsEl.checked = false;
  syncQuickSearchInputs();
  syncQuickSearchModeControls();
  syncQuickSearchOperatorsControls();
  updateColumnSummary();
  updateDateChipsActive();
  updateFilterBadge();
  if (typeof setSecondFilterVisible === "function") setSecondFilterVisible(false);
}

function setSidebarOpen(open) {
  const shouldOpen = !!open;
  rootEl.classList.toggle("sidebar-open", shouldOpen);
  if (sidebarScrim) sidebarScrim.classList.toggle("hidden", !shouldOpen);
  if (panelToggle) {
    panelToggle.setAttribute("aria-expanded", shouldOpen ? "true" : "false");
    panelToggle.textContent = shouldOpen ? t("panelOpen") : t("panelClosed");
  }
  requestAnimationFrame(() => syncSidebarHandle());
  window.setTimeout(() => syncSidebarHandle(), 270);
}

function openColumnPicker(key) {
  if (!currentHeaders.length) {
    toast(t("loadSheetToPickColumns"), "info");
    return;
  }
  activePickerKey = key;
  if (columnPickerTitleEl) {
    columnPickerTitleEl.textContent = key === "filter1"
      ? t("quickSearchColumnsTitleShort")
      : key === "filter2"
        ? t("textFilter2ColumnsTitle")
        : t("dateFilterColumnsTitle");
  }
  columnListEl.replaceChildren();
  columnSearchEl.value = "";
  const currentSet = columnSelections[key];
  const isAll = currentSet.size === 0;
  currentHeaders.forEach((h, idx) => {
    const row = document.createElement("div");
    row.className = "field checkbox";
    const input = document.createElement("input");
    input.type = "checkbox";
    input.id = `colpick-${key}-${idx}`;
    input.value = h;
    input.checked = isAll ? true : currentSet.has(h);
    const label = document.createElement("label");
    label.htmlFor = input.id;
    label.textContent = h;
    row.appendChild(input);
    row.appendChild(label);
    columnListEl.appendChild(row);
  });
  columnPickerEl.classList.remove("hidden");
  columnSearchEl.focus();
}

function closeColumnPicker() {
  columnPickerEl.classList.add("hidden");
  if (lastPickerTriggerEl) {
    lastPickerTriggerEl.focus();
    lastPickerTriggerEl = null;
  }
}

function updateGroupByMaxLimit() {
  if (activePickerKey !== "groupby") return;
  const checkboxes = Array.from(columnListEl.querySelectorAll("input[type=checkbox]"));
  const checkedCount = checkboxes.filter((cb) => cb.checked).length;
  checkboxes.forEach((cb) => {
    cb.disabled = !cb.checked && checkedCount >= 3;
  });
}

function openGroupByPicker() {
  if (!currentAggregationGroupOptions || !currentAggregationGroupOptions.length) {
    toast(t("aggregationNoOptions"), "info");
    return;
  }
  activePickerKey = "groupby";
  if (columnPickerTitleEl) {
    columnPickerTitleEl.textContent = t("groupByPickerTitle");
  }
  columnListEl.replaceChildren();
  columnSearchEl.value = "";
  const currentSelected = [
    aggregationWorkbenchState.groupBy,
    aggregationWorkbenchState.groupBy2,
    aggregationWorkbenchState.groupBy3,
  ].filter(Boolean);
  currentAggregationGroupOptions.forEach((option, idx) => {
    if (!option.value) return;
    const row = document.createElement("div");
    row.className = "field checkbox";
    const input = document.createElement("input");
    input.type = "checkbox";
    input.id = `groupbypick-${idx}`;
    input.value = option.value;
    input.checked = currentSelected.includes(option.value);
    if (currentSelected.length >= 3 && !input.checked) {
      input.disabled = true;
    }
    const label = document.createElement("label");
    label.htmlFor = input.id;
    label.textContent = option.label;
    row.appendChild(input);
    row.appendChild(label);
    columnListEl.appendChild(row);
  });
  columnPickerEl.classList.remove("hidden");
  columnSearchEl.focus();
}

function openMeasurePicker() {
  if (!currentAggregationMeasureCandidates || !currentAggregationMeasureCandidates.length) {
    toast(t("aggregationNoOptions"), "info");
    return;
  }
  activePickerKey = "measures";
  if (columnPickerTitleEl) {
    columnPickerTitleEl.textContent = t("selectMeasuresTitle");
  }
  columnListEl.replaceChildren();
  columnSearchEl.value = "";
  const currentSet = new Set(aggregationWorkbenchState.measures || []);
  currentAggregationMeasureCandidates.forEach((candidate, idx) => {
    const row = document.createElement("div");
    row.className = "field checkbox";
    const input = document.createElement("input");
    input.type = "checkbox";
    input.id = `measurepick-${idx}`;
    input.value = candidate.key;
    input.checked = currentSet.has(candidate.key);
    // count_rows (Liczebność) wzajemnie wyklucza się z miarami kolumnowymi:
    // silnik agregacji i tak nadpisuje pozostałe miary, gdy obecne jest count_rows,
    // więc nie pozwalamy zmieszać ich w jednym wyborze.
    input.addEventListener("change", () => {
      if (!input.checked) return;
      if (candidate.key === "count_rows") {
        columnListEl.querySelectorAll("input[type=checkbox]").forEach((cb) => {
          if (cb !== input) cb.checked = false;
        });
      } else {
        const countRowsCb = columnListEl.querySelector('input[value="count_rows"]');
        if (countRowsCb) countRowsCb.checked = false;
      }
    });
    const label = document.createElement("label");
    label.htmlFor = input.id;
    label.textContent = candidate.label;
    row.appendChild(input);
    row.appendChild(label);
    columnListEl.appendChild(row);
  });
  columnPickerEl.classList.remove("hidden");
  columnSearchEl.focus();
}

function getModalFocusables() {
  const modalContent = columnPickerEl.querySelector(".modal-content");
  if (!modalContent) return [];
  const all = Array.from(modalContent.querySelectorAll("button, input:not([type=hidden]), [tabindex]:not([tabindex^='-'])"));
  return all.filter((el) => {
    const row = el.closest(".field.checkbox");
    return !row || !row.classList.contains("hidden");
  });
}

function handlePickerKeydown(e) {
  if (columnPickerEl.classList.contains("hidden")) return;
  if (e.key === "Tab") {
    const focusables = getModalFocusables();
    if (focusables.length === 0) return;
    const idx = focusables.indexOf(document.activeElement);
    if (idx === -1) return;
    if (e.shiftKey && idx === 0) {
      e.preventDefault();
      focusables[focusables.length - 1].focus();
    } else if (!e.shiftKey && idx === focusables.length - 1) {
      e.preventDefault();
      focusables[0].focus();
    }
  }
}

function filterColumnList() {
  const q = columnSearchEl.value.trim().toLowerCase();
  columnListEl.querySelectorAll(".field.checkbox").forEach((row) => {
    const text = row.textContent.toLowerCase();
    row.classList.toggle("hidden", q && !text.includes(q));
  });
}

columnListEl.addEventListener("click", (e) => {
  const row = e.target.closest(".field.checkbox");
  if (!row) return;
  const cb = row.querySelector("input[type=checkbox]");
  if (!cb || cb.disabled) return;
  if (e.target === cb || e.target.tagName === "LABEL") return;
  cb.checked = !cb.checked;
  cb.dispatchEvent(new Event("change", { bubbles: true }));
});

columnListEl.addEventListener("change", () => {
  updateGroupByMaxLimit();
});


function attachResizeHandlers() {
  let active = null;
  let startX = 0;
  let startW = 0;
  let rafId = null;

  const start = (e) => {
    const handle = e.target.closest(".col-resizer");
    if (!handle) return;
    e.preventDefault();
    const colIndex = parseInt(handle.dataset.colIndex, 10);
    const th = handle.parentElement;
    active = { colIndex, th };
    startX = e.clientX || (e.touches && e.touches[0].clientX) || 0;
    startW = th.getBoundingClientRect().width;
    document.body.classList.add("resizing");
  };

  const move = (e) => {
    if (!active) return;
    const x = e.clientX || (e.touches && e.touches[0].clientX) || 0;
    const delta = x - startX;
    const next = Math.max(80, Math.min(520, Math.round(startW + delta)));
    manualColumnWidths[active.colIndex] = next;
    if (rafId) return;
    rafId = requestAnimationFrame(() => {
      rafId = null;
      if (active) renderActiveTable();
    });
  };

  const stop = () => {
    if (!active) return;
    active = null;
    if (rafId) { cancelAnimationFrame(rafId); rafId = null; }
    document.body.classList.remove("resizing");
  };

  tableEl.addEventListener("mousedown", start);
  tableEl.addEventListener("touchstart", start, { passive: true });
  window.addEventListener("mousemove", move);
  window.addEventListener("touchmove", move, { passive: true });
  window.addEventListener("mouseup", stop);
  window.addEventListener("touchend", stop);
}

function initTheme() {
  const saved = localStorage.getItem(THEME_KEY);
  const prefersDark = window.matchMedia && window.matchMedia("(prefers-color-scheme: dark)").matches;
  const theme = saved || (prefersDark ? "dark" : "light");
  setTheme(theme, false);
}

function initIntroSplash() {
  const splash = document.getElementById("heroSplash");
  const vid = document.getElementById("introVideo");
  if (!splash) return;

  if (sessionStorage.getItem(INTRO_PLAYED_KEY)) {
    splash.style.display = "none";
    document.body.classList.remove("splashing");
    return;
  }

  document.body.classList.add("splashing");

  const hideSplash = () => {
    if (!splash || splash.classList.contains("hide")) return;
    splash.classList.add("hide");
    sessionStorage.setItem(INTRO_PLAYED_KEY, "true");
    setTimeout(() => {
      splash.style.display = "none";
      document.body.classList.remove("splashing");
    }, 700);
  };

  if (vid) {
    try {
      vid.currentTime = 0;
      vid.muted = true;
      vid.playbackRate = 1.5;
      const playPromise = vid.play();
      if (playPromise && typeof playPromise.catch === "function") {
        playPromise.catch(() => hideSplash());
      }
    } catch {
      hideSplash();
    }

    const fallback = setTimeout(hideSplash, 10000);
    vid.addEventListener("ended", () => {
      clearTimeout(fallback);
      hideSplash();
    });
  } else {
    setTimeout(hideSplash, 6000);
  }
}

function createThemeIcon(isDark) {
  const svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
  svg.setAttribute("width", "18");
  svg.setAttribute("height", "18");
  svg.setAttribute("viewBox", "0 0 24 24");
  svg.setAttribute("fill", "none");
  svg.setAttribute("stroke", "currentColor");
  svg.setAttribute("stroke-width", "2");
  svg.setAttribute("stroke-linecap", "round");
  svg.setAttribute("stroke-linejoin", "round");

  if (isDark) {
    // Sun icon for dark mode active (click to switch to light)
    const circle = document.createElementNS("http://www.w3.org/2000/svg", "circle");
    circle.setAttribute("cx", "12");
    circle.setAttribute("cy", "12");
    circle.setAttribute("r", "5");
    svg.appendChild(circle);

    const rays = [
      { x1: "12", y1: "1", x2: "12", y2: "3" },
      { x1: "12", y1: "21", x2: "12", y2: "23" },
      { x1: "4.22", y1: "4.22", x2: "5.64", y2: "5.64" },
      { x1: "18.36", y1: "18.36", x2: "19.78", y2: "19.78" },
      { x1: "1", y1: "12", x2: "3", y2: "12" },
      { x1: "21", y1: "12", x2: "23", y2: "12" },
      { x1: "4.22", y1: "19.78", x2: "5.64", y2: "18.36" },
      { x1: "18.36", y1: "5.64", x2: "19.78", y2: "4.22" },
    ];

    rays.forEach(ray => {
      const line = document.createElementNS("http://www.w3.org/2000/svg", "line");
      Object.entries(ray).forEach(([attr, val]) => line.setAttribute(attr, val));
      svg.appendChild(line);
    });
  } else {
    // Moon icon for light mode active (click to switch to dark)
    const path = document.createElementNS("http://www.w3.org/2000/svg", "path");
    path.setAttribute("d", "M21 12.79A9 9 0 1 1 11.21 3a7 7 0 0 0 9.79 9.79z");
    svg.appendChild(path);
  }

  return svg;
}

function setTheme(theme, persist = true) {
  rootEl.setAttribute("data-theme", theme);
  themeToggle.setAttribute("aria-pressed", theme === "dark" ? "true" : "false");
  if (persist) localStorage.setItem(THEME_KEY, theme);
  
  // Clear existing content and append safe SVG element
  themeToggle.textContent = "";
  themeToggle.appendChild(createThemeIcon(theme === "dark"));
}

function updateNetworkBadge() {
  if (!networkBadgeEl) return;
  const isOnline = navigator.onLine;
  networkBadgeEl.textContent = isOnline ? t("online") : t("offline");
  networkBadgeEl.classList.toggle("offline", !isOnline);
  const safetyNote = t("networkSafety");
  networkBadgeEl.setAttribute(
    "title",
    isOnline ? t("networkOnlineTitle", { note: safetyNote }) : t("networkOfflineTitle", { note: safetyNote })
  );
}

async function hardRefreshApp() {
  try {
    if ("serviceWorker" in navigator) {
      const registrations = await navigator.serviceWorker.getRegistrations();
      await Promise.all(registrations.map((registration) => registration.update().catch(() => {})));
    }

    if ("caches" in window) {
      const keys = await caches.keys();
      const appKeys = keys.filter((key) => key.startsWith("excel-wb-"));
      await Promise.all(appKeys.map((key) => caches.delete(key).catch(() => false)));
    }

    toast(t("cacheRefresh"), "info");
  } catch {
    toast(t("refreshingApp"), "info");
  }

  window.location.reload();
}

function formatFileSize(bytes) {
  if (bytes >= 1048576) return `${(bytes / 1048576).toFixed(1)} MB`;
  if (bytes >= 1024) return `${Math.round(bytes / 1024)} KB`;
  return `${bytes} B`;
}

async function handleFile(file) {
  if (!file) return;
  if (!isXlsxAvailable(true)) return;
  try {
    const sizeHint = file.size > 0 ? ` (${formatFileSize(file.size)})` : "";
    setLoading(true, t("loadingFile") + sizeHint);
    const data = await file.arrayBuffer();
    try {
      workbook = XLSX.read(data, { cellDates: true, cellStyles: true });
    } catch {
      workbook = XLSX.read(data, { cellDates: true });
    }
    sheetSelect.replaceChildren();
    workbook.SheetNames.forEach((s) => {
      const opt = document.createElement("option");
      opt.value = s;
      opt.textContent = s;
    sheetSelect.appendChild(opt);
  });
    currentWorkbookStats = collectWorkbookStats(workbook, file.name);
    currentSheetStats = null;
    currentKpiEntries = [];
    currentKpiAnchorRow = 1;
    currentColumnProfiles = [];
    currentSections = [];
    currentRepeatingBlocks = [];
    currentDisplayModel = null;
    tableViewMode = "wide";
    multiSortState = [];
    sortState = { col: "", dir: "asc" };
    currentFileName = file.name;
    currentStartCol = 0;
    currentMerges = [];
    currentHeaderStyles = [];
    currentSheetColWidths = [];
    currentSheetRowHeights = {};
    fileNameTextEl.textContent = file.name;
    fileNameEl.classList.remove("hidden");
    dropZone.classList.add("has-file");
    setDirtyState(false);
    setStatus(t("statusFileLoaded"));
    renderInsights();
    renderKpiExtractor();
    renderColumnProfiles();
    renderSections();
    renderRepeatingBlocks();
    renderDurationAnalysis();
    populateSortColumnSelect();
    renderSortPresets();
    toast(t("fileLoaded"), "success");
    log(`Wczytano plik: ${file.name}`, "success");
  } catch (err) {
    toast(t("fileLoadFailed"), "error");
    log("Blad przy wczytywaniu pliku.", "error");
  } finally {
    setLoading(false);
  }
}

// Przykładowy arkusz procesowy/SLA generowany lokalnie w pamięci (xlsx-js-style).
// Pokazuje: powtarzalne bloki (od/do/Długość ×3) → Wide-to-Long, analizę czasu,
// agregacje, filtry dat oraz KPI (wiersze podsumowania nad nagłówkiem).
// Lekki (kilkanaście wierszy) — generacja jest natychmiastowa, nic nie jest wysyłane.
function buildSampleWorkbookArrayBuffer() {
  const owners = ["Anna Kowalska", "Jan Nowak", "Piotr Wiśniewski", "Maria Wójcik", "Tomasz Lewandowski"];
  const statuses = ["Zamknięte", "W trakcie", "PRZETERMINOWANY"];
  const base = new Date(2026, 0, 6);
  const addDays = (date, n) => new Date(date.getFullYear(), date.getMonth(), date.getDate() + n);
  const daysBetween = (a, b) => Math.round((b - a) / 86400000);
  // Daty zapisujemy jako stringi ISO (YYYY-MM-DD): ten build xlsx-js-style gubi
  // gołe obiekty Date przy zapisie, a parseDateFlexible aplikacji i tak parsuje ISO.
  const iso = (d) => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;

  const headerRow = ["Teren", "Opiekun", "Status", "od", "do", "Długość", "od2", "do2", "Długość2", "od3", "do3", "Długość3"];
  const rows = [];
  let closed = 0, inProgress = 0, overdue = 0;

  for (let i = 0; i < 16; i++) {
    const teren = `Teren ${String.fromCharCode(65 + (i % 4))}-${String(i + 1).padStart(2, "0")}`;
    const opiekun = owners[i % owners.length];
    const status = statuses[i % 3];
    if (status === "Zamknięte") closed += 1;
    else if (status === "W trakcie") inProgress += 1;
    else overdue += 1;

    // Cykl 1 — zawsze zamknięty
    const od1 = addDays(base, i * 6);
    const do1 = addDays(od1, 9 + (i % 5) * 3);
    // Cykl 2 — obecny dla ~2/3 wierszy
    const hasC2 = i % 3 !== 2;
    const od2 = hasC2 ? addDays(do1, 4) : null;
    const do2 = hasC2 ? addDays(od2, 7 + (i % 4) * 2) : null;
    // Cykl 3 — dla ~1/3 wierszy, część otwarta (brak "do" → liczone do dzisiaj)
    const hasC3 = i % 3 === 0;
    const startC3Base = hasC2 ? do2 : do1;
    const od3 = hasC3 ? addDays(startC3Base, 5) : null;
    const open3 = hasC3 && i % 2 === 0;
    const do3 = hasC3 && !open3 ? addDays(od3, 11 + (i % 3) * 4) : null;

    rows.push([
      teren, opiekun, status,
      iso(od1), iso(do1), daysBetween(od1, do1),
      od2 ? iso(od2) : "", do2 ? iso(do2) : "", hasC2 ? daysBetween(od2, do2) : "",
      od3 ? iso(od3) : "", do3 ? iso(do3) : "", (hasC3 && !open3) ? daysBetween(od3, do3) : "",
    ]);
  }

  const aoa = [
    ["Raport: Obieg terenów 2026"],
    ["Terenów łącznie", rows.length, "", "Zamknięte", closed, "", "W trakcie", inProgress, "", "Przeterminowane", overdue],
    [],
    headerRow,
    ...rows,
  ];

  const ws = XLSX.utils.aoa_to_sheet(aoa, { cellDates: true });
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Obieg terenów");
  return XLSX.write(wb, { bookType: "xlsx", type: "array" });
}

async function loadSampleFile() {
  if (!isXlsxAvailable(true)) return;
  let buffer;
  try {
    buffer = buildSampleWorkbookArrayBuffer();
  } catch (err) {
    toast(t("sampleLoadFailed"), "error");
    log("Blad przy generowaniu przykladowego pliku.", "error");
    return;
  }
  const file = new File([buffer], t("sampleFileName"), {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  await handleFile(file);
  // Auto-wykryj wiersz nagłówka (pomija wiersze KPI) i od razu wczytaj arkusz do tabeli.
  if (workbook) {
    if (autoHeaderRowEl) autoHeaderRowEl.checked = true;
    applyAutoHeaderRowIfEnabled();
    loadBtn.click();
  }
}

function escapeCsv(value) {
  const raw = String(value ?? "");
  if (raw.includes("\"") || raw.includes(",") || raw.includes("\n")) {
    return `"${raw.replace(/\"/g, '""')}"`;
  }
  return raw;
}

function exportCsv() {
  const model = currentDisplayModel || getDisplayModel();
  if (!model.headers.length || !model.rows.length) {
    toast(t("noDataForExport"), "warning");
    return;
  }
  const rows = [
    model.headers,
    ...model.rows.map((row) => row.values.map((v, i) => getDisplayValue(row, i))),
  ];
  const csv = rows.map((row) => row.map(escapeCsv).join(",")).join("\n");
  const base = currentFileName ? currentFileName.replace(/\.[^.]+$/, "") : "excel-workbench";
  const sheet = sheetSelect.value ? sheetSelect.value.replace(/\s+/g, "_") : "arkusz";
  const suffix = model.mode === "long" ? "long" : "wide";
  const filename = `${base}_${sheet}_${suffix}.csv`;
  const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
  toast(t("csvExported"), "success");
}

function saveWorkbook() {
  if (!isXlsxAvailable(true)) return;
  if (!workbook) {
    toast(t("noFileToSave"), "warning");
    return;
  }
  const base = currentFileName ? currentFileName.replace(/\.[^.]+$/, "") : "excel-workbench";
  const ext = currentFileName && currentFileName.toLowerCase().endsWith(".xlsm") ? "xlsm" : "xlsx";
  if (ext === "xlsm") {
    const ok = window.confirm(t("xlsmConfirm"));
    if (!ok) return;
  }
  const filename = `${base}_edited.${ext}`;
  XLSX.writeFile(workbook, filename, { bookType: ext });
  setDirtyState(false);
  toast(t("fileSaved"), "success");
  log(`Zapisano plik: ${filename}`, "success");
}

function saveWorkbookAs() {
  if (!isXlsxAvailable(true)) return;
  if (!workbook) {
    toast(t("noFileToSave"), "warning");
    return;
  }
  const base = currentFileName ? currentFileName.replace(/\.[^.]+$/, "") : "excel-workbench";
  const suggested = `${base}_edited.xlsx`;
  const nameRaw = window.prompt(t("saveAsPrompt"), suggested);
  if (!nameRaw) return;
  let name = nameRaw.trim();
  if (!name) return;
  if (!/\.(xlsx|xlsm)$/i.test(name)) {
    name = `${name}.xlsx`;
  }
  const ext = name.toLowerCase().endsWith(".xlsm") ? "xlsm" : "xlsx";
  if (ext === "xlsm") {
    const ok = window.confirm(t("xlsmConfirm"));
    if (!ok) return;
  }
  XLSX.writeFile(workbook, name, { bookType: ext });
  setDirtyState(false);
  toast(t("fileSaved"), "success");
  log(`Zapisano plik: ${name}`, "success");
}

fileInput.addEventListener("change", (e) => {
  const file = e.target.files[0];
  handleFile(file);
});

if (loadSampleBtn) {
  loadSampleBtn.addEventListener("click", loadSampleFile);
}

// Drugi filtr tekstowy: domyślnie ukryty, odsłaniany przyciskiem "+ Dodaj drugi filtr".
const addFilter2Btn = document.getElementById("addFilter2Btn");
const removeFilter2Btn = document.getElementById("removeFilter2Btn");
const filter2Block = document.getElementById("filter2Block");
function setSecondFilterVisible(show) {
  if (filter2Block) filter2Block.classList.toggle("hidden", !show);
  if (addFilter2Btn) addFilter2Btn.classList.toggle("hidden", show);
}
if (addFilter2Btn) {
  addFilter2Btn.addEventListener("click", () => {
    setSecondFilterVisible(true);
    if (searchQuery2El) searchQuery2El.focus();
  });
}
if (removeFilter2Btn) {
  removeFilter2Btn.addEventListener("click", () => {
    if (searchQuery2El) searchQuery2El.value = "";
    if (filterMode2El) filterMode2El.value = "contains";
    if (filterEmptyMode2El) filterEmptyMode2El.value = "all";
    if (filterNegate2El) filterNegate2El.checked = false;
    if (filterOperators2El) filterOperators2El.checked = false;
    columnSelections.filter2.clear();
    updateColumnSummary();
    updateFilterBadge();
    setSecondFilterVisible(false);
  });
}

loadBtn.addEventListener("click", () => {
  if (!isXlsxAvailable(true)) return;
  if (!workbook) {
    toast(t("chooseFileFirst"), "warning");
    log("Najpierw wybierz plik.", "warn");
    return;
  }
  setLoading(true, t("loadingSheet"));
  setTimeout(() => {
    try {
      const sheetName = sheetSelect.value;
      const sheet = workbook.Sheets[sheetName];
      if (!sheet) {
        toast(t("noSheet"), "error");
        log("Brak arkusza.", "error");
        return;
      }
      applyAutoHeaderRowIfEnabled();
      const headerRow = Math.max(1, parseInt(headerRowEl.value || "1", 10));
      currentHeaderRow = headerRow;
      if (!Number.isFinite(aggregationWorkbenchState.customHeaderRow) || aggregationWorkbenchState.customHeaderRow < 1 || aggregationWorkbenchState.headerRowChoice === "auto") {
        aggregationWorkbenchState.customHeaderRow = headerRow;
      }
      currentSheetName = sheetName;
      const data = buildRows(sheet, headerRow, workbook);
      currentHeaders = data.headers;
      currentStartCol = data.startCol || 0;
      currentMerges = Array.isArray(data.merges) ? data.merges : [];
      currentHeaderStyles = Array.isArray(data.headerStyles) ? data.headerStyles : [];
      currentSheetColWidths = Array.isArray(data.colWidths) ? data.colWidths : [];
      currentSheetRowHeights = data.rowHeights || {};
      currentSheetStats = data.stats || null;
      baseRows = markSubheaderRows(data.rows);
      const kpiData = collectKpiEntries(sheet, headerRow);
      currentKpiEntries = Array.isArray(kpiData?.entries) ? kpiData.entries : [];
      currentKpiAnchorRow = Number(kpiData?.anchorRow) || headerRow;
      currentColumnProfiles = collectColumnProfiles();
      currentSections = detectSections(sheet, headerRow, data);
      currentRepeatingBlocks = detectRepeatingBlocks(sheet, headerRow, data);
      currentFormulaEntries = collectFormulaEntries(sheet, data, headerRow);
      if (!canUseLongView()) tableViewMode = "wide";
      viewRows = baseRows.slice();
      multiSortState = [];
      sortState = { col: "", dir: "asc" };
      manualColumnWidths = {};
      columnSelections.filter1.clear();
      columnSelections.filter2.clear();
      columnSelections.date.clear();
      updateColumnSummary();
      updateFilterBadge();
      populateSortColumnSelect();
      renderActiveTable();
      renderInsights();
      renderKpiExtractor();
      renderSheetInspectorSummary();
      renderColumnProfiles();
      renderSections();
      renderRepeatingBlocks();
      renderDurationAnalysis();
      renderAggregationWorkbench();
      renderFormulaWorkbench();
      setDirtyState(false);
      if ((currentSheetStats?.trimmedColumns || 0) > 0) {
        log(`Przycięto puste kolumny poza realnym zakresem danych: ${currentSheetStats.trimmedColumns}`, "info");
      }
      if (currentSheetStats?.duplicateHeaderCount) {
        toast(t("duplicatedHeaders", { count: currentSheetStats.duplicateHeaderCount }), "warning");
      }
      toast(t("sheetLoaded"), "success");
      log(`Wczytano arkusz: ${sheetName}`, "success");
      setTimeout(() => {
        const panelFileSheet = document.getElementById("panel-file-sheet");
        if (panelFileSheet) panelFileSheet.removeAttribute("open");
      }, 100);
    } finally {
      setLoading(false);
    }
  }, 50);
});

applyFilterBtn.addEventListener("click", () => {
  if (!currentHeaders.length) return;
  applyFilters();
  sortRows();
  renderActiveTable();
  renderInsights();
  renderKpiExtractor();
  renderSheetInspectorSummary();
  renderColumnProfiles();
  renderSections();
  renderRepeatingBlocks();
  renderDurationAnalysis();
  renderAggregationWorkbench();
  updateFilterBadge();
  toast(t("filtersApplied"), "info");
});

function applyQuickSearch() {
  if (!currentHeaders.length) return;
  let value = "";
  if (quickSearchPopupEl && !quickSearchPopupEl.classList.contains("hidden") && quickSearchPopupInput) value = quickSearchPopupInput.value;
  else if (quickSearchEl) value = quickSearchEl.value;
  else value = searchQueryEl.value || "";
  const popupActive = quickSearchPopupEl && !quickSearchPopupEl.classList.contains("hidden");
  if (popupActive && quickSearchPopupModeEl) applyQuickSearchMode(getNormalizedSelectValue(quickSearchPopupModeEl));
  else if (quickSearchModeEl) applyQuickSearchMode(getNormalizedSelectValue(quickSearchModeEl));

  // Odczytaj tryb akcji (filtruj / zaznacz)
  const actionEl = (popupActive && quickSearchPopupActionEl) ? quickSearchPopupActionEl : quickSearchActionEl;
  quickSearchHighlightMode = actionEl ? actionEl.value === "highlight" : false;

  // Odczytaj i synchronizuj checkbox operatorów
  const operatorsEl = (popupActive && quickSearchPopupOperatorsEl) ? quickSearchPopupOperatorsEl : quickSearchOperatorsEl;
  quickSearchOperatorsEnabled = operatorsEl ? operatorsEl.checked : false;
  if (filterOperatorsEl) filterOperatorsEl.checked = quickSearchOperatorsEnabled;
  if (quickSearchOperatorsEl) quickSearchOperatorsEl.checked = quickSearchOperatorsEnabled;
  if (quickSearchPopupOperatorsEl) quickSearchPopupOperatorsEl.checked = quickSearchOperatorsEnabled;

  // Synchronizuj oba selecty akcji
  if (quickSearchActionEl && actionEl) quickSearchActionEl.value = actionEl.value;
  if (quickSearchPopupActionEl && actionEl) quickSearchPopupActionEl.value = actionEl.value;

  if (quickSearchPopupInput) quickSearchPopupInput.value = value;
  if (quickSearchEl) quickSearchEl.value = value;
  searchQueryEl.value = value;
  applyFilters();
  sortRows();
  renderActiveTable();
  renderInsights();
  renderSheetInspectorSummary();
  renderColumnProfiles();
  renderSections();
  renderRepeatingBlocks();
  renderDurationAnalysis();
  renderAggregationWorkbench();
  updateFilterBadge();
  if (quickSearchPopupEl && !quickSearchPopupEl.classList.contains("hidden")) {
    quickSearchPopupEl.classList.add("hidden");
  }
}

if (tableWrapEl && tableScrollbarEl) {
  tableWrapEl.addEventListener("touchstart", startTableTouchAxisLock, { passive: true });
  tableWrapEl.addEventListener("touchmove", updateTableTouchAxisLock, { passive: true });
  tableWrapEl.addEventListener("touchend", endTableTouchAxisLock, { passive: true });
  tableWrapEl.addEventListener("touchcancel", endTableTouchAxisLock, { passive: true });

  tableWrapEl.addEventListener("scroll", () => {
    hideCellTooltip();
    if (syncingHorizontalScroll) return;
    syncingHorizontalScroll = true;
    tableScrollbarEl.scrollLeft = tableWrapEl.scrollLeft;
    requestAnimationFrame(() => {
      syncingHorizontalScroll = false;
    });
  }, { passive: true });

  tableScrollbarEl.addEventListener("scroll", () => {
    if (syncingHorizontalScroll) return;
    syncingHorizontalScroll = true;
    tableWrapEl.scrollLeft = tableScrollbarEl.scrollLeft;
    requestAnimationFrame(() => {
      syncingHorizontalScroll = false;
    });
  }, { passive: true });
}

tbodyEl.addEventListener("pointerenter", (e) => {
  const td = e.target.closest("td");
  if (!td || td.classList.contains("row-head")) return;
  showCellTooltip(td);
}, true);

tbodyEl.addEventListener("pointerleave", (e) => {
  const td = e.target.closest("td");
  if (!td) return;
  hideCellTooltip();
}, true);

tbodyEl.addEventListener("touchstart", (e) => {
  const td = e.target.closest("td");
  if (!td || td.classList.contains("row-head")) return;
  showCellTooltip(td, true);
}, { passive: true });

window.addEventListener("resize", () => {
  syncHorizontalScrollbar();
  hideCellTooltip();
  syncSidebarHandle();
});

if (quickSearchBtn) {
  quickSearchBtn.addEventListener("click", applyQuickSearch);
}

if (quickSearchColumnsBtn) {
  quickSearchColumnsBtn.addEventListener("click", () => {
    lastPickerTriggerEl = quickSearchColumnsBtn;
    openColumnPicker("filter1");
  });
}

if (quickSearchEl) {
  quickSearchEl.addEventListener("keydown", (e) => {
    if (e.key === "Enter") applyQuickSearch();
  });
}

if (quickSearchModeEl) {
  quickSearchModeEl.addEventListener("change", () => {
    applyQuickSearchMode(getNormalizedSelectValue(quickSearchModeEl));
  });
}

if (quickSearchPopupInput) {
  quickSearchPopupInput.addEventListener("keydown", (e) => {
    if (e.key === "Enter") { e.preventDefault(); applyQuickSearch(); }
  });
}
if (quickSearchPopupModeEl) {
  quickSearchPopupModeEl.addEventListener("change", () => {
    applyQuickSearchMode(getNormalizedSelectValue(quickSearchPopupModeEl));
  });
}
if (quickSearchPopupBtn) {
  quickSearchPopupBtn.addEventListener("click", applyQuickSearch);
}
if (quickSearchPopupColumnsBtn) {
  quickSearchPopupColumnsBtn.addEventListener("click", () => {
    lastPickerTriggerEl = quickSearchPopupColumnsBtn;
    openColumnPicker("filter1");
  });
}
if (quickSearchPopupEl) {
  quickSearchPopupEl.addEventListener("click", (e) => {
    if (e.target === quickSearchPopupEl) quickSearchPopupEl.classList.add("hidden");
  });
}


resetFiltersBtn.addEventListener("click", () => {
  resetFilterInputs();
  viewRows = baseRows.slice();
  sortRows();
  renderActiveTable();
  renderInsights();
  renderColumnProfiles();
  renderSections();
  renderRepeatingBlocks();
  renderDurationAnalysis();
  renderAggregationWorkbench();
  toast(t("filtersReset"), "info");
});

filter1PickBtn.addEventListener("click", () => {
  lastPickerTriggerEl = filter1PickBtn;
  openColumnPicker("filter1");
});
filter2PickBtn.addEventListener("click", () => {
  lastPickerTriggerEl = filter2PickBtn;
  openColumnPicker("filter2");
});
datePickBtn.addEventListener("click", () => {
  lastPickerTriggerEl = datePickBtn;
  openColumnPicker("date");
});

quickRangeButtons.forEach((btn) => {
  btn.addEventListener("click", () => {
    const days = parseInt(btn.dataset.range || "30", 10);
    dateModeEl.value = "last_n_days";
    lastDaysEl.value = String(days);
    updateDateChipsActive();
    applyFilters();
    sortRows();
    renderActiveTable();
    updateFilterBadge();
  });
});

selectAllBtn.addEventListener("click", () => {
  if (activePickerKey === "groupby") {
    toast(t("groupByLimitReached"), "info");
    return;
  }
  columnListEl.querySelectorAll("input[type=checkbox]").forEach((cb) => {
    cb.checked = activePickerKey === "measures" && cb.value === "count_rows" ? false : true;
  });
});

clearAllBtn.addEventListener("click", () => {
  columnListEl.querySelectorAll("input[type=checkbox]").forEach((cb) => {
    cb.checked = false;
  });
});

applyPickBtn.addEventListener("click", () => {
  if (!activePickerKey) return;
  const checked = Array.from(columnListEl.querySelectorAll("input[type=checkbox]"))
    .filter((cb) => cb.checked)
    .map((cb) => cb.value);
  if (activePickerKey === "groupby") {
    const selected = checked.slice(0, 3);
    aggregationWorkbenchState.groupBy = selected[0] || "";
    aggregationWorkbenchState.groupBy2 = selected[1] || "";
    aggregationWorkbenchState.groupBy3 = selected[2] || "";
    renderAggregationWorkbench();
    closeColumnPicker();
    return;
  }
  if (activePickerKey === "measures") {
    aggregationWorkbenchState.measures = checked.length ? checked : ["count_rows"];
    renderAggregationWorkbench();
    closeColumnPicker();
    return;
  }
  if (checked.length === currentHeaders.length) {
    columnSelections[activePickerKey].clear();
  } else {
    columnSelections[activePickerKey] = new Set(checked);
  }
  updateColumnSummary();
  updateFilterBadge();
  closeColumnPicker();
});

if (addSortRuleBtn) {
  addSortRuleBtn.addEventListener("click", () => {
    if (!currentHeaders.length) {
      toast(t("firstLoadSheet"), "info");
      return;
    }
    const col = sortColumnSelectEl?.value;
    const dir = sortDirectionSelectEl?.value === "desc" ? "desc" : "asc";
    if (!col) return;
    multiSortState = multiSortState.filter((rule) => rule.col !== col);
    multiSortState.push({ col, dir });
    normalizeSortState();
    applyCurrentSort();
    toast(t("addedSortRule"), "info");
  });
}

if (sortRulesListEl) {
  sortRulesListEl.addEventListener("click", (e) => {
    const btn = e.target.closest("button[data-sort-action]");
    if (!btn) return;
    const action = btn.dataset.sortAction;
    const index = parseInt(btn.dataset.sortIndex || "", 10);
    if (!Number.isFinite(index) || index < 0 || index >= multiSortState.length) return;

    if (action === "remove") {
      multiSortState.splice(index, 1);
    } else if (action === "toggle") {
      multiSortState[index].dir = multiSortState[index].dir === "asc" ? "desc" : "asc";
    } else if (action === "up" && index > 0) {
      [multiSortState[index - 1], multiSortState[index]] = [multiSortState[index], multiSortState[index - 1]];
    } else if (action === "down" && index < multiSortState.length - 1) {
      [multiSortState[index + 1], multiSortState[index]] = [multiSortState[index], multiSortState[index + 1]];
    }

    normalizeSortState();
    applyCurrentSort();
  });
}

if (saveSortPresetBtn) {
  saveSortPresetBtn.addEventListener("click", () => {
    normalizeSortState();
    if (!multiSortState.length) {
      toast(t("noSortsToSave"), "warning");
      return;
    }
    const name = window.prompt(t("presetNamePrompt"), "");
    if (!name || !name.trim()) return;
    const trimmed = name.trim();
    const presets = loadSortPresets().filter((preset) => preset.name !== trimmed);
    presets.push({ name: trimmed, rules: multiSortState.map((rule) => ({ ...rule })) });
    presets.sort((a, b) => a.name.localeCompare(b.name, "pl"));
    saveSortPresets(presets);
    renderSortPresets();
    if (sortPresetSelectEl) sortPresetSelectEl.value = trimmed;
    toast(t("sortPresetSaved"), "success");
  });
}

if (applySortPresetBtn) {
  applySortPresetBtn.addEventListener("click", () => {
    const name = sortPresetSelectEl?.value;
    if (!name) {
      toast(t("choosePresetToast"), "info");
      return;
    }
    const preset = loadSortPresets().find((item) => item.name === name);
    if (!preset) {
      toast(t("presetNotFound"), "warning");
      renderSortPresets();
      return;
    }
    multiSortState = Array.isArray(preset.rules) ? preset.rules.map((rule) => ({ col: rule.col, dir: rule.dir })) : [];
    normalizeSortState();
    applyCurrentSort();
    toast(t("sortPresetLoaded"), "success");
  });
}

if (deleteSortPresetBtn) {
  deleteSortPresetBtn.addEventListener("click", () => {
    const name = sortPresetSelectEl?.value;
    if (!name) {
      toast(t("choosePresetToDelete"), "info");
      return;
    }
    const presets = loadSortPresets().filter((preset) => preset.name !== name);
    saveSortPresets(presets);
    renderSortPresets();
    toast(t("sortPresetDeleted"), "info");
  });
}

columnPickerEl.addEventListener("click", (e) => {
  if (e.target === columnPickerEl) closeColumnPicker();
});

columnPickerEl.addEventListener("keydown", handlePickerKeydown);

closePickerBtn.addEventListener("click", () => {
  closeColumnPicker();
});
columnSearchEl.addEventListener("input", filterColumnList);

exportCsvBtn.addEventListener("click", exportCsv);
if (resetSortBtn) {
  resetSortBtn.addEventListener("click", () => {
    multiSortState = [];
    normalizeSortState();
    applyCurrentSort();
    toast(t("defaultSortRestored"), "info");
  });
}
saveBtn.addEventListener("click", () => {
  toast(t("webSaveInfo"), "info");
});
saveAsBtn.addEventListener("click", saveWorkbookAs);
resetWidthsBtn.addEventListener("click", () => {
  manualColumnWidths = {};
  renderActiveTable();
  toast(t("widthsRestored"), "info");
});

tbodyEl.addEventListener("click", (e) => {
  const td = e.target.closest("td");
  if (!td || td.classList.contains("row-head")) return;
  const tr = td.parentElement;
  const rowKey = tr?.dataset.rowKey || "";
  const colIndex0 = parseInt(td.dataset.colIndex || "", 10);
  if (!rowKey || !Number.isFinite(colIndex0)) return;
  setFocusedCell(rowKey, colIndex0, { scroll: false });
});

tbodyEl.addEventListener("dblclick", (e) => {
  const td = e.target.closest("td");
  if (!td || td.classList.contains("row-head")) return;
  toast(t("cellEditingFuture"), "info");
});

[searchQueryEl, searchQuery2El, filterOperatorsEl, filterOperators2El, onlyNonEmptyEl, dateModeEl, dateFromEl, dateToEl, lastDaysEl].forEach((el) => {
  if (!el) return;
  el.addEventListener("input", updateFilterBadge);
  el.addEventListener("change", updateFilterBadge);
});
[dateModeEl, lastDaysEl].forEach((el) => {
  el.addEventListener("change", updateDateChipsActive);
  el.addEventListener("input", updateDateChipsActive);
});

searchQueryEl.addEventListener("input", syncQuickSearchInputs);
filterModeEl.addEventListener("change", syncQuickSearchModeControls);
if (filterOperatorsEl) {
  filterOperatorsEl.addEventListener("change", syncQuickSearchOperatorsControls);
}

maxRowsEl.addEventListener("change", () => {
  saveMaxRowsPreference();
  renderActiveTable();
});

zoomLevelEl.addEventListener("change", () => {
  setTimeout(() => {
    applyZoom();
    if (currentHeaders.length) {
      renderActiveTable();
    } else {
      syncHorizontalScrollbar();
    }
  }, 50);
});

if (excelLayoutToggleEl) {
  excelLayoutToggleEl.addEventListener("click", () => {
    setExcelLayoutEnabled(!isExcelLayoutEnabled());
    saveExcelLayoutPreference();
    renderActiveTable();
  });
}

if (langButtons.length) {
  langButtons.forEach((button) => {
    button.addEventListener("click", () => {
      const nextLang = button.dataset.lang;
      if (nextLang && nextLang !== currentLang) {
        applyLanguage(nextLang);
      }
    });
  });
}

window.addEventListener("resize", () => {
  updateLangSwitchIndicator();
});

applyLanguage(currentLang);
initIntroSplash();
initTheme();
loadMaxRowsPreference();
loadExcelLayoutPreference();
attachResizeHandlers();
applyZoom();
updateNetworkBadge();
window.addEventListener("online", updateNetworkBadge);
window.addEventListener("offline", updateNetworkBadge);

themeToggle.addEventListener("click", () => {
  const next = rootEl.getAttribute("data-theme") === "dark" ? "light" : "dark";
  setTheme(next);
});

if (brandRefreshBtn) {
  brandRefreshBtn.addEventListener("click", () => {
    hardRefreshApp();
  });

  const expandLogo = () => {
    brandRefreshBtn.classList.add("expanded");
    if (heroRightEl) heroRightEl.classList.add("expanded");
  };
  const collapseLogo = () => {
    brandRefreshBtn.classList.remove("expanded");
    if (heroRightEl) heroRightEl.classList.remove("expanded");
  };

  brandRefreshBtn.addEventListener("mouseenter", expandLogo);
  brandRefreshBtn.addEventListener("mouseleave", collapseLogo);
  brandRefreshBtn.addEventListener("pointerenter", expandLogo);
  brandRefreshBtn.addEventListener("pointerleave", collapseLogo);
  brandRefreshBtn.addEventListener("focus", expandLogo);
  brandRefreshBtn.addEventListener("blur", collapseLogo);
  brandRefreshBtn.addEventListener("touchstart", expandLogo, { passive: true });

  window.addEventListener("pageshow", collapseLogo);
  document.addEventListener("visibilitychange", () => {
    if (document.visibilityState === "hidden") collapseLogo();
  });
}

function toggleSidebar() {
  setSidebarOpen(!isSidebarOpen());
  syncSidebarHandle();
}

function syncSidebarHandle() {
  if (panelToggle) {
    panelToggle.setAttribute("aria-expanded", isSidebarOpen() ? "true" : "false");
    panelToggle.textContent = isSidebarOpen() ? t("panelOpen") : t("panelClosed");
  }
  if (panelHandle) {
    panelHandle.textContent = "";
    panelHandle.setAttribute("aria-expanded", isSidebarOpen() ? "true" : "false");
    panelHandle.setAttribute("aria-label", isSidebarOpen() ? t("sidebarCloseAria") : t("sidebarOpenAria"));
    panelHandle.setAttribute("title", isSidebarOpen() ? t("sidebarHideTitle") : t("sidebarShowTitle"));
    panelHandle.style.setProperty("--handle-open-label", `"${t("sidebarHandleLabel")}"`);

    if (isSidebarOpen() && sidebarEl) {
      const rect = sidebarEl.getBoundingClientRect();
      const overlap = 8;
      const nextLeft = Math.max(8, Math.round(rect.right - overlap)); // [EN] Allow tighter edge on narrow viewports; CSS handles closed state
      panelHandle.style.left = `${nextLeft}px`;
    } else {
      panelHandle.style.removeProperty("left"); // [EN] Let .sidebar-handle use fluid clamp() when closed
    }
  }
}

function setReadingMode(enabled) {
  rootEl.classList.toggle("reading", enabled);
  if (enabled) {
    if (quickSearchWrap) quickSearchWrap.classList.remove("hidden");
    if (readingToggle) readingToggle.textContent = t("readingStandard");
  } else {
    if (quickSearchWrap) quickSearchWrap.classList.add("hidden");
    if (readingToggle) readingToggle.textContent = t("readingQuick");
  }
  syncSidebarHandle();
}
