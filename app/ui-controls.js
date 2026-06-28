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
  if (validationState.showOnly) count += 1;

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
  if (highlightMatchCellsEl) highlightMatchCellsEl.checked = false;
  if (highlightMatchCellsDateEl) highlightMatchCellsDateEl.checked = false;
  highlightMatchedCells = false;
  matchedCellsByRow = new Map();
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
  quickSearchCellsMode = false;
  quickSearchFilterCellsMode = false;
  filtersCommitted = false;
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

// Re-triggerowalny pop: zdejmij klasę, wymuś reflow, dodaj — odpala animację od nowa.
function replayPop(el, cls) {
  if (!el) return;
  el.classList.remove(cls);
  void el.offsetWidth; // reflow, żeby animacja mogła zagrać ponownie
  el.classList.add(cls);
}

// Tryb „drawer" (nakładka, którą zamyka klik poza nią): telefon ZAWSZE; tablet TYLKO
// w pionie i tylko do ~900px szerokości; desktop, tablet w poziomie ORAZ szerokie tablety
// w pionie (iPad Pro 12.9" = 1024px) = panel STAŁY (klik poza nim NIE zamyka). Próg 900px:
// powyżej mieści się panel (~311px) + użyteczna tabela (~690px) obok siebie, jak na laptopie,
// więc drawer nie ma sensu. Zwykłe iPady w pionie (≤834px) nadal są drawerem.
const sidebarDrawerMQ = typeof matchMedia === "function"
  ? matchMedia("(max-width: 768px), (orientation: portrait) and (max-width: 900px)")
  : null;
function sidebarIsDrawer() { return !sidebarDrawerMQ || sidebarDrawerMQ.matches; }

let prevSidebarOpenState = null;
function setSidebarOpen(open) {
  const shouldOpen = !!open;
  // „Plumknięcie" handle tylko przy realnej zmianie stanu (nie przy starcie/init).
  const stateChanged = prevSidebarOpenState !== null && prevSidebarOpenState !== shouldOpen;
  prevSidebarOpenState = shouldOpen;
  rootEl.classList.toggle("sidebar-open", shouldOpen);
  if (sidebarScrim) sidebarScrim.classList.toggle("hidden", !shouldOpen);
  if (panelToggle) {
    panelToggle.setAttribute("aria-expanded", shouldOpen ? "true" : "false");
    panelToggle.textContent = shouldOpen ? t("panelOpen") : t("panelClosed");
  }
  if (stateChanged) replayPop(panelHandle, "handle-pop");
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
  // Pointer Events + setPointerCapture: bez GLOBALNYCH listenerów touchmove (te z passive:false
  // psuły płynne przewijanie na dotyku). W trakcie przeciągania aktualizujemy tylko szerokość
  // <col> / wysokość <tr> NA ŻYWO (bez przebudowy tabeli) — dzięki czemu uchwyt nie znika z DOM,
  // przechwycenie wskaźnika trwa i pointerup zawsze zwalnia resize (koniec „przyklejania się").
  let active = null; // { kind, index, start, startSize, handle, pointerId }
  let rafId = null;

  const colGroupCol = (i) => {
    const cg = tableEl.querySelector("colgroup");
    return cg ? cg.children[i + 1] : null; // +1: pierwszy <col> to kolumna numerów wierszy
  };

  const liveApply = () => {
    if (!active) return;
    if (active.kind === "col") {
      const col = colGroupCol(active.index);
      if (col) col.style.width = `${manualColumnWidths[active.index]}px`;
    } else {
      const tr = tbodyEl.querySelector(`tr[data-row-index="${active.index}"]`);
      if (tr) { tr.style.height = `${manualRowHeights[active.index]}px`; tr.classList.add("row-fixed-height"); }
    }
  };

  const onDown = (e) => {
    const colHandle = e.target.closest(".col-resizer");
    const rowHandle = !colHandle && e.target.closest(".row-resizer");
    const handle = colHandle || rowHandle;
    if (!handle) return;
    if (colHandle) {
      const th = handle.parentElement;
      active = { kind: "col", index: parseInt(handle.dataset.colIndex, 10), start: e.clientX, startSize: th.getBoundingClientRect().width, handle, pointerId: e.pointerId };
    } else {
      const tr = handle.closest("tr");
      active = { kind: "row", index: parseInt(handle.dataset.rowIndex, 10), start: e.clientY, startSize: tr ? tr.getBoundingClientRect().height : 28, handle, pointerId: e.pointerId };
    }
    document.body.classList.add("resizing");
    try { handle.setPointerCapture(e.pointerId); } catch {}
    e.preventDefault();
  };

  const onMove = (e) => {
    if (!active || e.pointerId !== active.pointerId) return;
    if (active.kind === "col") {
      manualColumnWidths[active.index] = Math.max(60, Math.min(900, Math.round(active.startSize + (e.clientX - active.start))));
    } else {
      manualRowHeights[active.index] = Math.max(16, Math.min(600, Math.round(active.startSize + (e.clientY - active.start))));
    }
    if (rafId) return;
    rafId = requestAnimationFrame(() => { rafId = null; liveApply(); });
  };

  const onUp = (e) => {
    if (!active || (e.pointerId != null && e.pointerId !== active.pointerId)) return;
    const finished = active;
    active = null;
    if (rafId) { cancelAnimationFrame(rafId); rafId = null; }
    document.body.classList.remove("resizing");
    try { finished.handle.releasePointerCapture(finished.pointerId); } catch {}
    renderActiveTable(); // finalizacja: colgroup, pasek przewijania, itd.
  };

  tableEl.addEventListener("pointerdown", onDown);
  tableEl.addEventListener("pointermove", onMove);
  tableEl.addEventListener("pointerup", onUp);
  tableEl.addEventListener("pointercancel", onUp);
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
    }, { once: true });
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
  // cursor-hint zamiast natywnego title; odświeżane przy zmianie języka i online/offline
  networkBadgeEl.setAttribute(
    "data-hint",
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

async function handleFile(file, fileHandle = null) {
  if (!file) return;
  if (!(await ensureXlsxLibs(true))) return; // dogrywa xlsx/jszip przy pierwszym użyciu
  try {
    const sizeHint = file.size > 0 ? ` (${formatFileSize(file.size)})` : "";
    setLoading(true, t("loadingFile") + sizeHint);
    const data = await file.arrayBuffer();
    originalFileBytes = new Uint8Array(data); // do zapisu metodą ZIP-patch (zachowanie pliku)
    pendingEdits = {}; // świeży plik → brak naniesionych edycji
    try {
      workbook = XLSX.read(data, { cellDates: true, cellStyles: true });
    } catch {
      workbook = XLSX.read(data, { cellDates: true });
    }
    // Odzyskaj mapę indeksów stylów z surowego pliku (font/kolor/rozmiar jak w Excelu).
    currentStyleIndexMap = await buildStyleIndexMap(originalFileBytes, workbook);
    // Wczytaj reguły formatowania warunkowego + dxf (ewaluowane leniwie przy renderze).
    await buildConditionalFormatting(originalFileBytes, workbook);
    // Wczytaj reguły Data Validation (listy/słowniki jak w Excelu) — type="list".
    // Konsumowane przez openCellEditor: dropdown + tryb (blokuj/ostrzegaj/podpowiadaj).
    if (typeof buildDataValidations === "function") await buildDataValidations(originalFileBytes, workbook);
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
    currentFileHandle = fileHandle; // uchwyt FSA (z showOpenFilePicker) lub null dla <input>/drop
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
    populateEditColumnSelect();
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

  // Drobny pokaz nowych możliwości stylów (czytanych z pliku przez aplikację):
  //  • wiersz nagłówka = zielony pasek z BIAŁYM POGRUBIONYM tekstem
  //    (naraz: wypełnienie wiersza + pogrubienie + biały tekst na tle),
  //  • kolumna „Status" = KOLORY CZCIONEK zależne od wartości.
  const headerAoaRow = 3; // [tytuł, KPI, pusty, headerRow, ...dane]
  for (let c = 0; c < headerRow.length; c++) {
    const ref = XLSX.utils.encode_cell({ r: headerAoaRow, c });
    if (ws[ref]) ws[ref].s = {
      font: { bold: true, color: { rgb: "FFFFFFFF" } },
      fill: { patternType: "solid", fgColor: { rgb: "FF2F6F5C" } },
      alignment: { horizontal: "center" },
    };
  }
  const STATUS_COLOR = { "Zamknięte": "FF1E7B45", "W trakcie": "FFB7791F", "PRZETERMINOWANY": "FFC00000" };
  rows.forEach((row, i) => {
    const color = STATUS_COLOR[row[2]];
    if (!color) return;
    const ref = XLSX.utils.encode_cell({ r: headerAoaRow + 1 + i, c: 2 }); // kolumna C = Status
    if (ws[ref]) ws[ref].s = { font: { bold: true, color: { rgb: color } } };
  });

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Obieg terenów");
  return XLSX.write(wb, { bookType: "xlsx", type: "array", cellStyles: true });
}

async function loadSampleFile() {
  if (!(await ensureXlsxLibs(true))) return; // dogrywa xlsx/jszip przy pierwszym użyciu
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

// ── Eksport / Raport: wybór kolumn + CSV + Drukuj/PDF ──
// Działa na BIEŻĄCYM (przefiltrowanym, posortowanym) modelu widoku — eksportujesz to,
// co widzisz. Wybór kolumn ogranicza wynik do zaznaczonych (po INDEKSIE, bo nagłówki
// bywają puste/zduplikowane). Druk buduje czysty #printArea (tylko wybrane kolumny)
// i woła window.print() → użytkownik zapisuje jako PDF systemowym dialogiem.
const exportModalEl = document.getElementById("exportModal");
const exportColumnListEl = document.getElementById("exportColumnList");

function exportModelOrNull() {
  const model = currentDisplayModel || getDisplayModel();
  if (!model.headers.length || !model.rows.length) {
    toast(t("noDataForExport"), "warning");
    return null;
  }
  return model;
}

function exportColLabel(header, idx) {
  const h = String(header ?? "").trim();
  return h || t("exportColumnFallback", { n: idx + 1 });
}

function getSelectedExportCols() {
  if (!exportColumnListEl) return [];
  return Array.from(exportColumnListEl.querySelectorAll("input[type=checkbox]"))
    .filter((cb) => cb.checked)
    .map((cb) => Number(cb.value));
}

function runCsvExport(cols) {
  const model = exportModelOrNull();
  if (!model) return;
  const useCols = cols && cols.length ? cols : model.headers.map((_, i) => i);
  const rows = [
    useCols.map((ci) => exportColLabel(model.headers[ci], ci)),
    ...model.rows.map((row) => useCols.map((ci) => getDisplayValue(row, ci))),
  ];
  const csv = rows.map((row) => row.map(escapeCsv).join(",")).join("\n");
  const base = currentFileName ? currentFileName.replace(/\.[^.]+$/, "") : "excel-workbench";
  const sheet = sheetSelect.value ? sheetSelect.value.replace(/\s+/g, "_") : "arkusz";
  const suffix = model.mode === "long" ? "long" : "wide";
  const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = `${base}_${sheet}_${suffix}.csv`;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
  toast(t("csvExported"), "success");
}

// Kontener druku jest BEZPOŚREDNIM dzieckiem <body>, żeby @media print mógł ukryć
// całą resztę przez `body > *:not(#printArea)`.
function ensurePrintArea() {
  let area = document.getElementById("printArea");
  if (!area) {
    area = document.createElement("div");
    area.id = "printArea";
    area.setAttribute("aria-hidden", "true");
    document.body.appendChild(area);
  }
  return area;
}

function runPrintExport(cols) {
  const model = exportModelOrNull();
  if (!model) return;
  const useCols = cols && cols.length ? cols : model.headers.map((_, i) => i);
  const area = ensurePrintArea();
  area.replaceChildren();

  const title = document.createElement("h1");
  title.className = "print-title";
  title.textContent = currentFileName || t("exportReportTitle");
  const meta = document.createElement("div");
  meta.className = "print-meta";
  const locale = (I18N[currentLang] && I18N[currentLang].locale) || "pl-PL";
  const dateStr = new Date().toLocaleString(locale);
  const sheetName = sheetSelect.value || "";
  meta.textContent = `${sheetName ? sheetName + " · " : ""}${dateStr} · ${t("exportRowsMeta", { count: model.rows.length })}`;

  const table = document.createElement("table");
  table.className = "print-table";
  const thead = document.createElement("thead");
  const htr = document.createElement("tr");
  useCols.forEach((ci) => {
    const th = document.createElement("th");
    th.textContent = exportColLabel(model.headers[ci], ci);
    htr.appendChild(th);
  });
  thead.appendChild(htr);
  table.appendChild(thead);
  const tbody = document.createElement("tbody");
  model.rows.forEach((row) => {
    const tr = document.createElement("tr");
    useCols.forEach((ci) => {
      const td = document.createElement("td");
      td.textContent = String(getDisplayValue(row, ci) ?? "");
      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  });
  table.appendChild(tbody);
  area.append(title, meta, table);
  window.print();
}

function openExportModal() {
  const model = exportModelOrNull();
  if (!model || !exportModalEl || !exportColumnListEl) return;
  exportColumnListEl.replaceChildren();
  model.headers.forEach((h, idx) => {
    const row = document.createElement("div");
    row.className = "field checkbox";
    const input = document.createElement("input");
    input.type = "checkbox";
    input.id = `exportcol-${idx}`;
    input.value = String(idx);
    input.checked = true;
    const label = document.createElement("label");
    label.htmlFor = input.id;
    label.textContent = exportColLabel(h, idx);
    row.appendChild(input);
    row.appendChild(label);
    exportColumnListEl.appendChild(row);
  });
  exportModalEl.classList.remove("hidden");
}

function closeExportModal() {
  if (exportModalEl) exportModalEl.classList.add("hidden");
}

// ── Walidacja listą referencyjną (data validation) ──
// Wskaż kolumnę + zbiór dozwolonych wartości (wpisanych albo z innej kolumny-słownika),
// a aplikacja policzy i pokaże wiersze z wartościami SPOZA listy. „Pokaż tylko niezgodne"
// komponuje się z filtrami tekstu/dat przez rdzenny applyFilters (rowIsValidationViolation).
const validationColumnEl = document.getElementById("validationColumn");
const validationDictColumnEl = document.getElementById("validationDictColumn");
const validationSourceEl = document.getElementById("validationSource");
const validationListWrapEl = document.getElementById("validationListWrap");
const validationColumnWrapEl = document.getElementById("validationColumnWrap");
const validationAllowedEl = document.getElementById("validationAllowed");
const validationIgnoreEmptyEl = document.getElementById("validationIgnoreEmpty");
const validationCaseEl = document.getElementById("validationCaseInsensitive");
const validationCheckBtn = document.getElementById("validationCheckBtn");
const validationClearBtn = document.getElementById("validationClearBtn");
const validationSummaryEl = document.getElementById("validationSummary");
const validationResultsEl = document.getElementById("validationResults");
const validationShowOnlyWrapEl = document.getElementById("validationShowOnlyWrap");
const validationShowOnlyEl = document.getElementById("validationShowOnly");

function populateValidationColumns() {
  if (!validationColumnEl) return;
  const prev = validationColumnEl.value;
  const prevDict = validationDictColumnEl ? validationDictColumnEl.value : "";
  const fill = (sel) => {
    if (!sel) return;
    sel.replaceChildren();
    currentHeaders.forEach((h, idx) => {
      const opt = document.createElement("option");
      opt.value = String(idx);
      opt.textContent = exportColLabel(h, idx);
      sel.appendChild(opt);
    });
  };
  fill(validationColumnEl);
  fill(validationDictColumnEl);
  if (prev && validationColumnEl.querySelector(`option[value="${prev}"]`)) validationColumnEl.value = prev;
  if (prevDict && validationDictColumnEl && validationDictColumnEl.querySelector(`option[value="${prevDict}"]`)) validationDictColumnEl.value = prevDict;
}

function resetValidationUi() {
  validationState = { colIdx: -1, allowed: null, ignoreEmpty: true, caseInsensitive: true, showOnly: false };
  if (validationSummaryEl) { validationSummaryEl.textContent = ""; validationSummaryEl.className = "validation-summary"; }
  if (validationResultsEl) validationResultsEl.replaceChildren();
  if (validationShowOnlyEl) validationShowOnlyEl.checked = false;
  if (validationShowOnlyWrapEl) validationShowOnlyWrapEl.classList.add("hidden");
}

function buildValidationAllowedSet(caseInsensitive) {
  const source = validationSourceEl ? validationSourceEl.value : "list";
  const set = new Set();
  if (source === "column" && validationDictColumnEl) {
    const dictIdx = Number(validationDictColumnEl.value);
    if (Number.isInteger(dictIdx) && dictIdx >= 0) {
      baseRows.forEach((row) => {
        const raw = getDisplayValue(row, dictIdx);
        if (String(raw == null ? "" : raw).trim() === "") return;
        set.add(normalizeValidationValue(raw, caseInsensitive));
      });
    }
  } else {
    const text = validationAllowedEl ? validationAllowedEl.value : "";
    text.split(/[\n,;]+/).forEach((tok) => {
      const v = tok.trim();
      if (v) set.add(normalizeValidationValue(v, caseInsensitive));
    });
  }
  return set;
}

function refreshValidationView() {
  applyFilters();
  sortRows();
  renderActiveTable();
  updateFilterBadge();
}

function renderValidationSummary(violations, total, distinctCount) {
  if (!validationSummaryEl) return;
  if (violations === 0) {
    validationSummaryEl.textContent = t("validationAllValid", { total });
    validationSummaryEl.className = "validation-summary ok";
  } else {
    validationSummaryEl.textContent = t("validationSummaryText", { bad: violations, total, values: distinctCount });
    validationSummaryEl.className = "validation-summary bad";
  }
}

function renderValidationResults(badValues) {
  if (!validationResultsEl) return;
  validationResultsEl.replaceChildren();
  if (!badValues.size) return;
  const title = document.createElement("div");
  title.className = "validation-results-title";
  title.textContent = t("validationBadValuesTitle");
  validationResultsEl.appendChild(title);
  Array.from(badValues.values())
    .sort((a, b) => b.count - a.count)
    .slice(0, 50)
    .forEach((v) => {
      const rowEl = document.createElement("div");
      rowEl.className = "validation-bad-row";
      const lab = document.createElement("span");
      lab.className = "validation-bad-label";
      lab.textContent = v.label;
      const cnt = document.createElement("span");
      cnt.className = "validation-bad-count";
      cnt.textContent = String(v.count);
      rowEl.append(lab, cnt);
      validationResultsEl.appendChild(rowEl);
    });
}

function runValidation() {
  if (!currentHeaders.length) { toast(t("loadSheetToPickColumns"), "info"); return; }
  const colIdx = validationColumnEl ? Number(validationColumnEl.value) : -1;
  if (!Number.isInteger(colIdx) || colIdx < 0) return;
  const caseInsensitive = !!(validationCaseEl && validationCaseEl.checked);
  const ignoreEmpty = !!(validationIgnoreEmptyEl && validationIgnoreEmptyEl.checked);
  const allowed = buildValidationAllowedSet(caseInsensitive);
  if (!allowed.size) { toast(t("validationNeedValues"), "warning"); return; }

  validationState = {
    colIdx,
    allowed,
    ignoreEmpty,
    caseInsensitive,
    showOnly: !!(validationShowOnlyEl && validationShowOnlyEl.checked),
  };

  // Liczymy po CAŁYM arkuszu (baseRows) — to samo, co rowIsValidationViolation.
  let total = 0;
  let violations = 0;
  const badValues = new Map(); // norm -> { label, count }
  baseRows.forEach((row) => {
    total += 1;
    const raw = getDisplayValue(row, colIdx);
    const str = String(raw == null ? "" : raw);
    if (str.trim() === "") { if (!ignoreEmpty) violations += 1; return; }
    const norm = normalizeValidationValue(raw, caseInsensitive);
    if (!allowed.has(norm)) {
      violations += 1;
      const ex = badValues.get(norm);
      if (ex) ex.count += 1;
      else badValues.set(norm, { label: str.trim(), count: 1 });
    }
  });

  renderValidationSummary(violations, total, badValues.size);
  renderValidationResults(badValues);
  if (validationShowOnlyWrapEl) validationShowOnlyWrapEl.classList.toggle("hidden", violations === 0);
  refreshValidationView();
}

function clearValidation() {
  resetValidationUi();
  if (validationAllowedEl) validationAllowedEl.value = "";
  refreshValidationView();
}

if (validationSourceEl) {
  validationSourceEl.addEventListener("change", () => {
    const useColumn = validationSourceEl.value === "column";
    if (validationListWrapEl) validationListWrapEl.classList.toggle("hidden", useColumn);
    if (validationColumnWrapEl) validationColumnWrapEl.classList.toggle("hidden", !useColumn);
  });
}
if (validationCheckBtn) validationCheckBtn.addEventListener("click", runValidation);
if (validationClearBtn) validationClearBtn.addEventListener("click", clearValidation);
if (validationShowOnlyEl) {
  validationShowOnlyEl.addEventListener("change", () => {
    validationState.showOnly = validationShowOnlyEl.checked;
    refreshValidationView();
  });
}

// Typy plików dla OTWIERANIA (showOpenFilePicker) — akceptuj oba formaty.
const FSA_FILE_TYPES = [
  {
    description: "Excel",
    accept: {
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": [".xlsx"],
      "application/vnd.ms-excel.sheet.macroEnabled.12": [".xlsm"],
    },
  },
];

// Typy dla ZAPISU (showSaveFilePicker) — TYLKO format zgodny z plikiem bazowym.
// Podanie obu (xlsx+xlsm) sprawiało, że mobilny picker (np. Samsung Internet)
// dolepiał .xlsm do pliku z treścią xlsx → plik nie do otwarcia. Jeden typ to wyklucza.
function fsaSaveTypes(ext) {
  if (ext === "xlsm") {
    return [{
      description: "Skoroszyt Excel z makrami",
      accept: { "application/vnd.ms-excel.sheet.macroEnabled.12": [".xlsm"] },
    }];
  }
  return [{
    description: "Skoroszyt Excel",
    accept: { "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": [".xlsx"] },
  }];
}

// Upewnij się, że mamy zgodę na zapis do uchwytu (FSA wymaga aktywacji użytkownika).
async function ensureWritePermission(handle) {
  const opts = { mode: "readwrite" };
  if ((await handle.queryPermission(opts)) === "granted") return true;
  if ((await handle.requestPermission(opts)) === "granted") return true;
  return false;
}

// Buduje bajty pliku wyjściowego. Preferuje ZIP-patch (zachowuje tabele, wykresy,
// style i formuły oryginału); gdy brak oryginału lub patch zawiedzie — fallback do
// XLSX.write (xlsx-js-style), który gubi część elementów, ale zawsze działa.
// Mapa rozwiniętych formuł z workbooka SheetJS: { sheetName: { ref: "A3*10+B2", … } }.
// Potrzebna do „od-dzielania" formuł dzielonych przy zapisie (unshareTouchedGroups) —
// SheetJS przy odczycie poprawnie tłumaczy formuły zależne z grupy shared, więc bierzemy
// gotowe `.f` zamiast samodzielnie przeliczać przesunięcia referencji.
function buildSheetFormulaMaps() {
  const maps = {};
  if (!workbook || !workbook.Sheets || !Array.isArray(workbook.SheetNames)) return maps;
  for (const name of workbook.SheetNames) {
    const sh = workbook.Sheets[name];
    if (!sh) continue;
    const m = {};
    for (const ref in sh) {
      if (ref.charCodeAt(0) === 33) continue; // pomiń metadane "!ref", "!cols" itp.
      const cell = sh[ref];
      if (cell && cell.f) m[ref] = cell.f;
    }
    if (Object.keys(m).length) maps[name] = m;
  }
  return maps;
}

async function buildOutputBytes(ext) {
  if (originalFileBytes && typeof buildPatchedXlsx === "function" && typeof JSZip !== "undefined") {
    try {
      return await buildPatchedXlsx(originalFileBytes, pendingEdits, buildSheetFormulaMaps());
    } catch (err) {
      log("ZIP-patch nie powiódł się — zapis przez xlsx-js-style (możliwa utrata tabel/wykresów).", "warning");
    }
  }
  return XLSX.write(workbook, { bookType: ext, type: "array", cellStyles: true });
}

// Zapis wprost do uchwytu pliku (nadpisanie w miejscu) bez pobierania.
// Dane owijamy w Blob — write(Blob) jest szerzej wspierane niż write(Uint8Array)
// (niektóre implementacje FSA, np. Samsung Internet, zapisywały 0 B z surowego bufora).
async function writeWorkbookToHandle(handle, ext) {
  const data = await buildOutputBytes(ext);
  const writable = await handle.createWritable();
  await writable.write(new Blob([data]));
  await writable.close();
}

// Pobranie pliku (fallback dla przeglądarek bez FSA, np. iOS Safari / Firefox).
// opts.toastKey / opts.toastType pozwalają dać inny komunikat, gdy pobranie jest
// ratunkiem po nieudanym zapisie w miejscu (a nie zwykłym "Zapisz jako…").
async function downloadWorkbook(name, ext, opts = {}) {
  const data = await buildOutputBytes(ext);
  const blob = new Blob([data], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = name;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
  setDirtyState(false);
  toast(t(opts.toastKey || "fileSaved"), opts.toastType || "success");
  log(`Zapisano plik: ${name}`, "success");
}

function workbookBaseName() {
  return currentFileName ? currentFileName.replace(/\.[^.]+$/, "") : "excel-workbench";
}

// Ratunkowe pobranie kopii, gdy zapis przez FSA zawiódł (np. wadliwa implementacja
// na Samsung Internet dająca plik 0 B). Pobranie działa na każdej przeglądarce Chromium,
// więc użytkownik dostaje działający plik zamiast błędu. Bez regresji na Chrome/iPad
// (uruchamia się tylko w gałęzi błędu zapisu w miejscu).
async function tryDownloadFallback(ext) {
  try {
    log("Zapis w miejscu nie powiódł się — pobieram kopię pliku.", "warning");
    await downloadWorkbook(`${workbookBaseName()}.${ext}`, ext, {
      toastKey: "saveFellBackToDownload",
      toastType: "warning",
    });
  } catch (err) {
    toast(t("saveFailed"), "error");
    log("Blad przy zapisie pliku.", "error");
  }
}

// "Zapisz": nadpisz oryginał w miejscu (FSA). Przy braku uchwytu — picker zapisu;
// bez FSA — fallback do "Zapisz jako…" (pobranie).
async function saveWorkbook() {
  if (!isXlsxAvailable(true)) return;
  if (!workbook) {
    toast(t("noFileToSave"), "warning");
    return;
  }
  // Nadpisanie istniejącego uchwytu (oryginalny plik otwarty przez FSA lub wybrany wcześniej).
  if (currentFileHandle) {
    const handleName = currentFileHandle.name || currentFileName || "";
    const ext = handleName.toLowerCase().endsWith(".xlsm") ? "xlsm" : "xlsx";
    // Ostrzeżenie przy KAŻDYM nadpisaniu w miejscu — to jedyna nieodwracalna operacja.
    // Docelowo (gdy testy round-tripu będą pewniejsze) można to poluzować np. do
    // jednorazowego potwierdzenia per plik/sesja.
    if (!window.confirm(t("saveInPlaceWarn"))) return;
    if (ext === "xlsm" && !window.confirm(t("xlsmConfirm"))) return;
    try {
      if (!(await ensureWritePermission(currentFileHandle))) {
        toast(t("savePermissionDenied"), "warning");
        return;
      }
      await writeWorkbookToHandle(currentFileHandle, ext);
      setDirtyState(false);
      toast(t("fileSaved"), "success");
      log(`Zapisano plik: ${handleName}`, "success");
    } catch (err) {
      if (err && err.name === "AbortError") return;
      await tryDownloadFallback(ext); // np. Samsung Internet: ratuj pobraniem kopii
    }
    return;
  }
  // Brak uchwytu, ale FSA dostępne → picker zapisu (zapamiętaj uchwyt do kolejnych zapisów).
  if (canFSA) {
    await saveWorkbookAs();
    return;
  }
  // Brak FSA → pobranie kopii.
  saveWorkbookAs();
}

// "Zapisz jako…": z FSA otwiera picker (i zapamiętuje uchwyt); bez FSA — pobranie kopii.
async function saveWorkbookAs() {
  if (!isXlsxAvailable(true)) return;
  if (!workbook) {
    toast(t("noFileToSave"), "warning");
    return;
  }
  const base = workbookBaseName();
  const defaultExt = currentFileName && currentFileName.toLowerCase().endsWith(".xlsm") ? "xlsm" : "xlsx";

  if (canFSA) {
    try {
      const handle = await window.showSaveFilePicker({
        suggestedName: `${base}.${defaultExt}`,
        types: fsaSaveTypes(defaultExt),
      });
      // Format = ten, który podaliśmy pickerowi (zgodny z plikiem bazowym).
      // Nie czytamy handle.name, bo mobilne pickery potrafią dolepić zły sufiks.
      const ext = defaultExt;
      if (ext === "xlsm" && !window.confirm(t("xlsmConfirm"))) return;
      await writeWorkbookToHandle(handle, ext);
      currentFileHandle = handle;
      setDirtyState(false);
      toast(t("fileSaved"), "success");
      log(`Zapisano plik: ${handle.name || base}`, "success");
    } catch (err) {
      if (err && err.name === "AbortError") return;
      // używamy defaultExt (ext żyje tylko w try) — np. Samsung Internet: ratuj pobraniem
      await tryDownloadFallback(defaultExt);
    }
    return;
  }

  // Fallback bez FSA: zapytaj o nazwę i pobierz kopię.
  const nameRaw = window.prompt(t("saveAsPrompt"), `${base}_edited.xlsx`);
  if (!nameRaw) return;
  let name = nameRaw.trim();
  if (!name) return;
  if (!/\.(xlsx|xlsm)$/i.test(name)) {
    name = `${name}.xlsx`;
  }
  const ext = name.toLowerCase().endsWith(".xlsm") ? "xlsm" : "xlsx";
  if (ext === "xlsm" && !window.confirm(t("xlsmConfirm"))) return;
  await downloadWorkbook(name, ext);
}

// Otwarcie pliku przez File System Access API — daje uchwyt do zapisu w miejscu.
// Po wczytaniu pliku przy ZAMKNIĘTYM sidebarze (np. przyciskiem w pustym stanie)
// prowadzi użytkownika do następnego kroku: otwiera sidebar z panelem „Plik i arkusz",
// pulsuje na wyborze arkusza i przycisku „Wczytaj arkusz" oraz pokazuje toast.
// Sterowanie wyglądem pulsu: klasa .guide-attention w CSS.
let guideAttentionTimer = null;
function guideToSheetConfig() {
  if (!workbook) return; // plik się nie wczytał — nie ma do czego prowadzić
  setSidebarOpen(true);
  const filePanel = document.getElementById("panel-file-sheet");
  if (filePanel) filePanel.open = true;
  const targets = [sheetSelect, loadBtn].filter(Boolean);
  if (guideAttentionTimer) window.clearTimeout(guideAttentionTimer);
  requestAnimationFrame(() => {
    if (filePanel) filePanel.scrollIntoView({ behavior: "smooth", block: "start" });
    targets.forEach((el) => {
      el.classList.remove("guide-attention");
      void el.offsetWidth; // restart animacji przy ponownym wczytaniu pliku
      el.classList.add("guide-attention");
    });
  });
  guideAttentionTimer = window.setTimeout(() => {
    targets.forEach((el) => el.classList.remove("guide-attention"));
    guideAttentionTimer = null;
  }, 4600);
  toast(t("guideChooseSheet"), "info");
}

async function openWorkbookViaFsa() {
  if (!canOpenFSA) return;
  const sidebarWasClosed = !isSidebarOpen();
  try {
    const [handle] = await window.showOpenFilePicker({
      types: FSA_FILE_TYPES,
      multiple: false,
    });
    if (!handle) return;
    const file = await handle.getFile();
    await handleFile(file, handle);
    if (sidebarWasClosed) guideToSheetConfig();
  } catch (err) {
    if (err && err.name === "AbortError") return;
    // Picker FSA mimo wszystko zawiódł (np. nierozpoznany webview/iframe) —
    // cofnij się do natywnego <input>, nie pokazuj błędu.
    log("FSA picker niedostępny — fallback do natywnego wyboru pliku.", "warning");
    fileInput.click();
  }
}

fileInput.addEventListener("change", async (e) => {
  const file = e.target.files[0];
  const sidebarWasClosed = !isSidebarOpen();
  await handleFile(file);
  if (sidebarWasClosed) guideToSheetConfig();
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
  // Użytkownik dotarł do przycisku — zgaś prowadzący puls (guideToSheetConfig)
  if (guideAttentionTimer) { window.clearTimeout(guideAttentionTimer); guideAttentionTimer = null; }
  [sheetSelect, loadBtn].forEach((el) => el && el.classList.remove("guide-attention"));
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
      currentColumnProfiles = []; // liczone leniwie z viewRows (ensureColumnProfilesFresh) — reaguje na filtr
      currentSections = detectSections(sheet, headerRow, data);
      currentRepeatingBlocks = detectRepeatingBlocks(sheet, headerRow, data);
      currentFormulaEntries = collectFormulaEntries(sheet, data, headerRow);
      if (!canUseLongView()) tableViewMode = "wide";
      viewRows = baseRows.slice();
      // Kolumny wyliczane (#6/#7): świeży arkusz → wyzeruj licznik i przelicz wirtualne
      // kolumny PRZED populacją selectów, żeby od razu były widoczne w sort/filtr/picker.
      if (typeof dcResetForNewSheet === "function") {
        dcResetForNewSheet();
        applyDerivedColumns({ silent: true });
      }
      filtersCommitted = false;
      multiSortState = [];
      sortState = { col: "", dir: "asc" };
      manualColumnWidths = {};
      manualRowHeights = {};
      columnSelections.filter1.clear();
      columnSelections.filter2.clear();
      columnSelections.date.clear();
      updateColumnSummary();
      updateFilterBadge();
      populateSortColumnSelect();
      populateEditColumnSelect();
      resetValidationUi(); // nowy arkusz → zacznij walidację od zera (nie filtruj po starej regule)
      populateValidationColumns();
      withSceneTransition(() => {
        // Zwiń wysoki panel „Plik i arkusz" W TYM SAMYM przebiegu co render tabeli.
        // Wcześniej robił to setTimeout(…,100) PO paint → otwarty sidebar skakał o ~400px
        // (panel jest wysoki), dając duży layout shift (CLS ~0.5 na mobile), zwłaszcza gdy
        // analiza trwała >0,5s i skok wypadał poza oknem wykluczenia po inpucie. Teraz collapse
        // jest częścią tej samej mutacji DOM: fallback = jeden layout (brak skoku), View
        // Transitions = płynny morf transformem (nie liczy się do CLS).
        const panelFileSheet = document.getElementById("panel-file-sheet");
        if (panelFileSheet) panelFileSheet.removeAttribute("open");
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
      });
      setDirtyState(false);
      if ((currentSheetStats?.trimmedColumns || 0) > 0) {
        log(`Przycięto puste kolumny poza realnym zakresem danych: ${currentSheetStats.trimmedColumns}`, "info");
      }
      if (currentSheetStats?.duplicateHeaderCount) {
        toast(t("duplicatedHeaders", { count: currentSheetStats.duplicateHeaderCount }), "warning");
      }
      toast(t("sheetLoaded"), "success");
      log(`Wczytano arkusz: ${sheetName}`, "success");
    } finally {
      setLoading(false);
    }
  }, 50);
});

applyFilterBtn.addEventListener("click", () => {
  if (!currentHeaders.length) return;
  // Świadomy „Filtruj" z sidebara = deterministyczne filtrowanie wierszy: zdejmij
  // tryby szybkiego szukania (zaznacz/cells), żeby się nie „lepiły" i nie blokowały ukrywania.
  quickSearchHighlightMode = false;
  quickSearchCellsMode = false;
  quickSearchFilterCellsMode = false;
  filtersCommitted = true;
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

  // Odczytaj tryb akcji: filtruj (ukryj niepasujące) / zaznacz (wyróżnij wiersze) /
  // cells (pokaż wszystkie, podświetl KOMÓRKI) / filter-cells (filtruj wiersze ORAZ
  // podświetl pasujące komórki w pozostałych).
  const actionEl = (popupActive && quickSearchPopupActionEl) ? quickSearchPopupActionEl : quickSearchActionEl;
  const action = actionEl ? actionEl.value : "filter";
  quickSearchHighlightMode = action === "highlight";
  quickSearchCellsMode = action === "cells";
  quickSearchFilterCellsMode = action === "filter-cells";
  if (action === "filter" || action === "filter-cells") filtersCommitted = true;

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
  // Delikatny feedback przeliczenia: filtrowanie przestawia/odsłania wiersze → FLIP;
  // „zaznacz" zostawia wiersze w miejscu → łagodny puls trafień; „cells" pokazuje
  // wszystkie wiersze (bez przestawiania) → bez FLIP-a.
  if (quickSearchHighlightMode) animateMatchPulseNextRender = true;
  else if (!quickSearchCellsMode) flipNextRender = true;
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

// ── Szybkie szukanie: zakres „wszystkie arkusze" + live-podgląd wyników ──
// JEDNA logika dla OBU wariantów paska (inline #quickSearchWrap oraz wyskakujące
// okno #quickSearchPopup ze skrótu Cmd/Ctrl+Shift+F). Każdy pasek to „kontekst"
// (input + select trybu + przełącznik + lista wyników + element-kotwica), a stan
// zakresu (bieżący ↔ wszystkie arkusze) jest WSPÓLNY i synchronizowany między nimi.
// Dzięki temu zmiana zachowania robi się w jednym miejscu i nie rozjeżdża między pasami.
let qsAllSheetsScope = false;
const qsContexts = [];

// Porównanie liczbowo/datowe na surowej komórce (raw=cell.v, display=cell.w/v) — odbicie
// cellSatisfiesComparison, ale bez obiektu-wiersza (live-podgląd skanuje surowe komórki,
// też z innych arkuszy). Reużywa globalnych parseCellNumber/parseDateFlexible/compareWithOp.
function qsCellSatisfiesComparison(raw, display, cmp) {
  if (cmp.kind === "date") {
    const ds = String(display);
    const looksDate = raw instanceof Date || /\d[-/.]\d/.test(ds) || (/\d/.test(ds) && /[a-ząćęłńóśźż]/i.test(ds));
    if (!looksDate) return false;
    const d = parseDateFlexible(raw != null ? raw : display);
    if (!(d instanceof Date) || Number.isNaN(d.getTime())) return false;
    return compareWithOp(cmp.op, comparisonDayValue(d), cmp.value);
  }
  if (raw instanceof Date) return false;
  const n = parseCellNumber(raw, display);
  if (n == null || !Number.isFinite(n)) return false;
  return compareWithOp(cmp.op, n, cmp.value);
}

// Pojedynczy term (z opcjonalnym „!") wobec komórki: porównanie >>/<< albo tekst zawiera/=.
function qsTermMatchesCell(raw, display, term, exact) {
  let t = term.trim();
  if (!t) return true;
  let neg = false;
  if (t.startsWith("!") && t.length > 1) { neg = true; t = t.slice(1).trim(); }
  const cmp = parseComparisonTerm(t);
  let ok;
  if (cmp) ok = qsCellSatisfiesComparison(raw, display, cmp);
  else {
    const hay = display.toLowerCase();
    ok = exact ? hay === t : hay.includes(t);
  }
  return neg ? !ok : ok;
}

// Operator-świadome dopasowanie KOMÓRKI: OR (||) grup AND (&&), z „!" i porównaniami.
// Gdy operatory wyłączone — zwykłe zawiera/dokładnie (jak dawniej). Składnia spójna z
// silnikiem filtra, tyle że oceniana per komórka (preview), więc międzykolumnowe && nie
// „połączy" dwóch komórek — ale porównania (główny przypadek operatorów) działają wiernie.
function qsCellMatches(raw, display, q, exact, operators) {
  if (!operators) {
    const hay = display.toLowerCase();
    return exact ? hay === q : hay.includes(q);
  }
  return q.split("||").some((part) => {
    const terms = part.split("&&").map((s) => s.trim()).filter(Boolean);
    if (!terms.length) return false;
    return terms.every((term) => qsTermMatchesCell(raw, display, term, exact));
  });
}

// Skan surowych komórek arkusza. `q` jest już .trim().toLowerCase(). Gdy `operators`,
// dopasowanie honoruje operatory wyszukiwania (>>/<</&&/||/!).
function qsScanSheet(sheetName, q, exact, perSheet, operators) {
  const sheet = workbook && workbook.Sheets ? workbook.Sheets[sheetName] : null;
  const res = { count: 0, hits: [] };
  if (!sheet) return res;
  for (const addr in sheet) {
    if (addr.charCodeAt(0) === 33) continue; // klucze meta („!ref", „!merges"…)
    const cell = sheet[addr];
    if (!cell) continue;
    const cellVal = cell.v; // surowa wartość (liczba/Data) — do porównań
    const display = cell.w != null ? String(cell.w) : (cell.v != null ? String(cell.v) : "");
    if (display === "") continue;
    if (qsCellMatches(cellVal, display, q, exact, operators)) {
      res.count += 1;
      if (res.hits.length < perSheet) res.hits.push({ addr, text: display });
    }
  }
  return res;
}

// Kolejność arkuszy do skanu: bieżący zawsze pierwszy; reszta tylko gdy zakres=wszystkie.
function qsSheetOrder() {
  if (!workbook || !Array.isArray(workbook.SheetNames)) return [];
  if (!qsAllSheetsScope) return currentSheetName ? [currentSheetName] : [];
  const rest = workbook.SheetNames.filter((n) => n !== currentSheetName);
  return currentSheetName ? [currentSheetName, ...rest] : workbook.SheetNames.slice();
}

function qsFirstMatchingSheet(q, exact, operators) {
  for (const name of qsSheetOrder()) {
    if (qsScanSheet(name, q, exact, 1, operators).count > 0) return name;
  }
  return currentSheetName;
}

function qsCtxExact(ctx) {
  return ctx.modeEl ? getNormalizedSelectValue(ctx.modeEl) === "exact" : false;
}
function hideQsLive(ctx) {
  if (ctx && ctx.liveEl) { ctx.liveEl.classList.add("hidden"); ctx.liveEl.replaceChildren(); }
}
function hideAllQsLive() {
  qsContexts.forEach(hideQsLive);
}
function qsSyncToggles() {
  qsContexts.forEach((c) => { if (c.toggleEl) c.toggleEl.setAttribute("aria-pressed", String(qsAllSheetsScope)); });
}

// Stosuje szukanie przez applyQuickSearch (respektuje tryb/akcję/kolumny/operatory
// AKTYWNEGO paska), skacząc najpierw do wskazanego arkusza, jeśli różni się od bieżącego.
// Ustawiamy wartość w OBU polach, bo applyQuickSearch czyta z aktywnego paska.
function qsApplyOnSheet(sheetName, value) {
  hideAllQsLive();
  if (quickSearchEl) quickSearchEl.value = value;
  if (quickSearchPopupInput) quickSearchPopupInput.value = value;
  if (sheetName && sheetName !== currentSheetName) {
    sheetSelect.value = sheetName;
    loadBtn.click(); // odbudowa arkusza jest asynchroniczna (setLoading + setTimeout)
    const deadline = Date.now() + 4000;
    (function waitLoaded() {
      if (currentSheetName === sheetName || Date.now() > deadline) {
        if (quickSearchEl) quickSearchEl.value = value;
        if (quickSearchPopupInput) quickSearchPopupInput.value = value;
        applyQuickSearch();
      } else setTimeout(waitLoaded, 50);
    })();
  } else {
    applyQuickSearch();
  }
}

// Wspólny commit dla Enter / przycisku „Szukaj" (oba paski): przy zakresie „wszystkie
// arkusze" skacze do pierwszego arkusza z trafieniem (bieżący ma priorytet).
function commitQuickSearch() {
  const popupActive = quickSearchPopupEl && !quickSearchPopupEl.classList.contains("hidden");
  const inputEl = popupActive ? quickSearchPopupInput : quickSearchEl;
  const modeEl = popupActive ? quickSearchPopupModeEl : quickSearchModeEl;
  const operatorsEl = popupActive ? quickSearchPopupOperatorsEl : quickSearchOperatorsEl;
  const value = inputEl ? inputEl.value : (searchQueryEl ? searchQueryEl.value : "");
  if (qsAllSheetsScope && value.trim().length >= 1) {
    const exact = modeEl ? getNormalizedSelectValue(modeEl) === "exact" : false;
    const operators = !!(operatorsEl && operatorsEl.checked);
    qsApplyOnSheet(qsFirstMatchingSheet(value.trim().toLowerCase(), exact, operators), value);
  } else {
    hideAllQsLive();
    applyQuickSearch();
  }
}

function renderQsLive(ctx) {
  if (!ctx || !ctx.liveEl || !ctx.inputEl) return;
  const value = ctx.inputEl.value || "";
  const q = value.trim().toLowerCase();
  if (q.length < 2 || !currentHeaders.length) { hideQsLive(ctx); return; }
  const exact = qsCtxExact(ctx);
  const operators = !!(ctx.operatorsEl && ctx.operatorsEl.checked);
  const PER_SHEET = 8;
  const groups = [];
  let total = 0;
  for (const name of qsSheetOrder()) {
    const r = qsScanSheet(name, q, exact, PER_SHEET, operators);
    if (r.count > 0) { groups.push({ sheet: name, count: r.count, hits: r.hits }); total += r.count; }
    if (groups.length >= 8) break; // ochrona długości listy
  }
  ctx.liveEl.replaceChildren();
  if (!total) {
    const empty = document.createElement("div");
    empty.className = "qs-live-empty";
    empty.textContent = t("globalSearchEmpty");
    ctx.liveEl.appendChild(empty);
    ctx.liveEl.classList.remove("hidden");
    return;
  }
  const frag = document.createDocumentFragment();
  groups.forEach((g) => {
    // Nagłówek arkusza tylko w trybie „wszystkie arkusze" (w jednym arkuszu zbędny).
    if (qsAllSheetsScope) {
      const title = document.createElement("div");
      title.className = "qs-live-group";
      title.textContent = `${g.sheet} · ${g.count}`;
      frag.appendChild(title);
    }
    g.hits.forEach((h) => {
      const item = document.createElement("button");
      item.type = "button";
      item.className = "qs-live-item";
      item.setAttribute("role", "option");
      const addr = document.createElement("span");
      addr.className = "qs-live-addr";
      addr.textContent = h.addr;
      const text = document.createElement("span");
      text.className = "qs-live-text";
      text.textContent = h.text.length > 80 ? `${h.text.slice(0, 80)}…` : h.text;
      item.append(addr, text);
      item.addEventListener("click", () => qsApplyOnSheet(g.sheet, value));
      frag.appendChild(item);
    });
    if (g.count > g.hits.length) {
      const more = document.createElement("div");
      more.className = "qs-live-more";
      more.textContent = t("globalSearchMore", { count: g.count - g.hits.length });
      frag.appendChild(more);
    }
  });
  ctx.liveEl.appendChild(frag);
  ctx.liveEl.classList.remove("hidden");
}

// Podpina identyczne zachowanie (przełącznik zakresu + live-podgląd) do jednego paska.
function wireQuickSearchScope(ctx) {
  if (!ctx.inputEl && !ctx.toggleEl) return;
  qsContexts.push(ctx);
  let timer = null;
  if (ctx.toggleEl) {
    ctx.toggleEl.addEventListener("click", () => {
      qsAllSheetsScope = !qsAllSheetsScope;
      qsSyncToggles(); // wspólny stan → druga belka też pokazuje aktualny zakres
      if (ctx.inputEl) ctx.inputEl.focus();
      renderQsLive(ctx);
    });
  }
  if (ctx.inputEl) {
    ctx.inputEl.addEventListener("input", () => {
      window.clearTimeout(timer);
      timer = window.setTimeout(() => renderQsLive(ctx), 180);
    });
    ctx.inputEl.addEventListener("focus", () => {
      if ((ctx.inputEl.value || "").trim().length >= 2) renderQsLive(ctx);
    });
    ctx.inputEl.addEventListener("keydown", (e) => {
      if (e.key === "Escape") hideQsLive(ctx);
    });
  }
  if (ctx.modeEl) ctx.modeEl.addEventListener("change", () => renderQsLive(ctx));
  // Przełączenie „Operatory wyszukiwania" → przelicz live-podgląd tym samym trybem.
  if (ctx.operatorsEl) ctx.operatorsEl.addEventListener("change", () => renderQsLive(ctx));
  document.addEventListener("click", (e) => {
    if (!ctx.liveEl || ctx.liveEl.classList.contains("hidden")) return;
    if (ctx.wrapEl && ctx.wrapEl.contains(e.target)) return;
    hideQsLive(ctx);
  });
}

wireQuickSearchScope({
  inputEl: quickSearchEl,
  modeEl: quickSearchModeEl,
  operatorsEl: quickSearchOperatorsEl,
  toggleEl: document.getElementById("qsAllSheets"),
  liveEl: document.getElementById("qsLiveResults"),
  wrapEl: quickSearchWrap,
});
wireQuickSearchScope({
  inputEl: quickSearchPopupInput,
  modeEl: quickSearchPopupModeEl,
  operatorsEl: quickSearchPopupOperatorsEl,
  toggleEl: document.getElementById("qsAllSheetsPopup"),
  liveEl: document.getElementById("qsLiveResultsPopup"),
  wrapEl: quickSearchPopupEl,
});

if (tableWrapEl && tableScrollbarEl) {
  // Touch axis-lock USUNIĘTY — psuł natywne przewijanie/momentum na tabletach i telefonach
  // (pionowe gesty bywały ignorowane, tabela „dryfowała" po puszczeniu, przeszkadzał resize).
  // Zostawiamy w pełni natywne przewijanie 2D z momentum przeglądarki.

  tableWrapEl.addEventListener("scroll", () => {
    hideCellTooltip();
    updateScrollTopFab();
    handleHeroScroll(tableWrapEl.scrollTop); // auto-chowanie nagłówka (mobile)
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

// FAB „do góry" — handle przy dolnej krawędzi tabeli, który (jak handle sidebara)
// rozwija się w pigułkę z etykietą na hover/focus. Widoczny dopiero po przewinięciu.
function updateScrollTopFab() {
  if (!scrollTopFabEl || !tableWrapEl) return;
  const show = tableWrapEl.scrollTop > 120;
  scrollTopFabEl.classList.toggle("is-visible", show);
}

if (scrollTopFabEl && tableWrapEl) {
  scrollTopFabEl.addEventListener("click", () => {
    tableWrapEl.scrollTo({ top: 0, behavior: prefersReducedMotion ? "auto" : "smooth" });
  });
}

// Dotykowy zamiennik Esc — odznacza zaznaczenie i fokus jednym tapnięciem.
if (clearSelectionFabEl) {
  clearSelectionFabEl.addEventListener("click", () => {
    setSelectedCell("", -1);
    setFocusedCell("", -1);
  });
}

// Lekkie „plumknięcie" przełączników trybu przy kliknięciu (ten sam feel co handle).
[wideLongToggleEl, excelLayoutToggleEl, readingToggle].forEach((btn) => {
  if (btn) btn.addEventListener("click", () => replayPop(btn, "btn-pop"));
});

// Zwijanie paska narzędzi nad tabelą — zsuwa rząd przycisków (i quick-search),
// zostawiając tylko status. Daje tabeli więcej miejsca, zwłaszcza na telefonach.
function updateToolbarToggleLabel() {
  if (!toolbarToggleEl) return;
  const collapsed = !!tablePanelEl && tablePanelEl.classList.contains("toolbar-collapsed");
  const label = collapsed ? t("toolbarExpand") : t("toolbarCollapse");
  toolbarToggleEl.setAttribute("aria-label", label);
  toolbarToggleEl.setAttribute("data-hint", label); // cursor-hint zamiast natywnego title
  toolbarToggleEl.setAttribute("aria-expanded", String(!collapsed));
}

function setToolbarCollapsed(collapsed) {
  if (!tablePanelEl) return;
  tablePanelEl.classList.toggle("toolbar-collapsed", !!collapsed);
  updateToolbarToggleLabel();
  try {
    localStorage.setItem(TOOLBAR_COLLAPSED_KEY, collapsed ? "1" : "0");
  } catch (e) {}
  syncTableViewportHeight();
}

if (toolbarToggleEl) {
  toolbarToggleEl.addEventListener("click", () => {
    const next = !tablePanelEl.classList.contains("toolbar-collapsed");
    setToolbarCollapsed(next);
    replayPop(toolbarToggleEl, "btn-pop");
  });
  setToolbarCollapsed(localStorage.getItem(TOOLBAR_COLLAPSED_KEY) === "1");
}

// ── Zwijany nagłówek (mobile): uchwyt tap/drag + auto-chowanie przy scrollu tabeli ──
// Odzyskane miejsce trafia do tabeli (syncTableViewportHeight liczy z offsetu panelu).
// Na desktopie (>768px) cały mechanizm jest bezczynny (CSS chowa uchwyt, JS wymusza off).
//
// REGUŁA ROZWIJANIA (świadoma, bez „migotania"): hero rozwija się TYLKO gdy:
//   1) przewiniesz tabelę do samej GÓRY (scrollTop ≈ 0),
//   2) klikniesz w uchwyt, albo
//   3) pociągniesz uchwyt palcem w dół.
// Zwija się przy scrollu w dół poza próg. Rozwijanie „przy każdym ruchu w górę"
// powodowało pętlę: zwinięcie rosło panel → przycinało scrollTop → fałszywy scroll-up
// → rozwinięcie → reflow → … . Histereza (dół-zwija / sama-góra-rozwija) + strażnik
// animacji rozrywają tę pętlę.
const heroEl = document.querySelector(".hero");
const heroGripEl = document.getElementById("heroGrip");
const heroNarrowMQ = typeof matchMedia === "function" ? matchMedia("(max-width: 768px)") : null;
let heroScrollLast = 0;
let heroAnimating = false;       // true w trakcie animacji zwijania — nie reaguj wtedy na scroll
let heroAnimTimer = null;
const HERO_COLLAPSE_AFTER = 48;  // zwijaj dopiero po zejściu poniżej tylu px
const HERO_EXPAND_AT_TOP = 6;    // rozwijaj dopiero przy samej górze widoku

function heroIsNarrow() { return !heroNarrowMQ || heroNarrowMQ.matches; }

// Zmierz wysokość ROZWINIĘTEGO hero do CSS var (nie da się animować z „auto").
function measureHeroHeight() {
  if (!heroEl || document.body.classList.contains("hero-collapsed")) return;
  const h = heroEl.getBoundingClientRect().height;
  if (h) document.documentElement.style.setProperty("--hero-h", `${Math.round(h)}px`);
}

function setHeroCollapsed(on) {
  if (!heroEl) return;
  if (!heroIsNarrow()) on = false; // desktop zawsze rozwinięty
  const was = document.body.classList.contains("hero-collapsed");
  if (was === !!on) return;
  if (on) measureHeroHeight(); // zmierz PRZED zwinięciem
  document.body.classList.toggle("hero-collapsed", !!on);
  // Strażnik anty-oscylacji: w trakcie animacji ignoruj scroll wywołany reflowem.
  // Na końcu (lub po fallbacku, gdy brak transition) przelicz wysokość panelu tabeli.
  heroAnimating = true;
  clearTimeout(heroAnimTimer);
  heroAnimTimer = setTimeout(() => {
    heroAnimating = false;
    if (typeof syncTableViewportHeight === "function") syncTableViewportHeight();
    heroScrollLast = tableWrapEl ? tableWrapEl.scrollTop : heroScrollLast; // świeży punkt odniesienia
  }, 340);
}

function handleHeroScroll(scrollTop) {
  if (!heroEl || !heroIsNarrow() || heroAnimating) { heroScrollLast = scrollTop; return; }
  const y = scrollTop;
  const collapsed = document.body.classList.contains("hero-collapsed");
  if (!collapsed && y > heroScrollLast + 4 && y > HERO_COLLAPSE_AFTER) {
    setHeroCollapsed(true);                 // scroll w dół poza próg → zwiń
  } else if (collapsed && y <= HERO_EXPAND_AT_TOP) {
    setHeroCollapsed(false);                // dotarcie do samej góry → rozwiń
  }
  heroScrollLast = y;
}

if (heroEl) {
  measureHeroHeight();
  if (heroGripEl) {
    // Gest: pociągnięcie uchwytu palcem (w dół rozwija, w górę zwija). Tap = toggle.
    let dragStartY = null;
    let dragMoved = false;
    heroGripEl.addEventListener("pointerdown", (e) => {
      dragStartY = e.clientY; dragMoved = false;
      try { heroGripEl.setPointerCapture(e.pointerId); } catch (_) {}
    });
    heroGripEl.addEventListener("pointermove", (e) => {
      if (dragStartY == null) return;
      const dy = e.clientY - dragStartY;
      if (Math.abs(dy) > 6) dragMoved = true;
      if (dy > 18) { setHeroCollapsed(false); dragStartY = e.clientY; }       // w dół → rozwiń
      else if (dy < -18) { setHeroCollapsed(true); dragStartY = e.clientY; }  // w górę → zwiń
    });
    const endDrag = () => { dragStartY = null; };
    heroGripEl.addEventListener("pointerup", endDrag);
    heroGripEl.addEventListener("pointercancel", endDrag);
    heroGripEl.addEventListener("click", () => {
      if (dragMoved) { dragMoved = false; return; } // to było przeciągnięcie, nie tap
      setHeroCollapsed(!document.body.classList.contains("hero-collapsed"));
      replayPop(heroGripEl, "btn-pop");
    });
  }
  // Snappy sync na końcu animacji (poza fallbackiem w setHeroCollapsed).
  heroEl.addEventListener("transitionend", (e) => {
    if (e.propertyName === "max-height" && typeof syncTableViewportHeight === "function") {
      syncTableViewportHeight();
    }
  });
  window.addEventListener("resize", () => {
    if (!heroIsNarrow()) setHeroCollapsed(false); // wejście w desktop → rozwiń
    if (!document.body.classList.contains("hero-collapsed")) measureHeroHeight();
  });
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

// Podpowiedź wartości komórki pokazujemy DOPIERO na tapnięcie (touchend bez ruchu),
// a NIE na touchstart. Wcześniej każdy start gestu przewijania odpalał showCellTooltip,
// który robi wymuszony reflow (scrollWidth/clientWidth + getBoundingClientRect ×2) i pokazuje
// tooltip z backdrop-filter blur(8px) — synchronicznie w handlerze touchstart, na starcie
// KAŻDEGO machnięcia i na każdej komórce z przyciętym tekstem (daty/nazwy). Na iOS to
// „hamowało" scroll pod palcem. Teraz: touchmove > próg = to przewijanie, nie tap → cisza.
// Wszystkie listenery passive → nigdy nie blokują natywnego przewijania.
let cellTapTd = null, cellTapX = 0, cellTapY = 0, cellTapMoved = false;
tbodyEl.addEventListener("touchstart", (e) => {
  const td = e.target.closest("td");
  if (!td || td.classList.contains("row-head")) { cellTapTd = null; return; }
  cellTapTd = td;
  cellTapMoved = false;
  const t = e.touches[0];
  cellTapX = t ? t.clientX : 0;
  cellTapY = t ? t.clientY : 0;
}, { passive: true });
tbodyEl.addEventListener("touchmove", (e) => {
  if (!cellTapTd || cellTapMoved) return;
  const t = e.touches[0];
  if (!t) return;
  if (Math.abs(t.clientX - cellTapX) > 8 || Math.abs(t.clientY - cellTapY) > 8) {
    cellTapMoved = true; // ruch = przewijanie, nie tapnięcie → nie pokazuj podpowiedzi
  }
}, { passive: true });
tbodyEl.addEventListener("touchend", () => {
  const td = cellTapTd;
  cellTapTd = null;
  if (!td || cellTapMoved) return;
  if (activeCellEditor) return; // podczas edycji nie pokazuj tooltipa (kolidowałby z dropdownem)
  showCellTooltip(td, true);
}, { passive: true });

window.addEventListener("resize", () => {
  syncHorizontalScrollbar();
  hideCellTooltip();
  syncSidebarHandle();
});

if (quickSearchBtn) {
  quickSearchBtn.addEventListener("click", commitQuickSearch);
}

if (quickSearchColumnsBtn) {
  quickSearchColumnsBtn.addEventListener("click", () => {
    lastPickerTriggerEl = quickSearchColumnsBtn;
    openColumnPicker("filter1");
  });
}

if (quickSearchModeEl) {
  quickSearchModeEl.addEventListener("change", () => {
    applyQuickSearchMode(getNormalizedSelectValue(quickSearchModeEl));
  });
}

// Enter uruchamia szukanie z DOWOLNEGO elementu szybkiego szukania (pole, selecty
// trybu/akcji, checkbox operatorów) — nie tylko z pola tekstowego. Wcześniej Enter
// wisiał na samym input; po kliknięciu innego elementu fokus przechodził na niego
// i Enter „przepadał". Dotyczyło OBU wariantów: paska pod przyciskiem (#quickSearchWrap)
// i okna ze skrótu Cmd+Shift+F (#quickSearchPopup). Wyjątek: przyciski (Kolumny/Szukaj)
// zachowują natywne działanie (klik), żeby Enter na „Kolumny" otwierał picker.
function attachQuickSearchEnter(container, handler) {
  if (!container) return;
  container.addEventListener("keydown", (e) => {
    if (e.key !== "Enter") return;
    if (e.target.closest && e.target.closest("button")) return; // przyciski robią swoje
    e.preventDefault();
    handler();
  });
}
attachQuickSearchEnter(quickSearchWrap, commitQuickSearch);   // pasek pod przyciskiem (z zakresem)
attachQuickSearchEnter(quickSearchPopupEl, applyQuickSearch); // okno ze skrótu Cmd+Shift+F
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
    filtersCommitted = true;
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

exportCsvBtn.addEventListener("click", openExportModal);
if (exportModalEl) {
  const exportCsvActionEl = document.getElementById("exportCsvAction");
  const exportPrintActionEl = document.getElementById("exportPrintAction");
  const exportSelectAllEl = document.getElementById("exportSelectAll");
  const exportClearAllEl = document.getElementById("exportClearAll");
  const closeExportEl = document.getElementById("closeExport");
  const setAllExportCols = (checked) => {
    exportColumnListEl.querySelectorAll("input[type=checkbox]").forEach((cb) => { cb.checked = checked; });
  };
  if (exportSelectAllEl) exportSelectAllEl.addEventListener("click", () => setAllExportCols(true));
  if (exportClearAllEl) exportClearAllEl.addEventListener("click", () => setAllExportCols(false));
  if (closeExportEl) closeExportEl.addEventListener("click", closeExportModal);
  if (exportCsvActionEl) exportCsvActionEl.addEventListener("click", () => {
    const cols = getSelectedExportCols();
    if (!cols.length) { toast(t("exportNoColumns"), "warning"); return; }
    closeExportModal();
    runCsvExport(cols);
  });
  if (exportPrintActionEl) exportPrintActionEl.addEventListener("click", () => {
    const cols = getSelectedExportCols();
    if (!cols.length) { toast(t("exportNoColumns"), "warning"); return; }
    closeExportModal();
    runPrintExport(cols);
  });
  exportModalEl.addEventListener("click", (e) => { if (e.target === exportModalEl) closeExportModal(); });
}
if (resetSortBtn) {
  resetSortBtn.addEventListener("click", () => {
    multiSortState = [];
    normalizeSortState();
    applyCurrentSort();
    toast(t("defaultSortRestored"), "info");
  });
}
saveBtn.addEventListener("click", () => {
  if (!canFSA) {
    toast(t("webSaveInfo"), "info");
    return;
  }
  saveWorkbook();
});
saveAsBtn.addEventListener("click", saveWorkbookAs);
resetWidthsBtn.addEventListener("click", () => {
  manualColumnWidths = {};
  manualRowHeights = {};
  renderActiveTable();
  toast(t("widthsRestored"), "info");
});

// Shift+klik buduje zakres komórek — wtedy zablokuj natywne (OS-owe) zaznaczanie
// tekstu, by nie „łapało" treści komórek. Zwykły drag bez Shift zaznacza tekst
// normalnie (np. żeby skopiować zawartość komórki).
tbodyEl.addEventListener("mousedown", (e) => {
  if (e.shiftKey) e.preventDefault();
});

tbodyEl.addEventListener("click", (e) => {
  const td = e.target.closest("td");
  if (!td || td.classList.contains("row-head")) return;
  const tr = td.parentElement;
  const rowKey = tr?.dataset.rowKey || "";
  const colIndex0 = parseInt(td.dataset.colIndex || "", 10);
  if (!rowKey || !Number.isFinite(colIndex0)) return;
  // Shift+klik:
  //  - jest już kotwica → rozszerz zaznaczenie do prostokąta (zakres myszą/dotykiem);
  //  - brak kotwicy → od razu zaznacz pojedynczą komórkę (zakres 1×1), bez kroku
  //    „najpierw cały wiersz, dopiero potem komórka".
  if (e.shiftKey) {
    if (!focusedCellState) {
      setFocusedCell(rowKey, colIndex0, { scroll: false });
    }
    setSelectedCell(rowKey, colIndex0, { scroll: false });
    clearTextSelection(); // sprzątnij ewentualną resztkę zaznaczenia tekstu
    return;
  }
  setFocusedCell(rowKey, colIndex0, { scroll: false });
});

// --- Edycja komórki (inline input) ---------------------------------------
// Edycja działa tylko w trybie "wide": wiersze long-view są syntetyczne i nie
// mapują się 1:1 na komórki arkusza. Zapis idzie przez updateSheetCell()
// (workbook.js), który zachowuje styl/format i blokuje formuły.
let activeCellEditor = null;

function getRowByKey(rowKey) {
  if (!rowKey || !currentDisplayModel || !Array.isArray(currentDisplayModel.rows)) return null;
  return currentDisplayModel.rows.find((r) => getRowSelectionKey(r) === rowKey) || null;
}

// Własny dropdown podpowiedzi pod edytorem komórki (Data Validation list).
// Powód istnienia: natywny <datalist> na dotyku kradnie focus z inputa → nasz
// blur zatwierdzał po jednej literze i zamykał edytor. Tu wybór pozycji NIE
// odbiera focusu (pointerdown/touch z flagą „interacting"), więc swobodne
// wpisywanie w trybie podpowiadaj/ostrzegaj jest płynne; tap autouzupełnia.
// Rozróżnia tap od przewijania (jak siatka), żeby scroll listy nie wybierał.
function createCellSuggestions(input, values, onPick) {
  const box = document.createElement("div");
  box.className = "cell-suggest hidden";
  box.setAttribute("role", "listbox");
  document.body.appendChild(box);

  let interacting = false;
  let interactTimer = null;
  const startInteract = () => { interacting = true; if (interactTimer) { clearTimeout(interactTimer); interactTimer = null; } };
  const endInteract = () => { if (interactTimer) clearTimeout(interactTimer); interactTimer = setTimeout(() => { interacting = false; }, 250); };

  const position = () => {
    const r = input.getBoundingClientRect();
    box.style.left = `${Math.round(r.left)}px`;
    box.style.top = `${Math.round(r.bottom + 2)}px`;
    box.style.minWidth = `${Math.round(r.width)}px`;
  };

  const choose = (v) => {
    interacting = false;
    if (typeof onPick === "function") onPick(v);
  };

  let touchStartY = 0, touchMoved = false;
  const makeItem = (v) => {
    const el = document.createElement("div");
    el.className = "cell-suggest-item";
    el.setAttribute("role", "option");
    el.textContent = v;
    // mysz/pen: pointerdown preventDefault => input nie traci focusu; wybór na click
    el.addEventListener("pointerdown", (e) => { startInteract(); if (e.pointerType !== "touch") e.preventDefault(); });
    el.addEventListener("pointerup", endInteract);
    el.addEventListener("click", () => choose(v));
    // dotyk: odróżnij tap od scrolla; tap (bez ruchu) wybiera bez gubienia focusu
    el.addEventListener("touchstart", (e) => { startInteract(); touchMoved = false; touchStartY = e.touches[0] ? e.touches[0].clientY : 0; }, { passive: true });
    el.addEventListener("touchmove", (e) => { const y = e.touches[0] ? e.touches[0].clientY : 0; if (Math.abs(y - touchStartY) > 8) touchMoved = true; }, { passive: true });
    el.addEventListener("touchend", (e) => { if (!touchMoved) { e.preventDefault(); choose(v); } else { endInteract(); } }, { passive: false });
    return el;
  };

  const render = () => {
    const q = input.value.trim().toLowerCase();
    const starts = [], has = [];
    let exact = false;
    for (const v of values) {
      const lv = String(v).toLowerCase();
      if (lv === q && q) { exact = true; continue; } // dokładnie wpisana wartość z listy → ukryj
      if (!q || lv.startsWith(q)) starts.push(v);
      else if (lv.includes(q)) has.push(v);
    }
    let items = starts.concat(has);
    // Brak dopasowań do tego, co wpisano (np. wartość SPOZA listy lub literówka),
    // a nie jest to dokładne trafienie → pokaż CAŁĄ listę (wtedy podpowiedź jest
    // najpotrzebniejsza: user widzi z czego wybierać). Dokładne trafienie → nic.
    if (!items.length && !exact) items = values.slice();
    items = items.slice(0, 50);
    box.replaceChildren();
    if (!items.length) { box.classList.add("hidden"); return; }
    items.forEach((v) => box.appendChild(makeItem(v)));
    box.classList.remove("hidden");
    position();
  };

  const onInput = () => render();
  const onScroll = () => { if (!box.classList.contains("hidden")) position(); };
  input.addEventListener("input", onInput);
  (tableWrapEl || window).addEventListener("scroll", onScroll, { passive: true });
  window.addEventListener("resize", onScroll, { passive: true });

  render();

  return {
    isInteracting: () => interacting,
    destroy() {
      if (interactTimer) clearTimeout(interactTimer);
      input.removeEventListener("input", onInput);
      (tableWrapEl || window).removeEventListener("scroll", onScroll);
      window.removeEventListener("resize", onScroll);
      box.remove();
    },
  };
}

function formatDateForEdit(d) {
  const pad = (n) => String(n).padStart(2, "0");
  const base = `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
  if (d.getHours() || d.getMinutes() || d.getSeconds()) {
    return `${base} ${pad(d.getHours())}:${pad(d.getMinutes())}`;
  }
  return base;
}

function cellEditString(row, i) {
  const raw = Array.isArray(row.values) ? row.values[i] : null;
  if (raw == null || raw === "") return "";
  if (raw instanceof Date) return formatDateForEdit(raw);
  return String(raw);
}

function openCellEditor(td, options = {}) {
  if (activeCellEditor || !td || td.classList.contains("row-head")) return;
  if (!workbook || !currentDisplayModel) return;
  if (currentDisplayModel.mode !== "wide") {
    toast(t("editWideOnly"), "info");
    return;
  }
  const tr = td.parentElement;
  const rowKey = tr?.dataset.rowKey || "";
  const colIndex0 = parseInt(td.dataset.colIndex || "", 10);
  if (!rowKey || !Number.isFinite(colIndex0)) return;
  const row = getRowByKey(rowKey);
  if (!row) return;
  if (row.isLongViewRow || row.isSubheader) {
    toast(t("editBlockedRow"), "info");
    return;
  }

  setFocusedCell(rowKey, colIndex0, { scroll: false });

  const prevText = td.dataset.fullText != null ? td.dataset.fullText : td.textContent;
  const input = document.createElement("input");
  input.type = "text";
  input.className = "cell-editor";
  input.setAttribute("aria-label", t("editCellAria"));
  const initialChar = options.initialChar;
  input.value = initialChar != null ? initialChar : cellEditString(row, colIndex0);
  // Wartość wyjściowa — do wykrycia, czy user faktycznie coś zmienił.
  // Start pisaniem (initialChar) traktujemy od razu jako zmianę.
  const baseValue = initialChar != null ? null : input.value;

  // Data Validation (lista jak w Excelu): jeśli komórka ma regułę type="list",
  // pokaż WŁASNY dropdown podpowiedzi (nie natywny <datalist> — ten na dotyku
  // kradnie focus z inputa, przez co blur zatwierdzał już po jednej literze).
  // Tryb rygoru egzekwujemy w commit (stop=blokuj, warning=ostrzeż, info=cicho).
  let dvRule = null;
  try {
    if (typeof getCellDataValidation === "function") {
      dvRule = getCellDataValidation(currentSheetName, row.rowIndex0, currentStartCol + colIndex0);
    }
  } catch { dvRule = null; }

  td.classList.add("cell-editing");
  td.appendChild(input);
  activeCellEditor = { td, input };

  // Własny popup podpowiedzi (poniżej). Tapnięcie pozycji NIE odbiera focusu
  // inputowi → blur nie zamyka edytora; tap autouzupełnia i zatwierdza.
  const dvSuggest = (dvRule && dvRule.values.length)
    ? createCellSuggestions(input, dvRule.values, (v) => { input.value = v; commit(null); })
    : null;
  if (dvSuggest) input.classList.add("cell-editor-has-list");

  let finished = false;
  const close = () => {
    if (finished) return;
    finished = true;
    if (dvSuggest) dvSuggest.destroy();
    input.remove();
    td.classList.remove("cell-editing");
    activeCellEditor = null;
  };
  const commit = (move) => {
    if (finished) return;
    // Brak zmian (np. samo przejechanie Shift+strzałką przez komórkę) = nie
    // przepisuj komórki: zachowaj jej formatowanie i nie ustawiaj "dirty".
    // Po prostu zamknij i ewentualnie przejdź dalej.
    if (baseValue != null && input.value === baseValue) {
      close();
      if (move) {
        setFocusedCell(rowKey, colIndex0, { scroll: false });
        moveFocusedCell(move.row || 0, move.col || 0);
      }
      updateCellStats();
      return;
    }
    const parsed = parseInputValue(input.value);
    if (parsed && parsed.type === "formula") {
      toast(t("formulaEditBlocked"), "warning");
      input.focus();
      return;
    }
    // Egzekwuj regułę Data Validation (lista). stop=blokuj commit, warning=ostrzeż
    // ale przepuść, info=cicho. Porównanie bez wielkości liter (jak w Excelu).
    if (dvRule) {
      const candidate = parsed ? parsed.value : null;
      const str = candidate == null ? "" : String(candidate).trim();
      let ok;
      if (str === "") ok = dvRule.allowBlank;
      else { const norm = str.toLowerCase(); ok = dvRule.values.some((v) => String(v).trim().toLowerCase() === norm); }
      if (!ok) {
        if (dvRule.mode === "stop") { toast(t("dvRejected"), "warning"); input.focus(); return; }
        if (dvRule.mode === "warning") toast(t("dvWarning"), "warning");
      }
    }
    updateSheetCell(row.rowIndex0, colIndex0, parsed);
    const newVal = parsed ? parsed.value : null;
    if (Array.isArray(row.values)) row.values[colIndex0] = newVal;
    if (Array.isArray(row.rawValues)) row.rawValues[colIndex0] = newVal;
    const display = newVal == null ? "" : toDisplay(newVal);
    if (Array.isArray(row.display)) row.display[colIndex0] = display;
    close();
    td.textContent = display;
    td.dataset.fullText = display;
    setDirtyState(true);
    if (move) {
      setFocusedCell(rowKey, colIndex0, { scroll: false });
      moveFocusedCell(move.row || 0, move.col || 0);
    }
    updateCellStats();
  };

  input.addEventListener("keydown", (e) => {
    e.stopPropagation();
    if (e.key === "Enter") {
      // Enter = zatwierdź i w dół; Shift+Enter = zatwierdź i w górę (jak w Excelu).
      e.preventDefault();
      commit({ row: e.shiftKey ? -1 : 1 });
    } else if (e.key === "Tab") {
      e.preventDefault();
      commit({ col: e.shiftKey ? -1 : 1 });
    } else if (
      e.shiftKey &&
      (e.key === "ArrowUp" || e.key === "ArrowDown" || e.key === "ArrowLeft" || e.key === "ArrowRight")
    ) {
      // Shift+strzałki = zatwierdź i przenieś EDYCJĘ do sąsiedniej komórki.
      // Same strzałki (bez Shift) zostają domyślnym ruchem kursora w tekście.
      e.preventDefault();
      const move =
        e.key === "ArrowUp" ? { row: -1 } :
        e.key === "ArrowDown" ? { row: 1 } :
        e.key === "ArrowLeft" ? { col: -1 } : { col: 1 };
      commit(move);
      // Zostań w trybie edycji: otwórz edytor na komórce, do której przeszliśmy.
      // Dzięki temu kolejne Shift+strzałki dalej przesuwają edycję (zamiast wpadać
      // w zaznaczanie zakresu w siatce) i zawsze widać, która komórka jest aktywna.
      if (!activeCellEditor && focusedCellState) {
        const nextTd = findCellElement(focusedCellState);
        if (nextTd) openCellEditor(nextTd);
      }
    } else if (e.key === "Escape") {
      e.preventDefault();
      close();
      td.textContent = prevText;
    }
  });
  input.addEventListener("blur", () => {
    // Tapnięcie/przewijanie listy podpowiedzi nie może zamykać edytora (na dotyku
    // input chwilowo traci focus) — w trakcie interakcji z popupem pomijamy commit.
    if (dvSuggest && dvSuggest.isInteracting()) return;
    commit(null);
  });

  input.focus();
  if (initialChar != null) {
    const end = input.value.length;
    input.setSelectionRange(end, end);
  } else {
    input.select();
  }
}

tbodyEl.addEventListener("dblclick", (e) => {
  const td = e.target.closest("td");
  if (!td || td.classList.contains("row-head")) return;
  openCellEditor(td);
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
loadCellStylePreferences();
loadRowHeightPreference();
loadColWidthPreference();
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
    // cursor-hint zamiast natywnego title — tekst zależny od stanu (schowaj/pokaż)
    panelHandle.setAttribute("data-hint-pl", isSidebarOpen() ? "Schowaj sidebar" : "Wysuń sidebar");
    panelHandle.setAttribute("data-hint-en", isSidebarOpen() ? "Hide the sidebar" : "Open the sidebar");
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
  if (readingToggle) readingToggle.setAttribute("aria-pressed", String(enabled));
  if (enabled) {
    if (quickSearchWrap) quickSearchWrap.classList.remove("hidden");
    if (readingToggle) readingToggle.textContent = t("readingStandard");
  } else {
    if (quickSearchWrap) quickSearchWrap.classList.add("hidden");
    if (readingToggle) readingToggle.textContent = t("readingQuick");
  }
  syncSidebarHandle();
}
