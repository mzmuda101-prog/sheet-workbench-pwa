// App bootstrap: event wiring, initial render, and runtime startup.

scheduleCursorHintInit();

panelToggle.addEventListener("click", toggleSidebar);
if (panelHandle) panelHandle.addEventListener("click", toggleSidebar);
if (sidebarScrim) sidebarScrim.addEventListener("click", () => setSidebarOpen(false));
document.querySelectorAll("details.panel").forEach((det) => {
  det.addEventListener("toggle", () => {
    // Rozwinięcie panelu → dorenderuj jego analizy, które były pominięte gdy był
    // zwinięty (leniwe renderowanie analiz, perf na słabszych urządzeniach).
    // Ciężkie analizy (duration ~1,9s) ODRACZAMY do następnej klatki: najpierw niech
    // panel płynnie się rozsunie i pojawi (paint), potem liczymy — inaczej synchroniczne
    // liczenie w handlerze zacina animację otwarcia. Re-otwarcia bez brudnych analiz
    // pomijają to całkowicie (renderDirtyAnalysesForPanel liczy tylko brudne).
    if (det.open && typeof renderDirtyAnalysesForPanel === "function"
        && typeof panelHasDirtyAnalyses === "function" && panelHasDirtyAnalyses(det.id)) {
      // Najpierw natychmiast pokaż „Liczę…" (panel nie wygląda na pusty/niedziałający),
      // dopiero potem (po paincie) policz — ciężkie liczenie nie zacina animacji otwarcia.
      if (typeof showHeavyComputingHint === "function") showHeavyComputingHint(det.id);
      requestAnimationFrame(() => requestAnimationFrame(() => {
        if (det.open) renderDirtyAnalysesForPanel(det.id);
      }));
    }
    // Formula Workbench nie jest w ANALYSIS_PANELS — pierwszy raz pokaż pusty stan po otwarciu.
    if (det.open && det.id === "panel-formula-workbench" && typeof renderFormulaWorkbench === "function"
        && formulaWorkbenchSummaryEl && !formulaWorkbenchSummaryEl.childElementCount) {
      renderFormulaWorkbench();
    }
    if (!isSidebarOpen()) return;
    requestAnimationFrame(() => syncSidebarHandle()); // [EN] :has() width changes — no resize event; keep handle aligned
    window.setTimeout(() => syncSidebarHandle(), 260);
  });
});
if (sectionNavigatorEl) {
  sectionNavigatorEl.addEventListener("click", (e) => {
    const btn = e.target.closest("button[data-section-index]");
    if (!btn) return;
    const idx = parseInt(btn.dataset.sectionIndex || "", 10);
    if (!Number.isFinite(idx) || idx < 0 || idx >= currentSections.length) return;
    focusSection(currentSections[idx]);
  });
}
if (repeatBlockDetectorEl) {
  repeatBlockDetectorEl.addEventListener("click", (e) => {
    const btn = e.target.closest("button[data-repeat-group-index]");
    if (!btn) return;
    const groupIndex = parseInt(btn.dataset.repeatGroupIndex || "", 10);
    const blockIndex = parseInt(btn.dataset.repeatBlockIndex || "", 10);
    if (!Number.isFinite(groupIndex) || !Number.isFinite(blockIndex)) return;
    focusRepeatingBlock(groupIndex, blockIndex);
  });
}
if (durationAnalysisSummaryEl) {
  durationAnalysisSummaryEl.addEventListener("click", (e) => {
    e.stopPropagation();
    const btn = e.target.closest("button[data-duration-action]");
    if (!btn) return;
    const action = btn.dataset.durationAction;

    if (action === "toggle-long" && canUseLongView()) {
      tableViewMode = tableViewMode === "long" ? "wide" : "long";
      manualColumnWidths = {};
      withSceneTransition(() => {
        scheduleViewRefresh({ table: true, analyses: true, sync: true });
      });
      toast(tableViewMode === "long" ? t("wideLongOn") : t("wideLongOff"), "info");
      return;
    }

    if (action === "reset-filters") {
      resetFiltersBtn.click();
    }
  });
  durationAnalysisSummaryEl.addEventListener("change", (e) => {
    e.stopPropagation();
    const control = e.target.closest("[data-duration-control]");
    if (!control) return;
    const kind = control.dataset.durationControl;
    if (kind === "status") {
      durationAnalysisState.statusFilter = control.value || "all";
    } else if (kind === "sort") {
      durationAnalysisState.sortMetric = control.value || "avg";
    } else if (kind === "count") {
      const next = parseInt(control.value || "14", 10);
      durationAnalysisState.showCount = Number.isFinite(next) && next > 0 ? next : 14;
    }
    scheduleViewRefresh({ analyses: true });
  });
}
if (durationAnalysisListEl) {
  durationAnalysisListEl.addEventListener("click", (e) => {
    e.stopPropagation();
    const btn = e.target.closest("button[data-duration-action='filter-entity']");
    if (!btn) return;
    const entity = (btn.dataset.durationEntity || "").trim();
    if (!entity) return;
    searchQueryEl.value = entity;
    filtersCommitted = true;
    applyFilters();
    sortRows();
    scheduleViewRefresh({ table: true, analyses: true, formula: true, filterBadge: true });
    toast(t("filteredFor", { value: entity }), "info");
  });
}
if (monthlySummaryEl) {
  const withScroll = (fn) => {
    const sb = document.querySelector(".sidebar");
    const savedScroll = sb ? sb.scrollTop : 0;
    fn();
    renderMonthlySummary();
    if (sb) sb.scrollTop = savedScroll;
  };
  monthlySummaryEl.addEventListener("change", (e) => {
    const control = e.target.closest("[data-monthly-control]");
    if (!control) return;
    e.stopPropagation();
    const kind = control.dataset.monthlyControl;
    withScroll(() => {
      if (kind === "metric") monthlySummaryState.metric = control.value || "occurrences";
      else if (kind === "months") monthlySummaryState.months = parseInt(control.value, 10);
      else if (kind === "anchor") monthlySummaryState.anchor = control.value === "today" ? "today" : "data";
      else if (kind === "split") monthlySummaryState.split = control.checked;
      else if (kind === "gap") monthlySummaryState.gap = control.checked;
    });
  });
  // chipy multi-wyboru: kolumny dat (min. 1) oraz kolumny miary (min. 1, jeśli już są wybrane)
  monthlySummaryEl.addEventListener("click", (e) => {
    const dateChip = e.target.closest("[data-monthly-datecol]");
    const measChip = e.target.closest("[data-monthly-measurecol]");
    const chip = dateChip || measChip;
    if (!chip) return;
    e.stopPropagation();
    const stateKey = dateChip ? "dateCols" : "measureCols";
    const idx = parseInt(chip.dataset[dateChip ? "monthlyDatecol" : "monthlyMeasurecol"], 10);
    const cur = Array.isArray(monthlySummaryState[stateKey]) ? monthlySummaryState[stateKey].slice() : [];
    const pos = cur.indexOf(idx);
    if (pos >= 0) { if (cur.length > 1) cur.splice(pos, 1); } // zostaw co najmniej jedną
    else cur.push(idx);
    withScroll(() => { monthlySummaryState[stateKey] = cur; });
  });
}

if (aggregationWorkbenchSummaryEl) {
  aggregationWorkbenchSummaryEl.addEventListener("change", (e) => {
    e.stopPropagation();
    const sidebarEl = document.querySelector(".sidebar");
    const savedSidebarScroll = sidebarEl ? sidebarEl.scrollTop : 0;
    const control = e.target.closest("[data-aggregation-control]");
    if (!control) return;
    const kind = control.dataset.aggregationControl;
    if (kind === "source") aggregationWorkbenchState.sourceMode = control.value || "auto";
    if (kind === "scope") aggregationWorkbenchState.scopeMode = control.value || "filtered";
    if (kind === "header") {
      aggregationWorkbenchState.headerRowChoice = control.value === "manual" ? "manual" : "auto";
      if (aggregationWorkbenchState.headerRowChoice === "manual") {
        const fallbackRow = Number.isFinite(aggregationWorkbenchState.customHeaderRow) && aggregationWorkbenchState.customHeaderRow > 0
          ? aggregationWorkbenchState.customHeaderRow
          : currentHeaderRow;
        aggregationWorkbenchState.customHeaderRow = fallbackRow;
      }
    }
    if (kind === "header-number") {
      const next = parseInt(control.value || "", 10);
      if (!Number.isFinite(next) || next < 1) {
        toast(t("positiveHeaderRow"), "warning");
        control.value = String(aggregationWorkbenchState.customHeaderRow || currentHeaderRow);
        return;
      }
      if (!isValidAggregationHeaderRow(next)) {
        toast(t("invalidHeaderRow", { row: next }), "error");
        control.value = String(aggregationWorkbenchState.customHeaderRow || currentHeaderRow);
        return;
      }
      aggregationWorkbenchState.customHeaderRow = next;
      aggregationWorkbenchState.headerRowChoice = "manual";
    }
    if (kind === "measure-pick") {
      // Obsluzone przez event click na przycisku
      return;
    }
    if (kind === "aggregation") aggregationWorkbenchState.aggregation = control.value || "count";
    if (kind === "match") aggregationWorkbenchState.matchMode = control.value || "contains";
    if (kind === "groupmode") aggregationWorkbenchState.groupMode = control.value || "exact";
    if (kind === "grouppattern") aggregationWorkbenchState.groupPattern = control.value || "";
    if (kind === "measurefilter") {
      aggregationWorkbenchState.measureFilterMode = control.value || "all";
      const valueInput = aggregationWorkbenchSummaryEl.querySelector("[data-aggregation-control=\"measurefilter-value\"]");
      if (valueInput) {
        valueInput.style.display = aggregationWorkbenchState.measureFilterMode === "all" ? "none" : "inline-block";
      }
    }
    if (kind === "measurefilter-value") {
      aggregationWorkbenchState.measureFilterValue = control.value || "";
    }
    if (kind === "count") {
      const next = parseInt(control.value || "20", 10);
      aggregationWorkbenchState.showCount = Number.isFinite(next) && next > 0 ? next : 20;
    }
    if (kind === "having") {
      aggregationWorkbenchState.havingMode = control.value || "all";
      const valueInput = aggregationWorkbenchSummaryEl.querySelector("[data-aggregation-control=\"having-value\"]");
      if (valueInput) {
        valueInput.style.display = aggregationWorkbenchState.havingMode === "all" ? "none" : "inline-block";
      }
    }
    if (kind === "having-value") {
      const next = parseFloat(control.value || "0", 10);
      aggregationWorkbenchState.havingValue = Number.isFinite(next) && next >= 0 ? next : 10;
    }
    renderAggregationWorkbench();
    if (sidebarEl) sidebarEl.scrollTop = savedSidebarScroll;
  });

  aggregationWorkbenchSummaryEl.addEventListener("click", (e) => {
    if (e.target.closest("[data-aggregation-control=\"measure-pick\"]")) {
      e.stopPropagation();
      openMeasurePicker();
      return;
    }
    if (e.target.closest("[data-aggregation-control=\"groupby-pick\"]")) {
      e.stopPropagation();
      openGroupByPicker();
    }
  });
}
if (aggregationWorkbenchListEl) {
  aggregationWorkbenchListEl.addEventListener("change", (e) => {
    const opsToggle = e.target.closest("[data-aggregation-control='result-search-ops']");
    if (opsToggle) {
      aggregationWorkbenchState.resultSearchOperators = !!opsToggle.checked;
      // Zapytanie zostaje — zmienia się tylko sposób jego czytania, więc od razu
      // widać różnicę na tej samej frazie.
      renderAggregationWorkbench();
      return;
    }
    const control = e.target.closest("[data-aggregation-control='match']");
    if (!control) return;
    aggregationWorkbenchState.matchMode = control.value || "contains";
    // Zsynchronizuj wszystkie inne selekty match w kartach
    aggregationWorkbenchListEl.querySelectorAll("[data-aggregation-control='match']").forEach((sel) => {
      sel.value = aggregationWorkbenchState.matchMode;
    });
  });
  aggregationWorkbenchListEl.addEventListener("keydown", (e) => {
    if (e.target.classList.contains("aggregation-result-search") && e.key === "Enter") {
      e.preventDefault();
      aggregationWorkbenchState.resultSearch = e.target.value || "";
      renderAggregationWorkbench();
    }
  });
  aggregationWorkbenchListEl.addEventListener("click", (e) => {
    e.stopPropagation();
    const btn = e.target.closest("button[data-aggregation-action='filter-group']");
    if (!btn) return;
    const value = (btn.dataset.aggregationValue || "").trim();
    if (!value) return;
    searchQueryEl.value = value;
    if (filterModeEl) {
      filterModeEl.value = aggregationWorkbenchState.matchMode === "exact" ? "equals" : "contains";
    }
    filtersCommitted = true;
    applyFilters();
    sortRows();
    scheduleViewRefresh({ table: true, analyses: true, formula: true, filterBadge: true });
    toast(t("filteredFor", { value }), "info");
  });
}
if (columnProfilerEl) {
  columnProfilerEl.addEventListener("click", (e) => {
    const btn = e.target.closest("button[data-profile-col-index]");
    if (!btn) return;
    const colIdx = parseInt(btn.dataset.profileColIndex || "", 10);
    if (!Number.isFinite(colIdx)) return;
    focusColumnProfile(colIdx);
  });
}
if (sheetInspectorSummaryEl) {
  sheetInspectorSummaryEl.addEventListener("click", (e) => {
    const btn = e.target.closest("button[data-inspector-action]");
    if (!btn) return;
    const action = btn.dataset.inspectorAction;

    if (action === "set-header") {
      const headerRow = parseInt(btn.dataset.inspectorHeaderRow || "", 10);
      if (!Number.isFinite(headerRow)) return;
      if (autoHeaderRowEl) autoHeaderRowEl.checked = false;
      headerRowEl.value = String(headerRow);
      loadBtn.click();
      return;
    }

    if (action === "toggle-long" && canUseLongView()) {
      tableViewMode = tableViewMode === "long" ? "wide" : "long";
      manualColumnWidths = {};
      withSceneTransition(() => {
        scheduleViewRefresh({ table: true, analyses: true, sync: true });
      });
      toast(tableViewMode === "long" ? t("wideLongOn") : t("wideLongOff"), "info");
      return;
    }

    if (action === "focus-col") {
      const colIdx = parseInt(btn.dataset.profileColIndex || "", 10);
      if (!Number.isFinite(colIdx)) return;
      focusColumnProfile(colIdx);
    }
  });
}
if (formulaWorkbenchListEl) {
  formulaWorkbenchListEl.addEventListener("click", (e) => {
    const btn = e.target.closest("button[data-formula-address]");
    if (!btn) return;
    focusFormulaEntry(btn.dataset.formulaAddress || "");
  });
}
if (kpiListEl) {
  kpiListEl.addEventListener("click", (e) => {
    const btn = e.target.closest("button[data-kpi-address]");
    if (!btn) return;
    focusKpiEntry(btn.dataset.kpiAddress || "");
  });
}
if (wideLongToggleEl) {
  wideLongToggleEl.addEventListener("click", () => {
    if (!canUseLongView()) return;
    tableViewMode = tableViewMode === "long" ? "wide" : "long";
    manualColumnWidths = {};
    withSceneTransition(() => {
      scheduleViewRefresh({ table: true, analyses: true, sync: true });
    });
    toast(tableViewMode === "long" ? t("wideLongOn") : t("wideLongOff"), "info");
  });
}
if (freezeHeadersEl) {
  freezeHeadersEl.addEventListener("change", () => {
    applyFreezeHeaders();
    toast(freezeHeadersEl.checked ? t("freezeHeadersOn") : t("freezeHeadersOff"), "info");
  });
  applyFreezeHeaders();
}
[showFontColorsEl, showCellFillsEl, showCellFontsEl, showCellBordersEl, showConditionalFormattingEl, showSubheadersEl, smartColWidthsEl].forEach((el) => {
  if (!el) return;
  el.addEventListener("change", () => {
    syncCellStyleFlags();
    saveCellStylePreferences();
    renderActiveTable();
  });
});
if (wrapCellsEl) {
  wrapCellsEl.addEventListener("change", () => {
    applyWrapCells();
    saveCellStylePreferences();
  });
}
if (rowHeightAllEl) {
  let _rowHeightRenderTid;
  rowHeightAllEl.addEventListener("input", () => {
    applyRowHeightAllPreference();
    clearTimeout(_rowHeightRenderTid);
    _rowHeightRenderTid = setTimeout(renderActiveTable, 200); // debounce — pełny render tylko po pauzie wpisywania
  });
}
if (colWidthAllEl) {
  let _colWidthRenderTid;
  colWidthAllEl.addEventListener("input", () => {
    applyColWidthAllPreference();
    clearTimeout(_colWidthRenderTid);
    _colWidthRenderTid = setTimeout(renderActiveTable, 200);
  });
}
if (freezeFirstColEl) {
  freezeFirstColEl.addEventListener("change", () => {
    applyFreezeFirstColumn();
    saveCellStylePreferences();
  });
}
if (recalcDatesEl) {
  recalcDatesEl.addEventListener("change", () => {
    syncCellStyleFlags();
    saveCellStylePreferences();
    rebuildCurrentSheetData(); // przeliczanie dzieje się w buildRows → trzeba odbudować dane
  });
}
// „Podświetl pasujące komórki" — wspólny stan dla filtra tekstowego i dat (oba checkboxy
// trzymane w zgodzie); podświetla od razu, ale nie ukrywa wierszy dopóki nie kliknięto „Filtruj".
[highlightMatchCellsEl, highlightMatchCellsDateEl].forEach((el) => {
  if (!el) return;
  el.addEventListener("change", () => {
    const on = !!el.checked;
    if (highlightMatchCellsEl) highlightMatchCellsEl.checked = on;
    if (highlightMatchCellsDateEl) highlightMatchCellsDateEl.checked = on;
    if (!currentHeaders.length) return;
    applyFilters();
    sortRows();
    renderActiveTable();
    updateFilterBadge();
  });
});
window.addEventListener("resize", () => {
  syncTableViewportHeight();
  syncFrozenHeaderMetrics();
}, { passive: true });
window.addEventListener("orientationchange", () => {
  window.setTimeout(() => {
    syncTableViewportHeight();
    syncFrozenHeaderMetrics();
  }, 120);
}, { passive: true });
syncTableViewportHeight();
if (readingToggle) {
  readingToggle.addEventListener("click", () => {
    const enabled = !rootEl.classList.contains("reading");
    setReadingMode(enabled);
  });
}
const _debouncedRenderFormula = (() => {
  let _tid;
  return () => { clearTimeout(_tid); _tid = setTimeout(renderFormulaWorkbench, 280); };
})();
[formulaSearchEl, formulaFilterEl, formulaFunctionFilterEl].forEach((el) => {
  if (!el) return;
  el.addEventListener("input", _debouncedRenderFormula);
  el.addEventListener("change", renderFormulaWorkbench);
});

document.addEventListener("click", (e) => {
  if (!isSidebarOpen()) return;
  // Klik poza panelem zamyka go TYLKO w trybie drawer (telefon / tablet w pionie).
  // Na desktopie i tablecie w poziomie panel jest stały — klik w tabelę go nie chowa.
  if (typeof sidebarIsDrawer === "function" && !sidebarIsDrawer()) return;
  if (sidebarEl && sidebarEl.contains(e.target)) return;
  if (panelToggle && panelToggle.contains(e.target)) return;
  if (panelHandle && panelHandle.contains(e.target)) return;
  if (columnPickerEl && !columnPickerEl.classList.contains("hidden") && columnPickerEl.contains(e.target)) return;
  if (quickSearchPopupEl && !quickSearchPopupEl.classList.contains("hidden") && quickSearchPopupEl.contains(e.target)) return;
  setSidebarOpen(false);
});


dropZone.addEventListener("dragover", (e) => {
  e.preventDefault();
  dropZone.classList.add("dragover");
});

dropZone.addEventListener("dragleave", () => {
  dropZone.classList.remove("dragover");
});

dropZone.addEventListener("drop", (e) => {
  e.preventDefault();
  dropZone.classList.remove("dragover");
  const file = e.dataTransfer.files[0];
  handleFile(file); // drag-drop nie daje uchwytu FSA → zapis przez picker / pobranie
});

// Gdy dostępne File System Access API, kliknięcie/aktywacja strefy otwiera plik
// przez showOpenFilePicker (zwraca uchwyt → zapis w miejscu). Bez FSA: natywny <input>.
if (canOpenFSA) {
  dropZone.addEventListener("click", (e) => {
    e.preventDefault();
    openWorkbookViaFsa();
  });
}

dropZone.addEventListener("keydown", (e) => {
  if (e.key === "Enter" || e.key === " ") {
    e.preventDefault();
    if (canOpenFSA) openWorkbookViaFsa();
    else fileInput.click();
  }
});

// Przycisk „Wybierz plik z dysku" w empty state — ten sam flow co drop-zone,
// żeby plik dało się wczytać bez wysuwania sidebara.
const emptyOpenBtn = document.getElementById("emptyOpenBtn");
if (emptyOpenBtn) {
  emptyOpenBtn.addEventListener("click", () => {
    if (canOpenFSA) openWorkbookViaFsa();
    else fileInput.click();
  });
}

sheetSelect.addEventListener("change", () => {
  if (!workbook) return;
  setStatus("Wybrano arkusz");
  applyAutoHeaderRowIfEnabled();
});

if (autoHeaderRowEl) {
  autoHeaderRowEl.addEventListener("change", () => {
    if (applyAutoHeaderRowIfEnabled()) {
      toast(t("headerDetected"), "info");
    }
  });
}

// ── Skoki między regionami (obsługa bez myszy) ───────────────────────────────
// Przy 17 panelach w sidebarze jedyną drogą do tabeli był Tab przez wszystko.
// Główne skróty siedzą na CYFRACH, nie na F1-F12: na klawiaturach do tabletów rząd
// funkcyjny bywa dostępny dopiero przez Fn (albo go nie ma). F6 zostaje jako alias
// dla desktopu. Rozpoznajemy po e.code (fizyczny klawisz) — e.key przy Alt potrafi
// zwrócić złożony znak (macOS: Option+s → "ß"), a na polskim układzie AltGr+litera
// to ą/ó/ś, więc litery pod Altem odpadają z definicji.
const APP_REGIONS = ["panel", "toolbar", "grid"];
const REGION_DIGIT_CODES = { Digit1: "panel", Digit2: "toolbar", Digit3: "grid" };

function firstFocusableIn(container) {
  if (!container) return null;
  const candidates = container.querySelectorAll(
    "summary, button:not([disabled]), input:not([type=hidden]):not([disabled]), select:not([disabled]), textarea:not([disabled]), [tabindex]:not([tabindex^='-'])"
  );
  for (const el of candidates) {
    if (el.closest(".hidden")) continue;
    if (el.offsetParent === null && getComputedStyle(el).position !== "fixed") continue;
    return el;
  }
  return null;
}

function currentAppRegion() {
  const active = document.activeElement;
  if (!active || active === document.body) return null;
  if (sidebarEl && sidebarEl.contains(active)) return "panel";
  if (tbodyEl && tbodyEl.contains(active)) return "grid";
  const toolbar = document.querySelector(".table-toolbar");
  if (toolbar && toolbar.contains(active)) return "toolbar";
  return null;
}

// Zwraca PRAWDĘ dopiero gdy fokus faktycznie usiadł na celu. To nie jest ozdobnik:
// element bywa niewidoczny (display:none), a wtedy .focus() jest cichym no-op.
// Ślepe zwracanie true zatrzymywało cykl F6 na pustym regionie — na desktopie pasek
// nad tabelą jest ukryty (szybkie szukanie żyje w trybie czytania, a #toolbarToggle
// to kontrolka mobilna), więc F6 grzązł i nigdy nie dochodził do tabeli.
function focusIfPossible(el) {
  if (!el) return false;
  el.focus();
  return document.activeElement === el;
}

function focusAppRegion(region) {
  if (region === "panel") {
    if (typeof isSidebarOpen === "function" && !isSidebarOpen()) setSidebarOpen(true);
    // setSidebarOpen zdejmuje inert synchronicznie, więc focus łapie od razu —
    // nie trzeba czekać na koniec animacji wsuwania.
    return focusIfPossible(firstFocusableIn(sidebarEl));
  }
  if (region === "toolbar") {
    // Najpierw szybkie szukanie, jeśli pasek jest odsłonięty; inaczej pierwsza
    // widoczna kontrolka paska nad tabelą (Wide-to-Long, układ z Excela itd.).
    const preferred = quickSearchWrap && !quickSearchWrap.classList.contains("hidden")
      ? firstFocusableIn(quickSearchWrap)
      : null;
    return focusIfPossible(preferred || firstFocusableIn(document.querySelector(".table-toolbar")));
  }
  if (region === "grid") {
    if (!tbodyEl) return false;
    const cell = (typeof findCellElement === "function" ? findCellElement(focusedCellState) : null)
      || tbodyEl.querySelector("tr[data-row-key] td[data-col-index]");
    if (!cell) return false;
    // Bez tabindex focus() na <td> jest no-op — upewnij się, że to właśnie ta komórka
    // trzyma teraz punkt wejścia siatki.
    if (typeof syncGridRovingTabindex === "function") syncGridRovingTabindex(cell);
    return focusIfPossible(cell);
  }
  return false;
}

// F6 / Shift+F6: kolejny (poprzedni) region względem tego, w którym stoi fokus.
// Regiony bez treści (np. pasek szukania przed wczytaniem pliku) przeskakujemy.
function cycleAppRegion(dir) {
  const len = APP_REGIONS.length;
  const from = currentAppRegion();
  const start = from ? APP_REGIONS.indexOf(from) : (dir > 0 ? -1 : 0);
  for (let step = 1; step <= len; step++) {
    const idx = ((start + dir * step) % len + len) % len;
    if (focusAppRegion(APP_REGIONS[idx])) return true;
  }
  return false;
}

// Skip-link: pierwszy przystanek Tab na stronie skacze wprost na aktywną komórkę.
// Gdy arkusza jeszcze nie ma, oddaj fokus panelowi (tam jest wczytywanie pliku).
const skipToTableBtn = document.getElementById("skipToTable");
if (skipToTableBtn) {
  skipToTableBtn.addEventListener("click", () => {
    if (!focusAppRegion("grid")) focusAppRegion("panel");
  });
}

document.addEventListener("keydown", (e) => {
  const meta = e.ctrlKey || e.metaKey;
  const popupOpen = typeof isQuickSearchPopupOpen === "function"
    ? isQuickSearchPopupOpen()
    : !!(quickSearchPopupEl && !quickSearchPopupEl.classList.contains("hidden"));

  // Popup szybkiego szukania ma priorytet: Enter = Szukaj (niezależnie od fokusu),
  // strzałki ↓/↑ = live wyniki, Esc obsługiwane niżej. Bez tego Enter na „luźnym"
  // fokusie otwierał edytor komórki zamiast szukać.
  if (popupOpen && !meta && !e.altKey) {
    if (e.key === "ArrowDown" || e.key === "ArrowUp") {
      // Tylko z pola szukania albo z samej listy trafień. Wcześniej ten warunek łapał
      // KAŻDY cel przy otwartej liście — więc ↓/↑ na <select> „Tryb"/„Akcja" nie
      // zmieniało opcji, tylko uciekało do wyników. Na klawiaturze bez myszy te dwa
      // selecty były przez to praktycznie nie do przestawienia.
      const target = e.target;
      const fromQsInput = target === quickSearchPopupInput || target === quickSearchEl;
      const fromLiveItem = !!(target && typeof target.closest === "function" && target.closest(".qs-live-item"));
      if ((fromQsInput || fromLiveItem)
        && typeof navigateQsLiveResults === "function"
        && navigateQsLiveResults(e.key === "ArrowDown" ? 1 : -1)) {
        e.preventDefault();
        return;
      }
    }
    if (e.key === "Enter") {
      // Checkbox/przycisk/live-hit: własna aktywacja (tablet + klawiatura bez myszy).
      if (typeof qsEnterShouldActivateControl === "function" && qsEnterShouldActivateControl(e.target)) {
        if (typeof qsActivateFocusedControl === "function" && qsActivateFocusedControl(e.target)) {
          e.preventDefault();
        }
        return;
      }
      e.preventDefault();
      if (typeof commitQuickSearch === "function") commitQuickSearch();
      return;
    }
  }

  // `/` — szybkie otwarcie popup szukania (gdy nie piszesz w polu/edytorze
  // i nie masz aktywnej komórki — wtedy `/` ma iść do edytora jak inny znak).
  if (!popupOpen && !meta && !e.altKey && e.key === "/" && !shouldIgnoreTableArrowNavigation() && !focusedCellState) {
    if (currentHeaders.length && quickSearchPopupEl && quickSearchPopupInput) {
      e.preventDefault();
      openQuickSearchPopup();
      return;
    }
  }

  if (!meta && !e.altKey && !shouldIgnoreTableArrowNavigation()) {
    let handled = false;
    // UWAGA: `e.shiftKey` jest prawdziwe także dla samego klawisza Shift, więc bez tej
    // bramki gołe przytrzymanie Shifta zakładało zaznaczenie 1×1 i przełączało tryb
    // na komórkowy, zanim padła jakakolwiek strzałka.
    const isArrowKey = e.key === "ArrowUp" || e.key === "ArrowDown"
      || e.key === "ArrowLeft" || e.key === "ArrowRight";
    if (e.shiftKey && isArrowKey) {
      // Zakres jest z natury operacją na komórkach — Shift zawsze schodzi na ten poziom.
      if (!selectedCellState && focusedCellState) {
        setSelectedCell(focusedCellState.rowKey, focusedCellState.colIndex0, { scroll: false });
      }
      setSelectionKind("cell", { repaint: false });
      if (e.key === "ArrowUp") handled = moveSelectedCell(-1, 0);
      if (e.key === "ArrowDown") handled = moveSelectedCell(1, 0);
      if (e.key === "ArrowLeft") handled = moveSelectedCell(0, -1);
      if (e.key === "ArrowRight") handled = moveSelectedCell(0, 1);
    } else if (e.shiftKey) {
      // Shift bez strzałki — nic tu po nas (Shift+Spacja obsłużona niżej).
    } else if (!isCellSelectionMode() && (e.key === "ArrowLeft" || e.key === "ArrowRight")) {
      // Zaznaczony jest cały wiersz — w bok nie ma czym ruszać, więc ←/→ przewijają
      // widok. Zaznaczenie zostaje nietknięte (patrz selectionKind w core.js).
      handled = scrollTableHorizontally(e.key === "ArrowRight" ? 1 : -1);
    } else {
      if (e.key === "ArrowUp") handled = moveFocusedCell(-1, 0);
      if (e.key === "ArrowDown") handled = moveFocusedCell(1, 0);
      if (e.key === "ArrowLeft") handled = moveFocusedCell(0, -1);
      if (e.key === "ArrowRight") handled = moveFocusedCell(0, 1);
    }
    if (handled) {
      e.preventDefault();
      if (e.shiftKey) clearTextSelection(); // Shift+strzałki nie zostawia zaznaczonego tekstu
      return;
    }
  }
  // Ctrl/⌘+C / Ctrl/⌘+V — kopiuj/wklej zaznaczony zakres komórek (TSV, kompatybilne
  // z Excelem/Arkuszami). Tylko gdy komórka ma fokus i nie jesteśmy w polu tekstowym
  // (inaczej normalne kopiowanie/wklejanie tekstu przestałoby działać wszędzie indziej).
  if (meta && !e.altKey && !e.shiftKey && focusedCellState && !shouldIgnoreTableArrowNavigation()) {
    if (e.key.toLowerCase() === "c") {
      e.preventDefault();
      copySelectionToClipboard();
      return;
    }
    if (e.key.toLowerCase() === "v") {
      e.preventDefault();
      pasteClipboardToSelection();
      return;
    }
  }
  // Shift+Spacja — przełącznik poziomu zaznaczenia: cały wiersz ↔ sama komórka.
  // Jedyne wejście w tryb komórki bez myszy (Shift+klik wymaga wskaźnika), a przy
  // okazji zgodne z arkuszami, gdzie Shift+Space zaznacza wiersz. MUSI stać przed
  // gałęzią „znak drukowalny otwiera edytor" — spacja to też znak o długości 1.
  if (!meta && !e.altKey && e.shiftKey && e.key === " "
    && focusedCellState && !shouldIgnoreTableArrowNavigation()) {
    e.preventDefault();
    // Kierunek ustalamy PRZED przestawieniem — isCellSelectionMode() czyta już nową
    // wartość, więc pytanie o nią po zmianie dawało odwrotny wynik.
    const goingToCell = !isCellSelectionMode();
    setSelectionKind(goingToCell ? "cell" : "row", { repaint: false });
    // Powrót do wiersza zwija zakres. setSelectedCell(null) nie odświeża podświetlenia
    // wiersza, więc przemalowanie robimy raz, po ustaleniu OBU stanów.
    if (!goingToCell) setSelectedCell("", -1);
    syncFocusedCellInDom({ clearMissing: false });
    return;
  }
  // Enter lub znak drukowalny na zaznaczonej komórce otwiera edytor (jak w Excelu).
  if (!meta && !e.altKey && focusedCellState && !shouldIgnoreTableArrowNavigation()) {
    if (e.key === "Enter") {
      const td = findCellElement(focusedCellState);
      if (td) {
        e.preventDefault();
        openCellEditor(td);
        return;
      }
    } else if (e.key.length === 1) {
      const td = findCellElement(focusedCellState);
      if (td) {
        e.preventDefault();
        openCellEditor(td, { initialChar: e.key });
        return;
      }
    }
  }
  if (meta && e.key === "Enter") {
    e.preventDefault();
    applyFilterBtn.click();
  }
  if (meta && e.altKey && e.key.toLowerCase() === "s") {
    e.preventDefault();
    saveAsBtn.click();
  }
  // Ctrl/⌘+S: zapis w miejscu (gdy włączony), inaczej "Zapisz jako…".
  if (meta && !e.altKey && !e.shiftKey && e.key.toLowerCase() === "s") {
    e.preventDefault();
    if (saveBtn && !saveBtn.disabled) saveBtn.click();
    else saveAsBtn.click();
  }
  if (meta && e.shiftKey && e.key.toLowerCase() === "e") {
    e.preventDefault();
    exportCsvBtn.click();
  }
  if (meta && e.shiftKey && e.key.toLowerCase() === "x") {
    e.preventDefault();
    resetFiltersBtn.click();
  }
  if (meta && e.altKey && e.key.toLowerCase() === "w") {
    e.preventDefault();
    resetWidthsBtn.click();
  }
  if (meta && e.altKey && e.key.toLowerCase() === "k") {
    e.preventDefault();
    lastPickerTriggerEl = filter1PickBtn;
    openColumnPicker("filter1");
  }
  // Ctrl/⌘+Alt+1/2/3 — skok wprost do panelu / paska szukania / tabeli.
  // F6 i Shift+F6 robią to samo cyklicznie (alias dla pełnowymiarowych klawiatur).
  if (meta && e.altKey && !e.shiftKey && REGION_DIGIT_CODES[e.code]) {
    if (focusAppRegion(REGION_DIGIT_CODES[e.code])) {
      e.preventDefault();
      return;
    }
  }
  if (e.key === "F6" && !meta && !e.altKey) {
    if (cycleAppRegion(e.shiftKey ? -1 : 1)) {
      e.preventDefault();
      return;
    }
  }
  if (meta && e.key === "/") {
    e.preventDefault();
    themeToggle.click();
  }
  if (meta && e.shiftKey && e.key.toLowerCase() === "f") {
    e.preventDefault();
    if (quickSearchPopupEl && !quickSearchPopupEl.classList.contains("hidden")) {
      closeQuickSearchPopup();
    } else if (currentHeaders.length && quickSearchPopupEl && quickSearchPopupInput) {
      openQuickSearchPopup();
    } else if (!currentHeaders.length) {
      toast(t("loadSheetToSearch"), "info");
    }
  }
  // Escape najpierw zamyka NAKŁADKI (modal / okno szukania), a dopiero potem rusza
  // zaznaczenie w tabeli. Odwrotna kolejność powodowała, że przy zaznaczonej komórce
  // Esc odznaczał ją i robił `return` — okno szukania zostawało otwarte i nie dało się
  // go zamknąć z klawiatury bez wcześniejszego odznaczenia komórki.
  if (e.key === "Escape" && !columnPickerEl.classList.contains("hidden")) {
    e.preventDefault();
    closeColumnPicker();
    return;
  }
  if (e.key === "Escape" && exportModalEl && !exportModalEl.classList.contains("hidden")) {
    e.preventDefault();
    closeExportModal();
    return;
  }
  if (e.key === "Escape" && quickSearchPopupEl && !quickSearchPopupEl.classList.contains("hidden")) {
    e.preventDefault();
    closeQuickSearchPopup(); // oddaje fokus tam, skąd otwarto okno
    return;
  }

  // Odznaczanie jak w arkuszach (tylko Escape — żadnych liter, by nie kolidowały
  // z wpisywaniem do komórki). Shift+Esc = pełne odznaczenie; Esc = progresywnie:
  // najpierw zwiń zakres do aktywnej komórki, dopiero potem zdejmij fokus wiersza.
  if (e.key === "Escape" && (selectedCellState || focusedCellState)) {
    if (e.shiftKey) {
      e.preventDefault();
      setSelectedCell("", -1);
      setFocusedCell("", -1);
      return;
    }
    if (hasActiveCellRange()) {
      e.preventDefault();
      setSelectedCell("", -1); // zwiń zakres, zostaw aktywną komórkę (fokus)
      return;
    }
    if (focusedCellState) {
      e.preventDefault();
      setFocusedCell("", -1);
      return;
    }
  }

  if (e.key === "Escape" && isSidebarOpen()) {
    setSidebarOpen(false);
  }
});

setEmptyState(DEFAULT_EMPTY_TITLE, DEFAULT_EMPTY_SUB);
updateDateChipsActive();
updateQuickSearchColumnButtons();
updateSortControls();
setDirtyState(false);
syncQuickSearchInputs();
// Sidebar otwarty na starcie tylko na większych ekranach. Na telefonie zasłaniałby
// cały widok, a jego późniejsze domknięcie + przeskok uchwytu na dolny pasek
// generowały duży layout shift (CLS) przy ładowaniu. Plik można wczytać
// przyciskiem w pustym stanie — sidebar nie jest potrzebny od pierwszej klatki.
setSidebarOpen(window.matchMedia("(min-width: 769px)").matches);
syncSidebarHandle();
// Bez wczytanego arkusza panele analiz zostają „brudne" — dorenderują się przy
// otwarciu <details> (renderDirtyAnalysesForPanel) lub po loadBtn. Oszczędza boot.
if (currentHeaders.length) {
  scheduleViewRefresh({ analyses: true, formula: true });
}
populateSortColumnSelect();
populateEditColumnSelect();
renderSortPresets();
updateWideLongToggle();

// Z File System Access API "Zapisz" nadpisuje plik w miejscu — odblokuj przycisk.
// Bez FSA pozostaje wyłączony (fallback: "Zapisz jako…"). Tytuł ustawia language.js.
if (canFSA && saveBtn) {
  saveBtn.disabled = false;
  saveBtn.removeAttribute("aria-disabled");
}

// Biblioteki XLSX/JSZip dogrywane są LENIWIE (ensureXlsxLibs) przy pierwszym
// wczytaniu/zapisie pliku — przy starcie traktujemy runtime jako dostępny, żeby
// UI (input pliku, przyciski, dropzone) NIE był zablokowany. Realny brak (np.
// offline bez cache) zgłosi bramka ensureXlsxLibs(true) w handleFile/loadSampleFile.
setRuntimeAvailability(true);

// Anty-lag pierwszego otwarcia panelu: w bezczynności po starcie wymuszamy jednorazową
// rasteryzację wysuwanego panelu (niewidocznie, pod treścią). Bez tego PIERWSZE otwarcie
// maluje całe poddrzewo paneli/SVG w jednej klatce (~150ms zacięcia); kolejne są już płynne.
//
// WAŻNE (fix glitcha „przezroczysty panel po wczytaniu arkusza"): pre-warm zakłada panel
// SCHOWANY. Pomijamy go, gdy:
//   • sidebar jest OTWARTY/widoczny (desktop) — `.prewarm` (opacity:0.001, z-index:-1)
//     zrobiłby z widocznego panelu przezroczystą dziurę za treścią;
//   • arkusz jest już wczytany (po interakcji) — wtedy idle-callback potrafił odpalić się
//     późno, w trakcie ciężkiego renderu, a zdjęcie klasy przez rAF było głodzone → panel
//     zostawał przezroczysty aż do następnego kliknięcia.
// Dodatkowo TWARDY bezpiecznik (setTimeout) zdejmuje klasę, nawet gdyby rAF utknął.
function prewarmSidebar() {
  const sidebar = document.querySelector(".sidebar");
  if (!sidebar) return;
  if ((typeof isSidebarOpen === "function" && isSidebarOpen()) || workbook) return;
  sidebar.classList.add("prewarm");
  const clear = () => sidebar.classList.remove("prewarm");
  requestAnimationFrame(() => requestAnimationFrame(clear));
  window.setTimeout(clear, 250); // bezpiecznik: .prewarm nigdy nie utknie
}
// timeout na idle-callback, żeby nie odpalił się „kiedyś później" w trakcie wczytywania
(window.requestIdleCallback
  ? (fn) => window.requestIdleCallback(fn, { timeout: 1200 })
  : (fn) => window.setTimeout(fn, 200))(prewarmSidebar);

window.addEventListener("beforeunload", (e) => {
  if (!hasUnsavedChanges) return;
  e.preventDefault();
  e.returnValue = "";
});

if ("serviceWorker" in navigator) {
  let waitingServiceWorker = null;
  let refreshingForUpdate = false;

  const showAppUpdate = (worker) => {
    waitingServiceWorker = worker;
    if (!appUpdateBtn) return;
    appUpdateBtn.classList.remove("hidden");
    appUpdateBtn.textContent = t("updateNow");
    // podpowiedź daje statyczny data-hint-pl/en w HTML (cursor-hint), nie natywny title
    toast(t("updateAvailable"), "info");
  };

  if (appUpdateBtn) {
    appUpdateBtn.addEventListener("click", () => {
      if (!waitingServiceWorker) {
        hardRefreshApp();
        return;
      }
      appUpdateBtn.disabled = true;
      appUpdateBtn.textContent = t("refreshingApp");
      waitingServiceWorker.postMessage({ type: "SKIP_WAITING" });
    });
  }

  // Czy strona była JUŻ kontrolowana przez SW w chwili startu. Jeśli tak, to
  // późniejszy controllerchange = wymiana SW = AKTUALIZACJA → przeładuj. Jeśli nie
  // (pierwsze wejście), controllerchange pochodzi z clients.claim() świeżo
  // zainstalowanego SW — to NIE jest aktualizacja, więc NIE przeładowujemy
  // (inaczej każde pierwsze otwarcie robi zbędny pełny reload/mignięcie).
  const hadControllerAtStart = !!navigator.serviceWorker.controller;
  navigator.serviceWorker.addEventListener("controllerchange", () => {
    if (refreshingForUpdate) return;
    if (!hadControllerAtStart) return; // pierwsze przejęcie (claim), nie aktualizacja
    refreshingForUpdate = true;
    window.location.reload();
  });

  navigator.serviceWorker.register(`sw.js?v=${APP_BUILD_VERSION}`).then((registration) => {
    if (registration.waiting && navigator.serviceWorker.controller) {
      showAppUpdate(registration.waiting);
    }

    registration.addEventListener("updatefound", () => {
      const worker = registration.installing;
      if (!worker) return;
      worker.addEventListener("statechange", () => {
        if (worker.state === "installed" && navigator.serviceWorker.controller) {
          showAppUpdate(worker);
        }
      });
    });

    registration.update().catch(() => {});

    // PWA wznowiona z tła / trzymana otwarta: sprawdź nową wersję przy powrocie na
    // pierwszy plan oraz okresowo. Bez tego użytkownik z ikony na ekranie głównym
    // siedzi na starym buildzie aż do pełnego przeładowania. update() po znalezieniu
    // nowego sw.js sam odpali updatefound → showAppUpdate (przycisk + toast).
    const checkForUpdate = () => registration.update().catch(() => {});
    document.addEventListener("visibilitychange", () => {
      if (document.visibilityState === "visible") checkForUpdate();
    });
    window.setInterval(checkForUpdate, 30 * 60 * 1000);
  }).catch(() => {});
}
