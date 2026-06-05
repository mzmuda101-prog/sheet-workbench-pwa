// App bootstrap: event wiring, initial render, and runtime startup.

if (window.MateuszCursorHint) {
  window.MateuszCursorHint.initCursorHints({
    fallbackHint: t("hintDefault"),
  });
}

panelToggle.addEventListener("click", toggleSidebar);
if (panelHandle) panelHandle.addEventListener("click", toggleSidebar);
if (sidebarScrim) sidebarScrim.addEventListener("click", () => setSidebarOpen(false));
document.querySelectorAll("details.panel").forEach((det) => {
  det.addEventListener("toggle", () => {
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
      renderActiveTable();
      renderSheetInspectorSummary();
      renderDurationAnalysis();
      renderAggregationWorkbench();
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
    renderDurationAnalysis();
    renderAggregationWorkbench();
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
    renderFormulaWorkbench();
    updateFilterBadge();
    toast(t("filteredFor", { value: entity }), "info");
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
    renderFormulaWorkbench();
    updateFilterBadge();
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
      renderActiveTable();
      renderSheetInspectorSummary();
      renderDurationAnalysis();
      renderAggregationWorkbench();
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
    renderActiveTable();
    renderDurationAnalysis();
    renderAggregationWorkbench();
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

document.addEventListener("keydown", (e) => {
  const meta = e.ctrlKey || e.metaKey;
  if (!meta && !e.altKey && !shouldIgnoreTableArrowNavigation()) {
    let handled = false;
    if (e.shiftKey) {
      if (!selectedCellState && focusedCellState) {
        setSelectedCell(focusedCellState.rowKey, focusedCellState.colIndex0, { scroll: false });
      }
      if (e.key === "ArrowUp") handled = moveSelectedCell(-1, 0);
      if (e.key === "ArrowDown") handled = moveSelectedCell(1, 0);
      if (e.key === "ArrowLeft") handled = moveSelectedCell(0, -1);
      if (e.key === "ArrowRight") handled = moveSelectedCell(0, 1);
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
  if (meta && e.shiftKey && e.key.toLowerCase() === "w") {
    e.preventDefault();
    resetWidthsBtn.click();
  }
  if (meta && e.key.toLowerCase() === "k") {
    e.preventDefault();
    lastPickerTriggerEl = filter1PickBtn;
    openColumnPicker("filter1");
  }
  if (meta && e.key === "/") {
    e.preventDefault();
    themeToggle.click();
  }
  if (meta && e.shiftKey && e.key.toLowerCase() === "f") {
    e.preventDefault();
    if (quickSearchPopupEl && !quickSearchPopupEl.classList.contains("hidden")) {
      quickSearchPopupEl.classList.add("hidden");
    } else if (currentHeaders.length && quickSearchPopupEl && quickSearchPopupInput) {
      quickSearchPopupInput.value = searchQueryEl.value || "";
      quickSearchPopupEl.classList.remove("hidden");
      quickSearchPopupInput.focus();
    } else if (!currentHeaders.length) {
      toast(t("loadSheetToSearch"), "info");
    }
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

  if (e.key === "Escape" && !columnPickerEl.classList.contains("hidden")) {
    e.preventDefault();
    closeColumnPicker();
    return;
  }
  if (e.key === "Escape" && quickSearchPopupEl && !quickSearchPopupEl.classList.contains("hidden")) {
    e.preventDefault();
    quickSearchPopupEl.classList.add("hidden");
    return;
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
setSidebarOpen(true);
syncSidebarHandle();
renderInsights();
renderKpiExtractor();
renderSheetInspectorSummary();
renderColumnProfiles();
renderSections();
renderRepeatingBlocks();
renderDurationAnalysis();
renderAggregationWorkbench();
renderFormulaWorkbench();
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

const xlsxReady = isXlsxAvailable(false);
setRuntimeAvailability(xlsxReady);
if (!xlsxReady) {
  setEmptyState(
    t("xlsxMissingStatus"),
    t("xlsxMissingEmpty")
  );
  setStatus(t("xlsxMissingStatus"));
}

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
    appUpdateBtn.setAttribute("title", t("updateReady"));
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

  navigator.serviceWorker.addEventListener("controllerchange", () => {
    if (refreshingForUpdate) return;
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
