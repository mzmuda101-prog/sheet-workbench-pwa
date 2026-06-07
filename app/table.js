// Table models, layout calculations, and DOM rendering for the main grid.

function canUseLongView() {
  const group = getActiveRepeatingGroup();
  return !!(group && Array.isArray(group.blocks) && group.blocks.length >= 2);
}

function updateWideLongToggle() {
  if (!wideLongToggleEl) return;
  const available = canUseLongView();
  if (!available) {
    tableViewMode = "wide";
  }
  wideLongToggleEl.classList.toggle("hidden", !available);
  wideLongToggleEl.setAttribute("aria-hidden", available ? "false" : "true");
  wideLongToggleEl.setAttribute("aria-pressed", tableViewMode === "long" ? "true" : "false");
  wideLongToggleEl.textContent = tableViewMode === "long" ? t("classicView") : "Wide-to-Long";
  wideLongToggleEl.title = tableViewMode === "long"
    ? t("backToClassicTitle")
    : t("switchToLongTitle");
}

function buildWideDisplayModelFromRows(rows, options = {}) {
  const headers = Array.isArray(options.headers) ? options.headers.slice() : currentHeaders.slice();
  const startCol = Number.isFinite(options.startCol) ? options.startCol : currentStartCol;
  const headerRow = Number.isFinite(options.headerRow) ? options.headerRow : currentHeaderRow;
  return {
    mode: "wide",
    headers,
    rows: rows.slice(),
    guideLabels: headers.map((_, i) => XLSX.utils.encode_col(i + startCol)),
    headerRowLabel: String(headerRow),
    rowHeadFormatter: (row) => String((row?.rowIndex0 ?? 0) + 1),
    editable: true,
  };
}

function buildWideDisplayModel() {
  return buildWideDisplayModelFromRows(viewRows);
}

function buildLongViewModelFromRows(rows, group = getActiveRepeatingGroup(), options = {}) {
  if (!group || !Array.isArray(group.blocks) || !group.blocks.length) return buildWideDisplayModelFromRows(rows);

  const firstBlock = group.blocks[0];
  const prefixCount = Math.max(0, Number(group.prefixCount) || 0);
  const sourceHeaders = Array.isArray(options.headers) ? options.headers.slice() : currentHeaders.slice();
  const headerRow = Number.isFinite(options.headerRow) ? options.headerRow : currentHeaderRow;
  const prefixHeaders = sourceHeaders.slice(0, prefixCount);
  const repeatedHeaders = Array.isArray(group.longHeaders) && group.longHeaders.length
    ? group.longHeaders.slice()
    : firstBlock.headers.map((header) => parseRepeatedHeader(header)?.base || cleanSectionLabel(header) || header);
  const headers = [...prefixHeaders, t("longColBlockNum"), t("longColBlock"), ...repeatedHeaders];
  const nextRows = [];

  rows.forEach((row) => {
    group.blocks.forEach((block, blockIndex) => {
      const valueIndexes = Array.isArray(block.valueIndexes) && block.valueIndexes.length
        ? block.valueIndexes
        : null;
      const blockValues = valueIndexes
        ? valueIndexes.map((idx) => idx >= 0 ? row.values[idx] : null)
        : row.values.slice(block.startIndex, block.endIndex + 1);
      const blockDisplay = valueIndexes
        ? valueIndexes.map((idx) => idx >= 0 ? getDisplayValue(row, idx) : "")
        : blockValues.map((_, idx) => getDisplayValue(row, block.startIndex + idx));
      const hasMeaningfulValue = blockDisplay.some((value) => String(value ?? "").trim() !== "");
      if (!hasMeaningfulValue) return;

      const prefixValues = row.values.slice(0, prefixCount);
      const prefixDisplay = prefixValues.map((_, idx) => getDisplayValue(row, idx));
      const values = [...prefixValues, blockIndex + 1, block.label, ...blockValues];
      const display = [...prefixDisplay, String(blockIndex + 1), block.label, ...blockDisplay];

      nextRows.push({
        values,
        rawValues: values.slice(),
        display,
        rowIndex0: row.rowIndex0,
        sourceRowIndex0: row.rowIndex0,
        sourceBlockIndex: blockIndex,
        sourceBlockLabel: block.label,
        isLongViewRow: true,
      });
    });
  });

  return {
    mode: "long",
    headers,
    rows: nextRows,
    guideLabels: headers.map((_, idx) => `${idx + 1}`),
    headerRowLabel: `${headerRow} -> long`,
    rowHeadFormatter: (row) => `${(row?.sourceRowIndex0 ?? row?.rowIndex0 ?? 0) + 1}.${(row?.sourceBlockIndex ?? 0) + 1}`,
    editable: false,
  };
}

function buildLongViewModel() {
  return buildLongViewModelFromRows(viewRows);
}

function getRowSelectionKey(row) {
  if (!row) return "";
  if (row.isLongViewRow) {
    return `long:${row.sourceRowIndex0 ?? row.rowIndex0}:${row.sourceBlockIndex ?? 0}`;
  }
  return `wide:${row.rowIndex0 ?? ""}`;
}

function buildFocusedRowStatusSuffix(model) {
  if (!focusedCellState || !model?.rows?.length) return "";
  const idx = model.rows.findIndex((row) => getRowSelectionKey(row) === focusedCellState.rowKey);
  if (idx < 0) return "";
  const limit = Math.max(1, parseInt(maxRowsEl.value || "200", 10));
  if (idx >= limit) return "";
  return t("statusFocusedRow", { pos: idx + 1 });
}

function updateTableStatus(model) {
  if (!model?.headers?.length) return;
  const rows = Array.isArray(model.rows) ? model.rows : [];
  if (!rows.length) return;
  const limit = Math.max(1, parseInt(maxRowsEl.value || "200", 10));
  const modeLabel = model.mode === "long" ? t("statusLongMode") : "";
  const focusedSuffix = buildFocusedRowStatusSuffix(model);
  setStatus(t("statusTableRows", {
    total: rows.length,
    shown: Math.min(rows.length, limit),
    mode: modeLabel,
    focused: focusedSuffix,
  }));
}

function findCellElement(cellState) {
  if (!cellState) return null;
  return tbodyEl.querySelector(
    `tr[data-row-key="${CSS.escape(cellState.rowKey)}"] td[data-col-index="${cellState.colIndex0}"]`
  );
}

function findFocusedRowElement() {
  if (!focusedCellState) return null;
  return tbodyEl.querySelector(`tr[data-row-key="${CSS.escape(focusedCellState.rowKey)}"]`);
}

function syncFocusedCellInDom(options = {}) {
  tbodyEl.querySelectorAll("tr.row-focused").forEach((row) => row.classList.remove("row-focused"));
  const rowEl = findFocusedRowElement();
  if (!rowEl) {
    if (options.clearMissing !== false) focusedCellState = null;
    return null;
  }
  rowEl.classList.add("row-focused");
  const cell = findCellElement(focusedCellState);
  if (options.scroll) {
    (cell || rowEl).scrollIntoView({ block: "nearest", inline: "nearest" });
  }
  return rowEl;
}

function syncSelectedCellInDom(options = {}) {
  tbodyEl.querySelectorAll("td.cell-selected").forEach((cell) => cell.classList.remove("cell-selected"));
  const cell = findCellElement(selectedCellState);
  if (!cell) {
    if (options.clearMissing !== false) selectedCellState = null;
    return null;
  }
  cell.classList.add("cell-selected");
  if (options.scroll) {
    cell.scrollIntoView({ block: "nearest", inline: "nearest" });
  }
  return cell;
}

function setFocusedCell(rowKey, colIndex0, options = {}) {
  if (!rowKey || !Number.isFinite(colIndex0) || colIndex0 < 0) {
    focusedCellState = null;
    syncFocusedCellInDom({ clearMissing: false });
    syncRangeHighlightInDom();
    updateCellStats();
    updateTableStatus(currentDisplayModel);
    return;
  }
  focusedCellState = { rowKey, colIndex0 };
  // Nowa aktywna komórka (kotwica) = świeży start: wyczyść poprzedni zakres,
  // chyba że jawnie go rozszerzamy (Shift → options.keepSelection). Dzięki temu
  // zwykły klik / strzałka daje pojedynczą komórkę, a nie zakres od starego końca.
  if (!options.keepSelection) selectedCellState = null;
  syncFocusedCellInDom(options);
  syncRangeHighlightInDom();
  updateCellStats();
  updateTableStatus(currentDisplayModel);
}

function setSelectedCell(rowKey, colIndex0, options = {}) {
  if (!rowKey || !Number.isFinite(colIndex0) || colIndex0 < 0) {
    selectedCellState = null;
    syncSelectedCellInDom({ clearMissing: false });
    syncRangeHighlightInDom();
    updateCellStats();
    return;
  }
  selectedCellState = { rowKey, colIndex0 };
  syncSelectedCellInDom(options);
  syncRangeHighlightInDom();
  updateCellStats();
}

function moveFocusedCell(rowDelta, colDelta) {
  if (!focusedCellState || !currentDisplayModel?.rows?.length || !currentDisplayModel?.headers?.length) return false;
  const rowIndex = currentDisplayModel.rows.findIndex((row) => getRowSelectionKey(row) === focusedCellState.rowKey);
  if (rowIndex < 0) {
    focusedCellState = null;
    return false;
  }
  const nextRowIndex = Math.max(0, Math.min(currentDisplayModel.rows.length - 1, rowIndex + rowDelta));
  const nextColIndex = Math.max(0, Math.min(currentDisplayModel.headers.length - 1, focusedCellState.colIndex0 + colDelta));
  const nextRow = currentDisplayModel.rows[nextRowIndex];
  setFocusedCell(getRowSelectionKey(nextRow), nextColIndex, { scroll: true });
  return true;
}

function moveSelectedCell(rowDelta, colDelta) {
  if (!selectedCellState || !currentDisplayModel?.rows?.length || !currentDisplayModel?.headers?.length) return false;
  const rowIndex = currentDisplayModel.rows.findIndex((row) => getRowSelectionKey(row) === selectedCellState.rowKey);
  if (rowIndex < 0) {
    selectedCellState = null;
    return false;
  }
  const nextRowIndex = Math.max(0, Math.min(currentDisplayModel.rows.length - 1, rowIndex + rowDelta));
  const nextColIndex = Math.max(0, Math.min(currentDisplayModel.headers.length - 1, selectedCellState.colIndex0 + colDelta));
  const nextRow = currentDisplayModel.rows[nextRowIndex];
  setSelectedCell(getRowSelectionKey(nextRow), nextColIndex, { scroll: true });
  return true;
}

// --- Pasek statystyk komórek (Excel-style) ------------------------------
// Zakres = prostokąt od kotwicy (focusedCellState) do ruchomego końca
// (selectedCellState). Istnieje tylko gdy obie komórki są ustawione.
function getSelectionRectangle() {
  if (!focusedCellState || !selectedCellState) return null;
  const model = currentDisplayModel;
  if (!model?.rows?.length) return null;
  const anchorRowIdx = model.rows.findIndex((r) => getRowSelectionKey(r) === focusedCellState.rowKey);
  const endRowIdx = model.rows.findIndex((r) => getRowSelectionKey(r) === selectedCellState.rowKey);
  if (anchorRowIdx < 0 || endRowIdx < 0) return null;
  const rowStart = Math.min(anchorRowIdx, endRowIdx);
  const rowEnd = Math.max(anchorRowIdx, endRowIdx);
  const colMin = Math.min(focusedCellState.colIndex0, selectedCellState.colIndex0);
  const colMax = Math.max(focusedCellState.colIndex0, selectedCellState.colIndex0);
  const rowKeys = new Set();
  for (let i = rowStart; i <= rowEnd; i++) rowKeys.add(getRowSelectionKey(model.rows[i]));
  return {
    model,
    rowStart,
    rowEnd,
    colMin,
    colMax,
    rowKeys,
    rowCount: rowEnd - rowStart + 1,
    colCount: colMax - colMin + 1,
  };
}

// Tolerancyjne parsowanie liczby z komórki: liczba wprost, albo string z
// separatorami PL/EN (spacje/nbsp jako tysiące, przecinek lub kropka jako
// dziesiętny), z opcjonalnym znakiem procenta.
function parseCellNumber(raw, display) {
  if (typeof raw === "number" && Number.isFinite(raw)) return raw;
  let s = display != null && String(display).trim() ? String(display) : String(raw ?? "");
  s = s.trim();
  if (!s) return null;
  const percent = /%\s*$/.test(s);
  s = s.replace(/[^\d.,\-]/g, "");
  if (!s || s === "-" || s === "." || s === ",") return null;
  const lastComma = s.lastIndexOf(",");
  const lastDot = s.lastIndexOf(".");
  let normalized;
  if (lastComma > lastDot) {
    // przecinek to separator dziesiętny → kropki to tysiące
    normalized = s.replace(/\./g, "").replace(/,/g, ".");
  } else {
    // kropka dziesiętna (lub brak przecinka) → przecinki to tysiące
    normalized = s.replace(/,/g, "");
  }
  const n = Number(normalized);
  if (!Number.isFinite(n)) return null;
  return percent ? n / 100 : n;
}

function computeSelectionStats(rect) {
  if (!rect) return null;
  const { model, rowStart, rowEnd, colMin, colMax } = rect;
  let cellCount = 0;
  let numericCount = 0;
  let sum = 0;
  let min = Infinity;
  let max = -Infinity;
  for (let r = rowStart; r <= rowEnd; r++) {
    const row = model.rows[r];
    if (!row) continue;
    for (let c = colMin; c <= colMax; c++) {
      const display = getDisplayValue(row, c);
      if (display != null && String(display).trim()) cellCount += 1;
      const num = parseCellNumber(row.values?.[c], display);
      if (num != null && Number.isFinite(num)) {
        numericCount += 1;
        sum += num;
        if (num < min) min = num;
        if (num > max) max = num;
      }
    }
  }
  return {
    cellCount,
    numericCount,
    sum,
    avg: numericCount ? sum / numericCount : 0,
    min: numericCount ? min : 0,
    max: numericCount ? max : 0,
  };
}

function formatStatNumber(n) {
  const locale = (I18N[currentLang] && I18N[currentLang].locale) || "pl-PL";
  const rounded = Math.round(n * 1e6) / 1e6;
  return rounded.toLocaleString(locale, { maximumFractionDigits: 6 });
}

const RANGE_EDGE_CLASSES = [
  "cell-in-range",
  "range-edge-top",
  "range-edge-bottom",
  "range-edge-left",
  "range-edge-right",
];

// Czyści natywne (OS-owe) zaznaczenie tekstu — używane przy zaznaczaniu zakresu
// komórek, żeby gest Shift nie pozostawiał podświetlonej treści komórek.
function clearTextSelection() {
  const sel = typeof window.getSelection === "function" ? window.getSelection() : null;
  if (sel && typeof sel.removeAllRanges === "function") sel.removeAllRanges();
}

// Czy istnieje realny zakres zaznaczenia (więcej niż jedna komórka), a nie sama kotwica.
function hasActiveCellRange() {
  const rect = getSelectionRectangle();
  return !!rect && !(rect.rowCount === 1 && rect.colCount === 1);
}

// Podświetla prostokąt zaznaczenia: wypełnienie + obwódkę całego zakresu
// (klasy krawędziowe na komórkach brzegowych — jak w Arkuszach Google).
function syncRangeHighlightInDom() {
  tbodyEl
    .querySelectorAll("td.cell-in-range")
    .forEach((c) => c.classList.remove(...RANGE_EDGE_CLASSES));
  const rect = getSelectionRectangle();
  if (!rect || (rect.rowCount === 1 && rect.colCount === 1)) {
    // brak zakresu → pełne podświetlenie wiersza fokusu wraca do normy
    tbodyEl.classList.remove("has-cell-range");
    return;
  }
  // jest zakres → stłum pas wiersza fokusu, by nie kolidował z prostokątem zaznaczenia
  tbodyEl.classList.add("has-cell-range");
  for (let r = rect.rowStart; r <= rect.rowEnd; r++) {
    const row = rect.model.rows[r];
    if (!row) continue;
    const tr = tbodyEl.querySelector(`tr[data-row-key="${CSS.escape(getRowSelectionKey(row))}"]`);
    if (!tr) continue;
    for (let c = rect.colMin; c <= rect.colMax; c++) {
      const td = tr.querySelector(`td[data-col-index="${c}"]`);
      if (!td) continue; // np. komórka pokryta przez scalenie — obwódka może mieć tam lukę
      td.classList.add("cell-in-range");
      if (r === rect.rowStart) td.classList.add("range-edge-top");
      if (r === rect.rowEnd) td.classList.add("range-edge-bottom");
      if (c === rect.colMin) td.classList.add("range-edge-left");
      if (c === rect.colMax) td.classList.add("range-edge-right");
    }
  }
}

// Przycisk „Odznacz" (dotykowy zamiennik Esc) — widoczny, gdy jest aktywny
// fokus lub zaznaczenie. Czyści wszystko jednym tapnięciem.
function updateClearSelectionFab() {
  if (!clearSelectionFabEl) return;
  const active = !!(focusedCellState || selectedCellState);
  clearSelectionFabEl.classList.toggle("is-visible", active);
}

function updateCellStats() {
  updateClearSelectionFab();
  if (!cellStatsBarEl) return;
  const rect = getSelectionRectangle();
  // Jedna komórka = cisza (jak w Excelu — pasek pojawia się przy zakresie).
  if (!rect || (rect.rowCount === 1 && rect.colCount === 1)) {
    cellStatsBarEl.classList.add("hidden");
    cellStatsBarEl.replaceChildren();
    return;
  }
  const stats = computeSelectionStats(rect);
  if (!stats || stats.cellCount === 0) {
    cellStatsBarEl.classList.add("hidden");
    cellStatsBarEl.replaceChildren();
    return;
  }
  const parts = [
    [t("cellStatsRange"), `${rect.rowCount}×${rect.colCount}`],
    [t("cellStatsCount"), formatStatNumber(stats.cellCount)],
  ];
  if (stats.numericCount > 0) {
    parts.push([t("cellStatsSum"), formatStatNumber(stats.sum)]);
    parts.push([t("cellStatsAvg"), formatStatNumber(stats.avg)]);
    parts.push([t("cellStatsMin"), formatStatNumber(stats.min)]);
    parts.push([t("cellStatsMax"), formatStatNumber(stats.max)]);
  }
  const frag = document.createDocumentFragment();
  parts.forEach(([label, value]) => {
    const chip = document.createElement("span");
    chip.className = "cell-stat";
    const l = document.createElement("span");
    l.className = "cell-stat-label";
    l.textContent = label;
    const v = document.createElement("span");
    v.className = "cell-stat-value";
    v.textContent = value;
    chip.append(l, v);
    frag.appendChild(chip);
  });
  cellStatsBarEl.replaceChildren(frag);
  cellStatsBarEl.classList.remove("hidden");
}

function shouldIgnoreTableArrowNavigation() {
  const active = document.activeElement;
  if (!active) return false;
  const tag = String(active.tagName || "").toLowerCase();
  return active.isContentEditable || ["input", "textarea", "select", "button"].includes(tag);
}

function getAggregationSourceRows(scopeMode) {
  return scopeMode === "all" ? baseRows.slice() : viewRows.slice();
}

function getAggregationHeaderCandidateRows() {
  const candidates = new Set([currentHeaderRow]);
  currentSections.forEach((section) => {
    if (section?.action === "set-header" && Number.isFinite(section.headerRow)) {
      candidates.add(section.headerRow);
    }
  });
  for (let row = Math.max(1, currentHeaderRow - 3); row <= currentHeaderRow + 4; row++) {
    candidates.add(row);
  }
  return Array.from(candidates)
    .filter((row) => Number.isFinite(row) && row > 0)
    .sort((a, b) => a - b);
}

function getAggregationHeaderSourceData(headerRow = currentHeaderRow, scopeMode = aggregationWorkbenchState.scopeMode) {
  if (!workbook || !currentSheetName || headerRow === currentHeaderRow) {
    return {
      headerRow: currentHeaderRow,
      rows: getAggregationSourceRows(scopeMode),
      headers: currentHeaders.slice(),
      startCol: currentStartCol,
      group: getActiveRepeatingGroup(),
      longAvailable: canUseLongView(),
      helperMode: false,
    };
  }

  const sheet = workbook.Sheets[currentSheetName];
  if (!sheet) {
    return {
      headerRow: currentHeaderRow,
      rows: getAggregationSourceRows(scopeMode),
      headers: currentHeaders.slice(),
      startCol: currentStartCol,
      group: getActiveRepeatingGroup(),
      longAvailable: canUseLongView(),
      helperMode: false,
    };
  }

  try {
    const data = buildRows(sheet, headerRow, workbook);
    const rows = markSubheaderRows(data.rows.slice());
    const visibleRowIndexes = scopeMode === "filtered"
      ? new Set(viewRows.map((row) => row.rowIndex0))
      : null;
    const scopedRows = visibleRowIndexes
      ? rows.filter((row) => visibleRowIndexes.has(row.rowIndex0))
      : rows;
    const groups = detectRepeatingBlocks(sheet, headerRow, data);
    const group = Array.isArray(groups) && groups.length ? groups[0] : null;
    return {
      headerRow,
      rows: scopedRows,
      headers: data.headers.slice(),
      startCol: data.startCol || 0,
      group,
      longAvailable: !!(group && Array.isArray(group.blocks) && group.blocks.length >= 2),
      helperMode: headerRow !== currentHeaderRow,
    };
  } catch {
    return {
      headerRow: currentHeaderRow,
      rows: getAggregationSourceRows(scopeMode),
      headers: currentHeaders.slice(),
      startCol: currentStartCol,
      group: getActiveRepeatingGroup(),
      longAvailable: canUseLongView(),
      helperMode: false,
    };
  }
}

function collectAggregationContextForHeaderRow(headerRow, sourceMode = aggregationWorkbenchState.sourceMode, scopeMode = aggregationWorkbenchState.scopeMode) {
  const source = getAggregationHeaderSourceData(headerRow, scopeMode);
  const normalizedSource = sourceMode === "auto"
    ? (source.longAvailable ? "long" : "wide")
    : sourceMode;
  const model = normalizedSource === "long" && source.longAvailable
    ? buildLongViewModelFromRows(source.rows, source.group, {
      headers: source.headers,
      headerRow: source.headerRow,
      startCol: source.startCol,
    })
    : buildWideDisplayModelFromRows(source.rows, {
      headers: source.headers,
      headerRow: source.headerRow,
      startCol: source.startCol,
    });
  const profiles = collectAggregationProfiles(model);
  const groupOptions = resolveAggregationGroupOptions(profiles);
  const measures = detectAggregationMeasureCandidates(model, profiles);
  return {
    ...source,
    model,
    profiles,
    groupOptions,
    measures,
    resolvedSourceMode: normalizedSource,
  };
}

function scoreAggregationContext(context) {
  if (!context?.model?.rows?.length) return -1;
  const dateRangeBonus = context.measures.some((candidate) => candidate.measureType === "date_range") ? 30 : 0;
  const textGroupBonus = context.groupOptions.some((option) => /\b(imie|nazwisko|osoba|pracownik|owner|assignee)\b/.test(normalizeAnalysisKey(option.label))) ? 12 : 0;
  const currentHeaderBonus = context.headerRow === currentHeaderRow ? 4 : 0;
  return (context.groupOptions.length * 8)
    + (context.measures.length * 10)
    + dateRangeBonus
    + textGroupBonus
    + currentHeaderBonus
    + Math.min(context.model.rows.length, 500) * 0.02;
}

function isValidAggregationHeaderRow(headerRow) {
  if (!Number.isFinite(headerRow) || headerRow < 1) return false;
  if (!workbook || !currentSheetName) return false;
  const sheet = workbook.Sheets[currentSheetName];
  if (!sheet) return false;
  try {
    const data = buildRows(sheet, headerRow, workbook);
    return Array.isArray(data?.headers) && data.headers.length > 0 && Array.isArray(data?.rows) && data.rows.length > 0;
  } catch {
    return false;
  }
}

function getAggregationSourceModel(sourceMode = aggregationWorkbenchState.sourceMode, scopeMode = aggregationWorkbenchState.scopeMode) {
  return collectAggregationContextForHeaderRow(currentHeaderRow, sourceMode, scopeMode).model;
}

function getDisplayModel() {
  if (tableViewMode === "long" && canUseLongView()) {
    return buildLongViewModel();
  }
  return buildWideDisplayModel();
}

function renderActiveTable() {
  currentDisplayModel = getDisplayModel();
  sortRowsForHeaders(currentDisplayModel.rows, currentDisplayModel.headers);
  renderTable(currentDisplayModel);
  updateWideLongToggle();
}

function updateSortControls() {
  if (!resetSortBtn) return;
  normalizeSortState();
  const active = multiSortState.length > 0;
  resetSortBtn.classList.toggle("hidden", !active);
  resetSortBtn.setAttribute("aria-hidden", active ? "false" : "true");
  renderSortRules();
}

function toPixelWidth(meta) {
  if (!meta || typeof meta !== "object") return null;
  if (Number.isFinite(meta.wpx)) return Math.max(40, Math.round(meta.wpx));
  if (Number.isFinite(meta.wch)) return Math.max(40, Math.round(meta.wch * 8 + 16));
  if (Number.isFinite(meta.width)) return Math.max(40, Math.round(meta.width * 7 + 8));
  return null;
}

function toPixelHeight(meta) {
  if (!meta || typeof meta !== "object") return null;
  if (Number.isFinite(meta.hpx)) return Math.max(18, Math.round(meta.hpx));
  if (Number.isFinite(meta.hpt)) return Math.max(18, Math.round((meta.hpt * 96) / 72));
  return null;
}

function normalizeHexColor(input) {
  if (!input) return null;
  const raw = String(input).replace(/^#/, "").trim();
  if (/^[A-Fa-f0-9]{8}$/.test(raw)) return `#${raw.slice(2)}`;
  if (/^[A-Fa-f0-9]{6}$/.test(raw)) return `#${raw}`;
  if (/^[A-Fa-f0-9]{3}$/.test(raw)) return `#${raw[0]}${raw[0]}${raw[1]}${raw[1]}${raw[2]}${raw[2]}`;
  return null;
}

// Legacy 56-kolorowa paleta indeksowana Excela (color indexed="N").
// 64/65 = automatyczny foreground/background → null (traktujemy jak domyślny).
const INDEXED_COLORS = [
  "#000000","#FFFFFF","#FF0000","#00FF00","#0000FF","#FFFF00","#FF00FF","#00FFFF",
  "#000000","#FFFFFF","#FF0000","#00FF00","#0000FF","#FFFF00","#FF00FF","#00FFFF",
  "#800000","#008000","#000080","#808000","#800080","#008080","#C0C0C0","#808080",
  "#9999FF","#993366","#FFFFCC","#CCFFFF","#660066","#FF8080","#0066CC","#CCCCFF",
  "#000080","#FF00FF","#FFFF00","#00FFFF","#800080","#800000","#008080","#0000FF",
  "#00CCFF","#CCFFFF","#CCFFCC","#FFFF99","#99CCFF","#FF99CC","#CC99FF","#FFCC99",
  "#3366FF","#33CCCC","#99CC00","#FFCC00","#FF9900","#FF6600","#666699","#969696",
  "#003366","#339966","#003300","#333300","#993300","#993366","#333399","#333333",
  null, null,
];

// Domyślna paleta motywu Office (2013+). Indeksy jak w atrybucie color theme="N":
// 0/1 (tło1/tekst1) są zamienione względem clrScheme — zgodnie z tym, co pokazuje Excel.
const THEME_COLORS = [
  "#FFFFFF", "#000000", "#E7E6E6", "#44546A",
  "#4472C4", "#ED7D31", "#A5A5A5", "#FFC000",
  "#5B9BD5", "#70AD47", "#0563C1", "#954F72",
];

// Tint motywu wg OOXML: modyfikuje LUMINANCJĘ w HSL (nie per-kanał).
function applyColorTint(hex, tint) {
  const m = /^#?([0-9a-fA-F]{6})$/.exec(hex || "");
  if (!m || !tint) return hex;
  const num = parseInt(m[1], 16);
  let r = ((num >> 16) & 255) / 255;
  let g = ((num >> 8) & 255) / 255;
  let b = (num & 255) / 255;
  const max = Math.max(r, g, b), min = Math.min(r, g, b);
  let h = 0, s = 0, l = (max + min) / 2;
  const d = max - min;
  if (d !== 0) {
    s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
    if (max === r) h = (g - b) / d + (g < b ? 6 : 0);
    else if (max === g) h = (b - r) / d + 2;
    else h = (r - g) / d + 4;
    h /= 6;
  }
  l = tint < 0 ? l * (1 + tint) : l * (1 - tint) + tint;
  l = Math.min(1, Math.max(0, l));
  let R, G, B;
  if (s === 0) { R = G = B = l; }
  else {
    const hue2rgb = (p, q, t) => {
      if (t < 0) t += 1;
      if (t > 1) t -= 1;
      if (t < 1 / 6) return p + (q - p) * 6 * t;
      if (t < 1 / 2) return q;
      if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6;
      return p;
    };
    const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
    const p = 2 * l - q;
    R = hue2rgb(p, q, h + 1 / 3); G = hue2rgb(p, q, h); B = hue2rgb(p, q, h - 1 / 3);
  }
  const toHex = (v) => Math.round(v * 255).toString(16).padStart(2, "0");
  return `#${toHex(R)}${toHex(G)}${toHex(B)}`;
}

// Rozwiązuje kolor węzła stylu. Przy podanym `resolveTheme` obsługuje też
// color indexed="N" oraz color theme="N" (+ tint) — Excel tak zwykle zapisuje
// kolory czcionek. Bez flagi: tylko rgb/auto (jak dotąd — dla teł/obramowań).
// Czy kolor (hex) jest jasny — do decyzji, czy pod zaznaczeniem wiersza zamienić tekst
// na czytelny (var(--ink)). Próg na percepcyjnej luminancji.
function isLightColor(hex) {
  const m = /^#?([0-9a-fA-F]{6})$/.exec(hex || "");
  if (!m) return false;
  const n = parseInt(m[1], 16);
  const r = (n >> 16) & 255, g = (n >> 8) & 255, b = n & 255;
  return (0.2126 * r + 0.7152 * g + 0.0722 * b) > 165;
}

function colorFromStyleNode(node, resolveTheme = false) {
  if (!node || typeof node !== "object") return null;
  const direct = normalizeHexColor(node.rgb ?? node.RGB);
  if (direct) return direct;
  if (resolveTheme) {
    const tint = Number(node.tint ?? node.Tint) || 0;
    const idx = node.indexed ?? node.Indexed;
    if (Number.isInteger(idx) && INDEXED_COLORS[idx]) {
      return tint ? applyColorTint(INDEXED_COLORS[idx], tint) : INDEXED_COLORS[idx];
    }
    const themeIdx = node.theme ?? node.Theme;
    if (Number.isInteger(themeIdx) && THEME_COLORS[themeIdx]) {
      return applyColorTint(THEME_COLORS[themeIdx], tint);
    }
  }
  const auto = normalizeHexColor(node.auto);
  if (auto) return auto;
  return null;
}

function isDefaultLikeFill(fill, fillColor) {
  if (!fill || typeof fill !== "object") return true;
  const patternType = String(fill.patternType || fill.PatternType || "none").toLowerCase();
  if (!patternType || patternType === "none") return true;
  if (!fillColor) return true;
  const fg = fill.fgColor || fill.FgColor || null;
  const hasExplicitFgColor = !!(
    fg
    && typeof fg === "object"
    && (
      fg.rgb != null
      || fg.RGB != null
      || fg.indexed != null
      || fg.Indexed != null
      || fg.theme != null
      || fg.Theme != null
      || fg.tint != null
      || fg.Tint != null
    )
  );
  const normalized = String(fillColor).toUpperCase();
  // White-ish fill can be intentionally chosen by the user (especially solid fill),
  // so treat it as custom when the fg color is explicitly present in style.
  if (normalized === "#FFFFFF" || normalized === "#FFFFFE") {
    return !(patternType === "solid" && hasExplicitFgColor);
  }
  if (normalized === "#000000") return true;
  return false;
}

function isDefaultLikeFontColor(fontColor) {
  if (!fontColor) return true;
  const normalized = String(fontColor).toUpperCase();
  return normalized === "#000000" || normalized === "#FFFFFF";
}

function isCustomAlignment(alignment) {
  if (!alignment || typeof alignment !== "object") return false;
  const horizontal = String(alignment.horizontal || alignment.Horizontal || "").toLowerCase();
  const vertical = String(alignment.vertical || alignment.Vertical || "").toLowerCase();
  const wrapText = !!(alignment.wrapText || alignment.wrap || alignment.WrapText);
  const isDefaultHorizontal = !horizontal || horizontal === "general";
  const isDefaultVertical = !vertical || vertical === "bottom";
  return !isDefaultHorizontal || !isDefaultVertical || wrapText;
}

function hasCustomBorder(border) {
  if (!border || typeof border !== "object") return false;
  const edges = [
    border.top || border.Top,
    border.right || border.Right,
    border.bottom || border.Bottom,
    border.left || border.Left,
    border.diagonal || border.Diagonal,
  ];
  return edges.some((edge) => {
    const style = String(edge?.style || edge?.Style || "").toLowerCase();
    return !!style && style !== "none";
  });
}

function resolveXfStyle(styleIndex, wb) {
  if (!Number.isFinite(styleIndex) || !wb || !wb.Styles) return null;
  const styles = wb.Styles;
  const xfs = styles.CellXfs || styles.cellXfs;
  const xf = Array.isArray(xfs) ? xfs[styleIndex] : null;
  if (!xf) return null;
  const fontId = xf.fontId ?? xf.FontId;
  const fillId = xf.fillId ?? xf.FillId;
  const borderId = xf.borderId ?? xf.BorderId;
  const alignment = xf.alignment || xf.Alignment || null;
  const numFmtId = xf.numFmtId ?? xf.NumFmtId;
  const fonts = styles.Fonts || styles.fonts || [];
  const fills = styles.Fills || styles.fills || [];
  const borders = styles.Borders || styles.borders || [];
  return {
    font: Number.isFinite(fontId) ? fonts[fontId] : null,
    fill: Number.isFinite(fillId) ? fills[fillId] : null,
    border: Number.isFinite(borderId) ? borders[borderId] : null,
    alignment,
    numFmtId,
  };
}

// Domyślny rozmiar czcionki skoroszytu (fonts[0]) — baza do skalowania względnego.
// Excel domyślnie to Calibri 11pt; jeśli nie da się odczytać, zakładamy 11.
function getWorkbookDefaultFontSize(wb) {
  const fonts = wb?.Styles?.Fonts || wb?.Styles?.fonts;
  const f0 = Array.isArray(fonts) ? fonts[0] : null;
  const sz = f0 ? Number(f0.sz ?? f0.Sz ?? f0.size ?? f0.Size) : 0;
  return sz > 0 ? sz : 11;
}

// Mapuje nazwę czcionki z pliku Excel na stos CSS: najpierw oryginalna nazwa
// (użyta, jeśli zainstalowana — np. Calibri na Windows), potem web-safe zamienniki
// i rodzina ogólna. Dzięki temu na każdym systemie wygląd jest najbliższy Excelowi
// bez ryzyka „braku" czcionki.
function excelFontToCssStack(name) {
  const n = String(name || "").trim();
  if (!n) return "";
  const FALLBACKS = {
    "calibri": "'Segoe UI', Candara, Optima, sans-serif",
    "calibri light": "'Segoe UI Light', 'Segoe UI', sans-serif",
    "aptos": "'Segoe UI', sans-serif",
    "aptos narrow": "'Segoe UI', sans-serif",
    "arial": "Helvetica, 'Helvetica Neue', sans-serif",
    "arial narrow": "'Arial Narrow', Arial, sans-serif",
    "arial black": "Arial, sans-serif",
    "helvetica": "'Helvetica Neue', Arial, sans-serif",
    "tahoma": "Geneva, Verdana, sans-serif",
    "verdana": "Geneva, sans-serif",
    "segoe ui": "Candara, sans-serif",
    "trebuchet ms": "'Lucida Grande', sans-serif",
    "century gothic": "'Apple Gothic', sans-serif",
    "times new roman": "Times, serif",
    "times": "serif",
    "cambria": "Georgia, serif",
    "georgia": "'Times New Roman', serif",
    "garamond": "'Apple Garamond', Georgia, serif",
    "book antiqua": "Palatino, 'Palatino Linotype', serif",
    "palatino linotype": "Palatino, serif",
    "courier new": "Courier, monospace",
    "consolas": "'Lucida Console', Monaco, monospace",
    "lucida console": "Monaco, monospace",
    "comic sans ms": "'Comic Sans', cursive",
  };
  const fallback = FALLBACKS[n.toLowerCase()] || "sans-serif";
  return `"${n}", ${fallback}`;
}

function decodeXmlEntities(s) {
  return String(s)
    .replace(/&lt;/g, "<").replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"').replace(/&apos;/g, "'")
    .replace(/&#(\d+);/g, (_, d) => String.fromCharCode(+d))
    .replace(/&amp;/g, "&");
}

// Parsuje surowy XML arkusza → Map<cellRef, xfIndex> z atrybutów <c r="A1" s="N">.
function parseCellStyleIndices(sheetXml) {
  const map = new Map();
  const re = /<c\s+r="([A-Z]+\d+)"([^>]*)>/g;
  let m;
  while ((m = re.exec(sheetXml))) {
    const sMatch = /\bs="(\d+)"/.exec(m[2]);
    if (sMatch) map.set(m[1], parseInt(sMatch[1], 10));
  }
  return map;
}

// Odzyskuje mapę stylów per arkusz wprost z pliku .xlsx (JSZip), bo ten build
// xlsx-js-style nie wystawia indeksu stylu na komórce. Zwraca
// Map<sheetName, Map<cellRef, xfIndex>> albo null (brak JSZip / błąd / nie-xlsx).
async function buildStyleIndexMap(bytes, wb) {
  if (typeof JSZip === "undefined" || !bytes || !wb) return null;
  try {
    const zip = await JSZip.loadAsync(bytes);
    const wbXml = await zip.file("xl/workbook.xml")?.async("string");
    const relsXml = await zip.file("xl/_rels/workbook.xml.rels")?.async("string");
    if (!wbXml || !relsXml) return null;

    const ridToPath = new Map();
    const relRe = /<Relationship\b[^>]*?Id="([^"]+)"[^>]*?Target="([^"]+)"[^>]*?\/?>/g;
    let rm;
    while ((rm = relRe.exec(relsXml))) {
      const target = rm[2];
      if (!/worksheets\//.test(target)) continue;
      let path = decodeXmlEntities(target);
      if (path.startsWith("/")) path = path.slice(1);
      else if (!path.startsWith("xl/")) path = "xl/" + path.replace(/^\.\//, "");
      ridToPath.set(rm[1], path);
    }

    const result = new Map();
    const sheetRe = /<sheet\b[^>]*?name="([^"]+)"[^>]*?r:id="([^"]+)"[^>]*?\/?>/g;
    let sm;
    while ((sm = sheetRe.exec(wbXml))) {
      const name = decodeXmlEntities(sm[1]);
      const path = ridToPath.get(sm[2]);
      const file = path ? zip.file(path) : null;
      if (!file) continue;
      const xml = await file.async("string");
      result.set(name, parseCellStyleIndices(xml));
    }
    return result.size ? result : null;
  } catch {
    return null;
  }
}

// Rozwiązuje styl z indeksu xf (s="N" z surowego XML komórki) → font/fill/border.
// KLUCZOWE: ten build xlsx-js-style przy odczycie wrzuca do `cell.s` tylko fill,
// gubiąc font/border/alignment — ale pełne tabele są w `wb.Styles.CellXf/Fonts/...`
// (Fonts mają już policzony color.rgb), więc po odzyskaniu indeksu wszystko wraca.
function resolveXfStyleFromIndex(sIdx, wb) {
  if (!Number.isInteger(sIdx) || !wb || !wb.Styles) return null;
  const st = wb.Styles;
  const xfs = st.CellXf || st.CellXfs || st.cellXfs;
  const xf = Array.isArray(xfs) ? xfs[sIdx] : null;
  if (!xf) return null;
  const pick = (a, b) => {
    const v = a ?? b;
    const n = Number(v);
    return Number.isFinite(n) ? n : null;
  };
  const fontId = pick(xf.fontId, xf.fontid);
  const fillId = pick(xf.fillId, xf.fillid);
  const borderId = pick(xf.borderId, xf.borderid);
  const fonts = st.Fonts || st.fonts || [];
  const fills = st.Fills || st.fills || [];
  const borders = st.Borders || st.borders || [];
  return {
    fontId,
    font: fontId != null ? fonts[fontId] : null,
    fill: fillId != null ? fills[fillId] : null,
    border: borderId != null ? borders[borderId] : null,
    alignment: xf.alignment || xf.Alignment || null,
  };
}

function extractCellStyle(cell, wb, xfStyle = null) {
  const cellS = cell && cell.s && typeof cell.s === "object" ? cell.s : null;
  if (!cellS && !xfStyle) return null;

  // Źródła: font/border/alignment bierzemy z xf (cell.s ich nie ma w tym buildzie),
  // a fill z xf (autorytatywny, ma rgb) lub z samego cell.s (bywa węzłem fill).
  const cellFillNode = cellS && (cellS.patternType != null || cellS.PatternType != null || cellS.fgColor || cellS.FgColor)
    ? cellS
    : (cellS?.fill || cellS?.Fill || null);
  const fill = (xfStyle && xfStyle.fill) || cellFillNode || null;
  const border = (xfStyle && xfStyle.border) || cellS?.border || cellS?.Border || null;
  const alignment = (xfStyle && xfStyle.alignment) || cellS?.alignment || cellS?.Alignment || null;
  const rawFont = (xfStyle && xfStyle.font) || cellS?.font || cellS?.Font || null;
  // Font domyślny skoroszytu (fontId 0) zostawiamy natywnemu wyglądowi aplikacji —
  // inaczej KAŻDA komórka dostałaby kolor/rodzinę domyślną (psułoby też dark mode).
  const isDefaultFont = xfStyle ? xfStyle.fontId === 0 : false;
  const font = isDefaultFont ? null : rawFont;

  const fillColor = colorFromStyleNode(fill?.fgColor || fill?.FgColor || fill?.bgColor || fill?.BgColor, true);
  const fontColor = colorFromStyleNode(font?.color || font?.Color, true);
  const hasCustomFill = !isDefaultLikeFill(fill, fillColor);
  // Czerń/biel zwykle pomijamy (żeby dark mode mógł sterować kolorem), ALE gdy komórka
  // ma własne wypełnienie, jawny kolor tekstu (np. biały na ciemnym tle) jest istotny
  // dla kontrastu — wtedy go honorujemy (jak w Excelu).
  const hasCustomFontColor = !!fontColor && (!isDefaultLikeFontColor(fontColor) || hasCustomFill);
  const hasCustomAlign = isCustomAlignment(alignment);
  const hasBorder = hasCustomBorder(border);

  // Czcionka jak w Excelu: rodzina (nazwa) + względny rozmiar wobec domyślnej.
  const fontName = (font && (font.name || font.Name || font.rFont)) || "";
  const rawFontSize = font ? Number(font.sz ?? font.Sz ?? font.size ?? font.Size) : 0;
  const baseFontSize = getWorkbookDefaultFontSize(wb);
  const fontFamily = fontName ? excelFontToCssStack(fontName) : "";
  const fontScale = rawFontSize > 0 && baseFontSize > 0 ? rawFontSize / baseFontSize : 0;

  const styleOut = {
    fillColor,
    hasCustomFill,
    fontColor,
    hasCustomFontColor,
    fontFamily,
    fontScale,
    bold: !!(font && (font.bold || font.b || font.Bold)),
    italic: !!(font && (font.italic || font.i || font.Italic)),
    underline: !!(font && (font.underline || font.u || font.Underline)),
    horizontal: hasCustomAlign ? (alignment?.horizontal || alignment?.Horizontal || "") : "",
    vertical: hasCustomAlign ? (alignment?.vertical || alignment?.Vertical || "") : "",
    wrapText: hasCustomAlign ? !!(alignment && (alignment.wrapText || alignment.wrap || alignment.WrapText)) : false,
    hasBorder,
    border,
  };

  return styleOut;
}

function applyEdgeBorder(td, edge) {
  if (!edge) return;
  const borderStyle = edge.style || edge.Style || "";
  if (!borderStyle || borderStyle === "none") return;
  const color = colorFromStyleNode(edge.color || edge.Color) || "rgba(0,0,0,0.32)";
  return `1px solid ${color}`;
}

function applyCellStyle(td, style) {
  if (!style) return;
  if (cellStyleShowFills && style.hasCustomFill && style.fillColor) {
    td.classList.add("cell-has-fill");
    // Gdy komórka ma jawny kolor tekstu (i pokazujemy kolory), renderuj tło pełnym
    // kryciem — to wierna para fg+bg z Excela (np. biały tekst na pełnym kolorze).
    // Pozostałe wypełnienia zostają subtelne (0.28), żeby nie przytłaczać widoku.
    const strongFill = cellStyleShowFontColors && style.hasCustomFontColor;
    const bg = hexToRgba(style.fillColor, strongFill ? 1 : 0.28);
    if (bg) td.style.background = bg;
  }
  if (cellStyleShowFontColors && style.hasCustomFontColor && style.fontColor) {
    td.style.color = style.fontColor;
    // Jasny/biały tekst oznaczamy klasą — pod zaznaczeniem wiersza (które zakrywa tło)
    // CSS zamieni go na czytelny (var(--ink)), żeby nie zniknął na jasnym tle zaznaczenia.
    if (isLightColor(style.fontColor)) td.classList.add("cell-light-text");
  }
  if (cellStyleShowFonts && style.fontFamily) td.style.fontFamily = style.fontFamily;
  if (cellStyleShowFonts && style.fontScale && Math.abs(style.fontScale - 1) > 0.01) {
    td.style.setProperty("--cell-font-scale", String(Math.round(style.fontScale * 1000) / 1000));
  }
  // Pogrubienie/kursywa/podkreślenie też pod „Pokaż formatowanie tekstu" (cellStyleShowFonts),
  // żeby dało się je wyłączyć (użytkownik nie zawsze chce dziedziczyć pogrubienia z pliku).
  if (cellStyleShowFonts && style.bold) td.style.fontWeight = "700";
  if (cellStyleShowFonts && style.italic) td.style.fontStyle = "italic";
  if (cellStyleShowFonts && style.underline) td.style.textDecoration = "underline";
  if (style.horizontal) td.style.textAlign = style.horizontal;
  if (style.vertical) td.style.verticalAlign = style.vertical;
  // Zawijanie tekstu NIE jest brane z pliku per komórka (dawało nieoczekiwane
  // łamanie wierszy). Steruje nim globalny przełącznik „Zawijaj tekst" w Widoku
  // (klasa #dataTable.wrap-cells), domyślnie wyłączony.

  if (cellStyleShowBorders && style.hasBorder && style.border && typeof style.border === "object") {
    const t = applyEdgeBorder(td, style.border.top || style.border.Top);
    const r = applyEdgeBorder(td, style.border.right || style.border.Right);
    const b = applyEdgeBorder(td, style.border.bottom || style.border.Bottom);
    const l = applyEdgeBorder(td, style.border.left || style.border.Left);
    if (t) td.style.borderTop = t;
    if (r) td.style.borderRight = r;
    if (b) td.style.borderBottom = b;
    if (l) td.style.borderLeft = l;
  }
}

function computeMergeLayout(rowsShown, colCount) {
  if (!Array.isArray(currentMerges) || !currentMerges.length || !rowsShown.length) return null;
  const rowPosByAbs = new Map();
  rowsShown.forEach((row, pos) => rowPosByAbs.set(row.rowIndex0, pos));
  const anchors = new Map();
  const covered = new Set();

  currentMerges.forEach((merge) => {
    if (!merge || !merge.s || !merge.e) return;
    if (merge.s.c < currentStartCol || merge.e.c >= currentStartCol + colCount) return;
    const topPos = rowPosByAbs.get(merge.s.r);
    if (topPos == null) return;
    for (let r = merge.s.r; r <= merge.e.r; r++) {
      const p = rowPosByAbs.get(r);
      if (p == null || p !== topPos + (r - merge.s.r)) return;
    }
    const startCol = merge.s.c - currentStartCol;
    const endCol = merge.e.c - currentStartCol;
    const rowspan = merge.e.r - merge.s.r + 1;
    const colspan = endCol - startCol + 1;
    if (rowspan < 2 && colspan < 2) return;

    const anchorKey = `${topPos}:${startCol}`;
    anchors.set(anchorKey, {
      rowspan,
      colspan,
      ref: XLSX.utils.encode_range({
        s: { r: merge.s.r, c: merge.s.c },
        e: { r: merge.e.r, c: merge.e.c },
      }),
    });
    for (let rp = topPos; rp < topPos + rowspan; rp++) {
      for (let c = startCol; c <= endCol; c++) {
        if (rp === topPos && c === startCol) continue;
        covered.add(`${rp}:${c}`);
      }
    }
  });

  return { anchors, covered };
}

function computeHeaderMergeLayout(colCount) {
  if (!Array.isArray(currentMerges) || !currentMerges.length) return null;
  const headerAbsRow = currentHeaderRow - 1;
  const anchors = new Map();
  const covered = new Set();

  currentMerges.forEach((merge) => {
    if (!merge || !merge.s || !merge.e) return;
    if (merge.s.r !== headerAbsRow || merge.e.r !== headerAbsRow) return;
    if (merge.s.c < currentStartCol || merge.e.c >= currentStartCol + colCount) return;
    const startCol = merge.s.c - currentStartCol;
    const endCol = merge.e.c - currentStartCol;
    const colspan = endCol - startCol + 1;
    if (colspan < 2) return;
    anchors.set(startCol, {
      colspan,
      ref: XLSX.utils.encode_range({
        s: { r: merge.s.r, c: merge.s.c },
        e: { r: merge.e.r, c: merge.e.c },
      }),
    });
    for (let c = startCol + 1; c <= endCol; c++) covered.add(c);
  });

  return { anchors, covered };
}

function computeColumnWidths(headers, rows, useExcelLayout) {
  const widths = headers.map(() => 0);
  const min = 80;
  const max = 520;

  // Globalna szerokość kolumn (pole „Szerokość kolumn (px)") ma pierwszeństwo nad
  // auto-dopasowaniem i Wymiarami z Excela; ręczne przeciągnięcie pojedynczej kolumny dalej wygrywa.
  if (manualColWidthAll > 0) {
    return widths.map((_, i) => Math.max(40, Math.min(900, manualColumnWidths[i] || manualColWidthAll)));
  }

  if (useExcelLayout && Array.isArray(currentSheetColWidths) && currentSheetColWidths.length) {
    return widths.map((_, i) => {
      const manual = manualColumnWidths[i];
      if (manual) return Math.max(min, Math.min(max, manual));
      const fromSheet = toPixelWidth(currentSheetColWidths[i]);
      if (fromSheet) return Math.max(min, Math.min(max, fromSheet));
      return 140;
    });
  }

  const canvas = document.createElement("canvas");
  const ctx = canvas.getContext("2d");
  // Mierzymy osobno wariantem normalnym i POGRUBIONYM — pogrubiony tekst jest szerszy,
  // a wcześniej mierzyliśmy wszystko jako normalny → kolumny (np. dat) bywały za wąskie
  // i ucinały ~połowę wartości. Czcionka brana z realnego stylu tabeli.
  const cs = getComputedStyle(tableEl);
  // WAŻNE: mierzymy rozmiarem realnie używanym przez <td> (CSS: 12.5px × zoom),
  // a NIE font-size samego <table> (bywa inny, np. 12px) — inaczej kolumny wychodzą
  // za wąskie i ucinają tekst, gdy formatowanie tekstu z pliku jest wyłączone.
  const zoom = (typeof zoomLevelEl !== "undefined" && parseFloat(zoomLevelEl.value)) || 1;
  const fSize = `${12.5 * zoom}px`;
  const fFamily = cs.fontFamily || "sans-serif";
  const normalFont = `${fSize} ${fFamily}`;
  const boldFont = `700 ${fSize} ${fFamily}`;
  const fontsShown = cellStyleShowFonts; // pogrubienie pokazywane tylko gdy włączone
  const measure = (text, bold, scale) => {
    ctx.font = bold ? boldFont : normalFont;
    let w = ctx.measureText(text).width;
    if (scale > 1) w *= scale;
    return w;
  };

  headers.forEach((h, i) => {
    const hs = currentHeaderStyles && currentHeaderStyles[i];
    widths[i] = Math.max(widths[i], measure(h, fontsShown && hs && hs.bold, (hs && hs.fontScale) || 1));
  });
  const limit = Math.min(rows.length, 300);
  const samples = headers.map(() => []);
  for (let r = 0; r < limit; r++) {
    rows[r].values.forEach((v, i) => {
      // Zabezpieczenie: wiersz może mieć więcej wartości niż nagłówków
      // (np. uszkodzony/pusty model long) — pomijamy kolumny poza zakresem nagłówków.
      if (i >= samples.length) return;
      const text = getDisplayValue(rows[r], i);
      // Puste komórki POMIJAMY w próbce — inaczej w kolumnach z dużą liczbą pustych
      // (np. od2/do2 w późniejszych cyklach) 90. percentyl spada do ~0, kolumna zwija
      // się do minimum i nieliczne realne daty zostają ucięte.
      if (!text || !text.trim()) return;
      const cellStyle = rows[r].cellStyles && rows[r].cellStyles[i];
      samples[i].push(measure(text, fontsShown && cellStyle && cellStyle.bold, (cellStyle && cellStyle.fontScale) || 1));
    });
  }
  const padding = 26;
  // Inteligentnie: dopasuj do większości (p90) i przytnij skrajnie długie teksty, by
  // kolumny nie robiły się ogromne. Wyłączone: zmieść CAŁĄ zawartość (najdłuższa komórka),
  // z wyższym limitem szerokości.
  const pct = cellStyleSmartWidths ? 0.9 : 1.0;
  const maxW = cellStyleSmartWidths ? max : 900;
  return widths.map((base, i) => {
    const manual = manualColumnWidths[i];
    if (manual) return Math.max(min, Math.min(900, manual));
    const colSamples = samples[i].sort((a, b) => a - b);
    const idx = Math.floor(colSamples.length * pct);
    const pVal = colSamples.length ? colSamples[Math.min(idx, colSamples.length - 1)] : base;
    const raw = Math.max(base, pVal) * 1.02 + padding; // +2% zapasu na różnice renderowania
    return Math.max(min, Math.min(maxW, Math.ceil(raw)));
  });
}

function renderTable(modelOrHeaders, maybeRows) {
  const model = Array.isArray(modelOrHeaders)
    ? {
        mode: "wide",
        headers: modelOrHeaders,
        rows: Array.isArray(maybeRows) ? maybeRows : [],
        guideLabels: modelOrHeaders.map((_, i) => XLSX.utils.encode_col(i + currentStartCol)),
        headerRowLabel: String(currentHeaderRow),
        rowHeadFormatter: (row) => String((row?.rowIndex0 ?? 0) + 1),
        editable: true,
      }
    : (modelOrHeaders || { headers: [], rows: [] });
  const headers = Array.isArray(model.headers) ? model.headers : [];
  const rows = Array.isArray(model.rows) ? model.rows : [];

  updateSortControls();
  if (!headers.length) {
    setStatus(t("tableNoData"));
    if (tableScrollbarEl) tableScrollbarEl.classList.add("hidden");
    setEmptyState(DEFAULT_EMPTY_TITLE, DEFAULT_EMPTY_SUB);
    return;
  }
  if (!rows.length) {
    setStatus(t("statusTableRowsEmpty"));
    if (tableScrollbarEl) tableScrollbarEl.classList.add("hidden");
    setEmptyState(t("tableNoResults"), t("tableNoResultsHint"));
    return;
  }

  showTable();
  theadEl.replaceChildren();
  tbodyEl.replaceChildren();

  const useExcelLayout = isExcelLayoutEnabled();
  const widths = computeColumnWidths(headers, rows, useExcelLayout);
  const rowHeaderDigits = String(rows.length + currentHeaderRow).length;
  const rowHeaderWidth = Math.max(42, rowHeaderDigits * 8 + 18);
  // Offset dla sticky 1. kolumny danych (blokada kolumny) = szerokość kolumny numerów.
  if (tableWrapEl) tableWrapEl.style.setProperty("--row-head-w", `${rowHeaderWidth}px`);

  const colgroup = document.createElement("colgroup");
  const rowHeadCol = document.createElement("col");
  rowHeadCol.style.width = `${rowHeaderWidth}px`;
  colgroup.appendChild(rowHeadCol);
  widths.forEach((w) => {
    const col = document.createElement("col");
    col.style.width = `${w}px`;
    colgroup.appendChild(col);
  });
  tableEl.replaceChildren();
  tableEl.appendChild(colgroup);
  tableEl.appendChild(theadEl);
  tableEl.appendChild(tbodyEl);

  const guideRow = document.createElement("tr");
  guideRow.className = "guide-row";
  const corner = document.createElement("th");
  corner.className = "corner-cell";
  corner.textContent = "";
  guideRow.appendChild(corner);
  headers.forEach((_, i) => {
    const th = document.createElement("th");
    th.className = "guide-cell";
    th.setAttribute("scope", "col");
    th.textContent = Array.isArray(model.guideLabels) && model.guideLabels[i] ? model.guideLabels[i] : XLSX.utils.encode_col(i + currentStartCol);
    const resizer = document.createElement("div");
    resizer.className = "col-resizer";
    resizer.dataset.colIndex = String(i);
    th.appendChild(resizer);
    guideRow.appendChild(th);
  });
  theadEl.appendChild(guideRow);

  const headRow = document.createElement("tr");
  headRow.className = "header-row";
  const rowHead = document.createElement("th");
  rowHead.className = "row-head";
  rowHead.setAttribute("scope", "row");
  rowHead.textContent = model.headerRowLabel || String(currentHeaderRow);
  headRow.appendChild(rowHead);
  const headerMergeLayout = model.mode === "wide" ? computeHeaderMergeLayout(headers.length) : null;
  for (let i = 0; i < headers.length; i++) {
    if (headerMergeLayout && headerMergeLayout.covered.has(i)) continue;
    const h = headers[i];
    const th = document.createElement("th");
    th.setAttribute("scope", "col");
    th.textContent = h;
    if (currentHeaderStyles[i]) applyCellStyle(th, currentHeaderStyles[i]);

    if (headerMergeLayout) {
      const merge = headerMergeLayout.anchors.get(i);
      if (merge) {
        th.colSpan = merge.colspan;
        th.classList.add("cell-merged");
        if (merge.ref) th.title = `Scalona komórka: ${merge.ref}`;
      }
    }

    th.addEventListener("click", () => {
      if (sortState.col === h) {
        setPrimarySort(h, sortState.dir === "asc" ? "desc" : "asc");
      } else {
        setPrimarySort(h, "asc");
      }
      if (model.mode === "wide") {
        sortRows();
      }
      updateSortControls();
      renderActiveTable();
    });

    const primarySort = multiSortState[0];
    if (primarySort && primarySort.col === h) {
      const arrow = document.createElement("span");
      arrow.className = "sort-arrow";
      arrow.textContent = primarySort.dir === "asc" ? "▲" : "▼";
      th.appendChild(arrow);
    }

    headRow.appendChild(th);
  }
  theadEl.appendChild(headRow);

  const limit = Math.max(1, parseInt(maxRowsEl.value || "200", 10));
  const rowsShown = rows.slice(0, limit);
  const mergeLayout = model.mode === "wide" ? computeMergeLayout(rowsShown, headers.length) : null;

  // Formatowanie warunkowe: mapa kolor/tło per ref (tylko widok „wide" — w „long"
  // komórki są przeukładane, więc odwołania CF nie mają sensu).
  const cfMapForRender = (cellStyleShowConditionalFormatting && model.mode === "wide")
    ? getSheetCFMap(currentSheetName)
    : null;

  const tbodyFragment = document.createDocumentFragment();
  rowsShown.forEach((row, rowPos) => {
    const tr = document.createElement("tr");
    tr.dataset.rowKey = getRowSelectionKey(row);
    if (focusedCellState && focusedCellState.rowKey === tr.dataset.rowKey) tr.classList.add("row-focused");
    if (cellStyleShowSubheaders && row.isSubheader) tr.classList.add("row-subheader");
    if (quickSearchHighlightMode && matchedRowIndexes.size > 0) {
      if (matchedRowIndexes.has(row.rowIndex0)) {
        tr.classList.add("row-matched");
      } else {
        tr.classList.add("row-unmatched");
      }
    }
    if (typeof row.rowIndex0 === "number") {
      tr.dataset.rowIndex = String(row.rowIndex0);
    }
    // Wysokość wiersza: ręczne przeciąganie > jednolita z pola > z pliku (Wymiary z Excela).
    let rowH = manualRowHeights[row.rowIndex0] || (manualRowHeightAll > 0 ? manualRowHeightAll : 0);
    if (!rowH && useExcelLayout) rowH = toPixelHeight(currentSheetRowHeights[row.rowIndex0]) || 0;
    if (rowH) {
      tr.style.height = `${rowH}px`;
      tr.classList.add("row-fixed-height");
    }
    const rowHead = document.createElement("td");
    rowHead.className = "row-head";
    rowHead.textContent = model.rowHeadFormatter ? model.rowHeadFormatter(row, rowPos) : String(row.rowIndex0 + 1);
    if (typeof row.rowIndex0 === "number") {
      const rowResizer = document.createElement("div");
      rowResizer.className = "row-resizer";
      rowResizer.dataset.rowIndex = String(row.rowIndex0);
      rowHead.appendChild(rowResizer);
    }
    tr.appendChild(rowHead);
    const matchedCols = highlightMatchedCells ? matchedCellsByRow.get(row.rowIndex0) : null;
    row.values.forEach((v, i) => {
      const mergeKey = `${rowPos}:${i}`;
      if (mergeLayout && mergeLayout.covered.has(mergeKey)) return;
      const td = document.createElement("td");
      const displayValue = getDisplayValue(row, i);
      td.textContent = displayValue;
      td.dataset.fullText = displayValue;
      td.dataset.colIndex = String(i);
      if (selectedCellState && selectedCellState.rowKey === tr.dataset.rowKey && selectedCellState.colIndex0 === i) {
        td.classList.add("cell-selected");
      }
      if (matchedCols && matchedCols.has(i)) td.classList.add("cell-filter-match");

      if (mergeLayout) {
        const anchor = mergeLayout.anchors.get(mergeKey);
        if (anchor) {
          if (anchor.rowspan > 1) td.rowSpan = anchor.rowspan;
          if (anchor.colspan > 1) td.colSpan = anchor.colspan;
          td.classList.add("cell-merged");
          if (anchor.ref) td.title = `Scalona komórka: ${anchor.ref}`;
        }
      }

      if (row.cellStyles && row.cellStyles[i]) applyCellStyle(td, row.cellStyles[i]);
      if (cfMapForRender) {
        const cf = cfMapForRender.get(XLSX.utils.encode_cell({ r: row.rowIndex0, c: currentStartCol + i }));
        if (cf) {
          if (cf.fontColor) { td.style.color = cf.fontColor; if (isLightColor(cf.fontColor)) td.classList.add("cell-light-text"); }
          if (cf.fillColor) { td.classList.add("cell-has-fill"); const cbg = hexToRgba(cf.fillColor, cf.fontColor ? 1 : 0.28); if (cbg) td.style.background = cbg; }
        }
      }
      tr.appendChild(td);
    });
    tbodyFragment.appendChild(tr);
  });
  tbodyEl.appendChild(tbodyFragment);

  updateTableStatus(model);
  syncFocusedCellInDom({ clearMissing: true });
  syncSelectedCellInDom({ clearMissing: true });
  syncRangeHighlightInDom();
  updateCellStats();
  if (typeof updateScrollTopFab === "function") updateScrollTopFab();
  syncHorizontalScrollbar();
  applyZoom();
  applyFreezeHeaders();
}

function buildRows(sheet, headerRow, wb) {
  const originalRange = XLSX.utils.decode_range(sheet["!ref"]);
  const range = computeEffectiveSheetRange(sheet, headerRow);
  const colMeta = sheet["!cols"] || [];
  const rowMeta = sheet["!rows"] || [];
  const merges = Array.isArray(sheet["!merges"]) ? sheet["!merges"] : [];
  const styleIndexForSheet = currentStyleIndexMap ? currentStyleIndexMap.get(currentSheetName) : null;
  const xfStyleAt = (ref) => (styleIndexForSheet ? resolveXfStyleFromIndex(styleIndexForSheet.get(ref), wb) : null);
  const rawHeaders = [];
  const headerStyles = [];
  for (let c = range.s.c; c <= range.e.c; c++) {
    const ref = XLSX.utils.encode_cell({ r: headerRow - 1, c });
    const cell = sheet[ref];
    const v = cell ? cell.v : null;
    rawHeaders.push(v ? String(v).trim() : XLSX.utils.encode_col(c));
    headerStyles.push(wb ? extractCellStyle(cell, wb, xfStyleAt(ref)) : null);
  }
  const headers = makeHeadersUnique(rawHeaders);
  const duplicateHeaderCount = rawHeaders.length - new Set(rawHeaders).size;
  const rows = [];
  let formulaCount = 0;
  let formulaMissingResultCount = 0;
  let commentCount = 0;
  let hyperlinkCount = 0;
  for (let r = headerRow; r <= range.e.r; r++) {
    const values = [];
    const display = [];
    const cellStyles = [];
    let any = false;
    for (let c = range.s.c; c <= range.e.c; c++) {
      const ref = XLSX.utils.encode_cell({ r, c });
      const cell = sheet[ref];
      let v = cell ? cell.v : null;
      let shown = cell && cell.w ? String(cell.w) : toDisplay(v);
      shown = localizeDisplayedDate(v, shown, cell);
      if (cell && cell.f) {
        formulaCount += 1;
        if (cell.v == null && cell.w == null) formulaMissingResultCount += 1;
      }
      if (cell && Array.isArray(cell.c) && cell.c.length) commentCount += 1;
      if (cell && cell.l && (cell.l.Target || cell.l.target)) hyperlinkCount += 1;
      if (displayModeEl.value === "formulas" && cell && cell.f) {
        v = "=" + cell.f;
        shown = v;
      }
      values.push(v);
      display.push(shown);
      cellStyles.push(wb ? extractCellStyle(cell, wb, xfStyleAt(ref)) : null);
      if (v !== null && v !== "") any = true;
    }
    if (!any) continue;
    rows.push({ values, display, rawValues: values, rowIndex0: r, cellStyles });
  }
  return {
    headers,
    headerStyles,
    rows,
    startCol: range.s.c,
    merges,
    stats: {
      duplicateHeaderCount,
      formulaCount,
      formulaMissingResultCount,
      commentCount,
      hyperlinkCount,
      mergeRegions: merges.length,
      mergedCells: merges.reduce((sum, merge) => sum + ((merge.e.r - merge.s.r + 1) * (merge.e.c - merge.s.c + 1)), 0),
      hiddenColumns: colMeta.filter((meta) => meta && meta.hidden).length,
      hiddenRows: rowMeta.filter((meta) => meta && meta.hidden).length,
      sourceRange: XLSX.utils.encode_range(originalRange),
      effectiveRange: XLSX.utils.encode_range(range),
    },
  };
}
