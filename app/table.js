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
  wideLongToggleEl.textContent = tableViewMode === "long" ? "Widok klasyczny" : "Wide-to-Long";
  wideLongToggleEl.title = tableViewMode === "long"
    ? "Wroc do klasycznego ukladu arkusza"
    : "Przelacz wykryte bloki kolumn na dlugi widok analityczny";
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
  const headers = [...prefixHeaders, "Nr bloku", "Blok", ...repeatedHeaders];
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
    return;
  }
  focusedCellState = { rowKey, colIndex0 };
  syncFocusedCellInDom(options);
}

function setSelectedCell(rowKey, colIndex0, options = {}) {
  if (!rowKey || !Number.isFinite(colIndex0) || colIndex0 < 0) {
    selectedCellState = null;
    syncSelectedCellInDom({ clearMissing: false });
    return;
  }
  selectedCellState = { rowKey, colIndex0 };
  syncSelectedCellInDom(options);
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

function colorFromStyleNode(node) {
  if (!node || typeof node !== "object") return null;
  const rgb = node.rgb ?? node.RGB;
  const direct = normalizeHexColor(rgb);
  if (direct) return direct;
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

function extractCellStyle(cell, wb) {
  if (!cell) return null;
  let style = null;
  if (cell.s && typeof cell.s === "object") {
    style = cell.s;
  } else if (Number.isFinite(cell.s)) {
    style = resolveXfStyle(cell.s, wb);
  }
  if (!style || typeof style !== "object") return null;

  const fill = style.fill || style.Fill || null;
  const font = style.font || style.Font || null;
  const border = style.border || style.Border || null;
  const alignment = style.alignment || style.Alignment || null;

  const fillColor = colorFromStyleNode(fill?.fgColor || fill?.FgColor || fill?.bgColor || fill?.BgColor);
  const fontColor = colorFromStyleNode(font?.color || font?.Color);
  const hasCustomFill = !isDefaultLikeFill(fill, fillColor);
  const hasCustomFontColor = !isDefaultLikeFontColor(fontColor);
  const hasCustomAlign = isCustomAlignment(alignment);
  const hasBorder = hasCustomBorder(border);

  const styleOut = {
    fillColor,
    hasCustomFill,
    fontColor,
    hasCustomFontColor,
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
  if (style.hasCustomFill && style.fillColor) {
    td.classList.add("cell-has-fill");
    td.style.background = hexToRgba(style.fillColor, 0.28) || td.style.background;
  }
  if (style.hasCustomFontColor && style.fontColor) td.style.color = style.fontColor;
  if (style.bold) td.style.fontWeight = "700";
  if (style.italic) td.style.fontStyle = "italic";
  if (style.underline) td.style.textDecoration = "underline";
  if (style.horizontal) td.style.textAlign = style.horizontal;
  if (style.vertical) td.style.verticalAlign = style.vertical;
  if (style.wrapText) td.style.whiteSpace = "normal";

  if (style.hasBorder && style.border && typeof style.border === "object") {
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
  const tableFont = getComputedStyle(tableEl).font;
  ctx.font = tableFont;

  headers.forEach((h, i) => {
    widths[i] = Math.max(widths[i], ctx.measureText(h).width);
  });
  const limit = Math.min(rows.length, 300);
  const samples = headers.map(() => []);
  for (let r = 0; r < limit; r++) {
    rows[r].values.forEach((v, i) => {
      const text = getDisplayValue(rows[r], i);
      const w = ctx.measureText(text).width;
      samples[i].push(w);
    });
  }
  const padding = 24;
  return widths.map((base, i) => {
    const colSamples = samples[i].sort((a, b) => a - b);
    const idx = Math.floor(colSamples.length * 0.9);
    const p90 = colSamples.length ? colSamples[Math.min(idx, colSamples.length - 1)] : base;
    const raw = Math.max(base, p90) + padding;
    const manual = manualColumnWidths[i];
    if (manual) return Math.max(min, Math.min(max, manual));
    return Math.max(min, Math.min(max, Math.ceil(raw)));
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
    setStatus("Brak danych");
    if (tableScrollbarEl) tableScrollbarEl.classList.add("hidden");
    setEmptyState(DEFAULT_EMPTY_TITLE, DEFAULT_EMPTY_SUB);
    return;
  }
  if (!rows.length) {
    setStatus("Wierszy: 0");
    if (tableScrollbarEl) tableScrollbarEl.classList.add("hidden");
    setEmptyState("Brak wynikow", "Zmien filtry albo wybierz inny arkusz.");
    return;
  }

  showTable();
  theadEl.replaceChildren();
  tbodyEl.replaceChildren();

  const useExcelLayout = isExcelLayoutEnabled();
  const widths = computeColumnWidths(headers, rows, useExcelLayout);
  const rowHeaderDigits = String(rows.length + currentHeaderRow).length;
  const rowHeaderWidth = Math.max(42, rowHeaderDigits * 8 + 18);

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

  rowsShown.forEach((row, rowPos) => {
    const tr = document.createElement("tr");
    tr.dataset.rowKey = getRowSelectionKey(row);
    if (focusedCellState && focusedCellState.rowKey === tr.dataset.rowKey) tr.classList.add("row-focused");
    if (row.isSubheader) tr.classList.add("row-subheader");
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
    if (useExcelLayout) {
      const rowMeta = currentSheetRowHeights[row.rowIndex0];
      const h = toPixelHeight(rowMeta);
      if (h) tr.style.height = `${h}px`;
    }
    const rowHead = document.createElement("td");
    rowHead.className = "row-head";
    rowHead.textContent = model.rowHeadFormatter ? model.rowHeadFormatter(row, rowPos) : String(row.rowIndex0 + 1);
    tr.appendChild(rowHead);
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
      tr.appendChild(td);
    });
    tbodyEl.appendChild(tr);
  });

  const modeLabel = model.mode === "long" ? " • tryb long" : "";
  setStatus(`Wierszy: ${rows.length} (pokazano: ${Math.min(rows.length, limit)})${modeLabel}`);
  syncFocusedCellInDom({ clearMissing: true });
  syncSelectedCellInDom({ clearMissing: true });
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
  const rawHeaders = [];
  const headerStyles = [];
  for (let c = range.s.c; c <= range.e.c; c++) {
    const cell = sheet[XLSX.utils.encode_cell({ r: headerRow - 1, c })];
    const v = cell ? cell.v : null;
    rawHeaders.push(v ? String(v).trim() : XLSX.utils.encode_col(c));
    headerStyles.push(wb ? extractCellStyle(cell, wb) : null);
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
      const cell = sheet[XLSX.utils.encode_cell({ r, c })];
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
      cellStyles.push(wb ? extractCellStyle(cell, wb) : null);
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
