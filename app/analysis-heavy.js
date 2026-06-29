// Ciężkie analizy (duration + aggregation) — leniwie przy pierwszym użyciu panelu.
// Zakłada globals z analysis.js (deferAnalysis, normalizeAnalysisKey, createEmptyInsight…).

function collectDurationBlockStats(group) {
  const firstBlock = group && Array.isArray(group.blocks) ? group.blocks[0] : null;
  if (!firstBlock || !Array.isArray(firstBlock.headers) || !firstBlock.headers.length) return [];

  const stats = firstBlock.headers.map((header, idx) => ({
    idx,
    header,
    nonEmptyCount: 0,
    dateCount: 0,
    durationCount: 0,
    textCount: 0,
    uniqueText: new Set(),
  }));

  const rowSample = viewRows.slice(0, 400);
  rowSample.forEach((row) => {
    group.blocks.forEach((block) => {
      stats.forEach((entry) => {
        const absIdx = block.startIndex + entry.idx;
        const raw = row.values[absIdx] ?? getDisplayValue(row, absIdx);
        const text = String(raw ?? "").trim();
        if (!text) return;
        entry.nonEmptyCount += 1;

        const asDate = parseDateFlexible(raw);
        if (asDate instanceof Date) {
          entry.dateCount += 1;
          return;
        }

        const asDuration = parseDurationDaysFlexible(raw);
        if (asDuration !== null) {
          entry.durationCount += 1;
          return;
        }

        entry.textCount += 1;
        entry.uniqueText.add(normalizeAnalysisKey(text));
      });
    });
  });

  return stats.map((entry) => ({
    ...entry,
    uniqueTextCount: entry.uniqueText.size,
  }));
}

function inferDurationAnalysisConfigFromData(group) {
  const stats = collectDurationBlockStats(group);
  if (!stats.length) return null;

  const entityCandidate = stats
    .filter((entry) => entry.textCount > 0)
    .sort((a, b) => {
      const scoreA = (a.textCount * 5) + Math.min(a.uniqueTextCount, 25);
      const scoreB = (b.textCount * 5) + Math.min(b.uniqueTextCount, 25);
      return scoreB - scoreA || a.idx - b.idx;
    })[0];

  const dateCandidates = stats
    .filter((entry) => entry.dateCount > 0)
    .sort((a, b) => b.dateCount - a.dateCount || a.idx - b.idx);

  const durationCandidate = stats
    .filter((entry) => entry.durationCount > 0)
    .sort((a, b) => b.durationCount - a.durationCount || a.idx - b.idx)[0];

  if (!entityCandidate) return null;
  if (!dateCandidates.length && !durationCandidate) return null;

  const orderedDateCandidates = dateCandidates.slice().sort((a, b) => a.idx - b.idx);
  const startCandidate = orderedDateCandidates[0] || null;
  const endCandidate = orderedDateCandidates[1] || null;

  return {
    entityIdx: entityCandidate.idx,
    startIdx: startCandidate ? startCandidate.idx : -1,
    endIdx: endCandidate ? endCandidate.idx : -1,
    durationIdx: durationCandidate ? durationCandidate.idx : -1,
    entityLabel: "Osoba",
    entityHeader: entityCandidate.header || "Osoba",
    inferred: true,
  };
}

function detectDurationAnalysisConfig(group) {
  const firstBlock = group && Array.isArray(group.blocks) ? group.blocks[0] : null;
  if (!firstBlock || !Array.isArray(firstBlock.headers) || !firstBlock.headers.length) return null;

  const candidates = firstBlock.headers.map((header, idx) => {
    const base = parseRepeatedHeader(header)?.base || cleanSectionLabel(header) || String(header || "");
    return {
      idx,
      header,
      base,
      norm: normalizeAnalysisKey(base),
    };
  });

  const entityIdx = findAnalysisColumnIndex(candidates, [
    (norm) => /\b(imie|nazwisko|osoba|pracownik|opiekun|wlasciciel|owner|assignee|user|agent|operator)\b/.test(norm),
    (norm) => norm.includes("imie") || norm.includes("nazwisk"),
  ]);
  const startIdx = findAnalysisColumnIndex(candidates, [
    (norm) => norm === "od" || norm === "data od",
    (norm) => /\b(start|from|poczatek|rozpoczecie|rozpoczecia)\b/.test(norm),
  ]);
  const endIdx = findAnalysisColumnIndex(candidates, [
    (norm) => norm === "do" || norm === "data do",
    (norm) => /\b(koniec|zakonczenie|end|to|until)\b/.test(norm),
  ]);
  const durationIdx = findAnalysisColumnIndex(candidates, [
    (norm) => norm.includes("dlugosc") || norm.includes("czas"),
    (norm) => /\b(duration|age|days)\b/.test(norm),
  ]);

  const inferred = inferDurationAnalysisConfigFromData(group);

  const resolvedEntityIdx = entityIdx >= 0 ? entityIdx : (inferred?.entityIdx ?? -1);
  const resolvedStartIdx = startIdx >= 0 ? startIdx : (inferred?.startIdx ?? -1);
  const resolvedEndIdx = endIdx >= 0 ? endIdx : (inferred?.endIdx ?? -1);
  const resolvedDurationIdx = durationIdx >= 0 ? durationIdx : (inferred?.durationIdx ?? -1);

  if (resolvedEntityIdx < 0 || (resolvedStartIdx < 0 && resolvedDurationIdx < 0)) return null;

  const entityBase = candidates[resolvedEntityIdx]?.base || inferred?.entityHeader || "Wartosc";
  const normEntity = normalizeAnalysisKey(entityBase);
  let entityLabel = "Wartosc";
  if (normEntity.includes("imie") || normEntity.includes("nazwisk") || normEntity.includes("osoba")) entityLabel = "Osoba";
  else if (normEntity.includes("pracownik")) entityLabel = "Pracownik";
  else if (normEntity.includes("owner") || normEntity.includes("wlasciciel")) entityLabel = "Wlasciciel";
  else if (inferred?.inferred) entityLabel = "Osoba";
  else if (entityBase) entityLabel = entityBase;

  return {
    entityIdx: resolvedEntityIdx,
    startIdx: resolvedStartIdx,
    endIdx: resolvedEndIdx,
    durationIdx: resolvedDurationIdx,
    entityLabel,
    entityHeader: entityBase,
    inferred: !!(inferred && (entityIdx < 0 || startIdx < 0 || endIdx < 0 || durationIdx < 0)),
  };
}

function buildDurationAnalysisFromRows(group, rows, meta = {}) {
  const config = detectDurationAnalysisConfig(group);
  if (!config) {
    return { status: "no-config", group, ...meta };
  }

  const today = new Date();
  const records = [];
  const aggregate = new Map();

  rows.forEach((row) => {
    group.blocks.forEach((block, blockIndex) => {
      const entityCol = block.startIndex + config.entityIdx;
      const startCol = config.startIdx >= 0 ? block.startIndex + config.startIdx : -1;
      const endCol = config.endIdx >= 0 ? block.startIndex + config.endIdx : -1;
      const durationCol = config.durationIdx >= 0 ? block.startIndex + config.durationIdx : -1;

      const entityValue = String(row.values[entityCol] ?? "").trim();
      if (!entityValue) return;

      const startDate = startCol >= 0 ? parseDateFlexible(row.values[startCol] ?? getDisplayValue(row, startCol)) : null;
      const endDate = endCol >= 0 ? parseDateFlexible(row.values[endCol] ?? getDisplayValue(row, endCol)) : null;
      let durationDays = null;
      let isOpen = false;

      if (startDate instanceof Date) {
        if (endDate instanceof Date) {
          durationDays = diffDays(startDate, endDate);
        } else {
          durationDays = diffDays(startDate, today);
          isOpen = durationDays !== null;
        }
      }

      if (durationDays === null && durationCol >= 0) {
        durationDays = parseDurationDaysFlexible(row.values[durationCol] ?? getDisplayValue(row, durationCol));
      }

      records.push({
        entity: entityValue,
        durationDays,
        isOpen,
        isClosed: durationDays !== null && !isOpen,
        blockLabel: block.label,
        blockIndex: blockIndex + 1,
        rowIndex0: row.rowIndex0,
      });
    });
  });

  const filteredRecords = records.filter((record) => {
    if (!Number.isFinite(record.durationDays)) return false;
    if (durationAnalysisState.statusFilter === "open") return record.isOpen;
    if (durationAnalysisState.statusFilter === "closed") return !record.isOpen;
    return true;
  });

  filteredRecords.forEach((record) => {
    const key = normalizeAnalysisKey(record.entity);
    const entry = aggregate.get(key) || {
      entity: record.entity,
      durations: [],
      openCount: 0,
      minDays: null,
      maxDays: null,
      blocks: new Set(),
      rowIndexes: new Set(),
    };
    entry.durations.push(record.durationDays);
    if (record.isOpen) entry.openCount += 1;
    entry.minDays = entry.minDays === null ? record.durationDays : Math.min(entry.minDays, record.durationDays);
    entry.maxDays = entry.maxDays === null ? record.durationDays : Math.max(entry.maxDays, record.durationDays);
    entry.blocks.add(record.blockLabel);
    entry.rowIndexes.add(record.rowIndex0);
    aggregate.set(key, entry);
  });

  const entries = Array.from(aggregate.values())
    .map((entry) => ({
      entity: entry.entity,
      averageDays: entry.durations.length ? entry.durations.reduce((sum, value) => sum + value, 0) / entry.durations.length : null,
      medianDays: computeMedian(entry.durations),
      count: entry.durations.length,
      openCount: entry.openCount,
      minDays: entry.minDays,
      maxDays: entry.maxDays,
      blockCount: entry.blocks.size,
      rowCount: entry.rowIndexes.size,
    }))
    .sort((a, b) => {
      const metricMap = {
        avg: "averageDays",
        median: "medianDays",
        count: "count",
        min: "minDays",
        max: "maxDays",
      };
      const metric = metricMap[durationAnalysisState.sortMetric] || "averageDays";
      const left = Number(a[metric] || 0);
      const right = Number(b[metric] || 0);
      const diff = right - left;
      if (Math.abs(diff) > 0.001) return diff;
      const countDiff = b.count - a.count;
      if (countDiff) return countDiff;
      return a.entity.localeCompare(b.entity, "pl");
    });

  if (!entries.length) {
    return { status: "no-records", config, group, records, filteredRecords, ...meta };
  }

  const totalDurationRecords = filteredRecords.length;
  const totalOpen = filteredRecords.filter((record) => record.isOpen).length;
  const totalClosed = filteredRecords.filter((record) => !record.isOpen).length;
  const allDurations = filteredRecords.map((record) => record.durationDays).filter((value) => Number.isFinite(value));
  const totalDays = allDurations.reduce((sum, value) => sum + value, 0);

  return {
    status: "ok",
    config,
    group,
    entries,
    records,
    filteredRecords,
    ...meta,
    summary: {
      uniqueEntities: entries.length,
      totalDurationRecords,
      totalOpen,
      totalClosed,
      averageDays: totalDurationRecords ? totalDays / totalDurationRecords : null,
      medianDays: computeMedian(allDurations),
      minDays: allDurations.length ? Math.min(...allDurations) : null,
      maxDays: allDurations.length ? Math.max(...allDurations) : null,
      visibleRows: rows.length,
      sourceRows: rows.length,
    },
  };
}

function tryBuildDurationAnalysisFromAlternateHeaders() {
  if (!workbook || !currentSheetName) return null;
  const sheet = workbook.Sheets[currentSheetName];
  if (!sheet) return null;

  const candidateRows = [];
  const seen = new Set();
  const minHeader = 1;
  const maxHeader = Math.max(minHeader, currentHeaderRow + 4);

  for (let row = Math.max(minHeader, currentHeaderRow - 3); row <= maxHeader; row++) {
    if (row === currentHeaderRow) continue;
    if (seen.has(row)) continue;
    seen.add(row);
    candidateRows.push(row);
  }

  let best = null;

  candidateRows.forEach((headerRow) => {
    try {
      const data = buildRows(sheet, headerRow, workbook);
      const groups = detectRepeatingBlocks(sheet, headerRow, data);
      const group = Array.isArray(groups) && groups.length ? groups[0] : null;
      if (!group || !Array.isArray(group.blocks) || group.blocks.length < 2) return;

      const shadowRows = markSubheaderRows(data.rows.slice());
      const result = buildDurationAnalysisFromRows(group, shadowRows, {
        helperHeaderRow: headerRow,
        helperMode: true,
      });
      if (!result || result.status !== "ok") return;

      const score = (result.summary.totalDurationRecords * 10) + result.summary.uniqueEntities;
      if (!best || score > best.score) {
        best = { ...result, score };
      }
    } catch {
      // Ignore helper header candidates that fail to parse well.
    }
  });

  return best;
}

function buildDurationAnalysis() {
  const group = getActiveRepeatingGroup();
  if (!group || !Array.isArray(group.blocks) || group.blocks.length < 2) {
    const fallback = tryBuildDurationAnalysisFromAlternateHeaders();
    return fallback || { status: "no-group" };
  }

  const currentResult = buildDurationAnalysisFromRows(group, viewRows, {
    helperHeaderRow: currentHeaderRow,
    helperMode: false,
  });

  if (currentResult.status === "ok") {
    return currentResult;
  }

  const fallback = tryBuildDurationAnalysisFromAlternateHeaders();
  return fallback || currentResult;
}

function renderDurationAnalysis() {
  if (deferAnalysis("duration")) return;
  if (!durationAnalysisSummaryEl || !durationAnalysisListEl) return;
  durationAnalysisSummaryEl.replaceChildren();
  durationAnalysisListEl.replaceChildren();

  const analysis = buildDurationAnalysis();

  if (analysis.status === "no-group") {
    durationAnalysisSummaryEl.appendChild(createEmptyInsight(t("durationNoGroup")));
    return;
  }

  if (analysis.status === "no-config") {
    durationAnalysisSummaryEl.appendChild(createEmptyInsight(t("durationNoConfig")));
    return;
  }

  if (analysis.status === "no-records") {
    durationAnalysisSummaryEl.appendChild(createEmptyInsight(t("durationNoRecords")));
    return;
  }

  const summaryGrid = document.createElement("div");
  summaryGrid.className = "sheet-inspector-summary";
  [
    { label: pluralizeEntityLabel(analysis.config.entityLabel), value: String(analysis.summary.uniqueEntities) },
    { label: t("durationRecordsCount"), value: String(analysis.summary.totalDurationRecords) },
    { label: t("durationAvgTime"), value: formatDurationDays(analysis.summary.averageDays) },
    { label: t("durationMedian"), value: formatDurationDays(analysis.summary.medianDays) },
    { label: t("durationMin"), value: formatDurationDays(analysis.summary.minDays) },
    { label: t("durationMax"), value: formatDurationDays(analysis.summary.maxDays) },
    { label: t("durationInProgress"), value: String(analysis.summary.totalOpen), tone: analysis.summary.totalOpen ? "info" : "" },
    { label: t("durationClosed"), value: String(analysis.summary.totalClosed) },
  ].forEach((item) => {
    const chip = document.createElement("div");
    chip.className = `sheet-inspector-chip${item.tone ? ` ${item.tone}` : ""}`;

    const label = document.createElement("div");
    label.className = "sheet-inspector-chip-label";
    label.textContent = item.label;

    const value = document.createElement("div");
    value.className = "sheet-inspector-chip-value";
    value.textContent = item.value;

    chip.appendChild(label);
    chip.appendChild(value);
    summaryGrid.appendChild(chip);
  });
  durationAnalysisSummaryEl.appendChild(summaryGrid);

  const note = document.createElement("div");
  note.className = "duration-analysis-note";
  const filtered = analysis.summary.visibleRows !== analysis.summary.sourceRows;
  note.textContent = filtered
    ? t("durationNoteFiltered", { visible: analysis.summary.visibleRows, source: analysis.summary.sourceRows })
    : t("durationNoteFull");
  if (analysis.config.inferred) {
    note.textContent += t("durationNoteInferred");
  }
  if (analysis.helperMode && Number.isFinite(analysis.helperHeaderRow) && analysis.helperHeaderRow !== currentHeaderRow) {
    note.textContent += t("durationNoteHelper", { row: analysis.helperHeaderRow, current: currentHeaderRow });
  }
  durationAnalysisSummaryEl.appendChild(note);

  const controls = document.createElement("div");
  controls.className = "duration-analysis-controls";

  const statusField = document.createElement("label");
  statusField.className = "field";
  statusField.append(t("durationStatusLabel"));
  const statusSelect = document.createElement("select");
  statusSelect.dataset.durationControl = "status";
  [
    { value: "all", label: t("durationStatusAll") },
    { value: "closed", label: t("durationStatusClosed") },
    { value: "open", label: t("durationStatusOpen") },
  ].forEach((item) => {
    const option = document.createElement("option");
    option.value = item.value;
    option.textContent = item.label;
    statusSelect.appendChild(option);
  });
  statusSelect.value = durationAnalysisState.statusFilter;
  statusField.appendChild(statusSelect);

  const sortField = document.createElement("label");
  sortField.className = "field";
  sortField.append(t("durationSortLabel"));
  const sortSelect = document.createElement("select");
  sortSelect.dataset.durationControl = "sort";
  [
    { value: "avg", label: t("durationSortAvg") },
    { value: "median", label: t("durationSortMedian") },
    { value: "count", label: t("durationSortCount") },
    { value: "max", label: t("durationSortMax") },
    { value: "min", label: t("durationSortMin") },
  ].forEach((item) => {
    const option = document.createElement("option");
    option.value = item.value;
    option.textContent = item.label;
    sortSelect.appendChild(option);
  });
  sortSelect.value = durationAnalysisState.sortMetric;
  sortField.appendChild(sortSelect);

  const countField = document.createElement("label");
  countField.className = "field";
  countField.append(t("durationShowLabel"));
  const countSelect = document.createElement("select");
  countSelect.dataset.durationControl = "count";
  ["14", "24", "40", "80", "999"].forEach((value) => {
    const option = document.createElement("option");
    option.value = value;
    option.textContent = value === "999" ? t("durationShowAll") : value;
    countSelect.appendChild(option);
  });
  countSelect.value = String(durationAnalysisState.showCount);
  countField.appendChild(countSelect);

  controls.appendChild(statusField);
  controls.appendChild(sortField);
  controls.appendChild(countField);
  durationAnalysisSummaryEl.appendChild(controls);

  const actions = document.createElement("div");
  actions.className = "section-nav-actions";
  if (canUseLongView()) {
    const toggleBtn = document.createElement("button");
    toggleBtn.className = "btn ghost btn-sm";
    toggleBtn.type = "button";
    toggleBtn.dataset.durationAction = "toggle-long";
    toggleBtn.textContent = tableViewMode === "long" ? t("durationViewClassic") : "Wide-to-Long";
    actions.appendChild(toggleBtn);
  }
  if (filtered) {
    const resetBtn = document.createElement("button");
    resetBtn.className = "btn ghost btn-sm";
    resetBtn.type = "button";
    resetBtn.dataset.durationAction = "reset-filters";
    resetBtn.textContent = t("durationShowFull");
    actions.appendChild(resetBtn);
  }
  durationAnalysisSummaryEl.appendChild(actions);

  const listNote = document.createElement("div");
  listNote.className = "duration-analysis-note";
  const visibleCount = Math.min(durationAnalysisState.showCount, analysis.entries.length);
  listNote.textContent = analysis.entries.length > visibleCount
    ? t("durationShownPartial", { shown: visibleCount, total: analysis.entries.length })
    : t("durationShownAll", { total: analysis.entries.length });
  durationAnalysisListEl.appendChild(listNote);

  analysis.entries.slice(0, durationAnalysisState.showCount).forEach((entry, index) => {
    const item = document.createElement("div");
    item.className = "duration-person-item";

    const top = document.createElement("div");
    top.className = "duration-person-top";

    const titleWrap = document.createElement("div");
    titleWrap.className = "duration-person-title-wrap";

    const rank = document.createElement("div");
    rank.className = "duration-person-rank";
    rank.textContent = String(index + 1);

    const title = document.createElement("div");
    title.className = "duration-person-title";
    title.textContent = entry.entity;

    const value = document.createElement("div");
    value.className = "duration-person-value";
    value.textContent = formatDurationDays(entry.averageDays);

    titleWrap.appendChild(rank);
    titleWrap.appendChild(title);
    top.appendChild(titleWrap);
    top.appendChild(value);

    const meta = document.createElement("div");
    meta.className = "duration-person-meta";
    const avgDaysText = entry.averageDays !== null ? t("durationDaysUnit", { n: Math.round(entry.averageDays * 10) / 10 }) : t("durationNone");
    const medianDaysText = entry.medianDays !== null ? t("durationDaysUnit", { n: Math.round(entry.medianDays * 10) / 10 }) : t("durationNone");
    meta.textContent = t("durationMetaLine", {
      avg: avgDaysText,
      median: medianDaysText,
      count: entry.count,
      open: entry.openCount,
      min: formatDurationDays(entry.minDays),
      max: formatDurationDays(entry.maxDays),
    });

    const actionsRow = document.createElement("div");
    actionsRow.className = "section-nav-actions";

    const filterBtn = document.createElement("button");
    filterBtn.className = "btn ghost btn-sm";
    filterBtn.type = "button";
    filterBtn.dataset.durationAction = "filter-entity";
    filterBtn.dataset.durationEntity = entry.entity;
    filterBtn.textContent = t("durationShowInTable");
    actionsRow.appendChild(filterBtn);

    item.appendChild(top);
    item.appendChild(meta);
    item.appendChild(actionsRow);
    durationAnalysisListEl.appendChild(item);
  });
}

function renderAggregationWorkbench() {
  renderMonthlySummary(); // panel „Podsumowanie miesięczne" odświeża się tym samym cyklem (ma własną bramkę)
  if (deferAnalysis("aggregation")) return;
  if (!aggregationWorkbenchSummaryEl || !aggregationWorkbenchListEl) return;
  aggregationWorkbenchSummaryEl.replaceChildren();
  aggregationWorkbenchListEl.replaceChildren();

  const result = buildAggregationWorkbenchResult();
  if (typeof currentAggregationMeasureCandidates !== "undefined") {
    currentAggregationMeasureCandidates = result.measures || [];
    currentAggregationGroupOptions = result.groupOptions || [];
  }
  if (result.status === "empty") {
    aggregationWorkbenchSummaryEl.appendChild(createEmptyInsight(t("aggregationNoData")));
    return;
  }
  if (result.status === "no-options") {
    aggregationWorkbenchSummaryEl.appendChild(createEmptyInsight(t("aggregationNoOptions")));
    return;
  }
  if (result.status === "no-results") {
    aggregationWorkbenchSummaryEl.appendChild(createEmptyInsight(t("aggregationNoResults")));
    return;
  }

  const measure = result.measure;
  const isDistinctMode = aggregationWorkbenchState.aggregation === "distinct";
  const primaryKind = isDistinctMode ? "number" : getPrimaryAggregationValueKind(result.selectedMeasures || [measure].filter(Boolean), aggregationWorkbenchState.aggregation);
  const aggregationLabels = {
    count: t("aggregationCount"),
    avg: t("aggregationAvg"),
    median: t("aggregationMedian"),
    min: t("aggregationMin"),
    max: t("aggregationMax"),
    sum: t("aggregationSum"),
    distinct: t("aggregationDistinct"),
    earliest: t("aggregationEarliest"),
    latest: t("aggregationLatest"),
  };
  const groupDepth = [
    aggregationWorkbenchState.groupBy,
    aggregationWorkbenchState.groupBy2,
    aggregationWorkbenchState.groupBy3,
  ].filter(Boolean).length;

  const summaryGrid = document.createElement("div");
  summaryGrid.className = "sheet-inspector-summary";
  [
    { label: t("aggregationGroups"), value: String(result.summary.groups) },
    { label: t("aggregationSourceRows"), value: String(result.summary.sourceRows) },
    { label: t("aggregationMeasuredRows"), value: String(result.summary.measuredRows) },
    { label: aggregationLabels[aggregationWorkbenchState.aggregation] || t("aggregationResult"), value: formatAggregationMetricValue(result.entries[0]?.primary, primaryKind), tone: "info" },
  ].forEach((item) => {
    const chip = document.createElement("div");
    chip.className = `sheet-inspector-chip${item.tone ? ` ${item.tone}` : ""}`;
    const label = document.createElement("div");
    label.className = "sheet-inspector-chip-label";
    label.textContent = item.label;
    const value = document.createElement("div");
    value.className = "sheet-inspector-chip-value";
    value.textContent = item.value;
    chip.appendChild(label);
    chip.appendChild(value);
    summaryGrid.appendChild(chip);
  });
  aggregationWorkbenchSummaryEl.appendChild(summaryGrid);

  const controls = document.createElement("div");
  controls.className = "aggregation-controls";

  const sourceField = document.createElement("label");
  sourceField.className = "field";
  sourceField.append(t("aggregationSource"));
  const sourceSelect = document.createElement("select");
  sourceSelect.dataset.aggregationControl = "source";
  [
    { value: "auto", label: t("aggregationAuto") },
    { value: "wide", label: t("aggregationClassic") },
    ...(result.longAvailable ? [{ value: "long", label: "Wide-to-Long" }] : []),
  ].forEach((item) => {
    const option = document.createElement("option");
    option.value = item.value;
    option.textContent = item.label;
    sourceSelect.appendChild(option);
  });
  sourceSelect.value = aggregationWorkbenchState.sourceMode;
  sourceField.appendChild(sourceSelect);

  const scopeField = document.createElement("label");
  scopeField.className = "field";
  scopeField.append(t("aggregationScope"));
  const scopeSelect = document.createElement("select");
  scopeSelect.dataset.aggregationControl = "scope";
  [
    { value: "filtered", label: t("aggregationCurrentView") },
    { value: "all", label: t("aggregationWholeSheet") },
  ].forEach((item) => {
    const option = document.createElement("option");
    option.value = item.value;
    option.textContent = item.label;
    scopeSelect.appendChild(option);
  });
  scopeSelect.value = aggregationWorkbenchState.scopeMode;
  scopeField.appendChild(scopeSelect);

  const headerField = document.createElement("label");
  headerField.className = "field";
  headerField.append(t("aggregationHeader"));
  const headerRowWrap = document.createElement("div");
  headerRowWrap.className = "aggregation-header-row";
  const headerSelect = document.createElement("select");
  headerSelect.dataset.aggregationControl = "header";
  const headerNumberInput = document.createElement("input");
  headerNumberInput.dataset.aggregationControl = "header-number";
  headerNumberInput.type = "number";
  headerNumberInput.min = "1";
  headerNumberInput.step = "1";
  headerNumberInput.inputMode = "numeric";
  headerNumberInput.placeholder = "nr wiersza";
  headerRowWrap.appendChild(headerSelect);
  headerRowWrap.appendChild(headerNumberInput);
  headerField.appendChild(headerRowWrap);
  const autoOpt = document.createElement("option");
  autoOpt.value = "auto";
  autoOpt.textContent = `${t("aggregationAuto")} (${result.resolvedHeaderRow})`;
  headerSelect.appendChild(autoOpt);
  const manualOpt = document.createElement("option");
  manualOpt.value = "manual";
  manualOpt.textContent = t("aggregationCustomNumber");
  headerSelect.appendChild(manualOpt);
  headerSelect.value = aggregationWorkbenchState.headerRowChoice;
  headerNumberInput.value = String(aggregationWorkbenchState.customHeaderRow || currentHeaderRow);
  headerNumberInput.disabled = aggregationWorkbenchState.headerRowChoice !== "manual";

  const groupField = document.createElement("label");
  groupField.className = "field";
  groupField.append(t("aggregationGroupBy"));
  const currentGroupLabels = [
    aggregationWorkbenchState.groupBy,
    aggregationWorkbenchState.groupBy2,
    aggregationWorkbenchState.groupBy3,
  ].filter(Boolean).map((val) => {
    const opt = result.groupOptions.find((o) => o.value === val);
    return opt ? opt.label : val;
  });
  const groupInput = document.createElement("input");
  groupInput.type = "text";
  groupInput.readOnly = true;
  groupInput.className = "aggregation-measure-input";
  groupInput.value = currentGroupLabels.length ? currentGroupLabels.join(", ") : t("aggregationNone");
  groupInput.title = currentGroupLabels.join(", ");
  const groupPickBtn = document.createElement("button");
  groupPickBtn.type = "button";
  groupPickBtn.className = "btn ghost btn-sm";
  groupPickBtn.dataset.aggregationControl = "groupby-pick";
  groupPickBtn.textContent = t("choose");
  const groupRow = document.createElement("div");
  groupRow.className = "picker-row";
  groupRow.appendChild(groupInput);
  groupRow.appendChild(groupPickBtn);
  groupField.appendChild(groupRow);

  // Scalanie grup — Dokładnie / Rozmyte / Wzorzec (rejestr AGGREGATION_GROUP_TRANSFORMS)
  // Steruje TYM, jak karty agregacji scalają warianty tej samej wartości (np. "Gr 1 J. Kowalski" → "J. Kowalski").
  const groupModeField = document.createElement("label");
  groupModeField.className = "field";
  groupModeField.append(t("aggGroupMode"));
  const groupModeSelect = document.createElement("select");
  groupModeSelect.dataset.aggregationControl = "groupmode";
  groupModeSelect.dataset.hint = "";
  groupModeSelect.dataset.hintTouch = "on";
  groupModeSelect.dataset.hintDuration = "7";
  groupModeSelect.dataset.hintFade = "";
  groupModeSelect.dataset.hintPl = "Scalanie grup: liczy różne warianty tej samej wartości /| jako jedną grupę w kartach. /| Rozmyte = automat (np. „Gr 1 J. Kowalski” → „J. Kowalski”). /| Wzorzec = własna reguła z = i * # @.";
  groupModeSelect.dataset.hintEn = "Group merging: counts variants of the same value /| as one group in the cards. /| Fuzzy = automatic (e.g. “Gr 1 J. Kowalski” → “J. Kowalski”). /| Pattern = your own rule with = and * # @.";
  AGGREGATION_GROUP_TRANSFORMS.forEach((def) => {
    const opt = document.createElement("option");
    opt.value = def.id;
    opt.textContent = t(def.labelKey);
    groupModeSelect.appendChild(opt);
  });
  groupModeSelect.value = aggregationWorkbenchState.groupMode;
  groupModeField.appendChild(groupModeSelect);

  const activeGroupDef = getAggregationGroupTransformDef(aggregationWorkbenchState.groupMode);
  const groupPatternInput = document.createElement("input");
  groupPatternInput.type = "text";
  groupPatternInput.className = "aggregation-measurefilter-value";
  groupPatternInput.dataset.aggregationControl = "grouppattern";
  groupPatternInput.value = aggregationWorkbenchState.groupPattern || "";
  groupPatternInput.placeholder = t("aggGroupPatternPlaceholder");
  // Bogata, interaktywna legenda składni przez silnik cursor-hint (PL/EN, touch, dłuższy czas)
  groupPatternInput.dataset.hint = "";
  groupPatternInput.dataset.hintTouch = "on";
  groupPatternInput.dataset.hintDuration = "9";
  groupPatternInput.dataset.hintFade = "0.4";
  groupPatternInput.dataset.hintPl = "Wzorzec scalania.  =  to RDZEŃ (zostaje, to liczymy). /| Przed/po = wpisz śmieć do zdjęcia: /| *  jeden token (krótki) · **  dłuższy /| #  cyfry · @  litery · ?  jeden znak /| „Gr” itp. = dosłownie (ignoruje wielkość liter). /| Przykłady:  Gr*=  → „Gr 1 J. Kowalski” = „J. Kowalski”;   =#  → „Faktura12” = „Faktura”.";
  groupPatternInput.dataset.hintEn = "Merge pattern.  =  is the CORE (kept, what we count). /| Before/after = type the junk to strip: /| *  one token (short) · **  longer /| #  digits · @  letters · ?  one char /| „Gr” etc. = literal (case-insensitive). /| Examples:  Gr*=  → „Gr 1 J. Kowalski” = „J. Kowalski”;   =#  → „Faktura12” = „Faktura”.";
  groupPatternInput.style.display = activeGroupDef.needsPattern ? "inline-block" : "none";
  groupModeField.appendChild(groupPatternInput);

  const measureField = document.createElement("label");
  measureField.className = "field";
  measureField.append(t("aggregationMeasure"));
  const measureInput = document.createElement("input");
  measureInput.type = "text";
  measureInput.readOnly = true;
  measureInput.className = "aggregation-measure-input";
  const selectedMeasureLabels = aggregationWorkbenchState.measures
    .map((key) => {
      const candidate = result.measures.find((c) => c.key === key);
      return candidate ? candidate.label : key;
    })
    .filter(Boolean);
  measureInput.value = selectedMeasureLabels.length
    ? selectedMeasureLabels.join(", ")
    : t("aggregationMeasure") + "...";
  measureInput.title = selectedMeasureLabels.join(", ");
  const measurePickBtn = document.createElement("button");
  measurePickBtn.type = "button";
  measurePickBtn.className = "btn ghost btn-sm";
  measurePickBtn.dataset.aggregationControl = "measure-pick";
  measurePickBtn.textContent = t("choose");
  const measureRow = document.createElement("div");
  measureRow.className = "picker-row";
  measureRow.appendChild(measureInput);
  measureRow.appendChild(measurePickBtn);
  measureField.appendChild(measureRow);

  const aggregationField = document.createElement("label");
  aggregationField.className = "field";
  aggregationField.append(t("aggregationMethod"));
  const aggregationSelect = document.createElement("select");
  aggregationSelect.dataset.aggregationControl = "aggregation";
  aggregationField.appendChild(aggregationSelect);
  const aggregationTooltips = {
    count: t("aggregationTooltipCount"),
    avg: t("aggregationTooltipAvg"),
    median: t("aggregationTooltipMedian"),
    min: t("aggregationTooltipMin"),
    max: t("aggregationTooltipMax"),
    sum: t("aggregationTooltipSum"),
    distinct: t("aggregationTooltipDistinct"),
    earliest: t("aggregationTooltipEarliest"),
    latest: t("aggregationTooltipLatest"),
  };
  result.allowedAggregations.forEach((key) => {
    const opt = document.createElement("option");
    opt.value = key;
    opt.textContent = aggregationLabels[key];
    opt.title = aggregationTooltips[key] || "";
    aggregationSelect.appendChild(opt);
  });
  aggregationSelect.value = aggregationWorkbenchState.aggregation;

  const measureFilterField = document.createElement("label");
  measureFilterField.className = "field";
  measureFilterField.append(t("aggregationMeasureFilter"));
  const measureFilterSelect = document.createElement("select");
  measureFilterSelect.dataset.aggregationControl = "measurefilter";
  // Doprecyzowanie: filtr miary odejmuje WIERSZE (co wpada do liczenia) —
  // to co innego niż „Scalanie grup", które skleja etykiety już policzonych grup.
  measureFilterSelect.dataset.hint = "";
  measureFilterSelect.dataset.hintTouch = "on";
  measureFilterSelect.dataset.hintDuration = "7";
  measureFilterSelect.dataset.hintFade = "";
  measureFilterSelect.dataset.hintPl = "Filtr miary: liczy tylko WIERSZE, /| w których wartość miary pasuje. /| (To filtr wierszy — co inne­go niż „Scalanie grup”, /| które skleja etykiety już policzonych grup.) /| Np. miara „Status”, Zawiera „aktywny” → licz tylko aktywne.";
  measureFilterSelect.dataset.hintEn = "Measure filter: counts only ROWS /| whose measure value matches. /| (A row filter — unlike „Group merging”, /| which merges labels of already-counted groups.) /| E.g. measure „Status”, Contains „active” → count only active.";
  [
    { value: "all", label: t("aggregationAll"), title: "" },
    { value: "contains", label: t("aggregationContains"), title: "" },
    { value: "exact", label: t("aggregationExact"), title: "" },
  ].forEach((item) => {
    const option = document.createElement("option");
    option.value = item.value;
    option.textContent = item.label;
    option.title = item.title || "";
    measureFilterSelect.appendChild(option);
  });
  measureFilterSelect.value = aggregationWorkbenchState.measureFilterMode || "all";
  measureFilterField.appendChild(measureFilterSelect);

  const measureFilterInput = document.createElement("input");
  measureFilterInput.type = "text";
  measureFilterInput.className = "aggregation-measurefilter-value";
  measureFilterInput.dataset.aggregationControl = "measurefilter-value";
  measureFilterInput.value = aggregationWorkbenchState.measureFilterValue || "";
  measureFilterInput.placeholder = t("aggregationMeasureSearchPlaceholder");
  measureFilterInput.title = t("aggregationMeasureFilter");
  measureFilterInput.style.display = aggregationWorkbenchState.measureFilterMode === "all" ? "none" : "inline-block";
  measureFilterField.appendChild(measureFilterInput);

  const showCountField = document.createElement("label");
  showCountField.className = "field";
  showCountField.append(t("aggregationShowResults"));
  const showCountSelect = document.createElement("select");
  showCountSelect.dataset.aggregationControl = "count";
  ["10", "20", "40", "80", "999"].forEach((value) => {
    const option = document.createElement("option");
    option.value = value;
    option.textContent = value === "999" ? t("aggregationAll") : value;
    showCountSelect.appendChild(option);
  });
  showCountSelect.value = String(aggregationWorkbenchState.showCount);
  showCountField.appendChild(showCountSelect);

  const havingField = document.createElement("label");
  havingField.className = "field";
  havingField.append(t("aggregationGroupFilter"));
  const havingSelect = document.createElement("select");
  havingSelect.dataset.aggregationControl = "having";
  [
    { value: "all", label: t("aggregationAll"), title: "" },
    { value: "above_value", label: t("aggregationHavingAboveValue"), title: "" },
    { value: "above_percent", label: t("aggregationHavingAboveTotal"), title: "" },
    { value: "above_max_percent", label: t("aggregationHavingAboveMax"), title: "" },
  ].forEach((item) => {
    const option = document.createElement("option");
    option.value = item.value;
    option.textContent = item.label;
    option.title = item.title;
    havingSelect.appendChild(option);
  });
  havingSelect.value = aggregationWorkbenchState.havingMode;
  havingField.appendChild(havingSelect);

  const havingValueInput = document.createElement("input");
  havingValueInput.type = "number";
  havingValueInput.className = "aggregation-having-value";
  havingValueInput.dataset.aggregationControl = "having-value";
  havingValueInput.value = String(aggregationWorkbenchState.havingValue);
  havingValueInput.min = "0";
  havingValueInput.step = "1";
  havingValueInput.title = t("aggHavingTitle");
  havingValueInput.style.display = aggregationWorkbenchState.havingMode === "all" ? "none" : "inline-block";
  havingField.appendChild(havingValueInput);

  [sourceField, scopeField, headerField, groupField, measureField, aggregationField, measureFilterField, groupModeField, showCountField, havingField].forEach((field) => controls.appendChild(field));
  aggregationWorkbenchSummaryEl.appendChild(controls);

  const note = document.createElement("div");
  note.className = "duration-analysis-note";
  const headerModeText = aggregationWorkbenchState.headerRowChoice === "auto"
    ? t("aggregationHeaderAuto", { row: result.resolvedHeaderRow })
    : t("aggregationHeaderRow", { row: result.resolvedHeaderRow });
  const havingText = aggregationWorkbenchState.havingMode === "all"
    ? ""
    : aggregationWorkbenchState.havingMode === "above_value"
      ? t("aggregationHavingValue", { value: aggregationWorkbenchState.havingValue })
      : t("aggregationHavingPercent", { value: aggregationWorkbenchState.havingValue });
  note.textContent = t("aggregationNote", {
    source: result.model.mode === "long" ? "Wide-to-Long" : t("aggregationClassic"),
    scope: aggregationWorkbenchState.scopeMode === "all" ? t("aggregationWholeSheet") : t("aggregationCurrentView"),
    header: headerModeText,
    helper: result.helperMode ? t("aggregationHelper") : "",
    depth: groupDepth,
    match: aggregationWorkbenchState.matchMode === "exact" ? t("aggregationExact").toLowerCase() : t("aggregationContains").toLowerCase(),
    having: havingText,
  });
  aggregationWorkbenchSummaryEl.appendChild(note);

  const measureNote = document.createElement("div");
  measureNote.className = "duration-analysis-note";
  const measureLabels = (result.selectedMeasures || [measure].filter(Boolean)).map((item) => item.label).join(", ");
  measureNote.textContent = t("aggregationMeasureNote", { measures: measureLabels || t("aggregationNone") });
  aggregationWorkbenchSummaryEl.appendChild(measureNote);

  const currentSearch = aggregationWorkbenchState.resultSearch || "";
  const filteredEntries = currentSearch
    ? result.entries.filter((e) => e.label.toLowerCase().includes(currentSearch.toLowerCase()))
    : result.entries;
  const searchWrap = document.createElement("div");
  searchWrap.className = "aggregation-result-search-wrap";

  const searchLabel = document.createElement("span");
  searchLabel.textContent = t("aggregationSearch");
  searchLabel.style.fontSize = "12px";
  searchLabel.style.color = "var(--muted)";
  searchWrap.appendChild(searchLabel);

  const resultSearchInput = document.createElement("input");
  resultSearchInput.type = "text";
  resultSearchInput.className = "aggregation-result-search";
  resultSearchInput.placeholder = t("aggregationSearchPlaceholder");
  resultSearchInput.title = t("aggregationSearch");
  resultSearchInput.style.flex = "1";
  resultSearchInput.style.minWidth = "100px";
  resultSearchInput.style.padding = "4px 8px";
  resultSearchInput.style.borderRadius = "var(--r-sm)";
  resultSearchInput.style.border = "1px solid var(--border)";
  resultSearchInput.style.fontSize = "13px";
  resultSearchInput.value = currentSearch;
  searchWrap.appendChild(resultSearchInput);

  const searchCount = document.createElement("span");
  searchCount.style.fontSize = "12px";
  searchCount.style.color = "var(--muted)";
  searchCount.style.marginLeft = "8px";
  searchCount.style.whiteSpace = "nowrap";
  searchCount.textContent = currentSearch ? t("aggregationSearchCount", { visible: filteredEntries.length, total: result.entries.length }) : "";
  searchWrap.appendChild(searchCount);

  aggregationWorkbenchListEl.appendChild(searchWrap);

  const showCount = Math.min(aggregationWorkbenchState.showCount, filteredEntries.length);
  filteredEntries.slice(0, showCount).forEach((entry, index) => {
    const item = document.createElement("div");
    item.className = "aggregation-item";

    const top = document.createElement("div");
    top.className = "duration-person-top";

    const titleWrap = document.createElement("div");
    titleWrap.className = "duration-person-title-wrap";

    const rank = document.createElement("div");
    rank.className = "duration-person-rank";
    rank.textContent = String(index + 1);

    const title = document.createElement("div");
    title.className = "duration-person-title";
    title.textContent = entry.label;

    const value = document.createElement("div");
    value.className = "duration-person-value";
    value.textContent = formatAggregationMetricValue(entry.primary, primaryKind);

    titleWrap.appendChild(rank);
    titleWrap.appendChild(title);
    top.appendChild(titleWrap);
    top.appendChild(value);

    const meta = document.createElement("div");
    meta.className = "duration-person-meta";
    if (primaryKind === "date") {
      meta.textContent = t("aggregationDateMeta", {
        count: entry.count,
        min: formatAggregationMetricValue(entry.min, "date"),
        max: formatAggregationMetricValue(entry.max, "date"),
      });
    } else if (measure?.kind === "text") {
      meta.textContent = t("aggregationTextMeta", {
        count: entry.count,
        distinct: entry.distinct,
      });
    } else {
      meta.textContent = t("aggregationMeta", {
        count: entry.count,
        avg: formatAggregationMetricValue(entry.average, primaryKind),
        median: formatAggregationMetricValue(entry.median, primaryKind),
        min: formatAggregationMetricValue(entry.min, primaryKind),
        max: formatAggregationMetricValue(entry.max, primaryKind),
      });
    }

    const actions = document.createElement("div");
    actions.className = "section-nav-actions";

    const matchSel = document.createElement("select");
    matchSel.className = "aggregation-card-match";
    matchSel.dataset.aggregationControl = "match";
    [
      { value: "contains", label: t("aggregationContains") },
      { value: "exact", label: t("aggregationExact") },
    ].forEach(({ value, label }) => {
      const opt = document.createElement("option");
      opt.value = value;
      opt.textContent = label;
      matchSel.appendChild(opt);
    });
    matchSel.value = aggregationWorkbenchState.matchMode;

    const btn = document.createElement("button");
    btn.className = "btn ghost btn-sm";
    btn.type = "button";
    btn.dataset.aggregationAction = "filter-group";
    btn.dataset.aggregationValue = entry.filterLabel || entry.label;
    btn.textContent = t("aggregationSearchTable");
    actions.appendChild(matchSel);
    actions.appendChild(btn);

    item.appendChild(top);
    item.appendChild(meta);
    item.appendChild(actions);
    aggregationWorkbenchListEl.appendChild(item);
  });
}
