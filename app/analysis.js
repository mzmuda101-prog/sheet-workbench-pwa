// Sheet/workbook insights and analytical helpers.

function valuesEqual(a, b) {
  if (a === b) return true;
  if (a instanceof Date && b instanceof Date) return a.getTime() === b.getTime();
  if ((a === null || a === undefined) && (b === null || b === undefined)) return true;
  return false;
}

function makeHeadersUnique(headers) {
  const seen = new Map();
  return headers.map((header, index) => {
    const base = String(header || t("colFallback", { n: index + 1 })).trim() || t("colFallback", { n: index + 1 });
    const count = seen.get(base) || 0;
    seen.set(base, count + 1);
    return count ? `${base} (${count + 1})` : base;
  });
}

function formatPercent(part, total) {
  if (!total) return "0%";
  return `${Math.round((part / total) * 100)}%`;
}

function createEmptyInsight(text) {
  const el = document.createElement("div");
  el.className = "insight-empty";
  el.textContent = text;
  return el;
}

function renderInsightList(container, items, emptyText) {
  if (!container) return;
  container.replaceChildren();
  if (!items || !items.length) {
    container.appendChild(createEmptyInsight(emptyText));
    return;
  }
  items.forEach((item) => {
    const row = document.createElement("div");
    row.className = `insight-item${item.tone ? ` ${item.tone}` : ""}`;

    const label = document.createElement("div");
    label.className = "insight-label";
    label.textContent = item.label;

    const value = document.createElement("div");
    value.className = "insight-value";
    value.textContent = item.value;

    row.appendChild(label);
    row.appendChild(value);
    container.appendChild(row);
  });
}

function renderInsightFlags(items) {
  if (!insightFlagsEl) return;
  insightFlagsEl.replaceChildren();
  if (!items || !items.length) {
    insightFlagsEl.appendChild(createEmptyInsight(t("noInsightFlags")));
    return;
  }
  items.forEach((item) => {
    const badge = document.createElement("div");
    badge.className = `insight-flag${item.tone ? ` ${item.tone}` : ""}`;
    badge.textContent = item.label;
    insightFlagsEl.appendChild(badge);
  });
}

function cleanSectionLabel(value) {
  return String(value ?? "").replace(/\s+/g, " ").trim();
}

function formatColRange(startColAbs, endColAbs = startColAbs) {
  const start = XLSX.utils.encode_col(startColAbs);
  const end = XLSX.utils.encode_col(endColAbs);
  return start === end ? start : `${start}:${end}`;
}

function getCellDisplayText(sheet, rowAbs, colAbs) {
  if (!sheet) return "";
  const ref = XLSX.utils.encode_cell({ r: rowAbs, c: colAbs });
  const cell = sheet[ref];
  if (!cell) return "";
  return cleanSectionLabel(cell.w ?? cell.v ?? "");
}

function inferSectionKindLabel(kind) {
  if (kind === "table") return t("secKindTable");
  if (kind === "group") return t("secKindGroup");
  if (kind === "candidate") return t("secKindCandidate");
  if (kind === "subheader") return t("secKindSubheader");
  return t("secKindLayout");
}

function addSection(sections, seen, entry) {
  if (!entry || !entry.label) return;
  const key = `${entry.kind}|${entry.label}|${entry.rowIndex0 ?? ""}|${entry.headerRow ?? ""}|${entry.colIndex ?? ""}`;
  if (seen.has(key)) return;
  seen.add(key);
  sections.push(entry);
}

function detectSections(sheet, headerRow, data) {
  if (!sheet || !data || !data.headers || !data.headers.length) return [];
  const range = XLSX.utils.decode_range(sheet["!ref"]);
  const sections = [];
  const seen = new Set();
  const headerAbsRow = headerRow - 1;

  addSection(sections, seen, {
    kind: "table",
    label: t("secTableData"),
    meta: t("secMetaHeader", { row: headerRow, cols: formatColRange(data.startCol || 0, (data.startCol || 0) + data.headers.length - 1) }),
    tone: "info",
    action: "scroll-top",
  });

  const scanHeaderMax = Math.min(range.e.r, range.s.r + 7);
  for (let r = range.s.r; r <= scanHeaderMax; r++) {
    const texts = [];
    let filled = 0;
    for (let c = range.s.c; c <= range.e.c; c++) {
      const txt = getCellDisplayText(sheet, r, c);
      if (!txt) continue;
      filled += 1;
      if (texts.length < 3) texts.push(txt);
    }
    if (filled < 2) continue;
    addSection(sections, seen, {
      kind: "candidate",
      label: r + 1 === headerRow ? t("secCurrentHeaderRow", { row: r + 1 }) : t("secPossibleHeaderRow", { row: r + 1 }),
      meta: texts.join(" • "),
      tone: r + 1 === headerRow ? "info" : "",
      action: r + 1 === headerRow ? "scroll-top" : "set-header",
      headerRow: r + 1,
    });
  }

  const merges = Array.isArray(data.merges) ? data.merges : [];
  merges
    .filter((merge) => merge && merge.s && merge.e)
    .sort((a, b) => (a.s.r - b.s.r) || (a.s.c - b.s.c))
    .forEach((merge) => {
      const colspan = merge.e.c - merge.s.c + 1;
      if (colspan < 2) return;
      const label = getCellDisplayText(sheet, merge.s.r, merge.s.c);
      if (!label) return;
      const isAboveHeader = merge.s.r < headerAbsRow;
      const overlapsTable = merge.e.c >= (data.startCol || 0) && merge.s.c <= (data.startCol || 0) + data.headers.length - 1;
      if (!isAboveHeader && !overlapsTable) return;
      addSection(sections, seen, {
        kind: "group",
        label,
        meta: t("secMetaRowCols", { row: merge.s.r + 1, cols: formatColRange(merge.s.c, merge.e.c) }),
        tone: isAboveHeader ? "" : "info",
        action: overlapsTable ? "scroll-col" : "set-header",
        colIndex: Math.max(0, merge.s.c - (data.startCol || 0)),
        headerRow: merge.s.r + 1,
      });
    });

  baseRows
    .filter((row) => row && row.isSubheader)
    .slice(0, 8)
    .forEach((row) => {
      const firstText = row.values.find((value) => typeof value === "string" && cleanSectionLabel(value));
      if (!firstText) return;
      addSection(sections, seen, {
        kind: "subheader",
        label: cleanSectionLabel(firstText),
        meta: t("secMetaDataRow", { row: row.rowIndex0 + 1 }),
        tone: "",
        action: "scroll-row",
        rowIndex0: row.rowIndex0,
      });
    });

  return sections.slice(0, 14);
}

function renderSections() {
  if (!sectionNavigatorEl) return;
  sectionNavigatorEl.replaceChildren();
  if (!currentSections.length) {
    sectionNavigatorEl.appendChild(createEmptyInsight(t("sectionsEmpty")));
    return;
  }

  currentSections.forEach((section, index) => {
    const item = document.createElement("div");
    item.className = `section-nav-item${section.tone ? ` ${section.tone}` : ""}`;

    const top = document.createElement("div");
    top.className = "section-nav-top";

    const title = document.createElement("div");
    title.className = "section-nav-title";
    title.textContent = section.label;

    const kind = document.createElement("div");
    kind.className = "section-nav-kind";
    kind.textContent = inferSectionKindLabel(section.kind);

    top.appendChild(title);
    top.appendChild(kind);

    const meta = document.createElement("div");
    meta.className = "section-nav-meta";
    meta.textContent = section.meta || t("secFallback");

    const actions = document.createElement("div");
    actions.className = "section-nav-actions";

    const primary = document.createElement("button");
    primary.className = "btn ghost btn-sm";
    primary.type = "button";
    primary.dataset.sectionIndex = String(index);
    primary.dataset.sectionAction = section.action || "scroll-top";
    primary.textContent = section.action === "set-header" ? t("secSetHeader") : t("secJump");
    actions.appendChild(primary);

    item.appendChild(top);
    item.appendChild(meta);
    item.appendChild(actions);
    sectionNavigatorEl.appendChild(item);
  });
}

function renderSheetInspectorSummary() {
  if (!sheetInspectorSummaryEl) return;
  sheetInspectorSummaryEl.replaceChildren();

  if (!currentHeaders.length || !baseRows.length) {
    sheetInspectorSummaryEl.appendChild(createEmptyInsight(t("sheetSummaryEmpty")));
    return;
  }

  const blockCount = currentRepeatingBlocks.reduce((sum, group) => sum + (Array.isArray(group.blocks) ? group.blocks.length : 0), 0);
  const flaggedProfiles = currentColumnProfiles.filter((profile) => Array.isArray(profile.flags) && profile.flags.length).length;
  const chips = [
    { label: t("inspColumns"), value: String(currentHeaders.length) },
    { label: t("inspChipSections"), value: String(currentSections.length), tone: currentSections.length ? "" : "info" },
    { label: t("inspChipBlocks"), value: String(blockCount), tone: blockCount ? "info" : "" },
    { label: t("inspChipFlagged"), value: String(flaggedProfiles), tone: flaggedProfiles ? "warning" : "" },
  ];

  const topProfile = currentColumnProfiles[0];
  if (topProfile) {
    chips.push({
      label: t("inspTopSignal"),
      value: topProfile.flags.length ? `${topProfile.header} • ${t("profFlag_" + topProfile.flags[0])}` : `${topProfile.header} • ${t("profType_" + topProfile.type)}`,
      tone: topProfile.flags.length ? "warning" : "info",
      wide: true,
    });
  }

  chips.forEach((chip) => {
    const item = document.createElement("div");
    item.className = `sheet-inspector-chip${chip.tone ? ` ${chip.tone}` : ""}${chip.wide ? " wide" : ""}`;

    const label = document.createElement("div");
    label.className = "sheet-inspector-chip-label";
    label.textContent = chip.label;

    const value = document.createElement("div");
    value.className = "sheet-inspector-chip-value";
    value.textContent = chip.value;

    item.appendChild(label);
    item.appendChild(value);
    sheetInspectorSummaryEl.appendChild(item);
  });

  const actions = document.createElement("div");
  actions.className = "sheet-inspector-actions";

  const suggestedHeader = currentSections.find((section) => section.action === "set-header" && section.headerRow && section.headerRow !== currentHeaderRow);
  if (suggestedHeader) {
    const btn = document.createElement("button");
    btn.className = "btn ghost btn-sm";
    btn.type = "button";
    btn.dataset.inspectorAction = "set-header";
    btn.dataset.inspectorHeaderRow = String(suggestedHeader.headerRow);
    btn.textContent = t("inspSetHeaderRow", { row: suggestedHeader.headerRow });
    actions.appendChild(btn);
  }

  if (canUseLongView()) {
    const btn = document.createElement("button");
    btn.className = "btn ghost btn-sm";
    btn.type = "button";
    btn.dataset.inspectorAction = "toggle-long";
    btn.textContent = tableViewMode === "long" ? t("inspBackToClassic") : t("inspSwitchToLong");
    actions.appendChild(btn);
  }

  if (topProfile) {
    const btn = document.createElement("button");
    btn.className = "btn ghost btn-sm";
    btn.type = "button";
    btn.dataset.inspectorAction = "focus-col";
    btn.dataset.profileColIndex = String(topProfile.colIdx);
    btn.textContent = t("inspJumpToCol", { header: topProfile.header });
    actions.appendChild(btn);
  }

  if (actions.childNodes.length) {
    sheetInspectorSummaryEl.appendChild(actions);
  }
}

function focusSection(section) {
  if (!section) return;
  if (section.action === "set-header" && section.headerRow) {
    if (autoHeaderRowEl) autoHeaderRowEl.checked = false;
    headerRowEl.value = String(section.headerRow);
    toast(t("sectionHeaderSet", { row: section.headerRow }), "info");
    loadBtn.click();
    return;
  }

  if (section.action === "scroll-row" && Number.isFinite(section.rowIndex0)) {
    const rowEl = tbodyEl.querySelector(`tr[data-row-index="${section.rowIndex0}"]`);
    if (rowEl) {
      rowEl.scrollIntoView({ behavior: "smooth", block: "center", inline: "nearest" });
      return;
    }
    toast(t("sectionOutsideLimit"), "info");
    return;
  }

  if (section.action === "scroll-col" && Number.isFinite(section.colIndex)) {
    const cells = theadEl.querySelectorAll(".guide-row .guide-cell");
    const cell = cells[section.colIndex];
    if (cell && tableWrapEl) {
      const targetLeft = Math.max(0, cell.offsetLeft - 64);
      tableWrapEl.scrollTo({ left: targetLeft, behavior: "smooth" });
      syncHorizontalScrollbar();
      return;
    }
  }

  if (tableWrapEl) {
    tableWrapEl.scrollTo({ top: 0, left: 0, behavior: "smooth" });
    syncHorizontalScrollbar();
  }
}

function parseRepeatedHeader(header) {
  const raw = cleanSectionLabel(header);
  if (!raw) return null;

  // Pattern 1: trailing digit — "Kwota1", "Wartość_2" → base="Kwota", order=1
  const trailingMatch = raw.match(/^(.*?)(\d+)$/);
  if (trailingMatch) {
    const base = cleanSectionLabel(trailingMatch[1]).replace(/[_\-.\s]+$/, "");
    const order = Number(trailingMatch[2]);
    if (base && Number.isFinite(order)) return { base, order };
  }

  // Pattern 2: middle digit — "Kw1_Kwota", "Q2_Revenue", "M3 Zysk"
  // Cyfra po krótkim prefiksie (max 6 znaków) + separator + reszta nazwy
  const middleMatch = raw.match(/^([A-Za-zĄąĆćĘęŁłŃńÓóŚśŹźŻż]{1,6})(\d+)([_\-. ].+)$/);
  if (middleMatch) {
    const prefix = middleMatch[1];
    const order = Number(middleMatch[2]);
    const suffix = middleMatch[3].replace(/^[_\-. ]+/, "");
    const base = `${prefix}_${suffix}`;
    if (Number.isFinite(order)) return { base, order };
  }

  return { base: raw, order: 1 };
}

function normalizeAnalysisKey(value) {
  return String(value ?? "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, " ")
    .trim();
}

function pluralizeDays(days) {
  const n = Math.abs(days);
  if (currentLang === "en") return n === 1 ? "day" : "days";
  if (n === 1) return "dzien";
  return "dni";
}

function formatDurationDays(days) {
  if (!Number.isFinite(days)) return t("durationNone");
  const rounded = Math.max(0, Math.round(days));
  const months = Math.floor(rounded / 30);
  const restDays = rounded % 30;
  const parts = [];
  if (months > 0) parts.push(`${months} ${t("durationMonthsShort")}`);
  if (restDays > 0 || !parts.length) parts.push(`${restDays} ${pluralizeDays(restDays)}`);
  return parts.join(" ");
}

function pluralizeEntityLabel(label) {
  if (label === "Osoba") return "Osoby";
  if (label === "Pracownik") return "Pracownicy";
  if (label === "Wlasciciel") return "Wlasciciele";
  return `${label}y`;
}

function computeMedian(values) {
  const nums = values.filter((value) => Number.isFinite(value)).slice().sort((a, b) => a - b);
  if (!nums.length) return null;
  const mid = Math.floor(nums.length / 2);
  if (nums.length % 2 === 1) return nums[mid];
  return (nums[mid - 1] + nums[mid]) / 2;
}

function diffDays(start, end) {
  if (!(start instanceof Date) || !(end instanceof Date)) return null;
  const a = new Date(start.getFullYear(), start.getMonth(), start.getDate());
  const b = new Date(end.getFullYear(), end.getMonth(), end.getDate());
  const ms = b.getTime() - a.getTime();
  if (!Number.isFinite(ms)) return null;
  const days = Math.round(ms / 86400000);
  return days >= 0 ? days : null;
}

function parseDurationDaysFlexible(value) {
  if (typeof value === "number" && Number.isFinite(value)) {
    return value >= 0 ? value : null;
  }
  if (typeof value !== "string") return null;
  const text = normalizeAnalysisKey(value);
  if (!text) return null;
  const monthMatch = text.match(/(\d+(?:[.,]\d+)?)\s*(m|mies|miesiac|miesiace|miesiecy|month|months)\b/);
  const dayMatch = text.match(/(\d+(?:[.,]\d+)?)\s*(d|dzien|dni|day|days)\b/);
  if (!monthMatch && !dayMatch) {
    if (/^\d+(?:[.,]\d+)?$/.test(text)) {
      const numeric = Number(text.replace(",", "."));
      return Number.isFinite(numeric) && numeric >= 0 ? numeric : null;
    }
    return null;
  }
  const months = monthMatch ? Number(monthMatch[1].replace(",", ".")) : 0;
  const days = dayMatch ? Number(dayMatch[1].replace(",", ".")) : 0;
  const total = (months * 30) + days;
  return Number.isFinite(total) && total >= 0 ? total : null;
}

function findAnalysisColumnIndex(candidates, matchers) {
  for (const matcher of matchers) {
    const hit = candidates.find((candidate) => matcher(candidate.norm, candidate.base));
    if (hit) return hit.idx;
  }
  return -1;
}

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

function inferAggregationValueKind(header, profile) {
  const baseHeader = parseRepeatedHeader(header)?.base || cleanSectionLabel(header) || header;
  const norm = normalizeAnalysisKey(baseHeader);
  if (profile?.measureType === "date_range") return "duration";
  if (profile?.dateCount > 0 || /\b(data|date|created|closed|start|end|od|do|termin|deadline)\b/.test(norm)) return "date";
  if (profile?.durationCount > 0 || norm.includes("dlugosc") || norm.includes("czas")) return "duration";
  if (profile?.numericCount > 0) return "number";
  return "text";
}

function formatAggregationMetricValue(value, kind = "number") {
  if (!Number.isFinite(value)) return t("aggregationMissing");
  if (kind === "duration") return formatDurationDays(value);
  if (kind === "date") {
    const date = new Date(value);
    if (Number.isNaN(date.getTime())) return t("aggregationMissing");
    return new Intl.DateTimeFormat(I18N[currentLang]?.locale || "pl-PL", {
      year: "numeric",
      month: "short",
      day: "2-digit",
    }).format(date);
  }
  const rounded = Math.round(value * 100) / 100;
  return String(rounded).replace(".", ",");
}

function collectAggregationProfiles(model) {
  if (!model || !Array.isArray(model.headers) || !Array.isArray(model.rows)) return [];
  return model.headers.map((header, idx) => {
    const profile = {
      header,
      idx,
      nonEmptyCount: 0,
      numericCount: 0,
      durationCount: 0,
      dateCount: 0,
      textCount: 0,
      uniqueValues: new Set(),
    };

    model.rows.forEach((row) => {
      const raw = row.values?.[idx];
      const display = getDisplayValue(row, idx);
      const text = String(display ?? raw ?? "").trim();
      if (!text) return;
      profile.nonEmptyCount += 1;
      profile.uniqueValues.add(normalizeAnalysisKey(text));

      const headerKind = classifyAggregationHeader(header);
      const asDate = parseDateFlexible(raw ?? display);
      if (asDate instanceof Date && (raw instanceof Date || typeof raw === "string" || headerKind === "time")) {
        profile.dateCount += 1;
        return;
      }

      if (typeof raw === "number" && Number.isFinite(raw)) {
        profile.numericCount += 1;
        return;
      }

      const duration = parseDurationDaysFlexible(raw ?? display);
      if (duration !== null) {
        profile.durationCount += 1;
        return;
      }

      if (asDate instanceof Date) {
        profile.dateCount += 1;
        return;
      }

      profile.textCount += 1;
    });

    profile.uniqueCount = profile.uniqueValues.size;
    return profile;
  });
}

function classifyAggregationHeader(header) {
  const baseHeader = parseRepeatedHeader(header)?.base || cleanSectionLabel(header) || header;
  const norm = normalizeAnalysisKey(baseHeader);
  if (/\b(id|uuid|guid|nr|lp|numer|number|no|kod|code|pesel|nip|regon|phone|telefon|email|mail)\b/.test(norm)) return "id";
  if (/\b(imie|nazwisko|osoba|pracownik|employee|person|owner|assignee|agent|user|uzytkownik|manager|opiekun)\b/.test(norm)) return "person";
  if (/\b(status|stage|etap|stan|alert|priority|priorytet|type|typ|rodzaj|category|kategoria|segment|tag)\b/.test(norm)) return "category";
  if (/\b(customer|klient|client|kontrahent|firma|company|vendor|supplier|dostawca|account)\b/.test(norm)) return "entity";
  if (/\b(product|produkt|sku|item|towar|material|usluga|service|projekt|project|task|zadanie)\b/.test(norm)) return "item";
  if (/\b(country|kraj|city|miasto|region|wojewodztwo|powiat|teren|lokalizacja|location|department|dzial|team|zespol)\b/.test(norm)) return "place";
  if (/\b(month|miesiac|week|tydzien|quarter|kwartal|year|rok|data|date|created|closed|start|end|od|do)\b/.test(norm)) return "time";
  return "other";
}

function scoreAggregationGroupProfile(profile, totalRows) {
  if (!profile || profile.nonEmptyCount <= 0) return -1000;
  const nonEmptyRatio = totalRows > 0 ? profile.nonEmptyCount / totalRows : 0;
  const uniqueRatio = profile.nonEmptyCount > 0 ? profile.uniqueCount / profile.nonEmptyCount : 1;
  const textRatio = profile.nonEmptyCount > 0 ? profile.textCount / profile.nonEmptyCount : 0;
  const dateRatio = profile.nonEmptyCount > 0 ? profile.dateCount / profile.nonEmptyCount : 0;
  const numericRatio = profile.nonEmptyCount > 0 ? profile.numericCount / profile.nonEmptyCount : 0;
  const kind = classifyAggregationHeader(profile.header);
  let score = 0;

  score += Math.min(24, nonEmptyRatio * 24);
  if (profile.uniqueCount >= 2 && profile.uniqueCount <= 80) score += 18;
  if (profile.uniqueCount >= 2 && uniqueRatio <= 0.65) score += 18;
  if (profile.uniqueCount >= 2 && uniqueRatio <= 0.25) score += 8;
  if (textRatio >= 0.5) score += 16;
  if (dateRatio >= 0.75) score -= 10;
  if (numericRatio >= 0.75) score -= 16;
  if (profile.uniqueCount <= 1) score -= 28;
  if (profile.uniqueCount > 120 || uniqueRatio > 0.9) score -= 30;
  if (profile.idx <= 4) score += 4;

  if (kind === "person") score += 34;
  if (kind === "category") score += 32;
  if (kind === "entity") score += 24;
  if (kind === "place") score += 22;
  if (kind === "item") score += uniqueRatio <= 0.8 ? 18 : 8;
  if (kind === "time") score += profile.uniqueCount <= 36 ? 10 : -8;
  if (kind === "id") score -= 42;

  return score;
}

function describeAggregationGroupProfile(profile) {
  if (!profile) return "";
  const knownKinds = ["person", "category", "entity", "place", "item", "time", "id", "other"];
  const kind = classifyAggregationHeader(profile.header);
  const kindKey = knownKinds.includes(kind) ? kind : "other";
  return t("aggGroupProfileMeta", {
    kind: t("aggGroupKind_" + kindKey),
    unique: profile.uniqueCount,
    nonEmpty: profile.nonEmptyCount,
  });
}

function detectAggregationDateRangeCandidates(model, profiles) {
  const candidates = [];
  const startRegex = /\b(od|start|data od|from|poczatek|rozpoczecie)\b/;
  const endRegex = /\b(do|end|data do|to|until|koniec|zakonczenie)\b/;

  profiles.forEach((profile, idx) => {
    if (profile.dateCount <= 0) return;
    const base = parseRepeatedHeader(model.headers[idx])?.base || cleanSectionLabel(model.headers[idx]) || model.headers[idx];
    const norm = normalizeAnalysisKey(base);
    if (!startRegex.test(norm)) return;

    let endIdx = -1;
    for (let next = idx + 1; next < profiles.length; next++) {
      if (profiles[next].dateCount <= 0) continue;
      const nextBase = parseRepeatedHeader(model.headers[next])?.base || cleanSectionLabel(model.headers[next]) || model.headers[next];
      const nextNorm = normalizeAnalysisKey(nextBase);
      if (endRegex.test(nextNorm)) {
        endIdx = next;
        break;
      }
      if (next > idx + 2) break;
    }
    if (endIdx < 0) return;

    candidates.push({
      key: `date_range:${idx}:${endIdx}`,
      label: `${model.headers[idx]} -> ${model.headers[endIdx]}`,
      kind: "duration",
      measureType: "date_range",
      startIdx: idx,
      endIdx,
      getValue: (row) => {
        const start = parseDateFlexible(row.values?.[idx] ?? getDisplayValue(row, idx));
        const end = parseDateFlexible(row.values?.[endIdx] ?? getDisplayValue(row, endIdx));
        if (!(start instanceof Date) || !(end instanceof Date)) return null;
        return diffDays(start, end);
      },
    });
  });

  return candidates;
}

function detectAggregationMeasureCandidates(model, profiles) {
  const candidates = [{
    key: "count_rows",
    label: t("measureOccurrences"),
    kind: "count",
    measureType: "count_rows",
    getValue: () => 1,
  }];

  detectAggregationDateRangeCandidates(model, profiles).forEach((candidate) => {
    candidates.push(candidate);
  });

  profiles.forEach((profile) => {
    if (profile.nonEmptyCount <= 0) return;
    const kind = inferAggregationValueKind(profile.header, profile);
    if (profile.numericCount <= 0 && profile.durationCount <= 0 && profile.dateCount <= 0 && kind !== "text") return;
    candidates.push({
      key: `column:${profile.idx}`,
      label: profile.header,
      kind,
      measureType: "column",
      colIdx: profile.idx,
      getValue: (row) => {
        if (kind === "text") {
          const raw = row.values?.[profile.idx];
          return String(getDisplayValue(row, profile.idx) || raw || "").trim();
        }
        if (kind === "date") {
          const date = parseDateFlexible(row.values?.[profile.idx] ?? getDisplayValue(row, profile.idx));
          return date instanceof Date && !Number.isNaN(date.getTime()) ? date.getTime() : null;
        }
        const raw = row.values?.[profile.idx];
        if (typeof raw === "number" && Number.isFinite(raw)) return raw;
        return parseDurationDaysFlexible(raw ?? getDisplayValue(row, profile.idx));
      },
      getRawText: (row) => {
        const raw = row.values?.[profile.idx];
        return String(getDisplayValue(row, profile.idx) || raw || "").trim();
      },
    });
  });

  return candidates;
}

function resolveAggregationGroupOptions(profiles) {
  const totalRows = Math.max(...profiles.map((profile) => profile.nonEmptyCount), 0);
  return profiles
    .filter((profile) => profile.nonEmptyCount > 0)
    .map((profile) => ({
      profile,
      score: scoreAggregationGroupProfile(profile, totalRows),
    }))
    .sort((a, b) => {
      if (Math.abs(b.score - a.score) > 0.001) return b.score - a.score;
      return a.profile.idx - b.profile.idx;
    })
    .map(({ profile, score }) => ({
      value: profile.header,
      label: profile.header,
      idx: profile.idx,
      score,
      meta: describeAggregationGroupProfile(profile),
    }));
}

function chooseDefaultAggregationGroup(groupOptions) {
  if (!groupOptions.length) return "";
  const preferred = groupOptions.find((option) => option.score >= 10);
  return (preferred || groupOptions[0]).value;
}

function chooseDefaultAggregationMeasure(measures) {
  if (!measures.length) return "count_rows";
  const dateRange = measures.find((candidate) => candidate.measureType === "date_range");
  if (dateRange) return dateRange.key;
  const duration = measures.find((candidate) => candidate.kind === "duration");
  if (duration) return duration.key;
  const date = measures.find((candidate) => candidate.kind === "date");
  if (date) return date.key;
  const numeric = measures.find((candidate) => candidate.measureType === "column" && candidate.kind === "number");
  return numeric ? numeric.key : "count_rows";
}

function chooseDefaultAggregationMethod(measure) {
  if (!measure || measure.measureType === "count_rows") return "count";
  if (measure.kind === "text") return "count";
  if (measure.kind === "date") return "earliest";
  return "avg";
}

function resolveSelectedAggregationMeasures(measures) {
  const selectedKeys = aggregationWorkbenchState.measures.filter(Boolean);
  const seen = new Set();
  return selectedKeys
    .map((key) => measures.find((candidate) => candidate.key === key))
    .filter(Boolean)
    .filter((candidate) => {
      if (seen.has(candidate.key)) return false;
      seen.add(candidate.key);
      return true;
    });
}

function getAggregationKindFamily(kind) {
  if (kind === "date") return "date";
  if (kind === "text") return "text";
  if (kind === "count") return "count";
  return "numeric";
}

function filterCompatibleAggregationMeasures(selectedMeasures, aggregation) {
  if (!selectedMeasures.length) return [];
  if (aggregation === "count" || aggregation === "distinct") return selectedMeasures;
  if (aggregation === "earliest" || aggregation === "latest") {
    return selectedMeasures.filter((measure) => measure.kind === "date");
  }
  return selectedMeasures.filter((measure) => {
    const family = getAggregationKindFamily(measure.kind);
    return family === "numeric";
  });
}

function getAllowedAggregationsForMeasures(selectedMeasures) {
  if (!selectedMeasures.length) return ["count"];
  const columnMeasures = selectedMeasures.filter((measure) => measure.measureType !== "count_rows");
  if (!columnMeasures.length) return ["count"];
  const families = new Set(columnMeasures.map((measure) => getAggregationKindFamily(measure.kind)));
  const allowed = ["count", "distinct"];
  if (families.has("numeric")) allowed.push("avg", "median", "min", "max", "sum");
  if (families.has("date")) allowed.push("earliest", "latest");
  return allowed;
}

function getPrimaryAggregationValueKind(measures, aggregation) {
  if (aggregation === "earliest" || aggregation === "latest") return "date";
  if (aggregation === "distinct" || aggregation === "count") return "number";
  if (measures.some((measure) => measure.kind === "duration")) return "duration";
  return "number";
}

function getNormalizedAggregationWorkbenchContext() {
  const headerCandidates = getAggregationHeaderCandidateRows();
  let resolvedHeaderRow = currentHeaderRow;
  let context = null;

  if (aggregationWorkbenchState.headerRowChoice === "auto") {
    headerCandidates.forEach((candidateRow) => {
      const candidateContext = collectAggregationContextForHeaderRow(
        candidateRow,
        aggregationWorkbenchState.sourceMode,
        aggregationWorkbenchState.scopeMode
      );
      const score = scoreAggregationContext(candidateContext);
      if (!context || score > context.score) {
        context = { ...candidateContext, score };
        resolvedHeaderRow = candidateRow;
      }
    });
  } else {
    const explicitRow = Number.isFinite(aggregationWorkbenchState.customHeaderRow)
      ? aggregationWorkbenchState.customHeaderRow
      : currentHeaderRow;
    resolvedHeaderRow = explicitRow > 0 ? explicitRow : currentHeaderRow;
    context = collectAggregationContextForHeaderRow(
      resolvedHeaderRow,
      aggregationWorkbenchState.sourceMode,
      aggregationWorkbenchState.scopeMode
    );
  }

  if (!context) {
    context = collectAggregationContextForHeaderRow(
      currentHeaderRow,
      aggregationWorkbenchState.sourceMode,
      aggregationWorkbenchState.scopeMode
    );
    resolvedHeaderRow = currentHeaderRow;
  }

  const { model, profiles, groupOptions, measures, longAvailable } = context;

  const nextGroupBy = groupOptions.some((option) => option.value === aggregationWorkbenchState.groupBy)
    ? aggregationWorkbenchState.groupBy
    : chooseDefaultAggregationGroup(groupOptions);
  const nextGroupBy2 = groupOptions.some((option) => option.value === aggregationWorkbenchState.groupBy2 && option.value !== nextGroupBy)
    ? aggregationWorkbenchState.groupBy2
    : "";
  const nextGroupBy3 = groupOptions.some((option) => option.value === aggregationWorkbenchState.groupBy3 && option.value !== nextGroupBy && option.value !== nextGroupBy2)
    ? aggregationWorkbenchState.groupBy3
    : "";
  const validMeasures = aggregationWorkbenchState.measures
    .filter((key) => measures.some((candidate) => candidate.key === key));
  const nextMeasures = validMeasures.length
    ? validMeasures
    : [chooseDefaultAggregationMeasure(measures)];
  aggregationWorkbenchState.measures = nextMeasures;
  const measure = measures.find((candidate) => candidate.key === nextMeasures[0]) || measures[0] || null;
  const selectedMeasures = resolveSelectedAggregationMeasures(measures);
  const allowedAggregations = getAllowedAggregationsForMeasures(selectedMeasures);
  const nextAggregation = allowedAggregations.includes(aggregationWorkbenchState.aggregation)
    ? aggregationWorkbenchState.aggregation
    : chooseDefaultAggregationMethod(measure);

  aggregationWorkbenchState.groupBy = nextGroupBy;
  aggregationWorkbenchState.groupBy2 = nextGroupBy2;
  aggregationWorkbenchState.groupBy3 = nextGroupBy3;
  aggregationWorkbenchState.aggregation = nextAggregation;
  if (aggregationWorkbenchState.sourceMode === "long" && !longAvailable) {
    aggregationWorkbenchState.sourceMode = "auto";
  }

  return {
    ...context,
    model,
    profiles,
    groupOptions,
    measures,
    measure,
    selectedMeasures: filterCompatibleAggregationMeasures(selectedMeasures, nextAggregation),
    selectedMeasuresRaw: selectedMeasures,
    longAvailable,
    allowedAggregations,
    headerCandidates,
    resolvedHeaderRow,
  };
}

function computeAggregateMetric(values, aggregation) {
  if (aggregation === "distinct") {
    const unique = new Set(values.map((v) => String(v).trim()).filter((v) => v));
    return unique.size;
  }
  if (aggregation === "count") {
    return values.filter((value) => {
      if (Number.isFinite(value)) return true;
      return String(value ?? "").trim() !== "";
    }).length;
  }
  const nums = values.filter((value) => Number.isFinite(value));
  if (!nums.length) return null;
  if (aggregation === "earliest") return Math.min(...nums);
  if (aggregation === "latest") return Math.max(...nums);
  if (aggregation === "sum") return nums.reduce((sum, value) => sum + value, 0);
  if (aggregation === "avg") return nums.reduce((sum, value) => sum + value, 0) / nums.length;
  if (aggregation === "median") return computeMedian(nums);
  if (aggregation === "min") return Math.min(...nums);
  if (aggregation === "max") return Math.max(...nums);
  return null;
}

// ── Scalanie grup wzorcem (pattern group-merging) ──────────────────────────────
// Cel: różne warianty tej samej wartości liczyć jako JEDNĄ grupę w kartach agregacji.
//   np. "Gr 1 J. Kowalski", "Gr1 J. Kowalski", "J. Kowalski" → wszystkie jako "J. Kowalski".
//
// Składnia wzorca (inspirowana SQL LIKE / glob, ale czytelniejsza):
//   =          RDZEŃ — część, która zostaje i definiuje grupę (to, co zliczamy)
//   literał    np. "Gr" — musi pasować dosłownie (ignoruje wielkość liter)
//   *          cokolwiek, krótko        (≤ GROUP_WILD_SHORT znaków)
//   **         cokolwiek, dłużej        (≤ GROUP_WILD_LONG znaków)
//   #          tylko cyfry              (ciąg 1..GROUP_WILD_LONG)
//   @          tylko litery             (ciąg 1..GROUP_WILD_LONG)
//   ?          dokładnie jeden znak
//   Tekst przed "=" to wiodący śmieć do ZDJĘCIA, po "=" to końcowy śmieć do zdjęcia.
//   Śmieć zdejmujemy tylko, gdy pasuje (opcjonalnie) — wartości bez śmiecia zostają.
// Przykłady:  Gr*=   → "Gr 1 J. Kowalski" → "J. Kowalski";  gołe "J. Kowalski" bez zmian
//             =#     → "Faktura12" → "Faktura"  (zdejmij końcowe cyfry)
//             @#=    → "AB12 Kowalski" → "Kowalski"
const GROUP_WILD_SHORT = 3;    // pojedyncze *  (jeden krótki token śmiecia)
const GROUP_WILD_LONG = 12;    // ** / # / @    (dłuższy ciąg)
const GROUP_LETTER_CLASS = "A-Za-zÀ-ÿĄ-ž";  // litery łac. + PL/diakrytyki dla @
const GROUP_SEP_CLASS = "\\s._\\-";          // separatory: spacja, _ . -

// Zamień jedną stronę wzorca (przed lub po "=") na fragment regexpa.
// `*`/`**` to "token śmiecia" ograniczony separatorem przy styku z rdzeniem —
// dzięki temu nie wjeżdża w rdzeń (np. "Gr1 J. Kowalski" → zdejmuje "Gr1 ", nie "Gr1 J").
// `#`/`@`/`?` są klasowe (cyfry/litery/jeden znak) i nie wymagają separatora.
// Zwraca null, jeśli strona jest pusta (brak śmiecia do zdjęcia).
function buildGroupNoiseRegex(side, anchor /* "lead" | "trail" */) {
  const s = String(side || "");
  if (!s) return null;
  let body = "";
  for (let i = 0; i < s.length; i++) {
    const ch = s[i];
    if (ch === "*") {
      if (s[i + 1] === "*") { body += `.{0,${GROUP_WILD_LONG}}`; i++; }
      else { body += `.{0,${GROUP_WILD_SHORT}}`; }
    } else if (ch === "#") {
      body += `[0-9]{1,${GROUP_WILD_LONG}}`;
    } else if (ch === "@") {
      body += `[${GROUP_LETTER_CLASS}]{1,${GROUP_WILD_LONG}}`;
    } else if (ch === "?") {
      body += ".";
    } else {
      body += ch.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"); // literał — escape
    }
  }
  let pattern;
  if (anchor === "lead") {
    // śmieć wiodący kończy się "any"-wildcardem → dopasowanie kończ na separatorze
    pattern = /\*$/.test(s) ? `^(?:${body})(?<=[${GROUP_SEP_CLASS}])` : `^(?:${body})`;
  } else {
    // śmieć końcowy zaczyna się "any"-wildcardem → musi startować po separatorze
    pattern = /^\*/.test(s) ? `(?<=[${GROUP_SEP_CLASS}])(?:${body})$` : `(?:${body})$`;
  }
  try { return new RegExp(pattern, "iu"); } catch { return null; }
}

function compileGroupPattern(pattern) {
  const raw = String(pattern || "").trim();
  const coreMatch = raw.match(/=+/);
  if (!coreMatch) {
    // bez kotwicy "=" nie wiadomo, co jest rdzeniem → tożsamość (bezpieczny fallback)
    return (label) => String(label ?? "").trim();
  }
  const lead = raw.slice(0, coreMatch.index);
  const trail = raw.slice(coreMatch.index + coreMatch[0].length);
  const leadRe = buildGroupNoiseRegex(lead, "lead");
  const trailRe = buildGroupNoiseRegex(trail, "trail");
  const edgeTrim = new RegExp(`^[${GROUP_SEP_CLASS}]+|[${GROUP_SEP_CLASS}]+$`, "gu");
  return (label) => {
    const original = String(label ?? "").trim();
    let core = original;
    if (trailRe) core = core.replace(trailRe, "");
    if (leadRe) core = core.replace(leadRe, "");
    core = core.replace(edgeTrim, "").trim();
    return core || original; // nigdy nie redukuj do pustego (np. "123" przy =#)
  };
}

// Preset "Rozmyte" — bez wpisywania. Zdejmuje wiodący znacznik grupy:
// krótki token liter + cyfry + separator, ale TYLKO gdy zawiera cyfrę i poprzedza literę.
// "Gr 1 J. Kowalski" → "J. Kowalski";  "J. Kowalski" (inicjał, brak cyfry) → bez zmian.
const FUZZY_LEAD_MARKER = /^[A-Za-zÀ-ÿĄ-ž]{1,6}[\s.]*[0-9]{1,4}[\s._\-]+(?=[A-Za-zÀ-ÿĄ-ž])/u;
function fuzzyGroupTransform(label) {
  const original = String(label ?? "").trim();
  const core = original.replace(FUZZY_LEAD_MARKER, "").trim();
  return core || original;
}

// Rejestr trybów scalania grup — ROZSZERZALNY: dodaj obiekt = nowa opcja w UI.
//   id           unikalny klucz (zapis w stanie)
//   labelKey     klucz i18n etykiety w <select>
//   needsPattern czy pokazać input na wzorzec
//   build(opts) → (label) => coreLabel   (opts.pattern = aktualny wzorzec)
const AGGREGATION_GROUP_TRANSFORMS = [
  {
    id: "exact",
    labelKey: "aggGroupExact",
    needsPattern: false,
    build: () => (label) => String(label ?? "").trim(),
  },
  {
    id: "fuzzy",
    labelKey: "aggGroupFuzzy",
    needsPattern: false,
    build: () => fuzzyGroupTransform,
  },
  {
    id: "pattern",
    labelKey: "aggGroupPattern",
    needsPattern: true,
    build: (opts) => compileGroupPattern(opts && opts.pattern),
  },
];

function getAggregationGroupTransformDef(id) {
  return AGGREGATION_GROUP_TRANSFORMS.find((def) => def.id === id) || AGGREGATION_GROUP_TRANSFORMS[0];
}

function getActiveGroupTransform() {
  const def = getAggregationGroupTransformDef(aggregationWorkbenchState.groupMode);
  try {
    return def.build({ pattern: aggregationWorkbenchState.groupPattern }) || ((label) => String(label ?? "").trim());
  } catch {
    return (label) => String(label ?? "").trim();
  }
}

function buildAggregationWorkbenchResult() {
  const context = getNormalizedAggregationWorkbenchContext();
  const { model, groupOptions, measures, measure } = context;
  if (!model?.headers?.length || !model?.rows?.length) {
    return { status: "empty", ...context };
  }
  if (!groupOptions.length || !measure) {
    return { status: "no-options", ...context };
  }

  const groupHeaders = [
    aggregationWorkbenchState.groupBy,
    aggregationWorkbenchState.groupBy2,
    aggregationWorkbenchState.groupBy3,
  ].filter(Boolean);
  const groupIndexes = groupHeaders.map((header) => model.headers.indexOf(header)).filter((idx) => idx >= 0);
  if (!groupIndexes.length) {
    return { status: "no-options", ...context };
  }

  const isDistinct = aggregationWorkbenchState.aggregation === "distinct";
  const activeMeasures = context.selectedMeasures?.length ? context.selectedMeasures : [measure].filter(Boolean);
  const measureFilterMode = aggregationWorkbenchState.measureFilterMode || "all";
  const measureFilterValue = aggregationWorkbenchState.measureFilterValue || "";
  const buckets = new Map();
  const groupTransform = getActiveGroupTransform();
  model.rows.forEach((row) => {
    if (measureFilterMode !== "all" && measureFilterValue && activeMeasures.length) {
      const rowMeasureText = activeMeasures
        .map((activeMeasure) => activeMeasure.getRawText
          ? activeMeasure.getRawText(row)
          : getDisplayValue(row, activeMeasure.colIdx) || "")
        .join(" ");
      const rowMeasureLower = rowMeasureText.toLowerCase();
      const filterLower = measureFilterValue.toLowerCase();
      if (measureFilterMode === "contains" && !rowMeasureLower.includes(filterLower)) return;
      if (measureFilterMode === "exact" && rowMeasureLower !== filterLower) return;
    }
    const groupLabels = groupIndexes.map((groupIdx) => {
      const rawGroup = row.values?.[groupIdx];
      const display = String(getDisplayValue(row, groupIdx) || rawGroup || "(puste)").trim() || "(puste)";
      // Wzorzec/preset scala warianty do wspólnego rdzenia (np. Osoba1 → Osoba)
      return groupTransform(display) || display;
    });
    const groupLabel = groupLabels.join(" / ");
    const key = groupLabels.map((label) => normalizeAnalysisKey(label) || "(puste)").join("\u001f");
    const entry = buckets.get(key) || {
      label: groupLabel,
      groupLabels,
      values: [],
      rowIndexes: new Set(),
    };
    if (isDistinct) {
      activeMeasures.forEach((activeMeasure) => {
        const rawValue = activeMeasure.getRawText
          ? activeMeasure.getRawText(row)
          : activeMeasure.colIdx != null
            ? getDisplayValue(row, activeMeasure.colIdx) || ""
            : "";
        if (rawValue) entry.values.push(rawValue);
      });
    } else {
      if (activeMeasures.some((activeMeasure) => activeMeasure.measureType === "count_rows")) {
        entry.values.push(1);
      } else {
        activeMeasures.forEach((activeMeasure) => {
          const measureValue = activeMeasure.getValue(row);
          if (Number.isFinite(measureValue)) {
            entry.values.push(measureValue);
          } else if (typeof measureValue === "string" && measureValue.trim()) {
            entry.values.push(measureValue.trim());
          }
        });
      }
    }
    entry.rowIndexes.add(row.rowIndex0);
    buckets.set(key, entry);
  });

  const rawEntries = Array.from(buckets.values())
    .map((entry) => {
      const count = computeAggregateMetric(entry.values, "count");
      return {
        label: entry.label,
        groupLabels: entry.groupLabels,
        filterLabel: entry.groupLabels?.[0] || entry.label,
        count,
        average: computeAggregateMetric(entry.values, "avg"),
        median: computeAggregateMetric(entry.values, "median"),
        min: computeAggregateMetric(entry.values, "min"),
        max: computeAggregateMetric(entry.values, "max"),
        sum: computeAggregateMetric(entry.values, "sum"),
        distinct: computeAggregateMetric(entry.values, "distinct"),
        primary: computeAggregateMetric(entry.values, aggregationWorkbenchState.aggregation),
      };
    })
    .filter((entry) => {
      if (aggregationWorkbenchState.aggregation === "distinct") return true;
      return entry.count > 0 || aggregationWorkbenchState.aggregation === "count";
    });

  const totalPrimary = rawEntries.reduce((sum, e) => sum + (e.primary || 0), 0);
  const maxPrimary = rawEntries.length > 0 ? Math.max(...rawEntries.map(e => e.primary || 0)) : 0;

  const entries = rawEntries
    .filter((entry) => {
      if (aggregationWorkbenchState.havingMode === "all") return true;
      const primary = entry.primary || 0;
      const value = aggregationWorkbenchState.havingValue;
      if (aggregationWorkbenchState.havingMode === "above_value") return primary > value;
      if (aggregationWorkbenchState.havingMode === "above_percent") return totalPrimary > 0 && (primary / totalPrimary) * 100 > value;
      if (aggregationWorkbenchState.havingMode === "above_max_percent") return maxPrimary > 0 && (primary / maxPrimary) * 100 > value;
      return true;
    })
    .sort((a, b) => {
      const sortMultiplier = aggregationWorkbenchState.aggregation === "earliest" ? 1 : -1;
      const diff = (Number(a.primary || 0) - Number(b.primary || 0)) * sortMultiplier;
      if (Math.abs(diff) > 0.001) return diff;
      const countDiff = b.count - a.count;
      if (countDiff) return countDiff;
      return a.label.localeCompare(b.label, "pl");
    });

  if (!entries.length) {
    return { status: "no-results", ...context };
  }

  return {
    status: "ok",
    ...context,
    entries,
    summary: {
      groups: entries.length,
      sourceRows: model.rows.length,
      measuredRows: entries.reduce((sum, entry) => sum + entry.count, 0),
    },
  };
}

function renderAggregationWorkbench() {
  renderMonthlySummary(); // panel „Podsumowanie miesięczne" odświeża się tym samym cyklem
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

function buildMergedLabelMap(merges, sheet, tableStartCol, tableEndCol, headerAbsRow) {
  const labels = new Map();
  merges
    .filter((merge) => merge && merge.s && merge.e && merge.s.r < headerAbsRow)
    .forEach((merge) => {
      if (merge.e.c < tableStartCol || merge.s.c > tableEndCol) return;
      const label = getCellDisplayText(sheet, merge.s.r, merge.s.c);
      if (!label) return;
      const startIndex = Math.max(0, merge.s.c - tableStartCol);
      labels.set(startIndex, label);
    });
  return labels;
}

function isMeaningfulSheetCell(cell) {
  if (!cell || typeof cell !== "object") return false;
  if (cell.f) return true;
  if (cell.l && (cell.l.Target || cell.l.target)) return true;
  if (Array.isArray(cell.c) && cell.c.length) return true;
  if (cell.v instanceof Date) return true;
  if (typeof cell.v === "number" && Number.isFinite(cell.v)) return true;
  if (typeof cell.v === "boolean") return true;
  if (typeof cell.v === "string" && cell.v.trim() !== "") return true;
  if (typeof cell.w === "string" && cell.w.trim() !== "") return true;
  return false;
}

function computeEffectiveSheetRange(sheet, headerRow) {
  const fallback = XLSX.utils.decode_range(sheet["!ref"]);
  const headerIndex0 = Math.max(0, (headerRow || 1) - 1);
  let minCol = fallback.e.c;
  let maxCol = fallback.s.c;
  let maxRow = headerIndex0;
  let found = false;

  Object.keys(sheet).forEach((key) => {
    if (!key || key[0] === "!") return;
    const cell = sheet[key];
    if (!isMeaningfulSheetCell(cell)) return;
    const ref = XLSX.utils.decode_cell(key);
    if (ref.r < headerIndex0) {
      minCol = Math.min(minCol, ref.c);
      maxCol = Math.max(maxCol, ref.c);
      found = true;
      return;
    }
    minCol = Math.min(minCol, ref.c);
    maxCol = Math.max(maxCol, ref.c);
    maxRow = Math.max(maxRow, ref.r);
    found = true;
  });

  for (let c = fallback.s.c; c <= fallback.e.c; c++) {
    const cell = sheet[XLSX.utils.encode_cell({ r: headerIndex0, c })];
    if (!isMeaningfulSheetCell(cell)) continue;
    minCol = Math.min(minCol, c);
    maxCol = Math.max(maxCol, c);
    found = true;
  }

  if (!found) {
    return fallback;
  }

  const merges = Array.isArray(sheet["!merges"]) ? sheet["!merges"] : [];
  merges.forEach((merge) => {
    if (!merge || !merge.s || !merge.e) return;
    const anchor = sheet[XLSX.utils.encode_cell({ r: merge.s.r, c: merge.s.c })];
    if (!isMeaningfulSheetCell(anchor)) return;
    minCol = Math.min(minCol, merge.s.c);
    maxCol = Math.max(maxCol, merge.e.c);
    if (merge.e.r >= headerIndex0) {
      maxRow = Math.max(maxRow, merge.e.r);
    }
  });

  return {
    s: { r: fallback.s.r, c: Math.max(fallback.s.c, minCol) },
    e: { r: Math.max(headerIndex0, maxRow), c: Math.max(Math.max(fallback.s.c, minCol), maxCol) },
  };
}

function buildGroupFromSignature(headers, startIndex, span, repeatCount, tableStartCol, mergedLabels) {
  const bases = headers
    .slice(startIndex, startIndex + span)
    .map((header) => parseRepeatedHeader(header)?.base || cleanSectionLabel(header))
    .filter(Boolean);
  const uniqueBases = Array.from(new Set(bases));
  const blocks = [];

  for (let i = 0; i < repeatCount; i++) {
    const blockStart = startIndex + (i * span);
    const blockEnd = blockStart + span - 1;
    blocks.push({
      label: mergedLabels.get(blockStart) || t("blkFallbackLabel", { n: i + 1 }),
      span,
      startIndex: blockStart,
      endIndex: blockEnd,
      startAbs: tableStartCol + blockStart,
      endAbs: tableStartCol + blockEnd,
      headers: headers.slice(blockStart, blockEnd + 1),
    });
  }

  return {
    label: repeatCount >= 2 ? t("blkSignatureLabel", { n: repeatCount }) : t("blkSignatureLabelSingle"),
    kind: "repeating-signature",
    meta: t("blkSignatureMeta", { n: repeatCount, span }),
    prefixCount: startIndex,
    prefixLabel: startIndex > 0 ? formatColRange(tableStartCol, tableStartCol + startIndex - 1) : "",
    longHeaders: uniqueBases.slice(),
    families: uniqueBases.slice(0, 8).map((label) => ({ label, count: repeatCount })),
    blocks,
  };
}

function detectSignatureRepeatingBlocks(headers, tableStartCol, mergedLabels) {
  if (!Array.isArray(headers) || headers.length < 4) return [];

  let best = null;
  const maxSpan = Math.min(12, Math.floor(headers.length / 2));

  for (let startIndex = 0; startIndex < headers.length - 3; startIndex++) {
    for (let span = 2; span <= maxSpan; span++) {
      if (startIndex + (span * 2) > headers.length) break;

      const template = headers.slice(startIndex, startIndex + span).map((header) => parseRepeatedHeader(header)?.base || cleanSectionLabel(header));
      if (!template.some(Boolean)) continue;

      let repeatCount = 1;
      let nextStart = startIndex + span;

      while (nextStart + span <= headers.length) {
        const candidate = headers.slice(nextStart, nextStart + span).map((header) => parseRepeatedHeader(header)?.base || cleanSectionLabel(header));
        if (candidate.length !== template.length) break;
        if (!candidate.every((value, idx) => value === template[idx])) break;
        repeatCount += 1;
        nextStart += span;
      }

      if (repeatCount < 2) continue;

      const score = (repeatCount * span * 100) - startIndex;
      if (!best || score > best.score) {
        best = { score, startIndex, span, repeatCount };
      }
    }
  }

  if (!best) return [];
  return [buildGroupFromSignature(headers, best.startIndex, best.span, best.repeatCount, tableStartCol, mergedLabels)];
}

function detectSuffixedCycleBlocks(headers, tableStartCol, mergedLabels) {
  if (!Array.isArray(headers) || headers.length < 2) return [];

  const baseOrder = [];
  const baseSeen = new Set();
  const byBase = new Map();
  const byOrder = new Map();

  headers.forEach((header, index) => {
    const parsed = parseRepeatedHeader(header);
    if (!parsed || !parsed.base) return;
    const base = parsed.base;
    const order = Number(parsed.order) || 1;
    if (!baseSeen.has(base)) {
      baseSeen.add(base);
      baseOrder.push(base);
    }
    const baseEntry = byBase.get(base) || [];
    baseEntry.push({ index, order, header });
    byBase.set(base, baseEntry);

    const orderEntry = byOrder.get(order) || new Map();
    if (!orderEntry.has(base)) orderEntry.set(base, { index, header });
    byOrder.set(order, orderEntry);
  });

  const repeatedBases = baseOrder.filter((base) => {
    const orders = new Set((byBase.get(base) || []).map((entry) => entry.order));
    return orders.size >= 2;
  });
  if (!repeatedBases.length) return [];

  const minPresent = repeatedBases.length === 1 ? 1 : Math.max(2, Math.ceil(repeatedBases.length * 0.5));
  const blocks = Array.from(byOrder.entries())
    .sort(([a], [b]) => a - b)
    .map(([order, entries]) => {
      const valueIndexes = repeatedBases.map((base) => entries.get(base)?.index ?? -1);
      const presentIndexes = valueIndexes.filter((idx) => idx >= 0);
      if (presentIndexes.length < minPresent) return null;
      const startIndex = Math.min(...presentIndexes);
      const endIndex = Math.max(...presentIndexes);
      return {
        label: mergedLabels.get(startIndex) || `Cykl ${order}`,
        order,
        span: repeatedBases.length,
        startIndex,
        endIndex,
        startAbs: tableStartCol + startIndex,
        endAbs: tableStartCol + endIndex,
        headers: repeatedBases.slice(),
        valueIndexes,
      };
    })
    .filter(Boolean);

  if (blocks.length < 2) return [];

  const firstStart = Math.min(...blocks.map((block) => block.startIndex));
  return [{
    label: t("blkCycleLabel", { n: blocks.length }),
    kind: "suffixed-cycle",
    meta: t("blkCycleMeta", { n: blocks.length, fields: repeatedBases.length }),
    prefixCount: firstStart,
    prefixLabel: firstStart > 0 ? formatColRange(tableStartCol, tableStartCol + firstStart - 1) : "",
    longHeaders: repeatedBases.slice(),
    families: repeatedBases.slice(0, 10).map((label) => ({ label, count: blocks.length })),
    blocks,
  }];
}

function detectRepeatingBlocks(sheet, headerRow, data) {
  if (!sheet || !data || !Array.isArray(data.headers) || !data.headers.length) return [];
  const merges = Array.isArray(data.merges) ? data.merges : [];
  const headerAbsRow = headerRow - 1;
  const tableStartCol = data.startCol || 0;
  const tableEndCol = tableStartCol + data.headers.length - 1;
  const mergedLabels = buildMergedLabelMap(merges, sheet, tableStartCol, tableEndCol, headerAbsRow);

  const signatureGroups = detectSignatureRepeatingBlocks(data.headers, tableStartCol, mergedLabels);
  if (signatureGroups.length) {
    return signatureGroups;
  }

  const groups = [];

  const mergeBlocks = merges
    .filter((merge) => merge && merge.s && merge.e && merge.s.r < headerAbsRow && merge.e.c >= tableStartCol && merge.s.c <= tableEndCol)
    .sort((a, b) => a.s.c - b.s.c);

  if (mergeBlocks.length >= 2) {
    const bySpan = new Map();
    mergeBlocks.forEach((merge) => {
      const span = merge.e.c - merge.s.c + 1;
      if (span < 2) return;
      const label = getCellDisplayText(sheet, merge.s.r, merge.s.c);
      if (!label) return;
      const list = bySpan.get(span) || [];
      const startIndex = Math.max(0, merge.s.c - tableStartCol);
      const endIndex = Math.min(data.headers.length - 1, merge.e.c - tableStartCol);
      list.push({
        label,
        span,
        startIndex,
        endIndex,
        startAbs: merge.s.c,
        endAbs: merge.e.c,
        headers: data.headers.slice(startIndex, endIndex + 1),
      });
      bySpan.set(span, list);
    });

    bySpan.forEach((blocks, span) => {
      if (blocks.length < 2) return;
      const familyMap = new Map();
      blocks.forEach((block) => {
        block.headers.forEach((header) => {
          const parsed = parseRepeatedHeader(header);
          const key = parsed ? parsed.base : header;
          familyMap.set(key, (familyMap.get(key) || 0) + 1);
        });
      });
      const families = Array.from(familyMap.entries())
        .filter(([, count]) => count >= 2)
        .map(([label, count]) => ({ label, count }))
        .slice(0, 8);

      groups.push({
        label: t("blkSignatureLabel", { n: blocks.length }),
        kind: "merged",
        meta: t("blkMergedMeta", { n: blocks.length, span, first: blocks[0].label, last: blocks[blocks.length - 1].label }),
        prefixCount: blocks[0].startIndex,
        prefixLabel: blocks[0].startIndex > 0 ? formatColRange(tableStartCol, tableStartCol + blocks[0].startIndex - 1) : "",
        families,
        blocks,
      });
    });
  }

  if (groups.length) return groups.slice(0, 4);

  return detectSuffixedCycleBlocks(data.headers, tableStartCol, mergedLabels);
}

function renderRepeatingBlocks() {
  if (!repeatBlockDetectorEl) return;
  repeatBlockDetectorEl.replaceChildren();
  if (!currentRepeatingBlocks.length) {
    repeatBlockDetectorEl.appendChild(createEmptyInsight(t("repeatBlocksEmpty")));
    return;
  }

  currentRepeatingBlocks.forEach((group, groupIndex) => {
    const summary = document.createElement("div");
    summary.className = "repeat-summary";
    const prefixNote = group.prefixCount ? t("blkPrefixNote", { cols: group.prefixLabel }) : "";
    summary.textContent = `${group.meta || group.label}${prefixNote}`;
    repeatBlockDetectorEl.appendChild(summary);

    group.blocks.slice(0, 10).forEach((block, blockIndex) => {
      const item = document.createElement("div");
      item.className = "repeat-block-item";

      const top = document.createElement("div");
      top.className = "repeat-block-top";

      const title = document.createElement("div");
      title.className = "repeat-block-title";
      title.textContent = block.label;

      const badge = document.createElement("div");
      badge.className = "repeat-block-badge";
      badge.textContent = t("blkSpanBadge", { n: block.span });

      top.appendChild(title);
      top.appendChild(badge);

      const meta = document.createElement("div");
      meta.className = "repeat-block-meta";
      const headerPreview = block.headers.slice(0, 4).join(" • ");
      meta.textContent = `${t("blkColumnsMeta", { range: formatColRange(block.startAbs, block.endAbs) })}${headerPreview ? ` • ${headerPreview}` : ""}`;

      const actions = document.createElement("div");
      actions.className = "section-nav-actions";

      const btn = document.createElement("button");
      btn.className = "btn ghost btn-sm";
      btn.type = "button";
      btn.dataset.repeatGroupIndex = String(groupIndex);
      btn.dataset.repeatBlockIndex = String(blockIndex);
      btn.textContent = t("blkJumpTo");
      actions.appendChild(btn);

      item.appendChild(top);
      item.appendChild(meta);
      item.appendChild(actions);

      if (group.families && group.families.length) {
        const familyList = document.createElement("div");
        familyList.className = "repeat-family-list";
        group.families.slice(0, 6).forEach((family) => {
          const chip = document.createElement("div");
          chip.className = "repeat-family-chip";
          chip.textContent = `${family.label} ×${family.count}`;
          familyList.appendChild(chip);
        });
        item.appendChild(familyList);
      }

      repeatBlockDetectorEl.appendChild(item);
    });
  });
}

function focusRepeatingBlock(groupIndex, blockIndex) {
  const group = currentRepeatingBlocks[groupIndex];
  const block = group && group.blocks ? group.blocks[blockIndex] : null;
  if (!block || !tableWrapEl) return;
  const cells = theadEl.querySelectorAll(".guide-row .guide-cell");
  const cell = cells[block.startIndex];
  if (!cell) {
    toast(t("blockOutsideView"), "info");
    return;
  }
  const targetLeft = Math.max(0, cell.offsetLeft - 64);
  tableWrapEl.scrollTo({ left: targetLeft, behavior: "smooth" });
  syncHorizontalScrollbar();
}

function collectWorkbookStats(wb, fileName) {
  const book = wb && wb.Workbook ? wb.Workbook : {};
  const sheetsMeta = Array.isArray(book.Sheets) ? book.Sheets : [];
  let hiddenSheets = 0;
  let veryHiddenSheets = 0;
  sheetsMeta.forEach((sheetMeta) => {
    const hidden = Number(sheetMeta && sheetMeta.Hidden);
    if (hidden === 1) hiddenSheets += 1;
    if (hidden === 2) veryHiddenSheets += 1;
  });

  const definedNames = Array.isArray(book.Names) ? book.Names.length : 0;
  const ext = String(fileName || "").toLowerCase();
  const hasMacros = !!wb?.vbaraw || ext.endsWith(".xlsm");

  return {
    sheets: Array.isArray(wb?.SheetNames) ? wb.SheetNames.length : 0,
    hiddenSheets,
    veryHiddenSheets,
    definedNames,
    hasMacros,
  };
}

function collectSheetInsights() {
  const workbookItems = currentWorkbookStats ? [
    { label: t("inspSheets"), value: String(currentWorkbookStats.sheets) },
    { label: t("inspHiddenSheets"), value: String(currentWorkbookStats.hiddenSheets), tone: currentWorkbookStats.hiddenSheets ? "warning" : "" },
    { label: t("inspVeryHidden"), value: String(currentWorkbookStats.veryHiddenSheets), tone: currentWorkbookStats.veryHiddenSheets ? "warning" : "" },
    { label: t("inspDefinedNames"), value: String(currentWorkbookStats.definedNames), tone: currentWorkbookStats.definedNames ? "info" : "" },
  ] : [];

  if (!currentHeaders.length || !baseRows.length) {
    return {
      workbookRows: workbookItems,
      rows: [],
      flags: currentWorkbookStats?.hasMacros ? [{ label: t("inspMacroFile"), tone: "warning" }] : [],
    };
  }

  const totalRows = baseRows.length;
  const visibleRows = viewRows.length;
  const totalCols = currentHeaders.length;
  const duplicateHeaders = currentSheetStats?.duplicateHeaderCount || 0;
  const duplicateRows = (() => {
    const keys = baseRows.map((row) => JSON.stringify(row.values.map((value) => value instanceof Date ? value.toISOString() : value ?? "")));
    return keys.length - new Set(keys).size;
  })();

  let numericColumns = 0;
  let dateColumns = 0;
  let longTextColumns = 0;
  let sparseColumns = 0;

  currentHeaders.forEach((_, colIdx) => {
    let nonEmpty = 0;
    let numeric = 0;
    let dates = 0;
    let maxLen = 0;
    baseRows.forEach((row) => {
      const value = row.values[colIdx];
      if (value === null || value === undefined || String(value).trim() === "") return;
      nonEmpty += 1;
      if (typeof value === "number") numeric += 1;
      if (parseDateFlexible(value) instanceof Date) dates += 1;
      maxLen = Math.max(maxLen, String(getDisplayValue(row, colIdx)).length);
    });
    if (nonEmpty && numeric / nonEmpty >= 0.8) numericColumns += 1;
    if (nonEmpty && dates / nonEmpty >= 0.8) dateColumns += 1;
    if (maxLen > 150) longTextColumns += 1;
    if (totalRows && nonEmpty / totalRows <= 0.4) sparseColumns += 1;
  });

  const sheetItems = [
    { label: t("inspVisibleAllRows"), value: `${visibleRows} / ${totalRows}`, tone: visibleRows !== totalRows ? "info" : "" },
    { label: t("inspColumns"), value: String(totalCols) },
    { label: t("inspFormulas"), value: String(currentSheetStats?.formulaCount || 0), tone: (currentSheetStats?.formulaCount || 0) ? "info" : "" },
    {
      label: t("inspMerges"),
      value: `${currentSheetStats?.mergeRegions || 0} / ${currentSheetStats?.mergedCells || 0}`,
      tone: (currentSheetStats?.mergeRegions || 0) ? "info" : "",
    },
    { label: t("inspHiddenColsRows"), value: `${currentSheetStats?.hiddenColumns || 0} / ${currentSheetStats?.hiddenRows || 0}`, tone: ((currentSheetStats?.hiddenColumns || 0) || (currentSheetStats?.hiddenRows || 0)) ? "warning" : "" },
    { label: t("inspNumericDateCols"), value: `${numericColumns} / ${dateColumns}` },
    { label: t("inspSparseCols"), value: `${sparseColumns} (${formatPercent(sparseColumns, totalCols)})`, tone: sparseColumns ? "warning" : "" },
    { label: t("inspLongText"), value: String(longTextColumns), tone: longTextColumns ? "info" : "" },
  ];

  const flags = [];
  // Bazowy nagłówek (bez sufiksu cyklu, np. "od2" -> "od"), żeby sygnały procesu/SLA
  // łapały też powtarzalne bloki w trybie wide, nie tylko po Wide-to-Long.
  const normalizedHeaders = currentHeaders.map((header) => normalizeAnalysisKey(parseRepeatedHeader(header)?.base || header));
  const hasStatusHeader = normalizedHeaders.some((header) => /\b(status|stage|etap|stan|alert|priority|priorytet)\b/.test(header));
  const hasOwnerHeader = normalizedHeaders.some((header) => /\b(imie|nazwisko|osoba|pracownik|owner|assignee|agent|opiekun)\b/.test(header));
  const hasStartDateHeader = normalizedHeaders.some((header) => /\b(od|start|data od|from|poczatek|rozpoczecie|created)\b/.test(header));
  const hasEndDateHeader = normalizedHeaders.some((header) => /\b(do|end|data do|to|until|koniec|zakonczenie|closed)\b/.test(header));
  const processSignalCount = [hasStatusHeader, hasOwnerHeader, hasStartDateHeader && hasEndDateHeader].filter(Boolean).length;
  if (processSignalCount >= 2) flags.push({ label: t("flagProcessSheet"), tone: "info" });
  if (currentWorkbookStats?.hasMacros) flags.push({ label: t("inspMacroFile"), tone: "warning" });
  if (duplicateHeaders) flags.push({ label: t("flagDuplicateHeaders", { n: duplicateHeaders }), tone: "warning" });
  if (duplicateRows) flags.push({ label: t("flagDuplicateRows", { n: duplicateRows }), tone: duplicateRows > 0 ? "warning" : "" });
  if ((currentSheetStats?.formulaMissingResultCount || 0) > 0) {
    flags.push({ label: t("flagFormulasNoResult", { n: currentSheetStats.formulaMissingResultCount }), tone: "warning" });
  }
  if ((currentSheetStats?.commentCount || 0) > 0) flags.push({ label: t("flagComments", { n: currentSheetStats.commentCount }), tone: "info" });
  if ((currentSheetStats?.hyperlinkCount || 0) > 0) flags.push({ label: t("flagLinks", { n: currentSheetStats.hyperlinkCount }), tone: "info" });
  if (currentWorkbookStats?.veryHiddenSheets) flags.push({ label: t("flagVeryHiddenSheets"), tone: "warning" });

  return {
    workbookRows: workbookItems,
    rows: sheetItems,
    flags,
  };
}

function isKpiLabelCandidate(text) {
  const label = cleanSectionLabel(text);
  if (!label || label.length < 3 || label.length > 48) return false;
  if (/^\d+$/.test(label)) return false;
  const lowered = label.toLowerCase();
  const keywords = [
    "suma",
    "razem",
    "wartosc",
    "wartość",
    "koszt",
    "budzet",
    "budżet",
    "roznica",
    "różnica",
    "saldo",
    "marza",
    "marża",
    "przychod",
    "przychód",
    "wynik",
    "netto",
    "brutto",
    "plan",
    "wykonanie",
    "liczba",
    "ilosc",
    "ilość",
    "procent",
    "udzial",
    "udział",
    "kpi",
  ];
  return keywords.some((keyword) => lowered.includes(keyword)) || /[:\-]$/.test(label);
}

function isKpiValueCandidate(cell, displayText) {
  if (!cell) return false;
  const display = cleanSectionLabel(displayText);
  if (!display) return false;
  if (typeof cell.v === "number" && Number.isFinite(cell.v)) return true;
  if (cell.v instanceof Date) return true;
  if (/%|zl|zł|pln|eur|usd/i.test(display)) return true;
  if (/^-?\d[\d\s,.\u00A0]*$/.test(display)) return true;
  return false;
}

function normalizeKpiLabel(label) {
  const raw = cleanSectionLabel(label).toLowerCase();
  if (!raw) return "";
  return raw
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^a-z0-9]+/g, " ")
    .replace(/\b(na|do|od|i|oraz|jeszcze|samochod|samochodu|wartosc|wartosc)\b/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function inferKpiSemanticBucket(label) {
  const normalized = normalizeKpiLabel(label);
  if (!normalized) return "";
  if (/budzet/.test(normalized)) return "budget";
  if (/koszt|suma|razem|subtotal/.test(normalized)) return "total";
  if (/roznic|saldo|wynik/.test(normalized)) return "difference";
  if (/marz|procent|udzial|wykonanie/.test(normalized)) return "ratio";
  return normalized;
}

function scoreKpiLabelQuality(label) {
  const clean = cleanSectionLabel(label);
  const normalized = normalizeKpiLabel(label);
  if (!clean || !normalized) return 0;
  let score = normalized.length;
  if (normalized.includes(" ")) score += 4;
  if (clean.length >= 16) score += 4;
  if (/^(suma|razem|subtotal|wynik)$/i.test(normalized)) score -= 8;
  if (/koszt|budzet|roznic|saldo|marza|wykonanie|przychod|brutto|netto/i.test(normalized)) score += 6;
  return score;
}

function dedupeKpiEntries(entries) {
  const bestByKey = new Map();
  entries.forEach((entry) => {
    const semantic = inferKpiSemanticBucket(entry.label);
    const valueKey = cleanSectionLabel(entry.value);
    const key = `${semantic}|${valueKey}`;
    const existing = bestByKey.get(key);
    if (!existing) {
      bestByKey.set(key, {
        ...entry,
        aliases: [],
      });
      return;
    }
    if (entry.label !== existing.label && !existing.aliases.includes(entry.label)) {
      existing.aliases.push(entry.label);
    }
    const existingCompositeScore = existing.score + scoreKpiLabelQuality(existing.label);
    const entryCompositeScore = entry.score + scoreKpiLabelQuality(entry.label);
    if (entryCompositeScore > existingCompositeScore) {
      bestByKey.set(key, {
        ...entry,
        aliases: Array.from(new Set([existing.label, ...existing.aliases])),
      });
    }
  });
  return Array.from(bestByKey.values());
}

function pushKpiEntry(entries, seen, labelText, valueText, valueRef, labelRef, valueCell, scoreExtras = 0) {
  const cleanLabel = cleanSectionLabel(labelText).replace(/[:\-]\s*$/, "");
  const cleanValue = cleanSectionLabel(valueText);
  if (!cleanLabel || !cleanValue) return;
  const seenKey = `${normalizeKpiLabel(cleanLabel)}|${valueRef}`;
  if (seen.has(seenKey)) return;
  seen.add(seenKey);

  const decoded = XLSX.utils.decode_cell(valueRef);
  let score = scoreExtras;
  if (typeof valueCell?.v === "number") score += 3;
  if (valueCell?.f) score += 2;
  if (/%|zl|zł|pln|eur|usd/i.test(cleanValue)) score += 2;
  if (isKpiLabelCandidate(cleanLabel)) score += 2;

  entries.push({
    label: cleanLabel,
    value: cleanValue,
    address: valueRef,
    labelAddress: labelRef,
    rowIndex0: decoded.r,
    colAbs: decoded.c,
    score,
  });
}

function findDistantSameRowKpiTarget(sheet, rowIndex0, labelColAbs, endColAbs) {
  if (!sheet) return null;
  const farCandidates = [];
  const maxOffset = Math.min(6, endColAbs - labelColAbs);
  for (let offset = 3; offset <= maxOffset; offset++) {
    const candidateCol = labelColAbs + offset;
    let gapHasContent = false;
    for (let mid = labelColAbs + 1; mid < candidateCol; mid++) {
      if (getCellDisplayText(sheet, rowIndex0, mid)) {
        gapHasContent = true;
        break;
      }
    }
    if (gapHasContent) continue;
    const valueRef = XLSX.utils.encode_cell({ r: rowIndex0, c: candidateCol });
    const valueCell = sheet[valueRef];
    const valueDisplay = cleanSectionLabel(getCellDisplayText(sheet, rowIndex0, candidateCol));
    if (!isKpiValueCandidate(valueCell, valueDisplay)) continue;
    farCandidates.push({
      valueRef,
      valueCell,
      valueDisplay,
      colAbs: candidateCol,
    });
  }
  return farCandidates.length === 1 ? farCandidates[0] : null;
}

function scanKpiZone(sheet, entries, seen, rowStart, rowEnd, colStart, colEnd) {
  for (let r = rowStart; r <= rowEnd; r++) {
    for (let c = colStart; c <= colEnd; c++) {
      const labelText = getCellDisplayText(sheet, r, c);
      if (!isKpiLabelCandidate(labelText)) continue;

      const candidates = [
        { row: r, col: c + 1 },
        { row: r, col: c + 2 },
        { row: r + 1, col: c },
      ];

      candidates.forEach((candidate) => {
        if (candidate.row > rowEnd || candidate.col > colEnd) return;
        const valueRef = XLSX.utils.encode_cell({ r: candidate.row, c: candidate.col });
        const valueCell = sheet[valueRef];
        const valueDisplay = cleanSectionLabel(getCellDisplayText(sheet, candidate.row, candidate.col));
        if (!isKpiValueCandidate(valueCell, valueDisplay)) return;
        let score = 0;
        if (candidate.row === r) score += 2;
        if (candidate.col === c + 1) score += 1;
        pushKpiEntry(
          entries,
          seen,
          labelText,
          valueDisplay,
          valueRef,
          XLSX.utils.encode_cell({ r, c }),
          valueCell,
          score
        );
      });

      const distantTarget = findDistantSameRowKpiTarget(sheet, r, c, colEnd);
      if (distantTarget) {
        pushKpiEntry(
          entries,
          seen,
          labelText,
          distantTarget.valueDisplay,
          distantTarget.valueRef,
          XLSX.utils.encode_cell({ r, c }),
          distantTarget.valueCell,
          4
        );
      }
    }
  }

  for (let r = rowStart; r <= rowEnd; r++) {
    for (let c = colStart; c <= colEnd; c++) {
      const valueRef = XLSX.utils.encode_cell({ r, c });
      const valueCell = sheet[valueRef];
      const valueDisplay = cleanSectionLabel(getCellDisplayText(sheet, r, c));
      if (!isKpiValueCandidate(valueCell, valueDisplay)) continue;

      const leftLabel = c > colStart ? getCellDisplayText(sheet, r, c - 1) : "";
      const twoLeftLabel = c - 2 >= colStart ? getCellDisplayText(sheet, r, c - 2) : "";
      const aboveLabel = r > rowStart ? getCellDisplayText(sheet, r - 1, c) : "";
      const diagLabel = (r > rowStart && c > colStart) ? getCellDisplayText(sheet, r - 1, c - 1) : "";
      const candidates = [
        { label: leftLabel, ref: c > colStart ? XLSX.utils.encode_cell({ r, c: c - 1 }) : valueRef, bonus: 2 },
        { label: twoLeftLabel, ref: c - 2 >= colStart ? XLSX.utils.encode_cell({ r, c: c - 2 }) : valueRef, bonus: 1 },
        { label: aboveLabel, ref: r > rowStart ? XLSX.utils.encode_cell({ r: r - 1, c }) : valueRef, bonus: 1 },
        { label: diagLabel, ref: (r > rowStart && c > colStart) ? XLSX.utils.encode_cell({ r: r - 1, c: c - 1 }) : valueRef, bonus: 0 },
      ];

      candidates.forEach((candidate) => {
        if (!isKpiLabelCandidate(candidate.label)) return;
        pushKpiEntry(
          entries,
          seen,
          candidate.label,
          valueDisplay,
          valueRef,
          candidate.ref,
          valueCell,
          candidate.bonus
        );
      });
    }
  }
}

function collectKpiEntries(sheet, headerRow) {
  if (!sheet) return { entries: [], anchorRow: headerRow || 1 };
  const inferredAnchorRow = Math.max(1, detectHeaderRowSimple(sheet));
  const anchorRow = Math.max(1, headerRow || 1, inferredAnchorRow);
  if (anchorRow <= 1) return { entries: [], anchorRow };
  const entries = [];
  const seen = new Set();
  const effectiveRange = computeEffectiveSheetRange(sheet, anchorRow);
  const startRow = Math.max(effectiveRange.s.r, anchorRow - 8);
  const endRow = Math.max(startRow, anchorRow - 1);
  const endCol = effectiveRange.e.c;

  scanKpiZone(sheet, entries, seen, startRow, endRow, effectiveRange.s.c, endCol);

  const bottomStartRow = Math.max(anchorRow, effectiveRange.e.r - 4);
  if (bottomStartRow <= effectiveRange.e.r) {
    scanKpiZone(sheet, entries, seen, bottomStartRow, effectiveRange.e.r, effectiveRange.s.c, endCol);
  }

  return {
    anchorRow,
    entries: dedupeKpiEntries(entries).sort((a, b) => {
      if (b.score !== a.score) return b.score - a.score;
      return a.address.localeCompare(b.address, "pl");
    })
    .slice(0, 8),
  };
}

function renderKpiExtractor() {
  if (!kpiSummaryEl || !kpiListEl) return;
  kpiSummaryEl.replaceChildren();
  kpiListEl.replaceChildren();

  if (!currentHeaders.length || !currentKpiEntries.length) {
    renderInsightList(kpiSummaryEl, [], t("kpiSummaryEmpty"));
    kpiListEl.appendChild(createEmptyInsight(t("kpiListEmpty")));
    return;
  }

  renderInsightList(kpiSummaryEl, [
    { label: t("kpiCandidates"), value: String(currentKpiEntries.length), tone: "info" },
    {
      label: t("kpiSource"),
      value: currentKpiAnchorRow === currentHeaderRow
        ? t("kpiRowsAboveHeader", { row: currentHeaderRow })
        : t("kpiRowsAboveDetected", { row: currentKpiAnchorRow }),
      tone: currentKpiAnchorRow === currentHeaderRow ? "" : "info",
    },
  ], t("kpiNoSummary"));

  currentKpiEntries.forEach((entry) => {
    const item = document.createElement("div");
    item.className = "kpi-card";

    const label = document.createElement("div");
    label.className = "kpi-label";
    label.textContent = entry.label;

    const value = document.createElement("div");
    value.className = "kpi-value";
    value.textContent = entry.value;
    value.title = entry.aliases?.length
      ? t("kpiTitleAliases", { label: entry.label, value: entry.value, aliases: entry.aliases.join(", ") })
      : `${entry.label}: ${entry.value}`;

    const meta = document.createElement("div");
    meta.className = "kpi-meta";
    meta.textContent = entry.aliases?.length
      ? t("kpiMetaAliases", {
          address: entry.address,
          labelAddress: entry.labelAddress,
          aliases: entry.aliases.slice(0, 2).join(", "),
          more: entry.aliases.length > 2 ? ` +${entry.aliases.length - 2}` : "",
        })
      : t("kpiMetaPlain", { address: entry.address, labelAddress: entry.labelAddress });

    const actions = document.createElement("div");
    actions.className = "section-nav-actions";

    const btn = document.createElement("button");
    btn.className = "btn ghost btn-sm";
    btn.type = "button";
    btn.dataset.kpiAddress = entry.address;
    btn.textContent = t("kpiShowSource");

    actions.appendChild(btn);
    item.appendChild(label);
    item.appendChild(value);
    item.appendChild(meta);
    item.appendChild(actions);
    kpiListEl.appendChild(item);
  });
}

function focusKpiEntry(address) {
  const entry = currentKpiEntries.find((item) => item.address === address);
  if (!entry) return;
  const rowEl = tbodyEl.querySelector(`tr[data-row-index="${entry.rowIndex0}"]`);
  if (rowEl) {
    rowEl.scrollIntoView({ behavior: "smooth", block: "center", inline: "nearest" });
  } else {
    toast(t("kpiAboveTable", { row: entry.rowIndex0 + 1 }), "info");
    if (tableWrapEl) {
      tableWrapEl.scrollTo({ top: 0, behavior: "smooth" });
    }
  }
  const relativeColIdx = entry.colAbs - currentStartCol;
  if (Number.isFinite(relativeColIdx) && relativeColIdx >= 0) {
    focusColumnProfile(relativeColIdx);
  }
}

function inferColumnProfileType(stats) {
  if (!stats || !stats.nonEmpty) return "pusta";
  const ratio = (count) => (stats.nonEmpty ? count / stats.nonEmpty : 0);
  const numberRatio = ratio(stats.numericCount);
  const dateRatio = ratio(stats.dateCount);
  const formulaRatio = ratio(stats.formulaCount);
  const textRatio = ratio(stats.textCount);

  if (formulaRatio >= 0.8) return "formuly";
  if (dateRatio >= 0.8) return "daty";
  if (numberRatio >= 0.8) return "liczby";
  if (textRatio >= 0.8) return "tekst";
  return "mixed";
}

function formatColumnProfileRange(profile) {
  if (!profile) return "";
  if (profile.type === "liczby" && Number.isFinite(profile.minValue) && Number.isFinite(profile.maxValue)) {
    return `${profile.minValue} -> ${profile.maxValue}`;
  }
  if (profile.type === "daty" && profile.minDate instanceof Date && profile.maxDate instanceof Date) {
    return `${toDisplay(profile.minDate)} -> ${toDisplay(profile.maxDate)}`;
  }
  return "";
}

function collectColumnProfiles() {
  if (!currentHeaders.length || !baseRows.length) return [];
  const totalRows = baseRows.length;
  const profiles = currentHeaders.map((header, colIdx) => {
    const stats = {
      nonEmpty: 0,
      numericCount: 0,
      dateCount: 0,
      textCount: 0,
      formulaCount: 0,
      longTextCount: 0,
      minValue: null,
      maxValue: null,
      minDate: null,
      maxDate: null,
      unique: new Map(),
    };

    baseRows.forEach((row) => {
      const value = row.values[colIdx];
      const displayValue = getDisplayValue(row, colIdx);
      const text = String(displayValue ?? "").trim();
      if (text === "") return;

      stats.nonEmpty += 1;
      stats.unique.set(text, (stats.unique.get(text) || 0) + 1);
      if (text.length > 60) stats.longTextCount += 1;

      if (typeof value === "string" && value.startsWith("=")) stats.formulaCount += 1;
      if (typeof value === "number") {
        stats.numericCount += 1;
        stats.minValue = stats.minValue == null ? value : Math.min(stats.minValue, value);
        stats.maxValue = stats.maxValue == null ? value : Math.max(stats.maxValue, value);
      }

      const asDate = parseDateFlexible(value);
      if (asDate instanceof Date) {
        stats.dateCount += 1;
        stats.minDate = !stats.minDate || asDate < stats.minDate ? asDate : stats.minDate;
        stats.maxDate = !stats.maxDate || asDate > stats.maxDate ? asDate : stats.maxDate;
      }

      if (typeof value === "string" && !(parseDateFlexible(value) instanceof Date) && !value.startsWith("=")) {
        stats.textCount += 1;
      } else if (!(value instanceof Date) && typeof value !== "number" && typeof value !== "string") {
        stats.textCount += 1;
      }
    });

    const emptyCount = totalRows - stats.nonEmpty;
    const uniqueCount = stats.unique.size;
    const topValues = Array.from(stats.unique.entries())
      .sort((a, b) => b[1] - a[1])
      .slice(0, 3)
      .map(([label, count]) => ({ label, count }));
    const type = inferColumnProfileType(stats);
    const flags = [];
    let score = 0;

    if (stats.nonEmpty && stats.nonEmpty / totalRows <= 0.4) {
      flags.push("sparse");
      score += 2;
    }
    if (type === "mixed") {
      flags.push("mixed");
      score += 3;
    }
    if (stats.longTextCount > 0) {
      flags.push("longText");
      score += 1;
    }
    if (uniqueCount > Math.max(20, totalRows * 0.9) && type === "tekst") {
      flags.push("mostlyUnique");
      score += 1;
    }
    if (stats.formulaCount > 0 && stats.formulaCount / stats.nonEmpty >= 0.8) {
      flags.push("formulaColumn");
      score += 1;
    }
    if (emptyCount === totalRows) {
      flags.push("empty");
      score += 4;
    }

    return {
      header,
      colIdx,
      colAbs: currentStartCol + colIdx,
      nonEmpty: stats.nonEmpty,
      emptyCount,
      emptyPct: totalRows ? Math.round((emptyCount / totalRows) * 100) : 0,
      uniqueCount,
      type,
      topValues,
      rangeLabel: formatColumnProfileRange({
        type,
        minValue: stats.minValue,
        maxValue: stats.maxValue,
        minDate: stats.minDate,
        maxDate: stats.maxDate,
      }),
      flags,
      score,
    };
  });

  return profiles.sort((a, b) => {
    if (b.score !== a.score) return b.score - a.score;
    if (a.emptyPct !== b.emptyPct) return b.emptyPct - a.emptyPct;
    return a.header.localeCompare(b.header, "pl");
  });
}

function renderColumnProfiles() {
  if (!columnProfilerEl) return;
  columnProfilerEl.replaceChildren();
  if (!currentColumnProfiles.length) {
    columnProfilerEl.appendChild(createEmptyInsight(t("columnProfilesEmpty")));
    return;
  }

  currentColumnProfiles.slice(0, 14).forEach((profile, index) => {
    const item = document.createElement("div");
    item.className = "column-profile-item";

    const top = document.createElement("div");
    top.className = "column-profile-top";

    const title = document.createElement("div");
    title.className = "column-profile-title";
    title.textContent = profile.header;

    const kind = document.createElement("div");
    kind.className = "column-profile-kind";
    kind.textContent = t("profType_" + profile.type);

    top.appendChild(title);
    top.appendChild(kind);

    const meta = document.createElement("div");
    meta.className = "column-profile-meta";
    meta.textContent = t("profMeta", { col: XLSX.utils.encode_col(profile.colAbs), empty: profile.emptyPct, unique: profile.uniqueCount });

    const stats = document.createElement("div");
    stats.className = "column-profile-stats";
    if (profile.rangeLabel) {
      const rangeChip = document.createElement("div");
      rangeChip.className = "column-profile-chip";
      rangeChip.textContent = profile.rangeLabel;
      stats.appendChild(rangeChip);
    }
    profile.topValues.forEach((entry) => {
      const chip = document.createElement("div");
      chip.className = "column-profile-chip";
      chip.textContent = `${entry.label.slice(0, 24)}${entry.label.length > 24 ? "..." : ""} ×${entry.count}`;
      stats.appendChild(chip);
    });

    if (profile.flags.length) {
      const flags = document.createElement("div");
      flags.className = "column-profile-flags";
      profile.flags.forEach((flag) => {
        const badge = document.createElement("div");
        badge.className = "column-profile-flag";
        badge.textContent = t("profFlag_" + flag);
        flags.appendChild(badge);
      });
      item.appendChild(top);
      item.appendChild(meta);
      if (stats.childNodes.length) item.appendChild(stats);
      item.appendChild(flags);
    } else {
      item.appendChild(top);
      item.appendChild(meta);
      if (stats.childNodes.length) item.appendChild(stats);
    }

    const actions = document.createElement("div");
    actions.className = "section-nav-actions";
    const btn = document.createElement("button");
    btn.className = "btn ghost btn-sm";
    btn.type = "button";
    btn.dataset.profileColIndex = String(profile.colIdx);
    btn.textContent = t("profJumpToCol");
    actions.appendChild(btn);
    item.appendChild(actions);

    columnProfilerEl.appendChild(item);
  });
}

function focusColumnProfile(colIdx) {
  if (!Number.isFinite(colIdx)) return;
  const cells = theadEl.querySelectorAll(".guide-row .guide-cell");
  const cell = cells[colIdx];
  if (cell && tableWrapEl) {
    const targetLeft = Math.max(0, cell.offsetLeft - 64);
    tableWrapEl.scrollTo({ left: targetLeft, behavior: "smooth" });
    syncHorizontalScrollbar();
    return;
  }
  toast(t("columnOutsideView"), "info");
}

function renderInsights() {
  const data = collectSheetInsights();
  renderInsightList(
    workbookInsightsEl,
    data.workbookRows || [],
    t("workbookInsightsEmpty")
  );
  renderInsightList(
    sheetInsightsEl,
    data.rows || [],
    t("sheetInsightsEmpty")
  );
  renderInsightFlags(data.flags || []);
}

/* ===================== Podsumowanie miesięczne =====================
   Generyczny rozkład danych na 12 ostatnich miesięcy wg dowolnej kolumny z datą.
   Działa na aktualnym widoku (po filtrach), więc np. po przefiltrowaniu po osobie
   pokazuje jej miesięczną częstotliwość. Miary: liczba wierszy / suma / średnia
   wybranej kolumny liczbowej. Pomyślane elastycznie — dla różnych arkuszy/przypadków. */
// dateCols: indeksy kolumn dat do grupowania po miesiącu (multi),
// measureCols: indeksy kolumn miary (multi — parowane pozycyjnie z dateCols per cykl),
// metric: occurrences|rows|sum|avg|min|max, months: okno, anchor: 'data'|'today', split: rozbicie.
let monthlySummaryState = { dateCols: null, metric: "occurrences", measureCols: null, months: 12, anchor: "data", split: true, gap: true };

const MONTH_ABBR = {
  pl: ["sty", "lut", "mar", "kwi", "maj", "cze", "lip", "sie", "wrz", "paź", "lis", "gru"],
  en: ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"],
};
// Paleta segmentów słupków (rozbicie per kolumna) — wyraziste, ale w klimacie palety.
const MONTHLY_SEG_COLORS = ["#2f6f5c", "#b7791f", "#3b6fb0", "#9b59b6", "#cf6679", "#3aa394", "#d98324", "#6b8e23"];
function monthKeyLabel(key) {
  const m = /^(\d{4})-(\d{2})$/.exec(key);
  if (!m) return key;
  const lang = (typeof currentLang !== "undefined" && MONTH_ABBR[currentLang]) ? currentLang : "pl";
  return `${MONTH_ABBR[lang][parseInt(m[2], 10) - 1]} ${m[1]}`;
}
function lastNMonthKeys(endKey, n) {
  const m = /^(\d{4})-(\d{2})$/.exec(endKey);
  if (!m) return [];
  let y = parseInt(m[1], 10), mo = parseInt(m[2], 10) - 1;
  const keys = [];
  for (let i = 0; i < n; i++) {
    keys.unshift(`${y}-${String(mo + 1).padStart(2, "0")}`);
    mo--; if (mo < 0) { mo = 11; y--; }
  }
  return keys;
}
function formatMonthlyDate(ms) {
  const d = new Date(ms);
  if (isNaN(d)) return "";
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
}
// Dni → „Xm Yd" (miesiąc = 30 dni, symetrycznie do parseDurationDaysFlexible).
function formatMonthlyDuration(days) {
  if (!isFinite(days)) return "0";
  const total = Math.max(0, Math.round(days));
  const m = Math.floor(total / 30);
  const d = total - m * 30;
  if (m && d) return `${m}m ${d}d`;
  if (m) return `${m}m`;
  return `${d}d`;
}
// Generyczne formatowanie wartości miary wg TYPU kolumny miary:
//  - 'date'     → śr./min/maks jako data (YYYY-MM-DD),
//  - 'duration' → suma/śr./min/maks jako „Xm Yd" (jak kolumna „Długość"),
//  - 'number'   → liczba; liczniki (occurrences/rows) → liczba całkowita.
function formatMonthlyValue(val, metric, measureType) {
  if (metric === "occurrences" || metric === "rows") return String(Math.round(val));
  if (measureType === "date" && (metric === "avg" || metric === "min" || metric === "max")) return formatMonthlyDate(val);
  if (measureType === "duration") return formatMonthlyDuration(val);
  const rounded = Math.round(val * 100) / 100;
  return (typeof formatStatNumber === "function") ? formatStatNumber(rounded) : String(rounded);
}

function monthlyHeaderBase(header) {
  const rep = (typeof parseRepeatedHeader === "function") ? parseRepeatedHeader(header) : null;
  const base = (rep && rep.base) || (typeof cleanSectionLabel === "function" ? cleanSectionLabel(header) : header) || header;
  return normalizeAnalysisKey(base);
}

function renderMonthlySummary() {
  if (!monthlySummaryEl) return;
  monthlySummaryEl.replaceChildren();
  const model = (typeof currentDisplayModel !== "undefined" && currentDisplayModel) ? currentDisplayModel : getDisplayModel();
  if (!model || !Array.isArray(model.rows) || !model.rows.length) {
    monthlySummaryEl.appendChild(createEmptyInsight(t("monthlyNoData")));
    return;
  }
  const profiles = collectAggregationProfiles(model);
  const dateCols = profiles.filter((p) => p.nonEmptyCount >= 2 && p.dateCount / p.nonEmptyCount >= 0.6);
  if (!dateCols.length) {
    monthlySummaryEl.appendChild(createEmptyInsight(t("monthlyNoDate")));
    return;
  }
  const validIdx = new Set(dateCols.map((p) => p.idx));
  const bestCol = dateCols.slice().sort((a, b) => b.dateCount - a.dateCount)[0];

  // Wybór kolumn dat (multi). Domyślnie SMART: wszystkie kolumny tego samego pola
  // (np. od/od2/od3) → łączny „obrót". Stan zapamiętany, jeśli wciąż poprawny.
  let selectedCols = Array.isArray(monthlySummaryState.dateCols)
    ? monthlySummaryState.dateCols.filter((i) => validIdx.has(i))
    : [];
  if (!selectedCols.length) {
    const base = monthlyHeaderBase(bestCol.header);
    selectedCols = dateCols.filter((p) => monthlyHeaderBase(p.header) === base).map((p) => p.idx);
    if (!selectedCols.length) selectedCols = [bestCol.idx];
  }
  selectedCols.sort((a, b) => a - b); // rosnąco po indeksie → parowanie pozycyjne z miarami (cykl k)
  const selectedSet = new Set(selectedCols);
  monthlySummaryState.dateCols = selectedCols.slice(); // zapamiętaj efektywny wybór (do chipów)

  // Kandydaci na miarę: kolumny LICZBOWE/DURACJA oraz DATOWE. Daty można wskazać także te,
  // po których grupujemy — bo z 2+ kolumn z datą liczymy ODSTĘP (np. od→do = czas trzymania).
  const allDateIdx = new Set(dateCols.map((p) => p.idx));
  const measureCandidates = profiles.filter((p) => p.nonEmptyCount >= 2 && (
    (p.numericCount + p.durationCount) / p.nonEmptyCount >= 0.6 || allDateIdx.has(p.idx)
  ));
  const aggMetric = (m) => m === "sum" || m === "avg" || m === "min" || m === "max";
  const levelMetric = (m) => m === "avg" || m === "min" || m === "max";
  const countMetric = (m) => m === "occurrences" || m === "rows";
  const typeOfMeasure = (idx) => {
    if (idx == null) return null;
    if (allDateIdx.has(idx)) return "date";
    const pr = profiles.find((p) => p.idx === idx);
    if (pr && pr.durationCount / pr.nonEmptyCount >= 0.5 && pr.durationCount >= pr.numericCount) return "duration";
    return "number";
  };

  // Miara = WIELE kolumn (chipy), parowana POZYCYJNIE z kolumnami dat (cykl k: data_k ↔ miara_k).
  // Dzięki temu „średni czas trzymania" liczy każdy Długość_k w miesiącu swojego od_k — uniwersalnie
  // dla wielu cykli. Smart domyślny: wszystkie kolumny tego samego pola co pierwsza (np. Długość, Długość2…).
  let selectedMeasures = [];
  if (measureCandidates.length) {
    selectedMeasures = Array.isArray(monthlySummaryState.measureCols)
      ? monthlySummaryState.measureCols.filter((i) => measureCandidates.some((p) => p.idx === i))
      : [];
    if (!selectedMeasures.length) {
      // Sensowny domyślny: preferuj DURACJĘ (np. „Długość"), potem liczbę nie-ID, na końcu cokolwiek.
      // Unikamy kolumn ID/„Nr." — średnia z nich jest bez sensu.
      const nonDate = measureCandidates.filter((p) => !allDateIdx.has(p.idx));
      const isId = (p) => (typeof classifyAggregationHeader === "function") && classifyAggregationHeader(p.header) === "id";
      const pref = nonDate.find((p) => typeOfMeasure(p.idx) === "duration")
        || nonDate.find((p) => !isId(p))
        || nonDate[0] || measureCandidates[0];
      const base = monthlyHeaderBase(pref.header);
      selectedMeasures = measureCandidates.filter((p) => monthlyHeaderBase(p.header) === base).map((p) => p.idx);
      if (!selectedMeasures.length) selectedMeasures = [pref.idx];
    }
  }
  selectedMeasures.sort((a, b) => a - b);
  monthlySummaryState.measureCols = selectedMeasures.slice(); // zapamiętaj (do chipów)
  // Typ miary wg pierwszej wybranej: 'date' → wynik jako data; 'duration' → „Xm Yd"; 'number' → liczba.
  const measureType = selectedMeasures.length ? typeOfMeasure(selectedMeasures[0]) : null;
  // ODSTĘP dat: gdy wybrane 2+ kolumny z DATĄ, MOŻNA policzyć span między nimi (czas między datami).
  // Sterowane checkboxem (domyślnie włączony); wyłączony → daty traktujemy jak zwykłą miarę (avg/min/maks dat).
  const canGap = measureType === "date" && selectedMeasures.length >= 2;
  const measureIsGap = canGap && monthlySummaryState.gap !== false;
  const effectiveType = measureIsGap ? "duration" : measureType; // typ do formatowania/skalowania wyniku

  // Dozwolone miary: zawsze liczniki; przy mierze — śr./min/maks; suma dla liczb/duracji/odstępu (nie dla pojedynczej daty).
  const allowedMetrics = ["occurrences", "rows"];
  if (measureCandidates.length) {
    if (effectiveType !== "date") allowedMetrics.push("sum");
    allowedMetrics.push("avg", "min", "max");
  }
  let metric = monthlySummaryState.metric || "occurrences";
  if (metric === "count") metric = "occurrences"; // migracja starej nazwy
  if (!allowedMetrics.includes(metric)) metric = "occurrences";

  const monthsCount = [3, 6, 12, 24, 36].includes(monthlySummaryState.months) ? monthlySummaryState.months : 12;
  const anchor = monthlySummaryState.anchor === "today" ? "today" : "data";
  const canSplit = selectedCols.length > 1;
  const split = canSplit && monthlySummaryState.split !== false;
  const rangeScale = effectiveType === "date" && levelMetric(metric); // daty: skaluj słupki do zakresu, nie od zera

  // Wartość miary z DANEJ kolumny (typ wg measureType). Data → ms; duracja → dni; liczba → liczba.
  const measureValueOfCol = (row, colIdx) => {
    if (colIdx == null) return null;
    const raw = (row.values ? row.values[colIdx] : null) ?? getDisplayValue(row, colIdx);
    if (measureType === "date") { const d = parseDateFlexible(raw); return (d instanceof Date && !isNaN(d)) ? d.getTime() : null; }
    if (measureType === "duration") return parseDurationDaysFlexible(raw);
    return parseCellNumber(row.values ? row.values[colIdx] : null, getDisplayValue(row, colIdx));
  };
  // Odstęp dat: span (max − min) wybranych kolumn z datą w wierszu, w dniach (→ duracja „Xm Yd").
  const DAY_MS = 86400000;
  const gapValueOf = (row) => {
    const ms = [];
    selectedMeasures.forEach((ci) => {
      const d = parseDateFlexible((row.values ? row.values[ci] : null) ?? getDisplayValue(row, ci));
      if (d instanceof Date && !isNaN(d)) ms.push(d.getTime());
    });
    if (ms.length < 2) return null;
    return (Math.max(...ms) - Math.min(...ms)) / DAY_MS;
  };
  // Parowanie pozycyjne: cykl k (k-ta wybrana data) ↔ k-ta wybrana miara. Jedna miara → dla
  // wszystkich cykli; za mało miar → ostatnia powtórzona; brak miary/licznik → null.
  const measureForPosition = (k) => {
    if (!aggMetric(metric) || !selectedMeasures.length) return null;
    if (selectedMeasures.length === 1) return selectedMeasures[0];
    return selectedMeasures[k] != null ? selectedMeasures[k] : selectedMeasures[selectedMeasures.length - 1];
  };
  const newAgg = () => ({ occ: 0, rows: new Set(), sum: 0, n: 0, min: Infinity, max: -Infinity });
  const addToAgg = (a, rowId, mv) => {
    a.occ += 1; a.rows.add(rowId);
    if (aggMetric(metric) && mv != null) { a.sum += mv; a.n += 1; if (mv < a.min) a.min = mv; if (mv > a.max) a.max = mv; }
  };
  const aggValue = (a) => {
    if (!a) return 0;
    if (metric === "rows") return a.rows.size;
    if (metric === "sum") return a.sum;
    if (metric === "avg") return a.n ? a.sum / a.n : 0;
    if (metric === "min") return a.n ? a.min : 0;
    if (metric === "max") return a.n ? a.max : 0;
    return a.occ;
  };
  const aggHas = (a) => (a ? (countMetric(metric) ? a.occ > 0 : a.n > 0) : false);

  // Kubełkowanie po roku-miesiącu; każda wybrana kolumna daty liczona osobno + per-kolumna.
  const buckets = new Map();
  let maxKey = null;
  model.rows.forEach((row, ri) => {
    if (row && row.isSubheader) return;
    const rowId = (typeof row.rowIndex0 === "number") ? row.rowIndex0 : ri;
    const gap = measureIsGap ? gapValueOf(row) : null; // odstęp liczony raz na wiersz
    selectedCols.forEach((ci, k) => {
      const d = parseDateFlexible((row.values ? row.values[ci] : null) ?? getDisplayValue(row, ci));
      if (!(d instanceof Date) || isNaN(d)) return;
      const mv = measureIsGap ? gap : measureValueOfCol(row, measureForPosition(k)); // odstęp lub miara cyklu
      const key = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`;
      let b = buckets.get(key);
      if (!b) { b = newAgg(); b.byCol = {}; buckets.set(key, b); }
      addToAgg(b, rowId, mv);
      let bc = b.byCol[ci];
      if (!bc) { bc = newAgg(); b.byCol[ci] = bc; }
      addToAgg(bc, rowId, mv);
      if (maxKey == null || key > maxKey) maxKey = key;
    });
  });
  if (!maxKey) { monthlySummaryEl.appendChild(createEmptyInsight(t("monthlyNoDate"))); return; }

  const todayKey = (() => { const d = new Date(); return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`; })();
  const endKey = anchor === "today" ? todayKey : maxKey;
  const months = lastNMonthKeys(endKey, monthsCount);
  const rowsData = months.map((key) => {
    const b = buckets.get(key);
    return { key, agg: b, val: aggValue(b), has: aggHas(b), occ: b ? b.occ : 0, byCol: b ? b.byCol : {} };
  });
  const totalCount = rowsData.reduce((s, r) => s + r.occ, 0);
  const segColorFor = (pos) => MONTHLY_SEG_COLORS[pos % MONTHLY_SEG_COLORS.length];

  // Skalowanie długości słupka. Zero-based dla liczników/sum/liczb; range-based dla dat
  // (śr./min/maks daty to „poziom", nie ilość — słupek od najwcześniejszej do najpóźniejszej).
  const presentVals = rowsData.filter((r) => r.has).map((r) => r.val);
  const minVal = presentVals.length ? Math.min(...presentVals) : 0;
  const maxVal = Math.max(1, ...presentVals.map((v) => Math.abs(v)));
  const scaleOverall = (val, has) => {
    if (!has) return 0;
    if (rangeScale) return maxVal > minVal ? Math.round(10 + ((val - minVal) / (maxVal - minVal)) * 90) : 60;
    return Math.max(val > 0 ? 3 : 0, Math.round((Math.abs(val) / maxVal) * 100));
  };

  // ── kontrolki ──
  const mkSelect = (ctrl, value, label, options) => {
    const lab = document.createElement("label");
    lab.className = "field";
    lab.textContent = label;
    const sel = document.createElement("select");
    sel.dataset.monthlyControl = ctrl;
    options.forEach((o) => {
      const op = document.createElement("option");
      op.value = String(o.value);
      op.textContent = o.text;
      if (String(o.value) === String(value)) op.selected = true;
      sel.appendChild(op);
    });
    lab.appendChild(sel);
    return lab;
  };

  const controls = document.createElement("div");
  controls.className = "monthly-controls";

  // Wspólny builder chipów (multi-select) — dla dat i dla miar.
  const mkChips = (attr, candidates, selSet, withTypeMarker) => {
    const wrap = document.createElement("div");
    wrap.className = "monthly-chips";
    candidates.forEach((p) => {
      const on = selSet.has(p.idx);
      const chip = document.createElement("button");
      chip.type = "button";
      chip.className = "monthly-chip" + (on ? " active" : "");
      chip.dataset[attr] = String(p.idx);
      chip.setAttribute("aria-pressed", on ? "true" : "false");
      let label = p.header;
      if (withTypeMarker) {
        const isDate = allDateIdx.has(p.idx);
        const isDur = !isDate && p.durationCount / p.nonEmptyCount >= 0.5 && p.durationCount >= p.numericCount;
        label += isDate ? " 📅" : isDur ? " ⏱" : "";
      }
      chip.textContent = label;
      wrap.appendChild(chip);
    });
    return wrap;
  };
  const fieldWrap = (labelText, cls, child, hint) => {
    const w = document.createElement("div");
    w.className = "field" + (cls ? " " + cls : "");
    w.textContent = labelText;
    if (hint) { w.setAttribute("data-hint", hint); w.setAttribute("data-hint-touch", "on"); w.setAttribute("data-hint-tap", ""); }
    w.appendChild(child);
    return w;
  };

  const metricLabels = {
    occurrences: t("monthlyMetricOccurrences"),
    rows: t("monthlyMetricRows"),
    sum: t("monthlyMetricSum"),
    avg: t("monthlyMetricAvg"),
    min: t("monthlyMetricMin"),
    max: t("monthlyMetricMax"),
  };
  // Kolejność jak naturalne zdanie: 1) Co policzyć → 2) z których kolumn → 3) wg jakiej daty (miesiąc).
  controls.appendChild(mkSelect("metric", metric, t("monthlyMetricQ"),
    allowedMetrics.map((m) => ({ value: m, text: metricLabels[m] }))));
  if (aggMetric(metric) && measureCandidates.length) {
    controls.appendChild(fieldWrap(t("monthlyMeasureOf"), "monthly-measure-field",
      mkChips("monthlyMeasurecol", measureCandidates, new Set(selectedMeasures), true), t("monthlyMeasureHint")));
    // Wykryto 2+ kolumny z datą → checkbox włączający liczenie ODSTĘPU między nimi.
    if (canGap) {
      const gapLabel = document.createElement("label");
      gapLabel.className = "field checkbox monthly-gap-field";
      const gcb = document.createElement("input");
      gcb.type = "checkbox";
      gcb.dataset.monthlyControl = "gap";
      gcb.checked = measureIsGap;
      gapLabel.appendChild(gcb);
      gapLabel.appendChild(document.createTextNode(" " + t("monthlyGapToggle")));
      controls.appendChild(gapLabel);
    }
  }
  controls.appendChild(fieldWrap(t("monthlyDateGroup"), "monthly-cols-field",
    mkChips("monthlyDatecol", dateCols, selectedSet, false)));
  controls.appendChild(mkSelect("months", monthsCount, t("monthlyWindow"),
    [3, 6, 12, 24, 36].map((n) => ({ value: n, text: t("monthlyMonthsOption", { n }) }))));
  controls.appendChild(mkSelect("anchor", anchor, t("monthlyAnchor"), [
    { value: "data", text: t("monthlyAnchorData") },
    { value: "today", text: t("monthlyAnchorToday") },
  ]));
  if (canSplit) {
    const splitLabel = document.createElement("label");
    splitLabel.className = "field checkbox";
    const cb = document.createElement("input");
    cb.type = "checkbox";
    cb.dataset.monthlyControl = "split";
    cb.checked = split;
    splitLabel.appendChild(cb);
    splitLabel.appendChild(document.createTextNode(" " + t("monthlySplit")));
    controls.appendChild(splitLabel);
  }
  monthlySummaryEl.appendChild(controls);

  // ── Podsumowanie zapytania prostym językiem (readback) — czytelny tytuł nad wynikami ──
  const readWhat = {
    occurrences: t("monthlyReadOccurrences"),
    rows: t("monthlyReadRows"),
    sum: t("monthlyMetricSum"),
    avg: t("monthlyMetricAvg"),
    min: t("monthlyMetricMin"),
    max: t("monthlyMetricMax"),
  }[metric] || metricLabels[metric];
  let whatPhrase = readWhat;
  if (aggMetric(metric) && selectedMeasures.length) {
    const names = selectedMeasures.map((i) => { const p = measureCandidates.find((x) => x.idx === i); return p ? p.header : String(i); }).join(", ");
    whatPhrase += " " + t(measureIsGap ? "monthlyReadGap" : "monthlyReadOf", { col: names });
  }
  const dateNames = selectedCols.map((ci) => { const p = dateCols.find((x) => x.idx === ci); return p ? p.header : String(ci); }).join(", ");
  const query = document.createElement("div");
  query.className = "monthly-query";
  query.textContent = t("monthlyReadback", { what: whatPhrase, dates: dateNames });
  monthlySummaryEl.appendChild(query);

  const note = document.createElement("div");
  note.className = "monthly-note";
  note.textContent = t("monthlyTotalRows", { count: totalCount });
  monthlySummaryEl.appendChild(note);

  // ── legenda (gdy rozbicie na kolumny) ──
  if (split) {
    const legend = document.createElement("div");
    legend.className = "monthly-legend";
    selectedCols.forEach((ci, pos) => {
      const profile = dateCols.find((p) => p.idx === ci);
      const li = document.createElement("span");
      li.className = "monthly-legend-item";
      const sw = document.createElement("span");
      sw.className = "monthly-swatch";
      sw.style.background = segColorFor(pos);
      li.appendChild(sw);
      li.appendChild(document.createTextNode(profile ? profile.header : String(ci)));
      legend.appendChild(li);
    });
    monthlySummaryEl.appendChild(legend);
  }

  // ── paski miesięczne ──
  // Zawsze JEDEN słupek; przy rozbiciu kolorowe segmenty per kolumna na tym samym słupku.
  // Dla miar addytywnych (wystąpienia/suma) szerokość segmentu = dokładny udział wartości.
  // Dla nie-addytywnych (wiersze/śr./min/maks) segmenty pokazują udział kolumn (kompozycja),
  // a DOKŁADNE wartości per kolumna są w tooltipie (cursor-hint) na wyniku przy słupku.
  const additive = metric === "occurrences" || metric === "sum";
  const list = document.createElement("div");
  list.className = "monthly-list";
  rowsData.forEach((r) => {
    const item = document.createElement("div");
    item.className = "monthly-row";
    const lab = document.createElement("div");
    lab.className = "monthly-label";
    lab.textContent = monthKeyLabel(r.key);
    const track = document.createElement("div");
    track.className = "monthly-track";
    const bar = document.createElement("div");
    bar.className = "monthly-bar";
    bar.style.width = `${scaleOverall(r.val, r.has)}%`;

    // dane per kolumna dla tego miesiąca
    const colData = selectedCols.map((ci, pos) => {
      const bc = r.byCol[ci];
      const ph = dateCols.find((p) => p.idx === ci);
      return { ci, pos, has: aggHas(bc), v: aggValue(bc), header: ph ? ph.header : String(ci) };
    });

    if (split && r.has) {
      bar.classList.add("split");
      const base = colData.reduce((s, c) => s + (c.has ? Math.max(0, c.v) : 0), 0) || 1;
      colData.forEach((c) => {
        if (!c.has || c.v <= 0) return;
        const seg = document.createElement("div");
        seg.className = "monthly-seg";
        seg.style.width = `${(Math.max(0, c.v) / base) * 100}%`;
        seg.style.background = segColorFor(c.pos);
        bar.appendChild(seg);
      });
    }
    track.appendChild(bar);

    const val = document.createElement("div");
    val.className = "monthly-value";
    val.textContent = r.has ? formatMonthlyValue(r.val, metric, effectiveType) : "0";
    // Tooltip (cursor-hint): dokładny wynik per kolumna — także dla miar nie-addytywnych.
    if (selectedCols.length > 1 && r.has) {
      const parts = colData.filter((c) => c.has).map((c) => `${c.header}: ${formatMonthlyValue(c.v, metric, effectiveType)}`);
      if (parts.length) {
        const hint = parts.join("  ·  ");
        val.classList.add("monthly-value-hint");
        val.setAttribute("data-hint", "");
        val.setAttribute("data-hint-pl", hint);
        val.setAttribute("data-hint-en", hint);
        val.setAttribute("data-hint-touch", "on");
        val.setAttribute("data-hint-tap", "");   // dotyk: tap = zerknięcie, przytrzymanie = trzymaj
        val.setAttribute("data-hint-fade", "");  // płynne zniknięcie po „zerknięciu"
        val.setAttribute("data-hint-delay", "0.4");
      }
    }

    item.appendChild(lab);
    item.appendChild(track);
    item.appendChild(val);
    list.appendChild(item);
  });
  monthlySummaryEl.appendChild(list);
}
