// Formula extraction and the formula workbench UI.

function extractFormulaFunctionName(formulaText) {
  const text = String(formulaText || "").replace(/^=/, "").trim();
  const match = text.match(/^([A-Z_][A-Z0-9\._]*)\s*\(/i);
  return match ? match[1].toUpperCase() : "INNE";
}

function collectFormulaEntries(sheet, data, headerRow) {
  if (!sheet || !data || !Array.isArray(data.headers) || !data.headers.length) return [];
  const entries = [];
  const range = computeEffectiveSheetRange(sheet, headerRow);

  Object.keys(sheet).forEach((key) => {
    if (!key || key[0] === "!") return;
    const cell = sheet[key];
    if (!cell || !cell.f) return;
    const ref = XLSX.utils.decode_cell(key);
    if (ref.r < range.s.r || ref.r > range.e.r || ref.c < range.s.c || ref.c > range.e.c) return;

    const formulaText = `=${cell.f}`;
    const functionName = extractFormulaFunctionName(formulaText);
    const resultText = localizeDisplayedDate(cell.v, cell.w != null ? String(cell.w) : toDisplay(cell.v), cell);
    const missingResult = cell.v == null && cell.w == null;
    const hasError = String(resultText || "").trim().startsWith("#");
    const colIdx = ref.c - data.startCol;
    const inTable = ref.r >= headerRow && colIdx >= 0 && colIdx < data.headers.length;

    entries.push({
      address: key,
      formulaText,
      functionName,
      resultText,
      missingResult,
      hasError,
      rowIndex0: ref.r,
      colAbs: ref.c,
      colIdx,
      inTable,
      header: inTable ? data.headers[colIdx] : XLSX.utils.encode_col(ref.c),
    });
  });

  return entries.sort((a, b) => {
    if (a.missingResult !== b.missingResult) return a.missingResult ? -1 : 1;
    if (a.hasError !== b.hasError) return a.hasError ? -1 : 1;
    if (a.functionName !== b.functionName) return a.functionName.localeCompare(b.functionName, "pl");
    return a.address.localeCompare(b.address, "pl");
  });
}

function getFilteredFormulaEntries() {
  const search = String(formulaSearchEl?.value || "").trim().toLowerCase();
  const filter = formulaFilterEl?.value || "all";
  const functionFilter = String(formulaFunctionFilterEl?.value || "").trim().toUpperCase();
  return currentFormulaEntries.filter((entry) => {
    if (filter === "missing" && !entry.missingResult) return false;
    if (filter === "error" && !entry.hasError) return false;
    if (functionFilter && entry.functionName !== functionFilter) return false;
    if (!search) return true;
    const haystack = [
      entry.address,
      entry.header,
      entry.functionName,
      entry.formulaText,
      entry.resultText,
    ].join(" ").toLowerCase();
    return haystack.includes(search);
  });
}

function renderFormulaFunctionFilter() {
  if (!formulaFunctionFilterEl) return;
  const previous = formulaFunctionFilterEl.value;
  const names = Array.from(new Set(currentFormulaEntries.map((entry) => entry.functionName))).sort((a, b) => a.localeCompare(b, "pl"));
  formulaFunctionFilterEl.replaceChildren();

  const allOpt = document.createElement("option");
  allOpt.value = "";
  allOpt.textContent = t("allFunctions");
  formulaFunctionFilterEl.appendChild(allOpt);

  names.forEach((name) => {
    const opt = document.createElement("option");
    opt.value = name;
    opt.textContent = name;
    formulaFunctionFilterEl.appendChild(opt);
  });

  formulaFunctionFilterEl.value = names.includes(previous) ? previous : "";
}

function truncateFormulaPreview(text, maxLength = 120) {
  const raw = String(text || "").trim();
  if (raw.length <= maxLength) return raw;
  const head = Math.max(36, Math.floor(maxLength * 0.55));
  const tail = Math.max(18, maxLength - head - 3);
  return `${raw.slice(0, head)}...${raw.slice(-tail)}`;
}

function formatFormulaAddressSample(entries, limit = 4) {
  const sample = entries.slice(0, limit).map((entry) => entry.address);
  if (entries.length <= limit) return sample.join(", ");
  return `${sample.join(", ")} +${entries.length - limit}`;
}

function aggregateFormulaEntries(entries) {
  const groups = new Map();
  entries.forEach((entry) => {
    const key = [
      entry.functionName,
      entry.formulaText,
      entry.header,
      entry.missingResult ? "1" : "0",
      entry.hasError ? "1" : "0",
      entry.inTable ? "1" : "0",
    ].join("||");
    const existing = groups.get(key);
    if (existing) {
      existing.entries.push(entry);
      return;
    }
    groups.set(key, {
      key,
      functionName: entry.functionName,
      formulaText: entry.formulaText,
      header: entry.header,
      missingResult: entry.missingResult,
      hasError: entry.hasError,
      inTable: entry.inTable,
      resultText: entry.resultText,
      entries: [entry],
      firstEntry: entry,
    });
  });
  return Array.from(groups.values()).sort((a, b) => {
    if (a.missingResult !== b.missingResult) return a.missingResult ? -1 : 1;
    if (a.hasError !== b.hasError) return a.hasError ? -1 : 1;
    if (b.entries.length !== a.entries.length) return b.entries.length - a.entries.length;
    if (a.functionName !== b.functionName) return a.functionName.localeCompare(b.functionName, "pl");
    return a.firstEntry.address.localeCompare(b.firstEntry.address, "pl");
  });
}

function renderFormulaWorkbench() {
  if (!formulaWorkbenchSummaryEl || !formulaWorkbenchListEl) return;
  formulaWorkbenchSummaryEl.replaceChildren();
  formulaWorkbenchListEl.replaceChildren();
  renderFormulaFunctionFilter();

  if (!currentHeaders.length || !currentFormulaEntries.length) {
    renderInsightList(
      formulaWorkbenchSummaryEl,
      [],
      t("formulaNoFormulas")
    );
    formulaWorkbenchListEl.appendChild(createEmptyInsight(t("formulaNoList")));
    return;
  }

  const filtered = getFilteredFormulaEntries();
  const grouped = aggregateFormulaEntries(filtered);
  const functionCounts = new Map();
  currentFormulaEntries.forEach((entry) => {
    functionCounts.set(entry.functionName, (functionCounts.get(entry.functionName) || 0) + 1);
  });
  const topFunction = Array.from(functionCounts.entries()).sort((a, b) => b[1] - a[1])[0];
  const summaryItems = [
    { label: t("formulaSummaryFormulas"), value: String(currentFormulaEntries.length) },
    {
      label: t("formulaSummaryMissing"),
      value: String(currentFormulaEntries.filter((entry) => entry.missingResult).length),
      tone: currentFormulaEntries.some((entry) => entry.missingResult) ? "warning" : "",
    },
    {
      label: t("formulaSummaryErrors"),
      value: String(currentFormulaEntries.filter((entry) => entry.hasError).length),
      tone: currentFormulaEntries.some((entry) => entry.hasError) ? "warning" : "",
    },
    {
      label: t("formulaSummaryTop"),
      value: topFunction ? `${topFunction[0]} ×${topFunction[1]}` : t("aggregationNone"),
      tone: topFunction ? "info" : "",
    },
    {
      label: t("formulaSummaryVisible"),
      value: String(filtered.length),
      tone: filtered.length !== currentFormulaEntries.length ? "info" : "",
    },
    {
      label: t("formulaSummaryGroups"),
      value: String(grouped.length),
      tone: grouped.length < filtered.length ? "info" : "",
    },
  ];

  renderInsightList(formulaWorkbenchSummaryEl, summaryItems, t("formulaNoSummary"));

  if (!filtered.length) {
    formulaWorkbenchListEl.appendChild(createEmptyInsight(t("formulaNoFilterMatch")));
    return;
  }

  grouped.slice(0, 60).forEach((group) => {
    const item = document.createElement("div");
    item.className = "formula-item";

    const top = document.createElement("div");
    top.className = "formula-item-top";

    const title = document.createElement("div");
    title.className = "formula-item-title";
    title.textContent = group.entries.length > 1
      ? `${group.header} • ${t("formulaSameCount", { count: group.entries.length })}`
      : `${group.firstEntry.address} • ${group.header}`;

    const kind = document.createElement("div");
    kind.className = `formula-item-kind${group.missingResult || group.hasError ? " warning" : ""}`;
    kind.textContent = group.functionName;

    top.appendChild(title);
    top.appendChild(kind);

    const formula = document.createElement("div");
    formula.className = "formula-item-formula";
    formula.textContent = truncateFormulaPreview(group.formulaText);
    formula.title = group.formulaText;

    const meta = document.createElement("div");
    meta.className = "formula-item-meta";
    const resultLabel = group.missingResult ? t("formulaSummaryMissing").toLowerCase() : (group.resultText || t("formulaEmptyResult"));
    const addressLabel = formatFormulaAddressSample(group.entries);
    const outsideTable = group.inTable ? "" : ` • ${t("formulaOutsideTable")}`;
    meta.textContent = `${t("formulaAddresses")}: ${addressLabel} • ${t("formulaResultLabel")}: ${resultLabel}${outsideTable}`;

    const actions = document.createElement("div");
    actions.className = "section-nav-actions";

    const btn = document.createElement("button");
    btn.className = "btn ghost btn-sm";
    btn.type = "button";
    btn.dataset.formulaAddress = group.firstEntry.address;
    btn.textContent = group.entries.length > 1 ? t("formulaJumpFirst") : t("formulaJumpCell");

    actions.appendChild(btn);
    item.appendChild(top);
    item.appendChild(formula);
    item.appendChild(meta);
    item.appendChild(actions);
    formulaWorkbenchListEl.appendChild(item);
  });
}

function focusFormulaEntry(address) {
  if (!address || !currentDisplayModel?.rows?.length) return;
  const ref = XLSX.utils.decode_cell(address);
  const displayRow = currentDisplayModel.rows.find((row) => (row.sourceRowIndex0 ?? row.rowIndex0) === ref.r);
  if (!displayRow) {
    toast(t("formulaOutsideView"), "info");
    return;
  }

  let targetColIndex0 = 0;
  if (currentDisplayModel.mode === "wide") {
    targetColIndex0 = Math.max(0, ref.c - currentStartCol);
  } else {
    const group = getActiveRepeatingGroup();
    const prefixCount = Math.max(0, Number(group?.prefixCount) || 0);
    if (ref.c < currentStartCol + prefixCount) {
      targetColIndex0 = Math.max(0, ref.c - currentStartCol);
    } else {
      const block = group?.blocks?.[displayRow.sourceBlockIndex ?? 0];
      if (block && ref.c >= currentStartCol + block.startIndex && ref.c <= currentStartCol + block.endIndex) {
        targetColIndex0 = prefixCount + 2 + (ref.c - (currentStartCol + block.startIndex));
      }
    }
  }

  const boundedColIndex0 = Math.min(targetColIndex0, currentDisplayModel.headers.length - 1);
  const rowKey = getRowSelectionKey(displayRow);
  setFocusedCell(rowKey, boundedColIndex0, { scroll: true });
  setSelectedCell(rowKey, boundedColIndex0, { scroll: true });
}
