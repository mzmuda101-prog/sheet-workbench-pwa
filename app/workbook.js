// Workbook IO, filtering, sorting, and shared view state transitions.

function isXlsxAvailable(showFeedback = false) {
  const available = typeof window !== "undefined" && !!window.XLSX;
  if (!available && showFeedback) {
    setStatus(t("xlsxMissingStatus"));
    toast(t("xlsxMissingToast"), "error");
    log("Brak biblioteki XLSX (window.XLSX).", "error");
  }
  return available;
}

function setRuntimeAvailability(isAvailable) {
  fileInput.disabled = !isAvailable;
  loadBtn.disabled = !isAvailable;
  if (loadSampleBtn) {
    loadSampleBtn.disabled = !isAvailable;
    loadSampleBtn.classList.toggle("hidden", !isAvailable);
  }
  saveAsBtn.disabled = !isAvailable;
  if (excelLayoutToggleEl) {
    excelLayoutToggleEl.disabled = !isAvailable;
    excelLayoutToggleEl.setAttribute("aria-disabled", isAvailable ? "false" : "true");
  }
  saveAsBtn.setAttribute("aria-disabled", isAvailable ? "false" : "true");
  dropZone.classList.toggle("disabled", !isAvailable);
  dropZone.setAttribute("aria-disabled", isAvailable ? "false" : "true");
}

function loadMaxRowsPreference() {
  const saved = localStorage.getItem(MAX_ROWS_KEY);
  const value = saved ? parseInt(saved, 10) : null;
  if (value && Number.isFinite(value) && value > 0) {
    maxRowsEl.value = String(value);
  }
}

function saveMaxRowsPreference() {
  const value = Math.max(1, parseInt(maxRowsEl.value || "200", 10));
  localStorage.setItem(MAX_ROWS_KEY, String(value));
}

function loadExcelLayoutPreference() {
  if (!excelLayoutToggleEl) return;
  setExcelLayoutEnabled(localStorage.getItem(EXCEL_LAYOUT_KEY) === "1");
}

function saveExcelLayoutPreference() {
  if (!excelLayoutToggleEl) return;
  localStorage.setItem(EXCEL_LAYOUT_KEY, isExcelLayoutEnabled() ? "1" : "0");
}

function isExcelLayoutEnabled() {
  if (!excelLayoutToggleEl) return false;
  return excelLayoutToggleEl.getAttribute("aria-pressed") === "true";
}

function setExcelLayoutEnabled(enabled) {
  if (!excelLayoutToggleEl) return;
  const next = !!enabled;
  excelLayoutToggleEl.setAttribute("aria-pressed", next ? "true" : "false");
  excelLayoutToggleEl.classList.toggle("active", next);
  excelLayoutToggleEl.textContent = next ? t("excelViewOn") : t("excelView");
}

function setEmptyState(title, subtitle) {
  emptyTitleEl.textContent = title;
  emptySubEl.textContent = subtitle;
  emptyStateEl.classList.remove("hidden");
  tableWrapEl.classList.add("hidden");
}

function showTable() {
  emptyStateEl.classList.add("hidden");
  tableWrapEl.classList.remove("hidden");
  if (tableScrollbarEl) tableScrollbarEl.classList.remove("hidden");
}

function hideCellTooltip() {
  if (!cellTooltipEl) return;
  if (tooltipHideTimer) {
    clearTimeout(tooltipHideTimer);
    tooltipHideTimer = null;
  }
  cellTooltipEl.classList.add("hidden");
  cellTooltipEl.textContent = "";
}

function getTooltipText(cell) {
  if (!cell) return "";
  return (cell.dataset.fullText || cell.textContent || "").trim();
}

function isCellTextTruncated(cell) {
  if (!cell) return false;
  return cell.scrollWidth - cell.clientWidth > 1;
}

function positionCellTooltip(cell) {
  if (!cellTooltipEl || !cell) return;
  const rect = cell.getBoundingClientRect();
  const tooltipRect = cellTooltipEl.getBoundingClientRect();
  const margin = 12;
  let left = rect.left + rect.width / 2 - tooltipRect.width / 2;
  left = Math.max(margin, Math.min(left, window.innerWidth - tooltipRect.width - margin));
  let top = rect.top - tooltipRect.height - 10;
  if (top < margin) top = Math.min(window.innerHeight - tooltipRect.height - margin, rect.bottom + 10);
  cellTooltipEl.style.left = `${left}px`;
  cellTooltipEl.style.top = `${top}px`;
}

function showCellTooltip(cell, persistent = false) {
  if (!cellTooltipEl || !isCellTextTruncated(cell)) return;
  const text = getTooltipText(cell);
  if (!text) return;
  hideCellTooltip();
  cellTooltipEl.textContent = text;
  cellTooltipEl.classList.remove("hidden");
  positionCellTooltip(cell);
  if (!persistent) return;
  tooltipHideTimer = window.setTimeout(() => {
    hideCellTooltip();
  }, 2200);
}

function syncHorizontalScrollbar() {
  if (!tableWrapEl || !tableScrollbarEl || !tableScrollbarInnerEl) return;
  const active = !tableWrapEl.classList.contains("hidden") && tableWrapEl.scrollWidth > tableWrapEl.clientWidth + 1;
  tableScrollbarEl.classList.toggle("hidden", !active);
  if (!active) return;
  tableScrollbarInnerEl.style.width = `${tableWrapEl.scrollWidth}px`;
  if (Math.abs(tableScrollbarEl.scrollLeft - tableWrapEl.scrollLeft) > 1) {
    tableScrollbarEl.scrollLeft = tableWrapEl.scrollLeft;
  }
}

function toDisplay(value) {
  if (value === null || value === undefined) return "";
  if (value instanceof Date) {
    const dd = String(value.getDate()).padStart(2, "0");
    const mm = String(value.getMonth() + 1).padStart(2, "0");
    const yy = String(value.getFullYear()).slice(-2);
    return `${dd}-${mm}-${yy}`;
  }
  return String(value);
}

function formatLocalizedDateDisplay(date, options = {}) {
  if (!(date instanceof Date) || Number.isNaN(date.getTime())) return "";
  const formatterOptions = {
    month: options.month || "short",
  };

  if (options.day !== false) formatterOptions.day = "2-digit";
  if (options.year !== false) formatterOptions.year = options.year || "2-digit";

  if (options.weekday) formatterOptions.weekday = options.weekday;
  if (options.timeStyle) {
    formatterOptions.hour = "2-digit";
    formatterOptions.minute = "2-digit";
    if (options.timeStyle === "withSeconds") {
      formatterOptions.second = "2-digit";
    }
  }

  return new Intl.DateTimeFormat(I18N[currentLang]?.locale || "pl-PL", formatterOptions)
    .format(date)
    .replace(/\.$/g, "")
    .replace(/\s+/g, " ")
    .trim();
}

function localizeDisplayedDate(value, shown, cell = null) {
  if (!shown || typeof shown !== "string") return shown;

  const formatHint = String(cell?.z || cell?.w || "").toLowerCase();
  const hasMonthNameFormatHint = /(mmmm|mmm)/i.test(formatHint);
  let date = parseDateFlexible(value);
  if (hasMonthNameFormatHint && typeof value === "number" && Number.isFinite(value) && value >= 1 && value <= 12) {
    date = new Date(2000, value - 1, 1);
  }
  if (!(date instanceof Date) || Number.isNaN(date.getTime())) return shown;

  const normalizedShown = shown.trim().toLowerCase();
  const hasEnglishWeekdayWord = /\b(mon|monday|tue|tues|tuesday|wed|wednesday|thu|thur|thurs|thursday|fri|friday|sat|saturday|sun|sunday)\b/i.test(
    normalizedShown
  );
  const hasPolishWeekdayWord = /\b(pon|poniedzialek|poniedziałek|wt|wtorek|sr|śr|sroda|środa|czw|czwartek|pt|piatek|piątek|sob|sobota|niedz|niedziela)\b/i.test(
    normalizedShown
  );
  const hasEnglishMonthWord = /\b(jan|feb|mar|apr|may|jun|jul|aug|sep|sept|oct|nov|dec|january|february|march|april|june|july|august|september|october|november|december)\b/i.test(
    normalizedShown
  );
  const hasPolishMonthWord = /\b(sty|styczen|styczeń|lut|luty|mar|marzec|kwi|kwiecien|kwiecień|maj|cze|czerwiec|lip|lipiec|sie|sierpien|sierpień|wrz|wrzesien|wrzesień|paz|paź|pazdziernik|październik|lis|listopad|gru|grudzien|grudzień)\b/i.test(
    normalizedShown
  );
  const hasWeekdayFormatHint = /(dddd|ddd)/i.test(formatHint);
  const hasTime = /\b\d{1,2}:\d{2}(:\d{2})?\b/.test(shown) || /\b(h|hh|m|mm|s|ss)\b/i.test(formatHint);
  const monthOnlyFormat = hasMonthNameFormatHint && !/[dy]/i.test(formatHint.replace(/m+/gi, ""));

  if (!hasEnglishMonthWord && !hasPolishMonthWord && !hasMonthNameFormatHint && !hasEnglishWeekdayWord && !hasPolishWeekdayWord && !hasWeekdayFormatHint) {
    return shown;
  }

  const weekday =
    hasEnglishWeekdayWord || hasPolishWeekdayWord || hasWeekdayFormatHint
      ? /dddd/i.test(formatHint) || /\b(monday|tuesday|wednesday|thursday|friday|saturday|sunday)\b/i.test(normalizedShown)
        ? "long"
        : "short"
      : undefined;

  const month =
    /mmmm/i.test(formatHint) || /\b(january|february|march|april|june|july|august|september|october|november|december|styczen|styczeń|kwiecien|kwiecień|czerwiec|sierpien|sierpień|wrzesien|wrzesień|pazdziernik|październik|listopad|grudzien|grudzień)\b/i.test(normalizedShown)
      ? "long"
      : "short";

  const year = /\b\d{4}\b/.test(shown) || /yyyy/i.test(formatHint) ? "numeric" : "2-digit";
  const timeStyle = hasTime ? (/\b\d{1,2}:\d{2}:\d{2}\b/.test(shown) || /ss/i.test(formatHint) ? "withSeconds" : "short") : null;

  return (
    formatLocalizedDateDisplay(date, {
      weekday,
      month,
      day: monthOnlyFormat ? false : "2-digit",
      year: monthOnlyFormat ? false : year,
      timeStyle,
    }) || shown
  );
}

function getDisplayValue(row, index) {
  if (row && Array.isArray(row.display) && index < row.display.length) {
    return row.display[index];
  }
  if (row && Array.isArray(row.values) && index < row.values.length) {
    return toDisplay(row.values[index]);
  }
  return "";
}

function parseDateFlexible(value) {
  if (value instanceof Date) return value;
  if (typeof value === "number" && Number.isFinite(value)) {
    const ms = (value - 25569) * 86400000;
    const d = new Date(ms);
    return Number.isNaN(d.getTime()) ? null : d;
  }
  if (typeof value !== "string") return null;
  let v = value.trim();
  if (!v) return null;

  if (/^\d+(\.\d+)?$/.test(v)) {
    const numeric = Number(v);
    if (Number.isFinite(numeric)) {
      const ms = (numeric - 25569) * 86400000;
      const d = new Date(ms);
      return Number.isNaN(d.getTime()) ? null : d;
    }
  }

  v = v.replace(/T.*$/, "");
  v = v.replace(/\s+\d{1,2}:\d{2}(:\d{2})?.*$/, "");
  const normalized = v.replace(/[.\/]/g, "-");

  const monthMap = {
    // PL
    "sty": 1, "stycz": 1, "stycznia": 1,
    "lut": 2, "lutego": 2,
    "mar": 3, "marca": 3,
    "kwi": 4, "kwie": 4, "kwietnia": 4,
    "maj": 5, "maja": 5,
    "cze": 6, "czer": 6, "czerwca": 6,
    "lip": 7, "lipca": 7,
    "sie": 8, "sier": 8, "sierpnia": 8,
    "wrz": 9, "wrzes": 9, "wrzesnia": 9,
    "paź": 10, "paz": 10, "paźdz": 10, "pazdz": 10, "października": 10, "pazdziernika": 10,
    "lis": 11, "list": 11, "listopada": 11,
    "gru": 12, "grud": 12, "grudnia": 12,
    // EN
    "jan": 1, "january": 1,
    "feb": 2, "february": 2,
    "mar": 3, "march": 3,
    "apr": 4, "april": 4,
    "may": 5,
    "jun": 6, "june": 6,
    "jul": 7, "july": 7,
    "aug": 8, "august": 8,
    "sep": 9, "sept": 9, "september": 9,
    "oct": 10, "october": 10,
    "nov": 11, "november": 11,
    "dec": 12, "december": 12,
  };

  let m = normalized.match(/^(\d{1,2})-(\d{1,2})-(\d{2}|\d{4})$/);
  if (m) {
    const y = m[3].length === 2 ? Number(`20${m[3]}`) : Number(m[3]);
    return new Date(y, Number(m[2]) - 1, Number(m[1]));
  }

  m = normalized.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (m) {
    return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  }

  const words = v.toLowerCase()
    .replace(/,/g, "")
    .replace(/\s+/g, " ")
    .trim();
  let wm = words.match(/^(\d{1,2})\s+([a-ząćęłńóśźż\.]+)\s+(\d{4}|\d{2})$/i);
  if (wm) {
    const day = Number(wm[1]);
    const monthKey = wm[2].replace(/\.$/, "");
    const month = monthMap[monthKey];
    const year = wm[3].length === 2 ? Number(`20${wm[3]}`) : Number(wm[3]);
    if (month) return new Date(year, month - 1, day);
  }
  wm = words.match(/^([a-ząćęłńóśźż\.]+)\s+(\d{1,2})\s+(\d{4}|\d{2})$/i);
  if (wm) {
    const monthKey = wm[1].replace(/\.$/, "");
    const month = monthMap[monthKey];
    const day = Number(wm[2]);
    const year = wm[3].length === 2 ? Number(`20${wm[3]}`) : Number(wm[3]);
    if (month) return new Date(year, month - 1, day);
  }

  const parsed = Date.parse(value);
  if (!Number.isNaN(parsed)) return new Date(parsed);
  return null;
}

function parseInputValue(raw) {
  const text = String(raw ?? "").trim();
  if (!text) return null;
  if (text.startsWith("=")) return { value: text, type: "formula" };
  const asDate = parseDateFlexible(text);
  if (asDate) return { value: asDate, type: "date" };
  if (/^-?\d+(\.\d+)?$/.test(text)) return { value: Number(text), type: "number" };
  return { value: text, type: "string" };
}

function updateSheetCell(rowIndex0, colIndex0, parsed) {
  if (!workbook || !currentSheetName) return;
  const sheet = workbook.Sheets[currentSheetName];
  if (!sheet) return;
  const absoluteCol = currentStartCol + colIndex0;
  const cellRef = XLSX.utils.encode_cell({ r: rowIndex0, c: absoluteCol });
  if (!parsed || parsed.value === null) {
    delete sheet[cellRef];
    return;
  }
  if (parsed.type === "formula") {
    toast(t("formulaEditBlocked"), "warning");
    return;
  }
  if (parsed.type === "date") {
    sheet[cellRef] = { v: parsed.value, t: "d" };
    return;
  }
  if (parsed.type === "number") {
    sheet[cellRef] = { v: parsed.value, t: "n" };
    return;
  }
  sheet[cellRef] = { v: parsed.value, t: "s" };
}

function getDateRange() {
  const mode = getNormalizedSelectValue(dateModeEl);
  if (mode === "last_n_days") {
    const days = Math.max(1, parseInt(lastDaysEl.value || "30", 10));
    const now = new Date();
    const from = new Date(now.getFullYear(), now.getMonth(), now.getDate() - days);
    return { from, to: now };
  }
  const from = parseDateFlexible(dateFromEl.value);
  const to = parseDateFlexible(dateToEl.value);
  if (mode === "before") return { from: null, to };
  if (mode === "after") return { from, to: null };
  return { from, to };
}

function rowMatchesEmptyMode(row, indexes, emptyMode) {
  if (!emptyMode || emptyMode === "all") return true;
  const resolvedIndexes = indexes && indexes.length ? indexes : row.values.map((_, i) => i);
  const emptyStates = resolvedIndexes.map((i) => {
    if (i >= row.values.length) return true;
    return getDisplayValue(row, i).trim().length === 0;
  });
  if (!emptyStates.length) return emptyMode === "any_empty";
  if (emptyMode === "any_empty") return emptyStates.some(Boolean);
  if (emptyMode === "all_empty") return emptyStates.every(Boolean);
  if (emptyMode === "any_non_empty") return emptyStates.some((isEmpty) => !isEmpty);
  if (emptyMode === "all_non_empty") return emptyStates.every((isEmpty) => !isEmpty);
  return true;
}

function combinePrimaryAndEmptyMatch(primaryMatched, emptyMatched, negated, hasPrimaryRule, hasEmptyRule) {
  if (!hasPrimaryRule && !hasEmptyRule) return true;
  if (hasPrimaryRule && negated) {
    if (hasEmptyRule) return !primaryMatched && emptyMatched;
    return !primaryMatched;
  }
  if (hasPrimaryRule && !negated) {
    if (hasEmptyRule) return primaryMatched && emptyMatched;
    return primaryMatched;
  }
  if (hasEmptyRule && negated) return !emptyMatched;
  if (hasEmptyRule) return emptyMatched;
  return true;
}

function rowMatchesSingleTerm(row, term, criterion) {
  const values = row.values;
  const q = term.trim().toLowerCase();
  if (!q) return true;
  for (const i of criterion.indexes) {
    if (i >= values.length) continue;
    const text = getDisplayValue(row, i).toLowerCase();
    const altDate = parseDateFlexible(values[i]);
    const candidates = [text];
    if (altDate instanceof Date) {
      const dd = String(altDate.getDate()).padStart(2, "0");
      const mm = String(altDate.getMonth() + 1).padStart(2, "0");
      const yyyy = String(altDate.getFullYear());
      const yy = yyyy.slice(-2);
      candidates.push(`${dd}-${mm}-${yy}`);
      candidates.push(`${dd}-${mm}-${yyyy}`);
    }
    if (criterion.mode === "equals" && candidates.some((c) => c === q)) return true;
    if (criterion.mode === "starts_with" && candidates.some((c) => c.startsWith(q))) return true;
    if (criterion.mode === "contains" && candidates.some((c) => c.includes(q))) return true;
  }
  return false;
}

function parseOperatorTerm(rawTerm, operatorsEnabled = false) {
  let term = String(rawTerm || "").trim();
  const parsed = { term, negated: false };
  if (!operatorsEnabled || !term) return parsed;

  if (term.startsWith("!") && term.length > 1) {
    parsed.negated = true;
    term = term.slice(1).trim();
  }
  parsed.term = term;
  return parsed;
}

// Tokenize a flat string (no brackets) into NOT-aware terms split by && and ||
// Returns an AST node: { type: "or", children: [ { type: "and", terms: [...] } ] }
function parseFlatExpr(text) {
  function parseAndPart(part) {
    const terms = [];
    const s = String(part || "");
    const notPattern = /(^|\s)!(\S+)/g;
    let cursor = 0;
    let match = notPattern.exec(s);
    while (match) {
      const before = s.slice(cursor, match.index).trim();
      if (before) terms.push({ term: before, negated: false });
      terms.push({ term: match[2], negated: true });
      cursor = match.index + match[0].length;
      match = notPattern.exec(s);
    }
    const after = s.slice(cursor).trim();
    if (after) terms.push({ term: after, negated: false });
    return terms.filter((t) => t.term);
  }

  const andGroups = text.split("||").map((orPart) =>
    orPart.split("&&").flatMap((t) => parseAndPart(t))
  ).filter((g) => g.length);

  return { type: "or", children: andGroups.map((g) => ({ type: "and", terms: g })) };
}

// Split a query string respecting {} brackets into segments with their surrounding operators.
// Returns array of { segment: string|AST, op: "&&"|"||"|null, negated: bool }
// bracket segments are pre-parsed into inner AST nodes.
function tokenizeBrackets(query) {
  const tokens = []; // { kind: "text"|"bracket", value: string, negated: bool }
  let i = 0;
  const s = String(query || "");
  while (i < s.length) {
    if (s[i] === "{") {
      // Find matching closing bracket
      let depth = 1;
      let j = i + 1;
      while (j < s.length && depth > 0) {
        if (s[j] === "{") depth++;
        else if (s[j] === "}") depth--;
        j++;
      }
      const inner = s.slice(i + 1, j - 1);
      tokens.push({ kind: "bracket", value: inner });
      i = j;
    } else {
      // Collect text until next {
      let j = i;
      while (j < s.length && s[j] !== "{") j++;
      tokens.push({ kind: "text", value: s.slice(i, j) });
      i = j;
    }
  }
  return tokens;
}

// Full query parser supporting {} grouping, &&, ||, !
// Returns { groups: [ [term, ...], ... ] }
// Evaluation: groups.some(group => group.every(term => match(term)))
// A term is { term: string, negated: bool } OR { subExpr: AST, negated: bool }
// where AST = { type:"or", children: [{type:"and", terms:[...]}] }
function parseQueryTerms(query, operatorsEnabled = false) {
  const raw = String(query || "").trim();
  const simpleTerm = { term: raw, negated: false };
  const fallback = { groups: raw ? [[simpleTerm]] : [] };
  if (!operatorsEnabled) return fallback;

  const bracketTokens = tokenizeBrackets(raw);

  // Re-assemble into a flat stream, replacing {…} with placeholders
  // Then split by || (OR) and && (AND)
  const BRACKET_PH = "\x00BRACKET_";
  const brackets = [];
  let flat = "";
  for (const tok of bracketTokens) {
    if (tok.kind === "bracket") {
      const idx = brackets.length;
      brackets.push(tok.value);
      flat += `${BRACKET_PH}${idx}\x00`;
    } else {
      flat += tok.value;
    }
  }

  // Split by || then &&
  const orParts = flat.split("||");
  // Each OR part is an AND group
  const groups = [];
  for (const orPart of orParts) {
    const andParts = orPart.split("&&");
    const andTerms = [];
    for (const andPart of andParts) {
      // Find bracket placeholders in this part
      const phPattern = /!?\x00BRACKET_(\d+)\x00/g;
      let cursor = 0;
      let m = phPattern.exec(andPart);
      const segments = [];
      while (m) {
        const before = andPart.slice(cursor, m.index).trim();
        if (before) segments.push({ kind: "text", value: before });
        segments.push({ kind: "bracket", index: parseInt(m[1], 10), negated: m[0].startsWith("!") });
        cursor = m.index + m[0].length;
        m = phPattern.exec(andPart);
      }
      const after = andPart.slice(cursor).trim();
      if (after) segments.push({ kind: "text", value: after });

      for (const seg of segments) {
        if (seg.kind === "bracket") {
          const innerAst = parseFlatExpr(brackets[seg.index]);
          if (innerAst.children.length) {
            andTerms.push({ subExpr: innerAst, negated: seg.negated });
          }
        } else {
          // Parse text part for NOT and individual terms
          const s = seg.value;
          const notPattern = /(^|\s)!(\S+)/g;
          let cur = 0;
          let mt = notPattern.exec(s);
          while (mt) {
            const bef = s.slice(cur, mt.index).trim();
            if (bef) andTerms.push({ term: bef, negated: false });
            andTerms.push({ term: mt[2], negated: true });
            cur = mt.index + mt[0].length;
            mt = notPattern.exec(s);
          }
          const aft = s.slice(cur).trim();
          if (aft) andTerms.push({ term: aft, negated: false });
        }
      }
    }
    const validTerms = andTerms.filter((t) => t.term || t.subExpr);
    if (validTerms.length) groups.push(validTerms);
  }

  return groups.length ? { groups } : fallback;
}

function rowMatchesAstNode(row, ast, criterion) {
  // ast = { type: "or", children: [ { type: "and", terms: [...] } ] }
  return ast.children.some((andGroup) =>
    andGroup.terms.every((t) => rowMatchesParsedTerm(row, t, criterion))
  );
}

function rowMatchesParsedTerm(row, parsedTerm, criterion) {
  let matched;
  if (parsedTerm.subExpr) {
    matched = rowMatchesAstNode(row, parsedTerm.subExpr, criterion);
  } else {
    matched = rowMatchesSingleTerm(row, parsedTerm.term, criterion);
  }
  return parsedTerm.negated ? !matched : matched;
}

function rowMatchesTextFilter(row, criteria, onlyNonEmpty) {
  const values = row.values;
  let usedIndexes = new Set();
  criteria.forEach((c) => c.indexes.forEach((i) => usedIndexes.add(i)));
  const indexes = usedIndexes.size ? Array.from(usedIndexes) : values.map((_, i) => i);

  if (onlyNonEmpty) {
    const anyNonEmpty = indexes.some((i) => {
      const txt = getDisplayValue(row, i).trim();
      return txt.length > 0;
    });
    if (!anyNonEmpty) return false;
  }

  for (const criterion of criteria) {
    const query = criterion.query;
    const emptyMode = criterion.emptyMode || "all";
    const hasQuery = !!query;
    const hasEmptyRule = emptyMode !== "all";
    if (!hasQuery && !hasEmptyRule) continue;

    let textMatched = !hasQuery;
    if (hasQuery) {
      const parsed = parseQueryTerms(query, criterion.operatorsEnabled);
      textMatched = parsed.groups.some((group) => group.every((term) => rowMatchesParsedTerm(row, term, criterion)));
    }

    const emptyMatched = rowMatchesEmptyMode(row, criterion.indexes, emptyMode);
    const matched = combinePrimaryAndEmptyMatch(textMatched, emptyMatched, criterion.negated, hasQuery, hasEmptyRule);
    if (!matched) return false;
  }

  return true;
}

function rowMatchesDateFilter(row, filter) {
  const indexes = filter.indexes || [];
  const dateRange = filter.range || { from: null, to: null };
  const hasRange = !!(dateRange.from || dateRange.to);
  const emptyMode = filter.emptyMode || "all";
  const hasEmptyRule = emptyMode !== "all";
  if (!hasRange && !hasEmptyRule) return true;

  let rangeMatched = !hasRange;
  if (hasRange) {
    rangeMatched = false;
    for (const idx of indexes) {
      if (idx >= row.values.length) continue;
      const raw = row.rawValues ? row.rawValues[idx] : row.values[idx];
      const d = parseDateFlexible(raw ?? getDisplayValue(row, idx));
      if (!d) continue;
      if (dateRange.from && d < dateRange.from) continue;
      if (dateRange.to && d > dateRange.to) continue;
      rangeMatched = true;
      break;
    }
  }

  const emptyMatched = rowMatchesEmptyMode(row, indexes, emptyMode);
  return combinePrimaryAndEmptyMatch(rangeMatched, emptyMatched, filter.negated, hasRange, hasEmptyRule);
}

function resolveIndexes(headers, selected) {
  if (!selected.size) return headers.map((_, i) => i);
  return headers.map((h, i) => (selected.has(h) ? i : -1)).filter((i) => i >= 0);
}

function compareSortValues(av, bv) {
  const ad = parseDateFlexible(av);
  const bd = parseDateFlexible(bv);
  if (ad && bd) return ad - bd;
  if (typeof av === "number" && typeof bv === "number") return av - bv;
  return String(av || "").localeCompare(String(bv || ""), "pl");
}

function normalizeSortState() {
  multiSortState = multiSortState
    .filter((rule) => rule && rule.col && currentHeaders.includes(rule.col))
    .map((rule) => ({ col: rule.col, dir: rule.dir === "desc" ? "desc" : "asc" }));
  const primary = multiSortState[0] || null;
  sortState = primary ? { col: primary.col, dir: primary.dir } : { col: "", dir: "asc" };
}

function setPrimarySort(col, dir = "asc") {
  if (!col) {
    multiSortState = [];
    normalizeSortState();
    return;
  }
  const next = [{ col, dir: dir === "desc" ? "desc" : "asc" }];
  multiSortState.forEach((rule) => {
    if (rule.col === col) return;
    next.push(rule);
  });
  multiSortState = next;
  normalizeSortState();
}

function populateSortColumnSelect() {
  if (!sortColumnSelectEl) return;
  sortColumnSelectEl.replaceChildren();
  if (!currentHeaders.length) {
    const opt = document.createElement("option");
    opt.value = "";
    opt.textContent = t("firstLoadSheet");
    sortColumnSelectEl.appendChild(opt);
    sortColumnSelectEl.disabled = true;
    if (addSortRuleBtn) addSortRuleBtn.disabled = true;
    return;
  }
  sortColumnSelectEl.disabled = false;
  if (addSortRuleBtn) addSortRuleBtn.disabled = false;
  currentHeaders.forEach((header) => {
    const opt = document.createElement("option");
    opt.value = header;
    opt.textContent = header;
    sortColumnSelectEl.appendChild(opt);
  });
}

function renderSortRules() {
  if (!sortRulesListEl) return;
  sortRulesListEl.replaceChildren();
  if (!multiSortState.length) {
    sortRulesListEl.appendChild(createEmptyInsight(t("sortRulesEmpty")));
    return;
  }
  multiSortState.forEach((rule, index) => {
    const item = document.createElement("div");
    item.className = "sort-rule-item";

    const label = document.createElement("div");
    label.className = "sort-rule-label";
    label.textContent = `${index + 1}. ${rule.col}`;

    const dir = document.createElement("div");
    dir.className = "sort-rule-dir";
    dir.textContent = rule.dir === "asc" ? t("sortAsc") : t("sortDesc");

    const actions = document.createElement("div");
    actions.className = "sort-rule-actions";

    const upBtn = document.createElement("button");
    upBtn.className = "btn ghost btn-sm";
    upBtn.type = "button";
    upBtn.dataset.sortAction = "up";
    upBtn.dataset.sortIndex = String(index);
    upBtn.textContent = t("moveUp");
    upBtn.disabled = index === 0;

    const downBtn = document.createElement("button");
    downBtn.className = "btn ghost btn-sm";
    downBtn.type = "button";
    downBtn.dataset.sortAction = "down";
    downBtn.dataset.sortIndex = String(index);
    downBtn.textContent = t("moveDown");
    downBtn.disabled = index === multiSortState.length - 1;

    const toggleBtn = document.createElement("button");
    toggleBtn.className = "btn ghost btn-sm";
    toggleBtn.type = "button";
    toggleBtn.dataset.sortAction = "toggle";
    toggleBtn.dataset.sortIndex = String(index);
    toggleBtn.textContent = t("changeDirection");

    const removeBtn = document.createElement("button");
    removeBtn.className = "btn ghost btn-sm";
    removeBtn.type = "button";
    removeBtn.dataset.sortAction = "remove";
    removeBtn.dataset.sortIndex = String(index);
    removeBtn.textContent = t("remove");

    actions.appendChild(upBtn);
    actions.appendChild(downBtn);
    actions.appendChild(toggleBtn);
    actions.appendChild(removeBtn);

    item.appendChild(label);
    item.appendChild(dir);
    item.appendChild(actions);
    sortRulesListEl.appendChild(item);
  });
}

function loadSortPresets() {
  try {
    const raw = localStorage.getItem(SORT_PRESETS_KEY);
    const parsed = raw ? JSON.parse(raw) : [];
    return Array.isArray(parsed) ? parsed : [];
  } catch {
    return [];
  }
}

function saveSortPresets(presets) {
  localStorage.setItem(SORT_PRESETS_KEY, JSON.stringify(presets));
}

function renderSortPresets() {
  if (!sortPresetSelectEl) return;
  const presets = loadSortPresets();
  sortPresetSelectEl.replaceChildren();
  if (!presets.length) {
    const opt = document.createElement("option");
    opt.value = "";
    opt.textContent = t("noSavedPresets");
    sortPresetSelectEl.appendChild(opt);
    return;
  }
  const placeholder = document.createElement("option");
  placeholder.value = "";
  placeholder.textContent = t("choosePreset");
  sortPresetSelectEl.appendChild(placeholder);
  presets.forEach((preset) => {
    const opt = document.createElement("option");
    opt.value = preset.name;
    opt.textContent = preset.name;
    sortPresetSelectEl.appendChild(opt);
  });
}

function applyCurrentSort() {
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
  renderFormulaWorkbench();
}

function applyFilters() {
  const criteria = [
    {
      query: (searchQueryEl.value || "").trim().toLowerCase(),
      mode: getNormalizedSelectValue(filterModeEl),
      indexes: resolveIndexes(currentHeaders, columnSelections.filter1),
      emptyMode: getNormalizedSelectValue(filterEmptyModeEl),
      negated: filterNegateEl.checked,
      operatorsEnabled: !!filterOperatorsEl?.checked,
    },
    {
      query: (searchQuery2El.value || "").trim().toLowerCase(),
      mode: getNormalizedSelectValue(filterMode2El),
      indexes: resolveIndexes(currentHeaders, columnSelections.filter2),
      emptyMode: getNormalizedSelectValue(filterEmptyMode2El),
      negated: filterNegate2El.checked,
      operatorsEnabled: !!filterOperators2El?.checked,
    },
  ];

  const dateFilter = {
    indexes: resolveIndexes(currentHeaders, columnSelections.date),
    range: getDateRange(),
    emptyMode: getNormalizedSelectValue(dateEmptyModeEl),
    negated: dateNegateEl.checked,
  };
  const onlyNonEmpty = onlyNonEmptyEl.checked;

  const rowPasses = (row) => {
    if (!rowMatchesTextFilter(row, criteria, onlyNonEmpty)) return false;
    if (!rowMatchesDateFilter(row, dateFilter)) return false;
    return true;
  };

  if (quickSearchHighlightMode) {
    // Tryb "zaznacz": wszystkie wiersze widoczne, zapamiętaj które pasują
    viewRows = baseRows.slice();
    matchedRowIndexes = new Set(baseRows.filter(rowPasses).map((r) => r.rowIndex0));
  } else {
    // Tryb "filtruj": tylko pasujące wiersze
    matchedRowIndexes = new Set();
    viewRows = baseRows.filter(rowPasses);
  }
}

function sortRows() {
  normalizeSortState();
  if (!multiSortState.length) return;
  viewRows.sort((a, b) => {
    for (const rule of multiSortState) {
      const idx = currentHeaders.indexOf(rule.col);
      if (idx < 0) continue;
      const av = a.rawValues ? a.rawValues[idx] : a.values[idx];
      const bv = b.rawValues ? b.rawValues[idx] : b.values[idx];
      const cmp = compareSortValues(av, bv);
      if (cmp !== 0) return rule.dir === "desc" ? -cmp : cmp;
    }
    return 0;
  });
}

function sortRowsForHeaders(rows, headers) {
  normalizeSortState();
  if (!multiSortState.length || !Array.isArray(rows) || !Array.isArray(headers)) return;
  rows.sort((a, b) => {
    for (const rule of multiSortState) {
      const idx = headers.indexOf(rule.col);
      if (idx < 0) continue;
      const av = a.rawValues ? a.rawValues[idx] : a.values[idx];
      const bv = b.rawValues ? b.rawValues[idx] : b.values[idx];
      const cmp = compareSortValues(av, bv);
      if (cmp !== 0) return rule.dir === "desc" ? -cmp : cmp;
    }
    return 0;
  });
}

function getActiveRepeatingGroup() {
  return Array.isArray(currentRepeatingBlocks) && currentRepeatingBlocks.length ? currentRepeatingBlocks[0] : null;
}