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

// Leniwe doładowanie CIĘŻKICH bibliotek (SheetJS CE 0.20.3 ~930K + JSZip ~96K).
// [EN] SheetJS CE 0.20.3 replaces xlsx-js-style (SheetJS 0.18.5 fork) — fixes
// CVE-2023-30533 (prototype pollution) i CVE-2024-22363 (ReDoS) przy parsowaniu plików.
// NIE są ładowane przy starcie (zob. index.html) — apka bootuje bez ~512K JS.
// Dogrywamy je dopiero przy pierwszym wczytaniu/zapisie pliku. Po dograniu
// (i zacache'owaniu przez SW) kolejne wywołania zwracają od razu true.
// Idempotentne: równoległe wywołania współdzielą jedną obietnicę.
let _xlsxLibsPromise = null;
function _loadScriptOnce(src) {
  return new Promise((resolve, reject) => {
    const s = document.createElement("script");
    s.src = src;
    s.async = false; // zachowaj kolejność wykonania (xlsx przed kodem, który go używa)
    s.onload = () => resolve();
    s.onerror = () => reject(new Error("Nie udało się wczytać: " + src));
    document.head.appendChild(s);
  });
}
async function ensureXlsxLibs(showFeedback = false) {
  if (window.XLSX && window.JSZip) return true;
  if (!_xlsxLibsPromise) {
    _xlsxLibsPromise = Promise.all([
      window.XLSX ? null : _loadScriptOnce("lib/xlsx.full.min.js"),
      window.JSZip ? null : _loadScriptOnce("lib/jszip.min.js"),
    ]).catch((e) => {
      _xlsxLibsPromise = null; // pozwól spróbować ponownie przy następnej akcji
      throw e;
    });
  }
  try {
    await _xlsxLibsPromise;
  } catch (e) {
    // np. offline i biblioteki nie ma jeszcze w cache SW
  }
  const ok = !!(window.XLSX && window.JSZip);
  if (!ok && showFeedback) {
    setStatus(t("xlsxMissingStatus"));
    toast(t("xlsxMissingToast"), "error");
    log("Nie udało się doładować bibliotek XLSX/JSZip.", "error");
  }
  return ok;
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
  // Puste pole = wróć do automatu (limit dobierany do szybkości urządzenia).
  if (!String(maxRowsEl.value || "").trim()) {
    localStorage.removeItem(MAX_ROWS_KEY);
    maxRowsEl.value = String(typeof renderBudget !== "undefined"
      ? renderBudget.recommendedRows(currentHeaders.length)
      : 200);
    return;
  }
  const value = Math.max(1, parseInt(maxRowsEl.value || "200", 10));
  localStorage.setItem(MAX_ROWS_KEY, String(value));
}

function loadRowHeightPreference() {
  const saved = parseInt(localStorage.getItem(ROW_HEIGHT_KEY) || "0", 10);
  manualRowHeightAll = Number.isFinite(saved) && saved > 0 ? saved : 0;
  if (rowHeightAllEl && manualRowHeightAll > 0) rowHeightAllEl.value = String(manualRowHeightAll);
}

function applyRowHeightAllPreference() {
  const v = parseInt((rowHeightAllEl && rowHeightAllEl.value) || "0", 10);
  manualRowHeightAll = Number.isFinite(v) && v > 0 ? v : 0;
  localStorage.setItem(ROW_HEIGHT_KEY, String(manualRowHeightAll));
}

function loadColWidthPreference() {
  const saved = parseInt(localStorage.getItem(COL_WIDTH_KEY) || "0", 10);
  manualColWidthAll = Number.isFinite(saved) && saved > 0 ? saved : 0;
  if (colWidthAllEl && manualColWidthAll > 0) colWidthAllEl.value = String(manualColWidthAll);
}

function applyColWidthAllPreference() {
  const v = parseInt((colWidthAllEl && colWidthAllEl.value) || "0", 10);
  manualColWidthAll = Number.isFinite(v) && v > 0 ? v : 0;
  localStorage.setItem(COL_WIDTH_KEY, String(manualColWidthAll));
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

// Wczytuje flagi pokazywania stylów komórek z checkboxów do zmiennych modułu
// (czytanych potem per komórka w applyCellStyle — szybciej niż odpytywać DOM).
function syncCellStyleFlags() {
  cellStyleShowFontColors = !showFontColorsEl || showFontColorsEl.checked;
  cellStyleShowFills = !showCellFillsEl || showCellFillsEl.checked;
  cellStyleShowFonts = !showCellFontsEl || showCellFontsEl.checked;
  cellStyleShowBorders = !showCellBordersEl || showCellBordersEl.checked;
  cellStyleShowConditionalFormatting = !showConditionalFormattingEl || showConditionalFormattingEl.checked;
  cellStyleShowSubheaders = !showSubheadersEl || showSubheadersEl.checked;
  cellStyleSmartWidths = !smartColWidthsEl || smartColWidthsEl.checked;
  recalcDateFormulas = !recalcDatesEl || recalcDatesEl.checked;
}

// Przebudowa danych bieżącego arkusza (np. po zmianie „Przeliczaj formuły z datą")
// BEZ kasowania filtrów/szerokości — odtwarza baseRows i ponownie filtruje/sortuje.
function rebuildCurrentSheetData() {
  if (!workbook || !currentSheetName) return;
  const sheet = workbook.Sheets[currentSheetName];
  if (!sheet) return;
  const data = buildRows(sheet, currentHeaderRow, workbook);
  if (typeof cfmtOverlayRows === "function") cfmtOverlayRows(data.rows);
  baseRows = markSubheaderRows(data.rows);
  applyFilters();
  sortRows();
  renderActiveTable();
}

// Globalny przełącznik zawijania tekstu (klasa na <table>). Trwa przez re-rendery,
// bo renderTable podmienia tylko thead/tbody, nie sam #dataTable.
function applyWrapCells() {
  if (!tableEl) return;
  tableEl.classList.toggle("wrap-cells", !!(wrapCellsEl && wrapCellsEl.checked));
}

function saveCellStylePreferences() {
  localStorage.setItem(CELL_STYLE_PREFS_KEY, JSON.stringify({
    fontColors: cellStyleShowFontColors,
    fills: cellStyleShowFills,
    fonts: cellStyleShowFonts,
    borders: cellStyleShowBorders,
    conditionalFormatting: cellStyleShowConditionalFormatting,
    subheaders: cellStyleShowSubheaders,
    smartWidths: cellStyleSmartWidths,
    recalcDates: recalcDateFormulas,
    wrap: !!(wrapCellsEl && wrapCellsEl.checked),
    freezeFirstCol: !!(freezeFirstColEl && freezeFirstColEl.checked),
  }));
}

function loadCellStylePreferences() {
  let prefs = {};
  try { prefs = JSON.parse(localStorage.getItem(CELL_STYLE_PREFS_KEY) || "{}") || {}; } catch { prefs = {}; }
  const set = (el, val) => { if (el) el.checked = val !== false; }; // brak zapisu = domyślnie włączone
  set(showFontColorsEl, prefs.fontColors);
  set(showCellFillsEl, prefs.fills);
  set(showCellFontsEl, prefs.fonts);
  set(showCellBordersEl, prefs.borders);
  set(showConditionalFormattingEl, prefs.conditionalFormatting);
  set(showSubheadersEl, prefs.subheaders);
  set(smartColWidthsEl, prefs.smartWidths);
  set(recalcDatesEl, prefs.recalcDates);
  if (wrapCellsEl) wrapCellsEl.checked = prefs.wrap === true; // zawijanie domyślnie WYŁĄCZONE
  if (freezeFirstColEl) freezeFirstColEl.checked = prefs.freezeFirstCol === true; // blokada kolumny domyślnie WYŁ.
  syncCellStyleFlags();
  applyWrapCells();
  applyFreezeFirstColumn();
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

// Data ostatniego zapisu skoroszytu — z docProps/core.xml (SheetJS `Props.ModifiedDate`).
// To ona tłumaczy różnicę: Excel zapisał wynik formuły z TODAY() w TAMTYM dniu.
function workbookSavedDateLabel() {
  try {
    const raw = workbook?.Props?.ModifiedDate || workbook?.Props?.CreatedDate;
    if (!raw) return "";
    const d = raw instanceof Date ? raw : new Date(raw);
    if (Number.isNaN(d.getTime())) return "";
    return formatLocalizedDateDisplay(d, { month: "short", year: "numeric" });
  } catch {
    return "";
  }
}

// Treść notki dla komórki oznaczonej jako przeliczona. Musi podawać OBIE liczby —
// samo „przeliczono" nie odpowiada na pytanie „to czemu w Excelu widzę co innego".
function recalcNoteForCell(cell) {
  if (!cell || !cell.classList || !cell.classList.contains("cell-recalced")) return "";
  const was = String(cell.dataset.recalcWas ?? "").trim() || "—";
  const today = formatLocalizedDateDisplay(new Date(), { month: "short", year: "numeric" });
  const saved = workbookSavedDateLabel();
  return saved
    ? t("recalcNoteSaved", { today, was, saved })
    : t("recalcNote", { today, was });
}

function showCellTooltip(cell, persistent = false) {
  if (!cellTooltipEl || !cell) return;
  const note = recalcNoteForCell(cell);
  // Notatka komórki (cell-notes.js) — jak w Excelu: widać ją po najechaniu / tapnięciu.
  const userNote = typeof swbNoteTextForTd === "function" ? swbNoteTextForTd(cell) : "";
  const truncated = isCellTextTruncated(cell);
  // Komórka przeliczona pokazuje notkę ZAWSZE, także gdy tekst się mieści —
  // inaczej znacznik w rogu byłby zagadką bez odpowiedzi.
  if (!note && !truncated && !userNote) return;
  const text = truncated ? getTooltipText(cell) : "";
  if (!text && !note && !userNote) return;
  hideCellTooltip();
  const parts = [];
  if (text) {
    const main = document.createElement("div");
    main.textContent = text;
    parts.push(main);
  }
  if (userNote) {
    const noteEl = document.createElement("div");
    noteEl.className = "cell-tooltip-usernote";
    noteEl.textContent = userNote;
    parts.push(noteEl);
  }
  if (note) {
    const noteEl = document.createElement("div");
    noteEl.className = "cell-tooltip-note";
    noteEl.textContent = note;
    parts.push(noteEl);
  }
  cellTooltipEl.replaceChildren(...parts);
  cellTooltipEl.classList.remove("hidden");
  positionCellTooltip(cell);
  if (!persistent) return;
  // notatkę trzeba zdążyć przeczytać — dłużej przy dłuższej treści (do 8 s)
  const holdMs = userNote ? Math.min(8000, 3000 + userNote.length * 40) : 2200;
  tooltipHideTimer = window.setTimeout(() => {
    hideCellTooltip();
  }, holdMs);
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

// ── Tekst komórki jak w polskim Excelu ──────────────────────────────────────────
// SheetJS (SSF 0.20.3) rzuca wyjątek na kropce w formacie daty — a „dd.mm.yyyy” to
// najczęstszy format w polskich plikach. Komórka zostawała bez tekstu (w), widok
// pokazywał zapasowe „12-03-25”, a data z godziną („dd.mm.yyyy hh:mm”) gubiła godzinę.
// Kropka poza cudzysłowem i [ ], nie po 0 # ? (separator dziesiętny) i nie przed 0
// (ułamek sekundy „ss.000”) jest w dacie zwykłym znakiem — escapujemy ją i próbujemy jeszcze raz.
function swbEscapeFormatDots(code) {
  let out = "";
  let quoted = false;
  let bracket = false;
  for (let i = 0; i < code.length; i++) {
    const ch = code[i];
    if (ch === "\\" && !quoted) { out += ch + (code[i + 1] ?? ""); i++; continue; }
    if (ch === '"') quoted = !quoted;
    else if (!quoted && ch === "[") bracket = true;
    else if (!quoted && ch === "]") bracket = false;
    else if (ch === "." && !quoted && !bracket && !/[0#?]/.test(code[i - 1] || "") && code[i + 1] !== "0") {
      out += "\\.";
      continue;
    }
    out += ch;
  }
  return out;
}

function swbSsfFormat(code, value) {
  const ssf = typeof XLSX !== "undefined" && XLSX.SSF;
  if (!ssf || code == null) return null;
  // SSF bierze Date jako UTC, a daty w aplikacji to lokalny czas „ścienny” (swbLocalizeFileDates)
  // — do SSF idzie liczba seryjna z lokalnych składowych, jak przy zapisie (xlsx-patch).
  if (value instanceof Date) {
    if (Number.isNaN(value.getTime())) return null;
    value = (Date.UTC(value.getFullYear(), value.getMonth(), value.getDate(), value.getHours(), value.getMinutes(), value.getSeconds(), value.getMilliseconds()) - Date.UTC(1899, 11, 30)) / 86400000;
  }
  try { return ssf.format(code, value); } catch { /* druga próba niżej */ }
  const fixed = swbEscapeFormatDots(String(code));
  if (fixed === code) return null;
  try { return ssf.format(fixed, value); } catch { return null; }
}

// SheetJS (0.20.3, cellDates) CZYTA liczbę seryjną jako północ UTC (46030 → 8.01 01:00 w Polsce,
// latem 02:00), a ZAPISUJE i liczy daty z lokalnych składowych. Reszta aplikacji (zapis
// xlsx-patch, silnik formuł, pasek formuły, eksport) też liczy lokalnie — więc data z pliku
// pokazywała na pasku „2026-10-15 02:00”, =B3=DATA(2026;1;8) dawało FAŁSZ, a skopiowana /
// wypełniona / poprawiona data szła do pliku z ukrytą godziną (film Mateusza 2026-10-25).
// Tu, raz przy wczytaniu: data z pliku = ten sam czas „ścienny” w strefie lokalnej.
function swbLocalizeFileDates(wb) {
  if (!wb || !wb.Sheets) return 0;
  let fixed = 0;
  for (const name of wb.SheetNames || Object.keys(wb.Sheets)) {
    const sheet = wb.Sheets[name];
    if (!sheet) continue;
    for (const ref in sheet) {
      if (ref.charCodeAt(0) === 33) continue;
      const cell = sheet[ref];
      const v = cell && cell.v;
      if (!(v instanceof Date) || Number.isNaN(v.getTime())) continue;
      cell.v = new Date(v.getUTCFullYear(), v.getUTCMonth(), v.getUTCDate(), v.getUTCHours(), v.getUTCMinutes(), v.getUTCSeconds(), v.getUTCMilliseconds());
      fixed += 1;
    }
  }
  return fixed;
}

// Po XLSX.read: uzupełnij tekst komórek, których SheetJS nie umiał sformatować.
function swbFillMissingCellText(wb) {
  if (!wb || !wb.Sheets) return 0;
  let filled = 0;
  for (const name of wb.SheetNames || Object.keys(wb.Sheets)) {
    const sheet = wb.Sheets[name];
    if (!sheet) continue;
    for (const ref in sheet) {
      if (ref.charCodeAt(0) === 33) continue; // „!ref”, „!merges”…
      const cell = sheet[ref];
      if (!cell || cell.w !== undefined || cell.v == null || typeof cell.z !== "string") continue;
      if (cell.t !== "n" && cell.t !== "d") continue;
      const w = swbSsfFormat(cell.z, cell.v);
      if (w != null) { cell.w = w; filled += 1; }
    }
  }
  return filled;
}

function swbIsDateFormat(code) {
  if (typeof code !== "string" || !code || code === "General") return false;
  try { return !!(XLSX.SSF && XLSX.SSF.is_date(code)); } catch { return /[dmyhs]/i.test(code.replace(/"[^"]*"|\[[^\]]*\]/g, "")); }
}

// SSF formatuje liczby po amerykańsku („1,234.50”, „12.5%”, TRUE), polski Excel pokazuje
// „1 234,50”, „12,5%”, PRAWDA. Separatory zamieniamy TYLKO w ciągach cyfr (tekst z formatu,
// np. „szt.”, zostaje) i tylko dla liczb w formacie innym niż data (tam kropka to separator daty).
// Błędy: plik i SheetJS trzymają angielskie kody (#DIV/0!), polski Excel pokazuje
// #DZIEL/0! (sprawdzone przez AppleScript `formula local`, 2026-10-17). Tłumaczymy TYLKO
// tekst w widoku — cell.w/cell.v zostają angielskie dla formuł i zapisu (xlsx-patch).
// Kopia w build-rows-core.js (wątek roboczy) — trzymać obie w zgodzie.
const SWB_ERR_PL = {
  "#NULL!": "#ZERO!", "#DIV/0!": "#DZIEL/0!", "#VALUE!": "#ARG!", "#REF!": "#ADR!",
  "#NAME?": "#NAZWA?", "#NUM!": "#LICZBA!", "#N/A": "#N/D", "#SPILL!": "#ROZLANIE!", "#CALC!": "#OBL!",
};
// Nazwy miesięcy i dni z SSF są angielskie („29 Oct 26”, „Wednesday”), polski Excel pokazuje
// „29 paź 26”, „środa”. Daty z PLIKU przechodziły przez localizeDisplayedDate, ale wartość
// WPISANA albo WYPEŁNIONA serią szła prosto z formatu (cfmtFormatValue) i zostawała po
// angielsku — film z iPhone'a 2026-10-31: „29 Oct 26” obok „17 wrz 26”. Tłumaczymy słowa
// w gotowym tekście, więc układ formatu (kolejność, separatory, rok) zostaje nietknięty.
// Jawny język formatu inny niż polski ([$-409] = angielski) Excel PL też pokazuje w nim.
const SWB_DATE_WORDS_EN = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December",
  "Jan", "Feb", "Mar", "Apr", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec",
  "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday",
  "Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"];
const SWB_DATE_WORDS_PL = ["styczeń", "luty", "marzec", "kwiecień", "maj", "czerwiec", "lipiec", "sierpień", "wrzesień", "październik", "listopad", "grudzień",
  "sty", "lut", "mar", "kwi", "cze", "lip", "sie", "wrz", "paź", "lis", "gru",
  "poniedziałek", "wtorek", "środa", "czwartek", "piątek", "sobota", "niedziela",
  "pon", "wt", "śr", "czw", "pt", "sob", "niedz"];
const SWB_DATE_WORDS_MAP = new Map(SWB_DATE_WORDS_EN.map((w, i) => [w, SWB_DATE_WORDS_PL[i]]));
const SWB_DATE_WORDS_RE = new RegExp(`\\b(${SWB_DATE_WORDS_EN.join("|")})\\b`, "g");
function swbLocalizeDateWords(shown, code) {
  if (typeof shown !== "string" || !/[A-Za-z]/.test(shown)) return shown;
  const lcid = /\[\$[^\]]*-([0-9a-f]+)\]/i.exec(String(code || ""));
  if (lcid && (parseInt(lcid[1], 16) & 0x3ff) !== 0x15) return shown;
  return shown.replace(SWB_DATE_WORDS_RE, (w) => SWB_DATE_WORDS_MAP.get(w) || w);
}
function swbLocalizeCellText(shown, cell, lang) {
  if (lang === "en" || typeof shown !== "string" || !cell) return shown;
  if (cell.t === "b") return cell.v ? "PRAWDA" : "FAŁSZ";
  if (cell.t === "e") return SWB_ERR_PL[shown] || shown;
  if ((cell.t === "n" || cell.t === "d") && swbIsDateFormat(cell.z)) return swbLocalizeDateWords(shown, cell.z);
  if (cell.t !== "n" || !/\d[.,]\d/.test(shown)) return shown;
  if (swbIsDateFormat(cell.z)) return shown;
  return shown.replace(/\d{1,3}(?:,\d{3})+(?:\.\d+)?|\d*\.\d+/g, (tok) => tok.replace(/,/g, " ").replace(".", ","));
}

// Wartość bez formatu (Ogólny) — jak pokazuje ją polski Excel: 12,5 zamiast 12.5,
// 0,3 zamiast 0.30000000000000004, PRAWDA zamiast true.
function swbGeneralText(value) {
  if (currentLang === "en") {
    if (typeof value === "number" && Number.isFinite(value)) return String(Number(value.toPrecision(15)));
    if (typeof value === "boolean") return value ? "TRUE" : "FALSE";
    return toDisplay(value);
  }
  if (typeof value === "number" && Number.isFinite(value)) return String(Number(value.toPrecision(15))).replace(".", ",");
  if (typeof value === "boolean") return value ? "PRAWDA" : "FAŁSZ";
  return toDisplay(value);
}

// Tworzenie Intl.DateTimeFormat jest DROGIE (buduje reguły lokalizacji), a tu leciało
// po jednym na każdą komórkę z datą — przy wczytywaniu arkusza 563×33 dało to 420 ms
// z 3,5 s całego profilu CPU. Formatery zależą wyłącznie od (język, zestaw opcji),
// więc trzymamy je w mapie; zestawów jest kilka, więc mapa nie rośnie.
const _dateFormatterCache = new Map();
function getCachedDateFormatter(locale, opts) {
  const key = `${locale}|${opts.month || ""}|${opts.day || ""}|${opts.year || ""}|${opts.weekday || ""}|${opts.hour || ""}|${opts.minute || ""}|${opts.second || ""}`;
  let fmt = _dateFormatterCache.get(key);
  if (!fmt) {
    fmt = new Intl.DateTimeFormat(locale, opts);
    _dateFormatterCache.set(key, fmt);
  }
  return fmt;
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

  return getCachedDateFormatter(I18N[currentLang]?.locale || "pl-PL", formatterOptions)
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

// Mapa nazw miesięcy (PL+EN) — stała modułu. Wcześniej ten ~50-kluczowy obiekt
// powstawał od nowa przy KAŻDYM wywołaniu parseDateFlexible, czyli setki tysięcy
// razy na jedno przeliczenie analiz.
const DATE_MONTH_MAP = {
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
  "march": 3,
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

// Parsowanie dat jest czyste (ten sam tekst -> ta sama data), a w całej aplikacji
// nic nie mutuje zwróconego obiektu Date, więc cache'ujemy znacznik czasu i
// odtwarzamy z niego świeży Date — bez współdzielenia instancji między wywołaniami.
const _dateParseCache = new Map();
const DATE_PARSE_CACHE_LIMIT = 20000;

// Numer seryjny Excela to czas „ścienny" (bez strefy): 45369 = 18.03.2024 00:00.
// Budujemy LOKALNĄ datę z komponentów UTC — dawniej new Date(ms) dawał północ UTC,
// czyli 01:00/02:00 w Polsce, a przy zapisie komórka dostawała ułamek (45369,0417).
function excelSerialToLocalDate(serial) {
  const u = new Date(Math.round((serial - 25569) * 86400000));
  if (Number.isNaN(u.getTime())) return null;
  return new Date(u.getUTCFullYear(), u.getUTCMonth(), u.getUTCDate(),
    u.getUTCHours(), u.getUTCMinutes(), u.getUTCSeconds());
}

// Data z komponentów TYLKO gdy istnieje naprawdę — new Date(2024, 12, 1) „przelewa się"
// na 2025-01-01, więc „01.13.2024” czy „31.02.24” dawały inną, prawdziwą datę.
function makeValidDate(y, m, d) {
  if (m < 1 || m > 12 || d < 1 || d > 31) return null;
  const dt = new Date(y, m - 1, d);
  return dt.getFullYear() === y && dt.getMonth() === m - 1 && dt.getDate() === d ? dt : null;
}

function parseDateFlexible(value) {
  if (value instanceof Date) return value;
  if (typeof value === "number" && Number.isFinite(value)) return excelSerialToLocalDate(value);
  if (typeof value !== "string") return null;
  let v = value.trim();
  if (!v) return null;

  // Klucz cache bierzemy PRZED normalizacją `v` (niżej obcinane są godziny/sufiks T),
  // żeby zapis i odczyt trafiały w to samo miejsce.
  const cacheKey = v;
  const cached = _dateParseCache.get(cacheKey);
  if (cached !== undefined) return cached === null ? null : new Date(cached);

  if (/^\d+(\.\d+)?$/.test(v)) {
    const numeric = Number(v);
    if (Number.isFinite(numeric)) return excelSerialToLocalDate(numeric);
  }

  // UWAGA: dawniej /T.*$/ (bez kotwiczenia do cyfry) obcinał WSZYSTKO od pierwszej
  // wielkiej litery "T" w stringu — więc zwykły tekst typu "Teren C-03" albo
  // "Total 5" tracił resztę treści i mylnie przechodził dalej jako data (przez
  // pustą resztę i domyślny rok silnika JS). Kotwiczymy sufiks czasu ISO do cyfry
  // bezpośrednio przed "T" (np. "2026-01-01T10:30:00" → "2026-01-01").
  v = v.replace(/(\d)T\d.*$/, "$1");
  v = v.replace(/\s+\d{1,2}:\d{2}(:\d{2})?.*$/, "");
  const normalized = v.replace(/[.\/]/g, "-");


  const memo = (d) => {
    if (_dateParseCache.size >= DATE_PARSE_CACHE_LIMIT) _dateParseCache.clear();
    _dateParseCache.set(cacheKey, d ? d.getTime() : null);
    return d;
  };

  let m = normalized.match(/^(\d{1,2})-(\d{1,2})-(\d{2}|\d{4})$/);
  if (m) {
    const y = m[3].length === 2 ? Number(`20${m[3]}`) : Number(m[3]);
    return memo(makeValidDate(y, Number(m[2]), Number(m[1])));
  }

  m = normalized.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (m) {
    return memo(makeValidDate(Number(m[1]), Number(m[2]), Number(m[3])));
  }

  const words = v.toLowerCase()
    .replace(/,/g, "")
    .replace(/\s+/g, " ")
    .trim();
  let wm = words.match(/^(\d{1,2})\s+([a-ząćęłńóśźż\.]+)\s+(\d{4}|\d{2})$/i);
  if (wm) {
    const day = Number(wm[1]);
    const monthKey = wm[2].replace(/\.$/, "");
    const month = DATE_MONTH_MAP[monthKey];
    const year = wm[3].length === 2 ? Number(`20${wm[3]}`) : Number(wm[3]);
    if (month) return memo(makeValidDate(year, month, day));
  }
  wm = words.match(/^([a-ząćęłńóśźż\.]+)\s+(\d{1,2})\s+(\d{4}|\d{2})$/i);
  if (wm) {
    const monthKey = wm[1].replace(/\.$/, "");
    const month = DATE_MONTH_MAP[monthKey];
    const day = Number(wm[2]);
    const year = wm[3].length === 2 ? Number(`20${wm[3]}`) : Number(wm[3]);
    if (month) return memo(makeValidDate(year, month, day));
  }

  // Fallback: natywny Date.parse jest skrajnie pobłażliwy — z adresu "Mariańska 27, 29"
  // wyłuska 2029-03-27, przez co teksty (nazwy ulic, etykiety) trafiają do filtra dat.
  // Pozwalamy na niego tylko, gdy każde słowo w wartości jest nazwą miesiąca
  // (formaty słowne typu "March 27, 2029" / "7 czerwca 2026"); inaczej → to nie data.
  const alphaTokens = v.toLowerCase().match(/[a-ząćęłńóśźż]+/g) || [];
  // bez ani jednego słowa nie ma czego tłumaczyć — Date.parse z „12,5” robił 2001-12-05
  const onlyMonthWords = alphaTokens.length > 0 && alphaTokens.every(tok => {
    const key = tok.replace(/\.$/, "");
    return DATE_MONTH_MAP[key] != null;
  });
  if (onlyMonthWords) {
    const parsed = Date.parse(value);
    if (!Number.isNaN(parsed)) return memo(new Date(parsed));
  }
  return memo(null);
}

// Tekst wpisany/wklejony do komórki → wartość, jak polski Excel. Wzorzec: wyniki z Excela
// (formula local przez AppleScript, 2026-10-17) i parser OnlyOffice (sdkjs FormatParser.parse).
// KOLEJNOŚĆ MA ZNACZENIE: najpierw liczba, potem data. Wcześniej data szła pierwsza,
// a parseDateFlexible traktuje gołą liczbę jak numer seryjny daty Excela — wpisane „42”
// zapisywało się jako 10.02.1900 (także do pliku), „12,5” jako data z 2001 r.
// Wynik: { value, type: "number"|"date"|"boolean"|"string"|"formula", numFmt? } albo null
// (pusto). numFmt = format, który Excel nadaje komórce w formacie Ogólny („12%” → 0%,
// „8:30” → h:mm, „1 234” → #,##0); komórka z własnym formatem go zachowuje (updateSheetCell).
// opts.numFmt = bieżący format komórki: „@” (Tekst) = wszystko zostaje tekstem, jak w Excelu.
//
// Świadome odstępstwa od Excela: „12.5” zostaje liczbą (Excel robi z tego 12 maja),
// „500 600 700” zostaje tekstem (numer telefonu, Excel daje liczbę), „1.234,5” jest liczbą.
const SWB_TYPED_ERRORS = ["#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A"];
function swbTypedErrorCode(text, en) {
  const up = String(text).trim().toUpperCase();
  return SWB_TYPED_ERRORS.find((code) => (en ? code : SWB_ERR_PL[code]) === up) || null;
}

function parseInputValue(raw, opts = {}) {
  const text = String(raw ?? "").trim();
  if (!text) return null;
  // Apostrof na początku = „zapisz jako tekst”: '007 → „007”, '=A1 → tekst „=A1”.
  // Wcześniej apostrof trafiał do komórki razem z treścią.
  if (text[0] === "'") {
    const s = text.slice(1);
    return s ? { value: s, type: "string" } : null;
  }
  if (opts.numFmt === "@") return { value: text, type: "string" };
  if (text.startsWith("=")) return { value: text, type: "formula" };
  const en = (opts.lang || (typeof currentLang === "string" ? currentLang : "pl")) === "en";

  const low = text.toLowerCase();
  if (low === (en ? "true" : "prawda")) return { value: true, type: "boolean" };
  if (low === (en ? "false" : "fałsz")) return { value: false, type: "boolean" };

  // Wpisany błąd: polski Excel bierze „#DZIEL/0!”, „#n/d”… (bez względu na wielkość liter)
  // za błąd, a angielskie „#DIV/0!” za zwykły tekst; #OBL!/#ROZLANIE! wpisać się nie da
  // (wyrocznia, 2026-10-17). value = angielski kod — taki idzie do pliku (t="e").
  if (text[0] === "#") {
    const code = swbTypedErrorCode(text, en);
    if (code) return { value: code, type: "error" };
  }

  const num = parseTypedNumberInfo(text);
  if (num) return withInputFmt(num.n, "number", num.grouped ? (num.frac ? "#,##0.00" : "#,##0") : null);

  // Ułamek z częścią całkowitą: „1 1/2” → 1,5 w formacie „# ?/?” (mianownik 2+ cyfr → „# ??/??”),
  // „0 3/4” → 0,75. Bez całości („1/2”) to data, jak w Excelu; „-3/4” i „1  1/2” zostają tekstem.
  const frac = parseTypedFraction(text);
  if (frac) return withInputFmt(frac.n, "number", frac.fmt);

  // (100) = −100; 12% = 0,12; 12 zł = 12 w formacie złotowym — także razem: (12 zł), -12%
  const paren = /^\((.+)\)$/.exec(text);
  let body = paren ? paren[1].trim() : text;
  let kind = null;
  let m = /^(.+?)\s?%$/.exec(body);
  if (m) { kind = "percent"; body = m[1]; }
  else if ((m = /^(.+?)\s?zł$/i.exec(body))) { kind = "pln"; body = m[1]; }
  if ((kind || paren) && !(paren && /^[-+]/.test(body))) {
    // „1 1/2%” → 0,015 (0,00%), „1 1/2 zł” → 1,5 w formacie ułamka — tak robi Excel
    const fr = kind && !paren ? parseTypedFraction(body) : null;
    if (fr) {
      return kind === "percent"
        ? withInputFmt(Number((fr.n / 100).toPrecision(15)), "number", "0.00%")
        : withInputFmt(fr.n, "number", fr.fmt);
    }
    const info = parseTypedNumberInfo(body);
    if (info) {
      const n = paren ? -info.n : info.n;
      if (kind === "percent") return withInputFmt(Number((n / 100).toPrecision(15)), "number", info.frac ? "0.00%" : "0%");
      if (kind === "pln") return withInputFmt(n, "number", info.frac ? '#,##0.00 "zł"' : '#,##0 "zł"');
      return withInputFmt(n, "number", info.grouped ? (info.frac ? "#,##0.00" : "#,##0") : null);
    }
  }

  // 1E5, 1,5E3 → zapis naukowy
  m = /^([+-]?(?:\d+(?:[.,]\d*)?|[.,]\d+))[eE]([+-]?\d{1,3})$/.exec(text);
  if (m) {
    const n = Number(`${m[1].replace(",", ".")}e${m[2]}`);
    if (Number.isFinite(n)) return withInputFmt(n, "number", "0.00E+00");
  }

  // 8:30, 8:30:15, 25:30, 8:30 PM → ułamek doby (Excel: liczba w formacie godziny)
  const time = parseTypedTime(text);
  if (time) return withInputFmt(time.value, "number", time.fmt);

  // data z godziną: 12.03.2025 10:30, 2025-03-12 10:30, 2025-03-12T10:30
  // (parseDateFlexible obcina godzinę — wpisana data z godziną traciła czas)
  m = /^(.+?)(?:\s+|T)(\d{1,2}:\d{1,2}(?::\d{1,2})?)$/.exec(text);
  if (m && /\d[-./]\d/.test(m[1])) {
    const d = parseDateFlexible(m[1]);
    const tm = parseTypedTime(m[2]);
    if (d && tm && tm.value < 1) {
      const secs = Math.round(tm.value * 86400);
      const dt = new Date(d.getFullYear(), d.getMonth(), d.getDate(), Math.floor(secs / 3600), Math.floor(secs / 60) % 60, secs % 60);
      const fmt = en ? (tm.seconds ? "yyyy-mm-dd hh:mm:ss" : "yyyy-mm-dd hh:mm") : (tm.seconds ? "dd.mm.yyyy hh:mm:ss" : "dd.mm.yyyy hh:mm");
      return withInputFmt(dt, "date", fmt);
    }
  }

  const noYear = parseTypedPartialDate(text);
  if (noYear) return { value: noYear, type: "date" };

  // data tylko z wyraźnym zapisem daty: cyfry z separatorem - . / albo nazwa miesiąca
  // (12-mar-2025 i 12.mar.2025 → jak „12 mar 2025”)
  if (/\d[-./]\d|[a-ząćęłńóśźż]/i.test(text)) {
    const asDate = parseDateFlexible(text.replace(/(\d)[-./](?=[a-ząćęłńóśźż])|([a-ząćęłńóśźż])[-./](?=\d)/gi, "$1$2 "));
    if (asDate) return { value: asDate, type: "date" };
  }
  return { value: text, type: "string" };
}

// „całość licznik/mianownik” z JEDNĄ spacją (sprawdzone w Excelu: dwie spacje = tekst,
// mianownik 0 = tekst, licznik większy od mianownika wolno: „1 10/3” → 4,333…).
function parseTypedFraction(text) {
  const m = /^([+-]?)(\d{1,15}) (\d{1,15})\/(\d{1,15})$/.exec(text);
  if (!m) return null;
  const den = Number(m[4]);
  if (!den) return null;
  const n = Number((Number(m[2]) + Number(m[3]) / den).toPrecision(15));
  return { n: m[1] === "-" ? -n : n, fmt: m[4].length > 1 ? "# ??/??" : "# ?/?" };
}

function withInputFmt(value, type, numFmt) {
  return numFmt ? { value, type, numFmt } : { value, type };
}

// Mianownik nazw miesięcy (marzec 2025) — TYLKO dla wpisywania. DATE_MONTH_MAP celowo ich
// nie ma: słowa „Maj”, „Lipiec” w danych nie mają się zamieniać w daty w analizach.
const INPUT_MONTH_NOMINATIVE = {
  "styczeń": 1, "styczen": 1, "luty": 2, "marzec": 3, "kwiecień": 4, "kwiecien": 4, "czerwiec": 6,
  "lipiec": 7, "sierpień": 8, "sierpien": 8, "wrzesień": 9, "wrzesien": 9, "październik": 10,
  "pazdziernik": 10, "listopad": 11, "grudzień": 12, "grudzien": 12,
};

function inputMonthNumber(word) {
  const key = String(word).toLowerCase().replace(/\.$/, "");
  return DATE_MONTH_MAP[key] || INPUT_MONTH_NOMINATIVE[key] || null;
}

// Daty bez roku albo bez dnia, jak w Excelu: 12-03 i 1/2 → bieżący rok, 12 mar → bieżący rok,
// 3/2025, mar 2025, marzec 2025, mar 25 → pierwszy dzień miesiąca. Z kropką („12.03”) nie —
// to zostaje liczbą (patrz parseInputValue).
function parseTypedPartialDate(text) {
  const year = new Date().getFullYear();
  let m = /^(\d{1,2})[-/](\d{1,2})$/.exec(text);
  if (m) return makeValidDate(year, Number(m[2]), Number(m[1]));
  m = /^(\d{1,2})[-/](\d{4})$/.exec(text);
  if (m) return makeValidDate(Number(m[2]), Number(m[1]), 1);
  m = /^(\d{1,2})[\s-]+([a-ząćęłńóśźż]+\.?)$/i.exec(text);
  if (m && inputMonthNumber(m[2])) return makeValidDate(year, inputMonthNumber(m[2]), Number(m[1]));
  m = /^([a-ząćęłńóśźż]+\.?)[\s-]+(\d{2}|\d{4})$/i.exec(text);
  if (m && inputMonthNumber(m[1])) return makeValidDate(m[2].length === 2 ? 2000 + Number(m[2]) : Number(m[2]), inputMonthNumber(m[1]), 1);
  return null;
}

// Godzina → { value: ułamek doby, fmt, seconds }. Godziny do 9999 (25:30 = 1,0625 doby,
// format [h]:mm:ss jak w Excelu), minuty i sekundy < 60, AM/PM tylko z godzinami 0–12.
function parseTypedTime(text) {
  const m = /^(\d{1,4}):(\d{1,2})(?::(\d{1,2}))?(?:\s*([AaPp])\.?[Mm]\.?)?$/.exec(String(text).trim());
  if (!m) return null;
  let h = Number(m[1]);
  const min = Number(m[2]);
  const sec = m[3] != null ? Number(m[3]) : 0;
  if (min > 59 || sec > 59) return null;
  const ampm = m[4] ? m[4].toUpperCase() : null;
  if (ampm) {
    if (h > 12) return null;
    if (ampm === "P" && h < 12) h += 12;
    if (ampm === "A" && h === 12) h = 0;
  }
  const value = (h * 3600 + min * 60 + sec) / 86400;
  const seconds = m[3] != null;
  let fmt;
  if (ampm) fmt = seconds ? "h:mm:ss AM/PM" : "h:mm AM/PM";
  else if (h >= 24) fmt = "[h]:mm:ss";
  else fmt = seconds ? "h:mm:ss" : "h:mm";
  return { value, fmt, seconds };
}

// Liczba w zapisie polskim lub angielskim: 42, -3,75, 12.5, „5,” (jak Excel = 5), 1 234,50
// (spacja/twarda spacja jako separator tysięcy), 1.234,5. Nie liczby: „1,5,7”, „12 zł”,
// „500 600 700”; kody z zerem wiodącym zostają liczbami jak w Excelu (007 → 7).
// grouped = użyto separatora tysięcy (Excel nadaje wtedy format #,##0), frac = część dziesiętna.
function parseTypedNumberInfo(text) {
  let s = String(text).trim().replace(/^\+/, "");
  let grouped = false;
  // odstęp tysięcy: „1 234” albo z częścią dziesiętną; „500 600 700” (telefon) zostaje tekstem
  if (/^-?\d{1,3}[   ]\d{3}([.,]\d+)?$/.test(s) || /^-?\d{1,3}([   ]\d{3})+[.,]\d+$/.test(s)) {
    s = s.replace(/[   ]/g, "");
    grouped = true;
  } else if (/^-?\d{1,3}(\.\d{3})+,\d+$/.test(s)) {
    s = s.replace(/\./g, "");
    grouped = true;
  }
  // „5,” = 5 jak w Excelu; „1.” zostaje tekstem (numeracja „1.”, „2.” w danych)
  if (!/^-?\d+(,\d*|\.\d+)?$/.test(s) && !/^-?[.,]\d+$/.test(s)) return null;
  const n = Number(s.replace(",", "."));
  return Number.isFinite(n) ? { n, grouped, frac: /[.,]\d/.test(s) } : null;
}

function parseTypedNumber(text) {
  const info = parseTypedNumberInfo(text);
  return info ? info.n : null;
}

function updateSheetCell(rowIndex0, colIndex0, parsed) {
  if (!workbook || !currentSheetName) return;
  const sheet = workbook.Sheets[currentSheetName];
  if (!sheet) return;
  // Kolumna wyliczana (wirtualna, dopisana na końcu) nie istnieje w pliku — zapis trafiłby
  // do prawdziwej kolumny za danymi.
  if (typeof cfmtIsRealCol === "function" && !cfmtIsRealCol(colIndex0)) return;
  const absoluteCol = currentStartCol + colIndex0;
  const cellRef = XLSX.utils.encode_cell({ r: rowIndex0, c: absoluteCol });
  if (!parsed || parsed.value === null) {
    if (typeof undoCaptureCell === "function") undoCaptureCell(rowIndex0, colIndex0); // stan „przed" do Cofnij
    const prevCell = sheet[cellRef];
    const hadFormula = !!(prevCell && prevCell.f);
    // „Wyczyść zawartość” jak w Excelu: znika wartość/formuła, FORMAT ZOSTAJE. Wcześniej
    // komórka znikała w całości — zapis i tak trzymał styl (xlsx-patch), ale w aplikacji
    // pasek pokazywał „Ogólny”, a nowa wartość wpisana w to miejsce traciła format (np.
    // data jako liczba) aż do ponownego otwarcia pliku. Pusta komórka-zaślepka (t:"z")
    // jest dla formuł i wykrywania zakresu pusta.
    // Notatka (.c) też zostaje — Delete w Excelu nie kasuje notatek.
    if (prevCell && (prevCell.s || prevCell.z || prevCell.c)) {
      const stub = { t: "z" };
      if (prevCell.s) stub.s = prevCell.s;
      if (prevCell.z) stub.z = prevCell.z;
      if (prevCell.c) stub.c = prevCell.c;
      sheet[cellRef] = stub;
    } else {
      delete sheet[cellRef];
    }
    bumpSheetDataStamp(); // zawartość arkusza się zmieniła -> unieważnij cache'e
    recordPendingEdit(currentSheetName, cellRef, null); // null = usunięcie przy zapisie
    if (hadFormula && typeof swbFormulaCellChanged === "function") swbFormulaCellChanged(currentSheetName, cellRef);
    if (typeof swbFormulaTouched === "function") swbFormulaTouched(currentSheetName, rowIndex0, absoluteCol);
    return;
  }
  if (parsed.type === "formula") {
    toast(t("formulaEditBlocked"), "warning");
    return;
  }
  // Zachowaj styl (.s) i format liczbowy (.z) z poprzedniej komórki, jeśli były —
  // edycja samej wartości nie powinna gubić formatowania odczytanego z pliku.
  if (typeof undoCaptureCell === "function") undoCaptureCell(rowIndex0, colIndex0); // stan „przed" do Cofnij
  const prev = sheet[cellRef];
  const cell = {};
  if (prev && prev.s) cell.s = prev.s;
  if (prev && prev.z) cell.z = prev.z;
  if (prev && prev.c) cell.c = prev.c; // notatka komórki (cell-notes.js)
  // Format jak w Excelu: „12%” w komórce Ogólny dostaje 0%, „8:30” — h:mm, „1 234” — #,##0.
  // Komórka z własnym formatem go zachowuje (Excel też: „12%” w formacie 0.00 daje 0,12).
  const autoFmt = parsed.numFmt && typeof cfmtNumFmtAt === "function"
    && cfmtNumFmtAt(rowIndex0, colIndex0) === "General" ? parsed.numFmt : null;
  if (autoFmt) cell.z = autoFmt;
  if (parsed.type === "date") {
    cell.v = parsed.value;
    cell.t = "d";
  } else if (parsed.type === "number") {
    cell.v = parsed.value;
    cell.t = "n";
  } else if (parsed.type === "boolean") {
    cell.v = !!parsed.value;
    cell.t = "b";
  } else if (parsed.type === "error") {
    // jak wynik formuły (fcStore): v i w = angielski kod, widok tłumaczy swbLocalizeCellText
    cell.v = parsed.value;
    cell.w = parsed.value;
    cell.t = "e";
  } else {
    cell.v = parsed.value;
    cell.t = "s";
  }
  sheet[cellRef] = cell;
  swbEnsureSheetRef(sheet, rowIndex0, absoluteCol);
  bumpSheetDataStamp(); // zawartość arkusza się zmieniła -> unieważnij cache'e
  // Zarejestruj edycję do naniesienia metodą ZIP-patch przy zapisie (zachowuje plik).
  recordPendingEdit(currentSheetName, cellRef, cell.t === "e" ? { t: "e", err: cell.v } : { v: cell.v, t: cell.t });
  if (autoFmt && typeof cfmtRecord === "function") cfmtRecord(currentSheetName, cellRef, { numFmt: autoFmt });
  // Formuły czytające tę komórkę — przeliczenie raz na koniec bieżącego zadania.
  if (prev && prev.f && typeof swbFormulaCellChanged === "function") swbFormulaCellChanged(currentSheetName, cellRef);
  if (typeof swbFormulaTouched === "function") swbFormulaTouched(currentSheetName, rowIndex0, absoluteCol);
}

// Zapis formuły wpisanej przez użytkownika (formula-calc.js: swbFormulaCommit).
// fileF = formuła w zapisie pliku (angielskie nazwy, „,”), value = wynik policzony przez
// silnik (null = funkcja, której silnik nie zna — wynik policzy Excel przy otwarciu).
function updateSheetFormulaCell(rowIndex0, colIndex0, fileF, value) {
  if (!workbook || !currentSheetName) return;
  const sheet = workbook.Sheets[currentSheetName];
  if (!sheet) return;
  if (typeof cfmtIsRealCol === "function" && !cfmtIsRealCol(colIndex0)) return;
  const absoluteCol = currentStartCol + colIndex0;
  const cellRef = XLSX.utils.encode_cell({ r: rowIndex0, c: absoluteCol });
  if (typeof undoCaptureCell === "function") undoCaptureCell(rowIndex0, colIndex0);
  const prev = sheet[cellRef];
  const cell = { f: fileF };
  if (prev && prev.s) cell.s = prev.s;
  if (prev && prev.z) cell.z = prev.z;
  if (prev && prev.c) cell.c = prev.c;
  let payload;
  if (value === null || value === undefined) {
    cell.t = "s";
    cell.v = "";
    cell.w = typeof t === "function" ? t("formulaExcelWillCalc") : "";
    payload = { f: fileF, t: "n", v: null };
  } else {
    fcStore(cell, value);
    if (cell.t === "e") payload = { f: fileF, t: "e", err: cell.v };
    else payload = { f: fileF, t: cell.t === "d" ? "n" : cell.t, v: cell.v instanceof Date ? FormulaEngine.jsDateToSerial(cell.v) : cell.v };
  }
  sheet[cellRef] = cell;
  swbEnsureSheetRef(sheet, rowIndex0, absoluteCol);
  bumpSheetDataStamp();
  recordPendingEdit(currentSheetName, cellRef, payload);
  if (typeof swbFormulaCellChanged === "function") swbFormulaCellChanged(currentSheetName, cellRef);
  if (typeof swbFormulaTouched === "function") swbFormulaTouched(currentSheetName, rowIndex0, absoluteCol);
}

// Komórka poza zakresem arkusza (np. suma pod ostatnim wierszem) — poszerz !ref,
// inaczej przebudowa tabeli by jej nie zobaczyła.
function swbEnsureSheetRef(sheet, r, c) {
  const ref = sheet["!ref"];
  if (!ref) { sheet["!ref"] = XLSX.utils.encode_range({ s: { r, c }, e: { r, c } }); return; }
  const rng = XLSX.utils.decode_range(ref);
  if (r <= rng.e.r && c <= rng.e.c && r >= rng.s.r && c >= rng.s.c) return;
  rng.s.r = Math.min(rng.s.r, r); rng.s.c = Math.min(rng.s.c, c);
  rng.e.r = Math.max(rng.e.r, r); rng.e.c = Math.max(rng.e.c, c);
  sheet["!ref"] = XLSX.utils.encode_range(rng);
}

// Zapamiętuje pojedynczą edycję komórki w pendingEdits[sheetName][cellRef].
// Format nadany w tej sesji (pasek „Narzędzia główne”, cell-format.js) i notatka
// (cell-notes.js) przeżywają zmianę wartości: nowa wartość dostaje je z powrotem,
// a wyczyszczenie zostawia sam format i notatkę.
function recordPendingEdit(sheetName, cellRef, payload) {
  if (!sheetName) return;
  if (!pendingEdits[sheetName]) pendingEdits[sheetName] = {};
  const prev = pendingEdits[sheetName][cellRef];
  if (prev && (prev.fmt || "note" in prev)) {
    const keep = {};
    if (prev.fmt && !(payload && payload.fmt)) keep.fmt = prev.fmt;
    if ("note" in prev && !(payload && "note" in payload)) keep.note = prev.note;
    if (Object.keys(keep).length) payload = payload === null ? { clear: true, ...keep } : { ...payload, ...keep };
  }
  pendingEdits[sheetName][cellRef] = payload;
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

// Cache jednego skompilowanego wzorca regex — w jednym przebiegu filtra/podświetlania
// ten sam wzorzec trafia tu setki/tysiące razy (raz na komórkę), więc kompilacja
// za każdym razem byłaby marnotrawstwem. Zawsze case-insensitive ("i"), spójnie
// z resztą wyszukiwania w apce, które nie rozróżnia wielkości liter.
let _searchRegexSrc = null;
let _searchRegexObj = null;
function compileSearchRegex(pattern) {
  if (_searchRegexSrc === pattern) return _searchRegexObj;
  _searchRegexSrc = pattern;
  try {
    _searchRegexObj = new RegExp(pattern, "i");
  } catch (e) {
    _searchRegexObj = null; // niepoprawny wzorzec → nic nie pasuje, zamiast wywalać filtr
  }
  return _searchRegexObj;
}

// Czy pojedyncza komórka (row, kolumna i) dopasowuje zapytanie q w danym trybie.
// Wyłuskane z rowMatchesSingleTerm, by współdzielić tę samą logikę kandydatów
// (m.in. warianty dat dd-mm-yy / dd-mm-yyyy) z podświetlaniem pasujących komórek.
function cellMatchesTerm(row, i, q, mode) {
  const values = row.values;
  if (i >= values.length) return false;
  if (mode === "regex") {
    const re = compileSearchRegex(q);
    return re ? re.test(getDisplayValue(row, i)) : false;
  }
  const display = getDisplayValue(row, i);
  const text = display.toLowerCase();
  const candidates = [text];
  // Kandydaci-daty (dd-mm-yy / dd-mm-yyyy) dokładamy TYLKO dla komórek, które faktycznie
  // wyglądają jak data. Goła liczba (np. kolumna „Nr." 4, 8, 190…) NIE jest datą — inaczej
  // parseDateFlexible zrobiłby z niej numer seryjny Excela (1900-…), a kandydat „dd-mm-1900"
  // zawiera podciąg „190"/„00" i w trybie „zawiera" pasowałby do prawie każdego wiersza.
  const raw = values[i];
  const ds = String(display);
  const looksDate = raw instanceof Date
    || /\d[-/.]\d/.test(ds)
    || (/\d/.test(ds) && /[a-ząćęłńóśźż]/i.test(ds));
  const altDate = looksDate ? parseDateFlexible(raw) : null;
  if (altDate instanceof Date) {
    const dd = String(altDate.getDate()).padStart(2, "0");
    const mm = String(altDate.getMonth() + 1).padStart(2, "0");
    const yyyy = String(altDate.getFullYear());
    const yy = yyyy.slice(-2);
    candidates.push(`${dd}-${mm}-${yy}`);
    candidates.push(`${dd}-${mm}-${yyyy}`);
  }
  if (mode === "equals") return candidates.some((c) => c === q);
  if (mode === "starts_with") return candidates.some((c) => c.startsWith(q));
  if (mode === "contains") return candidates.some((c) => c.includes(q));
  return false;
}

// ── Operatory porównań (liczby/daty) — TYLKO podwójne, tylko przy włączonych
// „Operatorach wyszukiwania". `>>` większe, `<<` mniejsze, `>>=`/`=>>` ≥, `<<=`/`=<<` ≤.
// Term staje się porównaniem dopiero gdy operand parsuje się jako liczba lub data —
// inaczej zostaje zwykłym tekstem (fallback). Komponuje się z &&/||/!/{} za darmo.
function comparisonDayValue(d) {
  return new Date(d.getFullYear(), d.getMonth(), d.getDate()).getTime();
}
function operandLooksLikeDate(s) {
  const t = String(s).trim();
  if (/[a-ząćęłńóśźż]/i.test(t)) return true;             // słowo-miesiąc (np. „1 stycznia 2026")
  if (/\d{4}/.test(t) && /[-/.]/.test(t)) return true;    // rok 4-cyfrowy + separator
  if ((t.match(/[-/.]/g) || []).length >= 2) return true; // dwa separatory → dd-mm-yyyy itp.
  return false;                                            // np. „5.5" → liczba, nie data
}
// Zwraca { op:">"|">="|"<"|"<=", kind:"number"|"date"|"length", value:number } albo null.
// Prefiks „dl"/„len" przed operatorem (np. „dl>>10") zamienia porównanie liczby/daty
// w porównanie DŁUGOŚCI tekstu komórki — bez nowego trybu w UI, tylko rozszerzenie
// już istniejącej składni operatorów (>> << i warianty z „=").
function parseComparisonTerm(q) {
  const m = String(q).match(/^(dl|len)?(>>=|=>>|<<=|=<<|>>|<<)\s*(.+)$/i);
  if (!m) return null;
  const lengthMode = !!m[1];
  const raw = m[2];
  const operand = m[3].trim();
  const op = (raw === ">>=" || raw === "=>>") ? ">="
           : (raw === "<<=" || raw === "=<<") ? "<="
           : (raw === ">>") ? ">" : "<";
  if (lengthMode) {
    const n = parseFloat(operand.replace(",", "."));
    return Number.isFinite(n) ? { op, kind: "length", value: n } : null;
  }
  if (operandLooksLikeDate(operand)) {
    const d = parseDateFlexible(operand);
    if (d instanceof Date && !Number.isNaN(d.getTime())) return { op, kind: "date", value: comparisonDayValue(d) };
  }
  const n = parseCellNumber(operand, operand);
  if (n != null && Number.isFinite(n)) return { op, kind: "number", value: n };
  return null; // operand nie jest liczbą ani datą → nie porównanie
}
function compareWithOp(op, a, b) {
  return op === ">" ? a > b : op === ">=" ? a >= b : op === "<" ? a < b : a <= b;
}
function cellSatisfiesComparison(row, i, cmp) {
  if (i >= row.values.length) return false;
  const raw = row.values[i];
  const display = getDisplayValue(row, i);
  if (cmp.kind === "length") return compareWithOp(cmp.op, String(display).trim().length, cmp.value);
  if (cmp.kind === "date") {
    // tylko komórki będące faktycznie datą — nie parsuj liczb jako numerów seryjnych.
    // Oprócz formatu z separatorem (23-05-26) dopuszczamy też daty słowne („01 sty 26",
    // „1 stycznia 2026"): kandydatem jest tekst z cyfrą i literą, a parseDateFlexible
    // i tak waliduje (formaty słowne przyjmuje tylko gdy wszystkie słowa to miesiące).
    const ds = String(display);
    const looksDate = raw instanceof Date
      || /\d[-/.]\d/.test(ds)
      || (/\d/.test(ds) && /[a-ząćęłńóśźż]/i.test(ds));
    if (!looksDate) return false;
    const d = parseDateFlexible(raw != null ? raw : display);
    if (!(d instanceof Date) || Number.isNaN(d.getTime())) return false;
    return compareWithOp(cmp.op, comparisonDayValue(d), cmp.value);
  }
  if (raw instanceof Date) return false; // porównanie liczbowe, a komórka to data
  const n = parseCellNumber(raw, display);
  if (n == null || !Number.isFinite(n)) return false;
  return compareWithOp(cmp.op, n, cmp.value);
}

// ── Zakres kolumny w zapytaniu: „Kolumna:wartość" (tylko przy operatorach) ──
// Prefiks przed dwukropkiem jest kolumną WYŁĄCZNIE, gdy dokładnie (bez wielkości liter,
// ze zwiniętymi spacjami) równa się nazwie nagłówka — „10:30" czy „Uwaga: brak" przy
// braku takiej kolumny zostają zwykłym tekstem. Powtarzające się nagłówki (bloki
// w szerokiej tabeli) → wszystkie ich kolumny. Najdłuższy pasujący prefiks wygrywa
// (nagłówek sam może mieć dwukropek). Wartość może mieć własne „!" (Status:!Anulowana)
// i cudzysłów (Klient:"Jan Kowalski"). Cały term w cudzysłowie = dosłowny tekst.
function normalizeHeaderKey(s) {
  return String(s == null ? "" : s).trim().replace(/\s+/g, " ").toLowerCase();
}
function criterionColumnMap(criterion) {
  if (criterion._colMap !== undefined) return criterion._colMap;
  let map = null;
  if (Array.isArray(criterion.headers) && criterion.headers.length) {
    map = new Map();
    criterion.headers.forEach((h, i) => {
      const key = normalizeHeaderKey(h);
      if (!key) return;
      if (!map.has(key)) map.set(key, []);
      map.get(key).push(i);
    });
  }
  criterion._colMap = map;
  return map;
}
function unquoteTerm(s) {
  const t = String(s || "").trim();
  return t.length >= 2 && t.startsWith('"') && t.endsWith('"') ? t.slice(1, -1) : t;
}
// → { term, indexes, negated, column } — term gotowy do dopasowania komórek.
function resolveScopedTerm(q, criterion) {
  const out = { term: q, indexes: criterion.indexes, negated: false, exact: false, column: null };
  if (!criterion.operatorsEnabled) return out;
  if (q.length >= 2 && q.startsWith('"') && q.endsWith('"')) {
    out.term = q.slice(1, -1);
    return out;
  }
  const map = q.includes(":") ? criterionColumnMap(criterion) : null;
  if (!map) return out;
  let best = -1;
  for (let k = q.indexOf(":"); k > 0; k = q.indexOf(":", k + 1)) {
    if (map.has(normalizeHeaderKey(q.slice(0, k)))) best = k;
  }
  if (best < 0) return out;
  out.column = normalizeHeaderKey(q.slice(0, best));
  out.indexes = map.get(out.column);
  // Tryb auto (rekordy-cykle): „Data:…" przy powtarzających się blokach ma patrzeć
  // tylko w kolumnę BIEŻĄCEGO cyklu — inaczej rozjechałaby się korelacja tekst+stan.
  if (criterion.blockScope) {
    const inBlock = out.indexes.filter((i) => criterion.blockScope.has(i));
    if (inBlock.length) out.indexes = inBlock;
  }
  let value = q.slice(best + 1).trim();
  if (value.startsWith("!")) {
    out.negated = true;
    value = value.slice(1).trim();
  }
  // „Kolumna:=wartość" = CAŁA komórka równa (niezależnie od trybu Filtra 1);
  // „Kolumna:=""" = pusta, „Kolumna:!=""" = niepusta. „=<<50" / „=>>5" zostają porównaniem.
  if (value.startsWith("=") && !/^=(<<|>>)/.test(value)) {
    out.exact = true;
    value = value.slice(1).trim();
  }
  out.term = unquoteTerm(value);
  return out;
}

// Regex zachowuje wielkość liter WZORCA (flaga "i" i tak ignoruje ją w danych) —
// lowercase zamieniałby \D w \d, \S w \s, \W w \w.
function normalizeTermForMode(term, mode) {
  const t = String(term || "").trim();
  return mode === "regex" ? t : t.toLowerCase();
}

// Wspólny rdzeń dopasowania jednego termu: które komórki wiersza trafiają.
// Zwraca null, gdy term jest pusty (np. samo „Status:" w trakcie pisania).
function scopedTermMatcher(term, criterion) {
  const scoped = resolveScopedTerm(normalizeTermForMode(term, criterion.mode), criterion);
  if (scoped.exact) {
    const want = scoped.term.trim().toLowerCase();
    scoped.cellHit = want
      ? (row, i) => i < row.values.length
        && (String(getDisplayValue(row, i)).trim().toLowerCase() === want || cellMatchesTerm(row, i, want, "equals"))
      : (row, i) => i >= row.values.length || String(getDisplayValue(row, i)).trim() === "";
    return scoped;
  }
  if (!scoped.term) return null;
  const cmp = (criterion.mode !== "regex" && criterion.operatorsEnabled) ? parseComparisonTerm(scoped.term) : null;
  const q = scoped.term;
  const mode = criterion.mode;
  scoped.cellHit = cmp
    ? (row, i) => cellSatisfiesComparison(row, i, cmp)
    : (row, i) => cellMatchesTerm(row, i, q, mode);
  // „Szukaj też w treści notatek” — notatka liczy się jak druga zawartość swojej komórki
  if (!cmp && criterion.searchNotes && typeof swbNoteMatchesTerm === "function") {
    const base = scoped.cellHit;
    scoped.cellHit = (row, i) => base(row, i) || swbNoteMatchesTerm(row, i, q, mode);
  }
  return scoped;
}

function rowMatchesSingleTerm(row, term, criterion) {
  const m = scopedTermMatcher(term, criterion);
  if (!m) return true;
  let hit = false;
  for (const i of m.indexes) {
    if (m.cellHit(row, i)) { hit = true; break; }
  }
  return m.negated ? !hit : hit;
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

// „!" wyklucza jedno słowo — albo całą frazę w cudzysłowie: !"Jan Kowalski",
// także z kolumną: !Klient:"Jan Kowalski". Bez cudzysłowu dalej tylko do spacji.
const NOT_TERM_RE_SRC = '(^|\\s)!(\\S*?"[^"]*"\\S*|\\S+)';

// Tokenize a flat string (no brackets) into NOT-aware terms split by && and ||
// Returns an AST node: { type: "or", children: [ { type: "and", terms: [...] } ] }
function parseFlatExpr(text) {
  function parseAndPart(part) {
    const terms = [];
    const s = String(part || "");
    const notPattern = new RegExp(NOT_TERM_RE_SRC, "g");
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
          const notPattern = new RegExp(NOT_TERM_RE_SRC, "g");
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

// Kolejność sortowania jak w Excelu (sprawdzone na polskim Excelu przez AppleScript,
// 2026-10-17): liczby i daty < tekst < FAŁSZ < PRAWDA < błędy, a PUSTE ZAWSZE NA KOŃCU —
// także przy sortowaniu malejącym (wtedy: błędy, PRAWDA, FAŁSZ, tekst Z→A, liczby 9→0).
// Tekst bez rozróżniania wielkości liter, polskie litery osobno (a < ą < b, l < ł),
// remisy zostają w kolejności z pliku — Intl.Collator("pl", accent) daje dokładnie to.
// Wcześniej: 0 i FAŁSZ lądowały wśród pustych („0 || ''”), puste szły NA POCZĄTEK
// rosnąco, błędy (#DZIEL/0! ma w SheetJS v = kod liczbowy) mieszały się z liczbami,
// a tekst „10” był traktowany jak data seryjna i wchodził między liczby.
// Świadome odstępstwo: tekst, który JEST datą („12.03.2025”, „12 mar 2025”), stoi
// razem z prawdziwymi datami po czasie — Excel zostawiłby go wśród tekstów.
const SWB_SORT_COLLATOR = new Intl.Collator("pl", { sensitivity: "accent" });
const SWB_ERROR_TEXT_RE = /^#(?:NULL!|DIV\/0!|VALUE!|REF!|NAME\?|NUM!|N\/A|GETTING_DATA|SPILL!|CALC!|FIELD!|BLOCKED!|CONNECT!|BUSY!|UNKNOWN!|ZERO!|DZIEL\/0!|ARG!|ADR!|NAZWA\?|LICZBA!|N\/D!?|ROZLANIE!|OBL!)$/;
const SWB_SORT_BLANK = 4;

function swbSerialFromDate(d) {
  return (Date.UTC(d.getFullYear(), d.getMonth(), d.getDate(), d.getHours(), d.getMinutes(), d.getSeconds(), d.getMilliseconds())
    - Date.UTC(1899, 11, 30)) / 86400000;
}

// Klucz sortowania komórki: [ranga, wartość]. Ranga: 0 liczba/data, 1 tekst, 2 logiczna,
// 3 błąd, 4 pusta. `shown` = tekst w tabeli — tylko po nim poznać błąd z pliku.
function swbSortKey(v, shown) {
  if (v == null || v === "") return [SWB_SORT_BLANK, 0];
  if (typeof v === "boolean") return [2, v ? 1 : 0];
  if (v instanceof Date) return Number.isNaN(v.getTime()) ? [1, ""] : [0, swbSerialFromDate(v)];
  if (typeof v === "number") {
    if (typeof shown === "string" && shown[0] === "#" && SWB_ERROR_TEXT_RE.test(shown)) return [3, 0];
    return Number.isFinite(v) ? [0, v] : [3, 0];
  }
  const s = String(v);
  if (!s.trim()) return [1, s];
  if (s[0] === "#" && SWB_ERROR_TEXT_RE.test(s.trim())) return [3, 0];
  // czysta liczba w tekście („10”, „0012”) zostaje tekstem — parseDateFlexible
  // wziąłby ją za numer seryjny daty
  if (!/^[+-]?[\d\s]*\d(?:[.,]\d+)?$/.test(s)) {
    const d = parseDateFlexible(s);
    if (d) return [0, swbSerialFromDate(d)];
  }
  return [1, s];
}

function swbCompareSortKeys(a, b) {
  if (a[0] !== b[0]) return a[0] - b[0];
  if (a[0] === 1) return SWB_SORT_COLLATOR.compare(a[1], b[1]);
  return a[1] - b[1];
}

// Zgodność wstecz (inne moduły porównują pojedyncze wartości).
function compareSortValues(av, bv) {
  return swbCompareSortKeys(swbSortKey(av), swbSortKey(bv));
}

// Sortowanie wierszy według reguł — klucze liczone RAZ na wiersz (komparator woła się
// n·log n razy, a parseDateFlexible/kolacja na każde porównanie to zbędny koszt).
function swbSortRowsByRules(rows, rules) {
  if (!rules.length || rows.length < 2) return;
  const keyed = rows.map((row, i) => {
    const vals = row.rawValues || row.values || [];
    const keys = rules.map((rule) => swbSortKey(vals[rule.idx], row.display ? row.display[rule.idx] : undefined));
    return { row, keys, i };
  });
  keyed.sort((a, b) => {
    for (let k = 0; k < rules.length; k++) {
      const ka = a.keys[k];
      const kb = b.keys[k];
      // puste zawsze na dole, niezależnie od kierunku (jak Excel)
      const ab = ka[0] === SWB_SORT_BLANK;
      const bb = kb[0] === SWB_SORT_BLANK;
      if (ab || bb) {
        if (ab && bb) continue;
        return ab ? 1 : -1;
      }
      const cmp = swbCompareSortKeys(ka, kb);
      if (cmp !== 0) return rules[k].dir === "desc" ? -cmp : cmp;
    }
    return a.i - b.i;
  });
  for (let i = 0; i < keyed.length; i++) rows[i] = keyed[i].row;
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
  scheduleViewRefresh({ table: true, analyses: true, formula: true });
}

// Zbiera „surowe" pozytywne (nie-zanegowane) szukane frazy z drzewa zapytania.
// Zanegowane terminy i całe zanegowane pod-wyrażenia pomijamy — nie ma jednej
// komórki „dzięki której" wiersz przeszedł negację, więc nie ma czego podświetlać.
function gatherPositiveTermStrings(parsed) {
  const out = [];
  const walkTerm = (term) => {
    if (!term || term.negated) return;
    if (term.subExpr) walkAst(term.subExpr);
    else if (term.term) out.push(term.term.trim());
  };
  const walkAst = (ast) => {
    ast.children.forEach((andGroup) => andGroup.terms.forEach(walkTerm));
  };
  parsed.groups.forEach((group) => group.forEach(walkTerm));
  return out.filter(Boolean);
}

// Dla wiersza, który przeszedł filtr, wyznacza zbiór indeksów kolumn, których
// zawartość pozytywnie pasuje do filtrów tekstowych / zakresu dat. To podstawa
// subtelnego podświetlenia „komórek powodujących trafienie".
function collectMatchingCellsForRow(row, criteria, dateFilter) {
  const cols = new Set();
  for (const criterion of criteria) {
    if (criterion.negated || !criterion.query) continue;
    const parsed = parseQueryTerms(criterion.query, criterion.operatorsEnabled);
    const positiveTerms = gatherPositiveTermStrings(parsed);
    for (const q of positiveTerms) {
      const m = scopedTermMatcher(q, criterion);
      // „Status:!Anulowana" ani „Uwagi:=""" (pusta) nie wskazują komórki-dowodu
      if (!m || m.negated || (m.exact && !m.term)) continue;
      for (const i of m.indexes) {
        if (m.cellHit(row, i)) cols.add(i);
      }
    }
  }
  const range = dateFilter && dateFilter.range;
  if (dateFilter && !dateFilter.negated && range && (range.from || range.to)) {
    for (const idx of dateFilter.indexes) {
      if (idx >= row.values.length) continue;
      const raw = row.rawValues ? row.rawValues[idx] : row.values[idx];
      const d = parseDateFlexible(raw ?? getDisplayValue(row, idx));
      if (!d) continue;
      if (range.from && d < range.from) continue;
      if (range.to && d > range.to) continue;
      cols.add(idx);
    }
  }
  return cols;
}

// Wspólna normalizacja wartości do porównań walidacji: przy „ignoruj wielkość/spacje"
// → trim + lowercase; inaczej dosłownie. Używana i przy budowie zbioru dozwolonych,
// i przy sprawdzaniu wierszy, żeby reguła była identyczna po obu stronach.
function normalizeValidationValue(raw, caseInsensitive) {
  const s = String(raw == null ? "" : raw);
  return caseInsensitive ? s.trim().toLowerCase() : s;
}

// Czy wiersz narusza listę dozwolonych wartości w kolumnie validationState.colIdx.
function rowIsValidationViolation(row) {
  const st = validationState;
  if (st.colIdx < 0 || !st.allowed) return false;
  const raw = getDisplayValue(row, st.colIdx);
  if (String(raw == null ? "" : raw).trim() === "") return !st.ignoreEmpty;
  return !st.allowed.has(normalizeValidationValue(raw, st.caseInsensitive));
}

// Jedna definicja „czy wiersz przechodzi filtr” — dla Entera (applyFilters) i dla podglądu
// szybkiego szukania (renderQsLive). Wcześniej podgląd miał własną, uboższą logikę: nie
// znał tokenów (@notatka, @wtoku, @dni>>30), filtra 2, dat ani „Tylko z notatkami”, więc
// np. „@notatka” przed Enterem pokazywało „brak wyników” (zgłoszenie Mateusza 2026-10-25).
// query1/operators1 — zapytanie z pola szybkiego szukania zamiast pola filtra 1.
// preview — bez skutków ubocznych (nie startuje przebiegu trybów auto, nie zapamiętuje trafień).
function swbBuildRowFilter({ query1 = null, operators1 = null, preview = false } = {}) {
  const criteria = [
    {
      query: normalizeTermForMode(query1 != null ? query1 : searchQueryEl.value, getNormalizedSelectValue(filterModeEl)),
      mode: getNormalizedSelectValue(filterModeEl),
      headers: currentHeaders,
      indexes: resolveIndexes(currentHeaders, columnSelections.filter1),
      emptyMode: getNormalizedSelectValue(filterEmptyModeEl),
      negated: filterNegateEl.checked,
      operatorsEnabled: operators1 != null ? !!operators1 : !!filterOperatorsEl?.checked,
    },
    {
      query: normalizeTermForMode(searchQuery2El.value, getNormalizedSelectValue(filterMode2El)),
      mode: getNormalizedSelectValue(filterMode2El),
      headers: currentHeaders,
      indexes: resolveIndexes(currentHeaders, columnSelections.filter2),
      emptyMode: getNormalizedSelectValue(filterEmptyMode2El),
      negated: filterNegate2El.checked,
      operatorsEnabled: !!filterOperators2El?.checked,
    },
  ];

  // Notatki komórek (cell-notes.js): „Szukaj też w treści notatek” + tokeny @notatka w polach szukania.
  const searchNotes = typeof swbNotesSearchEnabled === "function" && swbNotesSearchEnabled();
  const noteRules = [];
  criteria.forEach((c) => {
    c.searchNotes = searchNotes;
    if (typeof swbExtractNoteTokens !== "function") return;
    const ex = swbExtractNoteTokens(c.query);
    if (ex.rules) { c.query = ex.query; noteRules.push(...ex.rules); }
  });
  const notesOnly = typeof swbNotesOnly !== "undefined" && swbNotesOnly;

  const dateFilter = {
    indexes: resolveIndexes(currentHeaders, columnSelections.date),
    range: getDateRange(),
    emptyMode: getNormalizedSelectValue(dateEmptyModeEl),
    negated: dateNegateEl.checked,
  };
  const onlyNonEmpty = onlyNonEmptyEl.checked;

  // ── Tryby auto ──────────────────────────────────────────────────────────────
  // Stan efektywny = chipy z panelu + tokeny „@wtoku / @ostatni / @dni>>30" wpisane
  // wprost w pole szukania (te ostatnie wycinamy z zapytania, żeby parser termów
  // ich nie zobaczył). Gdy arkusz nie ma powtarzalnych bloków, model jest null
  // i cała ta warstwa jest przezroczysta.
  const smartModel = typeof getSmartModel === "function" ? getSmartModel() : null;
  const smartResolved = smartModel ? resolveEffectiveSmartState(criteria.map((c) => c.query)) : null;
  if (smartResolved) {
    smartResolved.queries.forEach((q, i) => { criteria[i].query = q; });
  }
  // Stan trybów auto podany JAWNIE (nie przez globalny przebieg) — podgląd szukania liczy
  // to samo co Enter, nie ruszając zapamiętanych trafień bieżącego widoku.
  const smartState = smartModel ? (smartResolved ? smartResolved.state : smartFilterState) : null;
  if (!preview && typeof smartBeginFilterPass === "function") smartBeginFilterPass(smartResolved ? smartResolved.state : null);
  const smartActive = !!smartState && (smartNeedsRecordScope(smartState) || smartState.rowFlags.length > 0);
  const smartScoped = smartActive && smartNeedsRecordScope(smartState);
  const scopeCache = preview ? new Map() : null; // podgląd: własna pamięć zakresów rekordów
  const hitRecords = preview ? new Map() : null; // podgląd: trafione rekordy (dowód trafienia)
  const remember = (row, recs) => { if (preview) hitRecords.set(row, recs); else smartRememberMatch(row, recs); };

  const rowPasses = (row) => {
    if (!rowMatchesDateFilter(row, dateFilter)) return false;
    // Walidacja listą: w trybie „pokaż tylko niezgodne" zostają wyłącznie naruszenia.
    if (validationState.showOnly && !rowIsValidationViolation(row)) return false;
    if (notesOnly && !swbRowHasNote(row)) return false;
    if (noteRules.length && !swbRowPassesNoteRules(row, noteRules)) return false;
    if (!smartActive) return rowMatchesTextFilter(row, criteria, onlyNonEmpty);

    const records = buildSmartRecords(row, smartModel);
    if (!smartRowFlagsPass(records, smartState)) return false;
    if (!smartScoped) {
      if (!rowMatchesTextFilter(row, criteria, onlyNonEmpty)) return false;
      remember(row, records.filter((rec) => rec.filled));
      return true;
    }
    // Korelacja: tekst filtra sprawdzamy w obrębie KAŻDEGO pasującego rekordu osobno,
    // więc „Kowalski + w toku" trafia tylko tam, gdzie Kowalski faktycznie nie skończył.
    const candidates = records.filter((rec) => smartRecordPasses(row, rec, smartModel, smartState));
    if (!candidates.length) return false;
    const hits = candidates.filter((rec) => rowMatchesTextFilter(row, smartScopeCriteria(criteria, rec, smartModel, scopeCache), onlyNonEmpty));
    if (!hits.length) return false;
    remember(row, hits);
    return true;
  };

  return { rowPasses, criteria, dateFilter, noteRules, notesOnly, smartModel, smartScoped, hitRecords };
}

function applyFilters() {
  // Filtrowanie i reset sortowania (przez applyCurrentSort) przechodzą tędy —
  // następny render ma animować odsłonięcie/przestawienie wierszy (FLIP).
  flipNextRender = true;
  const { rowPasses, criteria, dateFilter, notesOnly, smartScoped } = swbBuildRowFilter();
  // Podświetlanie komórek: checkbox z filtra tekstowego LUB dat (zsynchronizowane),
  // ALBO tryby szybkiego szukania „Podświetl pasujące" (cells) / „Filtruj + podświetl" (filter-cells).
  // Tryb auto sam z siebie włącza podświetlanie: inaczej w szerokim arkuszu nie widać,
  // KTÓRY cykl sprawił, że wiersz przeszedł filtr.
  highlightMatchedCells = !!((highlightMatchCellsEl && highlightMatchCellsEl.checked) || (highlightMatchCellsDateEl && highlightMatchCellsDateEl.checked) || quickSearchCellsMode || quickSearchFilterCellsMode || smartScoped);
  matchedCellsByRow = new Map();
  const collectCells = (row) => {
    if (!highlightMatchedCells) return;
    let cols = collectMatchingCellsForRow(row, criteria, dateFilter);
    if (smartScoped) cols = smartRestrictMatchedCols(row, cols);
    if (cols.size) matchedCellsByRow.set(row.rowIndex0, cols);
  };

  const shouldFilterRows = (filtersCommitted || validationState.showOnly || notesOnly) && !quickSearchHighlightMode && !quickSearchCellsMode;
  if (shouldFilterRows) {
    // Tryb „filtruj": tylko pasujące wiersze (po kliknięciu Filtruj lub równoważnej akcji)
    matchedRowIndexes = new Set();
    viewRows = baseRows.filter(rowPasses);
    viewRows.forEach(collectCells);
  } else {
    // Podgląd / zaznacz: wszystkie wiersze widoczne; opcjonalnie podświetl trafienia
    viewRows = baseRows.slice();
    matchedRowIndexes = new Set();
    baseRows.forEach((row) => {
      if (!rowPasses(row)) return;
      if (quickSearchHighlightMode) matchedRowIndexes.add(row.rowIndex0);
      collectCells(row);
    });
  }
  // Migawka dla paska „co teraz filtruje" (filter-bar.js) — liczby z TEGO przebiegu.
  lastAppliedFilters = {
    filtering: shouldFilterRows,
    marking: !shouldFilterRows && (quickSearchHighlightMode || quickSearchCellsMode),
    matched: shouldFilterRows ? viewRows.length
      : (quickSearchHighlightMode ? matchedRowIndexes.size : matchedCellsByRow.size),
    total: baseRows.length,
  };
  if (typeof renderActiveFilters === "function") renderActiveFilters();
  if (typeof swbNotesSyncFilterUi === "function") swbNotesSyncFilterUi();
}

function sortRows() {
  normalizeSortState();
  if (!multiSortState.length) return;
  flipNextRender = true; // następny render tabeli animuje przestawienie wierszy (FLIP)
  swbSortRowsByRules(viewRows, resolveSortRules(currentHeaders));
}

function sortRowsForHeaders(rows, headers) {
  normalizeSortState();
  if (!multiSortState.length || !Array.isArray(rows) || !Array.isArray(headers)) return;
  swbSortRowsByRules(rows, resolveSortRules(headers));
}

// `Array#sort` wywołuje komparator wielokrotnie, więc wyszukiwanie nazwy kolumny
// w nagłówkach w jego środku robiło tę samą pracę O(porównań × kolumn). Reguły są
// niezmienne przez jedno synchroniczne sortowanie — indeksy można bezpiecznie
// wyliczyć tylko raz, zachowując identyczny porządek i obsługę nieistniejących kolumn.
function resolveSortRules(headers) {
  return multiSortState
    .map((rule) => ({ idx: headers.indexOf(rule.col), dir: rule.dir }))
    .filter((rule) => rule.idx >= 0);
}

function getActiveRepeatingGroup() {
  return Array.isArray(currentRepeatingBlocks) && currentRepeatingBlocks.length ? currentRepeatingBlocks[0] : null;
}
