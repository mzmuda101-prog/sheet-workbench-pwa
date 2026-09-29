// report.js — Raport do druku / PDF (krok 1: wersja KRÓTKA, jedna strona A4).
//
// Po co: eksport „Drukuj / PDF" drukuje samą tabelę. Raport odpowiada na pytanie
// „co jest w tych danych?" — kafelki z liczbami, kilka wniosków zdaniami i jeden
// wykres — na jednej kartce, do przeczytania albo wpięcia do segregatora.
//
// Zasady:
//  • Źródło = BIEŻĄCY widok (viewRows: filtry, szukanie), jak w eksporcie i spisywaniu.
//    Pierwsza rzecz na kartce to „jaki to wycinek” — wydruk bez informacji o filtrze
//    łatwo wziąć za pełne dane.
//  • Wnioski liczone lokalnie z prostych reguł, bez AI — plik nie opuszcza urządzenia.
//    Reguła, która nie ma podstaw w danych (za mało wierszy, brak kolumny), milczy.
//    Lepiej 2 prawdziwe zdania niż 6 naciąganych.
//  • Kolumny rozpoznajemy po WARTOŚCIACH (liczby / daty / kategorie); nazwa nagłówka
//    tylko podbija wybór („Status”, „Kwota”), nigdy go nie wymusza.
//  • Podgląd na ekranie = to, co wyjdzie z drukarki: ta sama kartka 210 mm, tylko
//    przeskalowana. PDF robi systemowy dialog druku („Zapisz jako PDF”).
//  • Wygląd (styl, kolor, wielkość tekstu) zapamiętany dla wszystkich plików.

const RP_PREFS_KEY = "swb-report-prefs";
const RP_STYLES = ["modern", "classic", "ink"];
const RP_ACCENTS = {
  green: "#2f6f5c",
  blue: "#2b5d9b",
  violet: "#6a4c93",
  orange: "#b5602a",
  graphite: "#3d4145",
};
const RP_SIZES = ["normal", "large"];
const RP_MAX_FINDINGS = 6;        // w wersji krótkiej — ma się zmieścić na jednej kartce
const RP_MAX_NUM_COLS = 8;
const RP_MAX_OTHER_CATS = 3;
const RP_MAX_DATA_COLS = 8;       // więcej kolumn na pionowym A4 = nieczytelne; resztę daje Eksport
const RP_MAX_DATA_ROWS = 500;
const RP_MAX_MONTHS = 18;

// Sekcje raportu w stałej kolejności. „Zakres” (jaki to wycinek danych) NIE jest na
// liście — jest zawsze, bo wydruk bez niego łatwo wziąć za pełne dane.
// `has` = czy sekcja ma w tym arkuszu z czego powstać; bez tego pusty rozdział.
const RP_SECTIONS = [
  { id: "tiles", label: "rpSecTiles" },
  { id: "findings", label: "rpSecFindings" },
  { id: "chart", label: "rpSecChart", has: (d) => !!d.category },
  { id: "months", label: "rpSecMonths", has: (d) => !!(d.date && d.date.months.size >= 2) },
  { id: "numbers", label: "rpSecNumbers", has: (d) => d.numericCols.length > 0 },
  { id: "categories", label: "rpSecCategories", has: (d) => d.otherCategories.length > 0 },
  { id: "columns", label: "rpSecColumns" },
  { id: "data", label: "rpSecData" },
];
const RP_PRESETS = {
  short: ["tiles", "findings", "chart"],
  normal: ["tiles", "findings", "chart", "months", "numbers", "categories"],
  detailed: RP_SECTIONS.map((sec) => sec.id),
};
const RP_MAX_BARS = 6;
const RP_PAGE_MM = { w: 210, h: 297 };
const RP_PRINT_MARGIN_MM = 12;  // = @page rpA4 margin (góra/dół) w app.css
const RP_SCREEN_PAD_MM = 16;    // = padding-top kartki na ekranie
const RP_PRINT_PAD_TOP_MM = 4;  // = padding-top kartki w druku

// Podpowiedzi z nazw nagłówków (PL + EN). Tylko podbijają wynik — kolumna i tak musi
// mieć odpowiednie wartości.
const RP_CAT_RE = /status|stan|etap|faza|typ|rodzaj|kategor|grupa|dział|dzial|priorytet|miasto|region|type|state|stage|category|group|priority|city/i;
const RP_NUM_RE = /kwot|cen|wart|sum|koszt|brutto|netto|ilo|liczb|godz|czas|dni|waga|wynik|amount|price|value|cost|total|qty|quantity|hours|days|score/i;
const RP_ID_RE = /^\s*(nr|lp|l\.\s*p|id|numer|no|#)(\b|\.|$)/i;
const RP_DATE_TEXT_RE = /^\d{1,4}[-./]\d{1,2}[-./]\d{1,4}/;
const RP_NUMBER_TEXT_RE = /^[-+]?[\d\s .,]+%?$/;

const reportBtn = document.getElementById("reportBtn");
const rpOverlayEl = document.getElementById("reportOverlay");
const rpStageEl = document.getElementById("rpStage");
const rpWrapEl = document.getElementById("rpSheetWrap");
const rpPageEl = document.getElementById("rpPage");
const rpStyleEl = document.getElementById("rpStyle");
const rpAccentsEl = document.getElementById("rpAccents");
const rpSizeBtn = document.getElementById("rpSizeBtn");
const rpPrintBtn = document.getElementById("rpPrintBtn");
const rpCloseBtn = document.getElementById("rpCloseBtn");
const rpFitNoteEl = document.getElementById("rpFitNote");
const rpContentBtn = document.getElementById("rpContentBtn");
const rpContentPanelEl = document.getElementById("rpContentPanel");
const rpPresetsEl = document.getElementById("rpPresets");
const rpSectionListEl = document.getElementById("rpSectionList");

let rpIsOpen = false;
let rpReturnFocusEl = null;
let rpData = null;
let rpTitle = "";
let rpPrefs = { style: "modern", accent: "green", size: "normal", preset: "short", sections: RP_PRESETS.short.slice() };
let rpDocTitleBefore = "";

// ── Ustawienia wyglądu ──────────────────────────────────────────────────────

function rpLoadPrefs() {
  try {
    const p = JSON.parse(localStorage.getItem(RP_PREFS_KEY) || "{}");
    rpPrefs = {
      style: RP_STYLES.includes(p.style) ? p.style : "modern",
      accent: RP_ACCENTS[p.accent] ? p.accent : "green",
      size: RP_SIZES.includes(p.size) ? p.size : "normal",
      preset: RP_PRESETS[p.preset] || p.preset === "custom" ? p.preset : "short",
      sections: Array.isArray(p.sections)
        ? p.sections.filter((id) => RP_SECTIONS.some((sec) => sec.id === id))
        : RP_PRESETS.short.slice(),
    };
  } catch { /* prywatne okno — zostają domyślne */ }
}

function rpSavePrefs() {
  try { localStorage.setItem(RP_PREFS_KEY, JSON.stringify(rpPrefs)); } catch { /* bez pamięci */ }
}

// ── Formatowanie ────────────────────────────────────────────────────────────

function rpLocale() {
  return (typeof I18N !== "undefined" && I18N[currentLang] && I18N[currentLang].locale) || "pl-PL";
}

function rpNum(n, digits = 2) {
  if (!Number.isFinite(n)) return "—";
  return n.toLocaleString(rpLocale(), { maximumFractionDigits: digits });
}

function rpPct(part, all) {
  return all ? Math.round((part / all) * 100) : 0;
}

function rpDate(d, withYear = true) {
  if (!(d instanceof Date) || Number.isNaN(d.getTime())) return "";
  return d.toLocaleDateString(rpLocale(), withYear
    ? { day: "numeric", month: "short", year: "numeric" }
    : { day: "numeric", month: "short" });
}

function rpMonthLabel(key) {
  const [y, m] = key.split("-").map(Number);
  return new Date(y, m - 1, 1).toLocaleDateString(rpLocale(), { month: "long", year: "numeric" });
}

function rpColName(c) {
  return typeof exportColLabel === "function" ? exportColLabel(currentHeaders[c], c) : String(currentHeaders[c] ?? c + 1);
}

// ── Rozpoznanie kolumn po wartościach ───────────────────────────────────────
// Jeden przebieg po wierszach. Liczby i daty rozróżniamy po tym, co WIDZI użytkownik:
// komórka z liczbą, która wyświetla się jako „05.01.2026”, jest datą; parseDateFlexible
// sam z siebie zrobiłby datę z każdej liczby (numer seryjny Excela).

function rpClassifyCell(value, shown) {
  if (value instanceof Date) return "date";
  if (typeof value === "number") {
    if (RP_DATE_TEXT_RE.test(shown) || /[a-ząćęłńóśźż]/i.test(shown)) {
      return parseDateFlexible(value) instanceof Date ? "date" : "text";
    }
    return "num";
  }
  if (typeof value === "string") {
    const v = value.trim();
    if (RP_NUMBER_TEXT_RE.test(v)) return "text"; // „0012”, „1 200” jako tekst — nie zgadujemy
    if (RP_DATE_TEXT_RE.test(v) && parseDateFlexible(v) instanceof Date) return "date";
  }
  return "text";
}

function rpProfileColumns(rows) {
  const cols = currentHeaders.map((_, c) => ({
    c,
    name: rpColName(c),
    nonEmpty: 0,
    nums: [],
    dates: [],
    counts: new Map(),
  }));
  rows.forEach((row) => {
    cols.forEach((col) => {
      const shown = String(getDisplayValue(row, col.c) ?? "").trim();
      if (!shown) return;
      col.nonEmpty += 1;
      col.counts.set(shown, (col.counts.get(shown) || 0) + 1);
      const value = row.values ? row.values[col.c] : shown;
      const kind = rpClassifyCell(value, shown);
      if (kind === "num") col.nums.push(value);
      else if (kind === "date") {
        const d = parseDateFlexible(value);
        if (d instanceof Date) col.dates.push(d);
      }
    });
  });
  cols.forEach((col) => {
    col.unique = col.counts.size;
    col.fill = rows.length ? col.nonEmpty / rows.length : 0;
    col.kind = !col.nonEmpty ? "empty"
      : col.nums.length / col.nonEmpty >= 0.8 ? "num"
        : col.dates.length / col.nonEmpty >= 0.8 ? "date"
          : "text";
  });
  return cols;
}

// Numer porządkowy (1, 2, 3…) to też liczby, ale ich suma nic nie mówi.
function rpLooksLikeCounter(col) {
  if (RP_ID_RE.test(col.name)) return true;
  if (col.unique !== col.nonEmpty || col.nums.length < 3) return false;
  if (!col.nums.every((n) => Number.isInteger(n))) return false;
  const min = Math.min(...col.nums);
  const max = Math.max(...col.nums);
  return max - min + 1 === col.nums.length;
}

// Kolumny-kategorie (mało różnych wartości, zwykle wypełnione), najlepsza pierwsza.
function rpCategories(cols, rowCount) {
  if (rowCount < 3) return [];
  return cols
    .filter((col) => col.kind === "text" && col.fill >= 0.5
      && col.unique >= 2 && col.unique <= 12 && col.unique <= col.nonEmpty * 0.6)
    .map((col) => ({
      col,
      score: (RP_CAT_RE.test(col.name) ? 10 : 0) + col.fill * 3 - col.unique * 0.1,
      entries: Array.from(col.counts.entries()).sort((a, b) => b[1] - a[1]),
      total: col.nonEmpty,
    }))
    .sort((a, b) => b.score - a.score);
}

function rpMedian(sorted) {
  const n = sorted.length;
  if (!n) return NaN;
  return n % 2 ? sorted[(n - 1) / 2] : (sorted[n / 2 - 1] + sorted[n / 2]) / 2;
}

function rpNumStats(col) {
  const nums = col.nums.slice().sort((a, b) => a - b);
  const sum = nums.reduce((s, n) => s + n, 0);
  return {
    col,
    n: nums.length,
    sum,
    avg: sum / nums.length,
    min: nums[0],
    max: nums[nums.length - 1],
    median: rpMedian(nums),
  };
}

// Kolumny liczbowe warte statystyk (bez numerów porządkowych), najlepsza pierwsza.
function rpNumericCols(cols) {
  return cols
    .filter((col) => col.kind === "num" && !rpLooksLikeCounter(col))
    .map((col) => ({ col, score: (RP_NUM_RE.test(col.name) ? 10 : 0) + col.fill * 3 }))
    .sort((a, b) => b.score - a.score)
    .map(({ col }) => rpNumStats(col));
}

function rpPickDate(cols) {
  let best = null;
  cols.forEach((col) => {
    if (col.kind !== "date" || col.dates.length < 2) return;
    if (!best || col.dates.length > best.dates.length) best = col;
  });
  if (!best) return null;
  let min = best.dates[0];
  let max = best.dates[0];
  const months = new Map();
  best.dates.forEach((d) => {
    if (d < min) min = d;
    if (d > max) max = d;
    const key = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`;
    months.set(key, (months.get(key) || 0) + 1);
  });
  // Remis = wymieniamy wszystkie (do 3). Wybranie jednego z równych byłoby zmyśleniem.
  const peak = Math.max(...months.values());
  // Więcej niż 3 ex aequo = nie ma „najwięcej” — zdanie wtedy nie pada (pusta lista).
  const tied = Array.from(months.keys()).filter((k) => months.get(k) === peak).sort();
  const topMonths = tied.length <= 3 ? tied : [];
  return { col: best, min, max, months, topMonths, peak };
}

// Kolumna-identyfikator z powtórkami = możliwe duplikaty. Tylko gdy kolumna naprawdę
// wygląda na identyfikator (nazwa „Nr/ID” albo prawie same unikaty), inaczej każda
// powtarzająca się wartość byłaby „duplikatem”.
function rpPickDuplicates(cols) {
  for (const col of cols) {
    if (col.nonEmpty < 5 || col.unique === col.nonEmpty) continue;
    const idLike = RP_ID_RE.test(col.name) || (col.nonEmpty >= 10 && col.unique / col.nonEmpty >= 0.95);
    if (!idLike) continue;
    const dups = Array.from(col.counts.entries()).filter(([, n]) => n > 1);
    if (dups.length) return { col, dups };
  }
  return null;
}

function rpCollect() {
  const rows = Array.isArray(viewRows) ? viewRows : [];
  const total = Array.isArray(baseRows) ? baseRows.length : rows.length;
  const filtering = !!(typeof lastAppliedFilters !== "undefined" && lastAppliedFilters && lastAppliedFilters.filtering);
  const cols = rpProfileColumns(rows);
  const used = cols.filter((c) => c.nonEmpty > 0);
  const emptyCols = cols.filter((c) => c.nonEmpty === 0);
  const filled = used.reduce((s, c) => s + c.nonEmpty, 0);
  const view = typeof captureViewState === "function" ? captureViewState() : null;
  const cats = rpCategories(cols, rows.length);
  const numericCols = rpNumericCols(cols);
  return {
    rowsList: rows,
    rows: rows.length,
    total,
    filtering,
    viewText: view && typeof describeViewState === "function" ? describeViewState(view) : "",
    cols,
    used,
    emptyCols,
    completeness: used.length && rows.length ? rpPct(filled, used.length * rows.length) : 0,
    category: cats[0] || null,
    otherCategories: cats.slice(1, 1 + RP_MAX_OTHER_CATS),
    numeric: numericCols[0] || null,
    numericCols: numericCols.slice(0, RP_MAX_NUM_COLS),
    date: rpPickDate(cols),
    dups: rpPickDuplicates(cols),
    at: new Date(),
  };
}

// ── Wnioski zdaniami ────────────────────────────────────────────────────────
// Kolejność = przydatność. Każda reguła zwraca zdanie albo nic.

function rpFindings(d) {
  const out = [];
  if (d.category) {
    const { col, entries, total } = d.category;
    const [v1, n1] = entries[0];
    let text = t("rpFindTop", { col: col.name, value: v1, pct: rpPct(n1, total), n: n1, all: total });
    if (entries[1]) text += ` ${t("rpFindNext", { value: entries[1][0], pct: rpPct(entries[1][1], total) })}`;
    out.push({ text });
  }
  if (d.numeric) {
    const s = d.numeric;
    out.push({ text: t("rpFindSum", { col: s.col.name, sum: rpNum(s.sum), avg: rpNum(s.avg), min: rpNum(s.min), max: rpNum(s.max) }) });
  }
  if (d.date) {
    const dt = d.date;
    let text = t("rpFindDates", { col: dt.col.name, from: rpDate(dt.min), to: rpDate(dt.max) });
    if (dt.months.size >= 2 && dt.topMonths.length && dt.topMonths.length < dt.months.size) {
      text += ` ${dt.topMonths.length > 1
        ? t("rpFindTopMonths", { months: dt.topMonths.map(rpMonthLabel).join(", "), n: dt.peak })
        : t("rpFindTopMonth", { month: rpMonthLabel(dt.topMonths[0]), n: dt.peak })}`;
    }
    out.push({ text });
  }
  if (d.numeric && d.numeric.n >= 5 && d.numeric.median > 0 && d.numeric.max >= 3 * d.numeric.median) {
    const s = d.numeric;
    out.push({ text: t("rpFindOutlier", { col: s.col.name, max: rpNum(s.max), x: rpNum(s.max / s.median, 1), median: rpNum(s.median) }), tone: "warn" });
  }
  if (d.dups) {
    const { col, dups } = d.dups;
    const ex = dups.slice(0, 3).map(([v, n]) => `${v} ×${n}`).join(", ");
    out.push({ text: t("rpFindDups", { col: col.name, n: dups.length, ex }), tone: "warn" });
  }
  const gaps = d.used
    .map((c) => ({ c, pct: 100 - Math.round(c.fill * 100) }))
    .filter((g) => g.pct >= 20)
    .sort((a, b) => b.pct - a.pct)
    .slice(0, 3);
  if (gaps.length && d.rows >= 3) {
    out.push({ text: t("rpFindGaps", { list: gaps.map((g) => t("rpFindGapItem", { col: g.c.name, pct: g.pct })).join(", ") }), tone: "warn" });
  }
  if (d.emptyCols.length) out.push({ text: t("rpFindEmptyCols", { n: d.emptyCols.length }) });
  if (!out.length) out.push({ text: t("rpFindNone") });
  return out;
}

// ── Kafelki ─────────────────────────────────────────────────────────────────

function rpTiles(d) {
  const tiles = [];
  tiles.push({
    label: t("rpTileRows"),
    value: rpNum(d.rows, 0),
    sub: d.filtering ? t("rpTileRowsOf", { total: rpNum(d.total, 0) }) : t("rpTileRowsAll"),
  });
  tiles.push({
    label: t("rpTileCols"),
    value: rpNum(d.used.length, 0),
    sub: d.emptyCols.length ? t("rpTileColsEmpty", { n: d.emptyCols.length }) : "",
  });
  tiles.push({ label: t("rpTileComplete"), value: `${d.completeness}%`, sub: t("rpTileCompleteSub") });
  if (d.numeric) {
    tiles.push({ label: t("rpTileSum", { col: d.numeric.col.name }), value: rpNum(d.numeric.sum), sub: t("rpTileAvg", { avg: rpNum(d.numeric.avg) }) });
  }
  if (d.category) {
    const [v, n] = d.category.entries[0];
    tiles.push({ label: t("rpTileTop", { col: d.category.col.name }), value: v, sub: `${rpPct(n, d.category.total)}% (${rpNum(n, 0)})`, text: true });
  }
  if (d.date) {
    const sameYear = d.date.min.getFullYear() === d.date.max.getFullYear();
    tiles.push({
      label: t("rpTileDates", { col: d.date.col.name }),
      value: `${rpDate(d.date.min, !sameYear)} – ${rpDate(d.date.max)}`,
      sub: "",
      text: true,
    });
  }
  return tiles.slice(0, 6);
}

// ── Budowa kartki ───────────────────────────────────────────────────────────

function rpEl(tag, cls, text) {
  const el = document.createElement(tag);
  if (cls) el.className = cls;
  if (text != null) el.textContent = text;
  return el;
}

function rpDefaultTitle() {
  const base = currentFileName ? currentFileName.replace(/\.[^.]+$/, "") : t("rpTitleFallback");
  return t("rpTitleDefault", { name: base });
}

function rpRender() {
  if (!rpPageEl || !rpData) return;
  const d = rpData;
  rpPageEl.dataset.style = rpPrefs.style;
  rpPageEl.dataset.size = rpPrefs.size;
  rpPageEl.style.setProperty("--rp-accent", RP_ACCENTS[rpPrefs.accent] || RP_ACCENTS.green);
  rpPageEl.replaceChildren();

  // Nagłówek: tytuł edytowalny wprost na kartce (to raport dla siebie — szybciej
  // poprawić w miejscu niż szukać pola w ustawieniach).
  const head = rpEl("header", "rp-head");
  const title = rpEl("h1", "rp-title", rpTitle);
  title.contentEditable = "plaintext-only";
  if (title.contentEditable !== "plaintext-only") title.contentEditable = "true";
  title.spellcheck = false;
  title.setAttribute("role", "textbox");
  title.setAttribute("aria-label", t("rpTitleAria"));
  title.addEventListener("input", () => { rpTitle = title.textContent.trim(); });
  title.addEventListener("keydown", (e) => { if (e.key === "Enter") { e.preventDefault(); title.blur(); } });
  const sheet = (typeof sheetSelect !== "undefined" && sheetSelect?.value) || "";
  const meta = rpEl("div", "rp-meta", [currentFileName || "", sheet, t("rpGenerated", { when: d.at.toLocaleString(rpLocale(), { dateStyle: "medium", timeStyle: "short" }) })].filter(Boolean).join(" · "));
  head.append(title, meta);

  // Zakres: co dokładnie jest na tej kartce.
  const scope = rpEl("section", `rp-scope${d.filtering ? " is-filtered" : ""}`);
  scope.append(
    rpEl("span", "rp-scope-label", t("rpScopeLabel")),
    rpEl("span", "rp-scope-text", d.filtering
      ? t("rpScopeFiltered", { view: d.viewText, rows: rpNum(d.rows, 0), total: rpNum(d.total, 0) })
      : t("rpScopeAll", { rows: rpNum(d.rows, 0), view: d.viewText })),
  );

  const parts = [head, scope];
  const active = rpActiveSections();
  RP_SECTIONS.forEach((sec) => {
    if (!active.includes(sec.id) || !rpSectionAvailable(sec, d)) return;
    const el = RP_BUILDERS[sec.id](d);
    if (el) {
      el.dataset.section = sec.id;
      parts.push(el);
    }
  });
  parts.push(rpEl("footer", "rp-foot", t("rpFooter")));
  rpPageEl.append(...parts);
  rpSyncControls();
  rpRenderContentPanel();
  rpFit();
}

// ── Sekcje ──────────────────────────────────────────────────────────────────

function rpSection(titleText, cls = "") {
  const sec = rpEl("section", `rp-block${cls ? ` ${cls}` : ""}`);
  if (titleText) sec.append(rpEl("h2", "rp-h2", titleText));
  return sec;
}

function rpBars(sec, items, total) {
  const maxN = Math.max(1, ...items.map(([, n]) => n));
  items.forEach(([label, n]) => {
    const row = rpEl("div", "rp-bar-row");
    const bar = rpEl("div", "rp-bar");
    const fill = rpEl("div", "rp-bar-fill");
    fill.style.width = `${Math.max(2, Math.round((n / maxN) * 100))}%`;
    bar.appendChild(fill);
    row.append(rpEl("div", "rp-bar-label", label), bar, rpEl("div", "rp-bar-value", `${rpNum(n, 0)} · ${rpPct(n, total)}%`));
    sec.appendChild(row);
  });
}

function rpTopWithRest(entries, max) {
  const shown = entries.slice(0, max);
  const rest = entries.slice(max).reduce((sum, [, n]) => sum + n, 0);
  if (rest) shown.push([t("rpChartOther"), rest]);
  return shown;
}

function rpTable(headers, rows, { numCols = [] } = {}) {
  const isNum = (i) => numCols.includes(i);
  const table = rpEl("table", "rp-table");
  const thead = rpEl("thead");
  const htr = rpEl("tr");
  headers.forEach((h, i) => htr.appendChild(rpEl("th", isNum(i) ? "is-num" : "", h)));
  thead.appendChild(htr);
  const tbody = rpEl("tbody");
  rows.forEach((cells) => {
    const tr = rpEl("tr");
    cells.forEach((c, i) => tr.appendChild(rpEl("td", isNum(i) ? "is-num" : "", c)));
    tbody.appendChild(tr);
  });
  table.append(thead, tbody);
  return table;
}

const RP_KIND_LABEL = { num: "rpKindNum", date: "rpKindDate", text: "rpKindText", empty: "rpKindEmpty" };

const RP_BUILDERS = {
  tiles(d) {
    const el = rpEl("section", "rp-tiles");
    rpTiles(d).forEach((tile) => {
      const box = rpEl("div", `rp-tile${tile.text ? " is-text" : ""}`);
      box.append(rpEl("div", "rp-tile-label", tile.label), rpEl("div", "rp-tile-value", tile.value));
      if (tile.sub) box.append(rpEl("div", "rp-tile-sub", tile.sub));
      el.appendChild(box);
    });
    return el;
  },
  findings(d) {
    const sec = rpSection(t("rpFindingsTitle"), "rp-findings");
    const list = rpEl("ul", "rp-list");
    // Krótki raport ma się zmieścić na kartce — tam limit; w dłuższych wszystkie wnioski.
    const all = rpFindings(d);
    const shown = rpPrefs.preset === "short" ? all.slice(0, RP_MAX_FINDINGS) : all;
    shown.forEach((f) => list.appendChild(rpEl("li", f.tone ? `is-${f.tone}` : "", f.text)));
    sec.appendChild(list);
    return sec;
  },
  chart(d) {
    const { col, entries, total } = d.category;
    const sec = rpSection(t("rpChartTitle", { col: col.name }), "rp-chart");
    rpBars(sec, rpTopWithRest(entries, RP_MAX_BARS), total);
    return sec;
  },
  months(d) {
    const dt = d.date;
    const sec = rpSection(t("rpMonthsTitle", { col: dt.col.name }), "rp-chart");
    let keys = Array.from(dt.months.keys()).sort();
    if (keys.length > RP_MAX_MONTHS) {
      keys = keys.slice(-RP_MAX_MONTHS);
      sec.appendChild(rpEl("p", "rp-note", t("rpMonthsLast", { n: RP_MAX_MONTHS })));
    }
    rpBars(sec, keys.map((k) => [rpMonthLabel(k), dt.months.get(k)]), dt.dates ? dt.dates.length : dt.col.dates.length);
    return sec;
  },
  numbers(d) {
    const sec = rpSection(t("rpNumbersTitle"));
    sec.appendChild(rpTable(
      [t("rpColColumn"), t("rpColCount"), t("rpColSum"), t("rpColAvg"), t("rpColMedian"), t("rpColMin"), t("rpColMax")],
      d.numericCols.map((s) => [s.col.name, rpNum(s.n, 0), rpNum(s.sum), rpNum(s.avg), rpNum(s.median), rpNum(s.min), rpNum(s.max)]),
      { numCols: [1, 2, 3, 4, 5, 6] },
    ));
    return sec;
  },
  categories(d) {
    const sec = rpSection(t("rpCategoriesTitle"), "rp-cats");
    const grid = rpEl("div", "rp-cat-grid");
    d.otherCategories.forEach(({ col, entries, total }) => {
      const box = rpEl("div", "rp-cat");
      box.appendChild(rpEl("h3", "rp-h3", col.name));
      rpBars(box, rpTopWithRest(entries, 5), total);
      grid.appendChild(box);
    });
    sec.appendChild(grid);
    return sec;
  },
  columns(d) {
    const sec = rpSection(t("rpColumnsTitle"));
    sec.appendChild(rpTable(
      [t("rpColColumn"), t("rpColKind"), t("rpColFill"), t("rpColUnique"), t("rpColTopValue")],
      d.cols.map((col) => {
        const top = Array.from(col.counts.entries()).sort((a, b) => b[1] - a[1])[0];
        return [
          col.name,
          t(RP_KIND_LABEL[col.kind]),
          `${Math.round(col.fill * 100)}%`,
          rpNum(col.unique, 0),
          // Każda wartość raz (np. „Nr”) → „najczęstsza” nic nie znaczy.
          !top ? "—" : top[1] === 1 ? t("rpAllDistinct") : `${afShortSafe(top[0], 28)} (${rpNum(top[1], 0)})`,
        ];
      }),
      { numCols: [2, 3] },
    ));
    return sec;
  },
  data(d) {
    const sec = rpSection(t("rpDataTitle"), "rp-data");
    const cols = d.used.slice(0, RP_MAX_DATA_COLS);
    const rows = d.rowsList.slice(0, RP_MAX_DATA_ROWS);
    const notes = [];
    if (d.used.length > cols.length) notes.push(t("rpDataColsCut", { shown: cols.length, all: d.used.length }));
    if (d.rowsList.length > rows.length) notes.push(t("rpDataRowsCut", { shown: rpNum(rows.length, 0), all: rpNum(d.rowsList.length, 0) }));
    if (notes.length) sec.appendChild(rpEl("p", "rp-note", notes.join(" ")));
    sec.appendChild(rpTable(
      cols.map((c) => c.name),
      rows.map((row) => cols.map((c) => String(getDisplayValue(row, c.c) ?? ""))),
    ));
    return sec;
  },
};

function afShortSafe(text, max) {
  const str = String(text || "").replace(/\s+/g, " ").trim();
  return str.length > max ? `${str.slice(0, max - 1)}…` : str;
}

function rpSectionAvailable(sec, d) {
  return !sec.has || sec.has(d);
}

function rpActiveSections() {
  return rpPrefs.preset !== "custom" && RP_PRESETS[rpPrefs.preset]
    ? RP_PRESETS[rpPrefs.preset]
    : rpPrefs.sections;
}

// ── Panel „Zawartość”: presety + lista sekcji ───────────────────────────────

function rpRenderContentPanel() {
  if (!rpPresetsEl || !rpSectionListEl || !rpData) return;
  rpPresetsEl.querySelectorAll("[data-preset]").forEach((btn) => {
    btn.setAttribute("aria-pressed", btn.dataset.preset === rpPrefs.preset ? "true" : "false");
  });
  const customEl = document.getElementById("rpPresetCustom");
  if (customEl) customEl.classList.toggle("hidden", rpPrefs.preset !== "custom");
  if (rpContentBtn) {
    const name = rpPrefs.preset === "custom" ? t("rpPresetCustom") : t(`rpPreset_${rpPrefs.preset}`);
    rpContentBtn.textContent = t("rpContentBtn", { preset: name });
  }
  const active = rpActiveSections();
  rpSectionListEl.replaceChildren();
  RP_SECTIONS.forEach((sec) => {
    const available = rpSectionAvailable(sec, rpData);
    const label = rpEl("label", `rp-sec${available ? "" : " is-unavailable"}`);
    const cb = rpEl("input");
    cb.type = "checkbox";
    cb.value = sec.id;
    cb.checked = active.includes(sec.id);
    cb.disabled = !available;
    cb.addEventListener("change", () => rpToggleSection(sec.id, cb.checked));
    const text = rpEl("span", "rp-sec-text", t(sec.label));
    label.append(cb, text);
    // Sekcja bez danych zostaje na liście (żeby było wiadomo, że istnieje), ale mówi czemu jej nie ma.
    if (!available) label.appendChild(rpEl("span", "rp-sec-why", t("rpSecUnavailable")));
    rpSectionListEl.appendChild(label);
  });
}

function rpSetPreset(preset) {
  if (!RP_PRESETS[preset]) return;
  rpPrefs.preset = preset;
  rpPrefs.sections = RP_PRESETS[preset].slice();
  rpSavePrefs();
  rpRender();
}

// Ręczna zmiana = „Własny”. Startujemy od tego, co było widać, więc nic nie przeskakuje.
function rpToggleSection(id, on) {
  const next = new Set(rpActiveSections());
  if (on) next.add(id);
  else next.delete(id);
  rpPrefs.sections = RP_SECTIONS.map((sec) => sec.id).filter((x) => next.has(x));
  const match = Object.keys(RP_PRESETS).find((k) => RP_PRESETS[k].join() === rpPrefs.sections.join());
  rpPrefs.preset = match || "custom";
  rpSavePrefs();
  rpRender();
}

function rpToggleContentPanel(force) {
  if (!rpContentPanelEl) return;
  const open = typeof force === "boolean" ? force : rpContentPanelEl.classList.contains("hidden");
  rpContentPanelEl.classList.toggle("hidden", !open);
  if (rpContentBtn) rpContentBtn.setAttribute("aria-expanded", open ? "true" : "false");
  if (open) rpFit();
}

// ── Podgląd: skalowanie kartki do okna + „czy mieści się na 1 stronie” ──────

function rpMmToPx(mm) {
  return (mm * 96) / 25.4;
}

function rpFit() {
  if (!rpIsOpen || !rpPageEl || !rpWrapEl || !rpStageEl) return;
  const pageW = rpMmToPx(RP_PAGE_MM.w);
  const avail = Math.max(200, rpStageEl.clientWidth - 24);
  const scale = Math.min(1, avail / pageW);
  rpPageEl.style.transform = scale < 1 ? `scale(${scale})` : "";
  const h = rpPageEl.offsetHeight;
  rpWrapEl.style.width = `${Math.round(pageW * scale)}px`;
  rpWrapEl.style.height = `${Math.round(h * scale)}px`;
  if (rpFitNoteEl) {
    // Kartka na ekranie ma min. 297 mm (stopka dociśnięta do dołu), więc liczymy
    // wysokość TREŚCI: dół ostatniej sekcji + stopka. Na wydruku góra/dół każdej strony
    // to margines @page, a kartka ma tylko RP_PRINT_PAD_TOP_MM u góry.
    const foot = rpPageEl.querySelector(".rp-foot");
    const last = foot ? foot.previousElementSibling : null;
    const px = rpMmToPx;
    const contentBottom = last && foot
      ? last.offsetTop + last.offsetHeight + px(6) + foot.offsetHeight
      : h;
    const printed = contentBottom - px(RP_SCREEN_PAD_MM) + px(RP_PRINT_PAD_TOP_MM);
    const usable = px(RP_PAGE_MM.h - 2 * RP_PRINT_MARGIN_MM);
    const pages = Math.max(1, Math.ceil((printed - 2) / usable));
    rpFitNoteEl.textContent = pages > 1 ? t("rpFitPages", { n: pages }) : t("rpFitOne");
    rpFitNoteEl.classList.toggle("is-warn", pages > 1);
    rpDrawPageGuides(pages, usable);
  }
}

// Przerywane linie „tu kończy się strona N” — przybliżenie (przeglądarka przy druku
// przenosi całe bloki, więc realny podział bywa odrobinę wyżej).
function rpDrawPageGuides(pages, usablePx) {
  rpPageEl.querySelectorAll(".rp-page-break-guide").forEach((el) => el.remove());
  const firstTop = rpMmToPx(RP_SCREEN_PAD_MM - RP_PRINT_PAD_TOP_MM);
  for (let i = 1; i < pages; i++) {
    const g = document.createElement("div");
    g.className = "rp-page-break-guide";
    g.setAttribute("aria-hidden", "true");
    g.dataset.label = t("rpPageGuide", { n: i + 1 });
    g.style.top = `${Math.round(firstTop + usablePx * i)}px`;
    rpPageEl.appendChild(g);
  }
}

// ── Kontrolki wyglądu ───────────────────────────────────────────────────────

function rpSyncControls() {
  if (rpStyleEl) rpStyleEl.value = rpPrefs.style;
  if (rpContentBtn) rpContentBtn.addEventListener("click", () => rpToggleContentPanel());
if (rpPresetsEl) {
  rpPresetsEl.addEventListener("click", (e) => {
    const btn = e.target.closest("[data-preset]");
    if (btn) rpSetPreset(btn.dataset.preset);
  });
}
if (rpAccentsEl) {
    rpAccentsEl.querySelectorAll("[data-accent]").forEach((btn) => {
      const on = btn.dataset.accent === rpPrefs.accent;
      btn.setAttribute("aria-pressed", on ? "true" : "false");
      // Styl „Oszczędny” drukuje w czerni — kolor nic by nie zmienił, więc go wyłączamy
      // zamiast udawać, że działa.
      btn.disabled = rpPrefs.style === "ink";
    });
  }
  if (rpSizeBtn) {
    rpSizeBtn.textContent = rpPrefs.size === "large" ? "A+" : "A";
    rpSizeBtn.setAttribute("aria-pressed", rpPrefs.size === "large" ? "true" : "false");
  }
}

function rpSetPref(key, value) {
  rpPrefs[key] = value;
  rpSavePrefs();
  rpRender();
}

// ── Otwarcie / zamknięcie / druk ────────────────────────────────────────────

function rpBackgroundInert(on) {
  document.querySelectorAll(".app, .hero-overlay").forEach((el) => {
    if (on) el.setAttribute("inert", "");
    else el.removeAttribute("inert");
  });
}

function openReport() {
  if (!rpOverlayEl) return;
  if (!Array.isArray(currentHeaders) || !currentHeaders.length || !Array.isArray(viewRows) || !viewRows.length) {
    toast(t("noDataForExport"), "warning");
    return;
  }
  rpLoadPrefs();
  rpData = rpCollect();
  rpTitle = rpDefaultTitle();
  rpReturnFocusEl = document.activeElement;
  rpOverlayEl.classList.remove("hidden");
  document.body.classList.add("rp-active");
  rpBackgroundInert(true);
  rpIsOpen = true;
  rpRender();
  if (rpPrintBtn) rpPrintBtn.focus();
}

function closeReport() {
  if (!rpOverlayEl || !rpIsOpen) return;
  rpToggleContentPanel(false);
  rpIsOpen = false;
  rpOverlayEl.classList.add("hidden");
  document.body.classList.remove("rp-active");
  rpBackgroundInert(false);
  const back = rpReturnFocusEl;
  rpReturnFocusEl = null;
  if (back && document.contains(back) && !back.closest("[inert]")) back.focus();
  else if (reportBtn) reportBtn.focus();
}

function rpAfterPrint() {
  document.body.classList.remove("rp-printing");
  if (rpDocTitleBefore) document.title = rpDocTitleBefore;
  rpDocTitleBefore = "";
}

function printReport() {
  if (!rpIsOpen) return;
  // Tytuł dokumentu = domyślna nazwa pliku PDF w dialogu „Zapisz jako PDF”.
  rpDocTitleBefore = document.title;
  document.title = rpTitle || rpDefaultTitle();
  document.body.classList.add("rp-printing");
  try {
    window.print();
  } finally {
    // Safari potrafi wrócić z print() zanim dialog się zamknie — sprzątamy też w afterprint.
    setTimeout(() => { if (!window.matchMedia("print").matches) rpAfterPrint(); }, 0);
  }
}

window.addEventListener("afterprint", rpAfterPrint);

if (reportBtn) reportBtn.addEventListener("click", openReport);
if (rpCloseBtn) rpCloseBtn.addEventListener("click", closeReport);
if (rpPrintBtn) rpPrintBtn.addEventListener("click", printReport);
if (rpStyleEl) rpStyleEl.addEventListener("change", () => rpSetPref("style", RP_STYLES.includes(rpStyleEl.value) ? rpStyleEl.value : "modern"));
if (rpSizeBtn) rpSizeBtn.addEventListener("click", () => rpSetPref("size", rpPrefs.size === "large" ? "normal" : "large"));
if (rpAccentsEl) {
  rpAccentsEl.addEventListener("click", (e) => {
    const btn = e.target.closest("[data-accent]");
    if (btn && !btn.disabled && RP_ACCENTS[btn.dataset.accent]) rpSetPref("accent", btn.dataset.accent);
  });
}
if (rpOverlayEl) {
  // Capture: Esc nie może dotrzeć do globalnych skrótów tabeli pod spodem. Esc w trakcie
  // edycji tytułu tylko kończy edycję.
  document.addEventListener("keydown", (e) => {
    if (!rpIsOpen || e.key !== "Escape") return;
    e.preventDefault();
    e.stopPropagation();
    if (document.activeElement && document.activeElement.classList?.contains("rp-title")) {
      document.activeElement.blur();
      return;
    }
    if (rpContentPanelEl && !rpContentPanelEl.classList.contains("hidden")) {
      rpToggleContentPanel(false);
      if (rpContentBtn) rpContentBtn.focus();
      return;
    }
    closeReport();
  }, true);
}
window.addEventListener("resize", () => { if (rpIsOpen) rpFit(); });

// Hook testowy.
window.__report = {
  open: openReport,
  close: closeReport,
  print: printReport,
  data: () => rpData,
  findings: () => (rpData ? rpFindings(rpData).map((f) => f.text) : []),
  prefs: () => ({ ...rpPrefs, sections: rpPrefs.sections.slice() }),
  sections: () => Array.from(document.querySelectorAll("#rpPage [data-section]")).map((el) => el.dataset.section),
  setPreset: rpSetPreset,
  toggleSection: rpToggleSection,
  isOpen: () => rpIsOpen,
};
