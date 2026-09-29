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
  { id: "aggAuto", label: "rpSecAggAuto", has: (d) => !!(d.aggs && d.aggs.length) },
  { id: "months", label: "rpSecMonths", has: (d) => !!(d.date && d.date.months.size >= 2) },
  { id: "numbers", label: "rpSecNumbers", has: (d) => d.numericCols.length > 0 },
  { id: "categories", label: "rpSecCategories", has: (d) => d.otherCategories.length > 0 },
  { id: "aggPanel", label: "rpSecAggPanel", has: (d) => !!d.aggPanel },
  { id: "columns", label: "rpSecColumns" },
  { id: "data", label: "rpSecData" },
];
const RP_PRESETS = {
  short: ["tiles", "findings", "chart"],
  normal: ["tiles", "findings", "chart", "aggAuto", "months", "numbers", "categories"],
  detailed: RP_SECTIONS.map((sec) => sec.id),
};
const RP_MAX_BARS = 6;
// A4 w obu orientacjach. Pozioma przydaje się przy szerokich tabelach (tryb „tabela” z Eksportu).
const RP_ORIENTS = ["portrait", "landscape"];
const RP_TABLE_MODE_MAX_ROWS = 3000; // wydruk tabeli: więcej = setki stron; pełne dane daje CSV/Excel
let rpMode = "report";       // "report" | "table" (druk tabeli z okna Eksport)
let rpTableSpec = null;      // { model, cols } w trybie tabeli
let rpTableOrient = null;    // orientacja w trybie tabeli (na tę sesję, nie zmienia ustawień raportu)
let rpPageCount = 1;
let rpBusy = false;
function rpOrient() {
  if (rpMode === "table") return rpTableOrient === "landscape" ? "landscape" : "portrait";
  return rpPrefs.orient === "landscape" ? "landscape" : "portrait";
}
function rpPageMm() {
  return rpOrient() === "landscape" ? { w: 297, h: 210 } : { w: 210, h: 297 };
}
// Bezpieczny margines druku (góra/dół KAŻDEJ strony) = @page rpA4 w app.css. Słabsze drukarki
// nie drukują 5–6 mm od krawędzi; 15 mm to zapas z nawiązką. Boki = padding kartki (15 mm).
// Marginesy jak w Wordzie — do wyboru. v = góra/dół KAŻDEJ strony (@page rpA4-*), h = boki
// (padding kartki). Nawet „wąskie” (12,7 mm, jak w Wordzie) są ponad 2× szersze niż martwa
// strefa słabszej drukarki (5–6 mm), więc nic nie zostanie ucięte.
const RP_MARGINS = {
  narrow: { v: 12.7, h: 12.7 },
  normal: { v: 20, h: 20 },
  wide: { v: 25.4, h: 32 },
};
function rpMargin() {
  return RP_MARGINS[rpPrefs.margin] || RP_MARGINS.normal;
}
const RP_SHEET_GAP_MM = 8;      // szara przerwa między kartkami w podglądzie

// Podpowiedzi z nazw nagłówków (PL + EN). Tylko podbijają wynik — kolumna i tak musi
// mieć odpowiednie wartości.
// Dwa poziomy: „stan sprawy” (status, etap) mówi o danych więcej niż „gdzie/jaki” (miasto, typ),
// więc przy dwóch kandydatach głównym wykresem zostaje status.
const RP_CAT_STRONG_RE = /status|stan\b|etap|faza|priorytet|state|stage|priority/i;
const RP_CAT_RE = /typ|rodzaj|kategor|grupa|dział|dzial|miasto|region|oddział|type|category|group|city|branch/i;
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
const rpMarginEl = document.getElementById("rpMargin");
const rpOrientEl = document.getElementById("rpOrient");
const rpPdfBtn = document.getElementById("rpPdfBtn");
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
let rpPrefs = { style: "modern", accent: "green", size: "normal", margin: "normal", orient: "portrait", preset: "short", sections: RP_PRESETS.short.slice() };
let rpDocTitleBefore = "";

// ── Ustawienia wyglądu ──────────────────────────────────────────────────────

function rpLoadPrefs() {
  try {
    const p = JSON.parse(localStorage.getItem(RP_PREFS_KEY) || "{}");
    rpPrefs = {
      style: RP_STYLES.includes(p.style) ? p.style : "modern",
      accent: RP_ACCENTS[p.accent] ? p.accent : "green",
      size: RP_SIZES.includes(p.size) ? p.size : "normal",
      margin: RP_MARGINS[p.margin] ? p.margin : "normal",
      orient: RP_ORIENTS.includes(p.orient) ? p.orient : "portrait",
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
    groups: new Map(),
  }));
  rows.forEach((row) => {
    cols.forEach((col) => {
      const shown = String(getDisplayValue(row, col.c) ?? "").trim();
      if (!shown) return;
      col.nonEmpty += 1;
      col.counts.set(shown, (col.counts.get(shown) || 0) + 1);
      // Warianty pisowni tej samej wartości („aktywny / Aktywny / AKTYWNY”, podwójna spacja)
      // liczymy RAZEM — tym samym kluczem co silnik agregacji (normalizeAnalysisKey), inaczej
      // kafelek mówiłby „aktywny 17%”, a zestawienie obok „AKTYWNY 48%”.
      const key = rpValueKey(shown);
      let g = col.groups.get(key);
      if (!g) { g = { n: 0, spell: new Map() }; col.groups.set(key, g); }
      g.n += 1;
      g.spell.set(shown, (g.spell.get(shown) || 0) + 1);
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
    col.uniqueGroups = col.groups.size;
    col.fill = rows.length ? col.nonEmpty / rows.length : 0;
    col.kind = !col.nonEmpty ? "empty"
      : col.nums.length / col.nonEmpty >= 0.8 ? "num"
        : col.dates.length / col.nonEmpty >= 0.8 ? "date"
          : "text";
  });
  return cols;
}

function rpValueKey(shown) {
  const k = typeof normalizeAnalysisKey === "function" ? normalizeAnalysisKey(shown) : String(shown).toLowerCase().replace(/\s+/g, " ").trim();
  return k || shown;
}

// Grupy wariantów kolumny jako [[etykieta, liczba]] — etykieta = najczęstsza pisownia.
function rpGroupEntries(col) {
  return Array.from(col.groups.values())
    .map((g) => [Array.from(g.spell.entries()).sort((a, b) => b[1] - a[1])[0][0], g.n, g.spell.size > 1 ? Array.from(g.spell.keys()) : null])
    .sort((a, b) => b[1] - a[1]);
}

// Numer porządkowy (1, 2, 3…) to też liczby, ale ich suma nic nie mówi.
function rpLooksLikeCounter(col) {
  if (RP_ID_RE.test(col.name)) return true;
  // „Kwota”, „Ilość”… z wartościami po kolei (100, 101, 102) to nadal kwoty, nie numeracja.
  if (RP_NUM_RE.test(col.name)) return false;
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
      && col.uniqueGroups >= 2 && col.uniqueGroups <= 12 && col.uniqueGroups <= col.nonEmpty * 0.6)
    .map((col) => {
      const grouped = rpGroupEntries(col);
      return {
        col,
        score: (RP_CAT_STRONG_RE.test(col.name) ? 10 : RP_CAT_RE.test(col.name) ? 5 : 0) + col.fill * 3 - col.uniqueGroups * 0.1,
        entries: grouped.map(([label, n]) => [label, n]),
        variants: grouped.filter((g) => g[2]).map(([label, n, spells]) => ({ label, n, spells })),
        total: col.nonEmpty,
      };
    })
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

// Liczby, których suma nic nie znaczy: sumy kontrolne / hashe, rok (2022+2023+… = bzdura).
const RP_NOT_MEASURE_RE = /kontroln|checksum|hash|\bcrc\b|\bpesel\b|\bnip\b|\bregon\b|telefon|\bphone\b|kod poczt|zip/i;
const RP_YEAR_RE = /^(rok|year|lata)\b/i;
function rpNotAMeasure(col) {
  if (RP_NOT_MEASURE_RE.test(col.name) || RP_YEAR_RE.test(col.name)) return true;
  // Same liczby całkowite z zakresu lat i mało różnych wartości = rok, nie ilość.
  return col.nums.length >= 3 && col.unique <= 60 && col.nums.every((n) => Number.isInteger(n) && n >= 1900 && n <= 2100);
}

// Bloki powtarzane (Kw1_Kwota, Kw2_Kwota…; Kwota1, Kwota2…) rozpoznaje parser apki
// (parseRepeatedHeader). Taka rodzina to JEDNA miara — sumujemy wszystkie bloki, zamiast
// brać pierwszą z brzegu kolumnę albo „Sumę kontrolną”.
function rpNumericFamilies(cols) {
  if (typeof parseRepeatedHeader !== "function") return [];
  const byBase = new Map();
  cols.forEach((col) => {
    const raw = String(currentHeaders[col.c] ?? "");
    const rh = parseRepeatedHeader(raw);
    if (!rh || !rh.base || rh.base === raw.trim()) return;
    const middle = /^[A-Za-zĄąĆćĘęŁłŃńÓóŚśŹźŻż]{1,6}\d+[_\-. ]/.test(raw.trim());
    const pretty = middle ? (rh.base.split("_").slice(1).join("_") || rh.base) : rh.base;
    const key = rpValueKey(rh.base);
    if (!byBase.has(key)) byBase.set(key, { pretty, members: [] });
    byBase.get(key).members.push(col);
  });
  return Array.from(byBase.values())
    .filter((f) => f.members.length >= 2)
    .map((f) => ({
      c: null,
      isFamily: true,
      members: f.members,
      name: t("rpFamilyName", { base: f.pretty, n: f.members.length }),
      nums: f.members.flatMap((m) => m.nums),
      fill: f.members.reduce((s2, m) => s2 + m.fill, 0) / f.members.length,
      unique: Infinity,
    }));
}

// Miary liczbowe warte statystyk, najlepsza pierwsza: rodziny bloków + pojedyncze kolumny
// spoza rodzin. Bez numerów porządkowych, lat i sum kontrolnych.
function rpNumericCols(cols) {
  const eligible = cols.filter((col) => col.kind === "num" && !rpLooksLikeCounter(col) && !rpNotAMeasure(col));
  const families = rpNumericFamilies(eligible);
  const inFamily = new Set(families.flatMap((f) => f.members));
  return [...families, ...eligible.filter((col) => !inFamily.has(col))]
    .map((col) => ({ col, score: (RP_NUM_RE.test(col.name) ? 10 : 0) + (col.isFamily ? 3 : 0) + col.fill * 3 }))
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
    // Kwoty z groszami to nie identyfikatory — dwie równe kwoty to nie duplikat rekordu.
    if (col.kind === "num" && !RP_ID_RE.test(col.name) && col.nums.some((n) => !Number.isInteger(n))) continue;
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

// Agregacje dokładamy PO zebraniu podstaw (potrzebują wybranej kategorii i liczby).
function rpCollectAll() {
  const d = rpCollect();
  d.aggs = rpAutoAggregations(d);
  d.aggPanel = rpPanelAggregation();
  return d;
}

// ── Agregacje z silnika apki ────────────────────────────────────────────────
// Raport nie liczy grup sam — pożycza silnik panelu „Agregacje” (buildAggregationWorkbenchResult).
// Silnik czyta ustawienia z globalnego aggregationWorkbenchState i po drodze je normalizuje,
// więc: kopia stanu → nasze ustawienia → wynik → przywrócenie KAŻDEGO pola. Panel użytkownika
// zostaje dokładnie taki, jaki był (test to pilnuje).
// Zawsze: bieżący widok (scopeMode "filtered"), wiersz nagłówka = ten z tabeli (bez
// autodetekcji — inaczej grupy mogłyby wyjść z innych kolumn niż reszta raportu),
// układ szeroki, bez scalania grup i progów.

const RP_AGG_MAX_GROUPS = 8;
const RP_AGG_MAX_CROSS_COLS = 5;
const RP_AGG_PANEL_MAX = 15;

function rpAggAvailable() {
  return typeof buildAggregationWorkbenchResult === "function" && typeof aggregationWorkbenchState !== "undefined";
}

function rpWithAggState(override, fn) {
  if (!rpAggAvailable()) return null;
  const saved = JSON.parse(JSON.stringify(aggregationWorkbenchState));
  try {
    Object.assign(aggregationWorkbenchState, override);
    return fn();
  } catch {
    return null; // raport ma się otworzyć nawet, gdy agregacja na tym arkuszu się wyłoży
  } finally {
    Object.keys(aggregationWorkbenchState).forEach((k) => { if (!(k in saved)) delete aggregationWorkbenchState[k]; });
    Object.assign(aggregationWorkbenchState, saved);
  }
}

function rpAggBase() {
  return {
    sourceMode: "wide",
    scopeMode: "filtered",
    headerRowChoice: "custom",
    customHeaderRow: currentHeaderRow,
    groupBy: "",
    groupBy2: "",
    groupBy3: "",
    groupMode: "exact",
    groupPattern: "=*",
    havingMode: "all",
    measureFilterMode: "all",
    measureFilterValue: "",
  };
}

// Uruchom silnik i sprawdź, czy policzył TO, o co prosiliśmy (silnik po cichu podmienia
// niepasujące ustawienia na domyślne — wtedy wynik byłby o czymś innym niż tytuł sekcji).
function rpAggRun(override) {
  return rpWithAggState({ ...rpAggBase(), ...override }, () => {
    const res = buildAggregationWorkbenchResult();
    const st = aggregationWorkbenchState;
    const asked = (k) => override[k] === undefined || JSON.stringify(override[k]) === JSON.stringify(st[k]);
    if (!res || res.status !== "ok" || !["groupBy", "groupBy2", "measures", "aggregation"].every(asked)) return null;
    return res;
  });
}

function rpAggProbe() {
  return rpWithAggState({ ...rpAggBase(), measures: ["count_rows"], aggregation: "count" }, () => {
    const res = buildAggregationWorkbenchResult();
    return res ? { measures: res.measures || [] } : null;
  });
}

function rpFmtKind(value, kind) {
  if (kind === "duration" && typeof formatDurationDays === "function") return formatDurationDays(value);
  return rpNum(value);
}

function rpAutoAggregations(d) {
  if (!rpAggAvailable() || !d.category || d.rows < 3) return [];
  const probe = rpAggProbe();
  if (!probe) return [];
  const out = [];
  const catHeader = currentHeaders[d.category.col.c];

  // 1. Główna liczba wg głównej kategorii: suma, średnia, udział.
  // Rodzina bloków (Kw1…Kw4_Kwota) to nie jedna kolumna silnika — tu liczymy sami:
  // suma bloków w wierszu, grupy po tym samym kluczu co reszta raportu.
  if (d.numeric && d.numeric.col.isFamily) {
    const fam = d.numeric.col;
    const catC = d.category.col.c;
    const groups = new Map();
    d.rowsList.forEach((row) => {
      let rowSum = 0;
      let has = false;
      fam.members.forEach((m) => {
        const v = row.values ? row.values[m.c] : undefined;
        if (typeof v === "number" && Number.isFinite(v)) { rowSum += v; has = true; }
      });
      if (!has) return;
      const shown = String(getDisplayValue(row, catC) ?? "").trim() || t("rpEmptyGroup");
      const key = rpValueKey(shown);
      let g = groups.get(key);
      if (!g) { g = { spell: new Map(), count: 0, sum: 0 }; groups.set(key, g); }
      g.spell.set(shown, (g.spell.get(shown) || 0) + 1);
      g.count += 1;
      g.sum += rowSum;
    });
    const entries = Array.from(groups.values()).map((g) => ({
      label: Array.from(g.spell.entries()).sort((a, b) => b[1] - a[1])[0][0],
      count: g.count,
      sum: g.sum,
      average: g.sum / g.count,
    })).sort((a, b) => b.sum - a.sum);
    if (entries.length) {
      const total = entries.reduce((s2, e) => s2 + e.sum, 0);
      out.push({ type: "byGroup", title: t("rpAggByGroup", { measure: fam.name, group: d.category.col.name }), kind: "number", entries, total, measureName: fam.name });
    }
  } else if (d.numeric) {
    const m = probe.measures.find((x) => x.measureType === "column" && x.colIdx === d.numeric.col.c);
    const res = m && rpAggRun({ groupBy: catHeader, measures: [m.key], aggregation: "sum" });
    if (res) {
      const total = res.entries.reduce((sum, e) => sum + (e.sum || 0), 0);
      out.push({ type: "byGroup", title: t("rpAggByGroup", { measure: d.numeric.col.name, group: d.category.col.name }), kind: m.kind, entries: res.entries, total, measureName: d.numeric.col.name });
    }
  }

  // 2. Czas trwania start → koniec (para kolumn, którą wykrywa silnik) wg kategorii.
  const dur = probe.measures.find((x) => x.measureType === "date_range") || probe.measures.find((x) => x.kind === "duration" && x.measureType === "column");
  if (dur) {
    const res = rpAggRun({ groupBy: catHeader, measures: [dur.key], aggregation: "avg" });
    if (res && res.entries.some((e) => Number.isFinite(e.average))) {
      out.push({ type: "duration", title: t("rpAggDuration", { measure: dur.label.replace("->", "→"), group: d.category.col.name }), entries: res.entries.filter((e) => Number.isFinite(e.average)), measureName: dur.label.replace("->", "→") });
    }
  }

  // 3. Tabela krzyżowa: główna kategoria × druga kategoria (liczba wierszy).
  const second = d.otherCategories[0];
  if (second) {
    const res = rpAggRun({ groupBy: catHeader, groupBy2: currentHeaders[second.col.c], measures: ["count_rows"], aggregation: "count" });
    if (res) {
      out.push({ type: "cross", title: t("rpAggCross", { a: d.category.col.name, b: second.col.name }), entries: res.entries, rowsName: d.category.col.name });
    }
  }
  return out;
}

const RP_AGG_METHODS = ["count", "sum", "avg", "median", "min", "max", "distinct", "earliest", "latest"];

// Agregacja dokładnie tak, jak ustawiona w panelu — to „skąd ma brać” w rękach użytkownika.
function rpPanelAggregation() {
  if (!rpAggAvailable()) return null;
  return rpWithAggState({}, () => {
    const res = buildAggregationWorkbenchResult();
    if (!res || res.status !== "ok") return null;
    const st = aggregationWorkbenchState;
    const measures = res.selectedMeasures && res.selectedMeasures.length ? res.selectedMeasures : [res.measure].filter(Boolean);
    const kind = typeof getPrimaryAggregationValueKind === "function" ? getPrimaryAggregationValueKind(measures, st.aggregation) : "number";
    return {
      entries: res.entries,
      kind,
      method: RP_AGG_METHODS.includes(st.aggregation) ? st.aggregation : "count",
      measureNames: measures.map((m) => m.label),
      groups: [st.groupBy, st.groupBy2, st.groupBy3].filter(Boolean),
      wholeSheet: st.scopeMode === "all",
    };
  });
}

function rpAggFindings(d) {
  const out = [];
  (d.aggs || []).forEach((a) => {
    if (a.type === "byGroup" && a.entries.length >= 2 && a.total > 0) {
      const top = a.entries[0];
      out.push({ text: t("rpFindAggShare", { measure: a.measureName, group: top.label, pct: rpPct(top.sum, a.total), value: rpFmtKind(top.sum, a.kind) }) });
      const solid = a.entries.filter((e) => e.count >= 3 && Number.isFinite(e.average) && e.average > 0);
      if (solid.length >= 2) {
        const hi = solid.reduce((x, y) => (y.average > x.average ? y : x));
        const lo = solid.reduce((x, y) => (y.average < x.average ? y : x));
        if (hi !== lo && hi.average >= lo.average * 1.5) {
          out.push({ text: t("rpFindAggSpread", { measure: a.measureName, hi: hi.label, hiV: rpFmtKind(hi.average, a.kind), x: rpNum(hi.average / lo.average, 1), lo: lo.label, loV: rpFmtKind(lo.average, a.kind) }) });
        }
      }
    }
    if (a.type === "duration") {
      const solid = a.entries.filter((e) => e.count >= 3);
      const n = a.entries.reduce((s2, e) => s2 + e.count, 0);
      const overall = n ? a.entries.reduce((s2, e) => s2 + e.average * e.count, 0) / n : NaN;
      if (solid.length >= 2 && Number.isFinite(overall)) {
        const hi = solid.reduce((x, y) => (y.average > x.average ? y : x));
        out.push({ text: t("rpFindAggLongest", { group: hi.label, measure: a.measureName, avg: rpFmtKind(hi.average, "duration"), all: rpFmtKind(overall, "duration") }) });
      }
    }
  });
  return out;
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
    // Ta sama wartość zapisana różnie — policzone razem, ale warto o tym wiedzieć (i poprawić).
    const v = d.category.variants || [];
    if (v.length) {
      const list = v.slice(0, 2).map((g) => g.spells.slice(0, 4).map((x) => `„${x}”`).join(" / ")).join("; ");
      out.push({ text: t("rpFindVariants", { col: col.name, list, n: v.length }), tone: "warn" });
    }
  }
  if (d.numeric) {
    const s = d.numeric;
    out.push({ text: t("rpFindSum", { col: s.col.name, sum: rpNum(s.sum), avg: rpNum(s.avg), min: rpNum(s.min), max: rpNum(s.max) }) });
  }
  out.push(...rpAggFindings(d));
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
  rpPageEl.dataset.margin = RP_MARGINS[rpPrefs.margin] ? rpPrefs.margin : "normal";
  rpPageEl.dataset.orient = rpOrient();
  rpPageEl.style.setProperty("--rp-pw", `${rpPageMm().w}mm`);
  rpPageEl.style.setProperty("--rp-mv", `${rpMargin().v}mm`);
  rpPageEl.style.setProperty("--rp-mh", `${rpMargin().h}mm`);
  const accent = RP_ACCENTS[rpPrefs.accent] || RP_ACCENTS.green;
  rpPageEl.style.setProperty("--rp-accent", accent);
  // Odcienie akcentu jako zwykłe rgb (nie color-mix) — patrz komentarz przy --rp-tint w CSS.
  const [ar, ag, ab] = [1, 3, 5].map((i) => parseInt(accent.slice(i, i + 2), 16));
  const mixWhite = (c, a) => Math.round(255 + (c - 255) * a);
  rpPageEl.style.setProperty("--rp-tint", `rgb(${mixWhite(ar, 0.09)}, ${mixWhite(ag, 0.09)}, ${mixWhite(ab, 0.09)})`);
  rpPageEl.style.setProperty("--rp-accent-soft", `rgba(${ar}, ${ag}, ${ab}, 0.08)`);
  rpPageEl.style.setProperty("--rp-accent-ring", `rgba(${ar}, ${ag}, ${ab}, 0.4)`);
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
  if (rpMode === "table") {
    parts.push(rpBuildTableModeSection());
  }
  const active = rpMode === "table" ? [] : rpActiveSections();
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
  rpFitTables();
  rpSyncControls();
  if (rpMode !== "table") rpRenderContentPanel();
  rpFit();
}

// ── Tryb „tabela” (druk z okna Eksport) ──
// Stary „Drukuj / PDF” z Eksportu budował gołą tabelę bez marginesów, bez orientacji i bez
// informacji o filtrze. Teraz idzie przez ten sam podgląd co raport: te same marginesy,
// łamanie stron, powtórzony nagłówek, dopasowanie szerokości i pobieranie PDF.
function rpBuildTableModeSection() {
  const { model, cols } = rpTableSpec;
  const sec = rpEl("section", "rp-block rp-data");
  const rows = model.rows.slice(0, RP_TABLE_MODE_MAX_ROWS);
  if (model.rows.length > rows.length) {
    sec.appendChild(rpEl("p", "rp-note", t("rpTableModeCut", { shown: rpNum(rows.length, 0), all: rpNum(model.rows.length, 0) })));
  }
  const label = (ci) => (typeof exportColLabel === "function" ? exportColLabel(model.headers[ci], ci) : String(model.headers[ci] ?? ""));
  sec.appendChild(rpTable(
    cols.map(label),
    rows.map((row) => cols.map((ci) => String(getDisplayValue(row, ci) ?? ""))),
  ));
  return sec;
}

function rpCollectScope(model) {
  const view = typeof captureViewState === "function" ? captureViewState() : null;
  return {
    rows: model.rows.length,
    total: Array.isArray(baseRows) ? baseRows.length : model.rows.length,
    filtering: !!(typeof lastAppliedFilters !== "undefined" && lastAppliedFilters && lastAppliedFilters.filtering),
    viewText: view && typeof describeViewState === "function" ? describeViewState(view) : "",
    at: new Date(),
  };
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
    cells.forEach((c, i) => {
      // Nierozdzielne: jedno słowo/data/kod albo krótka wartość („Klient 7”, „W toku”).
      // Dłuższe opisy zawijają się normalnie na spacjach.
      const txt = String(c).trim();
      const token = !/\s/.test(txt) || txt.length <= 20;
      tr.appendChild(rpEl("td", [isNum(i) ? "is-num" : "", token ? "is-token" : ""].filter(Boolean).join(" "), c));
    });
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
  aggAuto(d) {
    const sec = rpSection(t("rpAggTitle"), "rp-aggs");
    d.aggs.forEach((a) => {
      const box = rpEl("div", "rp-agg");
      box.appendChild(rpEl("h3", "rp-h3", a.title));
      if (a.type === "byGroup") {
        const shown = a.entries.slice(0, RP_AGG_MAX_GROUPS);
        box.appendChild(rpTable(
          [t("rpColGroup"), t("rpColCount"), t("rpColSum"), t("rpColAvg"), t("rpColShare")],
          shown.map((e) => [e.label, rpNum(e.count, 0), rpFmtKind(e.sum, a.kind), rpFmtKind(e.average, a.kind), `${rpPct(e.sum, a.total)}%`]),
          { numCols: [1, 2, 3, 4] },
        ));
        if (a.entries.length > shown.length) box.appendChild(rpEl("p", "rp-note", t("rpAggMore", { n: a.entries.length - shown.length })));
      } else if (a.type === "duration") {
        const shown = a.entries.slice().sort((x, y) => y.average - x.average).slice(0, RP_AGG_MAX_GROUPS);
        box.appendChild(rpTable(
          [t("rpColGroup"), t("rpColCount"), t("rpColAvg"), t("rpColMedian"), t("rpColMax")],
          shown.map((e) => [e.label, rpNum(e.count, 0), rpFmtKind(e.average, "duration"), rpFmtKind(e.median, "duration"), rpFmtKind(e.max, "duration")]),
          { numCols: [1, 2, 3, 4] },
        ));
      } else if (a.type === "cross") {
        box.appendChild(rpCrossTable(a));
      }
      sec.appendChild(box);
    });
    return sec;
  },
  aggPanel(d) {
    const a = d.aggPanel;
    // Jeden wykres = jeden blok, który się nie łamie (inaczej tytuł z opisem zostaje sam na dole strony).
    const sec = rpSection(t("rpAggPanelTitle"), "rp-agg-panel");
    const how = t("rpAggPanelHow", {
      method: t(`rpAggMethod_${a.method}`),
      measure: a.measureNames.join(", "),
      groups: a.groups.join(" / "),
    });
    sec.appendChild(rpEl("p", "rp-note", a.wholeSheet && d.filtering ? `${how} ${t("rpAggPanelWhole")}` : how));
    const shown = a.entries.slice(0, RP_AGG_PANEL_MAX);
    const maxV = Math.max(1e-9, ...shown.map((e) => Math.abs(Number(e.primary) || 0)));
    shown.forEach((e) => {
      const row = rpEl("div", "rp-bar-row");
      const bar = rpEl("div", "rp-bar");
      const fill = rpEl("div", "rp-bar-fill");
      const v = Number(e.primary) || 0;
      // Daty (najwcześniej/najpóźniej) nie mają sensownej długości paska.
      fill.style.width = a.kind === "date" ? "0" : `${Math.max(2, Math.round((Math.abs(v) / maxV) * 100))}%`;
      bar.appendChild(fill);
      const value = typeof formatAggregationMetricValue === "function" ? formatAggregationMetricValue(v, a.kind) : rpNum(v);
      row.append(rpEl("div", "rp-bar-label", e.label), bar, rpEl("div", "rp-bar-value", value));
      sec.appendChild(row);
    });
    if (a.entries.length > shown.length) sec.appendChild(rpEl("p", "rp-note", t("rpAggMore", { n: a.entries.length - shown.length })));
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
        const top = col.groups.size ? rpGroupEntries(col)[0] : null;
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

// Tabela krzyżowa z wyniku silnika grupującego po dwóch kolumnach („A / B”).
// Wiersze i kolumny: najliczniejsze, reszta zbiorczo w „inne”.
function rpCrossTable(a) {
  const rowsTot = new Map();
  const colsTot = new Map();
  const cell = new Map();
  a.entries.forEach((e) => {
    const [r, c] = e.groupLabels || [e.label, ""];
    rowsTot.set(r, (rowsTot.get(r) || 0) + e.count);
    colsTot.set(c, (colsTot.get(c) || 0) + e.count);
    cell.set(`${r}\u0000${c}`, e.count);
  });
  const top = (m, n) => Array.from(m.entries()).sort((x, y) => y[1] - x[1]).slice(0, n).map(([k]) => k);
  const rows = top(rowsTot, RP_AGG_MAX_GROUPS);
  const cols = top(colsTot, RP_AGG_MAX_CROSS_COLS);
  const colOther = colsTot.size > cols.length;
  const headers = [a.rowsName, ...cols, ...(colOther ? [t("rpChartOther")] : []), t("rpColTotal")];
  const body = rows.map((r) => {
    const vals = cols.map((c) => cell.get(`${r}\u0000${c}`) || 0);
    const other = rowsTot.get(r) - vals.reduce((x, y) => x + y, 0);
    return [r, ...vals.map((v) => (v ? rpNum(v, 0) : "·")), ...(colOther ? [other ? rpNum(other, 0) : "·"] : []), rpNum(rowsTot.get(r), 0)];
  });
  const numCols = headers.map((_, i) => i).filter((i) => i > 0);
  return rpTable(headers, body, { numCols });
}

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
  const pageW = rpMmToPx(rpPageMm().w);
  const avail = Math.max(200, rpStageEl.clientWidth - 24);
  const scale = Math.min(1, avail / pageW);
  rpPageEl.style.transform = scale < 1 ? `scale(${scale})` : "";
  rpMeasureScale = scale;
  const pages = rpPaginate();
  const h = rpPageEl.offsetHeight;
  rpWrapEl.style.width = `${Math.round(pageW * scale)}px`;
  rpWrapEl.style.height = `${Math.round(h * scale)}px`;
  if (rpFitNoteEl) {
    rpFitNoteEl.textContent = pages > 1 ? t("rpFitPages", { n: pages }) : t("rpFitOne");
    rpFitNoteEl.classList.toggle("is-warn", pages > 1);
  }
}

// ── Tabele za szerokie na kartkę → pomniejszenie (jak „dopasuj do strony” w Excelu) ──
// Słowa łamią się tylko na spacjach. Gdy tabela i tak nie mieści się w szerokości kartki
// (dużo kolumn + szerokie marginesy + duży tekst), pomniejszamy czcionkę TEJ tabeli krokami
// po 5% (min. 65%), zamiast łamać słowa w środku („Zakończon-e”). Dopiero gdy nawet to nie
// pomoże (bardzo długie słowa/adresy), dopuszczamy łamanie w środku słowa (.is-tight).
const RP_TABLE_MIN_SCALE = 0.65;

function rpFitTables() {
  rpPageEl.querySelectorAll("table.rp-table").forEach((table) => {
    const box = table.parentElement;
    const avail = box.clientWidth;
    table.classList.remove("is-tight");
    table.style.fontSize = "";
    box.querySelectorAll(":scope > .rp-shrunk-note").forEach((n) => n.remove());
    if (!avail || table.offsetWidth <= avail + 0.5) return;
    let scale = 1;
    while (table.offsetWidth > avail + 0.5 && scale > RP_TABLE_MIN_SCALE + 0.001) {
      scale = Math.max(RP_TABLE_MIN_SCALE, scale - 0.05);
      table.style.fontSize = `${scale}em`;
    }
    if (table.offsetWidth > avail + 0.5) table.classList.add("is-tight");
    const note = document.createElement("p");
    note.className = "rp-note rp-shrunk-note";
    note.textContent = t("rpTableShrunk", { pct: Math.round(scale * 100) });
    table.after(note);
  });
}

// ── Podgląd stronami (= wydruk) ──
// Drukarka nie drukuje przy samej krawędzi kartki (słabsze nawet 5–6 mm), więc każda
// strona ma bezpieczny margines (rpMargin().v) u góry i u dołu (@page rpA4-*) oraz
// 15 mm po bokach (padding kartki). Podgląd pokazuje DOKŁADNIE te strony: łamiemy treść
// tak, jak zrobi to przeglądarka przy druku, i w miejscach podziału wstawiamy przekładki
// (tylko na ekranie) — koniec strony z marginesem, szara przerwa między kartkami,
// margines nowej strony. Dzięki temu widać, że nic nie leży przy krawędzi.
// Zasady łamania (jak w druku): bloki z break-inside: avoid przechodzą w całości na
// następną stronę; zestawienia łamią się między tabelami; tabela danych między wierszami,
// z powtórzonym nagłówkiem. Liczba stron wychodzi przy okazji (test porównuje ją z PDF).

// Pomiary UŁAMKOWE (getBoundingClientRect). offsetTop/offsetHeight zaokrąglają do pełnych
// pikseli — przy 40 wierszach tabeli na stronę błąd sumował się do ~1 wiersza i podgląd
// łamał stronę w innym miejscu niż druk. Dzielimy przez skalę podglądu (transform).
let rpMeasureScale = 1;
function rpRelTop(el) {
  return (el.getBoundingClientRect().top - rpPageEl.getBoundingClientRect().top) / rpMeasureScale;
}
function rpHeight(el) {
  return el.getBoundingClientRect().height / rpMeasureScale;
}

// Jednostki, których druk nie rozetnie, w kolejności. `start` = od czego zaczyna się nowa
// strona (tytuł sekcji jedzie razem z pierwszym elementem), `row` = wiersz tabeli danych.
function rpPrintUnits() {
  const units = [];
  const isGap = (el) => el.classList.contains("rp-sheet-gap") || el.classList.contains("rp-gap-row") || el.classList.contains("rp-head-repeat") || el.classList.contains("rp-sheet-band");
  Array.from(rpPageEl.children).forEach((el) => {
    if (isGap(el)) return;
    if (!el.matches(".rp-aggs, .rp-data")) { units.push({ start: el, end: el }); return; }
    let lead = null;
    Array.from(el.children).forEach((ch) => {
      if (isGap(ch)) return;
      if (ch.matches(".rp-h2, .rp-note")) { lead = lead || ch; return; }
      if (ch.matches("table") && el.matches(".rp-data") && ch.tBodies[0]) {
        Array.from(ch.tBodies[0].rows).forEach((tr, i) => {
          if (isGap(tr)) return;
          units.push(i === 0 ? { start: lead || ch, end: tr } : { start: tr, end: tr, row: tr, table: ch });
        });
      } else {
        units.push({ start: lead || ch, end: ch });
      }
      lead = null;
    });
    // Notka PO tabeli (np. „tabela pomniejszona…”) też musi zmieścić się na stronie.
    if (lead) units.push({ start: lead, end: lead });
  });
  return units;
}

function rpClearSheetGaps() {
  rpPageEl.querySelectorAll(".rp-sheet-gap, .rp-gap-row, .rp-head-repeat, .rp-sheet-band").forEach((el) => el.remove());
  rpPageEl.style.minHeight = "";
}

// Szara przerwa między kartkami = osobna warstwa na KARTCE (nie w przekładce): zawsze na
// całą szerokość, także gdy podział wypada w środku tabeli (komórka tabeli przycinała
// pasek do szerokości tabeli i dawała mu tło wiersza).
function rpGapBand(boundary, gapH, nextPage) {
  const band = document.createElement("div");
  band.className = "rp-sheet-band";
  band.setAttribute("aria-hidden", "true");
  band.style.top = `${boundary}px`;
  band.style.height = `${gapH}px`;
  band.dataset.label = t("rpPageGuide", { n: nextPage });
  return band;
}

// Ustaw wysokość przekładki tak, żeby `target` zaczynał się DOKŁADNIE na górze obszaru druku.
// Dwa przebiegi: pierwszy zgrubnie, drugi koryguje ułamki (zaokrąglenia układu, marginesy).
function rpSettle(spacer, target, wantTop) {
  let h = 0;
  spacer.style.height = "0px";
  for (let pass = 0; pass < 2; pass++) {
    h = Math.max(0, h + (wantTop - rpRelTop(target)));
    spacer.style.height = `${h}px`;
  }
}

function rpPaginate() {
  rpClearSheetGaps();
  const px = rpMmToPx;
  const P = px(rpPageMm().h);
  const G = px(RP_SHEET_GAP_MM);
  const M = px(rpMargin().v);
  const pageEnd = (k) => k * (P + G) + P - M;      // dół obszaru druku strony k
  const contentTop = (k) => k * (P + G) + M;       // góra obszaru druku strony k
  let k = 0;
  const bands = [];
  const foot = rpPageEl.querySelector(".rp-foot");
  if (foot) foot.classList.remove("is-dropped");
  rpPrintUnits().forEach((u) => {
    const top = rpRelTop(u.start);
    const bottom = rpRelTop(u.end) + rpHeight(u.end);
    // Bez tolerancji „na plus”: blok kończący się ułamek piksela za granicą druk przenosi dalej
    // (tak wylądowała sama stopka na nowej stronie, choć podgląd mówił, że się mieści).
    if (bottom <= pageEnd(k) - 0.5) return;
    // Stopka (jedna linijka „raport policzony na tym urządzeniu”) nie jest warta osobnej
    // kartki papieru — gdy nie mieści się na ostatniej stronie, po prostu jej nie ma
    // (także w druku i w PDF: klasa działa wszędzie).
    if (u.start === foot) {
      foot.classList.add("is-dropped");
      return;
    }
    if (top <= contentTop(k) + 1) {
      // Blok wyższy niż strona, już stoi na górze strony — druk i tak go potnie.
      while (bottom > pageEnd(k) + 0.5) k += 1;
      return;
    }
    k += 1;
    const boundary = (k - 1) * (P + G) + P;
    if (u.row) {
      // Wiersz tabeli: przekładka-wiersz + powtórzony nagłówek (jak thead w druku).
      const cols = u.row.cells.length || 1;
      const gapRow = document.createElement("tr");
      gapRow.className = "rp-gap-row";
      gapRow.setAttribute("aria-hidden", "true");
      const td = document.createElement("td");
      td.colSpan = cols;
      const filler = document.createElement("div");
      filler.className = "rp-gap-filler";
      td.appendChild(filler);
      gapRow.appendChild(td);
      u.row.parentNode.insertBefore(gapRow, u.row);
      let first = u.row;
      const headRow = u.table.tHead && u.table.tHead.rows[0];
      if (headRow) {
        // Kopia nagłówka ze ZWYKŁYCH komórek (td.rp-th): ogólne style apki dla <th> w <tbody>
        // (numery wierszy głównej tabeli) rozciągały kopię z <th> do ~47 px zamiast 28 px —
        // i podgląd mieścił na stronie o wiersz mniej niż druk. Wysokość = prawdziwy thead.
        const clone = document.createElement("tr");
        clone.className = "rp-head-repeat";
        clone.setAttribute("aria-hidden", "true");
        Array.from(headRow.cells).forEach((th) => {
          const td = document.createElement("td");
          td.className = `rp-th${th.classList.contains("is-num") ? " is-num" : ""}`;
          td.textContent = th.textContent;
          clone.appendChild(td);
        });
        clone.style.height = `${rpHeight(u.table.tHead)}px`;
        u.row.parentNode.insertBefore(clone, u.row);
        first = clone;
      }
      rpSettle(filler, first, contentTop(k));
      bands.push(rpGapBand(boundary, G, k + 1));
    } else {
      const gap = document.createElement("div");
      gap.className = "rp-sheet-gap";
      gap.setAttribute("aria-hidden", "true");
      u.start.parentNode.insertBefore(gap, u.start);
      rpSettle(gap, u.start, contentTop(k));
      bands.push(rpGapBand(boundary, G, k + 1));
    }
  });
  // Warstwy dopiero na końcu — absolutne, nie ruszają układu, więc pomiary wyżej są czyste.
  bands.forEach((b) => rpPageEl.appendChild(b));
  const pages = k + 1;
  rpPageCount = pages;
  rpPageEl.style.minHeight = `${Math.round(pages * (P + G) - G)}px`;
  return pages;
}

// ── Kontrolki wyglądu ───────────────────────────────────────────────────────

function rpSyncControls() {
  if (rpStyleEl) rpStyleEl.value = rpPrefs.style;
  if (rpMarginEl) rpMarginEl.value = rpPrefs.margin;
  if (rpOrientEl) rpOrientEl.value = rpOrient();
  // W trybie tabeli nie ma sekcji do wybierania.
  if (rpContentBtn) rpContentBtn.classList.toggle("hidden", rpMode === "table");
  if (rpContentBtn) rpContentBtn.addEventListener("click", () => rpToggleContentPanel());
document.getElementById("rpLookBtn")?.addEventListener("click", (e) => {
  const open = !rpOverlayEl.classList.contains("look-open");
  rpOverlayEl.classList.toggle("look-open", open);
  e.currentTarget.setAttribute("aria-expanded", open ? "true" : "false");
  rpFit();
});
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
  rpMode = "report";
  rpTableSpec = null;
  rpData = rpCollectAll();
  rpTitle = rpDefaultTitle();
  rpShowOverlay();
}

// Druk / PDF tabeli z okna Eksport: wybrane kolumny, wszystkie wiersze widoku.
// Szeroka tabela (> 6 kolumn) startuje poziomo — i tak można przełączyć.
function openReportTable({ cols } = {}) {
  if (!rpOverlayEl) return;
  const model = (typeof currentDisplayModel !== "undefined" && currentDisplayModel) || getDisplayModel();
  if (!model?.headers?.length || !model?.rows?.length) {
    toast(t("noDataForExport"), "warning");
    return;
  }
  const useCols = Array.isArray(cols) && cols.length ? cols : model.headers.map((_, i) => i);
  rpLoadPrefs();
  rpMode = "table";
  rpTableSpec = { model, cols: useCols };
  rpTableOrient = useCols.length > 6 ? "landscape" : "portrait";
  rpData = rpCollectScope(model);
  const base = currentFileName ? currentFileName.replace(/\.[^.]+$/, "") : t("rpTitleFallback");
  rpTitle = t("rpTableTitleDefault", { name: base });
  rpShowOverlay();
}

function rpShowOverlay() {
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
  rpAfterPrint();
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

// iPhone/iPad (iPadOS udaje Maca — stąd maxTouchPoints).
function rpIsIos() {
  const ua = navigator.userAgent || "";
  return /iPad|iPhone|iPod/.test(ua) || (/Macintosh/.test(ua) && navigator.maxTouchPoints > 1);
}

let rpPrintStarted = false;
window.addEventListener("beforeprint", () => { rpPrintStarted = true; });

function printReport() {
  if (!rpIsOpen) return;
  // Tytuł dokumentu = domyślna nazwa pliku PDF w dialogu „Zapisz jako PDF”.
  if (!rpDocTitleBefore) rpDocTitleBefore = document.title;
  document.title = rpTitle || rpDefaultTitle();
  // Klasa zostaje aż do „afterprint” albo zamknięcia raportu. NIE zdejmujemy jej zaraz po
  // print(): na iPadzie print() wraca natychmiast, a stronę do druku bierze CHWILĘ później —
  // bez klasy szła pusta kartka. Na ekranie ta klasa nic nie zmienia (działa tylko w @media print).
  document.body.classList.add("rp-printing");
  rpPrintStarted = false;
  window.print();
  // W apce dodanej do ekranu początkowego iOS potrafi zignorować print() — okno druku się
  // nie otwiera i „nic się nie dzieje”. Wtedy robimy PDF i dajemy arkusz „Udostępnij”
  // (jest w nim „Drukuj”). Tylko na iOS i tylko gdy druk faktycznie nie ruszył.
  if (rpIsIos()) {
    setTimeout(() => {
      if (rpPrintStarted || !rpIsOpen) return;
      toast(t("rpPrintIosFallback"), "info");
      rpDownloadPdf({ intent: "print" });
    }, 1200);
  }
}

// ── Pobierz PDF (bez okna drukowania) ──────────────────────────────────────
// „Drukuj / PDF” otwiera systemowe okno druku — na iPadzie to kilka kroków, zanim powstanie
// plik. Tu dostajesz gotowy plik .pdf do zapisania/wysłania i wydrukowania później.
// Jak: każda strona z podglądu (który już jest pocięty dokładnie jak wydruk, z marginesami)
// jest renderowana do obrazu (html2canvas, ładowane dopiero przy pierwszym użyciu), a z obrazów
// składamy PDF ręcznie — to kilkadziesiąt bajtów struktury, nie potrzeba 400 KB biblioteki.
// Uczciwie: tekst w takim PDF to obraz (nie da się go zaznaczyć). Do PDF z zaznaczalnym
// tekstem zostaje „Drukuj / PDF” → „Zapisz jako PDF”.
const RP_PDF_SCALE = 2;          // ~190 dpi na A4 — ostro na wydruku, rozsądny rozmiar pliku
const RP_PDF_JPEG_QUALITY = 0.9;
let rpH2cPromise = null;

function rpEnsureHtml2canvas() {
  if (typeof window.html2canvas === "function") return Promise.resolve();
  if (rpH2cPromise) return rpH2cPromise;
  rpH2cPromise = new Promise((resolve, reject) => {
    const sc = document.createElement("script");
    // Bez ?v= — tak samo jak xlsx/jszip: service worker trzyma go w cache „ciężkich” (offline).
    sc.src = "lib/html2canvas.min.js";
    sc.async = true;
    sc.onload = () => resolve();
    sc.onerror = () => { rpH2cPromise = null; reject(new Error("html2canvas")); };
    document.head.appendChild(sc);
  });
  return rpH2cPromise;
}

function rpCanvasToJpeg(canvas) {
  return new Promise((resolve, reject) => {
    canvas.toBlob((blob) => {
      if (!blob) { reject(new Error("toBlob")); return; }
      blob.arrayBuffer().then((buf) => resolve(new Uint8Array(buf)), reject);
    }, "image/jpeg", RP_PDF_JPEG_QUALITY);
  });
}

// Minimalny PDF 1.4: każda strona = jeden obraz JPEG (DCTDecode) na całą stronę.
function rpBuildPdf(images, wPt, hPt) {
  const enc = new TextEncoder();
  const chunks = [];
  const offsets = [];
  let len = 0;
  const push = (data) => {
    const bytes = typeof data === "string" ? enc.encode(data) : data;
    chunks.push(bytes);
    len += bytes.length;
  };
  const obj = (n, write) => {
    offsets[n] = len;
    push(`${n} 0 obj\n`);
    write();
    push("\nendobj\n");
  };
  const W = wPt.toFixed(2);
  const H = hPt.toFixed(2);
  push("%PDF-1.4\n%âãÏÓ\n");
  obj(1, () => push("<< /Type /Catalog /Pages 2 0 R >>"));
  const kids = images.map((_, i) => `${3 + i * 3} 0 R`).join(" ");
  obj(2, () => push(`<< /Type /Pages /Kids [${kids}] /Count ${images.length} >>`));
  images.forEach((im, i) => {
    const pageN = 3 + i * 3;
    obj(pageN, () => push(`<< /Type /Page /Parent 2 0 R /MediaBox [0 0 ${W} ${H}] /Resources << /XObject << /Im${i} ${pageN + 2} 0 R >> >> /Contents ${pageN + 1} 0 R >>`));
    const content = `q ${W} 0 0 ${H} 0 0 cm /Im${i} Do Q`;
    obj(pageN + 1, () => { push(`<< /Length ${content.length} >>\nstream\n${content}\nendstream`); });
    obj(pageN + 2, () => {
      push(`<< /Type /XObject /Subtype /Image /Width ${im.w} /Height ${im.h} /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /DCTDecode /Length ${im.bytes.length} >>\nstream\n`);
      push(im.bytes);
      push("\nendstream");
    });
  });
  const size = 3 + images.length * 3;
  const xref = len;
  let table = `xref\n0 ${size}\n0000000000 65535 f \n`;
  for (let n = 1; n < size; n++) table += `${String(offsets[n]).padStart(10, "0")} 00000 n \n`;
  push(table);
  push(`trailer\n<< /Size ${size} /Root 1 0 R >>\nstartxref\n${xref}\n%%EOF\n`);
  return new Blob(chunks, { type: "application/pdf" });
}

function rpSafeFileName(name, ext) {
  const base = String(name || "raport").replace(/[\\/:*?"<>|\u0000-\u001f]+/g, "-").replace(/\s+/g, " ").trim().slice(0, 90) || "raport";
  return `${base}.${ext}`;
}

// Pobranie pliku. iPad/iPhone jako zainstalowana apka (standalone) często nie obsługuje
// zwykłego „pobierz” — tam arkusz udostępniania („Zapisz w Plikach”, „Drukuj”, AirDrop).
// Na iPhonie/iPadzie arkusz „Udostępnij” (Zachowaj w Plikach, Drukuj, AirDrop) zamiast
// „pobierz” — w apce z ekranu początkowego zwykłe pobieranie często NIC nie robi.
// Arkusz wolno otworzyć tylko tuż po dotknięciu. Gdy plik powstawał dłużej (PDF z kilku stron,
// ładowanie biblioteki Excela), iOS odrzuca share() — wtedy pasek „Plik gotowy” z przyciskiem:
// świeże dotknięcie = arkusz się otworzy.
let rpReadyEl = null;
let rpReadyUrl = "";

function rpHideReady() {
  if (rpReadyEl) rpReadyEl.classList.add("hidden");
  if (rpReadyUrl) { URL.revokeObjectURL(rpReadyUrl); rpReadyUrl = ""; }
}

function rpShowReady(file, label) {
  if (!rpReadyEl) {
    rpReadyEl = document.createElement("div");
    rpReadyEl.className = "file-ready hidden";
    rpReadyEl.setAttribute("role", "status");
    const text = document.createElement("span");
    text.className = "file-ready-text";
    const go = document.createElement("button");
    go.type = "button";
    go.className = "btn btn-sm file-ready-go";
    const x = document.createElement("button");
    x.type = "button";
    x.className = "btn btn-sm ghost file-ready-x";
    x.textContent = "✕";
    x.setAttribute("aria-label", t("rpReadyClose"));
    x.addEventListener("click", rpHideReady);
    rpReadyEl.append(text, go, x);
    document.body.appendChild(rpReadyEl);
  }
  rpReadyEl.querySelector(".file-ready-text").textContent = t("rpReadyText", { name: file.name });
  const go = rpReadyEl.querySelector(".file-ready-go");
  go.textContent = label || t("rpReadySave");
  go.onclick = async () => {
    try {
      await navigator.share({ files: [file], title: file.name });
      rpHideReady();
    } catch (e) {
      if (e && e.name === "AbortError") return;
      // Ostatnia deska: otwórz plik (PDF otworzy się w przeglądarce z własnym „Udostępnij”).
      if (rpReadyUrl) URL.revokeObjectURL(rpReadyUrl);
      rpReadyUrl = URL.createObjectURL(file);
      window.open(rpReadyUrl, "_blank");
    }
  };
  rpReadyEl.classList.remove("hidden");
  go.focus();
}

async function rpDeliverFile(blob, fileName, { label } = {}) {
  if (rpIsIos() && typeof File === "function" && navigator.canShare) {
    const file = new File([blob], fileName, { type: blob.type });
    if (navigator.canShare({ files: [file] })) {
      const fresh = navigator.userActivation ? navigator.userActivation.isActive : false;
      if (fresh) {
        try {
          await navigator.share({ files: [file], title: fileName });
          return "shared";
        } catch (e) {
          if (e && e.name === "AbortError") return "aborted";
        }
      }
      rpShowReady(file, label);
      return "ready";
    }
  }
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = fileName;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 30000);
  return "downloaded";
}

async function rpDownloadPdf({ deliver = true, intent = "save" } = {}) {
  if (!rpIsOpen || rpBusy) return null;
  rpBusy = true;
  const buttons = [rpPdfBtn, rpPrintBtn].filter(Boolean);
  buttons.forEach((b) => { b.disabled = true; });
  const noteBefore = rpFitNoteEl ? rpFitNoteEl.textContent : "";
  const transformBefore = rpPageEl.style.transform;
  try {
    if (rpFitNoteEl) rpFitNoteEl.textContent = t("rpPdfLoading");
    await rpEnsureHtml2canvas();
    const { w, h } = rpPageMm();
    const px = rpMmToPx;
    const P = px(h);
    const G = px(RP_SHEET_GAP_MM);
    const W = px(w);
    // Render bez skali podglądu; reszty apki nie klonujemy (szybciej i mniej pamięci na iPadzie),
    // ale arkusze stylów zostają.
    rpPageEl.style.transform = "";
    const keep = (el) => el === rpOverlayEl || rpOverlayEl.contains(el) || el.contains(rpOverlayEl)
      || /^(HEAD|LINK|STYLE|META|TITLE)$/.test(el.tagName) || !!el.closest?.("head");
    const images = [];
    for (let k = 0; k < rpPageCount; k++) {
      if (rpFitNoteEl) rpFitNoteEl.textContent = t("rpPdfProgress", { n: k + 1, all: rpPageCount });
      await new Promise((r) => setTimeout(r, 0)); // oddaj wątek — pasek postępu ma się odświeżyć
      const canvas = await window.html2canvas(rpPageEl, {
        scale: RP_PDF_SCALE,
        backgroundColor: "#ffffff",
        x: 0,
        y: Math.round(k * (P + G)),
        width: Math.round(W),
        height: Math.round(P),
        logging: false,
        ignoreElements: (el) => !keep(el) || el.classList?.contains("rp-sheet-band"),
      });
      images.push({ bytes: await rpCanvasToJpeg(canvas), w: canvas.width, h: canvas.height });
      canvas.width = 0; // zwolnij pamięć od razu (iPad)
      canvas.height = 0;
    }
    const mmToPt = (mm) => (mm / 25.4) * 72;
    const blob = rpBuildPdf(images, mmToPt(w), mmToPt(h));
    const name = rpSafeFileName(rpTitle || rpDefaultTitle(), "pdf");
    if (deliver) {
      const how = await rpDeliverFile(blob, name, { label: intent === "print" ? t("rpReadyPrint") : "" });
      if (how === "shared" || how === "downloaded") toast(t("rpPdfDone", { n: images.length, name }), "success");
    }
    return { blob, name, pages: images.length };
  } catch (e) {
    console.warn("PDF", e);
    toast(t("rpPdfFailed"), "error");
    return null;
  } finally {
    rpPageEl.style.transform = transformBefore;
    if (rpFitNoteEl) rpFitNoteEl.textContent = noteBefore;
    buttons.forEach((b) => { b.disabled = false; });
    rpBusy = false;
  }
}

window.addEventListener("afterprint", rpAfterPrint);

if (reportBtn) reportBtn.addEventListener("click", openReport);
if (rpCloseBtn) rpCloseBtn.addEventListener("click", closeReport);
if (rpPrintBtn) rpPrintBtn.addEventListener("click", printReport);
if (rpStyleEl) rpStyleEl.addEventListener("change", () => rpSetPref("style", RP_STYLES.includes(rpStyleEl.value) ? rpStyleEl.value : "modern"));
if (rpOrientEl) {
  rpOrientEl.addEventListener("change", () => {
    const v = RP_ORIENTS.includes(rpOrientEl.value) ? rpOrientEl.value : "portrait";
    if (rpMode === "table") { rpTableOrient = v; rpRender(); } else rpSetPref("orient", v);
  });
}
if (rpPdfBtn) rpPdfBtn.addEventListener("click", () => rpDownloadPdf());
if (rpMarginEl) rpMarginEl.addEventListener("change", () => rpSetPref("margin", RP_MARGINS[rpMarginEl.value] ? rpMarginEl.value : "normal"));
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
  openTable: openReportTable,
  downloadPdf: (opts) => rpDownloadPdf(opts),
  mode: () => rpMode,
  pageCount: () => rpPageCount,
  data: () => rpData,
  findings: () => (rpData ? rpFindings(rpData).map((f) => f.text) : []),
  prefs: () => ({ ...rpPrefs, sections: rpPrefs.sections.slice() }),
  sections: () => Array.from(document.querySelectorAll("#rpPage [data-section]")).map((el) => el.dataset.section),
  setPreset: rpSetPreset,
  // Strażnik marginesu: każda jednostka treści leży w całości w obszarze druku swojej strony
  // (≥ margines strony od górnej i dolnej krawędzi kartki). Zwraca listę naruszeń.
  marginViolations: () => {
    const P = rpMmToPx(rpPageMm().h);
    const G = rpMmToPx(RP_SHEET_GAP_MM);
    const M = rpMmToPx(rpMargin().v);
    return rpPrintUnits().map((u) => {
      const top = rpRelTop(u.start);
      const bottom = rpRelTop(u.end) + rpHeight(u.end);
      const k = Math.floor(top / (P + G));
      const ok = top >= k * (P + G) + M - 1 && bottom <= k * (P + G) + P - M + 1;
      return ok ? null : { page: k + 1, top: Math.round(top), bottom: Math.round(bottom), text: u.start.textContent.slice(0, 40) };
    }).filter(Boolean);
  },
  aggState: () => (typeof aggregationWorkbenchState !== "undefined" ? JSON.stringify(aggregationWorkbenchState) : ""),
  toggleSection: rpToggleSection,
  isOpen: () => rpIsOpen,
};
