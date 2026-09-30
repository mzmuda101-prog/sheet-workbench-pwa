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
  { id: "state", label: "rpSecState", has: (d) => !!d.state },
  { id: "compare", label: "rpSecCompare", has: (d) => !!d.compare },
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
  // Sekcje kątów (Stan teraz / Porównanie) dokłada wybór kąta albo ręczne zaznaczenie — długość raportu ich nie włącza.
  detailed: RP_SECTIONS.map((sec) => sec.id).filter((id) => id !== "state" && id !== "compare"),
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
let rpPrefs = { style: "modern", accent: "green", size: "normal", margin: "normal", orient: "portrait", preset: "short", angle: "overview", sections: RP_PRESETS.short.slice() };
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
      angle: ["overview", "state", "compare"].includes(p.angle) ? p.angle : "overview",
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

// Odmiana polska: 2–4 strony, 5–21 stron, 22–24 strony… (EN ma tylko jedną formę „many”).
function rpPluralKey(base, n) {
  let form = "other";
  try { form = new Intl.PluralRules(rpLocale()).select(n); } catch { /* stare przeglądarki */ }
  return form === "few" ? `${base}Few` : base;
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

// ── Zakres raportu ──────────────────────────────────────────────────────────
// Raport nie musi liczyć tego samego, co widać w tabeli. Zakres to jedno z:
//  • "view"  — jak w tabeli (domyślnie: filtry, szukanie, tryby auto, sortowanie),
//  • "all"   — cały arkusz (sortowanie z tabeli zostaje),
//  • "query" — zapytanie w składni szybkiego szukania z operatorami, np. Status:="W toku".
// Podpowiedzi to też zwykłe zapytania: po kliknięciu stoją w polu, więc nic nie dzieje się
// „po cichu”, a przy okazji widać składnię. Tabela zostaje nietknięta, dopóki użytkownik sam
// nie kliknie „Pokaż w tabeli”. Liczymy tym samym dopasowaniem co filtr tabeli
// (rowMatchesTextFilter), więc ten sam tekst daje te same wiersze tu i tam.

const RP_SCOPE_MAX_SUGGEST = 5;
const RP_SCOPE_MIN_ROWS = 3;       // podpowiedź na 1–2 wiersze to nie przegląd, tylko wyszukanie
const RP_SCOPE_FRESH_DAYS = 45;    // „ostatnie 30 dni” tylko, gdy w danych są świeże daty
const RP_QUERY_UNSAFE_RE = /&&|\|\||[{}"]/;
let rpScope = { kind: "view", q: "" };
const rpScopeBySheet = new Map();  // plik + arkusz → zakres (na tę sesję)
let rpSuggestCache = null;
let rpScopeCustomOpen = false;
let rpScopeMsg = "";
let rpScopeTimer = 0;

const rpScopeBarEl = document.getElementById("rpScopeBar");
const rpScopeChipsEl = document.getElementById("rpScopeChips");
const rpScopeRow2El = document.getElementById("rpScopeRow2");
const rpScopeQueryEl = document.getElementById("rpScopeQuery");
const rpScopeStatusEl = document.getElementById("rpScopeStatus");
const rpScopeApplyEl = document.getElementById("rpScopeApply");

function rpSheetId() {
  return [currentFileName || "", currentSheetName || "", currentHeaderRow || 0].join("\u001f");
}

function rpQueryRows(q) {
  const query = String(q || "").trim();
  if (!query || !Array.isArray(baseRows)) return [];
  const criteria = [{
    query: normalizeTermForMode(query, "contains"),
    mode: "contains",
    headers: currentHeaders,
    indexes: resolveIndexes(currentHeaders, new Set()),
    emptyMode: "all",
    negated: false,
    operatorsEnabled: true,
  }];
  try {
    return baseRows.filter((row) => rowMatchesTextFilter(row, criteria, false));
  } catch {
    return [];
  }
}

// Kolejność jak w tabeli (tabela danych w raporcie ma wyglądać znajomo).
function rpSortedLikeTable(rows) {
  if (typeof sortRowsForHeaders === "function") sortRowsForHeaders(rows, currentHeaders);
  return rows;
}

function rpScopeRows(scope = rpScope) {
  if (scope.kind === "all") return rpSortedLikeTable(baseRows.slice());
  if (scope.kind === "query") return rpSortedLikeTable(rpQueryRows(scope.q));
  return Array.isArray(viewRows) ? viewRows : [];
}

// Zakres jako „stan widoku” (kształt jak captureViewState) — do opisu na kartce, do porównania
// z tabelą i do „Pokaż w tabeli” (applyViewState).
function rpScopeViewState(scope = rpScope) {
  const table = typeof captureViewState === "function" ? captureViewState() : { v: 1, filter: null, sort: [] };
  if (scope.kind === "view") return table;
  const filter = scope.kind === "query"
    ? { f1: { q: scope.q, mode: "contains", neg: false, empty: "all", ops: true, cols: [], qcols: typeof vsQueryColumns === "function" ? vsQueryColumns(scope.q, true) : [] } }
    : null;
  return { v: 1, filter, sort: table.sort || [] };
}

function rpScopeDiffers() {
  if (rpScope.kind === "view" || typeof viewFilterKey !== "function") return false;
  return viewFilterKey(rpScopeViewState()) !== viewFilterKey(captureViewState());
}

// Nagłówek nadaje się do zapytania „Kolumna:…”, gdy jednoznacznie wskazuje JEDNĄ kolumnę
// i nie ma w nim znaków, które parser czyta jako operatory.
function rpQueryHeader(c) {
  const h = String(currentHeaders[c] ?? "").trim();
  if (!h || RP_QUERY_UNSAFE_RE.test(h) || /^[!@]/.test(h)) return "";
  const key = normalizeHeaderKey(h);
  return currentHeaders.filter((x) => normalizeHeaderKey(x) === key).length === 1 ? h : "";
}

function rpIsoDay(d) {
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
}

// Podpowiedzi zakresu — liczone z CAŁEGO arkusza (nie znikają po wybraniu jednej z nich).
// Każda musi dawać sensowny wycinek: co najmniej kilka wierszy i mniej niż wszystko.
function rpScopeSuggestions() {
  const total = Array.isArray(baseRows) ? baseRows.length : 0;
  const key = [rpSheetId(), typeof sheetDataStamp === "number" ? sheetDataStamp : 0, total, currentLang].join("\u001f");
  if (rpSuggestCache && rpSuggestCache.key === key) return rpSuggestCache.list;
  const list = [];
  const add = (s) => {
    if (list.length >= RP_SCOPE_MAX_SUGGEST || list.some((x) => x.q === s.q)) return;
    const n = rpQueryRows(s.q).length;
    if (n < RP_SCOPE_MIN_ROWS || n >= total) return;
    list.push({ ...s, n });
  };
  if (total >= RP_SCOPE_MIN_ROWS * 2) {
    const cols = rpProfileColumns(baseRows);
    // 1) Stan sprawy (Status/Etap…) — a bez takiej kolumny najlepsza kategoria, o ile ma mało grup.
    const cats = rpCategories(cols, total);
    const cat = cats.find((c) => RP_CAT_STRONG_RE.test(c.col.name)) || (cats[0] && cats[0].col.uniqueGroups <= 6 ? cats[0] : null);
    const ch = cat ? rpQueryHeader(cat.col.c) : "";
    if (ch) {
      cat.entries.filter(([, n]) => n < total * 0.9).slice(0, 3).forEach(([label, n]) => {
        const v = String(label).trim();
        if (!v || RP_QUERY_UNSAFE_RE.test(v)) return;
        add({ kind: "cat", q: `${ch}:="${v}"`, label: v, why: t("rpScopeWhyCat", { col: ch, value: v, pct: rpPct(n, cat.total) }) });
      });
    }
    // 2) Czas: ostatnie 30 dni (tylko przy świeżych danych) i najnowszy miesiąc w danych.
    const dp = rpPickDate(cols);
    const dh = dp ? rpQueryHeader(dp.col.c) : "";
    if (dh) {
      const now = new Date();
      const today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
      const from30 = new Date(today.getFullYear(), today.getMonth(), today.getDate() - 30);
      const fresh = new Date(today.getFullYear(), today.getMonth(), today.getDate() - RP_SCOPE_FRESH_DAYS);
      if (dp.max >= fresh && dp.min < from30) {
        add({ kind: "date", q: `${dh}:>>=${rpIsoDay(from30)} && ${dh}:<<=${rpIsoDay(today)}`, label: t("rpScopeLast30"), why: t("rpScopeWhyLast30", { col: dh }) });
      }
      if (dp.months.size >= 2) {
        const lastKey = Array.from(dp.months.keys()).sort().pop();
        const [y, m] = lastKey.split("-").map(Number);
        add({ kind: "date", q: `${dh}:>>=${rpIsoDay(new Date(y, m - 1, 1))} && ${dh}:<<=${rpIsoDay(new Date(y, m, 0))}`, label: rpMonthLabel(lastKey), why: t("rpScopeWhyMonth", { col: dh }) });
      }
    }
  }
  rpSuggestCache = { key, list };
  return list;
}

// Zmiana zakresu. Zapytanie bez wyników NIE zmienia raportu (pusta kartka nic nie mówi) —
// zostaje poprzedni zakres i komunikat przy polu.
function rpSetScope(scope) {
  let next = scope && scope.kind ? { kind: scope.kind, q: String(scope.q || "").trim() } : { kind: "view", q: "" };
  if (next.kind === "query" && !next.q) next = { kind: "view", q: "" };
  if (next.kind === "query") {
    if (/(^|[\s&|{!])@\S/.test(next.q)) {
      rpScopeMsg = t("rpScopeSmartOnly");
      rpRenderScopeBar();
      return false;
    }
    if (!rpQueryRows(next.q).length) {
      rpScopeMsg = t("rpScopeNoRows");
      rpRenderScopeBar();
      return false;
    }
  }
  // Zapamiętujemy ZAMIAR („jak w tabeli”), nawet gdy tabela chwilowo jest pusta i liczymy całość.
  rpScopeBySheet.set(rpSheetId(), next);
  rpScopeMsg = "";
  if (next.kind === "view" && !(Array.isArray(viewRows) && viewRows.length)) {
    next = { kind: "all", q: "" };
    rpScopeMsg = t("rpScopeViewEmpty");
  }
  rpScope = next;
  rpData = rpCollectAll();
  rpRender();
  return true;
}

function rpApplyScopeToTable() {
  if (rpScope.kind === "view" || typeof applyViewState !== "function") return;
  const res = applyViewState(rpScopeViewState());
  if (!res || !res.ok) {
    toast((res && res.problems ? res.problems.join(" ") : "") || t("rpScopeNoRows"), "warning");
    return;
  }
  rpScopeCustomOpen = false;
  rpSetScope({ kind: "view" });
  toast(t("rpScopeApplied", { rows: rpNum(res.shown, 0) }), "success");
}

function rpScopeChip({ id, label, n, why, active }) {
  const btn = rpEl("button", "rp-scope-chip");
  btn.type = "button";
  btn.dataset.scope = id;
  btn.setAttribute("aria-pressed", active ? "true" : "false");
  btn.append(rpEl("span", "rp-scope-chip-text", label));
  if (Number.isFinite(n)) btn.append(rpEl("span", "rp-scope-chip-n", rpNum(n, 0)));
  if (why) btn.setAttribute("data-hint", why);
  return btn;
}

function rpRenderScopeBar() {
  if (!rpScopeBarEl || !rpScopeChipsEl) return;
  const show = rpMode !== "table";
  rpScopeBarEl.classList.toggle("hidden", !show);
  if (!show) return;
  rpSetText("rpScopeBarLabel", t("rpScopeBarLabel"));
  const total = Array.isArray(baseRows) ? baseRows.length : 0;
  const tableFiltering = !!(typeof lastAppliedFilters !== "undefined" && lastAppliedFilters && lastAppliedFilters.filtering);
  const suggestions = rpScopeSuggestions();
  const isSuggestion = rpScope.kind === "query" && suggestions.some((s) => s.q === rpScope.q);
  const chips = [rpScopeChip({ id: "view", label: t("rpScopeView"), n: viewRows.length, why: t("rpScopeHintView"), active: rpScope.kind === "view" })];
  // Bez filtra w tabeli „cały arkusz” = „jak w tabeli” — drugi chip byłby tym samym.
  if (tableFiltering || rpScope.kind === "all") {
    chips.push(rpScopeChip({ id: "all", label: t("rpScopeWhole"), n: total, why: t("rpScopeHintWhole"), active: rpScope.kind === "all" }));
  }
  if (suggestions.length) {
    chips.push(rpEl("span", "rp-scope-sep", t("rpScopeSuggestSep")));
    suggestions.forEach((s) => chips.push(rpScopeChip({ id: `q:${s.q}`, label: s.label, n: s.n, why: `${s.why} ${t("rpScopeHintQuery", { q: s.q })}`, active: rpScope.kind === "query" && rpScope.q === s.q })));
  }
  const customActive = rpScope.kind === "query" && !isSuggestion;
  chips.push(rpScopeChip({ id: "custom", label: t("rpScopeCustom"), why: t("rpScopeHintCustom"), active: customActive || (rpScopeCustomOpen && rpScope.kind !== "query") }));
  rpScopeChipsEl.replaceChildren(...chips);
  // Wybrany chip zawsze w polu widzenia (na telefonie pasek przewija się w bok).
  const on = rpScopeChipsEl.querySelector('[aria-pressed="true"]');
  if (on && rpScopeChipsEl.scrollWidth > rpScopeChipsEl.clientWidth + 1) {
    const box = rpScopeChipsEl.getBoundingClientRect();
    const r = on.getBoundingClientRect();
    if (r.left < box.left) rpScopeChipsEl.scrollLeft -= box.left - r.left + 12;
    else if (r.right > box.right) rpScopeChipsEl.scrollLeft += r.right - box.right + 12;
  }
  rpScopeFade();

  const row2 = rpScope.kind === "query" || rpScopeCustomOpen || !!rpScopeMsg || rpScopeDiffers();
  rpScopeRow2El.classList.toggle("hidden", !row2);
  if (rpScopeQueryEl) {
    rpScopeQueryEl.placeholder = t("rpScopeQueryPh");
    rpScopeQueryEl.setAttribute("aria-label", t("rpScopeQueryAria"));
    rpScopeQueryEl.classList.toggle("hidden", !(rpScope.kind === "query" || rpScopeCustomOpen));
    if (document.activeElement !== rpScopeQueryEl) rpScopeQueryEl.value = rpScope.kind === "query" ? rpScope.q : "";
  }
  if (rpScopeStatusEl) {
    rpScopeStatusEl.classList.toggle("is-warn", !!rpScopeMsg);
    rpScopeStatusEl.textContent = rpScopeMsg
      || (rpScopeDiffers() ? t("rpScopeDiffers", { rows: rpNum(rpData ? rpData.rows : 0, 0), total: rpNum(total, 0) }) : "");
  }
  if (rpScopeApplyEl) {
    rpScopeApplyEl.textContent = t("rpScopeApply");
    rpScopeApplyEl.setAttribute("data-hint", t("rpScopeApplyHint"));
    rpScopeApplyEl.classList.toggle("hidden", !rpScopeDiffers());
  }
}

// Wygaszenie krawędzi tylko tam, gdzie jest co przewinąć (lewa / prawa osobno).
function rpScopeFade() {
  if (!rpScopeChipsEl) return;
  const el = rpScopeChipsEl;
  const over = el.scrollWidth > el.clientWidth + 1;
  el.classList.toggle("has-overflow", over);
  el.classList.toggle("fade-left", over && el.scrollLeft > 1);
  el.classList.toggle("fade-right", over && el.scrollLeft + el.clientWidth < el.scrollWidth - 1);
}

function rpSetText(id, text) {
  const el = document.getElementById(id);
  if (el && el.textContent !== text) el.textContent = text;
}

if (rpScopeChipsEl) {
  rpScopeChipsEl.addEventListener("scroll", rpScopeFade, { passive: true });
  window.addEventListener("resize", () => { if (rpIsOpen) rpScopeFade(); });
  rpScopeChipsEl.addEventListener("click", (e) => {
    const btn = e.target.closest("[data-scope]");
    if (!btn) return;
    const id = btn.dataset.scope;
    if (id === "custom") {
      rpScopeCustomOpen = true;
      rpScopeMsg = "";
      rpRenderScopeBar();
      rpScopeQueryEl?.focus();
      return;
    }
    rpScopeCustomOpen = false;
    if (id === "view" || id === "all") rpSetScope({ kind: id });
    else if (id.startsWith("q:")) rpSetScope({ kind: "query", q: id.slice(2) });
  });
}
if (rpScopeQueryEl) {
  const commit = () => {
    clearTimeout(rpScopeTimer);
    const q = rpScopeQueryEl.value.trim();
    if (!q) {
      rpScopeMsg = "";
      if (rpScope.kind === "query") rpSetScope({ kind: "view" });
      else rpRenderScopeBar();
      return;
    }
    if (rpScope.kind === "query" && rpScope.q === q && !rpScopeMsg) return;
    rpSetScope({ kind: "query", q });
  };
  rpScopeQueryEl.addEventListener("input", () => {
    clearTimeout(rpScopeTimer);
    rpScopeTimer = setTimeout(commit, 400);
  });
  rpScopeQueryEl.addEventListener("keydown", (e) => {
    if (e.key === "Enter") { e.preventDefault(); commit(); }
  });
}
rpScopeApplyEl?.addEventListener("click", rpApplyScopeToTable);

// Agregacje (silnik apki) czytają wiersze z globalnego viewRows — na czas liczenia podstawiamy
// wiersze zakresu raportu i zaraz oddajemy (wszystko synchronicznie, tabela tego nie widzi).
function rpWithScopeRows(rows, fn) {
  if (rows === viewRows) return fn();
  const saved = viewRows;
  viewRows = rows;
  try {
    return fn();
  } finally {
    viewRows = saved;
  }
}

function rpCollect() {
  const rows = rpScopeRows();
  const total = Array.isArray(baseRows) ? baseRows.length : rows.length;
  const view = rpScopeViewState();
  const filtering = rpScope.kind === "view"
    ? !!(typeof lastAppliedFilters !== "undefined" && lastAppliedFilters && lastAppliedFilters.filtering)
    : !!(view && view.filter);
  const cols = rpProfileColumns(rows);
  const used = cols.filter((c) => c.nonEmpty > 0);
  const emptyCols = cols.filter((c) => c.nonEmpty === 0);
  const filled = used.reduce((s, c) => s + c.nonEmpty, 0);
  const cats = rpCategories(cols, rows.length);
  const numericCols = rpNumericCols(cols);
  return {
    roles: rpDateRoles(cols, rows),
    rowsList: rows,
    rows: rows.length,
    total,
    filtering,
    scopeKind: rpScope.kind,
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
  d.state = rpStateModel(d);
  d.compare = rpCompareModel(d);
  rpWithScopeRows(d.rowsList, () => {
    d.aggs = rpAutoAggregations(d);
    d.aggPanel = rpPanelAggregation();
  });
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
  const out = rpAngleFindings(d);
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

// ── Kąty raportu: „Stan teraz” i „Porównanie grup” ─────────────────────────
// Kąt = pytanie, na które raport ma odpowiedzieć. Preset (krótki/normalny/szczegółowy)
// mówi „ile”, kąt mówi „o czym”. Kąt dokłada swoją sekcję i swoje wnioski na początek listy.
// Jak wszędzie w raporcie: zgadujemy z danych, ale każde zgadnięcie jest napisane na kartce
// („otwarte = W toku, Nowe”), a reguła bez podstaw milczy.

const RP_ANGLES = ["overview", "state", "compare"];
const RP_ANGLE_SECTION = { state: "state", compare: "compare" };
const RP_DAY_MS = 86400000;
const RP_STATE_MAX_ROWS = 5;
const RP_COMPARE_MAX_GROUPS = 6;
const RP_COMPARE_MIN_ROWS = 3;     // grupa mniejsza nie wchodzi do wniosków (2 wiersze to nie trend)

// Stan z NAZWY wartości. „niezakończone” to otwarte, choć zawiera „zakończ” — negacja pierwsza.
const RP_NEG_CLOSED_RE = /\bnie\s*-?\s*(zako|zamk|wykon|zreal|zrob|gotow|rozlicz|zap[łl]ac|odebr|oddan)|\bnot\s+(done|closed|finished|completed|paid)|\bun(paid|resolved|finished)/i;
const RP_CLOSED_RE = /zako[nń]cz|zamkni|zamkn|gotow|wykonan|zrealizow|zrobion|odebran|rozliczon|zap[łl]acon|oddan|anulow|odrzuc|wycofan|\bdone\b|closed|finished|complete|cancel|reject|resolved|\bpaid\b/i;
const RP_OPEN_RE = /w toku|otwart|planow|\bnow[eay]\b|oczekuj|wstrzym|realizac|w trakcie|do zrobienia|zaplanow|rozpocz|aktywn|w przygot|w drodze|\bopen\b|progress|pending|\bnew\b|to ?do|planned|active|on hold|waiting|draft/i;
// Role kolumn z datami — z nazwy; bez podpowiedzi w nazwie decydują dane (patrz rpDateRoles).
const RP_DUE_RE = /termin|deadline|\bdue\b|do kiedy|planowan|wymagan/i;
const RP_END_RE = /koniec|zako[nń]cz|zamkn|wykonan|zwrot|oddan|\bdo\b|\bend\b|closed|finish|complet|resolved|\bto\b/i;
const RP_START_RE = /start|rozpocz|pocz[aą]t|zg[łl]osz|utworz|przyj[eę]|otwar|wp[łl]yn|\bod\b|created|opened|received|begin|\bfrom\b/i;
const RP_LABEL_RE = /nazw|klient|tytu|temat|teren|obiekt|projekt|osoba|kontrahent|firma|zadanie|sprawa|name|title|client|customer|subject|task|project/i;

let rpCompareBy = "";                    // wybór w panelu (na tę sesję); "" = automatycznie
const rpCompareBySheet = new Map();

function rpStateKind(label) {
  const v = String(label || "");
  if (RP_NEG_CLOSED_RE.test(v)) return "open";
  if (RP_CLOSED_RE.test(v)) return "closed";
  if (RP_OPEN_RE.test(v)) return "open";
  return "other";
}

function rpDaysLabel(n, digits = 0) {
  return t(n === 1 ? "rpDay1" : "rpDays", { n: rpNum(n, digits) });
}

function rpToday() {
  const now = new Date();
  return new Date(now.getFullYear(), now.getMonth(), now.getDate());
}

function rpDaysBetween(a, b) {
  return Math.round((b - a) / RP_DAY_MS);
}

// Data z komórki tylko wtedy, gdy UŻYTKOWNIK widzi datę (liczba 45123 to nie data — patrz rpClassifyCell).
function rpCellDate(row, c) {
  if (c == null || c < 0) return null;
  const shown = String(getDisplayValue(row, c) ?? "").trim();
  if (!shown) return null;
  const value = row.values ? row.values[c] : shown;
  if (rpClassifyCell(value, shown) !== "date") return null;
  const d = parseDateFlexible(value);
  return d instanceof Date && !Number.isNaN(d.getTime()) ? new Date(d.getFullYear(), d.getMonth(), d.getDate()) : null;
}

// Liczba z komórki (albo suma bloków rodziny Kw1…Kw4).
function rpCellNum(row, col) {
  const one = (c) => {
    const v = row.values ? row.values[c] : undefined;
    return typeof v === "number" && Number.isFinite(v) ? v : null;
  };
  if (col.isFamily) {
    let sum = 0;
    let has = false;
    col.members.forEach((m) => { const v = one(m.c); if (v != null) { sum += v; has = true; } });
    return has ? sum : null;
  }
  return one(col.c);
}

// Role kolumn z datami: start (od kiedy), koniec (kiedy zamknięte), termin (do kiedy ma być).
// Najpierw nazwy; dwie daty bez podpowiedzi w nazwie → para start→koniec, jeśli w danych
// druga prawie zawsze jest ≥ pierwszej.
// Daty „z życia”: mediana roku 1971…dziś+5. Kwoty albo numery wyświetlane jako daty (1900–1926)
// nie mogą udawać „od kiedy czeka” — wychodziło „czeka 44 296 dni”.
function rpPlausibleDates(col) {
  const years = col.dates.map((d) => d.getFullYear()).sort((a, b) => a - b);
  const mid = years[Math.floor(years.length / 2)];
  return mid >= 1971 && mid <= new Date().getFullYear() + 5;
}

function rpDateRoles(cols, rows) {
  const dates = cols.filter((c) => c.kind === "date" && c.dates.length >= 2 && rpPlausibleDates(c));
  const due = dates.find((c) => RP_DUE_RE.test(c.name)) || null;
  const rest = dates.filter((c) => c !== due);
  let end = rest.find((c) => RP_END_RE.test(c.name) && !RP_START_RE.test(c.name)) || null;
  let start = rest.find((c) => c !== end && RP_START_RE.test(c.name))
    || rest.find((c) => c !== end && !RP_END_RE.test(c.name))
    || null;
  if (start && !end) {
    const others = rest.filter((c) => c !== start);
    if (others.length === 1 && !RP_START_RE.test(others[0].name)) {
      let both = 0;
      let ordered = 0;
      rows.forEach((row) => {
        const a = rpCellDate(row, start.c);
        const b = rpCellDate(row, others[0].c);
        if (a && b) { both += 1; if (b >= a) ordered += 1; }
      });
      if (both >= 3 && ordered / both >= 0.9) end = others[0];
    }
  }
  if (start && end && start.c > end.c && !RP_START_RE.test(start.name)) [start, end] = [end, start];
  return { start, end, due };
}

// Kolumna, po której człowiek rozpozna wiersz („Klient”, „Nazwa”, „Teren”…).
function rpLabelCol(cols, exclude = []) {
  const skip = new Set(exclude.filter(Boolean).map((c) => c.c));
  const text = cols.filter((c) => c.kind === "text" && c.fill >= 0.6 && !skip.has(c.c));
  return text.find((c) => RP_LABEL_RE.test(c.name) && c.unique >= c.nonEmpty * 0.3)
    || text.filter((c) => c.unique >= c.nonEmpty * 0.5).sort((a, b) => b.unique - a.unique)[0]
    || cols.find((c) => RP_ID_RE.test(c.name))
    || null;
}

function rpRowLabel(row, col) {
  if (!col) return t("rpRowN", { n: (row.rowIndex0 ?? 0) + 1 });
  return afShortSafe(String(getDisplayValue(row, col.c) ?? "").trim() || "—", 36);
}

function rpStatusCol(cols, rows) {
  const cats = rpCategories(cols, rows.length);
  for (const cat of cats) {
    const groups = cat.entries.map(([label, n]) => ({ label, n, kind: rpStateKind(label) }));
    const hasOpen = groups.some((g) => g.kind === "open");
    const hasClosed = groups.some((g) => g.kind === "closed");
    if ((hasOpen && hasClosed) || (RP_CAT_STRONG_RE.test(cat.col.name) && (hasOpen || hasClosed))) {
      return { col: cat.col, groups };
    }
  }
  return null;
}

function rpStateModel(d) {
  const rows = d.rowsList;
  if (rows.length < 3) return null;
  const roles = d.roles;
  const status = rpStatusCol(d.cols, rows);
  // Arkusz z cyklami (Tryby auto): stan wiersza = stan jego cykli (otwarty cykl = zaczęty, nieskończony).
  const smart = !status && typeof getSmartModel === "function" ? getSmartModel() : null;
  const records = smart && smart.hasStateColumns && typeof buildSmartRecords === "function" ? smart : null;
  const mode = status ? "status" : records ? "records" : roles.start && roles.end ? "dates" : roles.start ? "activity" : null;
  if (!mode) return null;
  const kindByKey = status ? new Map(status.groups.map((g) => [rpValueKey(g.label), g.kind])) : null;
  const today = rpToday();
  const labelCol = rpLabelCol(d.cols, [status && status.col, roles.start, roles.end, roles.due]);
  const counts = { open: 0, closed: 0, other: 0 };
  const open = [];
  const overdue = [];
  let lastDate = null;
  let last30 = 0;
  rows.forEach((row) => {
    let kind = "other";
    if (mode === "status") {
      const shown = String(getDisplayValue(row, status.col.c) ?? "").trim();
      kind = shown ? kindByKey.get(rpValueKey(shown)) || "other" : "other";
    } else if (mode === "dates") {
      const s = rpCellDate(row, roles.start.c);
      const e = rpCellDate(row, roles.end.c);
      kind = e ? "closed" : s ? "open" : "other";
    }
    let start = roles.start ? rpCellDate(row, roles.start.c) : null;
    let who = "";
    const seen = [start, roles.end ? rpCellDate(row, roles.end.c) : null];
    if (mode === "records") {
      const recs = buildSmartRecords(row, records).filter((r) => r.filled);
      const openRec = recs.filter((r) => r.state === "open").pop();
      kind = openRec ? "open" : recs.some((r) => r.state === "closed") ? "closed" : "other";
      start = openRec ? rpCellDate(row, openRec.startCol) : null;
      // „Kto” z otwartego cyklu (kolumna osoby w bloku) — w arkuszu obiegu to najważniejsza informacja.
      if (openRec && records.entityIdx >= 0) who = String(getDisplayValue(row, openRec.blockStart + records.entityIdx) ?? "").trim();
      recs.forEach((r) => seen.push(rpCellDate(row, r.startCol), rpCellDate(row, r.endCol)));
    }
    counts[kind] += 1;
    seen.forEach((dt) => {
      if (dt && dt <= today && (!lastDate || dt > lastDate)) lastDate = dt;
    });
    const newest = seen.filter(Boolean).filter((dt) => dt <= today).sort((a, b) => b - a)[0];
    if (newest && newest <= today && rpDaysBetween(newest, today) <= 30) last30 += 1;
    if (kind === "open") {
      open.push({ row, start, who, age: start && start <= today ? rpDaysBetween(start, today) : null });
      const due = roles.due ? rpCellDate(row, roles.due.c) : null;
      if (due && due < today) overdue.push({ row, due, late: rpDaysBetween(due, today) });
    }
  });
  const ages = open.map((o) => o.age).filter((a) => a != null).sort((a, b) => a - b);
  const buckets = ages.length ? [
    [t("rpAgeB1"), ages.filter((a) => a <= 7).length],
    [t("rpAgeB2"), ages.filter((a) => a > 7 && a <= 30).length],
    [t("rpAgeB3"), ages.filter((a) => a > 30 && a <= 90).length],
    [t("rpAgeB4"), ages.filter((a) => a > 90).length],
  ] : null;
  const recStart = records && records.blocks[0] ? records.blocks[0].startIndex + records.startIdx : -1;
  const recEnd = records && records.blocks[0] && records.endIdx >= 0 ? records.blocks[0].startIndex + records.endIdx : -1;
  return {
    mode,
    status,
    // W trybie cykli „start” to kolumna startu cyklu (nazwa z pierwszego bloku, do nagłówków na kartce).
    roles: mode === "records" ? { ...roles, start: recStart >= 0 ? { c: recStart, name: rpColName(recStart) } : roles.start } : roles,
    recEndName: recEnd >= 0 ? rpColName(recEnd) : "",
    whoName: records && records.entityIdx >= 0 && records.blocks[0] ? rpColName(records.blocks[0].startIndex + records.entityIdx) : "",
    labelCol,
    counts,
    total: rows.length,
    openList: open.filter((o) => o.age != null).sort((a, b) => b.age - a.age).slice(0, RP_STATE_MAX_ROWS),
    openCount: open.length,
    recent: mode === "activity"
      ? rows.map((row) => ({ row, date: rpCellDate(row, roles.start.c) })).filter((o) => o.date && o.date <= today).sort((a, b) => b.date - a.date).slice(0, RP_STATE_MAX_ROWS)
      : [],
    ages,
    medianAge: ages.length ? Math.round(rpMedian(ages)) : null,
    buckets,
    overdue: overdue.sort((a, b) => b.late - a.late),
    lastDate,
    lastAgo: lastDate ? rpDaysBetween(lastDate, today) : null,
    last30,
    openLabels: status ? status.groups.filter((g) => g.kind === "open").map((g) => g.label) : [],
    closedLabels: status ? status.groups.filter((g) => g.kind === "closed").map((g) => g.label) : [],
  };
}

// Opis „co uznaliśmy za otwarte” — na kartce i w panelu, żeby zgadnięcie dało się sprawdzić.
function rpStateRuleText(s) {
  if (s.mode === "status") {
    const list = (arr) => arr.slice(0, 4).map((x) => `„${x}”`).join(", ") + (arr.length > 4 ? "…" : "");
    const parts = [];
    if (s.openLabels.length) parts.push(t("rpStateRuleOpen", { list: list(s.openLabels) }));
    if (s.closedLabels.length) parts.push(t("rpStateRuleClosed", { list: list(s.closedLabels) }));
    return t("rpStateRuleStatus", { col: s.status.col.name, parts: parts.join(" · ") });
  }
  if (s.mode === "dates") return t("rpStateRuleDates", { start: s.roles.start.name, end: s.roles.end.name });
  if (s.mode === "records") return t("rpStateRuleRecords", { start: s.roles.start ? s.roles.start.name : "—", end: s.recEndName || "—" });
  return t("rpStateRuleActivity", { col: s.roles.start.name });
}

function rpStateFindings(s) {
  const out = [];
  if (s.mode !== "activity") {
    let text = t("rpFindStateOpen", { n: rpNum(s.openCount, 0), all: rpNum(s.total, 0), pct: rpPct(s.openCount, s.total) });
    const oldest = s.openList[0];
    if (oldest) {
      const who = oldest.who ? ` (${oldest.who})` : "";
      text += ` ${t("rpFindStateOldest", { days: rpDaysLabel(oldest.age), label: rpRowLabel(oldest.row, s.labelCol), who })}`;
    }
    out.push({ text });
    if (s.ages.length >= 5 && s.medianAge > 0) out.push({ text: t("rpFindStateMedian", { days: rpDaysLabel(s.medianAge) }) });
  }
  if (s.overdue.length) {
    const top = s.overdue[0];
    out.push({ text: t("rpFindStateOverdue", { n: rpNum(s.overdue.length, 0), col: s.roles.due.name, days: rpDaysLabel(top.late), label: rpRowLabel(top.row, s.labelCol) }), tone: "warn" });
  }
  if (s.lastDate) {
    if (s.lastAgo > 60 && (s.openCount > 0 || s.mode === "activity")) {
      out.push({ text: t("rpFindStateStale", { date: rpDate(s.lastDate), days: rpDaysLabel(s.lastAgo) }), tone: "warn" });
    } else if (s.lastAgo <= 60) {
      out.push({ text: t("rpFindStateFresh", { date: rpDate(s.lastDate), n: rpNum(s.last30, 0) }) });
    }
  }
  return out;
}

// ── Porównanie grup ──

function rpCompareOptions(d) {
  const opts = [];
  // Zakres vs reszta: wystarczy, że po obu stronach coś jest (mały zakres też da się porównać z resztą).
  if (d.filtering && d.rows > 0 && d.rows < d.total) opts.push({ id: "scope", label: t("rpCompareScopeOpt") });
  // Grupy w obrębie zakresu: dopiero od kilku wierszy (inaczej każda „grupa” to 1–2 wiersze).
  if (d.rows < RP_COMPARE_MIN_ROWS * 2) return opts;
  rpCategories(d.cols, d.rows).forEach((cat) => {
    const h = String(currentHeaders[cat.col.c] ?? "");
    if (h) opts.push({ id: `col:${h}`, label: t("rpCompareColOpt", { col: cat.col.name }), col: cat.col });
  });
  return opts;
}

function rpCompareGroupStats(rows, d, ctx) {
  const n = rows.length;
  const stat = { n };
  stat.measures = ctx.measures.map((m) => {
    const vals = rows.map((row) => rpCellNum(row, m.col)).filter((v) => v != null).sort((a, b) => a - b);
    return vals.length ? { n: vals.length, avg: vals.reduce((a, b) => a + b, 0) / vals.length, median: rpMedian(vals) } : null;
  });
  if (ctx.duration) {
    const days = [];
    rows.forEach((row) => {
      const a = rpCellDate(row, ctx.duration.start.c);
      const b = rpCellDate(row, ctx.duration.end.c);
      if (a && b && b >= a) days.push(rpDaysBetween(a, b));
    });
    stat.duration = days.length ? { n: days.length, avg: days.reduce((x, y) => x + y, 0) / days.length } : null;
  }
  if (ctx.state) {
    let open = 0;
    rows.forEach((row) => {
      const shown = String(getDisplayValue(row, ctx.state.col.c) ?? "").trim();
      if (shown && ctx.state.kindByKey.get(rpValueKey(shown)) === "open") open += 1;
    });
    stat.openPct = n ? rpPct(open, n) : 0;
  }
  if (ctx.usedCols.length && n) {
    let filled = 0;
    rows.forEach((row) => ctx.usedCols.forEach((c) => { if (String(getDisplayValue(row, c.c) ?? "").trim()) filled += 1; }));
    stat.complete = rpPct(filled, ctx.usedCols.length * n);
  }
  if (ctx.second) {
    const counts = new Map();
    let nonEmpty = 0;
    rows.forEach((row) => {
      const shown = String(getDisplayValue(row, ctx.second.c) ?? "").trim();
      if (!shown) return;
      nonEmpty += 1;
      const key = rpValueKey(shown);
      const g = counts.get(key) || { n: 0, label: shown };
      g.n += 1;
      counts.set(key, g);
    });
    stat.second = { nonEmpty, counts };
  }
  return stat;
}

function rpCompareModel(d) {
  const opts = rpCompareOptions(d);
  if (!opts.length) return null;
  const want = rpCompareBy || rpCompareBySheet.get(rpSheetId()) || "";
  const opt = opts.find((o) => o.id === want) || opts[0];
  let groups;
  let groupCol = null;
  let skipped = 0;
  if (opt.id === "scope") {
    const inScope = new Set(d.rowsList.map((r) => r.rowIndex0));
    const rest = baseRows.filter((r) => !inScope.has(r.rowIndex0));
    groups = [{ label: t("rpCompareScope"), rows: d.rowsList }, { label: t("rpCompareRest"), rows: rest }];
  } else {
    groupCol = opt.col;
    const map = new Map();
    d.rowsList.forEach((row) => {
      const shown = String(getDisplayValue(row, groupCol.c) ?? "").trim();
      if (!shown) return;
      const key = rpValueKey(shown);
      let g = map.get(key);
      if (!g) { g = { spell: new Map(), rows: [] }; map.set(key, g); }
      g.spell.set(shown, (g.spell.get(shown) || 0) + 1);
      g.rows.push(row);
    });
    const all = Array.from(map.values())
      .map((g) => ({ label: Array.from(g.spell.entries()).sort((a, b) => b[1] - a[1])[0][0], rows: g.rows }))
      .sort((a, b) => b.rows.length - a.rows.length);
    groups = all.slice(0, RP_COMPARE_MAX_GROUPS);
    skipped = all.length - groups.length;
  }
  if (groups.length < 2 || groups.filter((g) => g.rows.length).length < 2) return null;
  // Miary: najwyżej 2 liczby + czas trwania + % otwartych + kompletność + najczęstsza wartość innej kategorii.
  const measures = d.numericCols.filter((s) => !groupCol || s.col.c !== groupCol.c).slice(0, 2).map((s) => ({ col: s.col, name: s.col.name }));
  const roles = d.roles;
  const duration = roles.start && roles.end ? { start: roles.start, end: roles.end, name: `${roles.start.name} → ${roles.end.name}` } : null;
  const st = d.state && d.state.mode === "status" && (!groupCol || d.state.status.col.c !== groupCol.c)
    ? { col: d.state.status.col, kindByKey: new Map(d.state.status.groups.map((g) => [rpValueKey(g.label), g.kind])) }
    : null;
  const second = rpCategories(d.cols, d.rows).map((c) => c.col).find((c) => (!groupCol || c.c !== groupCol.c) && (!st || c.c !== st.col.c)) || null;
  const ctx = { measures, duration, state: st, second, usedCols: d.used };
  const total = groups.reduce((s2, g) => s2 + g.rows.length, 0);
  const stats = groups.map((g) => ({ label: g.label, ...rpCompareGroupStats(g.rows, d, ctx) }));
  return { mode: opt.id === "scope" ? "scope" : "column", optId: opt.id, options: opts, groupCol, groups: stats, total, ctx, skipped };
}

function rpCompareFindings(c) {
  const out = [];
  const solid = c.groups.filter((g) => g.n >= RP_COMPARE_MIN_ROWS);
  if (solid.length < 2) return out;
  const scope = c.mode === "scope";
  const ratio = (name, pick, fmt) => {
    const vals = solid.map((g) => ({ g, v: pick(g) })).filter((x) => Number.isFinite(x.v) && x.v > 0);
    if (vals.length < 2) return;
    const hi = vals.reduce((a, b) => (b.v > a.v ? b : a));
    const lo = vals.reduce((a, b) => (b.v < a.v ? b : a));
    if (hi === lo || hi.v < lo.v * 1.5) return;
    const x = rpNum(hi.v / lo.v, 1);
    if (scope) {
      const inScope = vals.find((v) => v.g === c.groups[0]);
      const rest = vals.find((v) => v.g === c.groups[1]);
      if (!inScope || !rest) return;
      out.push({ text: t(inScope === hi ? "rpFindCmpScopeMore" : "rpFindCmpScopeLess", { name, v: fmt(inScope.v), x, rest: fmt(rest.v) }) });
    } else {
      out.push({ text: t("rpFindCmpSpread", { name, hi: hi.g.label, hiV: fmt(hi.v), lo: lo.g.label, loV: fmt(lo.v), x }) });
    }
  };
  c.ctx.measures.forEach((m, i) => ratio(t("rpCmpAvgOf", { col: m.name }), (g) => g.measures[i] && g.measures[i].avg, (v) => rpNum(v)));
  if (c.ctx.duration) ratio(t("rpCmpDurationOf", { col: c.ctx.duration.name }), (g) => g.duration && g.duration.avg, (v) => rpDaysLabel(v, 1));
  if (c.ctx.state) {
    const hi = solid.reduce((a, b) => (b.openPct > a.openPct ? b : a));
    const lo = solid.reduce((a, b) => (b.openPct < a.openPct ? b : a));
    if (hi.openPct - lo.openPct >= 20) out.push({ text: t("rpFindCmpOpen", { hi: hi.label, hiP: hi.openPct, lo: lo.label, loP: lo.openPct }) });
  }
  // Skład wg innej kategorii: największa różnica udziału jednej wartości (grupy ≥ 5 wierszy).
  if (c.ctx.second) {
    const big = solid.filter((g) => g.second && g.second.nonEmpty >= 5);
    let best = null;
    const keys = new Set(big.flatMap((g) => Array.from(g.second.counts.keys())));
    keys.forEach((k) => {
      const shares = big.map((g) => ({ g, p: rpPct((g.second.counts.get(k) || { n: 0 }).n, g.second.nonEmpty), label: (g.second.counts.get(k) || {}).label }));
      if (shares.length < 2) return;
      const hi = shares.reduce((a, b) => (b.p > a.p ? b : a));
      const lo = shares.reduce((a, b) => (b.p < a.p ? b : a));
      if (hi.p - lo.p >= 15 && (!best || hi.p - lo.p > best.diff)) best = { diff: hi.p - lo.p, hi, lo, value: hi.label || shares.find((s) => s.label).label };
    });
    if (best) out.push({ text: t("rpFindCmpMix", { col: c.ctx.second.name, value: best.value, hi: best.hi.g.label, hiP: best.hi.p, lo: best.lo.g.label, loP: best.lo.p }) });
  }
  const comp = solid.filter((g) => Number.isFinite(g.complete));
  if (comp.length >= 2) {
    const hi = comp.reduce((a, b) => (b.complete > a.complete ? b : a));
    const lo = comp.reduce((a, b) => (b.complete < a.complete ? b : a));
    if (hi.complete - lo.complete >= 15) out.push({ text: t("rpFindCmpComplete", { lo: lo.label, loP: lo.complete, hi: hi.label, hiP: hi.complete }), tone: "warn" });
  }
  if (!out.length) out.push({ text: t("rpFindCmpNone") });
  return out;
}

// Wnioski kątów idą NA POCZĄTEK listy — to odpowiedź na pytanie, które zadał użytkownik.
function rpAngleFindings(d) {
  const active = rpActiveSections();
  const out = [];
  if (active.includes("state") && d.state) out.push(...rpStateFindings(d.state));
  if (active.includes("compare") && d.compare) out.push(...rpCompareFindings(d.compare));
  return out;
}

function rpSetAngle(angle) {
  if (!RP_ANGLES.includes(angle)) return;
  rpPrefs.angle = angle;
  rpSavePrefs();
  rpRender();
}

function rpSetCompareBy(id) {
  rpCompareBy = id || "";
  rpCompareBySheet.set(rpSheetId(), rpCompareBy);
  if (rpData) rpData.compare = rpCompareModel(rpData);
  rpRender();
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
  // Styl „Oszczędny” = czerń. Akcent ustawiany tu inline wygrywał z regułą CSS [data-style="ink"],
  // więc cz-b drukowało zielone nagłówki — kolor wybieramy więc już tutaj.
  const accent = rpPrefs.style === "ink" ? "#000000" : (RP_ACCENTS[rpPrefs.accent] || RP_ACCENTS.green);
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
  rpRenderScopeBar();
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
    // Zero = pusty pasek (minimalna kreska sugerowałaby, że coś tam jest).
    fill.style.width = n > 0 ? `${Math.max(2, Math.round((n / maxN) * 100))}%` : "0";
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
      // Długa liczba (np. „349 653 101,54”) nie może się łamać w środku — zamiast tego mniejsza czcionka.
      const len = String(tile.value ?? "").length;
      const long = tile.text ? "" : len > 15 ? " is-xlong" : len > 11 ? " is-long" : "";
      const box = rpEl("div", `rp-tile${tile.text ? " is-text" : ""}${long}`);
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
  state(d) {
    const s = d.state;
    // Klasa rp-aggs = łamanie MIĘDZY częściami (każda .rp-agg w całości), jak w zestawieniach.
    const sec = rpSection(t("rpStateTitle"), "rp-aggs rp-state");
    sec.appendChild(rpEl("p", "rp-note", rpStateRuleText(s)));
    if (s.mode !== "activity") {
      const box = rpEl("div", "rp-agg");
      const parts = [["open", s.counts.open], ["closed", s.counts.closed], ["other", s.counts.other]].filter(([, n]) => n > 0);
      const bar = rpEl("div", "rp-stack");
      bar.setAttribute("aria-hidden", "true");
      parts.forEach(([k, n]) => {
        const seg = rpEl("div", `rp-stack-seg is-${k}`);
        seg.style.width = `${(n / s.total) * 100}%`;
        bar.appendChild(seg);
      });
      const legend = rpEl("div", "rp-stack-legend");
      parts.forEach(([k, n]) => {
        const item = rpEl("span", `rp-stack-key is-${k}`);
        item.append(rpEl("i", "rp-stack-dot"), rpEl("span", "", `${t(`rpState_${k}`)}: ${rpNum(n, 0)} (${rpPct(n, s.total)}%)`));
        legend.appendChild(item);
      });
      box.append(bar, legend);
      sec.appendChild(box);
    }
    if (s.buckets && s.ages.length >= 3) {
      const box = rpEl("div", "rp-agg");
      box.appendChild(rpEl("h3", "rp-h3", t("rpStateAges", { col: s.roles.start.name })));
      rpBars(box, s.buckets, s.ages.length);
      sec.appendChild(box);
    }
    const label = s.labelCol ? s.labelCol.name : t("rpColRow");
    // Bez pojęcia „otwarte” (sama kolumna daty) pokazujemy przynajmniej najnowsze wpisy.
    if (s.mode === "activity" && s.recent.length) {
      const box = rpEl("div", "rp-agg");
      box.appendChild(rpEl("h3", "rp-h3", t("rpStateRecent", { col: s.roles.start.name })));
      box.appendChild(rpTable([label, s.roles.start.name], s.recent.map((o) => [rpRowLabel(o.row, s.labelCol), rpDate(o.date)])));
      sec.appendChild(box);
    }
    if (s.openList.length) {
      const box = rpEl("div", "rp-agg");
      box.appendChild(rpEl("h3", "rp-h3", t("rpStateOldest")));
      const withStatus = s.mode === "status";
      const withWho = !!s.whoName;
      box.appendChild(rpTable(
        [label, ...(withStatus ? [s.status.col.name] : []), ...(withWho ? [s.whoName] : []), s.roles.start.name, t("rpColWaiting")],
        s.openList.map((o) => [
          rpRowLabel(o.row, s.labelCol),
          ...(withStatus ? [String(getDisplayValue(o.row, s.status.col.c) ?? "")] : []),
          ...(withWho ? [afShortSafe(o.who || "—", 30)] : []),
          rpDate(o.start),
          rpDaysLabel(o.age, 0),
        ]),
        { numCols: [2 + (withStatus ? 1 : 0) + (withWho ? 1 : 0)] },
      ));
      sec.appendChild(box);
    }
    if (s.overdue.length) {
      const box = rpEl("div", "rp-agg");
      box.appendChild(rpEl("h3", "rp-h3", t("rpStateOverdue", { col: s.roles.due.name, n: rpNum(s.overdue.length, 0) })));
      box.appendChild(rpTable(
        [label, s.roles.due.name, t("rpColLate")],
        s.overdue.slice(0, RP_STATE_MAX_ROWS).map((o) => [rpRowLabel(o.row, s.labelCol), rpDate(o.due), rpDaysLabel(o.late, 0)]),
        { numCols: [2] },
      ));
      sec.appendChild(box);
    }
    return sec;
  },
  compare(d) {
    const c = d.compare;
    const sec = rpSection(c.mode === "scope" ? t("rpCompareTitleScope") : t("rpCompareTitleCol", { col: c.groupCol.name }), "rp-aggs rp-compare");
    const note = c.mode === "scope"
      ? t("rpCompareNoteScope", { view: d.viewText, rows: rpNum(d.rows, 0), rest: rpNum(c.groups[1].n, 0) })
      : t("rpCompareNoteCol", { n: c.groups.length }) + (c.skipped ? ` ${t("rpCompareSkipped", { n: c.skipped })}` : "");
    sec.appendChild(rpEl("p", "rp-note", note));
    const ctx = c.ctx;
    const cols = [
      { h: c.mode === "scope" ? t("rpColGroup") : c.groupCol.name, v: (g) => g.label },
      { h: t("rpColRows"), v: (g) => rpNum(g.n, 0), num: true },
      { h: t("rpColShare"), v: (g) => `${rpPct(g.n, c.total)}%`, num: true },
      ...ctx.measures.map((m, i) => ({ h: t("rpCmpAvgOf", { col: m.name }), v: (g) => (g.measures[i] ? rpNum(g.measures[i].avg) : "—"), num: true })),
      ...(ctx.duration ? [{ h: t("rpCmpDurationShort"), v: (g) => (g.duration ? rpDaysLabel(g.duration.avg, 1) : "—"), num: true }] : []),
      ...(ctx.state ? [{ h: t("rpCmpOpenShort"), v: (g) => `${g.openPct}%`, num: true }] : []),
      { h: t("rpColComplete"), v: (g) => (Number.isFinite(g.complete) ? `${g.complete}%` : "—"), num: true, optional: true },
      ...(ctx.second ? [{ h: t("rpCmpTopOf", { col: ctx.second.name }), v: (g) => {
        if (!g.second || !g.second.nonEmpty) return "—";
        const top = Array.from(g.second.counts.values()).sort((a, b) => b.n - a.n)[0];
        return `${afShortSafe(top.label, 22)} (${rpPct(top.n, g.second.nonEmpty)}%)`;
      } }] : []),
    ];
    // Za dużo kolumn na pionowym A4 → najpierw wypada „Kompletność”.
    const shown = cols.length > 8 ? cols.filter((col) => !col.optional) : cols;
    const box = rpEl("div", "rp-agg");
    box.appendChild(rpTable(
      shown.map((col) => col.h),
      c.groups.map((g) => shown.map((col) => col.v(g))),
      { numCols: shown.map((col, i) => (col.num ? i : -1)).filter((i) => i >= 0) },
    ));
    sec.appendChild(box);
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

// Sekcje wybrane presetem albo ręcznie (bez dokładki z kąta).
function rpBaseSections() {
  return rpPrefs.preset !== "custom" && RP_PRESETS[rpPrefs.preset]
    ? RP_PRESETS[rpPrefs.preset]
    : rpPrefs.sections;
}

// + sekcja kąta (Stan teraz / Porównanie) — zawsze, niezależnie od długości raportu.
function rpActiveSections() {
  const base = rpBaseSections();
  const extra = RP_ANGLE_SECTION[rpPrefs.angle];
  if (!extra || base.includes(extra)) return base;
  const set = new Set([...base, extra]);
  return RP_SECTIONS.map((sec) => sec.id).filter((id) => set.has(id));
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
    const angle = rpPrefs.angle !== "overview" ? `${t(`rpAngle_${rpPrefs.angle}`)} · ` : "";
    rpContentBtn.textContent = t("rpContentBtn", { preset: `${angle}${name}` });
  }
  rpRenderAnglePanel();
  const active = rpActiveSections();
  const forced = RP_ANGLE_SECTION[rpPrefs.angle];
  rpSectionListEl.replaceChildren();
  RP_SECTIONS.forEach((sec) => {
    const available = rpSectionAvailable(sec, rpData);
    const label = rpEl("label", `rp-sec${available ? "" : " is-unavailable"}`);
    const cb = rpEl("input");
    cb.type = "checkbox";
    cb.value = sec.id;
    cb.checked = available && active.includes(sec.id); // „brak w tym arkuszu” nie może wyglądać na zaznaczone
    cb.disabled = !available || sec.id === forced;
    cb.addEventListener("change", () => rpToggleSection(sec.id, cb.checked));
    const text = rpEl("span", "rp-sec-text", t(sec.label));
    label.append(cb, text);
    if (available && sec.id === forced) label.appendChild(rpEl("span", "rp-sec-why", t("rpSecFromAngle")));
    // Sekcja bez danych zostaje na liście (żeby było wiadomo, że istnieje), ale mówi czemu jej nie ma.
    if (!available) label.appendChild(rpEl("span", "rp-sec-why", t("rpSecUnavailable")));
    rpSectionListEl.appendChild(label);
  });
}

// Wybór kąta + (dla porównania) „co z czym”. Każdy kąt mówi, co wykrył albo czemu go nie ma.
function rpRenderAnglePanel() {
  const box = document.getElementById("rpAngles");
  if (!box || !rpData) return;
  const d = rpData;
  const head = rpEl("div", "rp-angle-row");
  head.appendChild(rpEl("span", "rp-angle-label", t("rpAngleLabel")));
  const chips = rpEl("div", "rp-presets");
  chips.setAttribute("role", "group");
  chips.setAttribute("aria-label", t("rpAngleLabel"));
  const avail = { overview: true, state: !!d.state, compare: !!d.compare };
  ["overview", "state", "compare"].forEach((a) => {
    const btn = rpEl("button", "rp-preset", t(`rpAngle_${a}`));
    btn.type = "button";
    btn.dataset.angle = a;
    btn.setAttribute("aria-pressed", rpPrefs.angle === a ? "true" : "false");
    btn.setAttribute("data-hint", t(`rpAngleHint_${a}`));
    if (!avail[a]) {
      btn.disabled = true;
      btn.setAttribute("data-hint", t(`rpAngleNo_${a}`));
    }
    btn.addEventListener("click", () => rpSetAngle(a));
    chips.appendChild(btn);
  });
  head.appendChild(chips);
  const parts = [head];
  let info = "";
  if (rpPrefs.angle === "state") info = d.state ? rpStateRuleText(d.state) : t("rpAngleNo_state");
  if (rpPrefs.angle === "compare") {
    if (!d.compare) info = t("rpAngleNo_compare");
    else {
      const row = rpEl("label", "rp-field rp-compare-pick");
      row.appendChild(rpEl("span", "", t("rpComparePick")));
      const sel = rpEl("select");
      sel.id = "rpCompareBy";
      d.compare.options.forEach((o) => {
        const opt = rpEl("option", "", o.label);
        opt.value = o.id;
        sel.appendChild(opt);
      });
      sel.value = d.compare.optId;
      sel.addEventListener("change", () => rpSetCompareBy(sel.value));
      row.appendChild(sel);
      parts.push(row);
    }
  }
  if (info) parts.push(rpEl("p", "rp-angle-info", info));
  box.replaceChildren(...parts);
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
  const next = new Set(rpBaseSections());
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
    rpFitNoteEl.textContent = pages > 1 ? t(rpPluralKey("rpFitPages", pages), { n: pages }) : t("rpFitOne");
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

// Nasłuchy przycisków paska — rejestrowane RAZ. Dawniej stały w środku rpSyncControls
// (wołanego przy każdym renderze), więc po N odświeżeniach klik przełączał panel N razy:
// „Zawartość” / „Wygląd” czasem nic nie robiły, a preset przeliczał raport N razy.
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
  if (!Array.isArray(currentHeaders) || !currentHeaders.length || !Array.isArray(baseRows) || !baseRows.length) {
    toast(t("noDataForExport"), "warning");
    return;
  }
  rpLoadPrefs();
  rpMode = "report";
  rpTableSpec = null;
  // Zakres pamiętany dla tego arkusza (na tę sesję). Zapytanie, które po zmianach w danych nic
  // nie łapie, i pusta tabela → raport liczy cały arkusz i mówi o tym przy pasku zakresu.
  rpScopeMsg = "";
  rpScopeCustomOpen = false;
  rpScope = rpScopeBySheet.get(rpSheetId()) || { kind: "view", q: "" };
  if (rpScope.kind === "query" && !rpQueryRows(rpScope.q).length) rpScope = { kind: "view", q: "" };
  if (rpScope.kind === "view" && !viewRows.length) {
    rpScope = { kind: "all", q: "" };
    rpScopeMsg = t("rpScopeViewEmpty");
  }
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
  setAngle: rpSetAngle,
  setCompareBy: rpSetCompareBy,
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
  scope: () => ({ ...rpScope }),
  setScope: rpSetScope,
  suggestions: () => rpScopeSuggestions().map((x) => ({ ...x })),
};
