// autofill.js — wypełnianie jak w Excelu: uchwyt w rogu zaznaczenia (serie), Ctrl+D / Ctrl+R.
//
// Wzorzec: polski Excel na Macu (AppleScript `autofill … type fill default` i `fill down`,
// 2026-10-17). Co Excel robi i co tu jest odtworzone:
//   • jedna liczba → kopia (z Ctrl/Alt przy puszczeniu: seria +1); dwie i więcej → trend
//     liniowy (najmniejsze kwadraty: 1, 3, 4 → 5,67; 7,17…),
//   • data → +1 dzień, godzina → +1 godzina; dwie daty: krok w miesiącach, gdy dzień
//     miesiąca ten sam albo oba to końce miesięcy (31.01, 28.02 → 31.03, 30.04),
//     inaczej krok w dniach (1.03, 8.03 → 15.03),
//   • tekst z liczbą NA KOŃCU („Pozycja 1”, „Faktura 001”, „A-07”, „x1”) albo na początku
//     przed spacją („1 szt”) → kolejne numery z zachowaniem zer wiodących; liczba w środku
//     („Poz. 9 a”, „1x”, „2025r”) = kopia; kwartały („Kw1”, „Kw. 1”, „kwartał 1”, „1 kw”)
//     wracają po 4 do 1,
//   • dni tygodnia i miesiące po polsku (pełne i skróty: pon → wt., śr.…; sty → lut),
//     wielkość liter jak w źródle (Pon → Wt., PONIEDZIAŁEK → WTOREK),
//   • kilka komórek źródła powtarza się w cyklu, a każdy ciąg tego samego rodzaju liczy
//     się osobno (1, 2, x → 3, 4, x, 5, 6…; x, Pozycja 1 → x, Pozycja 2…),
//   • formuły przesuwają odwołania względne, format liczbowy idzie razem z wartością.
// Świadome odstępstwo: kopiujemy FORMAT LICZBOWY, ale nie wygląd (tło, czcionka) — Excel
// kopiuje całe formatowanie; tu wygląd komórki docelowej zostaje.

// ── Silnik (czysta funkcja, bez DOM) ─────────────────────────────────────────────

const SWB_FILL_LISTS = {
  pl: [
    { id: "dayFull", items: ["poniedziałek", "wtorek", "środa", "czwartek", "piątek", "sobota", "niedziela"] },
    { id: "monthFull", items: ["styczeń", "luty", "marzec", "kwiecień", "maj", "czerwiec", "lipiec", "sierpień", "wrzesień", "październik", "listopad", "grudzień"] },
    { id: "dayShort", items: ["pon.", "wt.", "śr.", "czw.", "pt.", "sob.", "niedz."] },
    { id: "monthShort", items: ["sty", "lut", "mar", "kwi", "maj", "cze", "lip", "sie", "wrz", "paź", "lis", "gru"] },
  ],
  en: [
    { id: "dayFull", items: ["monday", "tuesday", "wednesday", "thursday", "friday", "saturday", "sunday"] },
    { id: "monthFull", items: ["january", "february", "march", "april", "may", "june", "july", "august", "september", "october", "november", "december"] },
    { id: "dayShort", items: ["mon", "tue", "wed", "thu", "fri", "sat", "sun"] },
    { id: "monthShort", items: ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"] },
  ],
};

function swbFillLower(s) { return String(s).toLocaleLowerCase("pl"); }
function swbFillUpper(s) { return String(s).toLocaleUpperCase("pl"); }

function swbFillCaseOf(s) {
  const up = swbFillUpper(s);
  const low = swbFillLower(s);
  if (s === up && s !== low) return "upper";
  if (s[0] && s[0] === swbFillUpper(s[0]) && s[0] !== swbFillLower(s[0])) return "cap";
  return "lower";
}

function swbFillApplyCase(s, kind) {
  if (kind === "upper") return swbFillUpper(s);
  if (kind === "cap") return swbFillUpper(s[0]) + s.slice(1);
  return s;
}

function swbFillListMatch(s, lang) {
  const key = swbFillLower(s.trim()).replace(/\.$/, "");
  if (!key) return null;
  const lists = SWB_FILL_LISTS[lang === "en" ? "en" : "pl"];
  for (const list of lists) {
    const idx = list.items.findIndex((it) => it.replace(/\.$/, "") === key);
    if (idx >= 0) return { list, idx };
  }
  return null;
}

// Format liczbowy z samą godziną (h:mm, [h]:mm:ss, hh:mm AM/PM) — bez dni i lat.
function swbFillIsTimeFmt(code) {
  if (typeof code !== "string" || !code || code === "General") return false;
  const bare = code.replace(/"[^"]*"/g, "").replace(/\[(?![hms]+\])[^\]]*\]/gi, "");
  return /[hs]/i.test(bare) && !/[dy]/i.test(bare);
}

function swbFillIsDateFmt(code) {
  return typeof swbIsDateFormat === "function" ? swbIsDateFormat(code) : false;
}

function swbFillSerial(d) {
  return typeof swbSerialFromDate === "function" ? swbSerialFromDate(d)
    : (Date.UTC(d.getFullYear(), d.getMonth(), d.getDate(), d.getHours(), d.getMinutes(), d.getSeconds()) - Date.UTC(1899, 11, 30)) / 86400000;
}

function swbFillDateFromSerial(serial) {
  const u = new Date(Math.round((serial - 25569) * 86400000));
  return new Date(u.getUTCFullYear(), u.getUTCMonth(), u.getUTCDate(), u.getUTCHours(), u.getUTCMinutes(), u.getUTCSeconds());
}

// Rodzaj komórki źródła. key = co musi się zgadzać, żeby sąsiednie komórki tworzyły
// jeden ciąg (np. ten sam przedrostek „Pozycja ”).
function swbFillClassify(src, lang) {
  const v = src.v;
  if (src.formula) return { kind: "copy", key: null };
  if (v == null || v === "" || typeof v === "boolean") return { kind: "copy", key: null };
  if (v instanceof Date) {
    const serial = swbFillSerial(v);
    if (!Number.isFinite(serial)) return { kind: "copy", key: null };
    return serial < 1 ? { kind: "time", key: "tD", serial, asDate: true } : { kind: "date", key: "dD", serial, asDate: true };
  }
  if (typeof v === "number") {
    if (!Number.isFinite(v)) return { kind: "copy", key: null };
    if (swbFillIsTimeFmt(src.fmt)) return { kind: "time", key: "tN", serial: v, asDate: false };
    if (swbFillIsDateFmt(src.fmt)) return { kind: "date", key: "dN", serial: v, asDate: false };
    return { kind: "number", key: "n", n: v };
  }
  const s = String(v);
  const hit = swbFillListMatch(s, lang);
  if (hit) return { kind: "list", key: `L${hit.list.id}`, list: hit.list, idx: hit.idx, caseKind: swbFillCaseOf(s.trim()) };
  const quarterRe = lang === "en" ? /^(?:kw\.?|kwartał|q|qtr\.?|quarter)\s*$/i : /^(?:kw\.?|kwartał)\s*$/i;
  const quarterSufRe = lang === "en" ? /^\s+(?:kw\.?|kwartał|q|qtr|quarter)$/i : /^\s+(?:kw\.?|kwartał)$/i;
  // Ponad 15 cyfr (IBAN, nr konta, PESEL-e sklejone z kodem) nie mieści się w liczbie JS:
  // seria robiła z „PL61109010140000071219812874” zapis „PL6.11e+25”. Takie wartości kopiujemy.
  if (/\d{16,}/.test(s)) return { kind: "copy", key: null };
  let m = /^([\s\S]*?)(\d+)$/.exec(s);
  if (m) {
    const n = Number(m[2]);
    const quarter = quarterRe.test(m[1]) && n >= 1 && n <= 4;
    return { kind: "textnum", key: `T${m[1]}\u0000${quarter ? "q" : ""}`, pre: m[1], suf: "", n, width: m[2].length > 1 && m[2][0] === "0" ? m[2].length : 0, quarter };
  }
  m = /^(\d+)(\s+\S[\s\S]*)$/.exec(s);
  if (m) {
    const n = Number(m[1]);
    const quarter = quarterSufRe.test(m[2]) && n >= 1 && n <= 4;
    return { kind: "textnum", key: `T\u0000${m[2]}${quarter ? "q" : ""}`, pre: "", suf: m[2], n, width: m[1].length > 1 && m[1][0] === "0" ? m[1].length : 0, quarter };
  }
  return { kind: "copy", key: null };
}

// Prosta najmniejszych kwadratów dla x = 0..n−1.
function swbFillLinFit(ys) {
  const n = ys.length;
  if (n === 1) return { a: ys[0], b: 0 };
  const mx = (n - 1) / 2;
  const my = ys.reduce((s, y) => s + y, 0) / n;
  let sxy = 0;
  let sxx = 0;
  ys.forEach((y, x) => { sxy += (x - mx) * (y - my); sxx += (x - mx) * (x - mx); });
  const b = sxx ? sxy / sxx : 0;
  return { a: my - b * mx, b };
}

function swbFillDaysInMonth(y, m0) { return new Date(y, m0 + 1, 0).getDate(); }

// Krok miesięczny między datami? → liczba miesięcy albo null.
function swbFillMonthStep(dates) {
  const ym = dates.map((d) => d.getFullYear() * 12 + d.getMonth());
  const step = ym[1] - ym[0];
  if (!step) return null;
  for (let i = 2; i < ym.length; i++) if (ym[i] - ym[i - 1] !== step) return null;
  const eom = dates.every((d) => d.getDate() === swbFillDaysInMonth(d.getFullYear(), d.getMonth()));
  const sameDay = dates.every((d) => d.getDate() === dates[0].getDate());
  if (!eom && !sameDay) return null;
  return { step, eom };
}

function swbFillAddMonths(base, months, eom) {
  const y = base.getFullYear();
  const m0 = base.getMonth() + months;
  const first = new Date(y, m0, 1, base.getHours(), base.getMinutes(), base.getSeconds());
  const dim = swbFillDaysInMonth(first.getFullYear(), first.getMonth());
  first.setDate(eom ? dim : Math.min(base.getDate(), dim));
  return first;
}

// Generator wartości dla jednego ciągu (run) źródła: x = pozycja w ciągu liczona od jego
// początku przez wszystkie powtórzenia (x < L to samo źródło).
function swbFillRunGenerator(run, single, invert) {
  const c0 = run.cells[0];
  const L = run.cells.length;
  const copy = () => null;
  if (c0.kind === "copy") return copy;
  if (c0.kind === "number") {
    // jedna komórka = kopia (Excel), Ctrl/Alt odwraca: seria +1
    if (single && !invert) return copy;
    if (!single && invert) return copy;
    const fit = L === 1 ? { a: c0.n, b: 1 } : swbFillLinFit(run.cells.map((c) => c.n));
    return (x) => ({ value: Number((fit.a + fit.b * x).toPrecision(15)), type: "number" });
  }
  if (invert) return copy; // dla dat, tekstu z numerem i list Ctrl/Alt = kopia
  if (c0.kind === "time" || c0.kind === "date") {
    const out = (serial) => (c0.asDate
      ? { value: swbFillDateFromSerial(serial), type: "date" }
      : { value: Number(serial.toPrecision(15)), type: "number" });
    if (L === 1) {
      const step = c0.kind === "time" ? 1 / 24 : 1;
      return (x) => out(c0.serial + step * x);
    }
    if (c0.kind === "date") {
      const dates = run.cells.map((c) => swbFillDateFromSerial(c.serial));
      const ms = swbFillMonthStep(dates);
      if (ms) return (x) => out(swbFillSerial(swbFillAddMonths(dates[0], ms.step * x, ms.eom)));
    }
    const fit = swbFillLinFit(run.cells.map((c) => c.serial));
    return (x) => out(fit.a + fit.b * x);
  }
  if (c0.kind === "textnum") {
    const fit = L === 1 ? { a: c0.n, b: 1 } : swbFillLinFit(run.cells.map((c) => c.n));
    return (x) => {
      let n = Math.abs(Math.round(fit.a + fit.b * x)); // Excel: „x3, x1” → x1, x3, x5 (bez minusa)
      if (c0.quarter) n = ((n - 1) % 4 + 4) % 4 + 1;
      const digits = c0.width ? String(n).padStart(c0.width, "0") : String(n);
      return { value: `${c0.pre}${digits}${c0.suf}`, type: "string" };
    };
  }
  if (c0.kind === "list") {
    const len = c0.list.items.length;
    // indeksy „rozwinięte” przez koniec listy (nd → pon to +1, nie −6)
    const unwrapped = [];
    run.cells.forEach((c, i) => {
      unwrapped.push(i ? unwrapped[i - 1] + ((((c.idx - run.cells[i - 1].idx) % len) + len) % len) : c.idx);
    });
    const step = L === 1 ? 1 : Math.round(swbFillLinFit(unwrapped).b);
    return (x) => {
      const idx = (((c0.idx + step * x) % len) + len) % len;
      return { value: swbFillApplyCase(c0.list.items[idx], c0.caseKind), type: "string" };
    };
  }
  return copy;
}

// sources: [{ v, fmt, formula }] (komórki źródła po kolei), total = długość całego zakresu
// (źródło + cel). Zwraca tablicę dla pozycji total−k celów: { src, value, type } — nowa
// wartość — albo { src, copy: true } — kopia komórki źródła src (formuła z przesunięciem).
// opts.invert = Ctrl/Alt przy puszczeniu uchwytu; opts.copy = zawsze kopia (Ctrl+D).
function swbAutoFillPlan(sources, total, opts = {}) {
  const k = sources.length;
  const lang = opts.lang || (typeof currentLang === "string" ? currentLang : "pl");
  const out = [];
  if (!k || total <= k) return out;
  const cls = sources.map((s) => swbFillClassify(s, lang));
  // ciągi: sąsiednie komórki tego samego rodzaju (klucz) — każdy liczony osobno
  const runs = [];
  const runOf = new Array(k);
  cls.forEach((c, j) => {
    const last = runs[runs.length - 1];
    if (last && c.key && last.key === c.key) last.cells.push(c);
    else runs.push({ key: c.key, start: j, cells: [c] });
    runOf[j] = runs.length - 1;
  });
  const gens = runs.map((run) => (opts.copy ? () => null : swbFillRunGenerator(run, k === 1, !!opts.invert)));
  for (let p = k; p < total; p++) {
    const cyc = Math.floor(p / k);
    const j = p % k;
    const run = runs[runOf[j]];
    const x = cyc * run.cells.length + (j - run.start);
    const val = gens[runOf[j]](x);
    out.push(val ? { src: j, value: val.value, type: val.type } : { src: j, copy: true });
  }
  return out;
}

// ── Zapis do arkusza ─────────────────────────────────────────────────────────────

function swbFillTypeOf(v) {
  if (v instanceof Date) return "date";
  if (typeof v === "number") return "number";
  if (typeof v === "boolean") return "boolean";
  return "string";
}

// Co wypełnianie zniszczyło po drodze — liczone w swbFillWriteCell, zgłaszane po całej
// operacji (swbFillRange). Film 2026-10-31: przeciągnięcie pustej komórki po kolumnie
// z datami i formułami „Długość” po cichu wyczyściło dane — w Excelu też, ale tam od razu
// widać, a tu użytkownik dowiadywał się po fakcie. Teraz mówimy wprost i podpowiadamy Cofnij.
let swbFillDamage = null; // { formulas, cleared }

// Jedna komórka celu. Zwraca "ok" | "formula" (formuły nie dało się wpisać) | "skip".
function swbFillWriteCell(model, srcRow, srcCol, dstRow, dstCol, item) {
  if (!dstRow || dstRow.isLongViewRow || dstRow.isSubheader || !Number.isFinite(dstRow.rowIndex0)) return "skip";
  if (dstCol < 0 || dstCol >= model.headers.length) return "skip";
  if (typeof cfmtIsRealCol === "function" && !cfmtIsRealCol(dstCol)) return "skip";
  if (srcRow === dstRow && srcCol === dstCol) return "skip";
  // format liczbowy NAJPIERW — zapis wartości zachowuje format komórki (i nie nadaje
  // automatycznego, gdy komórka już ma własny)
  if (typeof cfmtCopyNumFmt === "function") cfmtCopyNumFmt(srcRow, srcCol, dstRow, dstCol);
  const srcRef = XLSX.utils.encode_cell({ r: srcRow.rowIndex0, c: currentStartCol + srcCol });
  const srcCell = workbook.Sheets[currentSheetName]?.[srcRef];
  if (item.copy && srcCell && srcCell.f) {
    const text = typeof swbFormulaShifted === "function"
      ? swbFormulaShifted(currentSheetName, srcRef, dstRow.rowIndex0 - srcRow.rowIndex0, dstCol - srcCol) : null;
    return text ? writeCellFromInput(model, dstRow, dstCol, text) : "formula";
  }
  const value = item.copy ? (Array.isArray(srcRow.values) ? srcRow.values[srcCol] : null) : item.value;
  const parsed = value == null || value === "" ? null : { value, type: item.copy ? swbFillTypeOf(value) : item.type };
  const dstRef = XLSX.utils.encode_cell({ r: dstRow.rowIndex0, c: currentStartCol + dstCol });
  const dstHadFormula = !!workbook.Sheets[currentSheetName]?.[dstRef]?.f;
  const cur = Array.isArray(dstRow.values) ? dstRow.values[dstCol] : null;
  if (!parsed) {
    // pusta komórka źródła czyści cel — ale pusty cel zostawiamy w spokoju (bez „zmiany”)
    if ((cur == null || cur === "") && !dstHadFormula) return "skip";
  }
  if (swbFillDamage) {
    if (dstHadFormula) swbFillDamage.formulas += 1;
    else if (!parsed && cur != null && cur !== "") swbFillDamage.cleared += 1;
  }
  updateSheetCell(dstRow.rowIndex0, dstCol, parsed);
  if (Array.isArray(dstRow.values)) dstRow.values[dstCol] = parsed ? parsed.value : null;
  if (Array.isArray(dstRow.rawValues) && dstRow.rawValues !== dstRow.values) dstRow.rawValues[dstCol] = parsed ? parsed.value : null;
  if (typeof cfmtRefreshDisplay === "function") cfmtRefreshDisplay(dstRow, dstCol);
  else if (Array.isArray(dstRow.display)) dstRow.display[dstCol] = parsed ? toDisplay(parsed.value) : "";
  return "ok";
}

// Wypełnienie prostokąta: src = { rowStart, rowEnd, colMin, colMax } (pozycje w widoku),
// dir = "down" | "right", toEnd = ostatni wiersz / kolumna celu. mode: "series" | "copy".
function swbFillRange(model, src, dir, toEnd, { mode = "series", invert = false, quiet = false } = {}) {
  swbFillDamage = { formulas: 0, cleared: 0 };
  let changed = 0;
  let skippedFormula = 0;
  const bump = (res) => { if (res === "ok") changed += 1; else if (res === "formula") skippedFormula += 1; };
  const sheetFmt = (row, c) => (typeof cfmtNumFmtAt === "function" ? cfmtNumFmtAt(row.rowIndex0, c) : "General");
  const isFormula = (row, c) => !!workbook.Sheets[currentSheetName]?.[XLSX.utils.encode_cell({ r: row.rowIndex0, c: currentStartCol + c })]?.f;
  const lineSources = (cells) => cells.map(([row, c]) => ({ v: row && Array.isArray(row.values) ? row.values[c] : null, fmt: row ? sheetFmt(row, c) : null, formula: row ? isFormula(row, c) : false }));
  if (dir === "down") {
    const k = src.rowEnd - src.rowStart + 1;
    const total = toEnd - src.rowStart + 1;
    for (let c = src.colMin; c <= src.colMax; c++) {
      const cells = [];
      for (let r = src.rowStart; r <= src.rowEnd; r++) cells.push([model.rows[r], c]);
      const plan = swbAutoFillPlan(lineSources(cells), total, { copy: mode === "copy", invert });
      plan.forEach((item, i) => {
        const [srcRow] = cells[item.src];
        bump(swbFillWriteCell(model, srcRow, c, model.rows[src.rowStart + k + i], c, item));
      });
    }
  } else {
    const k = src.colMax - src.colMin + 1;
    const total = toEnd - src.colMin + 1;
    for (let r = src.rowStart; r <= src.rowEnd; r++) {
      const row = model.rows[r];
      if (!row) continue;
      const cells = [];
      for (let c = src.colMin; c <= src.colMax; c++) cells.push([row, c]);
      const plan = swbAutoFillPlan(lineSources(cells), total, { copy: mode === "copy", invert });
      plan.forEach((item, i) => bump(swbFillWriteCell(model, row, src.colMin + item.src, row, src.colMin + k + i, item)));
    }
  }
  if (changed > 0) {
    setDirtyState(true);
    renderActiveTable();
    if (typeof scheduleViewRefresh === "function") scheduleViewRefresh({ analyses: true, filterBadge: true });
    if (!quiet) toast(t(mode === "copy" ? "cellsFilled" : "cellsFilledSeries", { count: changed }), "success");
    if (typeof sheetEditUi !== "undefined") sheetEditUi.sync();
    log(`Wypelniono ${changed} komorek (${mode}, ${dir})`, "success");
  } else if (!quiet) {
    toast(t("editToolNoChange"), "info");
  }
  if (skippedFormula > 0) toast(t("fillSkippedFormulas", { count: skippedFormula }), "warning");
  const damage = swbFillDamage;
  swbFillDamage = null;
  if (changed > 0 && (damage.formulas || damage.cleared)) {
    const parts = [];
    if (damage.formulas) parts.push(t("fillOverwroteFormulas", { count: damage.formulas }));
    if (damage.cleared) parts.push(t("fillClearedCells", { count: damage.cleared }));
    toast(t("fillDamageWarn", { what: parts.join(", ") }), "warning");
  }
  return changed;
}

function swbFillCanEdit() {
  if (typeof swbEditGuard === "function" && !swbEditGuard()) return false;
  if (!workbook || !currentDisplayModel || currentDisplayModel.mode !== "wide") {
    toast(t("editWideOnly"), "info");
    return false;
  }
  return true;
}

// Ctrl+D / Ctrl+R (i przyciski „Wypełnij w dół / w prawo”): KOPIA, jak w Excelu.
// Zakres wielu wierszy → pierwszy wiersz rozlewa się w dół; jeden wiersz (także sama
// komórka) → kopiuje wiersz NAD zaznaczeniem. Analogicznie w prawo z kolumną po lewej.
function swbFillCopy(direction) {
  if (!swbFillCanEdit()) return 0;
  const rect = rangeCellsToCopy();
  if (!rect) return 0;
  const model = rect.model;
  const src = { rowStart: rect.rowStart, rowEnd: rect.rowStart, colMin: rect.colMin, colMax: rect.colMax };
  if (direction === "down") {
    if (rect.rowCount === 1) {
      if (rect.rowStart === 0) { toast(t("fillNothingAbove"), "info"); return 0; }
      src.rowStart = src.rowEnd = rect.rowStart - 1;
    }
    return swbFillRange(model, src, "down", rect.rowEnd, { mode: "copy" });
  }
  src.rowEnd = rect.rowEnd;
  src.colMax = rect.colMin;
  if (rect.colCount === 1) {
    if (rect.colMin === 0) { toast(t("fillNothingLeft"), "info"); return 0; }
    src.colMin = src.colMax = rect.colMin - 1;
  }
  return swbFillRange(model, src, "right", rect.colMax, { mode: "copy" });
}

// ── Uchwyt wypełniania ───────────────────────────────────────────────────────────
// Kwadracik w prawym dolnym rogu zaznaczenia (albo aktywnej komórki). Przeciągnięcie
// w dół lub w prawo = seria (jak w Excelu), z Ctrl/Alt przy puszczeniu = odwrotnie
// (kopia zamiast serii, a dla jednej liczby seria +1). Dwuklik = w dół do końca danych
// w sąsiedniej kolumnie. Działa myszą i palcem (touch-action: none na samym uchwycie).
//
// Gdzie stoi: NIE w komórce. <td> ma overflow: hidden (wielokropek długiego tekstu), więc
// przycinał kwadracik do wnętrza ramki zaznaczenia. Uchwyt leży w #tableWrap (position:
// relative — przewija się razem z treścią) na narożniku ramki i wystaje na zewnątrz jak
// w Excelu. Pozycję liczymy w rAF (raz na klatkę, po naturalnym layoucie) i odświeżamy,
// gdy tabela zmieni rozmiar (szerokość kolumny, wysokość wiersza, zoom). Wyjątek: róg
// w ZAMROŻONEJ kolumnie (sticky) — tam uchwyt zostaje w komórce (klasa is-inset), bo
// w #tableWrap odjeżdżałby od niej przy przewijaniu w bok.
//
// PALEC: ten sam róg ROZCIĄGA ZAZNACZENIE (w każdą stronę, jak kółka w Arkuszach Google /
// Excelu na iPhonie), a niczego nie wypełnia. Dwa filmy z iPhone'a (2026-10-31): (1) palec,
// który chciał przesunąć tabelę w bok, złapał uchwyt i seria dat nadpisała formuły w kolejnych
// kolumnach; (2) nie dało się w ogóle zaznaczyć zakresu — Shift+klik nie istnieje na telefonie,
// a róg był wyłącznie wypełnianiem. Wypełnianie na dotyku = przytrzymanie zaznaczenia → menu
// („Wypełnij serią / w dół / w prawo”, filter-bar.js) — świadomy wybór zamiast gestu, który
// łatwo zrobić niechcący. Mysz i rysik zostają przy wypełnianiu jak w Excelu.

const swbFillHandle = (() => {
  if (typeof document === "undefined" || typeof tbodyEl === "undefined" || !tbodyEl) return { sync() {} };
  const el = document.createElement("span");
  el.className = "fill-handle";
  el.setAttribute("aria-hidden", "true");
  el.dataset.hint = typeof t === "function" ? t("fillHandleHint") : "";
  let drag = null;
  let host = null;      // komórka-róg, przy której stoi uchwyt
  let placeFrame = 0;
  let lastPointerType = "";

  function sourceRect() {
    const rect = typeof rangeCellsToCopy === "function" ? rangeCellsToCopy() : null;
    if (!rect || !rect.model || rect.model.mode !== "wide") return null;
    return rect;
  }

  function cornerTd(rect) {
    const row = rect.model.rows[rect.rowEnd];
    if (!row) return null;
    return findCellElement({ rowKey: getRowSelectionKey(row), colIndex0: rect.colMax });
  }

  // Kółko zaznaczania (palec) jest też w Podglądzie — zaznaczanie to nie edycja, a Excel na
  // telefonie pozwala zaznaczać zawsze. Kwadracik wypełniania (mysz) — tylko w Edycji.
  const touchMQ = typeof matchMedia === "function" ? matchMedia("(pointer: coarse)") : null;
  function allowed(pointerType) {
    if (!workbook || !currentDisplayModel || currentDisplayModel.mode !== "wide") return false;
    const edit = typeof swbIsEditMode !== "function" || swbIsEditMode();
    if (!edit) {
      if (pointerType ? pointerType !== "touch" : !(touchMQ && touchMQ.matches)) return false;
      if (typeof isCellSelectionMode === "function" && !isCellSelectionMode()) return false;
    }
    if (typeof activeCellEditor !== "undefined" && activeCellEditor) return false;
    return true;
  }

  function detach() {
    if (host) { host.classList.remove("has-fill-handle"); host = null; }
    if (el.parentElement) el.remove();
  }

  const inFrozenCol = (td) => tableWrapEl.classList.contains("freeze-first-col") && td.cellIndex <= 1;

  // Środek kwadracika = narożnik ramki (środek jej 2-pikselowej kreski), nie zewnętrzny róg
  // komórki: clientWidth/Height pomijają obramowanie dolne komórki.
  function place() {
    placeFrame = 0;
    if (!host || el.parentElement !== tableWrapEl) return;
    if (!host.isConnected) { detach(); return; }
    const wrap = tableWrapEl.getBoundingClientRect();
    const b = host.getBoundingClientRect();
    const x = b.left + host.clientLeft + host.clientWidth - 1 - wrap.left - tableWrapEl.clientLeft + tableWrapEl.scrollLeft;
    const y = b.top + host.clientTop + host.clientHeight - 1 - wrap.top - tableWrapEl.clientTop + tableWrapEl.scrollTop;
    el.style.transform = `translate(${x}px, ${y}px)`;
  }
  function schedulePlace() {
    if (!placeFrame) placeFrame = requestAnimationFrame(place);
  }
  if (typeof ResizeObserver === "function" && typeof tableEl !== "undefined" && tableEl) {
    new ResizeObserver(() => { if (host && !drag) schedulePlace(); }).observe(tableEl);
  }

  function sync() {
    if (drag) return;
    const rect = allowed() ? sourceRect() : null;
    const td = rect && typeof cfmtIsRealCol === "function" && cfmtIsRealCol(rect.colMax) ? cornerTd(rect) : null;
    if (!td) { detach(); return; }
    if (host !== td) {
      if (host) host.classList.remove("has-fill-handle");
      host = td;
      td.classList.add("has-fill-handle");
    }
    const inset = inFrozenCol(td);
    el.classList.toggle("is-inset", inset);
    const parent = inset ? td : tableWrapEl;
    if (el.parentElement !== parent) parent.appendChild(el);
    if (inset) el.style.transform = "";
    else schedulePlace();
  }

  function cellAt(x, y) {
    const hit = document.elementFromPoint(x, y);
    // dane albo wiersz nad nagłówkiem (oba mają tr[data-row-key]) — indeks w widoku zaznaczenia
    const td = hit && hit.closest ? hit.closest("#dataTable tr[data-row-key] > td[data-col-index]") : null;
    if (!td) return null;
    const tr = td.parentElement;
    const r = swbSelIndexOfKey(drag ? drag.rect.model : swbSelView(), tr.dataset.rowKey);
    const c = parseInt(td.dataset.colIndex, 10);
    return r >= 0 && Number.isFinite(c) ? { r, c } : null;
  }

  function preview(target) {
    if (drag.extend) {
      // rozciąganie zaznaczenia: ruchomy koniec idzie za palcem w dowolną stronę
      if (target && (target.r !== drag.end?.r || target.c !== drag.end?.c)) {
        drag.end = target;
        setSelectedCell(getRowSelectionKey(drag.rect.model.rows[target.r]), target.c, { scroll: false });
      }
      return;
    }
    const { rect } = drag;
    let dir = null;
    let end = null;
    if (target) {
      const dRow = target.r - rect.rowEnd;
      const dCol = target.c - rect.colMax;
      if (dRow > 0 && dRow >= dCol) { dir = "down"; end = target.r; }
      else if (dCol > 0) { dir = "right"; end = target.c; }
    }
    drag.dir = dir;
    drag.end = end;
    const model = rect.model;
    const endRow = dir === "down" ? end : rect.rowEnd;
    const endCol = dir === "right" ? end : rect.colMax;
    setSelectedCell(getRowSelectionKey(model.rows[endRow]), endCol, { scroll: false });
  }

  function autoScroll() {
    if (!drag || !tableWrapEl) return;
    const box = tableWrapEl.getBoundingClientRect();
    const { x, y } = drag.last;
    let dy = 0;
    let dx = 0;
    if (y > box.bottom - 28) dy = 18;
    else if (y < box.top + 40) dy = -18;
    if (x > box.right - 28) dx = 18;
    else if (drag.extend && x < box.left + 28) dx = -18; // zaznaczenie rośnie też w lewo
    if (dy || dx) {
      tableWrapEl.scrollBy(dx, dy);
      preview(cellAt(Math.min(Math.max(x, box.left + 4), box.right - 4), Math.min(Math.max(y, box.top + 2), box.bottom - 4)));
    }
    drag.raf = requestAnimationFrame(autoScroll);
  }

  el.addEventListener("pointerdown", (e) => {
    lastPointerType = e.pointerType || "";
    if (e.button !== 0 || !allowed(e.pointerType)) return;
    const rect = sourceRect();
    if (!rect) return;
    e.preventDefault();
    e.stopPropagation();
    const model = rect.model;
    const extend = e.pointerType === "touch";
    if (extend) {
      // kotwica zostaje tam, gdzie była (aktywna komórka); zakres = od niej do palca
      setSelectionKind("cell");
      if (!focusedCellState) setFocusedCell(getRowSelectionKey(model.rows[rect.rowStart]), rect.colMin, { scroll: false });
      swbTouchFillOrigin = swbRectKeys(rect);
    } else {
      // kotwica w lewym górnym rogu źródła, ruchomy koniec = podgląd celu
      setFocusedCell(getRowSelectionKey(model.rows[rect.rowStart]), rect.colMin, { scroll: false });
      setSelectedCell(getRowSelectionKey(model.rows[rect.rowEnd]), rect.colMax, { scroll: false });
    }
    drag = { rect: { ...rect }, extend, dir: null, end: null, last: { x: e.clientX, y: e.clientY }, id: e.pointerId, raf: 0 };
    // Nasłuch na oknie, nie na uchwycie: przy przewijaniu (nowy silnik tabeli) wiersz
    // z uchwytem może zniknąć z DOM, a razem z nim zdarzenia wysyłane do uchwytu.
    window.addEventListener("pointermove", onMove, true);
    window.addEventListener("pointerup", onUp, true);
    window.addEventListener("pointercancel", onCancel, true);
    document.documentElement.classList.add("swb-filling");
    drag.raf = requestAnimationFrame(autoScroll);
  });
  function onMove(e) {
    if (!drag || e.pointerId !== drag.id) return;
    e.preventDefault();
    drag.last = { x: e.clientX, y: e.clientY };
    preview(cellAt(e.clientX, e.clientY));
  }
  function onUp(e) { finish(e, false); }
  function onCancel(e) { finish(e, true); }
  const finish = (e, cancelled) => {
    if (!drag || (e && e.pointerId !== drag.id)) return;
    const d = drag;
    drag = null;
    cancelAnimationFrame(d.raf);
    window.removeEventListener("pointermove", onMove, true);
    window.removeEventListener("pointerup", onUp, true);
    window.removeEventListener("pointercancel", onCancel, true);
    document.documentElement.classList.remove("swb-filling");
    const model = d.rect.model;
    if (d.extend) {
      sync();
      if (!cancelled && d.end) swbTouchRangeHintOnce();
      return;
    }
    if (cancelled || !d.dir || !swbFillCanEdit()) {
      setSelectedCell(getRowSelectionKey(model.rows[d.rect.rowEnd]), d.rect.colMax, { scroll: false });
      sync();
      return;
    }
    const invert = !!(e && (e.ctrlKey || e.altKey || e.metaKey));
    swbFillRange(model, d.rect, d.dir, d.end, { mode: "series", invert });
    // zaznaczony zostaje cały wypełniony zakres, jak w Excelu
    const m = swbSelView();
    const endRow = d.dir === "down" ? d.end : d.rect.rowEnd;
    const endCol = d.dir === "right" ? d.end : d.rect.colMax;
    if (m && m.rows[d.rect.rowStart] && m.rows[endRow]) {
      setFocusedCell(getRowSelectionKey(m.rows[d.rect.rowStart]), d.rect.colMin, { scroll: false, keepSelection: true });
      setSelectedCell(getRowSelectionKey(m.rows[endRow]), endCol, { scroll: false });
    }
    sync();
  };
  el.addEventListener("click", (e) => { e.stopPropagation(); e.preventDefault(); });
  el.addEventListener("dblclick", (e) => {
    e.preventDefault();
    e.stopPropagation();
    // dwutap palcem to nie „wypełnij do końca danych” — na dotyku róg tylko zaznacza
    if (lastPointerType === "touch") return;
    if (!swbFillCanEdit()) return;
    const rect = sourceRect();
    if (!rect) return;
    const end = swbFillDownLimit(rect);
    if (end <= rect.rowEnd) { toast(t("fillNoNeighbor"), "info"); return; }
    swbFillRange(rect.model, rect, "down", end, { mode: "series" });
    const m = swbSelView();
    setFocusedCell(getRowSelectionKey(m.rows[rect.rowStart]), rect.colMin, { scroll: false, keepSelection: true });
    setSelectedCell(getRowSelectionKey(m.rows[end]), rect.colMax, { scroll: true });
    sync();
  });
  document.addEventListener("keydown", (e) => {
    if (e.key === "Escape" && drag) { e.preventDefault(); e.stopPropagation(); finish(null, true); }
  }, true);

  return { sync, get dragging() { return !!drag; } };
})();

// ── Wypełnianie zaznaczenia z menu (dotyk) ───────────────────────────────────────
// Zaznaczenie sprzed rozciągnięcia palcem — to ono jest „źródłem” serii, jak przy
// przeciąganiu uchwytu myszą (np. 1, 2 zaznaczone i rozciągnięte w dół → 3, 4, 5…).
// Trzymane po kluczach wierszy, bo model widoku podmienia się przy każdym renderze.
let swbTouchFillOrigin = null; // { rowStartKey, rowEndKey, colMin, colMax }

function swbRectKeys(rect) {
  return {
    rowStartKey: getRowSelectionKey(rect.model.rows[rect.rowStart]),
    rowEndKey: getRowSelectionKey(rect.model.rows[rect.rowEnd]),
    colMin: rect.colMin,
    colMax: rect.colMax,
  };
}

// Źródło i kierunek serii dla zaznaczonego zakresu: zapamiętany początek, gdy zakres jest
// jego przedłużeniem w dół albo w prawo; inaczej pierwszy wiersz (zakres wysoki) albo
// pierwsza kolumna (jeden wiersz) — jak „Wypełnij → Seria” w Excelu.
function swbFillSeriesPlan(rect) {
  if (!rect || (rect.rowCount === 1 && rect.colCount === 1)) return null;
  const o = swbTouchFillOrigin;
  if (o) {
    const idx = (key) => rect.model.rows.findIndex((row) => getRowSelectionKey(row) === key);
    const rs = idx(o.rowStartKey);
    const re = idx(o.rowEndKey);
    const src = { rowStart: rs, rowEnd: re, colMin: o.colMin, colMax: o.colMax };
    if (rs === rect.rowStart && o.colMin === rect.colMin && re >= rs) {
      if (o.colMax === rect.colMax && re < rect.rowEnd) return { src, dir: "down", end: rect.rowEnd };
      if (re === rect.rowEnd && o.colMax < rect.colMax) return { src, dir: "right", end: rect.colMax };
    }
  }
  if (rect.rowCount > 1) return { src: { rowStart: rect.rowStart, rowEnd: rect.rowStart, colMin: rect.colMin, colMax: rect.colMax }, dir: "down", end: rect.rowEnd };
  return { src: { rowStart: rect.rowStart, rowEnd: rect.rowEnd, colMin: rect.colMin, colMax: rect.colMin }, dir: "right", end: rect.colMax };
}

function swbFillSeriesSelection() {
  if (!swbFillCanEdit()) return 0;
  const rect = typeof getSelectionRectangle === "function" ? getSelectionRectangle() : null;
  const plan = swbFillSeriesPlan(rect);
  if (!plan) { toast(t("fillNeedRange"), "info"); return 0; }
  const n = swbFillRange(rect.model, plan.src, plan.dir, plan.end, { mode: "series" });
  swbTouchFillOrigin = null;
  return n;
}

// Jednorazowa podpowiedź po pierwszym rozciągnięciu palcem: bez niej nikt nie zgadnie,
// że wypełnianie przeszło do menu przytrzymania.
const SWB_TOUCH_RANGE_HINT_KEY = "swb-touch-range-hint-v1";
function swbTouchRangeHintOnce() {
  try {
    if (localStorage.getItem(SWB_TOUCH_RANGE_HINT_KEY)) return;
    localStorage.setItem(SWB_TOUCH_RANGE_HINT_KEY, "1");
  } catch { return; }
  toast(t("touchRangeHint"), "info");
}

// Dwuklik uchwytu: do ostatniego wiersza ciągłych danych w kolumnie obok (lewej, a gdy
// pusta — prawej), jak w Excelu.
function swbFillDownLimit(rect) {
  const rows = rect.model.rows;
  const filled = (r, c) => {
    const row = rows[r];
    const v = row && Array.isArray(row.values) ? row.values[c] : null;
    return v != null && v !== "";
  };
  for (const c of [rect.colMin - 1, rect.colMax + 1]) {
    if (c < 0 || c >= rect.model.headers.length) continue;
    if (!filled(rect.rowEnd + 1, c)) continue;
    let r = rect.rowEnd + 1;
    while (r + 1 < rows.length && filled(r + 1, c)) r++;
    return r;
  }
  return rect.rowEnd;
}

function swbFillHandleSync() {
  if (typeof swbFillHandle !== "undefined") swbFillHandle.sync();
}
