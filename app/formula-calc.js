// formula-calc.js — formuły w aplikacji: wpisywanie, przeliczanie zależnych, zapis.
//
// Silnik (app/formula-engine.js) jest czysty; tu jest klej z resztą aplikacji:
// - kontekst liczenia na żywym skoroszycie (workbook.Sheets — te same obiekty, które
//   zmienia edycja), z tabelami Excela i nazwami zdefiniowanymi,
// - indeks zależności: która formuła czyta którą komórkę (budowany leniwie przy pierwszej
//   zmianie — plik bez formuł nic nie kosztuje),
// - przeliczenie po każdej zmianie wartości (edytor, wklejanie, wypełnianie, Delete,
//   Cofnij…): zbieramy zmienione komórki w jednym zadaniu i liczymy RAZ, w mikrozadaniu,
// - wpisanie formuły (swbFormulaCommit) i wypełnianie formułą z przesunięciem odwołań.
//
// Wynik przeliczenia trafia do komórek w pamięci (tabela pokazuje aktualne liczby).
// Do pliku zapisujemy tylko formuły wpisane przez użytkownika; resztę i tak przelicza
// Excel przy otwarciu (xlsx-patch ustawia fullCalcOnLoad).

const FC = typeof FormulaEngine !== "undefined" ? FormulaEngine : null;

// ── kontekst liczenia ────────────────────────────────────────────────────────
let fcTablesCache = { src: null, out: null };
function fcTables() {
  const src = typeof currentTables !== "undefined" ? currentTables : null;
  if (fcTablesCache.src === src) return fcTablesCache.out;
  const out = {};
  Object.entries(src || {}).forEach(([key, t]) => {
    if (!t || !t.ref || !t.sheetName) return;
    let rng;
    try { rng = XLSX.utils.decode_range(t.ref); } catch { return; }
    const names = Array.isArray(t.colNames) && t.colNames.length
      ? t.colNames
      : Object.entries(t.columns || {}).sort((a, b) => a[1] - b[1]).map(([n]) => n);
    out[key] = { sheet: t.sheetName, r1: rng.s.r, c1: rng.s.c, r2: rng.e.r, c2: rng.e.c, columns: names, headerRows: t.headerRows ?? 1, totalsRows: t.totalsRows || 0 };
  });
  fcTablesCache = { src, out };
  return out;
}

function fcCellValue(cell) {
  if (!cell || cell.t === "z") return null;
  if (cell.t === "e") return FC.ERR[cell.w] || FC.ERR[cell.v] || { err: String(cell.w || "#VALUE!") };
  if (cell.v instanceof Date) return FC.jsDateToSerial(cell.v);
  if (cell.t === "b") return !!cell.v;
  if (cell.t === "n" || cell.t === "d") return typeof cell.v === "number" ? cell.v : (cell.v == null ? null : Number(cell.v));
  if (cell.v === undefined || cell.v === null) return null;
  return String(cell.v);
}

let fcSheetLookup = { wb: null, map: null };
function fcSheetFor(name) {
  if (fcSheetLookup.wb !== workbook) fcSheetLookup = { wb: workbook, map: new Map(workbook.SheetNames.map((n) => [n.toLowerCase(), n])) };
  return fcSheetLookup.map.get(String(name).toLowerCase()) ?? null;
}

let fcBaseCtx = null;
function fcContext(sheetName, r, c) {
  if (!fcBaseCtx || fcBaseCtx.wb !== workbook || fcBaseCtx.tablesSrc !== (typeof currentTables !== "undefined" ? currentTables : null)) {
    const boundsCache = new Map();
    fcBaseCtx = {
      wb: workbook,
      tablesSrc: typeof currentTables !== "undefined" ? currentTables : null,
      depth: 0,
      get: (sh, rr, cc) => {
        const ws = workbook.Sheets[sh];
        return ws ? fcCellValue(ws[XLSX.utils.encode_cell({ r: rr, c: cc })]) : null;
      },
      // tekst formuły komórki (SUMY.CZĘŚCIOWE pomija komórki z własnym SUBTOTAL)
      formula: (sh, rr, cc) => {
        const ws = workbook.Sheets[sh];
        const cell = ws && ws[XLSX.utils.encode_cell({ r: rr, c: cc })];
        return cell && cell.f ? String(cell.f) : "";
      },
      bounds: (sh) => {
        const ws = workbook.Sheets[sh];
        const ref = ws && ws["!ref"];
        let b = boundsCache.get(sh);
        if (!b || b.ref !== ref) {
          const d = ref ? XLSX.utils.decode_range(ref) : { e: { r: 0, c: 0 } };
          b = { ref, r2: d.e.r, c2: d.e.c };
          boundsCache.set(sh, b);
        }
        return b;
      },
      sheetFor: fcSheetFor,
      now: () => FC.jsDateToSerial(new Date()),
      structured: FC.makeStructuredResolver(fcTables()),
      name: FC.makeNameResolver((workbook.Workbook && workbook.Workbook.Names) || [], workbook.SheetNames),
      formatNumber: (fmt, n) => { try { return XLSX.SSF.format(fmt, n); } catch { return null; } },
    };
  }
  const lang = typeof currentLang === "string" ? currentLang : "pl";
  FC.setLocale(lang); // „1 234,5” jako liczba po polsku, „1,000.5” po angielsku
  return { ...fcBaseCtx, sheetName, row: r, col: c, lang };
}

// ── indeks zależności ───────────────────────────────────────────────────────
// entries: key "arkusz\u0000A1" → { sheet, r, c, f, ast|null, ok, deps:[{sheet,r1,c1,r2,c2}] }
// bySheet: arkusz → [{ r1,c1,r2,c2, key }] (zakresy CZYTANE przez formuły)
let fcIndex = null;
const fcAstCache = new Map();

function fcParseFile(f, r, c) {
  let e = fcAstCache.get(f);
  if (e === undefined) {
    e = r == null ? undefined : fcFromShape(f, r, c);
    if (e === undefined) {
      try {
        const ast = FC.parse(f, { lang: "en" });
        const fns = [...FC.functionsOf(ast)];
        e = { ast, ok: fns.every((n) => FC.isSupported(n)) };
      } catch { e = { ast: null, ok: false }; }
      if (r != null) fcShapeRemember(f, r, c, e);
    }
    if (fcAstCache.size > 50000) fcAstCache.clear();
    fcAstCache.set(f, e);
  }
  return e;
}

// ── wspólne „kształty” formuł ───────────────────────────────────────────────
// Plik z 28 tys. formuł ma ich zwykle kilkadziesiąt RÓŻNYCH: „=B2*C2” w wierszu 2 i „=B3*C3”
// w wierszu 3 to ta sama formuła przesunięta o wiersz. Klucz kształtu = tekst, w którym
// względne części adresów są zapisane jako odległość od komórki formuły (jak R1C1 w Excelu).
// Ten sam kształt → raz parsowane drzewo, kolejne komórki dostają FC.shift (to samo, co
// robi wypełnianie formułą). Bezpiecznik: przy PIERWSZYM parsowaniu kształtu adresy znalezione
// tu skanerem muszą się zgadzać co do sztuki z adresami w drzewie parsera (w tym zakresy
// kolumn/wierszy, których skaner nie przesuwa) — inaczej kształt jest wyłączony i każda
// komórka parsuje się jak dawniej. Wynik zawsze identyczny z FC.parse.
const fcShapes = new Map(); // klucz → { r, c, e } | null (kształt niebezpieczny)
const FC_REF_AT = /(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7})(?![A-Za-z0-9_(.!\[À-ɏ])/y;
const FC_WORD_CH = /[A-Za-z0-9_.\\$À-ɏ]/;

function fcColIdx(letters) {
  let n = 0;
  for (let i = 0; i < letters.length; i++) n = n * 26 + (letters.charCodeAt(i) & 0x1f);
  return n - 1;
}

// → { key, refs: [[r, c, ar, ac], …] } albo null (nie da się bezpiecznie opisać kształtu)
function fcShapeOf(f, r, c) {
  const s = String(f);
  let key = "";
  let last = 0;
  const refs = [];
  for (let i = 0; i < s.length; i++) {
    const ch = s[i];
    if (ch === '"' || ch === "'") { // literał tekstowy / nazwa arkusza w apostrofach — bez zmian
      const j = s.indexOf(ch, i + 1);
      if (j < 0) return null;
      i = j;
      continue;
    }
    if (i > 0 && FC_WORD_CH.test(s[i - 1])) continue;
    FC_REF_AT.lastIndex = i;
    const m = FC_REF_AT.exec(s);
    if (!m) continue;
    const rr = +m[4] - 1;
    const cc = fcColIdx(m[2]);
    const ar = !!m[3];
    const ac = !!m[1];
    refs.push([rr, cc, ar, ac]);
    key += s.slice(last, i) + "\u0001" + (ac ? "$" + cc : cc - c) + "," + (ar ? "$" + rr : rr - r) + "\u0002";
    last = i + m[0].length;
    i = last - 1;
  }
  return { key: key + s.slice(last), refs };
}

function fcShapeRefsMatch(ast, refs) {
  const fromAst = [];
  let bad = false;
  const walk = (n) => {
    if (!n || typeof n !== "object" || bad) return;
    if (n.t === "ref") fromAst.push(`${n.r},${n.c},${!!n.ar},${!!n.ac}`);
    else if (n.t === "range") {
      if (n.cols || n.rows) { bad = true; return; }
      fromAst.push(`${n.r1},${n.c1},${!!n.ar1},${!!n.ac1}`, `${n.r2},${n.c2},${!!n.ar2},${!!n.ac2}`);
    } else if (n.t === "call") n.args.forEach(walk);
    else if (n.t === "bin") { walk(n.a); walk(n.b); }
    else if (n.t === "un" || n.t === "pct" || n.t === "paren") walk(n.a);
    else if (n.t === "array") n.rows.forEach((row) => row.forEach(walk));
  };
  walk(ast);
  if (bad || fromAst.length !== refs.length) return false;
  const a = fromAst.sort();
  const b = refs.map(([rr, cc, ar, ac]) => `${rr},${cc},${ar},${ac}`).sort();
  return a.every((x, i) => x === b[i]);
}

function fcFromShape(f, r, c) {
  const sh = fcShapeOf(f, r, c);
  if (!sh || !sh.refs.length) return undefined;
  const base = fcShapes.get(sh.key);
  if (!base) return undefined;
  return { ast: FC.shift(base.e.ast, r - base.r, c - base.c), ok: base.e.ok };
}

function fcShapeRemember(f, r, c, e) {
  if (!e.ast) return;
  const sh = fcShapeOf(f, r, c);
  if (!sh || !sh.refs.length || fcShapes.has(sh.key)) return;
  if (fcShapes.size > 5000) fcShapes.clear();
  fcShapes.set(sh.key, fcShapeRefsMatch(e.ast, sh.refs) ? { r, c, e } : null);
}

function fcDepsOf(ast, sheet, r, c) {
  const out = [];
  if (!ast) return out;
  const ctx = fcContext(sheet, r, c);
  FC.refsOf(ast).forEach((ref) => {
    if (ref.table) {
      const rng = ctx.structured(ref.table, ref.spec, ctx);
      if (rng) out.push({ sheet: rng.sheet, r1: rng.r1, c1: rng.c1, r2: rng.r2, c2: rng.c2 });
      return;
    }
    if (ref.name) {
      const def = ctx.name(ref.name, ref.sheet, ctx);
      if (def && FC.isRange(def)) out.push({ sheet: def.sheet, r1: def.r1, c1: def.c1, r2: def.r2, c2: def.c2 });
      return;
    }
    const sh = ref.sheet == null ? sheet : fcSheetFor(ref.sheet);
    if (sh == null) return;
    out.push({ sheet: sh, r1: Math.min(ref.r1, ref.r2), c1: Math.min(ref.c1, ref.c2), r2: Math.max(ref.r1, ref.r2), c2: Math.max(ref.c1, ref.c2) });
  });
  return out;
}

function fcIndexAdd(sheet, ref, cell, idx = fcIndex) {
  const { r, c } = XLSX.utils.decode_cell(ref);
  const key = `${sheet}\u0000${ref}`;
  const p = fcParseFile(cell.f, r, c);
  const deps = fcDepsOf(p.ast, sheet, r, c);
  idx.entries.set(key, { sheet, ref, r, c, f: cell.f, ast: p.ast, ok: p.ok && !cell.F, deps });
  deps.forEach((d) => {
    if (!idx.bySheet.has(d.sheet)) idx.bySheet.set(d.sheet, []);
    idx.bySheet.get(d.sheet).push({ ...d, key });
  });
}

function fcIndexRemove(key) {
  const e = fcIndex && fcIndex.entries.get(key);
  if (!e) return;
  fcIndex.entries.delete(key);
  new Set(e.deps.map((d) => d.sheet)).forEach((sh) => {
    const list = fcIndex.bySheet.get(sh);
    if (list) fcIndex.bySheet.set(sh, list.filter((x) => x.key !== key));
  });
}

// Budowa indeksu w porcjach (swbFormulaPrewarm): duży plik (28 tys. formuł) budował go
// ~0,3–0,5 s przy PIERWSZEJ edycji — Enter po otwarciu pliku stał. Teraz po wczytaniu
// pliku buduje się w bezczynności po kilka ms; gdy edycja przyjdzie wcześniej,
// fcEnsureIndex dokańcza resztę od razu (nigdy gorzej niż dawniej). Do końca budowy
// fcIndex zostaje pusty — swbFormulaCellChanged nic nie psuje, a budowa czyta stan
// komórek z chwili, w której do nich dochodzi.
let fcIndexBuild = null; // { wb, idx, sheets, sheetI, refs, refI, ms }
function fcBuildStep(budgetMs) {
  const b = fcIndexBuild;
  const t0 = performance.now();
  const end = t0 + budgetMs;
  while (b.sheetI < b.sheets.length) {
    const sn = b.sheets[b.sheetI];
    const ws = b.wb.Sheets[sn];
    if (!b.refs) { b.refs = ws ? Object.keys(ws) : []; b.refI = 0; }
    while (b.refI < b.refs.length) {
      const ref = b.refs[b.refI++];
      if (ref.charCodeAt(0) !== 33) { const cell = ws[ref]; if (cell && cell.f) fcIndexAdd(sn, ref, cell, b.idx); }
      if ((b.refI & 127) === 0 && performance.now() > end) { b.ms += performance.now() - t0; return false; }
    }
    b.sheetI++;
    b.refs = null;
  }
  b.ms += performance.now() - t0;
  return true;
}
function fcFinishBuild() {
  const b = fcIndexBuild;
  fcIndexBuild = null;
  fcIndex = b.idx;
  if (typeof log === "function" && fcIndex.entries.size) log(`Formuły: indeks zależności ${fcIndex.entries.size} kom. (${Math.round(b.ms)} ms)`, "info");
}
function fcStartBuild() {
  fcIndexBuild = { wb: workbook, idx: { wb: workbook, entries: new Map(), bySheet: new Map() }, sheets: workbook.SheetNames.slice(), sheetI: 0, refs: null, refI: 0, ms: 0 };
}

function fcEnsureIndex() {
  if (!FC || typeof workbook === "undefined" || !workbook) return null;
  if (fcIndex && fcIndex.wb === workbook) return fcIndex;
  if (!fcIndexBuild || fcIndexBuild.wb !== workbook) fcStartBuild();
  fcBuildStep(Infinity);
  fcFinishBuild();
  return fcIndex;
}

// Po wczytaniu pliku (handleFile): indeks w tle, porcjami w bezczynności.
function swbFormulaPrewarm() {
  if (!FC || typeof workbook === "undefined" || !workbook) return;
  if ((fcIndex && fcIndex.wb === workbook) || (fcIndexBuild && fcIndexBuild.wb === workbook)) return;
  fcStartBuild();
  const wb = workbook;
  const ric = window.requestIdleCallback || ((fn) => setTimeout(() => fn({ timeRemaining: () => 8 }), 60));
  const tick = (dl) => {
    if (!fcIndexBuild || fcIndexBuild.wb !== wb || workbook !== wb) return; // dokończone przez edycję albo nowy plik
    const budget = Math.max(4, Math.min(10, dl && dl.timeRemaining ? dl.timeRemaining() : 8));
    if (fcBuildStep(budget)) fcFinishBuild();
    else ric(tick, { timeout: 1500 });
  };
  ric(tick, { timeout: 1500 });
}

// Komórka dostała/straciła formułę (wpisanie, nadpisanie wartością, Cofnij) — popraw indeks.
function swbFormulaCellChanged(sheet, ref) {
  // budowa w toku: komórka, którą budowa już minęła, miałaby stary wpis — dokończ teraz
  if (fcIndexBuild && fcIndexBuild.wb === workbook) fcEnsureIndex();
  if (!fcIndex || fcIndex.wb !== workbook) return;
  const key = `${sheet}\u0000${ref}`;
  fcIndexRemove(key);
  const cell = workbook.Sheets[sheet] && workbook.Sheets[sheet][ref];
  if (cell && cell.f) fcIndexAdd(sheet, ref, cell);
}

function fcDependents(sheet, r, c) {
  const list = fcIndex && fcIndex.bySheet.get(sheet);
  if (!list) return [];
  const out = [];
  for (let i = 0; i < list.length; i++) {
    const d = list[i];
    if (r >= d.r1 && r <= d.r2 && c >= d.c1 && c <= d.c2) out.push(d.key);
  }
  return out;
}

// ── wynik formuły → komórka ─────────────────────────────────────────────────
function fcTypeOf(v) {
  if (FC.isErr(v)) return "e";
  if (typeof v === "boolean") return "b";
  if (typeof v === "number") return "n";
  return "s";
}

function fcFormatted(cell, v) {
  if (FC.isErr(v)) return v.err;
  if (typeof v === "boolean") return v ? "TRUE" : "FALSE";
  if (typeof v === "string") return v;
  const z = cell && cell.z && cell.z !== "General" ? cell.z : null;
  if (z) { const w = swbSsfFormat(z, v); if (w != null) return w; } // też „dd.mm.yyyy”, na którym SSF się wywraca
  return FC.generalText(v);
}

// Liczba seryjna → Date w czasie lokalnym (tak jak SheetJS z cellDates daje daty z pliku).
function fcSerialToDate(serial) {
  const ms = Math.round(serial * 86400000);
  const u = new Date(Date.UTC(1899, 11, 30) + ms);
  return new Date(u.getUTCFullYear(), u.getUTCMonth(), u.getUTCDate(), u.getUTCHours(), u.getUTCMinutes(), u.getUTCSeconds());
}
function fcIsDateFormat(z) {
  return !!z && z !== "General" && typeof isDateFormatCode === "function" && isDateFormatCode(z);
}

// Wpisuje wynik do obiektu komórki (bez ruszania formuły i stylu). true = wynik się zmienił.
// Komórka w formacie daty dostaje Date (jak daty z pliku) — filtry dat i sortowanie widzą datę.
function fcStore(cell, v) {
  let t = fcTypeOf(v);
  let nv = FC.isErr(v) ? v.err : v;
  if (t === "n" && fcIsDateFormat(cell.z) && v >= 0) { nv = fcSerialToDate(v); t = "d"; }
  const before = cell.v instanceof Date ? cell.v.getTime() : cell.v;
  const after = nv instanceof Date ? nv.getTime() : nv;
  const changed = before !== after || cell.t !== t;
  cell.v = nv;
  cell.t = t;
  cell.w = fcFormatted(cell, v);
  return changed;
}

// Odświeża wiersz tabeli (bieżący arkusz) po zmianie komórki r,c.
let fcRowMapCache = { rows: null, map: null };
function fcRowFor(r) {
  if (fcRowMapCache.rows !== baseRows) {
    const map = new Map();
    baseRows.forEach((row) => { if (row && !row.isLongViewRow && Number.isFinite(row.rowIndex0) && !map.has(row.rowIndex0)) map.set(row.rowIndex0, row); });
    fcRowMapCache = { rows: baseRows, map };
  }
  return fcRowMapCache.map.get(r) || null;
}
function fcRefreshRow(sheet, r, c) {
  if (sheet !== currentSheetName) return false;
  const row = fcRowFor(r);
  const col = c - currentStartCol;
  if (!row || col < 0 || !Array.isArray(row.values) || col >= row.values.length) return false;
  const cell = workbook.Sheets[sheet][XLSX.utils.encode_cell({ r, c })];
  const v = cell ? (cell.t === "e" ? null : (cell.v === undefined ? null : cell.v)) : null;
  row.values[col] = v;
  if (Array.isArray(row.rawValues) && row.rawValues !== row.values) row.rawValues[col] = v;
  if (Array.isArray(row.display)) {
    let shown = cell ? (cell.w != null ? String(cell.w) : (v == null ? "" : toDisplay(v))) : "";
    if (v instanceof Date && typeof localizeDisplayedDate === "function") shown = localizeDisplayedDate(v, shown, cell);
    if (cell && cell.t === "e") shown = cell.w || String(cell.v);
    shown = swbLocalizeCellText(shown, cell, currentLang);
    row.display[col] = shown;
  }
  if (row.recalcCells && row.recalcCells[col] !== undefined) delete row.recalcCells[col];
  return true;
}

// ── przeliczenie ────────────────────────────────────────────────────────────
let fcPending = null; // Set "arkusz\u0000r\u0000c"
let fcRecalcing = false;

// Wołane przy każdej zmianie wartości (updateSheetCell / formuła / Cofnij). Liczymy raz,
// po zakończeniu bieżącego zadania — wklejenie 500 komórek = jedno przeliczenie.
function swbFormulaTouched(sheet, r, c) {
  if (!FC || fcRecalcing) return;
  if (!fcPending) {
    fcPending = new Set();
    queueMicrotask(fcFlush);
  }
  fcPending.add(`${sheet}\u0000${r}\u0000${c}`);
}

function fcFlush() {
  const pending = fcPending;
  fcPending = null;
  if (!pending || !pending.size || !workbook) return;
  const idx = fcEnsureIndex();
  if (!idx || !idx.entries.size) return;
  const t0 = performance.now();
  // zbiór formuł do przeliczenia (BFS po zależnościach)
  const order = [];
  const seen = new Set();
  const queue = [];
  pending.forEach((k) => {
    const [sheet, r, c] = k.split("\u0000");
    fcDependents(sheet, +r, +c).forEach((key) => { if (!seen.has(key)) { seen.add(key); queue.push(key); } });
  });
  while (queue.length) {
    const key = queue.shift();
    order.push(key);
    const e = idx.entries.get(key);
    if (!e) continue;
    fcDependents(e.sheet, e.r, e.c).forEach((k2) => { if (!seen.has(k2)) { seen.add(k2); queue.push(k2); } });
    if (order.length > 200000) break; // bezpiecznik
  }
  if (!order.length) return;
  let renderNeeded = false;
  let changedCount = 0;
  fcRecalcing = true;
  // wspólna pamięć zakresów dla wszystkich formuł tego przeliczenia (zrzut Mateusza
  // 2026-10-25: Enter/Wklej w arkuszu z ~300 LICZ.JEŻELI po całych kolumnach = 2 s blokady,
  // bo każda formuła budowała od nowa ten sam zakres komórka po komórce)
  fcContext(currentSheetName, 0, 0);
  const matrixCache = FC.makeMatrixCache ? FC.makeMatrixCache() : null;
  fcBaseCtx.matrixCache = matrixCache;
  try {
    // kolejność BFS bywa nie-topologiczna (romby) — dwa przebiegi wystarczają w praktyce,
    // trzeci tylko gdy coś się jeszcze zmieniało
    for (let pass = 0; pass < 3; pass++) {
      let changedThisPass = false;
      for (const key of order) {
        const e = idx.entries.get(key);
        if (!e || !e.ok || !e.ast) continue;
        const cell = workbook.Sheets[e.sheet] && workbook.Sheets[e.sheet][e.ref];
        if (!cell || !cell.f) continue;
        let v;
        try { v = FC.evaluate(e.ast, fcContext(e.sheet, e.r, e.c)); } catch { continue; }
        if (v === FC.UNSUPPORTED) continue;
        if (fcStore(cell, v)) {
          if (matrixCache) matrixCache.invalidate(e.sheet, e.r, e.c);
          changedThisPass = true;
          changedCount++;
          if (fcRefreshRow(e.sheet, e.r, e.c)) renderNeeded = true;
        }
      }
      if (!changedThisPass) break;
    }
  } finally {
    fcRecalcing = false;
    if (fcBaseCtx) fcBaseCtx.matrixCache = null; // poza przeliczeniem komórki zmienia edycja — bez pamięci
  }
  if (changedCount) {
    bumpSheetDataStamp();
    if (renderNeeded) {
      renderActiveTable();
      if (typeof scheduleViewRefresh === "function") scheduleViewRefresh({ analyses: true });
    }
    if (typeof log === "function") log(`Przeliczono ${changedCount} formuł (${Math.round(performance.now() - t0)} ms)`, "info");
  }
}

// ── wpisanie formuły ────────────────────────────────────────────────────────
// text: to, co wpisał użytkownik (z „=”), w jego języku. Zwraca "ok" | "error" | "skip".
function swbFormulaCommit(row, col, text, opts = {}) {
  if (!row || row.isLongViewRow || !Number.isFinite(row.rowIndex0)) return "skip";
  const res = swbFormulaCommitAt(row.rowIndex0, col, text, opts);
  if (res === "ok" && Array.isArray(row.display) && fcLastUnknown) row.display[col] = t("formulaExcelWillCalc");
  return res;
}

// To samo bez obiektu wiersza — np. suma pod danymi, w wierszu, którego tabela jeszcze nie ma.
let fcLastUnknown = false;
function swbFormulaCommitAt(rowIndex0, col, text, { silent = false } = {}) {
  fcLastUnknown = false;
  if (!FC || !workbook || !currentSheetName) return "skip";
  if (!Number.isFinite(rowIndex0)) return "skip";
  if (typeof cfmtIsRealCol === "function" && !cfmtIsRealCol(col)) return "skip";
  const lang = typeof currentLang === "string" ? currentLang : "pl";
  let ast;
  try {
    ast = FC.parse(text, { lang });
  } catch (e) {
    if (!silent) toast(t("formulaSyntax", { msg: e.message }), "warning");
    return "error";
  }
  const fileF = FC.serialize(ast, { file: true });
  const r = rowIndex0;
  const c = currentStartCol + col;
  const ref = XLSX.utils.encode_cell({ r, c });
  // odwołanie do samej siebie (bez pośrednika) — Excel: odwołanie cykliczne
  const selfRef = FC.refsOf(ast).some((d) => !d.table && !d.name && (d.sheet == null || fcSheetFor(d.sheet) === currentSheetName)
    && r >= Math.min(d.r1, d.r2) && r <= Math.max(d.r1, d.r2) && c >= Math.min(d.c1, d.c2) && c <= Math.max(d.c1, d.c2));
  if (selfRef) {
    if (!silent) toast(t("formulaCircular", { ref }), "warning");
    return "error";
  }
  const fns = [...FC.functionsOf(ast)];
  const unknown = fns.filter((n) => !FC.isSupported(n));
  let value = null;
  if (!unknown.length) {
    try { value = FC.evaluate(ast, fcContext(currentSheetName, r, c)); } catch { value = null; }
    if (value === FC.UNSUPPORTED) value = null;
  }
  updateSheetFormulaCell(r, col, fileF, value);
  // Excel przy formacie Ogólnym sam daje formatu daty wynikom DZIŚ(), DATA()…
  const cell = workbook.Sheets[currentSheetName][ref];
  if (value != null && typeof value === "number" && FC.resultLooksLikeDate(ast) && (!cell.z || cell.z === "General")
    && typeof cfmtRecord === "function") {
    cfmtRecord(currentSheetName, ref, { numFmt: "yyyy-mm-dd" });
    cell.z = "yyyy-mm-dd";
    cell.w = fcFormatted(cell, value);
  }
  fcRefreshRow(currentSheetName, r, c);
  fcLastUnknown = unknown.length > 0;
  if (unknown.length && !silent) toast(t("formulaUnknownFn", { fns: unknown.map((n) => FC.PL_NAMES[n] && lang === "pl" ? FC.PL_NAMES[n] : n).join(", ") }), "info");
  return "ok";
}

// Formuła z komórki do edycji (z „=”, w języku aplikacji). null = komórka bez formuły.
function swbFormulaText(sheet, ref) {
  const cell = workbook && workbook.Sheets[sheet] && workbook.Sheets[sheet][ref];
  if (!cell || !cell.f) return null;
  const lang = typeof currentLang === "string" ? currentLang : "pl";
  return "=" + (FC ? FC.localize(cell.f, lang) : cell.f);
}

// Formuła przesunięta o (dr, dc) — wypełnianie w dół/w prawo. Tekst w języku aplikacji.
function swbFormulaShifted(sheet, ref, dr, dc) {
  const cell = workbook && workbook.Sheets[sheet] && workbook.Sheets[sheet][ref];
  if (!cell || !cell.f || !FC) return null;
  try {
    const lang = typeof currentLang === "string" ? currentLang : "pl";
    return "=" + FC.serialize(FC.shift(FC.parse(cell.f, { lang: "en" }), dr, dc), { lang });
  } catch { return null; }
}

// ── Autosumowanie (Σ) ───────────────────────────────────────────────────────
// Jak w Excelu: na pustej komórce szuka ciągłego bloku liczb NAD nią (w arkuszu, nie
// w widoku po filtrze), a gdy go nie ma — po lewej. Zwraca tekst formuły do edytora
// (użytkownik zatwierdza Enterem) albo null.
function swbAutoSumFormula(fnName, rowIndex0, absCol) {
  if (!workbook || !currentSheetName) return null;
  const ws = workbook.Sheets[currentSheetName];
  const isNum = (r, c) => { const v = fcCellValue(ws[XLSX.utils.encode_cell({ r, c })]); return typeof v === "number"; };
  const lang = typeof currentLang === "string" ? currentLang : "pl";
  const name = lang === "pl" ? (FC.PL_NAMES[fnName] || fnName) : fnName;
  let r = rowIndex0 - 1;
  while (r >= 0 && !isNum(r, absCol) && r >= rowIndex0 - 1) r--; // jedna pusta przerwa tuż nad — nie
  r = rowIndex0 - 1;
  if (r >= 0 && isNum(r, absCol)) {
    let top = r;
    while (top - 1 >= 0 && isNum(top - 1, absCol)) top--;
    const a = XLSX.utils.encode_cell({ r: top, c: absCol });
    const b = XLSX.utils.encode_cell({ r, c: absCol });
    return `=${name}(${a}:${b})`;
  }
  let c = absCol - 1;
  if (c >= 0 && isNum(rowIndex0, c)) {
    let left = c;
    while (left - 1 >= 0 && isNum(rowIndex0, left - 1)) left--;
    const a = XLSX.utils.encode_cell({ r: rowIndex0, c: left });
    const b = XLSX.utils.encode_cell({ r: rowIndex0, c });
    return `=${name}(${a}:${b})`;
  }
  return `=${name}()`;
}
