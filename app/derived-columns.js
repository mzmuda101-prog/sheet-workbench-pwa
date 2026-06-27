// app/derived-columns.js
// Kolumny wyliczane (#6) + cross-sheet lookup / zamiennik VLOOKUP (#7) — JEDEN silnik.
//
// Idea: zamiast budować osobny moduł join + osobny moduł formuł, dokładamy lekki
// silnik wyrażeń na referencjach KOLUMN `[Nagłówek]` (zamiast A1) i REUŻYWAMY całą
// bibliotekę funkcji z conditional-formatting.js (cfFunc: IF, SUM, ROUND, DATEDIF,
// LEFT/RIGHT/MID, YEAR/MONTH… oraz cfNum/cfStr/cfCompare/cfBool/cfIsErr/CF_ERR).
// Cross-sheet lookup to po prostu funkcja WYSZUKAJ()/LOOKUP() w tym samym silniku.
//
// Integracja: wyliczone kolumny są DOPISYWANE na koniec `currentHeaders` oraz do
// `row.values`/`row.display` każdego wiersza w `baseRows`. Dzięki temu sortowanie,
// filtry, picker kolumn, walidacja, szybkie szukanie i eksport widzą je jak zwykłe
// kolumny — zero dublowania logiki. Kolumny są WIRTUALNE: nie trafiają do pliku xlsx.
//
// Bezpieczeństwo wyrównania: żadna inna funkcja nie zmienia liczby kolumn po
// wczytaniu (edit-tools edytuje tylko wartości), więc usuwanie sprowadza się do
// odcięcia ostatnich N kolumn (dcAppliedCount).

/* ============================ stan ============================ */

const DERIVED_COLUMNS_KEY = "excel-workbench-derived-columns";
const DC_NA = { __cfError: true, __na: true }; // "nie znaleziono" (#N/D) — odróżniamy od #BŁĄD

let derivedColumns = [];        // [{ id, name, expr }]
let dcAppliedCount = 0;         // ile kolumn aktualnie dopisanych do currentHeaders/baseRows
let dcIdSeq = 1;
let dcSheetTableCache = new Map(); // (sheet\0headerRow) -> { headers, rows } innego arkusza
let dcKeyMapCache = new Map();     // (sheet\0headerRow\0keyIdx) -> Map(normKey -> rowIndex)

function dcNextId() { return dcIdSeq++; }
function dcClearCaches() { dcSheetTableCache.clear(); dcKeyMapCache.clear(); }

/* ============================ tokenizer / parser ============================ */
// Jak cfTokenize, ale rozpoznaje `[Nagłówek]` jako referencję kolumny, akceptuje
// `;` ORAZ `,` jako separator argumentów (PL Excel używa `;`), oraz polskie litery
// w nazwach funkcji (JEŻELI…). Stringi w "…" albo '…'.
const DC_BINPREC = { "=": 1, "<>": 1, "<": 1, ">": 1, "<=": 1, ">=": 1, "&": 2, "+": 3, "-": 3, "*": 4, "/": 4, "^": 5 };

function dcTokenize(s) {
  const toks = [];
  const re = /\s*(?:\[([^\]]+)\]|(\d+(?:\.\d+)?(?:[eE][+-]?\d+)?)|"((?:[^"]|"")*)"|'((?:[^']|'')*)'|(<=|>=|<>|[-+*/^&=<>(),;])|([A-Za-z_ąćęłńóśźżĄĆĘŁŃÓŚŹŻ][A-Za-z0-9_.ąćęłńóśźżĄĆĘŁŃÓŚŹŻ]*))/g;
  let m;
  while ((m = re.exec(s)) !== null) {
    if (m[0].trim() === "") { if (re.lastIndex === m.index) re.lastIndex++; continue; }
    if (m[1] !== undefined) toks.push({ t: "col", v: m[1].trim() });
    else if (m[2] !== undefined) toks.push({ t: "num", v: parseFloat(m[2]) });
    else if (m[3] !== undefined) toks.push({ t: "str", v: m[3].replace(/""/g, '"') });
    else if (m[4] !== undefined) toks.push({ t: "str", v: m[4].replace(/''/g, "'") });
    else if (m[5] !== undefined) { let op = m[5]; if (op === ";") op = ","; toks.push({ t: "op", v: op }); }
    else if (m[6] !== undefined) {
      const up = m[6].toUpperCase();
      if (up === "TRUE" || up === "PRAWDA") toks.push({ t: "bool", v: true });
      else if (up === "FALSE" || up === "FAŁSZ" || up === "FALSZ") toks.push({ t: "bool", v: false });
      else toks.push({ t: "name", v: up });
    }
  }
  return toks;
}

function dcParse(tokens) {
  let pos = 0;
  const peek = () => tokens[pos];
  const next = () => tokens[pos++];
  function parseExpr(minPrec) {
    let left = parseUnary();
    while (true) {
      const tk = peek();
      if (!tk || tk.t !== "op") break;
      const prec = DC_BINPREC[tk.v];
      if (prec == null || prec < minPrec) break;
      next();
      const right = parseExpr(tk.v === "^" ? prec : prec + 1);
      left = { k: "bin", op: tk.v, a: left, b: right };
    }
    return left;
  }
  function parseUnary() {
    const tk = peek();
    if (tk && tk.t === "op" && (tk.v === "-" || tk.v === "+")) { next(); return { k: "un", op: tk.v, a: parseUnary() }; }
    return parsePrimary();
  }
  function parsePrimary() {
    const tk = next();
    if (!tk) throw 0;
    if (tk.t === "num") return { k: "num", v: tk.v };
    if (tk.t === "str") return { k: "str", v: tk.v };
    if (tk.t === "bool") return { k: "bool", v: tk.v };
    if (tk.t === "col") return { k: "col", v: tk.v };
    if (tk.t === "name") {
      if (peek() && peek().t === "op" && peek().v === "(") {
        next();
        const args = [];
        if (!(peek() && peek().t === "op" && peek().v === ")")) {
          args.push(parseExpr(0));
          while (peek() && peek().t === "op" && peek().v === ",") { next(); args.push(parseExpr(0)); }
        }
        const cl = next();
        if (!cl || cl.v !== ")") throw 0;
        return { k: "call", name: tk.v, args };
      }
      throw 0; // gołe słowo bez ( ) — błąd (kolumny piszemy w [ ])
    }
    if (tk.t === "op" && tk.v === "(") { const e = parseExpr(0); const c = next(); if (!c || c.v !== ")") throw 0; return e; }
    throw 0;
  }
  const ast = parseExpr(0);
  if (pos < tokens.length) throw 0;
  return ast;
}

/* ============================ ewaluator ============================ */
// Date → serial Excela, żeby arytmetyka dat działała spójnie z cfFunc (które operuje
// na serialach). Liczby zostają liczbami, stringi stringami.
function dcCoerceArg(v) {
  if (v instanceof Date) return cfDateToSerial(v);
  return v;
}
// Liczba „po ludzku": natywna liczba, data, albo string z polskim formatem (1 000,50).
function dcNum(v) {
  if (typeof v === "number") return v;
  if (v instanceof Date) return cfDateToSerial(v);
  const p = (typeof parseCellNumber === "function") ? parseCellNumber(v) : null;
  if (p != null) return p;
  return cfNum(v);
}

function dcEval(node, ctx) {
  if (!node) return CF_ERR;
  switch (node.k) {
    case "num": case "str": case "bool": return node.v;
    case "col": {
      const idx = ctx.colMap.get(node.v.trim().toLowerCase());
      if (idx === undefined) return CF_ERR; // nieznana kolumna
      const raw = ctx.row && Array.isArray(ctx.row.values) ? ctx.row.values[idx] : null;
      return dcCoerceArg(raw == null ? null : raw);
    }
    case "un": {
      const a = dcEval(node.a, ctx); if (cfIsErr(a)) return a;
      const n = dcNum(a); return isNaN(n) ? CF_ERR : (node.op === "-" ? -n : n);
    }
    case "bin": return dcEvalBin(node, ctx);
    case "call": return dcEvalCall(node, ctx);
  }
  return CF_ERR;
}

function dcEvalBin(node, ctx) {
  const a = dcEval(node.a, ctx); if (cfIsErr(a)) return a;
  const b = dcEval(node.b, ctx); if (cfIsErr(b)) return b;
  const op = node.op;
  if (op === "&") return cfStr(a) + cfStr(b);
  if (op === "=" || op === "<>" || op === "<" || op === ">" || op === "<=" || op === ">=") return cfCompare(op, a, b);
  const x = dcNum(a), y = dcNum(b);
  if (isNaN(x) || isNaN(y)) return CF_ERR;
  if (op === "+") return x + y;
  if (op === "-") return x - y;
  if (op === "*") return x * y;
  if (op === "/") return y === 0 ? CF_ERR : x / y;
  if (op === "^") return Math.pow(x, y);
  return CF_ERR;
}

function dcEvalCall(node, ctx) {
  const name = node.name, A = node.args;
  // Funkcje sterujące (leniwa/tolerancyjna ewaluacja) — własne, z aliasami PL.
  if (name === "IF" || name === "JEŻELI" || name === "JEZELI") {
    const c = cfBool(dcEval(A[0], ctx)); if (cfIsErr(c)) return c;
    return c ? dcEval(A[1], ctx) : (A[2] !== undefined ? dcEval(A[2], ctx) : false);
  }
  if (name === "AND" || name === "ORAZ" || name === "I") {
    for (const a of A) { const v = cfBool(dcEval(a, ctx)); if (cfIsErr(v)) return v; if (!v) return false; } return true;
  }
  if (name === "OR" || name === "LUB") {
    for (const a of A) { const v = cfBool(dcEval(a, ctx)); if (cfIsErr(v)) return v; if (v) return true; } return false;
  }
  if (name === "NOT" || name === "NIE") { const v = cfBool(dcEval(A[0], ctx)); if (cfIsErr(v)) return v; return !v; }
  if (name === "IFERROR" || name === "JEŻELIBŁĄD" || name === "JEZELIBLAD") {
    const v = dcEval(A[0], ctx); return cfIsErr(v) ? dcEval(A[1], ctx) : v;
  }
  if (name === "LOOKUP" || name === "WYSZUKAJ" || name === "VLOOKUP" || name === "WYSZUKAJ.PIONOWO") {
    return dcEvalLookup(A, ctx);
  }
  // Reszta: policz argumenty i oddaj wspólnej bibliotece cfFunc (jedno źródło prawdy).
  const args = A.map((a) => dcEval(a, ctx));
  const tolerant = name === "ISERROR" || name === "ISERR" || name === "ISNA";
  if (!tolerant) for (const a of args) if (cfIsErr(a)) return a;
  return cfFunc(name, args);
}

/* ============================ cross-sheet lookup (#7) ============================ */

function dcEvalLookup(A, ctx) {
  if (!A || A.length < 4) return CF_ERR;
  const key = dcEval(A[0], ctx); if (cfIsErr(key)) return key;
  const sheetName = cfStr(dcEval(A[1], ctx));
  const keyCol = dcEval(A[2], ctx);
  const retCol = dcEval(A[3], ctx);
  let headerRow = 1;
  if (A[4] !== undefined) { const hr = Math.round(dcNum(dcEval(A[4], ctx))); if (Number.isFinite(hr) && hr >= 1) headerRow = hr; }
  return dcLookup(key, sheetName, keyCol, retCol, headerRow);
}

function dcGetSheetTable(sheetName, headerRow) {
  const cacheKey = sheetName + " " + headerRow;
  if (dcSheetTableCache.has(cacheKey)) return dcSheetTableCache.get(cacheKey);
  let table = null;
  try {
    const sheet = (typeof workbook !== "undefined" && workbook && workbook.Sheets) ? workbook.Sheets[sheetName] : null;
    if (sheet) {
      const aoa = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: null, blankrows: false });
      const hdr = aoa[headerRow - 1] || [];
      const headers = hdr.map((h, i) => (h == null || String(h).trim() === "") ? XLSX.utils.encode_col(i) : String(h).trim());
      const rows = aoa.slice(headerRow);
      table = { headers, rows };
    }
  } catch { table = null; }
  dcSheetTableCache.set(cacheKey, table);
  return table;
}

function dcResolveColIndex(headers, colRef) {
  if (typeof colRef === "number" && Number.isFinite(colRef)) {
    const i = Math.round(colRef) - 1; // numer = 1-based
    return (i >= 0 && i < headers.length) ? i : -1;
  }
  const s = String(colRef == null ? "" : colRef).trim();
  if (!s) return -1;
  const norm = s.toLowerCase();
  let idx = headers.findIndex((h) => String(h).trim().toLowerCase() === norm);
  if (idx >= 0) return idx;
  if (/^[A-Za-z]{1,3}$/.test(s)) {
    try { const c = XLSX.utils.decode_col(s.toUpperCase()); if (c >= 0 && c < headers.length) return c; } catch { /* noop */ }
  }
  return -1;
}

// Klucz porównujemy „po ludzku": liczby (i liczbowe stringi) numerycznie, tekst
// bez wielkości liter i nadmiarowych spacji — 100 == "100", "Jan " == "jan".
function dcNormKey(v) {
  if (v == null) return "";
  if (typeof v === "number") return "n:" + v;
  if (v instanceof Date) return "n:" + cfDateToSerial(v);
  const s = String(v).trim();
  if (s === "") return "";
  if (/^[-+]?[\d\s.,]+$/.test(s)) {
    const num = (typeof parseCellNumber === "function") ? parseCellNumber(s) : Number(s);
    if (num != null && Number.isFinite(num)) return "n:" + num;
  }
  return "s:" + s.toLowerCase().replace(/\s+/g, " ");
}

function dcGetKeyMap(sheetName, headerRow, keyIdx) {
  const cacheKey = sheetName + " " + headerRow + " " + keyIdx;
  if (dcKeyMapCache.has(cacheKey)) return dcKeyMapCache.get(cacheKey);
  const table = dcGetSheetTable(sheetName, headerRow);
  const map = new Map();
  if (table) {
    for (let r = 0; r < table.rows.length; r++) {
      const k = dcNormKey((table.rows[r] || [])[keyIdx]);
      if (k !== "" && !map.has(k)) map.set(k, r); // pierwsze trafienie wygrywa (jak VLOOKUP)
    }
  }
  dcKeyMapCache.set(cacheKey, map);
  return map;
}

function dcLookup(key, sheetName, keyCol, retCol, headerRow) {
  const table = dcGetSheetTable(sheetName, headerRow);
  if (!table) return CF_ERR; // brak arkusza
  const keyIdx = dcResolveColIndex(table.headers, keyCol);
  const retIdx = dcResolveColIndex(table.headers, retCol);
  if (keyIdx < 0 || retIdx < 0) return CF_ERR; // zła nazwa kolumny
  const map = dcGetKeyMap(sheetName, headerRow, keyIdx);
  const r = map.get(dcNormKey(key));
  if (r === undefined) return DC_NA; // nie znaleziono → #N/D
  const val = (table.rows[r] || [])[retIdx];
  return val == null ? null : val;
}

/* ============================ wartość → komórka ============================ */

function dcFormatNumber(n) {
  const loc = (typeof I18N !== "undefined" && I18N[currentLang] && I18N[currentLang].locale) || "pl-PL";
  try { return new Intl.NumberFormat(loc, { maximumFractionDigits: 6 }).format(n); } catch { return String(n); }
}
function dcDisplayValue(v) {
  if (v && v.__na) return t("derivedNa");
  if (cfIsErr(v)) return t("derivedErr");
  if (v === null || v === undefined) return "";
  if (typeof v === "boolean") return v ? "TRUE" : "FALSE";
  if (typeof v === "number") return Number.isFinite(v) ? dcFormatNumber(v) : t("derivedErr");
  if (v instanceof Date) return toDisplay(v);
  return String(v);
}
// Co trafia do row.values (sort/agregacja/filtr): błędy → null (puste), reszta jak jest.
function dcStorableValue(v) {
  if (v && v.__cfError) return null;
  return v;
}

/* ============================ apply / remove ============================ */

function dcBuildColMap(headers) {
  const map = new Map();
  headers.forEach((h, i) => {
    const k = String(h == null ? "" : h).trim().toLowerCase();
    if (k && !map.has(k)) map.set(k, i);
  });
  return map;
}
function dcUniqueHeaderName(name, headers) {
  let base = String(name == null ? "" : name).trim() || t("derivedDefaultName");
  if (!headers.includes(base)) return base;
  let i = 2;
  while (headers.includes(`${base} (${i})`)) i++;
  return `${base} (${i})`;
}

// Odetnij dopisane wcześniej kolumny wyliczane (ostatnie dcAppliedCount).
function removeDerivedColumns() {
  const n = dcAppliedCount;
  if (n <= 0) return;
  if (typeof currentHeaders !== "undefined" && Array.isArray(currentHeaders) && currentHeaders.length >= n) {
    currentHeaders.splice(currentHeaders.length - n, n);
  }
  if (typeof currentHeaderStyles !== "undefined" && Array.isArray(currentHeaderStyles) && currentHeaderStyles.length >= n) {
    currentHeaderStyles.splice(currentHeaderStyles.length - n, n);
  }
  if (typeof currentSheetColWidths !== "undefined" && Array.isArray(currentSheetColWidths) && currentSheetColWidths.length >= n) {
    currentSheetColWidths.splice(currentSheetColWidths.length - n, n);
  }
  if (typeof baseRows !== "undefined" && Array.isArray(baseRows)) {
    baseRows.forEach((row) => {
      if (Array.isArray(row.values) && row.values.length >= n) row.values.splice(row.values.length - n, n);
      // rawValues to zwykle TA SAMA referencja co values (buildRows) — wtedy już obcięte.
      if (Array.isArray(row.rawValues) && row.rawValues !== row.values && row.rawValues.length >= n) row.rawValues.splice(row.rawValues.length - n, n);
      if (Array.isArray(row.display) && row.display.length >= n) row.display.splice(row.display.length - n, n);
      if (Array.isArray(row.cellStyles) && row.cellStyles.length >= n) row.cellStyles.splice(row.cellStyles.length - n, n);
    });
  }
  dcAppliedCount = 0;
}

// Przelicz i wstrzyknij wszystkie kolumny wyliczane do currentHeaders + baseRows.
function applyDerivedColumns(opts = {}) {
  if (typeof currentHeaders === "undefined" || !Array.isArray(currentHeaders)) return;
  removeDerivedColumns();
  dcClearCaches();
  const hasSheet = currentHeaders.length > 0 && Array.isArray(baseRows);
  let applied = 0;
  if (hasSheet) {
    derivedColumns.forEach((def) => {
      if (!def || !def.expr || !String(def.expr).trim()) { def._error = "empty"; return; }
      let ast = null; let parseErr = false;
      try { ast = dcParse(dcTokenize(def.expr)); } catch { parseErr = true; }
      def._error = parseErr ? "parse" : null;
      const name = dcUniqueHeaderName(def.name, currentHeaders);
      def._appliedName = name;
      currentHeaders.push(name);
      if (typeof currentHeaderStyles !== "undefined" && Array.isArray(currentHeaderStyles)) currentHeaderStyles.push(null);
      if (typeof currentSheetColWidths !== "undefined" && Array.isArray(currentSheetColWidths)) currentSheetColWidths.push(null);
      const colMap = dcBuildColMap(currentHeaders); // zawiera kolumny bazowe + wcześniejsze wyliczane
      let errRows = 0;
      baseRows.forEach((row) => {
        let val = parseErr ? CF_ERR : CF_ERR;
        if (!parseErr) { try { val = dcEval(ast, { row, colMap }); } catch { val = CF_ERR; } }
        if (val && val.__cfError) errRows++;
        if (Array.isArray(row.values)) row.values.push(dcStorableValue(val));
        if (Array.isArray(row.display)) row.display.push(dcDisplayValue(val));
        if (Array.isArray(row.cellStyles)) row.cellStyles.push(null);
      });
      def._errRows = errRows;
      applied++;
    });
  }
  dcAppliedCount = applied;
  dcPersist();
  if (!opts.silent) dcRefreshDependents();
  dcRenderPanel();
}

// Po zmianie zbioru kolumn: odśwież selecty kolumn i przerysuj tabelę/analizy.
function dcRefreshDependents() {
  if (typeof populateSortColumnSelect === "function") populateSortColumnSelect();
  if (typeof populateValidationColumns === "function") populateValidationColumns();
  if (typeof populateEditColumnSelect === "function") populateEditColumnSelect();
  if (typeof updateColumnSummary === "function") updateColumnSummary();
  if (typeof applyFilters === "function") applyFilters();
  if (typeof sortRows === "function") sortRows();
  if (typeof renderActiveTable === "function") renderActiveTable();
  if (typeof renderColumnProfiles === "function") renderColumnProfiles();
}

// Wczytano nowy arkusz: baseRows są świeże (nic nie dopisane) → wyzeruj licznik,
// wyczyść cache lookupów. Wołane PRZED applyDerivedColumns({silent:true}) w load path.
function dcResetForNewSheet() {
  dcAppliedCount = 0;
  dcClearCaches();
}

/* ============================ trwałość ============================ */

function dcPersist() {
  try {
    localStorage.setItem(DERIVED_COLUMNS_KEY, JSON.stringify(derivedColumns.map((d) => ({ name: d.name, expr: d.expr }))));
  } catch { /* prywatny tryb / brak miejsca — trudno */ }
}
function dcLoadPersisted() {
  try {
    const raw = localStorage.getItem(DERIVED_COLUMNS_KEY);
    if (!raw) return;
    const arr = JSON.parse(raw);
    if (Array.isArray(arr)) {
      derivedColumns = arr.filter((d) => d && d.expr).map((d) => ({ id: dcNextId(), name: String(d.name || "").trim(), expr: String(d.expr) }));
    }
  } catch { /* uszkodzony wpis — pomiń */ }
}

/* ============================ operacje publiczne (panel) ============================ */

function dcAddOrUpdate(name, expr, editId) {
  const cleanExpr = String(expr || "").trim();
  if (!cleanExpr) { toast(t("derivedNeedExpr"), "warning"); return false; }
  // szybka walidacja składni — zanim dopiszemy
  try { dcParse(dcTokenize(cleanExpr)); }
  catch { toast(t("derivedBadSyntax"), "error"); return false; }
  const cleanName = String(name || "").trim() || t("derivedDefaultName");
  if (editId != null) {
    const def = derivedColumns.find((d) => d.id === editId);
    if (def) { def.name = cleanName; def.expr = cleanExpr; }
  } else {
    derivedColumns.push({ id: dcNextId(), name: cleanName, expr: cleanExpr });
  }
  applyDerivedColumns();
  return true;
}
function dcRemove(id) {
  derivedColumns = derivedColumns.filter((d) => d.id !== id);
  applyDerivedColumns();
}
function dcClearAll() {
  if (!derivedColumns.length) return;
  derivedColumns = [];
  applyDerivedColumns();
}
function dcRecalc() { applyDerivedColumns(); }

// Buduje wyrażenie WYSZUKAJ z kreatora „Dołącz z arkusza" (#7).
function dcBuildLookupExpr(localKey, sheetName, srcKey, srcRet, headerRow) {
  const q = (s) => `"${String(s).replace(/"/g, '""')}"`;
  const hr = (headerRow && headerRow > 1) ? `; ${headerRow}` : "";
  return `WYSZUKAJ([${localKey}]; ${q(sheetName)}; ${q(srcKey)}; ${q(srcRet)}${hr})`;
}

/* ============================ panel UI ============================ */

const dcNameEl = document.getElementById("dcName");
const dcExprEl = document.getElementById("dcExpr");
const dcAddBtn = document.getElementById("dcAddBtn");
const dcFormulaMsgEl = document.getElementById("dcFormulaMsg");
const dcListEl = document.getElementById("dcList");
const dcRecalcBtn = document.getElementById("dcRecalcBtn");
const dcClearAllBtn = document.getElementById("dcClearAllBtn");
// kreator join (#7)
const dcJoinNameEl = document.getElementById("dcJoinName");
const dcJoinKeyColEl = document.getElementById("dcJoinKeyCol");
const dcJoinSheetEl = document.getElementById("dcJoinSheet");
const dcJoinSheetKeyColEl = document.getElementById("dcJoinSheetKeyCol");
const dcJoinSheetRetColEl = document.getElementById("dcJoinSheetRetCol");
const dcJoinBtn = document.getElementById("dcJoinBtn");

let dcEditId = null; // id edytowanej kolumny (null = dodawanie nowej)

function dcBaseHeaders() {
  // currentHeaders bez dopisanych wyliczanych (ostatnie dcAppliedCount)
  if (typeof currentHeaders === "undefined" || !Array.isArray(currentHeaders)) return [];
  return dcAppliedCount > 0 ? currentHeaders.slice(0, currentHeaders.length - dcAppliedCount) : currentHeaders.slice();
}

function dcFillSelect(sel, headers, withFallback) {
  if (!sel) return;
  const prev = sel.value;
  sel.replaceChildren();
  headers.forEach((h, i) => {
    const opt = document.createElement("option");
    opt.value = h;
    opt.textContent = (typeof exportColLabel === "function") ? exportColLabel(h, i) : (h || `#${i + 1}`);
    sel.appendChild(opt);
  });
  if (prev && Array.from(sel.options).some((o) => o.value === prev)) sel.value = prev;
}

function dcPopulateJoinSheets() {
  if (!dcJoinSheetEl) return;
  const prev = dcJoinSheetEl.value;
  dcJoinSheetEl.replaceChildren();
  const names = (typeof workbook !== "undefined" && workbook && Array.isArray(workbook.SheetNames)) ? workbook.SheetNames : [];
  names.forEach((nm) => {
    const opt = document.createElement("option");
    opt.value = nm; opt.textContent = nm;
    dcJoinSheetEl.appendChild(opt);
  });
  if (prev && names.includes(prev)) dcJoinSheetEl.value = prev;
  dcPopulateJoinSourceColumns();
}
function dcPopulateJoinSourceColumns() {
  if (!dcJoinSheetEl) return;
  const table = dcGetSheetTable(dcJoinSheetEl.value, 1);
  const headers = table ? table.headers : [];
  dcFillSelect(dcJoinSheetKeyColEl, headers);
  dcFillSelect(dcJoinSheetRetColEl, headers);
}

function dcRenderPanel() {
  // selecty kreatora join opierają się na kolumnach BAZOWYCH (nie wyliczanych)
  dcFillSelect(dcJoinKeyColEl, dcBaseHeaders());
  dcPopulateJoinSheets();

  if (dcListEl) {
    dcListEl.replaceChildren();
    if (!derivedColumns.length) {
      const empty = document.createElement("div");
      empty.className = "dc-empty";
      empty.textContent = t("derivedEmpty");
      dcListEl.appendChild(empty);
    } else {
      derivedColumns.forEach((def) => {
        const item = document.createElement("div");
        item.className = "dc-item";
        const main = document.createElement("div");
        main.className = "dc-item-main";
        const nameEl = document.createElement("span");
        nameEl.className = "dc-item-name";
        nameEl.textContent = def._appliedName || def.name || t("derivedDefaultName");
        const exprEl = document.createElement("code");
        exprEl.className = "dc-item-expr";
        exprEl.textContent = def.expr;
        main.append(nameEl, exprEl);
        if (def._error === "parse") {
          const badge = document.createElement("span");
          badge.className = "dc-item-badge bad";
          badge.textContent = t("derivedBadgeSyntax");
          main.appendChild(badge);
        } else if (def._errRows > 0) {
          const badge = document.createElement("span");
          badge.className = "dc-item-badge warn";
          badge.textContent = t("derivedBadgeErrRows", { n: def._errRows });
          main.appendChild(badge);
        }
        const actions = document.createElement("div");
        actions.className = "dc-item-actions";
        const editBtn = document.createElement("button");
        editBtn.type = "button"; editBtn.className = "btn btn-xs ghost";
        editBtn.textContent = t("derivedEdit");
        editBtn.addEventListener("click", () => dcStartEdit(def.id));
        const delBtn = document.createElement("button");
        delBtn.type = "button"; delBtn.className = "btn btn-xs ghost danger";
        delBtn.textContent = t("derivedRemove");
        delBtn.addEventListener("click", () => dcRemove(def.id));
        actions.append(editBtn, delBtn);
        item.append(main, actions);
        dcListEl.appendChild(item);
      });
    }
  }
  if (dcClearAllBtn) dcClearAllBtn.classList.toggle("hidden", !derivedColumns.length);
  if (dcRecalcBtn) dcRecalcBtn.classList.toggle("hidden", !derivedColumns.length);
}

function dcStartEdit(id) {
  const def = derivedColumns.find((d) => d.id === id);
  if (!def) return;
  dcEditId = id;
  if (dcNameEl) dcNameEl.value = def.name;
  if (dcExprEl) dcExprEl.value = def.expr;
  if (dcAddBtn) dcAddBtn.textContent = t("derivedUpdate");
  if (dcFormulaMsgEl) dcFormulaMsgEl.textContent = "";
  if (dcExprEl) dcExprEl.focus();
}
function dcResetForm() {
  dcEditId = null;
  if (dcNameEl) dcNameEl.value = "";
  if (dcExprEl) dcExprEl.value = "";
  if (dcAddBtn) dcAddBtn.textContent = t("derivedAdd");
  if (dcFormulaMsgEl) dcFormulaMsgEl.textContent = "";
}

/* ============================ wiring ============================ */

if (dcAddBtn) {
  dcAddBtn.addEventListener("click", () => {
    if (typeof currentHeaders === "undefined" || !currentHeaders.length) { toast(t("loadSheetToPickColumns"), "info"); return; }
    const ok = dcAddOrUpdate(dcNameEl ? dcNameEl.value : "", dcExprEl ? dcExprEl.value : "", dcEditId);
    if (ok) dcResetForm();
  });
}
if (dcExprEl) {
  dcExprEl.addEventListener("input", () => {
    if (!dcFormulaMsgEl) return;
    const v = dcExprEl.value.trim();
    if (!v) { dcFormulaMsgEl.textContent = ""; dcFormulaMsgEl.className = "dc-msg"; return; }
    try { dcParse(dcTokenize(v)); dcFormulaMsgEl.textContent = t("derivedSyntaxOk"); dcFormulaMsgEl.className = "dc-msg ok"; }
    catch { dcFormulaMsgEl.textContent = t("derivedSyntaxBad"); dcFormulaMsgEl.className = "dc-msg bad"; }
  });
}
if (dcRecalcBtn) dcRecalcBtn.addEventListener("click", dcRecalc);
if (dcClearAllBtn) dcClearAllBtn.addEventListener("click", dcClearAll);
if (dcJoinSheetEl) dcJoinSheetEl.addEventListener("change", dcPopulateJoinSourceColumns);
if (dcJoinBtn) {
  dcJoinBtn.addEventListener("click", () => {
    if (typeof currentHeaders === "undefined" || !currentHeaders.length) { toast(t("loadSheetToPickColumns"), "info"); return; }
    const localKey = dcJoinKeyColEl ? dcJoinKeyColEl.value : "";
    const sheetName = dcJoinSheetEl ? dcJoinSheetEl.value : "";
    const srcKey = dcJoinSheetKeyColEl ? dcJoinSheetKeyColEl.value : "";
    const srcRet = dcJoinSheetRetColEl ? dcJoinSheetRetColEl.value : "";
    if (!localKey || !sheetName || !srcKey || !srcRet) { toast(t("derivedJoinNeedAll"), "warning"); return; }
    const name = (dcJoinNameEl && dcJoinNameEl.value.trim()) || srcRet;
    const expr = dcBuildLookupExpr(localKey, sheetName, srcKey, srcRet, 1);
    if (dcAddOrUpdate(name, expr, null) && dcJoinNameEl) dcJoinNameEl.value = "";
  });
}

// start: wczytaj zapamiętane definicje i narysuj panel (kolumny zostaną przeliczone
// dopiero po wczytaniu arkusza — load path woła applyDerivedColumns({silent:true})).
dcLoadPersisted();
dcRenderPanel();
