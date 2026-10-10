// formula-oracle.mjs — silnik formuł (app/formula-engine.js) kontra EXCEL.
//
// Każda formuła z prawdziwych plików jest liczona naszym silnikiem na danych wejściowych
// zapisanych przez Excela (wartości komórek z pliku, w tym wyniki innych formuł) i
// porównywana z wynikiem, który Excel zapisał w tej komórce. „Dziś” = data ostatniego
// zapisu pliku (tego dnia Excel liczył TODAY()). Formuły z funkcjami, których silnik nie
// zna, są liczone osobno (nie są błędem — Excel przeliczy je przy otwarciu).
//
// Do tego zestaw przypadków z wynikami sprawdzonymi w Excelu (kolejność działań, puste
// komórki, porównania tekstu, zaokrąglanie, daty, wyszukiwanie, kryteria z * ?).
//
// Pliki: scripts/*.xlsx obecne na dysku (prywatne pliki Mateusza są w .gitignore —
// bez nich zostaje syntetyczny fixture z gen_fixtures.py). FILES="a.xlsx,b.xlsx" podmienia.
// Wynik: kod 1, gdy zgodność < PRÓG albo któryś przypadek z listy się nie zgadza.

import fs from "node:fs";
import path from "node:path";
import { createRequire } from "node:module";
import { fileURLToPath } from "node:url";
import { execFileSync } from "node:child_process";

const require = createRequire(import.meta.url);
const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const XLSX = require(path.join(ROOT, "lib/xlsx.full.min.js"));
const JSZip = require(path.join(ROOT, "lib/jszip.min.js"));
const FE = require(path.join(ROOT, "app/formula-engine.js"));
const THRESHOLD = Number(process.env.THRESHOLD || 0.995);

const failures = [];

// ── 1. przypadki z wynikami Excela ──────────────────────────────────────────
{
  const cells = {
    A1: 1, A2: 2, A3: 3, A4: null, A5: "5", A6: true, A7: "tekst",
    B1: "Jan", B2: "anna", B3: "Ola", B4: "Jan", B5: "Jana",
    C1: 45658, C2: 45688, // 2025-01-01, 2025-01-31
    D1: 10, D2: 20, D3: 30, D4: 40,
    E1: "W trakcie", E2: "Zamknięte", E3: "W trakcie - pilne", E4: null,
  };
  const ctx = mkCtx("S", (sh, r, c) => {
    const ref = XLSX.utils.encode_cell({ r, c });
    return Object.prototype.hasOwnProperty.call(cells, ref) ? cells[ref] : null;
  }, { r2: 10, c2: 5 }, { today: 45900 });
  const CASES = [
    ["=-2^2", 4], ["=2^3^2", 64], ["=1+2*3", 7], ["=(1+2)*3", 9], ["=10/4", 2.5], ["=1/0", "#DIV/0!"],
    ["=50%", 0.5], ["=2*50%", 1], ["=\"a\"&1&TRUE", "a1TRUE"], ["=A4+1", 1], ["=A5+1", 6], ["=A7+1", "#VALUE!"],
    ["=SUM(A1:A7)", 6], ["=SUM(A5,1)", 1], ["=COUNT(A1:A7)", 3], ["=COUNTA(A1:A7)", 6], ["=COUNTBLANK(A1:A7)", 1],
    ["=AVERAGE(A1:A3)", 2], ["=AVERAGE(A4)", "#DIV/0!"], ["=MAX(D1:D4)", 40], ["=MIN(D1:D4,5)", 5],
    ["=A4=\"\"", true], ["=A4=0", true], ["=\"a\"=\"A\"", true], ["=\"b\">\"A\"", true], ["=1<\"a\"", true], ["=TRUE>\"z\"", true],
    ["=ROUND(2.675,2)", 2.68], ["=ROUND(-2.5,0)", -3], ["=ROUND(1234.567,-2)", 1200], ["=ROUNDUP(1.21,1)", 1.3], ["=ROUNDDOWN(-1.29,1)", -1.2],
    ["=INT(-1.5)", -2], ["=TRUNC(-1.5)", -1], ["=MOD(-3,2)", 1], ["=MOD(3,-2)", -1],
    ["=IF(A1>0,\"tak\",\"nie\")", "tak"], ["=IF(A4,1,2)", 2], ["=IF(1>2,1)", false], ["=IFERROR(1/0,\"x\")", "x"], ["=IF(TRUE,,1)", 0],
    ["=AND(A1,A2>1)", true], ["=OR(A4,0)", false], ["=NOT(0)", true],
    ["=COUNTIF(B1:B5,\"jan\")", 2], ["=COUNTIF(B1:B5,\"Jan*\")", 3], ["=COUNTIF(B1:B5,\"?la\")", 1], ["=COUNTIF(D1:D4,\">=20\")", 3],
    ["=COUNTIF(E1:E4,\"*W trakcie*\")", 2], ["=COUNTIF(E1:E4,\"\")", 1], ["=COUNTIF(E1:E4,\"<>\")", 3],
    ["=SUMIF(B1:B4,\"Jan\",D1:D4)", 50], ["=SUMIFS(D1:D4,B1:B4,\"Jan\",D1:D4,\">15\")", 40], ["=AVERAGEIF(D1:D4,\">10\")", 30],
    ["=VLOOKUP(\"ola\",B1:D4,3,FALSE)", 30], ["=VLOOKUP(\"x\",B1:D4,3,FALSE)", "#N/A"], ["=VLOOKUP(25,D1:D4,1)", 20], ["=VLOOKUP(5,D1:D4,1)", "#N/A"],
    ["=MATCH(30,D1:D4,0)", 3], ["=MATCH(35,D1:D4)", 3], ["=INDEX(D1:D4,2)", 20], ["=INDEX(B1:D4,3,3)", 30],
    ["=LEN(\"Zażółć\")", 6], ["=LEFT(\"abc\")", "a"], ["=MID(\"abcdef\",2,3)", "bcd"], ["=RIGHT(\"abc\",5)", "abc"], ["=TRIM(\"  a   b \")", "a b"],
    ["=FIND(\"B\",\"abcB\")", 4], ["=SEARCH(\"b\",\"abcB\")", 2], ["=SEARCH(\"c*B\",\"abcB\")", 3], ["=FIND(\"x\",\"abc\")", "#VALUE!"],
    ["=SUBSTITUTE(\"a-b-c\",\"-\",\"+\")", "a+b+c"], ["=SUBSTITUTE(\"a-b-c\",\"-\",\"+\",2)", "a-b+c"], ["=UPPER(\"ąę\")", "ĄĘ"], ["=PROPER(\"jan KOWALSKI\")", "Jan Kowalski"],
    ["=DATE(2025,2,30)", 45718], ["=YEAR(C1)", 2025], ["=MONTH(C2)", 1], ["=DAY(C2)", 31], ["=EDATE(C2,1)", 45716], ["=EOMONTH(C1,1)", 45716],
    ["=DATEDIF(C1,C2,\"D\")", 30], ["=DATEDIF(DATE(2024,1,31),DATE(2024,3,1),\"M\")", 1], ["=DATEDIF(DATE(2024,1,31),DATE(2024,3,1),\"MD\")", -1],
    ["=DATEDIF(DATE(2020,5,15),DATE(2025,5,14),\"Y\")", 4], ["=DATEDIF(C2,C1,\"D\")", "#NUM!"], ["=WEEKDAY(C1)", 4], ["=WEEKDAY(C1,2)", 3],
    ["=NETWORKDAYS(C1,C2)", 23], ["=WORKDAY(C1,5)", 45665], ["=TODAY()", 45900], ["=DAYS(C2,C1)", 30],
    ["=SUMPRODUCT(A1:A3,D1:D3)", 140], ["=SUMPRODUCT((D1:D4>15)*D1:D4)", 90], ["=LARGE(D1:D4,2)", 30], ["=SMALL(D1:D4,1)", 10],
    ["=RANK.EQ(30,D1:D4)", 2], ["=MEDIAN(1,2,3,4)", 2.5], ["=SUBTOTAL(9,D1:D4)", 100], ["=SUBTOTAL(109,D1:D4)", 100],
    ["=CHOOSE(2,\"a\",\"b\")", "b"], ["=SWITCH(2,1,\"a\",2,\"b\",\"c\")", "b"], ["=IFS(A1>5,\"a\",A1>0,\"b\")", "b"],
    ["=XLOOKUP(\"Ola\",B1:B4,D1:D4)", 30], ["=XLOOKUP(\"x\",B1:B4,D1:D4,\"brak\")", "brak"], ["=TEXTJOIN(\"-\",TRUE,B1:B3,A4)", "Jan-anna-Ola"],
    ["=CONCATENATE(\"a\",A1)", "a1"], ["=ISNUMBER(A5)", false], ["=ISBLANK(A4)", true], ["=ISERROR(1/0)", true], ["=ISNA(NA())", true],
    ["=VALUE(\"1,000.5\")", 1000.5], ["=VALUE(\"(5)\")", -5], ["=\"$4.00\"+1", 5], ["=\"12:00\"*2", 1], ["=0^-1", "#DIV/0!"], ["=0^0", "#NUM!"], ["=(-8)^(1/3)", -1.9999999999999998], ["=ROUND(1.005,2)", 1.01], ["=ROUND(1234567890123.375,2)", 1234567890123.37], ["=1234.5678-1234.5678000000003", 0], ["=(1234.5678-1234.5678000000003)", -2.2737367544323206e-13], ["=SUM(3,1E16,-1E16,0)", 4], ["=+TRUE", true], ["=0.1+0.2=0.3", true], ["=\"x\"&0.1+0.2", "x0.3"], ["=CEILING.MATH(4.3)", 5], ["=FLOOR.MATH(-4.3)", -5],
    ["=MROUND(10,3)", 9], ["=ROW(B7)", 7], ["=COLUMNS(A1:D1)", 4], ["=ABS(-3)", 3], ["=SQRT(-1)", "#NUM!"], ["=POWER(2,10)", 1024],
  ];
  // polski zapis — ten sam wynik
  const CASES_PL = [
    ["=SUMA(A1:A3;1,5)", 7.5], ["=JEŻELI(A1>0;\"tak\";\"nie\")", "tak"], ["=LICZ.JEŻELI(B1:B5;\"Jan*\")", 3],
    ["=WYSZUKAJ.PIONOWO(\"ola\";B1:D4;3;FAŁSZ)", 30], ["=ZAOKR(2,675;2)", 2.68], ["=srednia(A1:A3)", 2], ["=A1*1,5", 1.5], ["=IF(A1>0,1,0)", 1],
  ];
  const check = (src, want, lang) => {
    let got;
    try { got = FE.evaluate(FE.parse(src, { lang }), ctx); } catch (e) { got = "SYNTAX: " + e.message; }
    const g = FE.isErr(got) ? got.err : got;
    const ok = typeof want === "number" && typeof g === "number" ? Math.abs(g - want) < 1e-9 * Math.max(1, Math.abs(want)) : g === want;
    if (!ok) failures.push(`${src} (${lang}) = ${JSON.stringify(g)}, Excel: ${JSON.stringify(want)}`);
  };
  CASES.forEach(([s, w]) => check(s, w, "en"));
  FE.setLocale("pl"); // tekst-liczba po polsku (jak Excel z polskimi ustawieniami regionalnymi)
  CASES_PL.forEach(([s, w]) => check(s, w, "pl"));
  check('=WARTOŚĆ("1 234,5")', 1234.5, "pl");
  check('="12,5 zł"*2', 25, "pl");
  FE.setLocale("en");
  // zapis do pliku i z powrotem
  const roundtrip = [["=JEŻELI.BŁĄD(X.WYSZUKAJ(A2;B:B;C:C);\"brak\")", "IFERROR(_xlfn.XLOOKUP(A2,B:B,C:C),\"brak\")"], ["=SUMA(A1:A3;1,5)", "SUM(A1:A3,1.5)"]];
  roundtrip.forEach(([src, want]) => {
    const got = FE.serialize(FE.parse(src, { lang: "pl" }), { file: true });
    if (got !== want) failures.push(`zapis do pliku: ${src} → ${got}, oczekiwane ${want}`);
  });
  // wypełnianie w dół (przesunięcie odwołań względnych)
  const sh = FE.serialize(FE.shift(FE.parse("=A1+$B$1+B$2+$C3+SUM(D1:D2)", { lang: "en" }), 2, 1), { file: true });
  if (sh !== "B3+$B$1+C$2+$C5+SUM(E3:E4)") failures.push(`przesunięcie: ${sh}`);
  console.log(`przypadki: ${CASES.length + CASES_PL.length + roundtrip.length + 1}, błędów: ${failures.length}`);
}

// ── 2. pliki: każda formuła vs wynik Excela ─────────────────────────────────
function mkCtx(sheetName, get, bounds, extra = {}) {
  return {
    sheetName, row: 0, col: 0, depth: 0, lang: "en",
    get, bounds: typeof bounds === "function" ? bounds : () => bounds,
    sheetFor: extra.sheetFor || ((n) => n),
    now: () => extra.today ?? 45900,
    structured: extra.structured, name: extra.name,
    formatNumber: (fmt, n) => { try { return XLSX.SSF.format(fmt, n); } catch { return null; } },
  };
}

async function tablesOf(bytes, sheetNames) {
  const zip = await JSZip.loadAsync(bytes);
  const wbXml = await zip.file("xl/workbook.xml").async("string");
  const rels = await zip.file("xl/_rels/workbook.xml.rels").async("string");
  const ridTarget = {};
  for (const m of rels.matchAll(/<Relationship\b[^>]*>/g)) {
    const id = (m[0].match(/\bId="([^"]+)"/) || [])[1];
    const tg = (m[0].match(/\bTarget="([^"]+)"/) || [])[1];
    if (id && tg) ridTarget[id] = tg.replace(/^\/?xl\//, "").replace(/^\//, "");
  }
  const sheetByPath = {};
  for (const m of wbXml.matchAll(/<sheet\b[^>]*>/g)) {
    const name = (m[0].match(/\bname="([^"]+)"/) || [])[1];
    const rid = (m[0].match(/\br:id="([^"]+)"/) || [])[1];
    if (name && rid && ridTarget[rid]) sheetByPath["xl/" + ridTarget[rid]] = name.replace(/&amp;/g, "&").replace(/&quot;/g, '"').replace(/&apos;/g, "'");
  }
  const tables = {};
  for (const f of Object.keys(zip.files).filter((x) => /^xl\/worksheets\/_rels\/.+\.rels$/.test(x))) {
    const sheetPath = f.replace("/_rels/", "/").replace(/\.rels$/, "");
    const sheet = sheetByPath[sheetPath];
    const x = await zip.file(f).async("string");
    for (const m of x.matchAll(/Target="([^"]*tables\/(table\d+\.xml))"/g)) {
      const t = await zip.file("xl/tables/" + m[2]).async("string");
      const ref = XLSX.utils.decode_range((t.match(/\bref="([^"]+)"/) || [])[1]);
      const cols = [...t.matchAll(/<tableColumn\b[^>]*\bname="([^"]*)"/g)].map((c) => c[1].replace(/&amp;/g, "&").replace(/_x000a_/gi, "\n"));
      const meta = {
        sheet, r1: ref.s.r, c1: ref.s.c, r2: ref.e.r, c2: ref.e.c, columns: cols,
        headerRows: Number((t.match(/\bheaderRowCount="(\d+)"/) || [])[1] ?? 1),
        totalsRows: Number((t.match(/\btotalsRowCount="(\d+)"/) || [])[1] ?? 0),
      };
      for (const key of ["name", "displayName"]) {
        const n = (t.match(new RegExp(`\\b${key}="([^"]+)"`)) || [])[1];
        if (n) tables[n.toLowerCase()] = meta;
      }
    }
  }
  return tables;
}

function cellValue(cell) {
  if (!cell || cell.t === "z") return null;
  if (cell.t === "e") return FE.ERR[cell.w] || { err: cell.w || "#VALUE!" };
  if (cell.t === "b") return !!cell.v;
  if (cell.t === "n") return cell.v;
  if (cell.t === "s" || cell.t === "str") return cell.v == null ? "" : String(cell.v);
  return cell.v ?? null;
}

function same(a, b) {
  if (FE.isErr(a) || FE.isErr(b)) return FE.isErr(a) && FE.isErr(b) && a.err === b.err;
  if (typeof a === "number" && typeof b === "number") return Math.abs(a - b) <= 1e-9 * Math.max(1, Math.abs(a), Math.abs(b));
  if (a === null && b === "") return true;
  return a === b;
}

const PY = fs.existsSync(path.join(ROOT, ".venv/bin/python")) ? path.join(ROOT, ".venv/bin/python") : "python3";
let fixture = null;
try { fixture = execFileSync(PY, [path.join(ROOT, "scripts/py/gen_fixtures.py")], { encoding: "utf8" }).trim(); } catch { /* bez Pythona */ }
const files = process.env.FILES
  ? process.env.FILES.split(",").map((f) => path.resolve(ROOT, f.trim()))
  : [fixture, ...fs.readdirSync(path.join(ROOT, "scripts")).filter((f) => /\.xlsx$/i.test(f)).map((f) => path.join(ROOT, "scripts", f))].filter((f) => f && fs.existsSync(f));

for (const file of files) {
  const bytes = fs.readFileSync(file);
  const wb = XLSX.read(bytes, { cellFormula: true, cellDates: false, cellNF: true });
  const tables = await tablesOf(bytes, wb.SheetNames);
  const mod = wb.Props && wb.Props.ModifiedDate ? new Date(wb.Props.ModifiedDate) : new Date();
  const today = FE.dateSerial(mod.getFullYear(), mod.getMonth() + 1, mod.getDate());
  const lowerSheets = new Map(wb.SheetNames.map((n) => [n.toLowerCase(), n]));
  const boundsCache = new Map();
  const base = mkCtx(null, (sh, r, c) => {
    const ws = wb.Sheets[sh];
    return ws ? cellValue(ws[XLSX.utils.encode_cell({ r, c })]) : null;
  }, (sh) => {
    if (!boundsCache.has(sh)) {
      const ref = wb.Sheets[sh] && wb.Sheets[sh]["!ref"];
      const r = ref ? XLSX.utils.decode_range(ref) : { e: { r: 0, c: 0 } };
      boundsCache.set(sh, { r2: r.e.r, c2: r.e.c });
    }
    return boundsCache.get(sh);
  }, { today, sheetFor: (n) => lowerSheets.get(String(n).toLowerCase()) ?? null, structured: FE.makeStructuredResolver(tables) });
  base.name = FE.makeNameResolver((wb.Workbook && wb.Workbook.Names) || [], wb.SheetNames);

  let total = 0, okN = 0, unsupported = 0, noCache = 0;
  const bad = new Map();
  const astCache = new Map();
  for (const sn of wb.SheetNames) {
    const ws = wb.Sheets[sn];
    for (const ref of Object.keys(ws)) {
      if (ref[0] === "!") continue;
      const cell = ws[ref];
      if (!cell || !cell.f) continue;
      if (cell.F && cell.F !== ref) { /* członek formuły tablicowej — pomijamy */ }
      if (cell.v === undefined && cell.t !== "e") { noCache++; continue; }
      total++;
      let ast = astCache.get(cell.f);
      if (ast === undefined) {
        try { ast = FE.parse(cell.f, { lang: "en" }); } catch (e) { ast = { syntax: e.message }; }
        astCache.set(cell.f, ast);
      }
      const want = cellValue(cell);
      if (ast.syntax) { const k = "SKŁADNIA: " + ast.syntax; bad.set(k, [...(bad.get(k) || []), `${sn}!${ref} = ${cell.f}`]); continue; }
      const fns = [...FE.functionsOf(ast)];
      if (fns.some((f) => !FE.isSupported(f))) { unsupported++; continue; }
      const { r, c } = XLSX.utils.decode_cell(ref);
      let got;
      try { got = FE.evaluate(ast, { ...base, sheetName: sn, row: r, col: c }); } catch (e) { got = { err: "WYJĄTEK " + e.message }; }
      if (same(got, want)) okN++;
      else {
        const key = fns.sort().join("+") || "(bez funkcji)";
        const list = bad.get(key) || [];
        if (list.length < 3) list.push(`${sn}!${ref} = ${cell.f.slice(0, 140)} → nasz ${JSON.stringify(FE.isErr(got) ? got.err : got)}, Excel ${JSON.stringify(FE.isErr(want) ? want.err : want)}`);
        list.n = (list.n || 0) + 1;
        bad.set(key, list);
      }
    }
  }
  const checked = total - unsupported;
  const rate = checked ? okN / checked : 1;
  console.log(`${rate >= THRESHOLD ? "✅" : "❌"} ${path.basename(file)}: ${okN}/${checked} zgodnych z Excelem (${(rate * 100).toFixed(2)}%)` +
    `${unsupported ? `, ${unsupported} z funkcjami spoza silnika` : ""}${noCache ? `, ${noCache} bez zapisanego wyniku` : ""} · dziś=${mod.toISOString().slice(0, 10)}`);
  for (const [k, list] of [...bad.entries()].sort((a, b) => (b[1].n || b[1].length) - (a[1].n || a[1].length)).slice(0, 12)) {
    console.log(`   ${k}: ${list.n || list.length}×`);
    list.slice(0, 2).forEach((x) => console.log("      " + x));
  }
  if (rate < THRESHOLD) failures.push(`${path.basename(file)}: zgodność ${(rate * 100).toFixed(2)}% < ${(THRESHOLD * 100).toFixed(1)}%`);
}

if (failures.length) {
  console.error("\n❌ formula-oracle:\n  - " + failures.join("\n  - "));
  process.exit(1);
}
console.log("\n✅ formula-oracle: silnik zgodny z Excelem");
