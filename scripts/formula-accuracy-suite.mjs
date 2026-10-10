// formula-accuracy-suite.mjs — silnik formuł kontra zewnętrzny zestaw „Spreadsheet Accuracy Suite”
// (github.com/MacJedi42/Spreadsheet-Accuracy-Suite, MIT): 44 543 formuły z odpowiedziami Excela
// (excel:observed — Excel for Mac 16.113, bit w bit) albo LibreOffice tam, gdzie zgadza się z Excelem
// (lo), plus kody błędów (error:<kod>). Mocno skupiony na precyzji: remisy przy ZAOKR, odejmowanie
// prawie równych liczb, porównania 15-cyfrowe, potęgi, sumy dziesiętne; do tego 29 tys. formuł
// z prawdziwych plików (SpreadsheetBench, FUSE, Enron).
//
// Dane NIE są w repozytorium (17 MB) — pobiera je jednorazowo:
//   npm run suite:fetch      (do .cache/accuracy-suite, katalog w .gitignore)
// Bez danych test się pomija (z ostrzeżeniem).
//
// Raport: zgodność per źródło odpowiedzi (excel / lo / błąd) i per grupa; dokładnie (ten sam double)
// i „praktycznie” (różnica ≤ 1e-13 względnie — to, czego nie widać przy 15 cyfrach Excela).
// Próg (PROG, domyślnie 0.999) dotyczy wierszy excel:* + error:* liczonych praktycznie.

import fs from "node:fs";
import path from "node:path";
import { createRequire } from "node:module";
import { fileURLToPath } from "node:url";

const require = createRequire(import.meta.url);
const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const FE = require(path.join(ROOT, "app/formula-engine.js"));
const TSV = path.join(ROOT, ".cache/accuracy-suite/formula_golden.tsv");
const PROG = Number(process.env.PROG || 0.999);
const VERBOSE = !!process.env.VERBOSE;

if (!fs.existsSync(TSV)) {
  console.log("⚠️  Pominięto: brak danych Spreadsheet Accuracy Suite (npm run suite:fetch).");
  process.exit(0);
}

const col = (s) => FE.colToIndex(s);
function parseInputs(field) {
  const cells = new Map(); // "sheet!r!c" (sheet lower) → value
  if (!field || field === "-") return cells;
  for (const part of field.split("\u001f")) {
    const m = part.match(/^('(?:[^']|'')+'|[^!]+)!([A-Z]+)(\d+)=(n|b|s|e):([\s\S]*)$/);
    if (!m) continue;
    const sheet = (m[1].startsWith("'") ? m[1].slice(1, -1).replace(/''/g, "'") : m[1]).toLowerCase();
    let v;
    if (m[4] === "n") v = Number(m[5]);
    else if (m[4] === "b") v = m[5] === "TRUE";
    else if (m[4] === "e") v = FE.ERR[m[5]] || { err: m[5] };
    else v = m[5];
    cells.set(`${sheet}!${+m[3] - 1}!${col(m[2])}`, v);
  }
  return cells;
}

const rows = fs.readFileSync(TSV, "utf8").split("\n").filter((l) => l.startsWith("F\t"));
const stats = new Map(); // klucz → { n, exact, close, unsupported, wrong: [] }
const bump = (key, field, sample) => {
  if (!stats.has(key)) stats.set(key, { n: 0, exact: 0, close: 0, unsupported: 0, syntax: 0, wrong: [] });
  const s = stats.get(key);
  s.n++;
  if (field) s[field]++;
  if (sample && s.wrong.length < 4) s.wrong.push(sample);
};

for (const line of rows) {
  const [, id, home, formula, inputs, expect, basis] = line.split("\t");
  if (!expect || expect === "refuse" && !/^error:/.test(basis)) continue; // nierozstrzygnięte / tekst — nie oceniamy
  if (/^undecided/.test(basis)) continue;
  const src = /^error:/.test(basis) ? "błąd" : /^excel/.test(basis) ? "excel" : /^lo/.test(basis) ? "lo" : basis.split(/[:(]/)[0];
  const group = id.split(".")[0];
  const cellAt = id.split("@")[1];
  const fc = cellAt ? XLSXdecode(cellAt) : { r: 0, c: 0 };
  const cells = parseInputs(inputs);
  const sheets = new Set([home.toLowerCase(), ...[...cells.keys()].map((k) => k.split("!")[0])]);
  const ctx = {
    sheetName: home.toLowerCase(), row: fc.r, col: fc.c, depth: 0, lang: "en",
    get: (sh, r, c) => (cells.has(`${sh}!${r}!${c}`) ? cells.get(`${sh}!${r}!${c}`) : null),
    bounds: () => ({ r2: 1048575, c2: 16383 }),
    sheetFor: (n) => (sheets.has(String(n).toLowerCase()) ? String(n).toLowerCase() : String(n).toLowerCase()),
    now: () => 46300,
    structured: () => null, name: () => null,
  };
  const key = `${src} · ${group.length > 14 ? group.slice(0, 5) : group}`;
  let ast;
  try { ast = FE.parse(formula, { lang: "en" }); } catch (e) { bump(key, "syntax", `${id}: ${formula} → składnia: ${e.message}`); continue; }
  const fns = [...FE.functionsOf(ast)];
  if (fns.some((f) => !FE.isSupported(f))) { bump(key, "unsupported"); continue; }
  let got;
  try { got = FE.evaluate(ast, ctx); } catch (e) { got = { err: "WYJĄTEK " + e.message }; }
  let ok = false, close = false;
  if (/^error:/.test(basis)) {
    const code = basis.slice(6).split(/[ (]/)[0];
    ok = FE.isErr(got) && got.err === code;
  } else if (expect.startsWith("n:")) {
    const want = Number(expect.slice(2));
    if (typeof got === "number") {
      ok = Object.is(got, want) || got === want;
      close = !ok && Math.abs(got - want) <= 1e-13 * Math.max(1, Math.abs(want));
    }
  } else if (expect.startsWith("b:")) {
    ok = got === (expect === "b:TRUE");
  }
  const shown = FE.isErr(got) ? got.err : typeof got === "number" ? String(got) : JSON.stringify(got);
  if (ok) bump(key, "exact");
  else if (close) bump(key, "close", VERBOSE ? `${id}: ${formula} [${inputs.replace(/\u001f/g, " ")}] → ${shown}, oczekiwane ${expect}` : null);
  else bump(key, null, `${id}: ${formula.slice(0, 120)} [${inputs.replace(/\u001f/g, " ").slice(0, 120)}] → ${shown}, oczekiwane ${expect} (${basis.slice(0, 60)})`);
}

function XLSXdecode(ref) {
  const m = String(ref).match(/^\$?([A-Z]+)\$?(\d+)$/);
  return m ? { r: +m[2] - 1, c: col(m[1]) } : { r: 0, c: 0 };
}

// podsumowanie per źródło
const bySrc = new Map();
for (const [key, s] of stats) {
  const src = key.split(" · ")[0];
  const t = bySrc.get(src) || { n: 0, exact: 0, close: 0, unsupported: 0, syntax: 0 };
  ["n", "exact", "close", "unsupported", "syntax"].forEach((k) => { t[k] += s[k]; });
  bySrc.set(src, t);
}
const pct = (a, b) => (b ? ((a / b) * 100).toFixed(2) + "%" : "—");
console.log("Spreadsheet Accuracy Suite — silnik formuł vs Excel");
for (const [src, t] of bySrc) {
  const judged = t.n - t.unsupported;
  console.log(`  ${src.padEnd(6)} ${String(t.exact).padStart(6)} dokładnie + ${String(t.close).padStart(4)} ≤1e-13 / ${judged} ocenionych (${pct(t.exact + t.close, judged)}), poza silnikiem: ${t.unsupported}, składnia: ${t.syntax}`);
}
console.log("\nNajwięcej rozbieżności:");
[...stats.entries()]
  .map(([k, s]) => [k, s, s.n - s.unsupported - s.exact - s.close])
  .filter(([, , bad]) => bad > 0)
  .sort((a, b) => b[2] - a[2])
  .slice(0, Number(process.env.TOP || 14))
  .forEach(([k, s, bad]) => {
    console.log(`  ${k}: ${bad} z ${s.n - s.unsupported}`);
    s.wrong.slice(0, 2).forEach((w) => console.log("      " + w));
  });

const ex = ["excel", "błąd"].map((k) => bySrc.get(k) || { n: 0, exact: 0, close: 0, unsupported: 0 });
const judged = ex.reduce((a, t) => a + t.n - t.unsupported, 0);
const good = ex.reduce((a, t) => a + t.exact + t.close, 0);
const rate = judged ? good / judged : 1;
if (rate < PROG) {
  console.error(`\n❌ accuracy-suite: zgodność z Excelem ${pct(good, judged)} < ${(PROG * 100).toFixed(1)}%`);
  process.exit(1);
}
console.log(`\n✅ accuracy-suite: zgodność z Excelem ${pct(good, judged)} (próg ${(PROG * 100).toFixed(1)}%)`);
