// formula-shape-cache-playwright.js — strażnik pamięci „kształtów” formuł (formula-calc.js).
//
// Indeks zależności parsuje każdą formułę pliku. Formuły różniące się tylko przesunięciem
// względnych adresów („=B2*C2”, „=B3*C3”…) dzielą jedno drzewo przesuwane FC.shift.
// Test: dla KAŻDEJ formuły z plików drzewo z pamięci kształtów == FC.parse tej formuły
// (JSON), plus przypadki złośliwe (adresy w tekście, arkusze z cyframi, zakresy kolumn
// i wierszy, nazwy podobne do adresów, $ mieszane, apostrofy w nazwach arkuszy).
//
//   node scripts/formula-shape-cache-playwright.js
//   XLSX_DIR="/ścieżka/do/plików" node scripts/formula-shape-cache-playwright.js

const { chromium } = require("playwright");
const fs = require("fs");
const path = require("path");

const APP_URL = process.env.APP_URL || "http://127.0.0.1:4175/";
const FILES = [
  path.join(__dirname, "RODO Obieg terenów 2026 V5.xlsx"),
  path.join(__dirname, "Kosztorys Renault Laguna 3 .xlsx"),
  path.join(__dirname, "fixtures/openpyxl-features.xlsx"),
  path.join(__dirname, "fixtures/threaded-comments.xlsx"),
];
if (process.env.XLSX_DIR) {
  for (const n of fs.readdirSync(process.env.XLSX_DIR)) if (/\.xlsx$/i.test(n) && !n.startsWith("~$")) FILES.push(path.join(process.env.XLSX_DIR, n));
}

// [formuła, wiersz, kolumna] — kolejne wpisy tej samej grupy mają ten sam klucz kształtu,
// więc drugi i dalsze idą ścieżką FC.shift (albo bezpiecznikiem do pełnego parsowania).
const TRICKY = [
  ["B2*C2", 1, 3], ["B3*C3", 2, 3], ["C3*D3", 2, 4], ["b4*c4", 3, 3],
  ['"A2"&A2', 1, 1], ['"A2"&A3', 2, 1], ['"A3"&A3', 2, 1],
  ["Sheet2!A2+A2", 1, 2], ["Sheet2!A3+A3", 2, 2], ["'Ark 2'!B2+'O''Brien 3'!B2", 1, 3], ["'Ark 2'!B3+'O''Brien 3'!B3", 2, 3],
  ["SUM(A:A)+B2", 1, 2], ["SUM(A:A)+C2", 1, 3], ["SUM(A:A)+C3", 2, 3],
  ["SUM(2:2)+B2", 1, 2], ["SUM(2:2)+B3", 2, 2],
  ["SUM($A$1:A2)", 1, 1], ["SUM($A$1:A3)", 2, 1], ["SUM($A$1:B3)", 2, 2],
  ["$A2+B$1", 1, 2], ["$A3+B$1", 2, 2], ["$A3+C$1", 2, 3],
  ["LOG10(A2)", 1, 1], ["LOG10(A3)", 2, 1],
  ["Tax2023+A2", 1, 1], ["Tax2023+A3", 2, 1],
  ["Tabela1[@Kwota]*A2", 1, 1], ["Tabela1[@Kwota]*A3", 2, 1],
  ["IF(A2>0,\"tak\",\"nie\")", 1, 2], ["IF(A3>0,\"tak\",\"nie\")", 2, 2],
  ["SUM({1,2;3,4})+A2", 1, 1], ["SUM({1,2;3,4})+A3", 2, 1],
  ["A1", 5, 5], ["A2", 6, 5], ["B1", 5, 6], ["A1", 0, 0],
  ["VLOOKUP(A2,$D$2:$E$100,2,FALSE)", 1, 1], ["VLOOKUP(A3,$D$2:$E$100,2,FALSE)", 2, 1],
  ["SUM(Sheet1!A1:A10)", 0, 0], ["SUM(Sheet1!A2:A11)", 1, 0],
  ["A1.5", 0, 0], ["#REF!+A2", 1, 0], ["#REF!+A3", 2, 0],
  ["XFD2+A2", 1, 0], ["XFD3+A3", 2, 0],
];

(async () => {
  const browser = await chromium.launch({ headless: true });
  const context = await browser.newContext({ serviceWorkers: "block" });
  await context.addInitScript(() => localStorage.setItem("introPlayed", "true"));
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(() => { try { ensureXlsxLibs && ensureXlsxLibs(false); } catch {} });
  let fail = 0;

  const tricky = await page.evaluate((list) => {
    fcAstCache.clear(); fcShapes.clear();
    const bad = [];
    let viaShape = 0;
    for (const [f, r, c] of list) {
      let direct;
      try { direct = JSON.stringify(FC.parse(f, { lang: "en" })); } catch { direct = "null"; }
      const before = fcShapes.size;
      const got = JSON.stringify(fcParseFile(f, r, c).ast);
      if (fcShapes.size === before) viaShape++;
      if (got !== direct) bad.push({ f, r, c, got, direct });
      fcAstCache.clear(); // tylko pamięć kształtów może dać trafienie
    }
    return { bad, viaShape, n: list.length };
  }, TRICKY);
  console.log(`przypadki złośliwe: ${tricky.n}, z pamięci kształtu ${tricky.viaShape}, różnic ${tricky.bad.length}`);
  tricky.bad.forEach((b) => console.log("  ✗", JSON.stringify(b)));
  fail += tricky.bad.length;

  for (const file of FILES) {
    if (!fs.existsSync(file)) { console.log(`(brak pliku ${file})`); continue; }
    await page.setInputFiles("#fileInput", file);
    await page.waitForFunction((name) => document.getElementById("sheetSelect")?.options?.length > 0 && typeof workbook !== "undefined" && workbook && window.__lastFileName !== name, path.basename(file), { timeout: 60000 }).catch(() => {});
    await page.waitForTimeout(300);
    const r = await page.evaluate(() => {
      fcAstCache.clear(); fcShapes.clear();
      let n = 0, diffs = 0;
      const sample = [];
      const t0 = performance.now();
      for (const sn of workbook.SheetNames) {
        const ws = workbook.Sheets[sn];
        for (const ref in ws) {
          if (ref.charCodeAt(0) === 33) continue;
          const cell = ws[ref];
          if (!cell || !cell.f) continue;
          const { r, c } = XLSX.utils.decode_cell(ref);
          n++;
          const got = JSON.stringify(fcParseFile(cell.f, r, c).ast);
          let direct;
          try { direct = JSON.stringify(FC.parse(cell.f, { lang: "en" })); } catch { direct = "null"; }
          if (got !== direct) { diffs++; if (sample.length < 5) sample.push({ sn, ref, f: cell.f }); }
        }
      }
      let shapes = 0, unsafe = 0;
      fcShapes.forEach((v) => { shapes++; if (!v) unsafe++; });
      return { n, diffs, sample, shapes, unsafe, ms: Math.round(performance.now() - t0) };
    });
    console.log(`${path.basename(file)}: formuł ${r.n}, kształtów ${r.shapes} (wyłączonych ${r.unsafe}), różnic ${r.diffs}`);
    r.sample.forEach((s) => console.log("  ✗", JSON.stringify(s)));
    fail += r.diffs;
  }

  if (errors.length) { console.log("błędy strony:", errors.slice(0, 5)); fail++; }
  await browser.close();
  console.log(fail ? `\nFAIL (${fail})` : "\nOK — drzewa z pamięci kształtów identyczne z FC.parse");
  process.exit(fail ? 1 : 0);
})().catch((e) => { console.error(e); process.exit(1); });
