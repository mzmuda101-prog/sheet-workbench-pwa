// entry-playwright.js — Wpisywanie: zapis kartki (swbEntry.plan / commit / undoCard), bez UI.
//
// Co musi być prawdą (plik z entry-fixture.js, tabela „Prace” A1:E6):
//   1. uzupełnienie istniejącego wiersza: data i status w wierszu 3, jeden krok Cofnij,
//   2. nowy wiersz: formuła kolumny obliczeniowej i styl z wiersza wyżej, tabela rośnie w zapisie,
//      walidator OOXML bez nowych błędów,
//   3. nadpisanie niepustej komórki tylko z overwrite,
//   4. lista „stop” blokuje CAŁĄ kartkę (zero zapisów),
//   5. kolumna z formułą — błąd „formula”,
//   6. puste pole nie czyści komórki,
//   7. Cofnij kartki z nowym wierszem — wiersz znika, tabela w zapisie wraca do A1:E6,
//   8. tabela posortowana — cel po numerze wiersza arkusza, nie po pozycji w widoku,
//   9. arkusz zmieniony między plan() a commit() — „stale”,
//  10. filtr + nowy wiersz — „filtered”; tryb Podgląd — „guard”.
// EXCEL_DIR=… zapisuje pliki do sprawdzenia w prawdziwym Excelu (scripts/excel-real-check.py).

const pw = require("playwright");
const { autoAnswerAppDialogs } = require("./app-dialog-auto");
const { makeEntryFixture } = require("./entry-fixture");
const path = require("path");
const fs = require("fs");
const os = require("os");
const JSZip = require("../lib/jszip.min.js");
let ooxml = null;
try { ooxml = require("./ooxml-validate"); } catch { ooxml = null; }

const APP_URL = process.env.APP_URL || "http://127.0.0.1:4175/";
const ENGINE = process.env.ENGINE || "chromium";
const EXCEL_DIR = process.env.EXCEL_DIR || "";
const SLEEP_SCALE = Math.min(1, Math.max(0.1, Number(process.env.SLEEP_SCALE || 1)));
const sleep = (ms) => new Promise((r) => setTimeout(r, Math.round(ms * SLEEP_SCALE)));

async function run() {
  const tmp = fs.mkdtempSync(path.join(os.tmpdir(), "swb-entry-"));
  const FILE = path.join(tmp, "entry.xlsx");
  makeEntryFixture(FILE);
  const browser = await pw[ENGINE].launch({ headless: true });
  const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1280, height: 860 } });
  await context.addInitScript(() => { localStorage.setItem("introPlayed", "true"); });
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  await autoAnswerAppDialogs(page);
  const failures = [];
  const ok = (name, cond, got) => { if (!cond) failures.push(`${name} — dostałem: ${JSON.stringify(got)}`); };

  const open = async () => {
    await page.goto(APP_URL, { waitUntil: "load" });
    await page.evaluate(() => { document.getElementById("heroSplash")?.remove(); return ensureXlsxLibs(false); });
    await page.setInputFiles("#fileInput", FILE);
    await page.waitForFunction(() => currentHeaderRow === 1 && !!document.querySelector("#dataTable tbody tr[data-row-key]"), null, { timeout: 20000 });
    await page.evaluate(() => { sheetEditUi.setMode("edit"); });
    await sleep(300);
  };
  const cell = (ref) => page.evaluate((ref) => {
    const x = workbook.Sheets[currentSheetName][ref];
    if (!x) return null;
    const ymd = (d) => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
    return { v: x.v instanceof Date ? ymd(x.v) : x.v, t: x.t, f: x.f || null, s: x.s ?? null };
  }, ref);
  const styleIdx = (ref) => page.evaluate((ref) => {
    const p = pendingEdits[currentSheetName] && pendingEdits[currentSheetName][ref];
    if (p && Number.isInteger(p.xf)) return p.xf;
    const m = currentStyleIndexMap && currentStyleIndexMap.get(currentSheetName);
    return m ? m.get(ref) ?? null : null;
  }, ref);
  const commit = (input) => page.evaluate((input) => window.__entry.commit(input), input);
  const plan = (input) => page.evaluate((input) => window.__entry.plan(input), input);
  const saveTableRef = async (name) => {
    const bytes = Buffer.from(await page.evaluate(async () => Array.from(await buildOutputBytes("xlsx"))));
    const out = path.join(tmp, name);
    fs.writeFileSync(out, bytes);
    if (EXCEL_DIR) { fs.mkdirSync(EXCEL_DIR, { recursive: true }); fs.writeFileSync(path.join(EXCEL_DIR, name), bytes); }
    const zip = await JSZip.loadAsync(bytes);
    const tbl = Object.keys(zip.files).find((k) => /^xl\/tables\/table\d+\.xml$/.test(k));
    const xml = tbl ? await zip.file(tbl).async("string") : "";
    const sheetXml = await zip.file("xl/worksheets/sheet1.xml").async("string");
    // Excel wymaga komórek w wierszu rosnąco po kolumnach — inaczej okno „naprawić plik?”
    const disorder = [];
    for (const row of sheetXml.matchAll(/<row r="(\d+)"[^>]*>([\s\S]*?)<\/row>/g)) {
      const cols = [...row[2].matchAll(/<c r="([A-Z]+)\d+"/g)].map((m) => m[1].split("").reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0));
      if (cols.some((c, i) => i && c <= cols[i - 1])) disorder.push(row[1]);
    }
    const wbXml = await zip.file("xl/workbook.xml").async("string");
    // wyniki formuł liczonych z innych komórek nie są przepisywane do pliku (formula-calc.js) —
    // dlatego plik MUSI kazać Excelowi przeliczyć wszystko przy otwarciu
    const fullCalc = /<calcPr\b[^>]*fullCalcOnLoad="1"/.test(wbXml);
    return { out, ref: (/<table[^>]*\sref="([^"]+)"/.exec(xml) || [])[1] || null, disorder, fullCalc };
  };

  await open();
  ok("0. moduł dostępny", await page.evaluate(() => typeof window.__entry === "object"), null);

  // 1. uzupełnienie
  const r1 = await commit({ target: { rowIndex0: 2 }, fields: { 1: "12.03.2026", 2: "zdana" } });
  ok("1. commit ok", r1 && r1.ok === true && r1.rowIndex0 === 2, r1);
  ok("1. undoSteps = 1", r1 && r1.undoSteps === 1, r1);
  const b3 = await cell("B3"), c3 = await cell("C3");
  ok("1. B3 = 2026-03-12 (data)", b3 && b3.t === "d" && b3.v === "2026-03-12", b3);
  ok("1. C3 = zdana", c3 && c3.v === "zdana", c3);
  ok("1. A3 bez zmian", (await cell("A3")).v === "Ewa Lis", await cell("A3"));

  // 2. nowy wiersz
  const r2 = await commit({ target: { newRow: true }, fields: { 0: "Olga Wrona", 3: "4" } });
  ok("2. commit nowego wiersza ok, rowIndex0 = 6", r2 && r2.ok && r2.rowIndex0 === 6, r2);
  const e7 = await cell("E7");
  ok("2. E7 = formuła D7*2 z wynikiem 8", e7 && e7.f === "D7*2" && e7.v === 8, e7);
  ok("2. A7 = Olga Wrona", (await cell("A7"))?.v === "Olga Wrona", await cell("A7"));
  const sB6 = await styleIdx("B6"), sB7 = await styleIdx("B7");
  ok("2. wygląd B7 (pusta, format daty) = B6", sB6 != null && sB6 === sB7, { sB6, sB7 });
  const s2 = await saveTableRef("entry-nowy-wiersz.xlsx");
  ok("2. tabela w zapisie A1:E7", s2.ref === "A1:E7", s2.ref);
  ok("2. komórki w każdym wierszu rosnąco po kolumnach (Excel bez naprawy)", s2.disorder.length === 0, s2.disorder);
  ok("2. plik wymusza przeliczenie przy otwarciu (fullCalcOnLoad)", s2.fullCalc === true, s2.fullCalc);
  const ins = await page.evaluate(() => insertCellInRow('<c r="A7" t="s"><v>1</v></c><c r="B7" s="1"/><c r="E7"><f>D7*2</f><v>0</v></c>', "D7", '<c r="D7"><v>4</v></c>'));
  ok("2c. insertCellInRow: komórka samozamykająca nie połyka następnej", /<c r="D7"><v>4<\/v><\/c><c r="E7">/.test(ins), ins);
  if (ooxml && ooxml.available && ooxml.available()) {
    const [b0, a0] = ooxml.validate([FILE, s2.out]);
    const fresh = ooxml.newErrors(b0, a0).map(ooxml.short);
    ok("2. OOXML bez nowych błędów", fresh.length === 0, fresh);
  }

  // 7. Cofnij kartki z nowym wierszem
  const u7 = await page.evaluate(() => window.__entry.undoCard());
  await sleep(200);
  ok("7. undoCard true", u7 === true, u7);
  const a7 = await cell("A7"), e7b = await cell("E7");
  ok("7. wiersz 7 pusty po Cofnij kartki", (!a7 || a7.v == null || a7.t === "z") && (!e7b || !e7b.f), { a7, e7b });
  const s7 = await saveTableRef("entry-po-cofnij.xlsx");
  ok("7. tabela w zapisie wraca do A1:E6", s7.ref === "A1:E6", s7.ref);
  ok("7. wpis z kartki 1 został", (await cell("C3"))?.v === "zdana", await cell("C3"));

  // 3. nadpisanie
  const r3a = await commit({ target: { rowIndex0: 1 }, fields: { 1: "01.01.2027" } });
  ok("3. bez overwrite — invalid / overwrite-not-allowed", r3a && !r3a.ok && r3a.reason === "invalid" && r3a.plan.errors[0]?.reason === "overwrite-not-allowed", r3a);
  ok("3. B2 bez zmian", (await cell("B2")).v === "2026-03-01", await cell("B2"));
  const r3b = await commit({ target: { rowIndex0: 1 }, fields: { 1: "01.01.2027" }, overwrite: [1] });
  ok("3. z overwrite — zapis", r3b && r3b.ok && (await cell("B2")).v === "2027-01-01", { r3b, b2: await cell("B2") });

  // 4. lista stop blokuje całą kartkę
  const r4 = await commit({ target: { rowIndex0: 3 }, fields: { 2: "oddana", 3: "9" }, overwrite: [2, 3] });
  ok("4. stop — invalid dv-stop", r4 && !r4.ok && r4.reason === "invalid" && r4.plan.errors.some((e) => e.col === 2 && e.reason === "dv-stop"), r4);
  ok("4. D4 bez zmian (zero zapisów)", (await cell("D4")).v === 1, await cell("D4"));
  const p4 = await plan({ target: { rowIndex0: 3 }, fields: { 2: "ZDANA" }, overwrite: [2] });
  ok("4b. lista bez wielkości liter", p4.ok && p4.errors.length === 0, p4);

  // 5. kolumna z formułą
  const p5 = await plan({ target: { rowIndex0: 1 }, fields: { 4: "99" }, overwrite: [4] });
  ok("5. formuła — error formula", !p5.ok && p5.errors[0]?.col === 4 && p5.errors[0]?.reason === "formula", p5);
  const p5n = await plan({ target: { newRow: true }, fields: { 4: "99" } });
  ok("5b. formuła w nowym wierszu (kolumna obliczeniowa) — error formula", !p5n.ok && p5n.errors[0]?.reason === "formula", p5n);

  // 6. puste pole nie czyści
  const p6 = await plan({ target: { rowIndex0: 1 }, fields: { 0: "", 3: "   " } });
  ok("6. puste pola — brak zmian", p6.ok && p6.changes.length === 0, p6);
  const p6b = await plan({ target: { rowIndex0: 1 }, fields: { 3: "2" } });
  ok("6b. ta sama wartość — brak zmian", p6b.ok && p6b.changes.length === 0, p6b);

  // 8. sortowanie
  await page.evaluate(() => { setPrimarySort(currentDisplayModel.headers[0], "desc"); renderActiveTable(); });
  await sleep(300);
  const r8 = await commit({ target: { rowIndex0: 4 }, fields: { 3: "7" }, overwrite: [3] });
  ok("8. po sortowaniu wpis w D5 (Ewa Lis 2.)", r8.ok && (await cell("D5")).v === 7 && (await cell("A5")).v === "Ewa Lis", { r8, d5: await cell("D5") });
  await page.evaluate(() => { setPrimarySort("", "asc"); renderActiveTable(); });

  // 9. stale
  const in9 = { target: { rowIndex0: 2 }, fields: { 3: "11" }, overwrite: [3] };
  const p9 = await plan(in9);
  await page.evaluate(() => { const row = baseRows.find((r) => r.rowIndex0 === 2); writeCellFromInput(currentDisplayModel, row, 3, "8"); });
  await sleep(50);
  const r9 = await commit({ ...in9, expect: Object.fromEntries(p9.changes.map((c) => [c.col, c.before])) });
  ok("9. zmiana pod spodem — stale", r9 && !r9.ok && r9.reason === "stale", r9);
  ok("9. D3 = 8 (nie 11)", (await cell("D3")).v === 8, await cell("D3"));

  // 10. filtr i Podgląd
  await page.evaluate(() => { quickSearchEl.value = "Nowak"; applyQuickSearch(); });
  await sleep(400);
  const r10 = await commit({ target: { newRow: true }, fields: { 0: "Zenon" } });
  ok("10. filtr + nowy wiersz — filtered", r10 && !r10.ok && r10.reason === "filtered", r10);
  await page.evaluate(() => { quickSearchEl.value = ""; applyQuickSearch(); });
  await sleep(400);
  await page.evaluate(() => sheetEditUi.setMode("view"));
  const r10b = await commit({ target: { rowIndex0: 2 }, fields: { 3: "12" }, overwrite: [3] });
  ok("10. Podgląd — guard", r10b && !r10b.ok && r10b.reason === "guard", r10b);
  await page.evaluate(() => sheetEditUi.setMode("edit"));

  // 12. plik po wszystkich kartkach (uzupełnienia + nadpisania): tabela bez zmian zakresu, OOXML czysty
  const s12 = await saveTableRef("entry-uzupelnienie.xlsx");
  ok("12. tabela po uzupełnieniach A1:E6", s12.ref === "A1:E6", s12.ref);
  if (ooxml && ooxml.available && ooxml.available()) {
    const [b0, a0] = ooxml.validate([FILE, s12.out]);
    const fresh = ooxml.newErrors(b0, a0).map(ooxml.short);
    ok("12. OOXML bez nowych błędów", fresh.length === 0, fresh);
  }

  ok("11. brak pageerror", errors.length === 0, errors);
  await browser.close();
  if (failures.length) {
    console.log(`entry (${ENGINE}): ${failures.length} błędów`);
    failures.forEach((f) => console.log("  ✗", f));
    process.exit(1);
  }
  console.log(`entry (${ENGINE}): OK`);
}

run().catch((e) => { console.error(e); process.exit(1); });
