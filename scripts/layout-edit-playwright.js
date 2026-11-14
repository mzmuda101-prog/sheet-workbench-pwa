// layout-edit-playwright.js — szerokość kolumny i wysokość wiersza z aplikacji trafiają do pliku
// (app/layout-edit.js + app/xlsx-layout.js; docs/OD-ZERA.md, poz. 1).
//   node scripts/layout-edit-playwright.js        ENGINE=webkit node scripts/layout-edit-playwright.js
//   EXCEL_DIR=/private/tmp/swb-excel-check …       (+ kopia zapisanego pliku dla excel-real-check.py)
//
// Pilnuje:
//   0. (bez przeglądarki) xlsx-layout: podział zakresu <col min max> z zachowaniem stylu, nowy <cols>
//      przed <sheetData>, wiersze w kolejności, pusty <sheetData/>, przeliczenie px → szerokość jak Excel Mac;
//   1. Edycja: przeciągnięcie kolumny B i wiersza 8 = 2 zmiany do zapisu, plik „niezapisany”;
//   2. Podgląd: przeciąganie zmienia tylko widok (nic do zapisu);
//   3. zapis: openpyxl widzi szerokość B (px / 6) i wysokość wiersza 8, pozostałe kolumny bez zmian,
//      bez NOWYCH błędów schematu;
//   4. przełączenie arkusza i powrót: kolumna dalej szeroka;
//   5. Cofnij ×2 = wymiary i plik jak na początku; Ponów przywraca;
//   6. zapisany plik otwarty ponownie: kolumna B ma przeciągniętą szerokość;
//   7. bez błędów strony.

const { chromium, webkit } = require("playwright");
const fs = require("fs");
const os = require("os");
const path = require("path");
const { execFileSync } = require("child_process");

const APP_URL = process.env.APP_URL || "http://127.0.0.1:4175/";
const ENGINE_NAME = process.env.ENGINE === "webkit" ? "webkit" : "chromium";
const ENGINE = ENGINE_NAME === "webkit" ? webkit : chromium;
const PY = path.join(__dirname, "..", ".venv", "bin", "python");
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
let ooxml = null;
try { ooxml = require("./ooxml-validate"); } catch { ooxml = null; }
const ooxmlOn = () => !!(ooxml && ooxml.available && ooxml.available());
const L = require("../app/xlsx-layout.js");

async function run() {
  const fails = [];
  const check = (name, ok, detail) => {
    console.log(`${ok ? "✅" : "❌"} ${name}${ok ? "" : ` — ${JSON.stringify(detail)}`}`);
    if (!ok) fails.push(name);
  };

  // 0. sam zapis układu
  const x0 = '<worksheet><sheetFormatPr defaultRowHeight="15"/><cols><col min="1" max="5" width="12" style="3" customWidth="1"/></cols><sheetData><row r="2" spans="1:3"><c r="A2"/></row><row r="9"/></sheetData></worksheet>';
  const x1 = L.patchSheetLayout(x0, { cols: { 2: 20.5, 7: 9 }, rows: { 1: 30, 4: 22.5, 0: 18 } });
  check("0. zakres A:E dzielony, styl zostaje", x1.includes('<col min="1" max="2" width="12" style="3" customWidth="1"/><col min="3" max="3" width="20.5" style="3" customWidth="1"/><col min="4" max="5" width="12" style="3" customWidth="1"/><col min="8" max="8" width="9" customWidth="1"/>'), x1);
  check("0. wiersze: istniejący dostaje ht, nowe w kolejności", /<row r="1" ht="18" customHeight="1"\/><row r="2" spans="1:3" ht="30" customHeight="1"><c r="A2"\/><\/row><row r="5" ht="22.5" customHeight="1"\/><row r="9"\/>/.test(x1), x1);
  const x2 = L.patchSheetLayout('<worksheet><dimension ref="A1"/><sheetData/><pageMargins/></worksheet>', { cols: { 0: L.pxToWidth(125) }, rows: { 2: 30 } });
  check("0. pusty arkusz: <cols> przed <sheetData>, wiersz w środku", x2 === '<worksheet><dimension ref="A1"/><cols><col min="1" max="1" width="20.83203125" customWidth="1"/></cols><sheetData><row r="3" ht="30" customHeight="1"/></sheetData><pageMargins/></worksheet>', x2);
  check("0. px → szerokość jak Excel Mac (185 → 30.83203125)", L.pxToWidth(185) === 30.83203125, L.pxToWidth(185));

  const browser = await ENGINE.launch({ headless: true });
  const ctx = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1280, height: 900 } });
  await ctx.addInitScript(() => { localStorage.setItem("introPlayed", "true"); });
  const p = await ctx.newPage();
  const errors = [];
  p.on("pageerror", (e) => errors.push(e.message));
  await p.goto(APP_URL, { waitUntil: "load" });
  await p.evaluate(() => { document.getElementById("heroSplash")?.remove(); return loadSampleFile(); });
  await p.waitForFunction(() => document.querySelector("#dataTable tbody tr[data-row-key]"), null, { timeout: 20000 });
  await sleep(600);
  const before = path.join(os.tmpdir(), `layout-before-${ENGINE_NAME}.xlsx`);
  fs.writeFileSync(before, Buffer.from(await p.evaluate(async () => Array.from(await buildOutputBytes("xlsx")))));

  // uchwyt kolumny: lewa połowa (prawą przykrywa następny nagłówek), wiersza: górna krawędź
  const dragCol = async (dx) => {
    const r = await p.locator('#dataTable thead tr.guide-row .col-resizer[data-col-index="1"]').boundingBox();
    const x = r.x + 3, y = r.y + r.height / 2;
    await p.mouse.move(x, y); await p.mouse.down(); await p.mouse.move(x + dx / 2, y); await p.mouse.move(x + dx, y); await p.mouse.up();
    await sleep(400);
  };
  const dragRow = async (dy) => {
    const r = await p.locator('#dataTable tbody tr[data-row-index="7"] .row-resizer').boundingBox();
    const x = r.x + r.width / 2, y = r.y + 1;
    await p.mouse.move(x, y); await p.mouse.down(); await p.mouse.move(x, y + dy / 2); await p.mouse.move(x, y + dy); await p.mouse.up();
    await sleep(400);
  };
  const st = () => p.evaluate(() => ({ n: swbLayoutEdit.count(), dirty: hasUnsavedChanges, colPx: Math.round(document.querySelector('#dataTable thead tr.guide-row .col-resizer[data-col-index="1"]').parentElement.getBoundingClientRect().width), pl: JSON.parse(JSON.stringify(pendingLayout)) }));

  // 2. Podgląd
  await p.evaluate(() => sheetEditUi.setMode("view"));
  await dragCol(40);
  let s = await st();
  check("2. Podgląd: tylko widok, nic do zapisu", s.n === 0 && !s.dirty, s);
  await p.evaluate(() => { manualColumnWidths = {}; renderActiveTable(); sheetEditUi.setMode("edit"); document.querySelectorAll(".toast").forEach((t) => t.remove()); });
  await sleep(300);

  // 1. Edycja
  const startPx = (await st()).colPx;
  await dragCol(60);
  await dragRow(25);
  s = await st();
  const colW = s.pl["Obieg zadań"] && s.pl["Obieg zadań"].cols[1];
  const rowH = s.pl["Obieg zadań"] && s.pl["Obieg zadań"].rows[7];
  check("1. dwie zmiany do zapisu, niezapisany", s.n === 2 && s.dirty && colW > 0 && rowH > 0, s);
  check("1. szerokość = px / 6 (jak Excel Mac)", Math.abs(colW * 6 - s.colPx) < 1.5, { colW, px: s.colPx });

  // 3. zapis
  const out = path.join(os.tmpdir(), `layout-${ENGINE_NAME}.xlsx`);
  fs.writeFileSync(out, Buffer.from(await p.evaluate(async () => Array.from(await buildOutputBytes("xlsx")))));
  if (process.env.EXCEL_DIR) { fs.mkdirSync(process.env.EXCEL_DIR, { recursive: true }); fs.copyFileSync(out, path.join(process.env.EXCEL_DIR, `layout-${ENGINE_NAME}.xlsx`)); }
  if (fs.existsSync(PY)) {
    const py = `import openpyxl,sys,json
a=openpyxl.load_workbook(sys.argv[1])['Obieg zadań']; b=openpyxl.load_workbook(sys.argv[2])['Obieg zadań']
print(json.dumps({"B":b.column_dimensions['B'].width,"C":[a.column_dimensions['C'].width,b.column_dimensions['C'].width],"r8":b.row_dimensions[8].height,"r9":[a.row_dimensions[9].height,b.row_dimensions[9].height]}))`;
    let o;
    try { o = JSON.parse(execFileSync(PY, ["-I", "-c", py, before, out], { encoding: "utf8" })); } catch (e) { o = { err: String(e.stderr || e.message).slice(0, 300) }; }
    check("3. openpyxl: szerokość B i wysokość wiersza 8, sąsiedzi bez zmian", o.B === colW && o.r8 === rowH && o.C[0] === o.C[1] && o.r9[0] === o.r9[1], o);
  }
  if (ooxmlOn()) {
    const r = ooxml.validate([before, out]);
    const fresh = ooxml.newErrors(r[before] || [], r[out] || []);
    check("3. bez NOWYCH błędów schematu", fresh.length === 0, fresh.slice(0, 3));
  }

  // 4. inny arkusz i z powrotem
  await p.evaluate(() => { sheetSelect.value = workbook.SheetNames[1]; loadBtn.click(); });
  await sleep(1200);
  await p.evaluate(() => { sheetSelect.value = workbook.SheetNames[0]; loadBtn.click(); });
  await sleep(1500);
  s = await st();
  check("4. po powrocie na arkusz kolumna dalej szeroka, plik dalej niezapisany", Math.abs(s.colPx - colW * 6) < 2 && s.dirty, { s, colW });

  // 5. Cofnij / Ponów
  await p.evaluate(() => { undoLast(); undoLast(); });
  await sleep(400);
  s = await st();
  check("5. Cofnij ×2: nic do zapisu, kolumna jak na początku", s.n === 0 && Math.abs(s.colPx - startPx) < 2, { s, startPx });
  await p.evaluate(() => { redoLast(); redoLast(); });
  await sleep(400);
  s = await st();
  check("5. Ponów ×2: obie zmiany wracają", s.n === 2, s);

  // 6. ponowne otwarcie zapisanego pliku
  await p.evaluate(() => { hasUnsavedChanges = false; });
  await p.setInputFiles("#fileInput", out);
  await p.waitForFunction(() => currentFileName && !/Przyklad/.test(currentFileName) || document.querySelector("#dataTable tbody tr[data-row-key]"), null, { timeout: 20000 });
  await sleep(1500);
  s = await st();
  check("6. ponownie otwarty plik: kolumna B ma zapisaną szerokość", Math.abs(s.colPx - colW * 6) < 2, { s, colW });

  check("7. bez błędów strony", errors.length === 0, errors.slice(0, 3));
  await browser.close();
  console.log(fails.length ? `\n❌ ${fails.length} nie przeszło` : "\n✅ wszystko przeszło");
  process.exit(fails.length ? 1 : 0);
}

run().catch((e) => { console.error(e); process.exit(1); });
