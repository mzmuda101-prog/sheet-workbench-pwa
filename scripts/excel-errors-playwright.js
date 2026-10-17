// excel-errors-playwright.js — nazwy błędów jak w polskim Excelu (2026-10-17).
//
// Wzorzec: Microsoft Excel na Macu przez AppleScript (`formula local`):
//   - widok: #DZIEL/0!, #N/D, #ARG!, #ZERO!, #LICZBA!, #ADR!, #NAZWA? (plik trzyma #DIV/0!…),
//   - wpis „#DZIEL/0!” / „#n/d” (dowolna wielkość liter) = BŁĄD (CZY.TEKST = FAŁSZ),
//     angielskie „#DIV/0!” = zwykły TEKST, „#OBL!” / „#ROZLANIE!” = tekst.
// Co musi być prawdą:
//   1. błędy z pliku widać po polsku (wątek główny i wątek roboczy), cell.w zostaje angielskie,
//   2. edytor pokazuje „#DZIEL/0!” (nie kod liczbowy SheetJS), Enter bez zmian nic nie psuje,
//   3. wpisane polskie nazwy → błąd, angielska nazwa / #OBL! → tekst; formuły widzą błąd,
//   4. wynik formuły =1/0 i =A2 po polsku,
//   5. zapis: openpyxl widzi błędy (data_type „e”, angielskie kody), w XML nie ma polskich nazw,
//   6. tryb EN: angielskie nazwy w widoku, wpis „#DIV/0!” = błąd, „#DZIEL/0!” = tekst.

const pw = require("playwright");
const path = require("path");
const fs = require("fs");
const os = require("os");
const { spawnSync } = require("child_process");
const VENV_PY = path.join(__dirname, "..", ".venv/bin/python");
const PY = fs.existsSync(VENV_PY) ? VENV_PY : "python3";

const APP_URL = process.env.APP_URL || "http://127.0.0.1:4175/";
const ENGINE = process.env.ENGINE || "chromium";
const SLEEP_SCALE = Math.min(1, Math.max(0.1, Number(process.env.SLEEP_SCALE || 1)));
const sleep = (ms) => new Promise((r) => setTimeout(r, Math.round(ms * SLEEP_SCALE)));

const tmp = fs.mkdtempSync(path.join(os.tmpdir(), "excel-errors-"));
const FILE = path.join(tmp, "bledy.xlsx");

const EN = ["#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A"];
const PL = ["#ZERO!", "#DZIEL/0!", "#ARG!", "#ADR!", "#NAZWA?", "#LICZBA!", "#N/D"];

// openpyxl zapisuje tekst „#DIV/0!” z listy kodów jako komórkę błędu (t="e")
const gen = spawnSync(PY, ["-c", `
import sys
from openpyxl import Workbook
wb = Workbook(); ws = wb.active; ws.title = "Bledy"
ws.append(["Plik", "Wpis", "Formula", "Opis"])
for i, e in enumerate(${JSON.stringify(EN)}):
    ws.append([e, None, None, "w%d" % (i + 1)])
wb.save(sys.argv[1])
`, FILE], { encoding: "utf8" });
if (gen.status !== 0) { console.error("❌ excel-errors: fixture —", gen.stderr); process.exit(1); }

async function openApp(browser, lang) {
  const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1400, height: 900 } });
  await context.addInitScript((lang) => {
    localStorage.setItem("introPlayed", "true");
    localStorage.setItem("excel-workbench-lang", lang);
  }, lang);
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  page.on("dialog", (d) => d.accept());
  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(() => { document.getElementById("heroSplash")?.remove(); return ensureXlsxLibs(false); });
  await page.setInputFiles("#fileInput", FILE);
  await page.waitForFunction(() => currentFileName === "bledy.xlsx" && currentSheetName === "Bledy" && !!document.querySelector("#dataTable tbody tr[data-row-key]"), null, { timeout: 20000 });
  await sleep(400);
  return { context, page, errors };
}

async function run() {
  const browser = await pw[ENGINE].launch({ headless: true });
  const failures = [];
  const ok = (name, cond, got) => { if (!cond) failures.push(`${name} — dostałem: ${JSON.stringify(got)}`); };

  const { context, page, errors } = await openApp(browser, "pl");
  const td = (r, c) => `#dataTable tbody tr[data-row-key="wide:${r}"] td[data-col-index="${c}"]`;
  const text = (r, c) => page.evaluate((sel) => (document.querySelector(sel)?.textContent ?? "").trim(), td(r, c));
  const colA = () => Promise.all([1, 2, 3, 4, 5, 6, 7].map((r) => text(r, 0)));

  // 1. widok z pliku
  const shownA = await colA();
  ok("1. błędy z pliku po polsku (tabela)", JSON.stringify(shownA) === JSON.stringify(PL), shownA);
  const fileW = await page.evaluate(() => ["A2", "A3", "A8"].map((ref) => workbook.Sheets.Bledy[ref].w));
  ok("1. cell.w zostaje angielskie (formuły, zapis)", fileW.join() === "#NULL!,#DIV/0!,#N/A", fileW);
  const viaWorker = await page.evaluate(async () => {
    const res = await buildRowsViaWorker(workbook.Sheets.Bledy, 1, workbook);
    return res.rows.map((row) => row.display[0]);
  });
  ok("1. wątek roboczy (duże arkusze) też po polsku", JSON.stringify(viaWorker) === JSON.stringify(PL), viaWorker);

  // 2. edytor: nazwa jak w widoku, Enter bez zmian = bez zapisu
  await page.dblclick(td(2, 0));
  await page.waitForSelector(`${td(2, 0)} input.cell-editor`, { timeout: 3000 });
  const editorVal = await page.$eval(`${td(2, 0)} input.cell-editor`, (el) => el.value);
  ok("2. edytor na błędzie z pliku pokazuje „#DZIEL/0!”", editorVal === "#DZIEL/0!", editorVal);
  await page.keyboard.press("Enter");
  await sleep(200);
  const afterEnter = await page.evaluate(() => ({ c: workbook.Sheets.Bledy.A3, p: (pendingEdits.Bledy || {}).A3 || null }));
  ok("2. Enter bez zmian: A3 dalej błędem, bez wpisu do zapisu", afterEnter.c.t === "e" && afterEnter.c.w === "#DIV/0!" && !afterEnter.p, afterEnter);
  ok("2. A3 dalej „#DZIEL/0!” w tabeli", (await text(2, 0)) === "#DZIEL/0!", await text(2, 0));

  // 3. wpisy
  const typeIn = async (r, c, s) => {
    await page.dblclick(td(r, c));
    await page.waitForSelector(`${td(r, c)} input.cell-editor`, { timeout: 3000 });
    await page.keyboard.press("ControlOrMeta+a");
    await page.keyboard.type(s);
    await page.keyboard.press("Enter");
    await sleep(200);
  };
  await typeIn(1, 1, "#n/d");
  await typeIn(2, 1, "#DIV/0!");
  await typeIn(3, 1, "#OBL!");
  await typeIn(4, 1, "#Dziel/0!");
  await typeIn(3, 0, "#DZIEL/0!"); // A4 (#VALUE! z pliku) przepisane ręcznie na inny błąd
  const typed = await page.evaluate(() => ["B2", "B3", "B4", "B5", "A4"].map((ref) => {
    const c = workbook.Sheets.Bledy[ref];
    return { ref, t: c.t, v: c.v, w: c.w ?? null, p: pendingEdits.Bledy[ref] };
  }));
  ok("3. „#n/d” → błąd #N/A", typed[0].t === "e" && typed[0].v === "#N/A" && typed[0].p.t === "e" && typed[0].p.err === "#N/A", typed[0]);
  ok("3. „#DIV/0!” w trybie PL → tekst (jak Excel)", typed[1].t === "s" && typed[1].v === "#DIV/0!", typed[1]);
  ok("3. „#OBL!” → tekst (Excel nie przyjmuje go jako błąd)", typed[2].t === "s" && typed[2].v === "#OBL!", typed[2]);
  ok("3. „#Dziel/0!” → błąd #DIV/0!", typed[3].t === "e" && typed[3].v === "#DIV/0!", typed[3]);
  ok("3. A4 przepisane na „#DZIEL/0!” → błąd #DIV/0!", typed[4].t === "e" && typed[4].w === "#DIV/0!" && typed[4].p.err === "#DIV/0!", typed[4]);
  const shownB = await Promise.all([1, 2, 3, 4].map((r) => text(r, 1)));
  ok("3. tabela po wpisach: #N/D, #DIV/0!, #OBL!, #DZIEL/0!", shownB.join() === "#N/D,#DIV/0!,#OBL!,#DZIEL/0!", shownB);
  ok("3. A4 w tabeli „#DZIEL/0!”", (await text(3, 0)) === "#DZIEL/0!", await text(3, 0));
  await page.dblclick(td(1, 1));
  await page.waitForSelector(`${td(1, 1)} input.cell-editor`, { timeout: 3000 });
  const editorTyped = await page.$eval(`${td(1, 1)} input.cell-editor`, (el) => el.value);
  ok("3. edytor na wpisanym błędzie pokazuje „#N/D”", editorTyped === "#N/D", editorTyped);
  await page.keyboard.press("Escape");
  await sleep(150);

  // 4. formuły
  await typeIn(1, 2, "=1/0");
  await typeIn(2, 2, "=A2");
  await typeIn(3, 2, "=CZY.BŁĄD(B2)");
  await typeIn(4, 2, "=CZY.TEKST(B3)");
  await typeIn(5, 2, "=JEŻELI.BŁĄD(B5;\"zł\")");
  const shownC = await Promise.all([1, 2, 3, 4, 5].map((r) => text(r, 2)));
  ok("4. =1/0 → #DZIEL/0!, =A2 → #ZERO!, CZY.BŁĄD(wpis) = PRAWDA, CZY.TEKST(„#DIV/0!”) = PRAWDA, JEŻELI.BŁĄD działa",
    shownC.join() === "#DZIEL/0!,#ZERO!,PRAWDA,PRAWDA,zł", shownC);
  const fW = await page.evaluate(() => workbook.Sheets.Bledy.C2.w);
  ok("4. wynik formuły w komórce zostaje angielski (#DIV/0!)", fW === "#DIV/0!", fW);

  // 5. zapis + drugi czytnik
  const out = path.join(tmp, "zapis.xlsx");
  const b64 = await page.evaluate(async () => {
    const u8 = await buildOutputBytes("xlsx");
    let s = ""; for (let i = 0; i < u8.length; i += 0x8000) s += String.fromCharCode(...u8.subarray(i, i + 0x8000));
    return btoa(s);
  });
  fs.writeFileSync(out, Buffer.from(b64, "base64"));
  const py = spawnSync(PY, ["-c", `
import json, sys, zipfile
from openpyxl import load_workbook
ws = load_workbook(sys.argv[1])["Bledy"]
cells = {r: {"v": ws[r].value, "t": ws[r].data_type} for r in ["A2", "A3", "A4", "A5", "A8", "B2", "B3", "B4", "B5", "C2"]}
xml = zipfile.ZipFile(sys.argv[1]).read("xl/worksheets/sheet1.xml").decode("utf8")
print(json.dumps({"cells": cells, "pl": [n for n in ${JSON.stringify(PL)} if n in xml]}))
`, out], { encoding: "utf8" });
  let f = null;
  try { f = JSON.parse(py.stdout); } catch { failures.push("5. openpyxl nie przeczytał zapisu: " + (py.stderr || "").slice(0, 300)); }
  if (f) {
    const c = f.cells;
    ok("5. błędy z pliku zostały błędami", c.A2.t === "e" && c.A2.v === "#NULL!" && c.A3.t === "e" && c.A3.v === "#DIV/0!" && c.A5.v === "#REF!" && c.A8.v === "#N/A", c);
    ok("5. A4 przepisane: błąd #DIV/0!", c.A4.t === "e" && c.A4.v === "#DIV/0!", c.A4);
    ok("5. wpisane „#n/d” → t=e #N/A, „#Dziel/0!” → t=e #DIV/0!", c.B2.t === "e" && c.B2.v === "#N/A" && c.B5.t === "e" && c.B5.v === "#DIV/0!", { B2: c.B2, B5: c.B5 });
    ok("5. tekst „#DIV/0!” / „#OBL!” zapisany jako tekst", c.B3.t === "s" && c.B3.v === "#DIV/0!" && c.B4.t === "s" && c.B4.v === "#OBL!", { B3: c.B3, B4: c.B4 });
    ok("5. formuła =1/0 zapisana jako formuła", c.C2.t === "f" && c.C2.v === "=1/0", c.C2);
    ok("5. w XML arkusza nie ma polskich nazw błędów", f.pl.length === 0, f.pl);
  }

  // 6. Cofnij wpisany błąd
  await page.evaluate(() => { for (let i = 0; i < 6; i++) undoLast(); }); // 5 formuł + A4
  await sleep(250);
  const undoneA4 = await page.evaluate(() => ({ t: workbook.Sheets.Bledy.A4.t, w: workbook.Sheets.Bledy.A4.w }));
  ok("6. Cofnij przywraca #VALUE! z pliku w A4", undoneA4.t === "e" && undoneA4.w === "#VALUE!", undoneA4);
  ok("6. A4 w tabeli znowu „#ARG!”", (await text(3, 0)) === "#ARG!", await text(3, 0));
  if (errors.length) failures.push("błędy strony (PL): " + errors.join(" | "));
  await context.close();

  // 7. tryb angielski
  const en = await openApp(browser, "en");
  const shownEn = await Promise.all([1, 2, 3, 4, 5, 6, 7].map((r) => en.page.evaluate((sel) => (document.querySelector(sel)?.textContent ?? "").trim(), td(r, 0))));
  ok("7. EN: angielskie nazwy w tabeli", JSON.stringify(shownEn) === JSON.stringify(EN), shownEn);
  const enWorker = await en.page.evaluate(async () => (await buildRowsViaWorker(workbook.Sheets.Bledy, 1, workbook)).rows.map((row) => row.display[0]));
  ok("7. EN: wątek roboczy po angielsku", JSON.stringify(enWorker) === JSON.stringify(EN), enWorker);
  const enParse = await en.page.evaluate(() => ["#DIV/0!", "#n/a", "#DZIEL/0!"].map((s) => parseInputValue(s)));
  ok("7. EN: „#DIV/0!”/„#n/a” = błąd, „#DZIEL/0!” = tekst",
    enParse[0].type === "error" && enParse[0].value === "#DIV/0!" && enParse[1].type === "error" && enParse[1].value === "#N/A" && enParse[2].type === "string", enParse);
  if (en.errors.length) failures.push("błędy strony (EN): " + en.errors.join(" | "));

  fs.rmSync(tmp, { recursive: true, force: true });
  await browser.close();
  if (failures.length) {
    console.error(`❌ excel-errors (${ENGINE}):\n  - ` + failures.join("\n  - "));
    process.exit(1);
  }
  console.log(`✅ excel-errors (${ENGINE}): błędy po polsku w widoku i edytorze, wpis #DZIEL/0!/#N/D = błąd, #DIV/0! = tekst, formuły, zapis t="e", Cofnij, tryb EN`);
}

run().catch((e) => { console.error(e); process.exit(1); });
