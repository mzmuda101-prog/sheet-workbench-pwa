// calc-column-playwright.js — kolumna obliczeniowa i wyjście z automatu, jak w Excelu (2026-11-12).
//
// Plik z openpyxl: tabela Excela „Dane” A1:D6 (Imię, Kwota, Wynik — pusta, Uwagi — z jedną ręczną wartością).
// Co musi być prawdą:
//   1. formuła wpisana w pustą kolumnę tabeli wypełnia całą kolumnę; dymek z 2 przyciskami,
//   2. 1. Cofnij zdejmuje tylko wypełnienie (formuła w komórce zostaje), 2. cofa formułę,
//   3. zmiana formuły w kolumnie obliczeniowej poprawia całą kolumnę,
//   4. kolumna z ręczną wartością nie jest wypełniana,
//   5. „Cofnij uzupełnienie” w dymku = 1. Cofnij,
//   6. nowy wiersz: formuły z wierszy wyżej + dymek tylko za pierwszym razem w pliku,
//   7. „Nie uzupełniaj automatycznie” wyłącza przełącznik (zapamiętany): nowa kolumna i nowy wiersz bez automatu.

const pw = require("playwright");
const { autoAnswerAppDialogs } = require("./app-dialog-auto");
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

function makeFixture(file) {
  const py = spawnSync(PY, ["-I", "-c", `
import sys
from openpyxl import Workbook
from openpyxl.worksheet.table import Table, TableStyleInfo
wb = Workbook(); ws = wb.active; ws.title = "Arkusz"
ws.append(["Imię", "Kwota", "Wynik", "Uwagi", "Ile"])
for i, (n, k) in enumerate([("Anna", 10), ("Bartek", 20), ("Celina", 30), ("Darek", 40), ("Ewa", 50)]):
    ws.append([n, k, None, "pilne" if i == 2 else None, None])
for r in range(2, 7): ws.cell(r, 5, f"=B{r}+1")
t = Table(displayName="Dane", ref="A1:E6"); t.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
ws.add_table(t)
wb.save(sys.argv[1])
`, file], { encoding: "utf8" });
  if (py.status !== 0) throw new Error("openpyxl: " + py.stderr);
}

async function run() {
  const tmp = fs.mkdtempSync(path.join(os.tmpdir(), "swb-calc-col-"));
  const FILE = path.join(tmp, "calc.xlsx");
  makeFixture(FILE);
  const browser = await pw[ENGINE].launch({ headless: true });
  const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1280, height: 860 } });
  await context.addInitScript(() => { localStorage.setItem("introPlayed", "true"); });
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  await autoAnswerAppDialogs(page);
  const failures = [];
  const ok = (name, cond, got) => { if (!cond) failures.push(`${name} — dostałem: ${JSON.stringify(got)}`); };

  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(() => { document.getElementById("heroSplash")?.remove(); return ensureXlsxLibs(false); });
  await page.setInputFiles("#fileInput", FILE);
  await page.waitForFunction(() => currentHeaderRow === 1 && !!document.querySelector("#dataTable tbody tr[data-row-key]"), null, { timeout: 20000 });
  await page.evaluate(() => { sheetEditUi.setMode("edit"); autoFillFormulasEl.checked = true; syncCellStyleFlags(); });
  await sleep(300);

  const td = (r, c) => `#dataTable tbody tr[data-row-key="wide:${r}"] td[data-col-index="${c}"]`;
  const col = (L) => page.evaluate((L) => [2, 3, 4, 5, 6].map((r) => { const x = workbook.Sheets[currentSheetName][L + r]; return x ? (x.f ? "f:" + x.f + "=" + x.v : x.v) : null; }), L);
  const clearToasts = () => page.evaluate(() => document.querySelectorAll(".toast").forEach((t) => t.remove()));
  const toasts = () => page.evaluate(() => [...document.querySelectorAll(".toast")].map((t) => ({ msg: t.querySelector(".toast-msg").textContent, btns: [...t.querySelectorAll(".toast-action")].map((b) => b.textContent) })));
  const type = async (r, c, text) => {
    await page.evaluate(({ r, c }) => setFocusedCell(`wide:${r}`, c, { scroll: true }), { r, c });
    await page.dblclick(td(r, c));
    await page.waitForSelector(`${td(r, c)} input.cell-editor`, { timeout: 3000 });
    await page.keyboard.press("ControlOrMeta+a");
    await page.keyboard.type(text);
    await page.keyboard.press("Tab");
    await sleep(400);
    await page.keyboard.press("Escape");
    await sleep(150);
  };

  // 1.
  await clearToasts();
  await type(1, 2, "=B2*2");
  const c1 = await col("C");
  const t1 = await toasts();
  ok("1. formuła w pustej kolumnie tabeli wypełnia całą kolumnę", c1.every((x, i) => x === `f:B${i + 2}*2=${(i + 1) * 20}`), c1);
  ok("1. dymek z „Cofnij uzupełnienie” i „Nie uzupełniaj automatycznie”", t1.some((x) => /całą kolumnę/.test(x.msg) && x.btns.length === 2), t1);
  const shown1 = await page.evaluate(() => document.querySelector('#dataTable tbody tr[data-row-key="wide:5"] td[data-col-index="2"]').textContent);
  ok("1. wyniki widać w tabeli (ostatni wiersz)", shown1 === "100", shown1);

  // 2.
  await page.evaluate(() => undoLast()); await sleep(300);
  const c2a = await col("C");
  ok("2. 1. Cofnij: zostaje tylko formuła wpisana ręcznie", c2a[0] && c2a[0].startsWith("f:B2*2") && c2a.slice(1).every((x) => x == null), c2a);
  await page.evaluate(() => undoLast()); await sleep(300);
  const c2b = await col("C");
  ok("2. 2. Cofnij: kolumna pusta", c2b.every((x) => x == null), c2b);
  await page.evaluate(() => { redoLast(); redoLast(); }); await sleep(300);
  const c2c = await col("C");
  ok("2. Ponów ×2: znów cała kolumna", c2c.every((x) => x && x.startsWith("f:")), c2c);

  // 3.
  await clearToasts();
  await type(3, 2, "=B4*3");
  const c3 = await col("C");
  ok("3. zmiana formuły w kolumnie obliczeniowej poprawia całą kolumnę", c3.every((x, i) => x === `f:B${i + 2}*3=${(i + 1) * 30}`), c3);

  // 4.
  await clearToasts();
  await type(1, 3, "=A2&\"!\"");
  const d4 = await col("D");
  ok("4. kolumna z ręczną wartością („pilne”) nie jest wypełniana", d4[0] === "f:A2&\"!\"=Anna!" && d4[2] === "pilne" && d4[1] == null && d4[3] == null, d4);
  ok("4. bez dymka o kolumnie", !(await toasts()).some((x) => /całą kolumnę/.test(x.msg)));

  // 5. „Cofnij uzupełnienie” w dymku
  await clearToasts();
  await type(2, 2, "=B3*4");
  const before5 = await col("C");
  await page.click(".toast .toast-action:not(.toast-action-2)");
  await sleep(300);
  const c5 = await col("C");
  ok("5. „Cofnij uzupełnienie” w dymku cofa tylko wypełnienie", before5.every((x) => x && x.includes("*4")) && c5[1] && c5[1].includes("B3*4") && c5[0].includes("*3") && c5[4].includes("*3"), { before5, c5 });

  // 6. nowy wiersz: formuły z góry (kolumny E i C) + dymek raz
  await clearToasts();
  await page.click("#addRowBtn"); await sleep(300);
  await page.keyboard.type("Franek"); await page.keyboard.press("Tab"); await sleep(400); await page.keyboard.press("Escape"); await sleep(150);
  const r7 = await page.evaluate(() => ["E7", "C7"].map((k) => { const x = workbook.Sheets[currentSheetName][k]; return x && x.f ? x.f : null; }));
  const t6 = await toasts();
  ok("6. nowy wiersz dostaje formuły z wierszy wyżej", r7[0] === "B7+1", r7);
  ok("6. dymek o nowym wierszu (za pierwszym razem) z 2 przyciskami", t6.some((x) => /Nowy wiersz dostał formuły/.test(x.msg) && x.btns.length === 2), t6);
  await clearToasts();
  await page.click("#addRowBtn"); await sleep(300);
  await page.keyboard.type("Gosia"); await page.keyboard.press("Tab"); await sleep(400); await page.keyboard.press("Escape"); await sleep(150);
  ok("6. przy kolejnym nowym wierszu dymka już nie ma", !(await toasts()).some((x) => /Nowy wiersz dostał/.test(x.msg)));

  // 7. „Nie uzupełniaj automatycznie”
  await page.evaluate(() => { const ws = workbook.Sheets[currentSheetName]; ["C2", "C3", "C4", "C5", "C6", "C7", "C8"].forEach((k) => { if (ws[k] && ws[k].f) delete ws[k]; }); });
  await clearToasts();
  await type(1, 2, "=B2*5");
  await page.click(".toast .toast-action-2");
  await sleep(300);
  const st7 = await page.evaluate(() => ({ on: swbAutoFillOn, box: autoFillFormulasEl.checked, saved: JSON.parse(localStorage.getItem(CELL_STYLE_PREFS_KEY) || "{}").autoFill }));
  ok("7. „Nie uzupełniaj automatycznie” wyłącza przełącznik i go zapamiętuje", !st7.on && !st7.box && st7.saved === false, st7);
  await page.evaluate(() => { const ws = workbook.Sheets[currentSheetName]; ["C3", "C4", "C5", "C6", "C7", "C8"].forEach((k) => { if (ws[k] && ws[k].f) delete ws[k]; }); });
  await type(2, 2, "=B3*6");
  const c7 = await col("C");
  ok("7. przy wyłączonym automacie formuła zostaje w jednej komórce", c7[1] && c7[1].includes("B3*6") && c7[2] == null && c7[3] == null, c7);
  await page.click("#addRowBtn"); await sleep(300);
  await page.keyboard.type("Hania"); await page.keyboard.press("Tab"); await sleep(400); await page.keyboard.press("Escape"); await sleep(150);
  const r9 = await page.evaluate(() => { const x = workbook.Sheets[currentSheetName].E9; return x && x.f ? x.f : null; });
  ok("7. przy wyłączonym automacie nowy wiersz bez formuł z góry", r9 === null, r9);
  await page.evaluate(() => { autoFillFormulasEl.checked = true; syncCellStyleFlags(); saveCellStylePreferences(); });

  await browser.close();
  fs.rmSync(tmp, { recursive: true, force: true });
  if (errors.length) failures.push("błędy strony: " + errors.join(" | "));
  if (failures.length) {
    console.error(`❌ calc-column (${ENGINE}):\n  - ` + failures.join("\n  - "));
    process.exit(1);
  }
  console.log(`✅ calc-column (${ENGINE}): kolumna obliczeniowa, Cofnij w 2 krokach, dymek z 2 przyciskami, przełącznik automatu`);
}

run().catch((e) => { console.error(e); process.exit(1); });
