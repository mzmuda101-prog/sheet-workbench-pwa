// cell-format-playwright.js — edycja „jak w Excelu”: tryb Edycja/Podgląd, pasek „Narzędzia
// główne” (format → plik), pasek formuły, pole nazwy, Delete, malarz formatów.
//
// Co musi być prawdą:
//   1. domyślnie Edycja: po wczytaniu widać pasek formatu i formuły; pole nazwy = adres
//      klikniętej komórki, pasek formuły = jej zawartość,
//   2. B na zakresie 2×2 pogrubia wszystkie 4 komórki (widać w tabeli), drugi B zdejmuje,
//      Ctrl/⌘+B działa jak przycisk; licznik „Zapisz” liczy zmienione komórki,
//   3. wypełnienie z palety daje pełny kolor tła (nie przygaszony), „Brak wypełnienia” zdejmuje,
//   4. format „Waluta zł” zmienia WYGLĄD liczby (wartość ta sama), „→,00” dokłada miejsce,
//   5. Ctrl+Z cofa format (styl i tekst komórki), Ctrl+Shift+Z ponawia,
//   6. Delete czyści zawartość zaznaczenia, a format zostaje (pendingEdits: clear + fmt),
//   7. przebudowa danych (zmiana wiersza nagłówka / ponowne wczytanie arkusza) nie gubi formatu,
//   8. pasek formuły: nowa wartość + Enter zapisuje komórkę i schodzi wiersz niżej,
//   9. pole nazwy: „B12” + Enter przenosi na B12; wiersz ukryty filtrem → komunikat,
//  10. malarz formatów przenosi format na kliknięty zakres,
//  11. Podgląd: brak pasków, dwuklik nie otwiera edytora, Ctrl+B i Delete nic nie zmieniają;
//      tryb pamiętany po przeładowaniu; powrót do Edycji odblokowuje,
//  12. zapis → openpyxl widzi pogrubienie, wypełnienie, format liczby i obramowanie; ponowne
//      wczytanie ZAPISANEGO pliku w aplikacji rysuje je z pliku (bez sesji).
//
// Uruchom z serwerem na APP_URL (domyślnie http://127.0.0.1:4175/). ENGINE=webkit — Safari.

const { chromium, webkit } = require("playwright");
const path = require("path");
const fs = require("fs");
const os = require("os");
const { spawnSync } = require("child_process");

const APP_URL = process.env.APP_URL || "http://127.0.0.1:4175/";
const FILE = path.join(__dirname, "stress-test-workbench.xlsx");
const ENGINE = process.env.ENGINE === "webkit" ? webkit : chromium;
const SLEEP_SCALE = Math.min(1, Math.max(0.1, Number(process.env.SLEEP_SCALE || 1)));
const sleep = (ms) => new Promise((r) => setTimeout(r, Math.round(ms * SLEEP_SCALE)));
const VENV_PY = path.join(__dirname, "..", ".venv/bin/python");
const PY = fs.existsSync(VENV_PY) ? VENV_PY : "python3";
const MOD = process.platform === "darwin" ? "Meta" : "Control";

async function run() {
  const browser = await ENGINE.launch({ headless: true });
  const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1360, height: 900 }, acceptDownloads: true });
  await context.addInitScript(() => localStorage.setItem("introPlayed", "true"));
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push("pageerror: " + e.message));
  page.on("console", (m) => { if (m.type() === "error") errors.push("console: " + m.text()); });
  const failures = [];
  const ok = (name, cond, got) => { if (!cond) failures.push(`${name} — dostałem: ${JSON.stringify(got)}`); };

  async function openFile(file) {
    await page.goto(APP_URL, { waitUntil: "load" });
    await page.evaluate(() => { document.getElementById("heroSplash")?.remove(); try { ensureXlsxLibs(false); } catch {} });
    await page.setInputFiles("#fileInput", file);
    await page.waitForFunction(() => document.getElementById("sheetSelect")?.options?.length > 0, null, { timeout: 20000 });
    await sleep(300);
    await page.evaluate(() => loadBtn.click());
    await page.waitForFunction(() => currentHeaders.length > 0 && document.querySelector("#dataTable tbody tr[data-row-key]"), null, { timeout: 20000 });
    await sleep(500);
    await page.evaluate(() => { if (typeof setSidebarOpen === "function") setSidebarOpen(false); });
    await sleep(200);
  }
  await openFile(FILE);

  const td = (r, c) => `tbody tr[data-row-key="wide:${r}"] td[data-col-index="${c}"]`;
  const select = (r1, c1, r2, c2) => page.evaluate(({ r1, c1, r2, c2 }) => {
    setFocusedCell(`wide:${r1}`, c1, { scroll: true });
    if (r2 != null) setSelectedCell(`wide:${r2}`, c2, { scroll: false });
    findCellElement(focusedCellState)?.focus({ preventScroll: true });
  }, { r1, c1, r2, c2 });
  const cellView = (r, c) => page.evaluate(({ sel, r, c }) => {
    const el = document.querySelector(sel);
    const ref = XLSX.utils.encode_cell({ r, c: currentStartCol + c });
    const cs = el ? getComputedStyle(el) : null;
    return {
      ref,
      text: el ? el.textContent : null,
      weight: cs ? cs.fontWeight : null,
      italic: cs ? cs.fontStyle : null,
      bg: el ? el.style.background || el.style.backgroundColor : null,
      align: cs ? cs.textAlign : null,
      borderBottom: el ? el.style.borderBottom : null,
      pending: pendingEdits[currentSheetName]?.[ref] ?? undefined,
      value: workbook.Sheets[currentSheetName][ref]?.v ?? null,
    };
  }, { sel: td(r, c), r, c });
  const saveCount = () => page.evaluate(() => document.getElementById("heroSaveCount")?.textContent || "");

  // ── 1. rama ───────────────────────────────────────────────────────────────
  const frame = await page.evaluate(() => ({
    mode: sheetEditUi.mode,
    bars: !document.getElementById("editBars").hidden,
    active: document.querySelector("#sheetModeSwitch .mode-btn.is-active")?.dataset.mode,
  }));
  ok("1. domyślnie Edycja + paski widoczne", frame.mode === "edit" && frame.bars && frame.active === "edit", frame);
  await page.click(td(3, 1));
  await sleep(150);
  const fx = await page.evaluate(() => ({ name: nameBox.value, fx: formulaInput.value, cell: getDisplayValue(baseRows.find((r) => r.rowIndex0 === 3), 1) }));
  ok("1. pole nazwy B4 + pasek formuły = zawartość", fx.name === "B4" && fx.fx === String(fx.cell), fx);

  // ── 2. pogrubienie na zakresie ───────────────────────────────────────────
  await select(2, 1, 3, 2); // B3:C4
  await page.click('#formatRibbon [data-fmt="b"]');
  await sleep(200);
  const bold = await Promise.all([[2, 1], [2, 2], [3, 1], [3, 2]].map(([r, c]) => cellView(r, c)));
  ok("2. B pogrubia 4 komórki zakresu", bold.every((v) => Number(v.weight) >= 700), bold.map((v) => v.weight));
  ok("2. pendingEdits ma fmt.b bez wartości", bold.every((v) => v.pending && v.pending.fmt && v.pending.fmt.b === true && !("t" in v.pending)), bold.map((v) => v.pending));
  ok("2. przycisk B wciśnięty", await page.getAttribute('#formatRibbon [data-fmt="b"]', "aria-pressed") === "true", null);
  ok("2. licznik Zapisz = 4", (await saveCount()) === "4", await saveCount());
  await select(2, 1, 3, 2);
  await page.keyboard.press(`${MOD}+b`);
  await sleep(200);
  const unbold = await cellView(2, 1);
  ok("2. Ctrl/⌘+B zdejmuje pogrubienie", Number(unbold.weight) < 700 && unbold.pending.fmt.b === false, unbold);

  // ── 3. wypełnienie ─────────────────────────────────────────────────────────
  await select(4, 1);
  await page.click("#fmtFillMore");
  await page.waitForSelector("#fmtPop:not([hidden]) .fmt-pop-swatch");
  await page.click('#fmtPop .fmt-pop-swatch[aria-label="#FFFF00"]');
  await sleep(200);
  const filled = await cellView(4, 1);
  ok("3. żółte wypełnienie pełnym kolorem", /rgb\(255, 255, 0\)|rgba\(255, 255, 0, 1\)/.test(filled.bg), filled.bg);
  await page.click("#fmtFillMore");
  await page.click("#fmtPop .fmt-pop-auto");
  await sleep(200);
  const unfilled = await cellView(4, 1);
  ok("3. „Brak wypełnienia” zdejmuje tło", !unfilled.bg && unfilled.pending.fmt.fill === null, unfilled);

  // ── 4. format liczb ────────────────────────────────────────────────────────
  const numCol = await page.evaluate(() => currentHeaders.indexOf("Kw1_Kwota"));
  await select(1, numCol);
  const numBefore = await cellView(1, numCol);
  await page.selectOption("#fmtNumFormat", "pln");
  await sleep(200);
  const pln = await cellView(1, numCol);
  // jak w polskim Excelu: „9 885,66 zł” (SSF daje „9,885.66 zł” — widok zamienia separatory)
  const expected = await page.evaluate((v) => swbLocalizeCellText(XLSX.SSF.format('#,##0.00 "zł"', v), { t: "n", z: '#,##0.00 "zł"' }, "pl"), numBefore.value);
  ok("4. Waluta zł zmienia wygląd", pln.text === expected && / zł$/.test(pln.text), { pln: pln.text, expected });
  ok("4. wartość bez zmian", pln.value === numBefore.value, { before: numBefore.value, after: pln.value });
  await page.click("#fmtDecMore");
  await sleep(150);
  const dec3 = await cellView(1, numCol);
  ok("4. →,00 = 3 miejsca", dec3.pending.fmt.numFmt === '#,##0.000 "zł"', dec3.pending.fmt);
  await page.click("#fmtDecLess");
  await page.click("#fmtDecLess");
  await sleep(150);
  ok("4. ,0← dwa razy = 1 miejsce", (await cellView(1, numCol)).pending.fmt.numFmt === '#,##0.0 "zł"', null);

  // ── 5. cofanie formatu ─────────────────────────────────────────────────────
  await select(6, 1);
  await page.click('#formatRibbon [data-fmt="i"]');
  await sleep(150);
  ok("5. kursywa nałożona", (await cellView(6, 1)).italic === "italic", null);
  await page.evaluate(() => findCellElement(focusedCellState)?.focus());
  await page.keyboard.press(`${MOD}+z`);
  await sleep(200);
  const undone = await cellView(6, 1);
  ok("5. Ctrl+Z cofa kursywę i wpis do zapisu", undone.italic !== "italic" && undone.pending === undefined, undone);
  await page.keyboard.press(`${MOD}+Shift+z`);
  await sleep(200);
  ok("5. Ctrl+Shift+Z ponawia", (await cellView(6, 1)).italic === "italic", null);

  // ── 6. Delete: zawartość znika, format zostaje ─────────────────────────────
  await select(7, 1);
  await page.click('#formatRibbon [data-fmt="b"]');
  await sleep(100);
  await page.keyboard.press("Delete");
  await sleep(200);
  const cleared = await cellView(7, 1);
  ok("6. Delete czyści wartość", cleared.text === "" && cleared.value === null, cleared);
  ok("6. format zostaje (clear + fmt)", cleared.pending && cleared.pending.clear === true && cleared.pending.fmt.b === true && Number(cleared.weight) >= 700, cleared.pending);

  // ── 7. przebudowa danych nie gubi formatu ──────────────────────────────────
  await page.evaluate(() => rebuildCurrentSheetData());
  await sleep(300);
  const rebuilt = await cellView(6, 1);
  ok("7. po przebudowie kursywa dalej jest", rebuilt.italic === "italic", rebuilt);
  ok("7. po przebudowie format liczby dalej jest", (await cellView(1, numCol)).text.endsWith("zł"), null);

  // ── 8. pasek formuły ───────────────────────────────────────────────────────
  await select(10, 1);
  await page.click("#formulaInput");
  await page.fill("#formulaInput", "Zmienione z paska");
  await page.keyboard.press("Enter");
  await sleep(200);
  const viaFx = await cellView(10, 1);
  const focusAfter = await page.evaluate(() => focusedCellState && focusedCellState.rowKey);
  ok("8. pasek formuły zapisuje komórkę", viaFx.value === "Zmienione z paska" && viaFx.text === "Zmienione z paska", viaFx);
  ok("8. Enter schodzi wiersz niżej", focusAfter === "wide:11", focusAfter);

  // ── 9. pole nazwy ──────────────────────────────────────────────────────────
  await page.click("#nameBox");
  await page.fill("#nameBox", "C20");
  await page.keyboard.press("Enter");
  await sleep(200);
  const jumped = await page.evaluate(() => ({ f: focusedCellState, name: nameBox.value }));
  ok("9. C20 + Enter = komórka C20", jumped.f && jumped.f.rowKey === "wide:19" && jumped.f.colIndex0 === 2 && jumped.name === "C20", jumped);
  await page.click("#nameBox");
  await page.fill("#nameBox", "nie-adres");
  await page.keyboard.press("Enter");
  await sleep(150);
  ok("9. zły adres = komunikat, fokus bez zmian", await page.evaluate(() => focusedCellState.rowKey === "wide:19"), null);

  // ── 10. malarz formatów ────────────────────────────────────────────────────
  await select(12, 1);
  await page.click('#formatRibbon [data-fmt="b"]');
  await page.click('#formatRibbon [data-align="center"]');
  await sleep(100);
  await page.click("#fmtPainterBtn");
  await sleep(100);
  await page.click(td(14, 2));
  await sleep(250);
  const painted = await cellView(14, 2);
  ok("10. malarz: pogrubienie + wyśrodkowanie", Number(painted.weight) >= 700 && painted.align === "center", painted);
  ok("10. malarz wyłącza się po użyciu", await page.getAttribute("#fmtPainterBtn", "aria-pressed") === "false", null);

  // obramowanie na zakres (sprawdzane potem w pliku)
  await select(16, 1, 17, 2);
  await page.click("#fmtBorderBtn");
  await page.click('#fmtPop .fmt-pop-border:nth-child(1)'); // wszystkie krawędzie
  await sleep(200);
  ok("10b. obramowanie widać w tabeli", !!(await cellView(16, 1)).borderBottom, await cellView(16, 1));

  // ── 11. Podgląd ────────────────────────────────────────────────────────────
  await page.click('#sheetModeSwitch [data-mode="view"]');
  await sleep(250);
  const view = await page.evaluate(() => ({ bars: !document.getElementById("editBars").hidden, cls: document.documentElement.classList.contains("swb-view-mode") }));
  ok("11. Podgląd chowa paski", !view.bars && view.cls, view);
  const countBefore = await saveCount();
  await page.dblclick(td(20, 1));
  await sleep(200);
  ok("11. dwuklik nie otwiera edytora", !(await page.$("input.cell-editor")), null);
  await select(20, 1);
  await page.keyboard.press(`${MOD}+b`);
  await page.keyboard.press("Delete");
  await sleep(200);
  const locked = await cellView(20, 1);
  ok("11. Ctrl+B / Delete nic nie zmieniają", locked.pending === undefined && locked.text !== "" && (await saveCount()) === countBefore, locked);
  await page.reload({ waitUntil: "load" });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  ok("11. tryb pamiętany po przeładowaniu", await page.evaluate(() => sheetEditUi.mode === "view"), null);
  await page.evaluate(() => sheetEditUi.setMode("edit"));
  // wróć do pliku z edycjami (szkic) — tu prościej: zapis robimy na świeżo przygotowanych edycjach poniżej
  await openFile(FILE);

  // ── 12. zapis → openpyxl + ponowne wczytanie ───────────────────────────────
  await select(2, 1);
  await page.click('#formatRibbon [data-fmt="b"]');
  await page.click("#fmtFillMore");
  await page.click('#fmtPop .fmt-pop-swatch[aria-label="#FFFF00"]');
  await select(1, numCol);
  await page.selectOption("#fmtNumFormat", "pln");
  await select(4, 1, 5, 2);
  await page.click("#fmtBorderBtn");
  await page.click('#fmtPop .fmt-pop-border:nth-child(2)'); // zewnętrzne
  await sleep(200);
  const tmp = fs.mkdtempSync(path.join(os.tmpdir(), "cell-format-"));
  const out = path.join(tmp, "zapis.xlsx");
  const b64 = await page.evaluate(async () => {
    const u8 = await buildOutputBytes("xlsx");
    let s = ""; for (let i = 0; i < u8.length; i += 0x8000) s += String.fromCharCode(...u8.subarray(i, i + 0x8000));
    return btoa(s);
  });
  fs.writeFileSync(out, Buffer.from(b64, "base64"));
  const sheetName = await page.evaluate(() => currentSheetName);
  const numRef = await page.evaluate((c) => XLSX.utils.encode_cell({ r: 1, c: currentStartCol + c }), numCol);
  const py = spawnSync(PY, ["-c", `
import json, sys
from openpyxl import load_workbook
ws = load_workbook(sys.argv[1])[sys.argv[2]]
def side(s): return s.style
out = {
  "b3_bold": ws["B3"].font.b, "b3_fill": ws["B3"].fill.fgColor.rgb if ws["B3"].fill.fill_type else None,
  "num_fmt": ws[sys.argv[3]].number_format,
  "b5": [side(ws["B5"].border.top), side(ws["B5"].border.left), side(ws["B5"].border.right), side(ws["B5"].border.bottom)],
  "c6": [side(ws["C6"].border.top), side(ws["C6"].border.left), side(ws["C6"].border.right), side(ws["C6"].border.bottom)],
}
print(json.dumps(out))
`, out, sheetName, numRef], { encoding: "utf8" });
  let pyOut = null;
  try { pyOut = JSON.parse(py.stdout); } catch { failures.push("12. openpyxl nie przeczytał zapisu: " + py.stderr.slice(0, 300)); }
  if (pyOut) {
    ok("12. plik: B3 pogrubione + żółte", pyOut.b3_bold === true && pyOut.b3_fill === "FFFFFF00", pyOut);
    ok("12. plik: format zł", pyOut.num_fmt === '#,##0.00 "zł"', pyOut.num_fmt);
    ok("12. plik: krawędzie zewnętrzne zakresu B5:C6", JSON.stringify(pyOut.b5) === JSON.stringify(["thin", "thin", null, null])
      && JSON.stringify(pyOut.c6) === JSON.stringify([null, null, "thin", "thin"]), pyOut);
  }
  // ponowne wczytanie zapisanego pliku — formatowanie rysowane z pliku
  await openFile(out);
  const reread = await cellView(2, 1);
  const rereadNum = await cellView(1, numCol);
  ok("12. po ponownym wczytaniu B3 pogrubione z pliku", Number(reread.weight) >= 700 && reread.pending === undefined, reread);
  ok("12. po ponownym wczytaniu liczba w zł", rereadNum.text.endsWith("zł"), rereadNum.text);
  fs.rmSync(tmp, { recursive: true, force: true });

  await browser.close();
  if (errors.length) failures.push("błędy strony: " + errors.join(" | "));
  if (failures.length) {
    console.error(`❌ cell-format (${process.env.ENGINE || "chromium"}):\n  - ` + failures.join("\n  - "));
    process.exit(1);
  }
  console.log(`✅ cell-format (${process.env.ENGINE || "chromium"}): tryby, format, pasek formuły, pole nazwy, malarz, zapis`);
}

run().catch((e) => { console.error(e); process.exit(1); });
