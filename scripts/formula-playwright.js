// formula-playwright.js — formuły w aplikacji (formula-engine/-calc/-ui).
//
// Co musi być prawdą:
//   1. =B2*C2 wpisane w komórkę liczy wynik; pasek formuły pokazuje formułę, nie wynik,
//   2. zmiana B2 przelicza D2 (zależne), Ctrl+Z wraca też z wynikiem formuły,
//   3. „Wypełnij w dół” z formuły przesuwa odwołania (D5 = B5*C5),
//   4. polska składnia: =JEŻELI(B2>5;"dużo";"mało"), =SUMA(…;1,5) — przecinek dziesiętny,
//   5. inny arkusz: =C2*Kurs!B1,
//   6. błąd składni zostawia edytor otwarty; odwołanie do samej siebie — odmowa,
//   7. Σ na zakresie D2:D5 wstawia =SUM(D2:D5) pod spodem (nowy wiersz w tabeli),
//   8. wskazywanie: „=SUMA(” + klik B3 + Shift+klik B5 + „)” = =SUMA(B3:B5),
//   9. podpowiedź funkcji: „=SUMA.J” + Tab → „=SUMA.JEŻELI(”,
//  10. funkcja spoza silnika: zapis formuły, w komórce „(policzy Excel)”,
//  11. zapis → openpyxl: formuły po angielsku z „,”, wyniki zapisane; schemat Office bez
//      nowych błędów; ponowne wczytanie pokazuje wyniki i formułę w pasku.
//
// Fixture budowany w przeglądarce (XLSX → bajty → input pliku), bez prywatnych plików.

const { chromium, webkit } = require("playwright");
const path = require("path");
const fs = require("fs");
const os = require("os");
const { spawnSync } = require("child_process");

const APP_URL = process.env.APP_URL || "http://127.0.0.1:4175/";
const ENGINE = process.env.ENGINE === "webkit" ? webkit : chromium;
const SLEEP_SCALE = Math.min(1, Math.max(0.1, Number(process.env.SLEEP_SCALE || 1)));
const sleep = (ms) => new Promise((r) => setTimeout(r, Math.round(ms * SLEEP_SCALE)));
const VENV_PY = path.join(__dirname, "..", ".venv/bin/python");
const PY = fs.existsSync(VENV_PY) ? VENV_PY : "python3";
const MOD = process.platform === "darwin" ? "Meta" : "Control";

async function run() {
  const browser = await ENGINE.launch({ headless: true });
  const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1360, height: 900 } });
  await context.addInitScript(() => { localStorage.setItem("introPlayed", "true"); localStorage.setItem("excel-workbench-lang", "pl"); });
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push("pageerror: " + e.message));
  page.on("console", (m) => { if (m.type() === "error") errors.push("console: " + m.text()); });
  const failures = [];
  const ok = (name, cond, got) => { if (!cond) failures.push(`${name} — dostałem: ${JSON.stringify(got)}`); };
  const tmp = fs.mkdtempSync(path.join(os.tmpdir(), "formula-"));

  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(() => { document.getElementById("heroSplash")?.remove(); try { ensureXlsxLibs(false); } catch {} });
  await page.waitForFunction(() => typeof XLSX !== "undefined" && typeof FormulaEngine !== "undefined");
  const fixtureB64 = await page.evaluate(() => {
    const dane = XLSX.utils.aoa_to_sheet([
      ["Nazwa", "Ilość", "Cena", "Wartość", "W euro", "Notatka"],
      ["Śruby", 4, 2.5, null, null, null],
      ["Nakrętki", 10, 0.5, null, null, null],
      ["Podkładki", 3, 1, null, null, null],
      ["Wkręty", 8, 1.25, null, null, null],
    ]);
    const kurs = XLSX.utils.aoa_to_sheet([["EUR", 4.25]]);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, dane, "Dane");
    XLSX.utils.book_append_sheet(wb, kurs, "Kurs");
    const u8 = XLSX.write(wb, { bookType: "xlsx", type: "array" });
    let s = ""; const b = new Uint8Array(u8); for (let i = 0; i < b.length; i++) s += String.fromCharCode(b[i]);
    return btoa(s);
  });
  const fixture = path.join(tmp, "formuly.xlsx");
  fs.writeFileSync(fixture, Buffer.from(fixtureB64, "base64"));

  async function openFile(file) {
    await page.setInputFiles("#fileInput", file);
    await page.waitForFunction(() => document.getElementById("sheetSelect")?.options?.length > 0, null, { timeout: 15000 });
    await sleep(200);
    await page.evaluate(() => { sheetSelect.value = "Dane"; });
    await page.evaluate(() => loadBtn.click());
    await page.waitForFunction(() => currentSheetName === "Dane" && document.querySelector("#dataTable tbody tr[data-row-key]"), null, { timeout: 15000 });
    await sleep(400);
    await page.evaluate(() => { if (typeof setSidebarOpen === "function") setSidebarOpen(false); });
  }
  await openFile(fixture);

  const td = (r, c) => `tbody tr[data-row-key="wide:${r}"] td[data-col-index="${c}"]`;
  const shown = (r, c) => page.evaluate((sel) => document.querySelector(sel)?.textContent ?? null, td(r, c));
  const typeInCell = async (r, c, text, key = "Enter") => {
    await page.evaluate(({ r, c }) => {
      setFocusedCell(`wide:${r}`, c, { scroll: true });
      openCellEditor(findCellElement(focusedCellState));
    }, { r, c });
    await page.waitForSelector("input.cell-editor");
    await page.fill("input.cell-editor", text);
    if (key) await page.keyboard.press(key);
    await sleep(200);
  };
  const fx = () => page.evaluate(() => formulaInput.value);

  // 1. wpisanie formuły
  await typeInCell(1, 3, "=B2*C2");
  ok("1. D2 = 4*2,5 = 10", (await shown(1, 3)) === "10", await shown(1, 3));
  await page.evaluate(() => setFocusedCell("wide:1", 3, { scroll: false }));
  await sleep(100);
  ok("1. pasek formuły pokazuje formułę", (await fx()) === "=B2*C2", await fx());
  const pend = await page.evaluate(() => pendingEdits.Dane.D2);
  ok("1. do zapisu: f=B2*C2, v=10", pend && pend.f === "B2*C2" && pend.v === 10, pend);

  // 2. przeliczenie zależnych + cofanie
  await typeInCell(1, 1, "6");
  ok("2. B2=6 → D2 = 15", (await shown(1, 3)) === "15", await shown(1, 3));
  await page.evaluate(() => findCellElement(focusedCellState)?.focus());
  await page.keyboard.press(`${MOD}+z`);
  await sleep(250);
  ok("2. Ctrl+Z → B2 = 4, D2 = 10", (await shown(1, 1)) === "4" && (await shown(1, 3)) === "10", [await shown(1, 1), await shown(1, 3)]);

  // 3. wypełnianie w dół
  await page.evaluate(() => { setFocusedCell("wide:1", 3, { scroll: true }); setSelectedCell("wide:4", 3, { scroll: false }); fillSelection("down"); });
  await sleep(250);
  const filled = await page.evaluate(() => ({ d5: workbook.Sheets.Dane.D5?.f, v: workbook.Sheets.Dane.D5?.v }));
  ok("3. D5 = B5*C5 = 10", filled.d5 === "B5*C5" && filled.v === 10, filled);
  ok("3. D3 widać 5", (await shown(2, 3)) === "5", await shown(2, 3));

  // 4. polska składnia
  await typeInCell(1, 5, '=JEŻELI(B2>5;"dużo";"mało")');
  ok("4. JEŻELI → mało", (await shown(1, 5)) === "mało", await shown(1, 5));
  await typeInCell(2, 5, "=SUMA(B3:B4;1,5)");
  ok("4. SUMA(…;1,5) = 14,5", (await page.evaluate(() => workbook.Sheets.Dane.F3.v)) === 14.5, await page.evaluate(() => workbook.Sheets.Dane.F3));
  ok("4. do pliku po angielsku", (await page.evaluate(() => pendingEdits.Dane.F3.f)) === "SUM(B3:B4,1.5)", await page.evaluate(() => pendingEdits.Dane.F3));

  // 5. inny arkusz
  await typeInCell(1, 4, "=C2*Kurs!B1");
  ok("5. E2 = 2,5*4,25", Math.abs((await page.evaluate(() => workbook.Sheets.Dane.E2.v)) - 10.625) < 1e-9, await page.evaluate(() => workbook.Sheets.Dane.E2));

  // 6. błąd składni / cykl
  await typeInCell(3, 5, "=SUMA(B2");
  ok("6. błąd składni — edytor dalej otwarty", !!(await page.$("input.cell-editor")), null);
  await page.keyboard.press("Escape");
  await sleep(150);
  await typeInCell(3, 5, "=F4+1");
  ok("6. odwołanie do samej siebie — edytor otwarty, nic nie zapisane", !!(await page.$("input.cell-editor")) && !(await page.evaluate(() => workbook.Sheets.Dane.F4?.f)), null);
  await page.keyboard.press("Escape");
  await sleep(150);

  // 7. autosumowanie zakresu
  await page.evaluate(() => { setFocusedCell("wide:1", 3, { scroll: true }); setSelectedCell("wide:4", 3, { scroll: false }); });
  await page.click("#fmtSumBtn");
  await sleep(400);
  const sum = await page.evaluate(() => ({ f: workbook.Sheets.Dane.D6?.f, v: workbook.Sheets.Dane.D6?.v, row: !!document.querySelector('tbody tr[data-row-key="wide:5"]') }));
  ok("7. Σ → D6 = SUM(D2:D5) = 28 i wiersz widać", sum.f === "SUM(D2:D5)" && sum.v === 28 && sum.row, sum);

  // 8. wskazywanie komórek myszą
  await page.evaluate(() => { setFocusedCell("wide:4", 5, { scroll: true }); openCellEditor(findCellElement(focusedCellState)); });
  await page.waitForSelector("input.cell-editor");
  await page.keyboard.type("=SUMA(");
  await page.click(td(2, 1));
  await page.click(td(4, 1), { modifiers: ["Shift"] });
  const pointed = await page.evaluate(() => document.querySelector("input.cell-editor")?.value);
  ok("8. wskazanie B3 + Shift B5", pointed === "=SUMA(B3:B5", pointed);
  await page.keyboard.type(")");
  await page.keyboard.press("Enter");
  await sleep(200);
  ok("8. F5 = 21", (await page.evaluate(() => workbook.Sheets.Dane.F5?.v)) === 21, await page.evaluate(() => workbook.Sheets.Dane.F5));

  // 9. podpowiedź funkcji
  await page.evaluate(() => { setFocusedCell("wide:3", 4, { scroll: true }); openCellEditor(findCellElement(focusedCellState)); });
  await page.waitForSelector("input.cell-editor");
  await page.keyboard.type("=SUMA.J");
  await sleep(150);
  const sugg = await page.evaluate(() => [...document.querySelectorAll(".fx-suggest:not([hidden]) .fx-suggest-name")].map((x) => x.textContent));
  ok("9. lista ma SUMA.JEŻELI", sugg.includes("SUMA.JEŻELI"), sugg);
  await page.keyboard.press("Tab");
  const acc = await page.evaluate(() => document.querySelector("input.cell-editor")?.value);
  ok("9. Tab wstawia SUMA.JEŻELI(", acc === "=SUMA.JEŻELI(", acc);
  await page.keyboard.press("Escape"); // zamyka podpowiedź
  await page.keyboard.press("Escape"); // zamyka edytor
  await sleep(150);

  // 10. funkcja spoza silnika
  await typeInCell(4, 4, "=FORECAST(1;B2:B5;C2:C5)");
  ok("10. komórka: (policzy Excel)", (await shown(4, 4)) === "(policzy Excel)", await shown(4, 4));
  ok("10. formuła do zapisu", (await page.evaluate(() => pendingEdits.Dane.E5?.f)) === "FORECAST(1,B2:B5,C2:C5)", await page.evaluate(() => pendingEdits.Dane.E5));

  // 11. zapis
  const b64 = await page.evaluate(async () => {
    const u8 = await buildOutputBytes("xlsx");
    let s = ""; for (let i = 0; i < u8.length; i += 0x8000) s += String.fromCharCode(...u8.subarray(i, i + 0x8000));
    return btoa(s);
  });
  const out = path.join(tmp, "zapis.xlsx");
  fs.writeFileSync(out, Buffer.from(b64, "base64"));
  const py = spawnSync(PY, ["-c", `
import json, sys
from openpyxl import load_workbook
f = load_workbook(sys.argv[1])["Dane"]
v = load_workbook(sys.argv[1], data_only=True)["Dane"]
refs = ["D2","D5","D6","E2","F2","F3","F5","E5"]
print(json.dumps({r: [f[r].value, v[r].value] for r in refs}))
`, out], { encoding: "utf8" });
  let res = null;
  try { res = JSON.parse(py.stdout); } catch { failures.push("11. openpyxl: " + py.stderr.slice(0, 300)); }
  if (res) {
    ok("11. D2 =B2*C2 / 10", res.D2[0] === "=B2*C2" && res.D2[1] === 10, res.D2);
    ok("11. D5 =B5*C5", res.D5[0] === "=B5*C5" && res.D5[1] === 10, res.D5);
    ok("11. D6 =SUM(D2:D5) / 28", res.D6[0] === "=SUM(D2:D5)" && res.D6[1] === 28, res.D6);
    ok("11. E2 =C2*Kurs!B1", res.E2[0] === "=C2*Kurs!B1", res.E2);
    ok("11. F2 =IF(B2>5,\"dużo\",\"mało\")", res.F2[0] === '=IF(B2>5,"dużo","mało")' && res.F2[1] === "mało", res.F2);
    ok("11. F3 =SUM(B3:B4,1.5)", res.F3[0] === "=SUM(B3:B4,1.5)" && res.F3[1] === 14.5, res.F3);
    ok("11. E5 bez wyniku (policzy Excel)", res.E5[0] === "=FORECAST(1,B2:B5,C2:C5)" && res.E5[1] === null, res.E5);
  }
  const ooxml = require("./ooxml-validate.js");
  if (ooxml.available()) {
    const [vo, vs] = ooxml.validate([fixture, out]);
    const fresh = ooxml.newErrors(vo, vs);
    ok("11. schemat Office bez nowych błędów", !vs.error && !fresh.length, vs.error || fresh.slice(0, 3).map(ooxml.short));
  }
  await openFile(out);
  ok("11. po wczytaniu D6 = 28", (await shown(5, 3)) === "28", await shown(5, 3));
  await page.evaluate(() => setFocusedCell("wide:1", 4, { scroll: false }));
  await sleep(100);
  ok("11. po wczytaniu pasek formuły E2", (await fx()) === "=C2*Kurs!B1", await fx());

  await browser.close();
  fs.rmSync(tmp, { recursive: true, force: true });
  if (errors.length) failures.push("błędy strony: " + errors.join(" | "));
  if (failures.length) {
    console.error(`❌ formula (${process.env.ENGINE || "chromium"}):\n  - ` + failures.join("\n  - "));
    process.exit(1);
  }
  console.log(`✅ formula (${process.env.ENGINE || "chromium"}): wpisywanie PL/EN, zależne, wypełnianie, Σ, wskazywanie, podpowiedzi, zapis`);
}

run().catch((e) => { console.error(e); process.exit(1); });
