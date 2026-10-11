// clear-keeps-format-playwright.js — „wyczyść zawartość” zostawia FORMAT komórki (jak Excel).
//
// Zgłoszenie Mateusza (2026-10-11): wejście w komórkę, zaznacz wszystko, usuń, wyjdź →
// pasek pokazywał „Ogólny”. Plik po zapisie miał styl (xlsx-patch), ale w aplikacji komórka
// znikała w całości: pasek kłamał, a nowa wartość wpisana w to miejsce traciła format.
//
// Co musi być prawdą (plik z openpyxl: B = data „yyyy-mm-dd”, E = kwota „#,##0.00 "zł"”):
//   1. wyczyszczenie przez edytor (pusty tekst + Enter) — pasek dalej pokazuje format kolumny,
//   2. to samo przez Delete na zaznaczeniu,
//   3. nowa liczba wpisana w wyczyszczoną komórkę kwoty dostaje format „zł”,
//   4. pusta, ale sformatowana komórka z PLIKU (bez obiektu komórki) pokazuje format z xf,
//   5. wyczyszczony wiersz nie staje się „pełnym” wierszem (zaślepka liczy się jako pusta).

const pw = require("playwright");
const path = require("path");
const fs = require("fs");
const os = require("os");
const { spawnSync } = require("child_process");
const ooxml = require("./ooxml-validate");
const VENV_PY = path.join(__dirname, "..", ".venv/bin/python");
const PY = fs.existsSync(VENV_PY) ? VENV_PY : "python3";

const APP_URL = process.env.APP_URL || "http://127.0.0.1:4175/";
const ENGINE = process.env.ENGINE || "chromium";
const SLEEP_SCALE = Math.min(1, Math.max(0.1, Number(process.env.SLEEP_SCALE || 1)));
const sleep = (ms) => new Promise((r) => setTimeout(r, Math.round(ms * SLEEP_SCALE)));
const FILE = path.join(__dirname, "fixtures", "openpyxl-features.xlsx");

async function run() {
  const browser = await pw[ENGINE].launch({ headless: true });
  const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1280, height: 860 } });
  await context.addInitScript(() => { localStorage.setItem("introPlayed", "true"); });
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  page.on("dialog", (d) => d.accept());
  const failures = [];
  const ok = (name, cond, got) => { if (!cond) failures.push(`${name} — dostałem: ${JSON.stringify(got)}`); };

  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(() => { document.getElementById("heroSplash")?.remove(); return ensureXlsxLibs(false); });
  await page.setInputFiles("#fileInput", FILE);
  await page.waitForFunction(() => currentHeaderRow === 2 && !!document.querySelector("#dataTable tbody tr[data-row-key]"), null, { timeout: 20000 });
  await sleep(400);

  const td = (r, c) => `#dataTable tbody tr[data-row-key="wide:${r}"] td[data-col-index="${c}"]`;
  const numFmtAt = (r, c) => page.evaluate(({ r, c }) => { setFocusedCell(`wide:${r}`, c, { scroll: false }); return cfmtActiveCell().numFmt; }, { r, c });
  const clearViaEditor = async (r, c) => {
    await page.dblclick(td(r, c));
    await page.waitForSelector(`${td(r, c)} input.cell-editor`, { timeout: 3000 });
    await page.keyboard.press("ControlOrMeta+a");
    await page.keyboard.press("Backspace");
    await page.keyboard.press("Enter");
    await sleep(250);
  };

  // 1. data przez edytor
  const dateBefore = await numFmtAt(2, 1);
  await clearViaEditor(2, 1);
  const d1 = await page.evaluate(() => ({ cell: workbook.Sheets[currentSheetName].B3, text: document.querySelector('#dataTable tbody tr[data-row-key="wide:2"] td[data-col-index="1"]').textContent }));
  const dateAfter = await numFmtAt(2, 1);
  ok("1. wyczyszczona data zostaje sformatowana (pasek nie pokazuje „Ogólny”)", /y/.test(dateBefore) && dateAfter === dateBefore && d1.text === "" && d1.cell && d1.cell.v === undefined, { dateBefore, dateAfter, d1 });

  // 2. kwota przez Delete na zaznaczeniu
  const moneyBefore = await numFmtAt(3, 4);
  await page.evaluate(() => { setFocusedCell("wide:3", 4, { scroll: false }); focusGridCell && focusGridCell(); });
  await page.evaluate(() => cfmtClearContents());
  await sleep(250);
  const moneyAfter = await numFmtAt(3, 4);
  ok("2. Delete na kwocie zostawia format „zł”", /zł/.test(moneyBefore) && moneyAfter === moneyBefore, { moneyBefore, moneyAfter });

  // 3. nowa liczba w wyczyszczonej komórce kwoty — z formatem
  await page.dblclick(td(3, 4));
  await page.waitForSelector(`${td(3, 4)} input.cell-editor`, { timeout: 3000 });
  await page.keyboard.type("1234,5");
  await page.keyboard.press("Enter");
  await sleep(250);
  const m3 = await page.evaluate(() => ({ z: workbook.Sheets[currentSheetName].E4.z, v: workbook.Sheets[currentSheetName].E4.v }));
  ok("3. nowa kwota w wyczyszczonej komórce ma dawny format", m3.v === 1234.5 && /zł/.test(m3.z || ""), m3);

  // 4. pusta komórka z pliku (bez obiektu komórki) — format z xf
  const m4 = await page.evaluate(() => {
    const sheet = workbook.Sheets[currentSheetName];
    const saved = sheet.E5;
    delete sheet.E5; // tak wygląda pusta, sformatowana komórka wczytana z pliku
    const code = cfmtFileNumFmt("E5");
    sheet.E5 = saved;
    return code;
  });
  ok("4. format pustej, sformatowanej komórki z pliku czytany z xf", /zł/.test(m4 || ""), m4);

  // 5. cały wiersz wyczyszczony — zaślepki nie robią z niego wiersza z danymi
  const rowsBefore = await page.evaluate(() => baseRows.length);
  await page.evaluate(() => {
    // wiersz 21 (bez treści w J–K, gdzie stoją podsumowania)
    for (let c = 0; c < currentHeaders.length; c++) updateSheetCell(20, c, null);
  });
  await page.evaluate(() => loadBtn.click());
  await page.waitForFunction((n) => baseRows.length === n - 1, rowsBefore, { timeout: 8000 }).catch(() => {});
  const rowsAfter = await page.evaluate(() => ({ n: baseRows.length, has21: baseRows.some((r) => r.rowIndex0 === 20) }));
  ok("5. wyczyszczony wiersz (z zaślepkami formatu) znika z tabeli jak pusty", rowsAfter.n === rowsBefore - 1 && !rowsAfter.has21, { rowsBefore, rowsAfter });

  // 6. wpisywane wartości: liczba to liczba (wcześniej „42” zapisywało się jako 10.02.1900),
  // polski przecinek, odstęp tysięcy; daty tylko z zapisem daty; telefon/adres zostają tekstem.
  const parsed = await page.evaluate(() => Object.fromEntries(["42", "12,5", "-3,75", "1 234,50", "007", "2026-05-01", "01.05.2026", "7 maja 2026", "500 600 700", "Mariańska 27, 29", "12%"].map((s) => {
    const r = parseInputValue(s);
    return [s, r ? `${r.type}:${r.value instanceof Date ? `${r.value.getFullYear()}-${r.value.getMonth() + 1}-${r.value.getDate()}` : r.value}` : "null"];
  })));
  const want = { "42": "number:42", "12,5": "number:12.5", "-3,75": "number:-3.75", "1 234,50": "number:1234.5", "007": "number:7", "2026-05-01": "date:2026-5-1", "01.05.2026": "date:2026-5-1", "7 maja 2026": "date:2026-5-7", "500 600 700": "string:500 600 700", "Mariańska 27, 29": "string:Mariańska 27, 29", "12%": "string:12%" };
  const bad = Object.keys(want).filter((k) => parsed[k] !== want[k]).map((k) => `${k} → ${parsed[k]} (chciałem ${want[k]})`);
  ok("6. wpisywane liczby/daty/teksty rozpoznane jak w Excelu (PL)", bad.length === 0, bad);
  await page.dblclick(td(4, 0));
  await page.waitForSelector(`${td(4, 0)} input.cell-editor`, { timeout: 3000 });
  await page.keyboard.press("ControlOrMeta+a");
  await page.keyboard.type("42");
  await page.keyboard.press("Enter");
  await sleep(250);
  const a5 = await page.evaluate(() => ({ v: workbook.Sheets[currentSheetName].A5.v, t: workbook.Sheets[currentSheetName].A5.t, p: JSON.stringify(pendingEdits[currentSheetName].A5) }));
  ok("6. „42” wpisane w komórkę zapisuje się jako liczba (nie data)", a5.v === 42 && a5.t === "n" && /"t":"n"/.test(a5.p), a5);

  // 7. ZAPIS do pliku (to, co naprawdę zostaje): openpyxl jako drugi czytnik, schemat Office
  // (Open XML SDK) i ponowne wczytanie w aplikacji. Stan po sekcjach 1–6:
  //   B3 wyczyszczona (data), E4 = 1234,5 po wyczyszczeniu (zł), A5 = 42, wiersz 21 pusty.
  const tmp = fs.mkdtempSync(path.join(os.tmpdir(), "clear-format-"));
  const out = path.join(tmp, "zapis.xlsx");
  const b64 = await page.evaluate(async () => {
    const u8 = await buildOutputBytes("xlsx");
    let s = ""; for (let i = 0; i < u8.length; i += 0x8000) s += String.fromCharCode(...u8.subarray(i, i + 0x8000));
    return btoa(s);
  });
  fs.writeFileSync(out, Buffer.from(b64, "base64"));
  const py = spawnSync(PY, ["-c", `
import json, sys
from openpyxl import load_workbook
ws = load_workbook(sys.argv[1])["Dane"]
def c(ref):
    x = ws[ref]
    return {"v": None if x.value is None else str(x.value), "type": type(x.value).__name__, "fmt": x.number_format}
print(json.dumps({k: c(k) for k in ["B3", "E4", "A5", "B21", "E21", "A21", "B4", "E3"]}))
`, out], { encoding: "utf8" });
  let f = null;
  try { f = JSON.parse(py.stdout); } catch { failures.push("7. openpyxl nie przeczytał zapisu: " + (py.stderr || "").slice(0, 300)); }
  if (f && process.env.DEBUG_SAVE) console.log("openpyxl:", JSON.stringify(f));
  if (f) {
    ok("7. plik: wyczyszczona data B3 pusta, format daty został", f.B3.v === null && /yy/.test(f.B3.fmt), f.B3);
    ok("7. plik: E4 = 1234,5 jako liczba w formacie „zł”", f.E4.type === "float" && f.E4.v === "1234.5" && /zł/.test(f.E4.fmt), f.E4);
    ok("7. plik: A5 = 42 jako liczba (nie data)", f.A5.type === "int" && f.A5.v === "42", f.A5);
    ok("7. plik: wyczyszczony wiersz 21 pusty, formaty kolumn zostały", f.A21.v === null && f.B21.v === null && /yy/.test(f.B21.fmt) && f.E21.v === null && /zł/.test(f.E21.fmt), { A21: f.A21, B21: f.B21, E21: f.E21 });
    ok("7. plik: nieruszane komórki bez zmian", f.B4.type === "datetime" && f.E3.type === "float", { B4: f.B4, E3: f.E3 });
  }
  if (ooxml.available()) {
    const [vo, vs] = ooxml.validate([FILE, out]);
    const fresh = ooxml.newErrors(vo, vs).map(ooxml.short);
    if (process.env.DEBUG_SAVE) console.log("schemat Office:", vs && vs.ok ? "OK" : JSON.stringify(vs && (vs.error || vs.errors.length)), "nowe błędy:", fresh.length);
    ok("7. plik zgodny ze schematem Office (bez nowych błędów)", vs && !vs.error && fresh.length === 0, fresh.slice(0, 5));
  } else {
    console.warn("⚠️  schemat Office pominięty — brak .NET / tools/ooxml-validator");
  }
  // ponowne wczytanie zapisanego pliku w aplikacji
  await page.evaluate(() => { theadEl.replaceChildren(); tbodyEl.replaceChildren(); });
  await page.setInputFiles("#fileInput", out);
  await page.waitForFunction(() => currentFileName === "zapis.xlsx" && !!document.querySelector("#dataTable tbody tr[data-row-key]"), null, { timeout: 20000 });
  await sleep(300);
  const back = await page.evaluate(() => {
    const txt = (r, c) => document.querySelector(`#dataTable tbody tr[data-row-key="wide:${r}"] td[data-col-index="${c}"]`)?.textContent ?? null;
    return { a5: txt(4, 0), e4: txt(3, 4), b3: txt(2, 1), a5v: workbook.Sheets[currentSheetName].A5.v };
  });
  const fmtB3 = await numFmtAt(2, 1);
  ok("7. po ponownym wczytaniu: 42 jako liczba, kwota w zł, pusta data z formatem", back.a5 === "42" && back.a5v === 42 && /zł/.test(back.e4 || "") && back.b3 === "" && /yy/.test(fmtB3), { ...back, fmtB3 });
  fs.rmSync(tmp, { recursive: true, force: true });

  await browser.close();
  if (errors.length) failures.push("błędy strony: " + errors.join(" | "));
  if (failures.length) {
    console.error(`❌ clear-keeps-format (${ENGINE}):\n  - ` + failures.join("\n  - "));
    process.exit(1);
  }
  console.log(`✅ clear-keeps-format (${ENGINE}): wyczyszczenie zostawia format, nowa wartość go dostaje, format pustych komórek z pliku`);
}

run().catch((e) => { console.error(e); process.exit(1); });
