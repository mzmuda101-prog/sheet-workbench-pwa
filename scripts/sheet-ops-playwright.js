// sheet-ops-playwright.js — „od zera”: Nowy skoroszyt i arkusze w aplikacji (app/sheet-ops-ui.js).
//   node scripts/sheet-ops-playwright.js            (Chromium, komputer)
//   ENGINE=webkit TOUCH=1 node scripts/sheet-ops-playwright.js   (Safari, telefon)
//   EXCEL_DIR=/private/tmp/swb-excel-check …         (+ kopie zapisanych plików dla excel-real-check.py)
//
// Pilnuje:
//   1. „Nowy skoroszyt” z ekranu startowego: Zeszyt1.xlsx / Arkusz1, od razu Edycja, pusta siatka
//      (8 kolumn A–H, wiersze-duchy od 2), plik nie jest „niezapisany”, nie trafia do „Ostatnio otwierane”;
//      w Podglądzie pusty arkusz = komunikat, nie siatka;
//   2. klik w pusty nagłówek = nazwa kolumny w A1 (puste pole = bez zmian i bez komunikatu);
//      klik w ducha = wpisywanie w A2; Enter schodzi do A3;
//   3. lista arkuszy (▾) przy jednym arkuszu: „Nowy arkusz / Zmień nazwę… / Usuń arkusz…”;
//      Nowy arkusz = Arkusz2 za bieżącym, wpisane dane zostają, plik „niezapisany” także po
//      przełączeniu arkusza, historia Cofnij wyczyszczona (jak w Excelu);
//   4. zmiana nazwy: powtórzona nazwa odrzucona w oknie; nowa nazwa poprawia formułę w innym arkuszu
//      (='Moje dane'!A2 liczy dalej); usunięcie: pytanie, formuła → #ADR!, ostatniego arkusza — odmowa;
//   5. zapis: openpyxl czyta arkusze, A1/A2 i formułę; bez błędów schematu (Open XML SDK);
//   6. istniejący plik (przykład): nowy arkusz i zapis bez NOWYCH błędów schematu, dane przykładu zostają;
//   7. bez błędów strony i natywnych okien.

const { chromium, webkit } = require("playwright");
const fs = require("fs");
const os = require("os");
const path = require("path");
const { execFileSync } = require("child_process");

const APP_URL = process.env.APP_URL || "http://127.0.0.1:4175/";
const ENGINE_NAME = process.env.ENGINE === "webkit" ? "webkit" : "chromium";
const ENGINE = ENGINE_NAME === "webkit" ? webkit : chromium;
const TOUCH = process.env.TOUCH === "1";
const PY = path.join(__dirname, "..", ".venv", "bin", "python");
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
let ooxml = null;
try { ooxml = require("./ooxml-validate"); } catch { ooxml = null; }
const ooxmlOn = () => !!(ooxml && ooxml.available && ooxml.available());
const tag = `${ENGINE_NAME}${TOUCH ? "-touch" : ""}`;

async function run() {
  const fails = [];
  const check = (name, ok, detail) => {
    console.log(`${ok ? "✅" : "❌"} ${name}${ok ? "" : ` — ${JSON.stringify(detail)}`}`);
    if (!ok) fails.push(name);
  };
  const browser = await ENGINE.launch({ headless: true });
  const ctx = await browser.newContext({
    serviceWorkers: "block",
    viewport: TOUCH ? { width: 390, height: 844 } : { width: 1280, height: 900 },
    hasTouch: TOUCH, isMobile: TOUCH && ENGINE_NAME !== "webkit" ? true : undefined,
  });
  await ctx.addInitScript(() => { localStorage.setItem("introPlayed", "true"); });
  const p = await ctx.newPage();
  const errors = [];
  p.on("pageerror", (e) => errors.push(e.message));
  p.on("dialog", (d) => { errors.push(`natywne okno: ${d.message()}`); d.dismiss(); });
  await p.goto(APP_URL, { waitUntil: "load" });
  await p.evaluate(() => document.getElementById("heroSplash")?.remove());
  const idle = () => p.waitForFunction(() => document.body.getAttribute("aria-busy") !== "true" && typeof workbook !== "undefined" && workbook && currentSheetName, null, { timeout: 20000 });
  const tap = (sel) => (TOUCH ? p.tap(sel) : p.click(sel));
  const shot = (name) => (process.env.SHOTS ? p.screenshot({ path: path.join(process.env.SHOTS, `${tag}-${name}.png`) }) : null);
  const dropToasts = () => p.evaluate(() => document.querySelectorAll(".toast").forEach((t) => t.remove()));
  const openSheetList = async () => {
    // telefon: nagłówek zwija się przy pracy w tabeli — jak użytkownik, najpierw rozwijamy go uchwytem
    if (await p.evaluate(() => document.body.classList.contains("hero-collapsed"))) { await tap("#heroGrip"); await sleep(500); }
    // telefon z plikiem (etap 7a): przycisk listy arkuszy jest na dole — otwarty od dołu „Aa” go zasłania,
    // więc jak użytkownik najpierw „Gotowe”
    // (pasek formatu bez zaznaczenia ma aria-disabled — Playwright nie stuka w nic pod nim, palec tak)
    if (await p.evaluate(() => { const d = document.querySelector(".fmt-sheet-host .fmt-sheet-done"); if (!d || !d.offsetParent) return false; d.click(); return true; })) await sleep(350);
    await tap("#heroSheetBtn");
    await p.waitForSelector(TOUCH ? "#sheetPickerOverlay.open .sheet-op-item" : "#sheetMenu:not([hidden]) .sheet-op-item", { timeout: 5000 });
    await sleep(350);
  };
  const sheetOp = async (op) => { await openSheetList(); await tap(`${TOUCH ? "#sheetPickerList" : "#sheetMenu"} [data-sheet-op="${op}"]`); };
  const dialogInput = () => p.waitForSelector(".app-dialog-box input", { state: "visible", timeout: 5000 });
  const state = () => p.evaluate(() => ({
    file: currentFileName, names: workbook.SheetNames.slice(), cur: currentSheetName, dirty: hasUnsavedChanges,
    mode: swbEditMode, headers: currentHeaders.slice(), rows: baseRows.length, undo: undoStack.length,
  }));

  // 1. nowy skoroszyt
  await tap("#emptyNewBtn");
  await idle();
  await sleep(500);
  let s = await state();
  check("1. Zeszyt1.xlsx / Arkusz1, Edycja, niezmieniony", s.file === "Zeszyt1.xlsx" && s.names.join() === "Arkusz1" && s.mode === "edit" && !s.dirty, s);
  check("1. pusta siatka: kolumny A–H, bez wierszy danych", s.headers.join("") === "ABCDEFGH" && s.rows === 0, s);
  const ghosts = await p.evaluate(() => [...document.querySelectorAll(".ghost-row .ghost-head")].map((td) => td.textContent));
  check("1. duchy od wiersza 2", ghosts[0] === "2" && ghosts.length >= 10, ghosts.slice(0, 3));
  const recent = await p.evaluate(() => (typeof swbRecent !== "undefined" ? swbRecent._current() : ""));
  check("1. nie trafia do „Ostatnio otwierane”", !recent, recent);
  await p.evaluate(() => sheetEditUi.setMode("view"));
  await sleep(300);
  const viewEmpty = await p.evaluate(() => !document.getElementById("emptyState").classList.contains("hidden") && document.getElementById("tableWrap").classList.contains("hidden"));
  check("1. Podgląd pustego arkusza: komunikat zamiast siatki", viewEmpty, null);
  await p.evaluate(() => sheetEditUi.setMode("edit"));
  await sleep(300);
  await dropToasts();

  // 1b. klikanie po pustych wierszach (nagranie 2026-11-14 19:47): bez duplikatów, bez skoków numerów,
  //     puste wiersze nie liczą się jako „N wierszy”
  const ghostClick = (i, c) => p.evaluate(([i, c]) => document.querySelectorAll(".ghost-row")[i].querySelectorAll("td")[c + 1].click(), [i, c]);
  await ghostClick(1, 0); await sleep(300);
  await ghostClick(4, 1); await sleep(300);
  await ghostClick(2, 2); await sleep(300);
  const clicks = await p.evaluate(() => ({ rows: baseRows.map((r) => r.rowIndex0 + 1), status: document.getElementById("status").textContent }));
  const contiguous = clicks.rows.every((r, i) => r === i + 2);
  check("1b. klikane puste wiersze: kolejne numery od 2, bez powtórek", contiguous && clicks.rows.length >= 3, clicks.rows);
  check("1b. puste wiersze nie liczą się jako wiersze", !/[1-9]\d* wiersz/.test(clicks.status), clicks.status);
  await p.keyboard.press("Escape");
  await p.evaluate(() => { focusedCellState = null; renderActiveTable(); });
  await sleep(300);
  const pruned = await p.evaluate(() => baseRows.length);
  check("1b. po odejściu kursora puste wiersze robocze znikają", pruned === 0, pruned);

  // 2. nazwa kolumny i dane
  const th0 = '#dataTable thead tr.header-row th[data-col-index="0"]';
  await tap(th0);
  await p.waitForSelector("input.hdr-editor", { timeout: 5000 });
  await p.keyboard.press("Escape"); // puste → bez zmian
  await tap(th0);
  await p.waitForSelector("input.hdr-editor", { timeout: 5000 });
  check("2. pole nazwy pustej kolumny startuje puste", (await p.inputValue("input.hdr-editor")) === "", null);
  await p.fill("input.hdr-editor", "Produkt");
  await p.keyboard.press("Enter");
  await sleep(300);
  s = await state();
  check("2. nazwa kolumny = A1", s.headers[0] === "Produkt" && (await p.evaluate(() => workbook.Sheets.Arkusz1.A1 && workbook.Sheets.Arkusz1.A1.v)) === "Produkt", s.headers);
  await tap('.ghost-row td[data-ghost-col="0"]');
  await p.waitForSelector("input.cell-editor", { timeout: 5000 });
  await p.keyboard.type("Chleb");
  await p.keyboard.press("Enter");
  await sleep(400);
  const a2 = await p.evaluate(() => ({ v: workbook.Sheets.Arkusz1.A2 && workbook.Sheets.Arkusz1.A2.v, focus: focusedCellState && focusedCellState.rowKey }));
  check("2. A2 wpisane, Enter schodzi niżej", a2.v === "Chleb" && a2.focus !== "wide:0", a2);
  await p.evaluate(() => { updateSheetCell(1, 1, { value: 5, type: "number" }); setDirtyState(true); });

  // 3. nowy arkusz
  await dropToasts();
  await sheetOp("add");
  await idle();
  await sleep(500);
  s = await state();
  check("3. Arkusz2 za bieżącym, na wierzchu", s.names.join() === "Arkusz1,Arkusz2" && s.cur === "Arkusz2", s);
  check("3. niezapisany, Cofnij wyczyszczone", s.dirty && s.undo === 0, s);
  const kept = await p.evaluate(() => [workbook.Sheets.Arkusz1.A1.v, workbook.Sheets.Arkusz1.A2.v, workbook.Sheets.Arkusz1.B2.v]);
  check("3. dane Arkusz1 zostają", kept.join() === "Produkt,Chleb,5", kept);
  // pusty arkusz: duch kliknięty, zanim kolumna ma nazwę — wpis w A2, nie w wierszu nagłówka
  await dropToasts();
  await tap('.ghost-row td[data-ghost-col="2"]');
  await p.waitForSelector("input.cell-editor", { timeout: 5000 });
  await p.keyboard.type("x");
  await p.keyboard.press("Enter");
  await sleep(400);
  const c2 = await p.evaluate(() => ({ c1: workbook.Sheets.Arkusz2.C1 && workbook.Sheets.Arkusz2.C1.v, c2: workbook.Sheets.Arkusz2.C2 && workbook.Sheets.Arkusz2.C2.v }));
  check("3. pierwszy wpis w pustym arkuszu trafia pod nagłówek (C2)", c2.c2 === "x" && c2.c1 == null, c2);
  // formuła w Arkusz2 → Arkusz1
  await p.evaluate(() => { updateSheetFormulaCell(0, 0, "Arkusz1!B2*2", 10); setDirtyState(true); });
  await p.evaluate(() => { sheetSelect.value = "Arkusz1"; loadBtn.click(); });
  await idle();
  await sleep(400);
  s = await state();
  check("3. przełączenie arkusza nie gasi „niezapisane”", s.dirty && s.cur === "Arkusz1", s);

  // 4. zmiana nazwy Arkusz1 → „Moje dane”
  await dropToasts();
  await sheetOp("rename");
  await dialogInput();
  await p.fill(".app-dialog-box input", "arkusz2");
  await p.keyboard.press("Enter");
  await sleep(200);
  const err = await p.evaluate(() => { const e = document.querySelector(".app-dialog-error"); return e && !e.hidden ? e.textContent : ""; });
  check("4. powtórzona nazwa odrzucona w oknie", /już jest/.test(err), err);
  await p.fill(".app-dialog-box input", "Moje dane");
  await p.keyboard.press("Enter");
  await idle();
  await sleep(500);
  s = await state();
  const f = await p.evaluate(() => workbook.Sheets.Arkusz2.A1 && workbook.Sheets.Arkusz2.A1.f);
  check("4. nowa nazwa, formuła w innym arkuszu poszła za nią", s.names.join() === "Moje dane,Arkusz2" && s.cur === "Moje dane" && f === "'Moje dane'!B2*2", { s, f });

  // 5. zapis
  const out = path.join(os.tmpdir(), `sheet-ops-${tag}.xlsx`);
  fs.writeFileSync(out, Buffer.from(await p.evaluate(async () => Array.from(await buildOutputBytes("xlsx")))));
  if (process.env.EXCEL_DIR) { fs.mkdirSync(process.env.EXCEL_DIR, { recursive: true }); fs.copyFileSync(out, path.join(process.env.EXCEL_DIR, `sheet-ops-app-${tag}.xlsx`)); }
  if (fs.existsSync(PY)) {
    const py = `import openpyxl,sys,json
wb=openpyxl.load_workbook(sys.argv[1])
print(json.dumps({"sheets":wb.sheetnames,"a1":wb["Moje dane"]["A1"].value,"a2":wb["Moje dane"]["A2"].value,"f":wb["Arkusz2"]["A1"].value}, ensure_ascii=False))`;
    let o;
    try { o = JSON.parse(execFileSync(PY, ["-I", "-c", py, out], { encoding: "utf8" })); } catch (e) { o = { err: String(e.stderr || e.message).slice(0, 300) }; }
    check("5. openpyxl: arkusze, dane i formuła", o.sheets && o.sheets.join() === "Moje dane,Arkusz2" && o.a1 === "Produkt" && o.a2 === "Chleb" && o.f === "='Moje dane'!B2*2", o);
  }
  if (ooxmlOn()) {
    const e = ooxml.validate([out])[out] || [];
    check("5. bez błędów schematu", e.length === 0, e.slice(0, 3));
  }

  // 4b. usunięcie „Moje dane” → formuła #ADR!, ostatni arkusz — odmowa
  await dropToasts();
  await sheetOp("delete");
  await p.waitForSelector('.app-dialog-box [data-kind="danger"]', { timeout: 5000 });
  await tap('.app-dialog-box [data-kind="danger"]');
  await idle();
  await sleep(500);
  s = await state();
  const f2 = await p.evaluate(() => workbook.Sheets.Arkusz2.A1 && workbook.Sheets.Arkusz2.A1.f);
  const v2 = await p.evaluate(() => { const c = workbook.Sheets.Arkusz2.A1; return c && [c.t, c.v, c.w]; });
  check("4. wartość formuły przeliczona na błąd (nie stara liczba)", v2 && v2[0] === "e", v2);
  check("4. usunięty, odwołanie → #REF!", s.names.join() === "Arkusz2" && s.cur === "Arkusz2" && /^#REF!/.test(f2 || ""), { s, f2 });
  await dropToasts();
  await sheetOp("delete");
  await sleep(400);
  s = await state();
  check("4. ostatniego arkusza nie da się usunąć", s.names.join() === "Arkusz2", s);

  // 8. Formatuj jako tabelę (po usunięciu zostaje Arkusz2: A1 = formuła #REF!, C2 = „x”)
  await p.evaluate(() => { updateSheetCell(0, 1, { value: 2026, type: "number" }); updateSheetCell(1, 1, { value: 7, type: "number" }); setDirtyState(true); });
  await dropToasts();
  await p.evaluate(() => { renderActiveTable(); setFocusedCell("wide:1", 0, { scroll: true }); }); // wstążka działa na aktywnej komórce
  await sleep(300);
  // telefon: wstążka formatu jest pod „Aa” (arkusz od dołu)
  if (TOUCH) { await tap("#fmtSheetBtn"); await sleep(500); }
  await shot("table-ribbon");
  // pasek formuły: przy ponownym wczytaniu pliku nie miga napis o Wide-to-Long (nagranie 2026-11-19)
  await p.evaluate(() => { const f = document.getElementById("formulaInput"); window.__ph = []; if (f) new MutationObserver(() => window.__ph.push(f.placeholder)).observe(f, { attributes: true, attributeFilter: ["placeholder"] }); });
  await tap("#fmtTableBtn");
  await idle();
  await sleep(800);
  const ph = await p.evaluate(() => window.__ph.filter((x) => /Wide-to-Long/.test(x)));
  check("8. przy tworzeniu tabeli bez mignięcia „Wide-to-Long” w pasku formuły", ph.length === 0, ph);
  const tb = await p.evaluate(() => ({ t: [...new Set(Object.values(currentTables))].map((x) => [x.displayName, x.ref, x.styleName]), h: ["A1", "B1", "C1"].map((r) => { const c = workbook.Sheets.Arkusz2[r]; return c && [c.t, c.v]; }), toast: [...document.querySelectorAll(".toast")].map((x) => x.textContent).join("|") }));
  check("8. Tabela1 na obszarze z danymi (A1:C2), styl jak Excel", JSON.stringify(tb.t) === JSON.stringify([["Tabela1", "A1:C2", "TableStyleMedium2"]]) && /Tabela1/.test(tb.toast), tb);
  check("8. nagłówki jak Excel: formuła i liczba → tekst, pusty → Kolumna1", JSON.stringify(tb.h) === JSON.stringify([["s", "#REF!"], ["s", "2026"], ["s", "Kolumna1"]]), tb.h);
  await dropToasts();
  await p.evaluate(() => setFocusedCell("wide:1", 0, { scroll: true }));
  await sleep(300);
  if (TOUCH && !(await p.evaluate(() => document.documentElement.classList.contains("fmt-sheet-open")))) { await tap("#fmtSheetBtn"); await sleep(500); }
  await tap("#fmtTableBtn");
  await sleep(400);
  const again = await p.evaluate(() => [...document.querySelectorAll(".toast")].map((x) => x.textContent).join("|"));
  check("8. drugi raz na tym samym: komunikat zamiast drugiej tabeli", /już tabela/.test(again), again);
  const outT = path.join(os.tmpdir(), `sheet-ops-table-${tag}.xlsx`);
  fs.writeFileSync(outT, Buffer.from(await p.evaluate(async () => Array.from(await buildOutputBytes("xlsx")))));
  if (process.env.EXCEL_DIR) fs.copyFileSync(outT, path.join(process.env.EXCEL_DIR, `sheet-ops-table-${tag}.xlsx`));
  if (ooxmlOn()) {
    const e = ooxml.validate([outT])[outT] || [];
    check("8. plik z tabelą: bez błędów schematu", e.length === 0, e.slice(0, 3));
  }

  // 6. istniejący plik: przykład + nowy arkusz
  await p.evaluate(() => { hasUnsavedChanges = false; swbStructureDirty = false; return loadSampleFile(); });
  await idle();
  await sleep(600);
  const sampleSheets = await p.evaluate(() => workbook.SheetNames.slice());
  const before = path.join(os.tmpdir(), `sheet-ops-sample-before-${tag}.xlsx`);
  fs.writeFileSync(before, Buffer.from(await p.evaluate(async () => Array.from(await buildOutputBytes("xlsx")))));
  await p.evaluate(() => sheetEditUi.setMode("edit"));
  await dropToasts();
  await sheetOp("add");
  await idle();
  await sleep(500);
  s = await state();
  check("6. przykład: nowy arkusz za bieżącym", s.names.length === sampleSheets.length + 1 && s.names[1] === "Arkusz1", s.names);
  const after = path.join(os.tmpdir(), `sheet-ops-sample-after-${tag}.xlsx`);
  fs.writeFileSync(after, Buffer.from(await p.evaluate(async () => Array.from(await buildOutputBytes("xlsx")))));
  if (process.env.EXCEL_DIR) fs.copyFileSync(after, path.join(process.env.EXCEL_DIR, `sheet-ops-sample-${tag}.xlsx`));
  if (ooxmlOn()) {
    const r = ooxml.validate([before, after]);
    const fresh = ooxml.newErrors(r[before] || [], r[after] || []);
    check("6. przykład z nowym arkuszem: bez NOWYCH błędów schematu", fresh.length === 0, fresh.slice(0, 3));
  }
  const rows = await p.evaluate(() => { sheetSelect.value = workbook.SheetNames[0]; loadBtn.click(); return new Promise((r) => setTimeout(() => r(baseRows.length), 1500)); });
  check("6. dane przykładu zostają", rows >= 100, rows);

  check("7. bez błędów strony i natywnych okien", errors.length === 0, errors.slice(0, 3));
  await browser.close();
  console.log(fails.length ? `\n❌ ${fails.length} nie przeszło` : "\n✅ wszystko przeszło");
  process.exit(fails.length ? 1 : 0);
}

run().catch((e) => { console.error(e); process.exit(1); });
