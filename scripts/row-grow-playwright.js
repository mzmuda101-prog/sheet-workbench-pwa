// row-grow-playwright.js — nowe wiersze pod danymi (app/row-grow.js + app/xlsx-grow.js, Etap 1a).
//   node scripts/row-grow-playwright.js            (Chromium)
//   ENGINE=webkit node scripts/row-grow-playwright.js
//
// Pilnuje (przykład „Obieg zadań”: 4 kolumny „Długość” to formuły, kolumny dat mają format daty):
//   1. „＋ Nowy wiersz” tylko w Edycji; klik = wiersz za ostatnim wierszem danych, pisanie od 1. kolumny;
//   2. wpis: wartości w arkuszu, wygląd z wiersza wyżej (indeks stylu w pustych kolumnach też),
//      formuły „Długość” przeniesione i liczą się z dat wpisanych w nowym wierszu;
//   3. Enter w ostatnim wierszu = kolejny nowy wiersz;
//   4. jedno Cofnij cofa cały nowy wiersz (wartość, formuły, wygląd); ponowny wpis znów daje formuły;
//      Delete w pustym nowym wierszu niczego nie dopisuje;
//   5. zapis: plik bez błędów schematu (Open XML SDK, jeśli jest .NET), openpyxl czyta nowy wiersz
//      z formułą i formatem daty, pusty wiersz roboczy NIE trafia do pliku;
//   6. (Chromium) plik z tabelą Excela: tabela, jej filtr, formatowanie warunkowe i listy wyboru
//      rosną; formuły [#This Row] w nowym wierszu bez #REF!; lista wyboru działa od razu;
//   7. tabela z wierszem sumy na końcu danych: komunikat zamiast wiersza pod sumą.

const { chromium, webkit } = require("playwright");
const fs = require("fs");
const os = require("os");
const path = require("path");
const { execFileSync } = require("child_process");

const APP_URL = process.env.APP_URL || "http://127.0.0.1:4175/";
const ENGINE_NAME = process.env.ENGINE === "webkit" ? "webkit" : "chromium";
const ENGINE = ENGINE_NAME === "webkit" ? webkit : chromium;
const ROOT = path.join(__dirname, "..");
const RODO = path.join(__dirname, "RODO Obieg terenów 2026 V5.xlsx");
const PY = path.join(ROOT, ".venv", "bin", "python");
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
let ooxml = null;
try { ooxml = require("./ooxml-validate"); } catch { ooxml = null; }

// EXCEL_DIR=/private/tmp/swb-excel-check: kopia zapisanych plików do sprawdzenia w prawdziwym Excelu
// (scripts/excel-real-check.py, npm run test:excel)
const toExcelDir = (file, name) => { if (process.env.EXCEL_DIR) { fs.mkdirSync(process.env.EXCEL_DIR, { recursive: true }); fs.copyFileSync(file, path.join(process.env.EXCEL_DIR, `row-grow-${ENGINE_NAME}-${name}.xlsx`)); } };

async function run() {
  const fails = [];
  const check = (name, ok, detail) => {
    console.log(`${ok ? "✅" : "❌"} ${name}${ok ? "" : ` — ${JSON.stringify(detail)}`}`);
    if (!ok) fails.push(name);
  };
  const browser = await ENGINE.launch({ headless: true });
  const ctx = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1280, height: 900 } });
  await ctx.addInitScript(() => { localStorage.setItem("introPlayed", "true"); });
  const p = await ctx.newPage();
  const errors = [];
  p.on("pageerror", (e) => errors.push(e.message));
  p.on("dialog", (d) => { errors.push(`natywne okno: ${d.message()}`); d.dismiss(); });
  await p.goto(APP_URL, { waitUntil: "load" });
  await p.evaluate(() => { document.getElementById("heroSplash")?.remove(); return loadSampleFile(); });
  await p.waitForFunction(() => document.querySelector("#dataTable tbody tr[data-row-key]"), null, { timeout: 20000 });
  await sleep(600);
  await p.evaluate(() => { sheetEditUi.setMode("edit"); document.querySelectorAll(".toast").forEach((t) => t.remove()); });

  const cell = (ref) => p.evaluate((r) => { const c = workbook.Sheets[currentSheetName][r]; return c ? { t: c.t, v: c.v instanceof Date ? c.v.toISOString() : c.v, f: c.f || null } : null; }, ref);
  const xfOf = (ref) => p.evaluate((r) => { const m = currentStyleIndexMap.get(currentSheetName); return m ? m.get(r) ?? null : null; }, ref);
  const typeIn = async (text, key) => { await p.keyboard.type(text); await p.keyboard.press(key); await sleep(150); };

  // 7. tabela z wierszem sumy (symulowana na modelu przykładu)
  const s7 = await p.evaluate(() => {
    const lastR = baseRows.reduce((m, r) => Math.max(m, r.rowIndex0), 0); // jeszcze bez wierszy roboczych
    const keep = currentTables;
    currentTables = { fake: { sheetName: currentSheetName, ref: `A5:S${lastR + 1}`, totalsRows: 1, columns: {}, colNames: [] } };
    document.querySelectorAll(".toast").forEach((t) => t.remove());
    const n0 = baseRows.length;
    const k = swbRowGrow.addDraft({ col: 0 });
    const msg = [...document.querySelectorAll(".toast")].map((t) => t.textContent).join(" | ");
    currentTables = keep;
    return { k, added: baseRows.length - n0, msg };
  });
  check("7. tabela z wierszem sumy: komunikat, bez wiersza pod sumą", !s7.k && s7.added === 0 && /wierszem sumy/.test(s7.msg), s7);

  // 1. przycisk
  const last = await p.evaluate(() => baseRows[baseRows.length - 1].rowIndex0 + 1); // 1-based
  const R = last + 1; // nowy wiersz (1-based)
  check("1. przycisk widoczny w Edycji", await p.evaluate(() => !addRowBtn.hidden && addRowBtn.offsetParent !== null));
  await p.evaluate(() => sheetEditUi.setMode("view"));
  check("1. w Podglądzie przycisku nie ma", await p.evaluate(() => addRowBtn.offsetParent === null));
  await p.evaluate(() => sheetEditUi.setMode("edit"));
  await p.click("#addRowBtn");
  await sleep(300);
  const s1 = await p.evaluate(() => ({ focus: focusedCellState && `${focusedCellState.rowKey}:${focusedCellState.colIndex0}`, editor: !!document.querySelector("input.cell-editor") }));
  check("1. klik = nowy wiersz za danymi, pisanie w 1. kolumnie", s1.focus === `wide:${R - 1}:0` && s1.editor, { s1, R });

  // 2. wpis: Nr, Zakres, Status, Imię, od, do
  await typeIn("121", "Tab");
  await typeIn("Zakres T-01", "Tab");
  await typeIn("W trakcie", "Tab");
  await typeIn("Test Testowy", "Tab");
  await typeIn("01.03.2026", "Tab");
  await typeIn("11.03.2026", "Tab");
  await sleep(300);
  const vals = await Promise.all([`A${R}`, `B${R}`, `E${R}`, `F${R}`].map(cell));
  check("2. wartości w arkuszu", vals[0]?.v === 121 && vals[1]?.v === "Zakres T-01" && !!vals[2] && !!vals[3], vals);
  const formulas = await Promise.all(["G", "K", "O", "S"].map((c) => cell(`${c}${R}`)));
  check("2. formuły „Długość” przeniesione (G, K, O, S)", formulas.every((f) => f && f.f && f.f.includes(`${R}`)), formulas.map((f) => f && f.f));
  check("2. „Długość” liczy się z dat nowego wiersza (10 dni)", formulas[0] && Number(formulas[0].v) === 10, formulas[0]);
  const xfs = await Promise.all([[`E${last}`, `E${R}`], [`I${last}`, `I${R}`], [`G${last}`, `G${R}`]].map(async ([a, b]) => [await xfOf(a), await xfOf(b)]));
  check("2. wygląd jak wiersz wyżej (także pusta kolumna dat I)", xfs.every(([a, b]) => a != null && a === b), xfs);

  // 3. Enter w ostatnim wierszu
  await p.evaluate((k) => setFocusedCell(k, 1, { scroll: true }), `wide:${R - 1}`);
  await p.keyboard.press("F2"); await sleep(100); // edycja: F2 (Enter na komórce schodzi w dół, 2026-11-14)
  await p.keyboard.press("ControlOrMeta+a"); await p.keyboard.type("Zakres T-01 b"); await p.keyboard.press("Enter"); await sleep(300); // F2 stawia kursor na końcu — zastępujemy cały wpis
  const s3 = await p.evaluate(() => focusedCellState && focusedCellState.rowKey);
  check("3. Enter w ostatnim wierszu = kolejny nowy wiersz", s3 === `wide:${R}`, s3);

  // 4. Cofnij / ponowny wpis / Delete
  await p.keyboard.type("Drugi"); await p.keyboard.press("Enter"); await sleep(300);
  const before4 = await cell(`G${R + 1}`);
  // jak w Excelu: 1. Cofnij zdejmuje tylko automat (formuły, wygląd z góry) — wpis zostaje; 2. cofa wpis
  await p.evaluate(() => undoLast()); await sleep(300);
  const mid4 = await Promise.all([`B${R + 1}`, `G${R + 1}`, `E${R + 1}`].map(cell));
  check("4. 1. Cofnij: wpis zostaje, formuła i wygląd z góry znikają", !!(before4 && before4.f) && mid4[0] && mid4[0].v === "Drugi" && !mid4[1] && !mid4[2], { before4, mid4 });
  await p.evaluate(() => undoLast()); await sleep(300);
  const after4 = await Promise.all([`B${R + 1}`, `G${R + 1}`, `E${R + 1}`].map(cell));
  check("4. 2. Cofnij: nowy wiersz cały cofnięty", after4.every((c) => !c), { after4 });
  await p.evaluate(() => redoLast()); await sleep(300);
  await p.evaluate(() => redoLast()); await sleep(300);
  const redo4 = await Promise.all([`B${R + 1}`, `G${R + 1}`].map(cell));
  check("4. Ponów ×2 przywraca wpis i formułę", redo4[0] && redo4[0].v === "Drugi" && redo4[1] && !!redo4[1].f, redo4);
  await p.evaluate(() => { undoLast(); undoLast(); }); await sleep(300);
  await p.evaluate((k) => setFocusedCell(k, 1, { scroll: true }), `wide:${R}`);
  await p.keyboard.press("F2"); await sleep(100); // edycja: F2 (Enter na komórce schodzi w dół, 2026-11-14)
  await p.keyboard.type("Drugi raz"); await p.keyboard.press("Tab"); await sleep(300);
  const again = await cell(`G${R + 1}`);
  check("4. po Cofnij ponowny wpis znów dostaje formuły", !!(again && again.f), again);
  // Delete w pustym wierszu roboczym (R+2)
  await p.evaluate((k) => setFocusedCell(k, 2, { scroll: true }), `wide:${R}`);
  await p.keyboard.press("F2"); await sleep(100); // edycja: F2 (Enter na komórce schodzi w dół, 2026-11-14)
  await p.keyboard.press("Enter"); await sleep(300); // pusty wpis w R+1 → Enter → nowy wiersz R+2
  const draftKey = await p.evaluate(() => focusedCellState && focusedCellState.rowKey);
  await p.keyboard.press("Delete"); await sleep(300);
  const del = await p.evaluate((r) => Object.keys(pendingEdits[currentSheetName] || {}).filter((k) => k.endsWith(String(r))).length, R + 2);
  check("4. Delete w pustym nowym wierszu niczego nie dopisuje", draftKey === `wide:${R + 1}` && del === 0, { draftKey, del });

  // 5. zapis
  const tmp = fs.mkdtempSync(path.join(os.tmpdir(), "row-grow-"));
  const out = path.join(tmp, "sample.xlsx");
  fs.writeFileSync(out, Buffer.from(await p.evaluate(async () => Array.from(await buildOutputBytes("xlsx")))));
  toExcelDir(out, "przyklad");
  if (ooxml && ooxml.available && ooxml.available()) {
    const res = ooxml.validate([out]);
    const errs = res && res[0] ? res[0].errors || [] : [];
    check("5. plik bez błędów schematu (Open XML SDK)", errs.length === 0, errs.slice(0, 3));
  } else console.log("⚠️  5. walidator Open XML niedostępny (brak .NET) — pomijam");
  if (fs.existsSync(PY)) {
    const py = `import openpyxl,sys,json
ws=openpyxl.load_workbook(sys.argv[1])['Obieg zadań']
R=int(sys.argv[2])
print(json.dumps({'B':ws['B%d'%R].value,'G':ws['G%d'%R].value,'Efmt':ws['E%d'%R].number_format,'Ifmt':ws['I%d'%R].number_format,'Iprev':ws['I%d'%(R-1)].number_format,'max':ws.max_row,'B2':ws['B%d'%(R+1)].value}))`;
    let o = null;
    try { o = JSON.parse(execFileSync(PY, ["-I", "-c", py, out, String(R)], { encoding: "utf8" })); } catch (e) { o = { err: String(e.stderr || e.message).slice(0, 300) }; }
    check("5. openpyxl: nowy wiersz z formułą i formatem daty (także w pustej kolumnie)", o && o.B === "Zakres T-01 b" && String(o.G || "").startsWith("=") && /yy/.test(o.Efmt) && o.Ifmt === o.Iprev, o);
    check("5. pusty wiersz roboczy nie trafia do pliku", o && o.max === R + 1 && o.B2 === "Drugi raz", o);
  } else console.log("⚠️  5. brak .venv z openpyxl — pomijam odczyt");

  // 6. prawdziwa tabela Excela (duży plik — tylko Chromium)
  if (ENGINE_NAME === "chromium") {
    await p.evaluate(() => { hasUnsavedChanges = false; setDirtyState(false); });
    await p.setInputFiles("#fileInput", RODO);
    await p.waitForFunction(() => typeof workbook !== "undefined" && workbook && currentFileName.startsWith("RODO") && sheetSelect.options.length > 0, null, { timeout: 60000 });
    await p.evaluate(() => { sheetSelect.value = "2025-2026"; loadBtn.click(); });
    await p.waitForFunction(() => currentSheetName === "2025-2026" && document.querySelector("#dataTable tbody tr[data-row-key]"), null, { timeout: 60000 });
    await sleep(1000);
    await p.evaluate(() => { sheetEditUi.setMode("edit"); document.querySelectorAll(".toast").forEach((t) => t.remove()); });
    await p.click("#addRowBtn"); await sleep(300);
    await p.keyboard.type("Nowy teren"); await p.keyboard.press("Tab"); await sleep(300);
    const s6 = await p.evaluate(() => {
      const ws = workbook.Sheets[currentSheetName];
      const errs = Object.keys(ws).filter((k) => /^[A-Z]+566$/.test(k) && ws[k].t === "e");
      return { tbl: currentTables["opracowanie_terenow"]?.ref, errs, dv: !!getCellDataValidation("2025-2026", 565, 1), f: Object.keys(ws).filter((k) => /^[A-Z]+566$/.test(k) && ws[k].f).length };
    });
    check("6. tabela w aplikacji rośnie, [#This Row] bez #REF!, lista wyboru w nowym wierszu", s6.tbl === "A2:AG566" && s6.errs.length === 0 && s6.dv && s6.f > 0, s6);
    const rodoOut = path.join(tmp, "rodo.xlsx");
    fs.writeFileSync(rodoOut, Buffer.from(await p.evaluate(async () => Array.from(await buildOutputBytes("xlsx")))));
    toExcelDir(rodoOut, "rodo");
    const JSZip = require(path.join(ROOT, "lib", "jszip.min.js"));
    const zip = await JSZip.loadAsync(fs.readFileSync(rodoOut));
    const tables = await Promise.all(Object.keys(zip.files).filter((n) => /^xl\/tables\/table\d+\.xml$/.test(n)).map((n) => zip.file(n).async("string")));
    const t = tables.find((x) => /name="opracowanie_terenow"/.test(x)) || "";
    const sheet = await zip.file("xl/worksheets/sheet2.xml").async("string");
    check("6. plik: tabela i jej filtr do wiersza 566", /\sref="A2:AG566"/.test(t) && /<autoFilter ref="A2:AG566"/.test(t), (t.match(/ ref="[^"]*"/g) || []).join(" "));
    check("6. plik: formatowanie warunkowe i listy wyboru do 566, pojedyncza komórka (P6) bez zmian",
      /<conditionalFormatting sqref="B3:B566 /.test(sheet) && /sqref="AD3:AD566 /.test(sheet) && /sqref="P6"/.test(sheet) && /<dimension ref="A1:AH566"\/>/.test(sheet));
    if (ooxml && ooxml.available && ooxml.available() && typeof ooxml.newErrors === "function") {
      const [before, after] = ooxml.validate([RODO, rodoOut]);
      const fresh = ooxml.newErrors(before, after);
      check("6. plik: bez NOWYCH błędów schematu względem oryginału", !fresh || fresh.length === 0, (fresh || []).slice(0, 3));
    }

    // 8. „Baza danych” (zgłoszenie 2026-11-12): wpis w A, formuły w C i E mają się dopisywać ZA KAŻDYM
    //    razem — także po Cofnij (dawniej kolejny wiersz stawał o jeden niżej, nad nim zostawał pusty
    //    i formuły już się nie kopiowały) i gdy wiersz wyżej formuły nie ma (reguła Excela 3 z 5)
    await p.evaluate(() => { hasUnsavedChanges = false; setDirtyState(false); sheetSelect.value = "Baza danych"; loadBtn.click(); });
    await p.waitForFunction(() => currentSheetName === "Baza danych" && document.querySelector("#dataTable tbody tr[data-row-key]"), null, { timeout: 60000 });
    await sleep(800);
    await p.evaluate(() => { sheetEditUi.setMode("edit"); document.querySelectorAll(".toast").forEach((t) => t.remove()); });
    const R8 = await p.evaluate(() => baseRows.reduce((m, r) => Math.max(m, r.rowIndex0), 0) + 2); // 1-based numer nowego wiersza
    const cells8 = (r) => p.evaluate((r) => ["A", "C", "E"].map((c) => { const x = workbook.Sheets[currentSheetName][c + r]; return x ? (x.f ? "f:" + x.v : x.t === "z" ? null : x.v) : null; }), r);
    const add8 = async (text) => { await p.click("#addRowBtn"); await sleep(300); await p.keyboard.type(text); await p.keyboard.press("Tab"); await sleep(300); await p.keyboard.press("Escape"); await sleep(150); };
    await add8("Kowalski Jan");
    const a8 = await cells8(R8);
    await p.evaluate(() => { undoLast(); undoLast(); }); await sleep(300); // automat, potem wpis
    await p.evaluate(() => { setFocusedCell(getRowSelectionKey(baseRows[0]), 0, { scroll: false }); renderActiveTable(); }); await sleep(300); // pusty wiersz roboczy sprzątnięty
    await add8("Nowak Anna");
    const b8 = await cells8(R8);
    check("8. po Cofnij nowy wiersz wraca w TO SAMO miejsce i znów dostaje formuły", a8[1] === "f:Jan Kowalski" && b8[0] === "Nowak Anna" && b8[1] === "f:Anna Nowak" && b8[2] === "f:A. Nowak" && (await cells8(R8 + 1)).every((x) => x == null), { a8, b8, R8 });
    // wiersz wyżej bez formuły w C → następny nowy wiersz i tak ją dostaje (3 z 5 wierszy wyżej)
    await p.evaluate((r) => { const row = baseRows.find((x) => x.rowIndex0 === r - 1); setFocusedCell(getRowSelectionKey(row), 2, { scroll: true }); cfmtClearContents(); }, R8);
    await sleep(300);
    await add8("Lis Ewa");
    const c8 = await cells8(R8 + 1);
    check("8. wiersz wyżej bez formuły nie zrywa łańcucha (reguła Excela: 3 z 5 wierszy)", c8[0] === "Lis Ewa" && c8[1] === "f:Ewa Lis" && c8[2] === "f:E. Lis", c8);
    // to, co WIDAĆ w tabeli (film 2026-11-11 22:07: wynik był w arkuszu, a komórka pusta)
    const shown8 = await p.evaluate((r) => [...document.querySelectorAll(`#dataTable tbody tr[data-row-key="wide:${r - 1}"] td[data-col-index]`)].slice(0, 5).map((td) => td.textContent), R8 + 1);
    check("8. wynik formuły widać w tabeli także w KOLEJNYM nowym wierszu", shown8[2] === "Ewa Lis" && shown8[4] === "E. Lis", shown8);
    // ręcznie: prawy klik na pojedynczej komórce → „Wypełnij z komórki wyżej” dociąga formułę
    const m8 = await p.evaluate((r) => {
      const row = baseRows.find((x) => x.rowIndex0 === r - 1);
      setFocusedCell(getRowSelectionKey(row), 2, { scroll: true });
      const item = cellMenuFillItems({ rect: rangeCellsToCopy() }).find((x) => x && x.id === "fill-down");
      if (item) item.run();
      return { label: item ? item.label : null };
    }, R8);
    await sleep(300);
    const d8 = await cells8(R8);
    check("8. menu komórki: „Wypełnij z komórki wyżej” na jednej komórce kopiuje formułę", /wyżej|above/.test(m8.label || "") && d8[1] === "f:Anna Nowak", { m8, d8 });
  }

  check("bez błędów strony", errors.length === 0, errors.slice(0, 5));
  await browser.close();
  if (fails.length) { console.log(`❌ row-grow (${ENGINE_NAME}): ${fails.length} niezaliczone`); process.exit(1); }
  console.log(`✅ row-grow (${ENGINE_NAME}): wszystko przeszło`);
}
run().catch((e) => { console.error(e); process.exit(1); });
