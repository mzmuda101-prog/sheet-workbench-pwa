// col-grow-playwright.js — nowe kolumny z prawej (app/row-grow.js + app/xlsx-grow.js, Etap 1b).
//   node scripts/col-grow-playwright.js            (Chromium)
//   ENGINE=webkit node scripts/col-grow-playwright.js
//
// Pilnuje:
//   1. „＋ Nowa kolumna” tylko w Edycji; okno nazwy: domyślna „Kolumna 20”, duplikat i znaki
//      składni szukania odrzucone (okno zostaje), Anuluj = nic się nie zmienia;
//   2. dodanie: nagłówek w arkuszu, wygląd nagłówka i komórek jak w kolumnie po lewej, filtr
//      i ręczna szerokość kolumny zostają, fokus w nowej kolumnie — pisanie trafia do niej;
//   3. jedno Cofnij usuwa kolumnę (filtr zostaje, bez błędów strony), Ponów przywraca;
//   4. zapis: bez błędów schematu (Open XML SDK), openpyxl czyta nagłówek, wartość i wygląd;
//   5. (Chromium) tabela Excela: w aplikacji i w pliku rośnie w bok, <tableColumn> z nazwą
//      z arkusza (& escapowane), openpyxl widzi kolumnę tabeli; razem z nowym wierszem —
//      tabela rośnie w obie strony; bez NOWYCH błędów schematu.

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
const ooxmlOn = () => !!(ooxml && ooxml.available && ooxml.available());

// EXCEL_DIR=/private/tmp/swb-excel-check: kopia zapisanych plików do sprawdzenia w prawdziwym Excelu
// (scripts/excel-real-check.py, npm run test:excel)
const toExcelDir = (file, name) => { if (process.env.EXCEL_DIR) { fs.mkdirSync(process.env.EXCEL_DIR, { recursive: true }); fs.copyFileSync(file, path.join(process.env.EXCEL_DIR, `col-grow-${ENGINE_NAME}-${name}.xlsx`)); } };

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

  const dialogOpen = () => p.evaluate(() => !!document.querySelector(".app-dialog-box") && document.querySelector(".app-dialog-box").offsetParent !== null);
  const openDialog = async () => { await p.click("#addColBtn"); await p.waitForSelector(".app-dialog-box input", { state: "visible", timeout: 5000 }); };
  const headers = () => p.evaluate(() => currentHeaders.slice());

  // 1. przycisk i okno nazwy
  check("1. przycisk widoczny w Edycji", await p.evaluate(() => !addColBtn.hidden && addColBtn.offsetParent !== null));
  await p.evaluate(() => sheetEditUi.setMode("view"));
  check("1. w Podglądzie przycisku nie ma", await p.evaluate(() => addColBtn.offsetParent === null));
  await p.evaluate(() => sheetEditUi.setMode("edit"));
  const h0 = await headers();
  await openDialog();
  check("1. domyślna nazwa „Kolumna 20”", (await p.inputValue(".app-dialog-box input")) === "Kolumna 20", await p.inputValue(".app-dialog-box input"));
  await p.fill(".app-dialog-box input", "Długość"); await p.keyboard.press("Enter"); await sleep(200);
  const dupErr = await p.evaluate(() => document.querySelector(".app-dialog-error")?.textContent || "");
  check("1. duplikat odrzucony, okno zostaje", (await dialogOpen()) && /już jest/.test(dupErr), dupErr);
  await p.fill(".app-dialog-box input", "a:b"); await p.keyboard.press("Enter"); await sleep(200);
  check("1. znaki składni szukania odrzucone", await dialogOpen());
  await p.keyboard.press("Escape"); await sleep(300);
  check("1. Anuluj = nic się nie zmienia", !(await dialogOpen()) && JSON.stringify(await headers()) === JSON.stringify(h0));

  // 2. dodanie z filtrem i ręczną szerokością
  await p.evaluate(() => { quickSearchEl.value = "Nowak"; applyQuickSearch(); manualColumnWidths[3] = 222; renderActiveTable(); });
  await sleep(500);
  const before = await p.evaluate(() => ({ view: viewRows.length, real: swbRealHeaderCount() }));
  await openDialog();
  await p.fill(".app-dialog-box input", "Uwagi");
  await p.keyboard.press("Enter");
  await p.waitForFunction(() => currentHeaders.includes("Uwagi"), null, { timeout: 20000 });
  await sleep(500);
  const s2 = await p.evaluate((real) => {
    const ws = workbook.Sheets[currentSheetName];
    const sm = currentStyleIndexMap.get(currentSheetName);
    const L = XLSX.utils.encode_col(currentStartCol + real - 1);
    const N = XLSX.utils.encode_col(currentStartCol + real);
    const hr = currentHeaderRow;
    return {
      hdr: ws[`${N}${hr}`] && ws[`${N}${hr}`].v,
      xfHdr: [sm.get(`${L}${hr}`), sm.get(`${N}${hr}`)],
      xfData: [sm.get(`${L}${hr + 1}`), sm.get(`${N}${hr + 1}`)],
      last: currentHeaders[real],
      view: viewRows.length, q: quickSearchEl.value, w3: manualColumnWidths[3],
      focusCol: focusedCellState && focusedCellState.colIndex0,
    };
  }, before.real);
  check("2. nagłówek „Uwagi” w arkuszu i w widoku", s2.hdr === "Uwagi" && s2.last === "Uwagi", s2);
  check("2. wygląd nagłówka i komórek jak w kolumnie po lewej", s2.xfHdr[0] != null && s2.xfHdr[0] === s2.xfHdr[1] && s2.xfData[0] === s2.xfData[1], s2);
  check("2. filtr i ręczna szerokość zostają", s2.view === before.view && s2.q === "Nowak" && s2.w3 === 222, { s2, before });
  await p.keyboard.type("notatka 1"); await p.keyboard.press("Enter"); await sleep(300);
  const typed = await p.evaluate((real) => {
    const ws = workbook.Sheets[currentSheetName];
    const N = XLSX.utils.encode_col(currentStartCol + real);
    return Object.keys(ws).filter((k) => new RegExp(`^${N}\\d+$`).test(k) && ws[k].v === "notatka 1").length;
  }, before.real);
  check("2. fokus w nowej kolumnie — pisanie trafia do niej", s2.focusCol === before.real && typed === 1, { focusCol: s2.focusCol, typed });

  // 4. zapis (zanim cofniemy)
  const tmp = fs.mkdtempSync(path.join(os.tmpdir(), "col-grow-"));
  const out = path.join(tmp, "sample.xlsx");
  fs.writeFileSync(out, Buffer.from(await p.evaluate(async () => Array.from(await buildOutputBytes("xlsx")))));
  toExcelDir(out, "przyklad");
  if (ooxmlOn()) {
    const res = ooxml.validate([out]);
    check("4. plik bez błędów schematu (Open XML SDK)", !(res[0].errors || []).length, (res[0].errors || []).slice(0, 3));
  } else console.log("⚠️  4. walidator Open XML niedostępny — pomijam");
  if (fs.existsSync(PY)) {
    const py = `import openpyxl,sys,json
ws=openpyxl.load_workbook(sys.argv[1])['Obieg zakresów']
vals=[ws.cell(row=r,column=20).value for r in range(5,126)]
h=ws['T5']; s=ws['S5']
print(json.dumps({'hdr':h.value,'bold':[h.font.b,s.font.b],'fill':[h.fill.fgColor.rgb,s.fill.fgColor.rgb],'note':'notatka 1' in vals,'fmt':[ws['T7'].number_format,ws['S7'].number_format]}))`;
    let o = null;
    try { o = JSON.parse(execFileSync(PY, ["-I", "-c", py, out], { encoding: "utf8" })); } catch (e) { o = { err: String(e.stderr || e.message).slice(0, 300) }; }
    check("4. openpyxl: nagłówek, wartość i wygląd jak kolumna obok", o && o.hdr === "Uwagi" && o.note && o.bold[0] === o.bold[1] && o.fill[0] === o.fill[1] && o.fmt[0] === o.fmt[1], o);
  }

  // 3. Cofnij / Ponów
  await p.evaluate(() => undoLast()); await sleep(200); // wpis
  await p.evaluate(() => undoLast()); // kolumna
  await p.waitForFunction(() => !currentHeaders.includes("Uwagi"), null, { timeout: 20000 }).catch(() => {});
  await sleep(400);
  const s3 = await p.evaluate(() => ({ n: swbRealHeaderCount(), q: quickSearchEl.value, view: viewRows.length }));
  check("3. jedno Cofnij usuwa kolumnę, filtr zostaje", s3.n === before.real && s3.q === "Nowak" && s3.view === before.view, s3);
  await p.evaluate(() => redoLast());
  await p.waitForFunction(() => currentHeaders.includes("Uwagi"), null, { timeout: 20000 }).catch(() => {});
  check("3. Ponów przywraca kolumnę", (await headers()).includes("Uwagi"));

  // 5. tabela Excela (duży plik — tylko Chromium)
  if (ENGINE_NAME === "chromium") {
    await p.evaluate(() => { hasUnsavedChanges = false; setDirtyState(false); });
    await p.setInputFiles("#fileInput", RODO);
    await p.waitForFunction(() => typeof workbook !== "undefined" && workbook && currentFileName.startsWith("RODO") && sheetSelect.options.length > 0, null, { timeout: 60000 });
    await p.evaluate(() => { sheetSelect.value = "2025-2026"; loadBtn.click(); });
    await p.waitForFunction(() => currentSheetName === "2025-2026" && document.querySelector("#dataTable tbody tr[data-row-key]"), null, { timeout: 60000 });
    await sleep(1000);
    await p.evaluate(() => { sheetEditUi.setMode("edit"); document.querySelectorAll(".toast").forEach((t) => t.remove()); });
    await openDialog();
    await p.fill(".app-dialog-box input", "Uwagi & nowe");
    await p.keyboard.press("Enter");
    await p.waitForFunction(() => currentHeaders.includes("Uwagi & nowe"), null, { timeout: 30000 });
    await sleep(400);
    await p.keyboard.type("pierwsza uwaga"); await p.keyboard.press("Enter"); await sleep(300);
    const s5 = await p.evaluate(() => ({ tbl: currentTables["opracowanie_terenow"].ref, last: currentTables["opracowanie_terenow"].colNames.slice(-1)[0] }));
    check("5. tabela w aplikacji rośnie w bok (A2:AH565, kolumna w tabeli)", s5.tbl === "A2:AH565" && s5.last === "Uwagi & nowe", s5);
    // i nowy wiersz — tabela w obie strony
    await p.click("#addRowBtn"); await sleep(300);
    await p.keyboard.type("Nowy teren"); await p.keyboard.press("Tab"); await sleep(300);
    const rodoOut = path.join(tmp, "rodo.xlsx");
    fs.writeFileSync(rodoOut, Buffer.from(await p.evaluate(async () => Array.from(await buildOutputBytes("xlsx")))));
    toExcelDir(rodoOut, "rodo");
    const JSZip = require(path.join(ROOT, "lib", "jszip.min.js"));
    const zip = await JSZip.loadAsync(fs.readFileSync(rodoOut));
    const tables = await Promise.all(Object.keys(zip.files).filter((n) => /^xl\/tables\/table\d+\.xml$/.test(n)).map((n) => zip.file(n).async("string")));
    const t = tables.find((x) => /name="opracowanie_terenow"/.test(x)) || "";
    check("5. plik: tabela A2:AH566 (w bok i w dół) z filtrem", /\sref="A2:AH566"/.test(t) && /<autoFilter ref="A2:AH566"/.test(t), (t.match(/ ref="[^"]*"/g) || []).join(" "));
    check("5. plik: <tableColumn> z nazwą (& escapowane), count +1", /<tableColumn id="\d+" name="Uwagi &amp; nowe"\/><\/tableColumns>/.test(t) && /<tableColumns count="34"/.test(t));
    if (fs.existsSync(PY)) {
      let o = null;
      try { o = JSON.parse(execFileSync(PY, ["-I", "-c", "import openpyxl,sys,json\nws=openpyxl.load_workbook(sys.argv[1])['2025-2026']\nt=ws.tables['opracowanie_terenow']\nprint(json.dumps({'ref':t.ref,'last':t.tableColumns[-1].name,'AH2':ws['AH2'].value}))", rodoOut], { encoding: "utf8" })); } catch (e) { o = { err: String(e.stderr || e.message).slice(0, 300) }; }
      check("5. openpyxl widzi kolumnę tabeli i nagłówek", o && o.ref === "A2:AH566" && o.last === "Uwagi & nowe" && o.AH2 === "Uwagi & nowe", o);
    }
    if (ooxmlOn()) {
      const [b0, a0] = ooxml.validate([RODO, rodoOut]);
      const fresh = ooxml.newErrors(b0, a0);
      check("5. bez NOWYCH błędów schematu względem oryginału", fresh.length === 0, fresh.slice(0, 3));
    }
  }

  check("bez błędów strony", errors.length === 0, errors.slice(0, 5));
  await browser.close();
  if (fails.length) { console.log(`❌ col-grow (${ENGINE_NAME}): ${fails.length} niezaliczone`); process.exit(1); }
  console.log(`✅ col-grow (${ENGINE_NAME}): wszystko przeszło`);
}
run().catch((e) => { console.error(e); process.exit(1); });
