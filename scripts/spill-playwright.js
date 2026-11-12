// spill-playwright.js — formuły rozlewane (tablice dynamiczne) przeliczane w aplikacji, jak w Excelu.
//
// Plik z openpyxl: A2:A7 imiona (z powtórzeniami), D2 = UNIQUE(TOCOL(A2:A12;1)) jako formuła tablicowa
// D2:D4 (jak zapis Excela: kotwica + zakres), F2 = COUNTA(D2:D20) (czyta rozlany wynik).
// Co musi być prawdą:
//   1. zmiana źródła → lista przeliczona i rozlana (wydłuża się), widać ją w tabeli,
//   2. formuła czytająca rozlany wynik też się przelicza (drugi przebieg),
//   3. lista się skraca → zbędne komórki wyczyszczone,
//   4. coś stoi na drodze → kotwica #ROZLANIE!, reszta pusta; po usunięciu przeszkody — lista wraca,
//   5. Cofnij przywraca poprzednią listę,
//   6. wpis w środku rozlanej listy → #ROZLANIE! (wpis zostaje), po jego usunięciu lista wraca,
//   7. formuła rozlewana wpisana w aplikacji: rozlewa się od razu, zapis jak Excel 365, Cofnij ją zdejmuje.

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
from openpyxl.worksheet.formula import ArrayFormula
wb = Workbook(); ws = wb.active; ws.title = "Dane"
ws.append(["Osoba", None, None, "Lista", None, "Ile"])
for n in ["Anna", "Bartek", "Anna", "Celina", "Bartek", "Anna"]: ws.append([n])
ws["D2"] = ArrayFormula("D2:D4", "=_xlfn.UNIQUE(_xlfn.TOCOL(A2:A12,1))")
ws["F2"] = "=COUNTA(D2:D20)"
wb.save(sys.argv[1])
`, file], { encoding: "utf8" });
  if (py.status !== 0) throw new Error("openpyxl: " + py.stderr);
}

async function run() {
  const tmp = fs.mkdtempSync(path.join(os.tmpdir(), "swb-spill-"));
  const FILE = path.join(tmp, "spill.xlsx");
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
  await page.evaluate(() => sheetEditUi.setMode("edit"));
  await sleep(300);

  const td = (r, c) => `#dataTable tbody tr[data-row-key="wide:${r}"] td[data-col-index="${c}"]`;
  const type = async (r, c, text) => {
    await page.evaluate(({ r, c }) => setFocusedCell(`wide:${r}`, c, { scroll: true }), { r, c });
    await page.dblclick(td(r, c));
    await page.waitForSelector(`${td(r, c)} input.cell-editor`, { timeout: 3000 });
    await page.keyboard.press("ControlOrMeta+a");
    if (text) await page.keyboard.type(text); else await page.keyboard.press("Backspace");
    await page.keyboard.press("Tab");
    await sleep(500);
    await page.keyboard.press("Escape");
    await sleep(150);
  };
  const state = () => page.evaluate(() => {
    const ws = workbook.Sheets[currentSheetName];
    const list = [];
    for (let r = 2; r <= 9; r++) { const x = ws["D" + r]; list.push(x && x.t !== "z" && x.v != null ? (x.t === "e" ? x.w || x.v : x.v) : null); }
    const shown = [];
    for (let r = 1; r <= 8; r++) { const el = document.querySelector(`#dataTable tbody tr[data-row-key="wide:${r}"] td[data-col-index="3"]`); shown.push(el ? el.textContent : null); }
    return { F: ws.D2.F, list, shown, ile: ws.F2 && ws.F2.v };
  });

  // 1. + 2.
  await type(2, 0, "Darek"); // A3: Bartek → Darek  (Anna, Darek, Anna, Celina, Bartek, Anna)
  const s1 = await state();
  ok("1. lista przeliczona i rozlana na 4 wiersze", s1.F === "D2:D5" && s1.list.slice(0, 5).join() === "Anna,Darek,Celina,Bartek,", s1);
  ok("1. lista widoczna w tabeli", s1.shown.slice(0, 4).join() === "Anna,Darek,Celina,Bartek", s1.shown);
  ok("2. formuła czytająca rozlany wynik przeliczona (ILE.NIEPUSTYCH = 4)", s1.ile === 4, s1.ile);

  // 3. skracanie
  await type(4, 0, "Anna"); // A5: Celina → Anna
  await type(5, 0, "Anna"); // A6: Bartek → Anna  (Anna, Darek, Anna, Anna, Anna, Anna)
  const s3 = await state();
  ok("3. lista skrócona do 2, reszta wyczyszczona", s3.F === "D2:D3" && s3.list.join() === "Anna,Darek,,,,,,", s3);
  ok("3. ILE.NIEPUSTYCH = 2", s3.ile === 2, s3.ile);

  // 4. przeszkoda
  await type(4, 3, "X"); // D5 = X (D2:D5 potrzebne po dopisaniu 2 osób)
  await type(6, 0, "Ela");
  await type(5, 0, "Franek"); // lista: Anna, Darek, Franek, Ela → D2:D5, a D5 zajęte
  const s4 = await state();
  ok("4. przeszkoda w zakresie → kotwica #ROZLANIE!, reszta pusta", /#(ROZLANIE|SPILL)!/.test(String(s4.list[0])) && s4.list.slice(1, 3).every((x) => x == null) && s4.list[3] === "X", s4);
  ok("4. w tabeli widać #ROZLANIE!", /#ROZLANIE!/.test(String(s4.shown[0])), s4.shown);
  await type(4, 3, ""); // usuń przeszkodę
  const s4b = await state();
  ok("4. po usunięciu przeszkody lista wraca", s4b.F === "D2:D5" && s4b.list.slice(0, 4).join() === "Anna,Darek,Franek,Ela", s4b);

  // 5. Cofnij usunięcia przeszkody → znów #ROZLANIE!; Cofnij dalej → wcześniejsza lista
  await page.evaluate(() => undoLast()); await sleep(500);
  const s5 = await state();
  ok("5. Cofnij (przeszkoda wraca) → znów #ROZLANIE!", /#(ROZLANIE|SPILL)!/.test(String(s5.list[0])) && s5.list[3] === "X", s5);

  // 6. wpis w środku rozlanej listy (jak w Excelu) → #ROZLANIE!, wpis zostaje; usunięty → lista wraca
  await page.evaluate(() => redoLast()); await sleep(500); // przeszkoda X znów usunięta → lista D2:D5
  await type(2, 3, "Y"); // D3
  const s6 = await state();
  ok("6. wpis w środku listy → #ROZLANIE!, wpis zostaje", /#(ROZLANIE|SPILL)!/.test(String(s6.list[0])) && s6.list[1] === "Y" && s6.list[2] == null, s6);
  await type(2, 3, "");
  const s6b = await state();
  ok("6. wpis usunięty → lista wraca", s6b.F === "D2:D5" && s6b.list.slice(0, 4).join() === "Anna,Darek,Franek,Ela", s6b);

  // 7. formuła rozlewana WPISANA w aplikacji (po polsku): rozlewa się od razu, Cofnij ją zdejmuje,
  //    zapis = tablica dynamiczna jak w Excelu (<f t="array" ref>, cm, metadata.xml)
  await type(1, 1, "=SORTUJ(UNIKATOWE(DO.KOLUMNY(A2:A12;1));;-1)"); // B2 (pusta kolumna)
  const s7 = await page.evaluate(() => { const ws = workbook.Sheets[currentSheetName]; return { F: ws.B2 && ws.B2.F, vals: ["B2", "B3", "B4", "B5", "B6"].map((k) => (ws[k] && ws[k].t !== "z" ? ws[k].v : null)), shown: [1, 2, 3, 4].map((r) => document.querySelector(`#dataTable tbody tr[data-row-key="wide:${r}"] td[data-col-index="1"]`)?.textContent) }; });
  ok("7. wpisana formuła rozlewa się od razu (malejąco)", s7.F === "B2:B5" && s7.vals.join() === "Franek,Ela,Darek,Anna,", s7);
  ok("7. rozlany wynik widać w tabeli", s7.shown.join() === "Franek,Ela,Darek,Anna", s7.shown);
  const xml7 = await page.evaluate(async () => {
    const bytes = await buildOutputBytes("xlsx");
    const zip = await JSZip.loadAsync(bytes);
    const sheet = await zip.file(Object.keys(zip.files).find((f) => /xl\/worksheets\/sheet\d+\.xml$/.test(f))).async("string");
    return { cell: (sheet.match(/<c r="B2"[^>]*>.*?<\/c>/) || [""])[0], meta: zip.file("xl/metadata.xml") ? await zip.file("xl/metadata.xml").async("string") : "", ct: await zip.file("[Content_Types].xml").async("string"), rels: await zip.file("xl/_rels/workbook.xml.rels").async("string") };
  });
  ok("7. zapis: <f t=\"array\" ref=\"B2:B5\"> i cm=\"1\", funkcje z _xlfn.", /cm="1"/.test(xml7.cell) && /<f t="array" ref="B2:B5">_xlfn\.SORT\(_xlfn\.UNIQUE\(_xlfn\.TOCOL\(A2:A12,1\)\),,-1\)<\/f>/.test(xml7.cell), xml7.cell);
  ok("7. zapis: metadata.xml z XLDAPR, wpis w [Content_Types] i relacja", /XLDAPR/.test(xml7.meta) && /fDynamic="1"/.test(xml7.meta) && /\/xl\/metadata\.xml/.test(xml7.ct) && /sheetMetadata/.test(xml7.rels), { meta: xml7.meta.slice(0, 120) });
  await page.evaluate(() => undoLast()); await sleep(400);
  const s7b = await page.evaluate(() => { const ws = workbook.Sheets[currentSheetName]; return ["B2", "B3", "B4", "B5"].map((k) => (ws[k] && ws[k].t !== "z" ? ws[k].v : null)); });
  ok("7. Cofnij zdejmuje formułę razem z rozlanym wynikiem", s7b.every((x) => x == null || x === ""), s7b);

  await browser.close();
  fs.rmSync(tmp, { recursive: true, force: true });
  if (errors.length) failures.push("błędy strony: " + errors.join(" | "));
  if (failures.length) {
    console.error(`❌ spill (${ENGINE}):\n  - ` + failures.join("\n  - "));
    process.exit(1);
  }
  console.log(`✅ spill (${ENGINE}): rozlewanie, skracanie, #ROZLANIE!, formuły zależne, Cofnij`);
}

run().catch((e) => { console.error(e); process.exit(1); });
