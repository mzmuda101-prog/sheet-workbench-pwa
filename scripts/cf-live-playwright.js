// cf-live-playwright.js — formatowanie warunkowe na żywo, jak w Excelu.
//
// Zgłoszenie Mateusza (film 2026-11-11): zmiana osoby w grafiku podbijała licznik 1→2
// (LICZ.JEŻELI), ale kolor licznika zostawał z otwarcia pliku; wpisanie nazwiska
// w komórkę „WYSZUKIWARKA” niczego nie podświetlało. Kolory CF liczone były raz na arkusz.
//
// Co musi być prawdą (plik z openpyxl, nagłówek w wierszu 1):
//   B2:B4 = LICZ.JEŻELI($D$2:$D$10;A2), reguła „komórka = 2” → fiolet,
//   D2:D10 reguła „=D2=$G$2” (wyszukiwarka) → zielona ramka + tło,
//   A2:A4 reguła „=CZY.LICZBA(PODAJ.POZYCJĘ(A2;$D:$D;0))” → żółte (cała kolumna, PODAJ.POZYCJĘ —
//   tego mini-silnik CF nie znał),
//   1. zmiana w grafiku → licznik przeliczony I przekolorowany od razu,
//   2. druga zmiana → kolor znika tam, gdzie licznik spadł,
//   3. wpis w wyszukiwarkę → podświetlone pasujące komórki grafiku, zmiana wpisu → inne,
//   4. Cofnij → kolory wracają razem z wartością,
//   5. reguły x14 z <extLst> (Excel 2010+) są czytane: zakres z <xm:sqref>, formuła z <xm:f>, styl z <x14:dxf>.

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
from openpyxl.styles import PatternFill, Border, Side
from openpyxl.formatting.rule import CellIsRule, FormulaRule
wb = Workbook(); ws = wb.active; ws.title = "Grafik"
ws.append(["Osoba", "Ile", None, "Grafik", None, None, "Szukaj"])
for i, (a, d) in enumerate([("Anna", "Anna"), ("Bartek", "Bartek"), ("Celina", "Anna"), (None, "Celina"), (None, "Ewa")]):
    r = i + 2
    if a: ws.cell(r, 1, a); ws.cell(r, 2, f"=COUNTIF($D$2:$D$10,A{r})")
    ws.cell(r, 4, d)
fill = lambda c: PatternFill(start_color=c, end_color=c, fill_type="solid")
ws.conditional_formatting.add("B2:B4", CellIsRule(operator="equal", formula=["2"], fill=fill("B4A7D6")))
side = Side(style="thin", color="00B050")
ws.conditional_formatting.add("D2:D10", FormulaRule(formula=["D2=$G$2"], fill=fill("C6EFCE"), border=Border(left=side, right=side, top=side, bottom=side)))
ws.conditional_formatting.add("A2:A4", FormulaRule(formula=["ISNUMBER(MATCH(A2,$D:$D,0))"], fill=fill("FFFF00")))
wb.save(sys.argv[1])
`, file], { encoding: "utf8" });
  if (py.status !== 0) throw new Error("openpyxl: " + py.stderr);
}

async function run() {
  const tmp = fs.mkdtempSync(path.join(os.tmpdir(), "swb-cf-live-"));
  const FILE = path.join(tmp, "cf-live.xlsx");
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
  await sleep(400);

  const td = (r, c) => `#dataTable tbody tr[data-row-key="wide:${r}"] td[data-col-index="${c}"]`;
  const look = (cells) => page.evaluate((cells) => cells.map(([r, c]) => {
    const el = document.querySelector(`#dataTable tbody tr[data-row-key="wide:${r}"] td[data-col-index="${c}"]`);
    return el ? { t: el.textContent, bg: el.style.background || "", ring: el.style.boxShadow || "" } : null;
  }), cells);
  const type = async (r, c, text) => {
    await page.dblclick(td(r, c));
    await page.waitForSelector(`${td(r, c)} input.cell-editor`, { timeout: 3000 });
    await page.keyboard.press("ControlOrMeta+a");
    await page.keyboard.type(text);
    await page.keyboard.press("Enter");
    await page.keyboard.press("Escape"); // lista podpowiedzi / kolejny edytor
    await sleep(350);
  };
  const purple = (x) => /180, 167, 214/.test(x.bg);
  const yellow = (x) => /255, 255, 0/.test(x.bg);
  const green = (x) => /0, 176, 80/.test(x.ring);

  // 1. D5 Celina → Bartek: Bartek ma 2 (fiolet), Celina 0 (bez żółtego w A)
  await type(4, 3, "Bartek");
  let [b2, b3, b4, a3, a4] = await look([[1, 1], [2, 1], [3, 1], [2, 0], [3, 0]]);
  ok("1. Anna = 2 → fiolet", b2.t === "2" && purple(b2), b2);
  ok("1. Bartek 1→2 → fiolet od razu (bez ponownego otwarcia)", b3.t === "2" && purple(b3), b3);
  ok("1. Celina 0 → bez koloru", b4.t === "0" && !purple(b4), b4);
  ok("1. PODAJ.POZYCJĘ po całej kolumnie: Bartek jest w grafiku → żółte", yellow(a3), a3);
  ok("1. Celina zniknęła z grafiku → żółte znika", !yellow(a4), a4);

  // 2. D2 Anna → Ewa: Anna spada do 1 → fiolet znika
  await type(1, 3, "Ewa");
  [b2] = await look([[1, 1]]);
  ok("2. Anna 2→1 → fiolet znika", b2.t === "1" && !purple(b2), b2);

  // 3. wyszukiwarka G2
  await type(1, 6, "Bartek");
  let [d2, d3, d4, d5] = await look([[1, 3], [2, 3], [3, 3], [4, 3]]);
  ok("3. „Bartek” w wyszukiwarce → D3 i D5 z ramką i tłem", green(d3) && green(d5) && /198, 239, 206/.test(d3.bg), { d3, d5 });
  ok("3. pozostałe bez ramki", !green(d2) && !green(d4), { d2, d4 });
  await type(1, 6, "Ewa");
  [d2, d3] = await look([[1, 3], [2, 3]]);
  ok("3. zmiana wpisu → podświetlenie przechodzi na Ewę", green(d2) && !green(d3), { d2, d3 });

  // 4. Cofnij wpis w wyszukiwarce → znów Bartek
  await page.evaluate(() => { setFocusedCell("wide:1", 0, { scroll: false }); });
  await page.keyboard.press("ControlOrMeta+z");
  await sleep(400);
  [d2, d3] = await look([[1, 3], [2, 3]]);
  ok("4. Cofnij → podświetlenie wraca do Bartka", !green(d2) && green(d3), { d2, d3 });

  // 5. reguła x14 z extLst
  const x14 = await page.evaluate(() => {
    const xml = `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"><sheetData/><extLst><ext uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}"><x14:conditionalFormattings><x14:conditionalFormatting><x14:cfRule type="expression" priority="2" id="{1}"><xm:f>COUNTIF(Inny!$A:$A,A1)&gt;0</xm:f><x14:dxf><font><color rgb="FFFF0000"/></font><fill><patternFill><bgColor rgb="FFFFFF00"/></patternFill></fill></x14:dxf></x14:cfRule><xm:sqref>A1:B5</xm:sqref></x14:conditionalFormatting></x14:conditionalFormattings></ext></extLst></worksheet>`;
    const b = parseSheetCF(xml);
    return b.length === 1 ? { ref: XLSX.utils.encode_range(b[0].ranges[0]), f: b[0].rules[0].formulas, type: b[0].rules[0].type, dxf: b[0].rules[0].dxf } : b;
  });
  ok("5. x14: zakres, formuła i styl z reguły", x14.ref === "A1:B5" && x14.f[0] === "COUNTIF(Inny!$A:$A,A1)>0" && x14.type === "expression" && x14.dxf && /^#?FF0000$/i.test(String(x14.dxf.fontColor).replace("#", "")) && x14.dxf.fillColor, x14);

  await browser.close();
  fs.rmSync(tmp, { recursive: true, force: true });
  if (errors.length) failures.push("błędy strony: " + errors.join(" | "));
  if (failures.length) {
    console.error(`❌ cf-live (${ENGINE}):\n  - ` + failures.join("\n  - "));
    process.exit(1);
  }
  console.log(`✅ cf-live (${ENGINE}): kolory CF przeliczane po każdej zmianie, wyszukiwarka podświetla, PODAJ.POZYCJĘ/całe kolumny, reguły x14`);
}

run().catch((e) => { console.error(e); process.exit(1); });
