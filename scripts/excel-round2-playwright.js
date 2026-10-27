// excel-round2-playwright.js — druga runda „jak w polskim Excelu” (2026-10-17).
//
// Wzorzec: wyniki z Microsoft Excel na Macu przez AppleScript (`formula local`, `sort`,
// `autofill … type fill default`, `fill down`). Co musi być prawdą:
//   1. sortowanie typów mieszanych: liczby < tekst < FAŁSZ < PRAWDA < błędy, puste NA KOŃCU
//      w obu kierunkach, tekst bez wielkości liter z polskimi literami, remisy stabilnie,
//   2. ułamki przy wpisie: „1 1/2” → 1,5 w formacie „# ?/?”, „0 7/16” → „# ??/??”,
//   3. kolor z formatu liczbowego: [Red] ujemne na czerwono, [Blue]/[ColorN]/warunki,
//   4. silnik serii = dokładnie to, co Excel dał przy autofill (≈60 przypadków),
//   5. uchwyt wypełniania myszą: data +1 dzień, „Pozycja 1” → 2, 3…, jedna liczba = kopia
//      (z Alt = seria), dwie liczby = trend, format „zł” idzie z wartością, dwuklik = do
//      końca danych obok; Ctrl+D (tekst „007” zostaje tekstem, formuła przesunięta), Ctrl+R,
//   6. Cofnij zdejmuje całe wypełnienie jednym krokiem.

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

const tmp = fs.mkdtempSync(path.join(os.tmpdir(), "excel-round2-"));
const FILE = path.join(tmp, "runda2.xlsx");

const gen = spawnSync(PY, ["-c", `
import sys, datetime
from openpyxl import Workbook
wb = Workbook(); ws = wb.active; ws.title = "Sort"
ws.append(["Mix", "Idx", "Kolor", "Opis"])
mix = ["zebra", 10, "Ąbc", "abc", True, "#DIV/0!", None, 2, "Łódź", "lody", False, "#N/A", "ćma", "cma", "Zebra", -5, "10", 0, None, "ostatni"]
kol = [(-5, '#,##0.00;[Red]-#,##0.00', "neg"), (5, '#,##0.00;[Red]-#,##0.00', "pos"), (3, '[Blue]0;[Red]0', "blue"), (-3, '[Blue]0;[Red]0', "blueneg"),
       (7, '[Color10]0', "c10"), (50, '[Red][<=100]0;[Blue][>100]0', "le100"), (150, '[Red][<=100]0;[Blue][>100]0', "gt100"), ("abc", '0;0;0;[Magenta]@', "txt"),
       (0, '0;[Red]-0;[Green]0', "zero"), (-2, '0.00', "plain")]
for i, v in enumerate(mix):
    k = kol[i] if i < len(kol) else (None, None, None)
    ws.append([v, i + 1, k[0], k[2]])
    if k[1]: ws.cell(ws.max_row, 3).number_format = k[1]
f = wb.create_sheet("Fill")
f.append(["Liczba", "Data", "Tekst", "Kwota", "Para", "Kod", "Formula", "X", "Pusta", "Nr", "Kw"])
f.append([1, datetime.date(2025, 3, 12), "Pozycja 1", 12.5, 2, "007", "=A2*2", "x", None, 1, "Kw1"])
f.cell(2, 2).number_format = "dd.mm.yyyy"
f.cell(2, 4).number_format = '#,##0.00 "zł"'
f.append([None, None, None, None, 4, None, None, None, None, 2, None])
for n in range(3, 11): f.append([None] * 9 + [n, None])
wb.save(sys.argv[1])
`, FILE], { encoding: "utf8" });
if (gen.status !== 0) { console.error("❌ excel-round2: fixture —", gen.stderr); process.exit(1); }

async function run() {
  const browser = await pw[ENGINE].launch({ headless: true });
  const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1400, height: 900 } });
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
  await page.waitForFunction(() => currentFileName === "runda2.xlsx" && currentSheetName === "Sort" && !!document.querySelector("#dataTable tbody tr[data-row-key]"), null, { timeout: 20000 });
  await sleep(400);

  // 1. sortowanie — kolejność z Excela (Idx = numer wiersza źródła)
  const order = (dir) => page.evaluate((dir) => {
    setPrimarySort("Mix", dir); sortRows(); renderActiveTable();
    return viewRows.map((r) => r.values[1]);
  }, dir);
  const asc = await order("asc");
  const desc = await order("desc");
  const wantAsc = [16, 18, 8, 2, 17, 4, 3, 14, 13, 10, 9, 20, 1, 15, 11, 5, 6, 12, 7, 19];
  const wantDesc = [6, 12, 5, 11, 1, 15, 20, 9, 10, 13, 14, 3, 4, 17, 2, 8, 18, 16, 7, 19];
  ok("1. rosnąco jak Excel: liczby, tekst (a<ą, l<ł, bez wielkości liter), FAŁSZ, PRAWDA, błędy, puste", JSON.stringify(asc) === JSON.stringify(wantAsc), asc);
  ok("1. malejąco jak Excel: błędy, PRAWDA, FAŁSZ, tekst Z→A, liczby, puste dalej NA KOŃCU", JSON.stringify(desc) === JSON.stringify(wantDesc), desc);
  await page.evaluate(() => { setPrimarySort(""); renderActiveTable(); });

  // 3. kolor z formatu liczbowego (w tym samym arkuszu, kolumna Kolor)
  await page.evaluate(() => { applyFilters(); renderActiveTable(); });
  await sleep(150);
  const colors = await page.evaluate(() => {
    const out = {};
    document.querySelectorAll("#dataTable tbody tr[data-row-key]").forEach((tr) => {
      const k = tr.querySelector('td[data-col-index="3"]')?.textContent;
      const td = tr.querySelector('td[data-col-index="2"]');
      if (!k || !td) return;
      out[k] = { cls: Array.from(td.classList).filter((c) => c.startsWith("nf-")).join(" "), color: getComputedStyle(td).color, text: td.textContent };
    });
    return out;
  });
  ok("3. [Red]: ujemna kwota czerwona (rgb 255,0,0), dodatnia bez koloru", colors.neg?.cls === "nf-red" && colors.neg.color === "rgb(255, 0, 0)" && colors.pos?.cls === "", colors);
  ok("3. [Blue]0;[Red]0: dodatnia niebieska, ujemna czerwona", colors.blue?.cls === "nf-blue" && colors.blueneg?.cls === "nf-red", colors);
  ok("3. [Color10] = kolor 10 palety Excela (#008000)", colors.c10?.cls === "nf-color" && colors.c10.color === "rgb(0, 128, 0)", colors.c10);
  ok("3. warunki [<=100]/[>100]: 50 czerwone, 150 niebieskie", colors.le100?.cls === "nf-red" && colors.gt100?.cls === "nf-blue", colors);
  ok("3. sekcja tekstu [Magenta]@ i zero [Green]", colors.txt?.cls === "nf-magenta" && colors.zero?.cls === "nf-green", colors);
  ok("3. format bez koloru nic nie dokłada", colors.plain?.cls === "", colors.plain);

  // 2. ułamki — parser (wyniki Excela) …
  const fr = await page.evaluate(() => Object.fromEntries(["1 1/2", "0 1/2", "-1 1/2", "2 3/8", "1 10/3", "0 7/16", "2 1/100", "1 1/0", "-3/4", "1  1/2", "1 1/2%", "1 1/2 zł", "5 1/4"].map((s) => {
    const r = parseInputValue(s);
    return [s, r ? `${r.type}:${typeof r.value === "number" ? Number(r.value.toPrecision(12)) : r.value instanceof Date ? "date" : r.value}${r.numFmt ? ` [${r.numFmt}]` : ""}` : "null"];
  })));
  const wantFr = {
    "1 1/2": "number:1.5 [# ?/?]", "0 1/2": "number:0.5 [# ?/?]", "-1 1/2": "number:-1.5 [# ?/?]", "2 3/8": "number:2.375 [# ?/?]",
    "1 10/3": "number:4.33333333333 [# ?/?]", "0 7/16": "number:0.4375 [# ??/??]", "2 1/100": "number:2.01 [# ??/??]",
    "1 1/0": "string:1 1/0", "-3/4": "string:-3/4", "1  1/2": "string:1  1/2", "1 1/2%": "number:0.015 [0.00%]", "1 1/2 zł": "number:1.5 [# ?/?]",
    "5 1/4": "number:5.25 [# ?/?]",
  };
  const badFr = Object.keys(wantFr).filter((k) => fr[k] !== wantFr[k]).map((k) => `${k} → ${fr[k]} (Excel: ${wantFr[k]})`);
  ok("2. ułamki rozpoznane jak w Excelu", badFr.length === 0, badFr);

  // 4. silnik serii — wyniki Excela (autofill do 8 komórek)
  const series = await page.evaluate(() => {
    const D = (y, m, d, h = 0, mi = 0) => new Date(y, m - 1, d, h, mi);
    const fmtOut = (o, src) => {
      if (o.copy) { const v = src[o.src].v; return v instanceof Date ? `D${v.getFullYear()}-${v.getMonth() + 1}-${v.getDate()}` : (typeof v === "number" ? Number(v.toPrecision(10)) : v); }
      if (o.value instanceof Date) { const v = o.value; return `D${v.getFullYear()}-${v.getMonth() + 1}-${v.getDate()}` + (v.getHours() || v.getMinutes() ? ` ${v.getHours()}:${v.getMinutes()}` : ""); }
      return typeof o.value === "number" ? Number(o.value.toPrecision(10)) : o.value;
    };
    const cases = {
      one: [[1]], two: [[1, 2]], step2: [[2, 4]], kw1: [["Kw1"]], pon: [["poniedziałek"]], sty_full: [["styczeń"]], sty: [["sty"]], ponShort: [["pon"]],
      date: [[D(2025, 3, 12)]], poz: [["Pozycja 1"]], a1b: [["a1b"]], q1: [["Q1"]], x: [["x"]], time: [[10 / 24], "h:mm"],
      eom: [[D(2025, 1, 31), D(2025, 2, 28)]], mid: [[D(2025, 1, 15), D(2025, 2, 15)]], Pon: [["Pon"]], PON: [["PONIEDZIAŁEK"]], Sty: [["Styczeń"]],
      nd: [["Nd"]], gru: [["grudzień"]], dec: [[1.5]], numx: [[1, "x"]], kw4: [["Kw4"]], pct: [[0.1], "0%"], kwLead: [["1 kw"]], etap: [["Etap 9"]],
      a07: [["A-07"]], neg: [[-1]], trend3: [[1, 3, 4]], szt: [["1 szt"]], r2025: [["2025r"]], fak: [["Faktura 001"]], kwSp: [["Kw 1"]],
      kwartal: [["kwartał 1"]], x13: [["x1", "x3"]], x1x: [["1x"]], poz13: [["Pozycja 1", "Pozycja 3"]], styMar: [["sty", "mar"]], ponSr: [["pon", "śr"]],
      ab: [["a", "b"]], n12x: [[1, 2, "x"]], xPoz: [["x", "Pozycja 1"]], dt: [[D(2025, 3, 12, 8, 30)]], years: [[D(2024, 1, 1), D(2025, 1, 1)]],
      week: [[D(2025, 3, 1), D(2025, 3, 8)]], poz9a: [["Poz. 9 a"]], wtCzw: [["wtorek", "czwartek"]], n24xy: [[2, 4, "x", "y"]], kw13: [["Kw1", "Kw3"]],
      half: [[0.5, 1]], pt: [["pt."]], Wrz: [["Wrz"]], maj: [["maj"]], pozNeg: [["Pozycja -1"]], kwDot: [["Kw. 1"]], Ikw: [["I kwartał"]],
      tydz: [["Tydzień 52"]], ten: [[10]], tenTxt: [["10"]], a15: [["a 1,5"]], x31: [["x3", "x1"]],
    };
    const out = {};
    Object.entries(cases).forEach(([k, [vals, fmt]]) => {
      const src = vals.map((v) => ({ v, fmt: fmt || "General", formula: false }));
      out[k] = swbAutoFillPlan(src, 8, { lang: "pl" }).map((o) => fmtOut(o, src)).join(" ; ");
    });
    return out;
  });
  const want = {
    one: "1 ; 1 ; 1 ; 1 ; 1 ; 1 ; 1", two: "3 ; 4 ; 5 ; 6 ; 7 ; 8", step2: "6 ; 8 ; 10 ; 12 ; 14 ; 16",
    kw1: "Kw2 ; Kw3 ; Kw4 ; Kw1 ; Kw2 ; Kw3 ; Kw4", pon: "wtorek ; środa ; czwartek ; piątek ; sobota ; niedziela ; poniedziałek",
    sty_full: "luty ; marzec ; kwiecień ; maj ; czerwiec ; lipiec ; sierpień", sty: "lut ; mar ; kwi ; maj ; cze ; lip ; sie",
    ponShort: "wt. ; śr. ; czw. ; pt. ; sob. ; niedz. ; pon.", date: "D2025-3-13 ; D2025-3-14 ; D2025-3-15 ; D2025-3-16 ; D2025-3-17 ; D2025-3-18 ; D2025-3-19",
    poz: "Pozycja 2 ; Pozycja 3 ; Pozycja 4 ; Pozycja 5 ; Pozycja 6 ; Pozycja 7 ; Pozycja 8", a1b: "a1b ; a1b ; a1b ; a1b ; a1b ; a1b ; a1b",
    q1: "Q2 ; Q3 ; Q4 ; Q5 ; Q6 ; Q7 ; Q8", x: "x ; x ; x ; x ; x ; x ; x",
    time: [11, 12, 13, 14, 15, 16, 17].map((h) => Number((h / 24).toPrecision(10))).join(" ; "),
    eom: "D2025-3-31 ; D2025-4-30 ; D2025-5-31 ; D2025-6-30 ; D2025-7-31 ; D2025-8-31", mid: "D2025-3-15 ; D2025-4-15 ; D2025-5-15 ; D2025-6-15 ; D2025-7-15 ; D2025-8-15",
    Pon: "Wt. ; Śr. ; Czw. ; Pt. ; Sob. ; Niedz. ; Pon.", PON: "WTOREK ; ŚRODA ; CZWARTEK ; PIĄTEK ; SOBOTA ; NIEDZIELA ; PONIEDZIAŁEK",
    Sty: "Luty ; Marzec ; Kwiecień ; Maj ; Czerwiec ; Lipiec ; Sierpień", nd: "Nd ; Nd ; Nd ; Nd ; Nd ; Nd ; Nd",
    gru: "styczeń ; luty ; marzec ; kwiecień ; maj ; czerwiec ; lipiec", dec: "1.5 ; 1.5 ; 1.5 ; 1.5 ; 1.5 ; 1.5 ; 1.5", numx: "2 ; x ; 3 ; x ; 4 ; x",
    kw4: "Kw1 ; Kw2 ; Kw3 ; Kw4 ; Kw1 ; Kw2 ; Kw3", pct: "0.1 ; 0.1 ; 0.1 ; 0.1 ; 0.1 ; 0.1 ; 0.1", kwLead: "2 kw ; 3 kw ; 4 kw ; 1 kw ; 2 kw ; 3 kw ; 4 kw",
    etap: "Etap 10 ; Etap 11 ; Etap 12 ; Etap 13 ; Etap 14 ; Etap 15 ; Etap 16", a07: "A-08 ; A-09 ; A-10 ; A-11 ; A-12 ; A-13 ; A-14",
    neg: "-1 ; -1 ; -1 ; -1 ; -1 ; -1 ; -1", trend3: "5.666666667 ; 7.166666667 ; 8.666666667 ; 10.16666667 ; 11.66666667",
    szt: "2 szt ; 3 szt ; 4 szt ; 5 szt ; 6 szt ; 7 szt ; 8 szt", r2025: "2025r ; 2025r ; 2025r ; 2025r ; 2025r ; 2025r ; 2025r",
    fak: "Faktura 002 ; Faktura 003 ; Faktura 004 ; Faktura 005 ; Faktura 006 ; Faktura 007 ; Faktura 008", kwSp: "Kw 2 ; Kw 3 ; Kw 4 ; Kw 1 ; Kw 2 ; Kw 3 ; Kw 4",
    kwartal: "kwartał 2 ; kwartał 3 ; kwartał 4 ; kwartał 1 ; kwartał 2 ; kwartał 3 ; kwartał 4", x13: "x5 ; x7 ; x9 ; x11 ; x13 ; x15",
    x1x: "1x ; 1x ; 1x ; 1x ; 1x ; 1x ; 1x", poz13: "Pozycja 5 ; Pozycja 7 ; Pozycja 9 ; Pozycja 11 ; Pozycja 13 ; Pozycja 15",
    styMar: "maj ; lip ; wrz ; lis ; sty ; mar", ponSr: "pt. ; niedz. ; wt. ; czw. ; sob. ; pon.", ab: "a ; b ; a ; b ; a ; b",
    n12x: "3 ; 4 ; x ; 5 ; 6", xPoz: "x ; Pozycja 2 ; x ; Pozycja 3 ; x ; Pozycja 4",
    dt: "D2025-3-13 8:30 ; D2025-3-14 8:30 ; D2025-3-15 8:30 ; D2025-3-16 8:30 ; D2025-3-17 8:30 ; D2025-3-18 8:30 ; D2025-3-19 8:30",
    years: "D2026-1-1 ; D2027-1-1 ; D2028-1-1 ; D2029-1-1 ; D2030-1-1 ; D2031-1-1", week: "D2025-3-15 ; D2025-3-22 ; D2025-3-29 ; D2025-4-5 ; D2025-4-12 ; D2025-4-19",
    poz9a: "Poz. 9 a ; Poz. 9 a ; Poz. 9 a ; Poz. 9 a ; Poz. 9 a ; Poz. 9 a ; Poz. 9 a", wtCzw: "sobota ; poniedziałek ; środa ; piątek ; niedziela ; wtorek",
    n24xy: "6 ; 8 ; x ; y", kw13: "Kw1 ; Kw3 ; Kw1 ; Kw3 ; Kw1 ; Kw3", half: "1.5 ; 2 ; 2.5 ; 3 ; 3.5 ; 4", pt: "sob. ; niedz. ; pon. ; wt. ; śr. ; czw. ; pt.",
    Wrz: "Paź ; Lis ; Gru ; Sty ; Lut ; Mar ; Kwi", maj: "czerwiec ; lipiec ; sierpień ; wrzesień ; październik ; listopad ; grudzień",
    pozNeg: "Pozycja -2 ; Pozycja -3 ; Pozycja -4 ; Pozycja -5 ; Pozycja -6 ; Pozycja -7 ; Pozycja -8", kwDot: "Kw. 2 ; Kw. 3 ; Kw. 4 ; Kw. 1 ; Kw. 2 ; Kw. 3 ; Kw. 4",
    Ikw: "I kwartał ; I kwartał ; I kwartał ; I kwartał ; I kwartał ; I kwartał ; I kwartał", tydz: "Tydzień 53 ; Tydzień 54 ; Tydzień 55 ; Tydzień 56 ; Tydzień 57 ; Tydzień 58 ; Tydzień 59",
    ten: "10 ; 10 ; 10 ; 10 ; 10 ; 10 ; 10", tenTxt: "11 ; 12 ; 13 ; 14 ; 15 ; 16 ; 17", a15: "a 1,6 ; a 1,7 ; a 1,8 ; a 1,9 ; a 1,10 ; a 1,11 ; a 1,12",
    x31: "x1 ; x3 ; x5 ; x7 ; x9 ; x11",
  };
  const badSeries = Object.keys(want).filter((k) => series[k] !== want[k]).map((k) => `${k}: „${series[k]}” (Excel: „${want[k]}”)`);
  ok("4. serie wypełniania = wyniki Excela", badSeries.length === 0, badSeries);

  // 5. uchwyt i skróty — arkusz Fill
  await page.evaluate(() => { const s = document.getElementById("sheetSelect"); s.value = "Fill"; s.dispatchEvent(new Event("change", { bubbles: true })); loadBtn.click(); });
  await page.waitForFunction(() => currentSheetName === "Fill" && !!document.querySelector('#dataTable tbody tr[data-row-key="wide:1"]'), null, { timeout: 15000 });
  await sleep(300);
  const td = (r, c) => `#dataTable tbody tr[data-row-key="wide:${r}"] td[data-col-index="${c}"]`;
  const text = (r, c) => page.evaluate((sel) => (document.querySelector(sel)?.textContent ?? "").replace(/ /g, " "), td(r, c));
  const col = async (c, from, to) => { const out = []; for (let r = from; r <= to; r++) out.push(await text(r, c)); return out; };
  const center = async (sel) => { const b = await page.locator(sel).boundingBox(); return { x: b.x + b.width / 2, y: b.y + b.height / 2 }; };
  const dragHandle = async (toR, toC, { alt = false } = {}) => {
    await page.waitForSelector("#tableWrap .fill-handle", { timeout: 3000 });
    const h = await center("#tableWrap .fill-handle");
    const t = await center(td(toR, toC));
    await page.mouse.move(h.x, h.y);
    await page.mouse.down();
    await page.mouse.move((h.x + t.x) / 2, (h.y + t.y) / 2, { steps: 4 });
    await page.mouse.move(t.x, t.y, { steps: 4 });
    if (alt) await page.keyboard.down("Alt");
    await page.mouse.up();
    if (alt) await page.keyboard.up("Alt");
    await sleep(250);
  };
  const pick = async (r, c) => { await page.click(td(r, c)); await sleep(120); };

  await pick(1, 1);
  // uchwyt leży w #tableWrap (nie w komórce — td go przycinało); róg wskazuje td.has-fill-handle
  const handleHost = await page.evaluate(() => { const h = document.querySelector("#tableWrap .fill-handle"); const td = h && document.querySelector("#dataTable td.has-fill-handle"); return td ? td.dataset.colIndex + "@" + td.closest("tr")?.dataset.rowKey : null; });
  ok("5. uchwyt stoi w prawym dolnym rogu aktywnej komórki", handleHost === "1@wide:1", handleHost);
  await dragHandle(5, 1);
  ok("5. data przeciągnięta uchwytem → kolejne dni, format daty", JSON.stringify(await col(1, 2, 5)) === JSON.stringify(["13.03.2025", "14.03.2025", "15.03.2025", "16.03.2025"]), await col(1, 2, 5));
  const sel = await page.evaluate(() => { const r = getSelectionRectangle(); return r ? [r.rowStart, r.rowEnd, r.colMin, r.colMax] : null; });
  ok("5. po wypełnieniu zaznaczony cały zakres (jak Excel)", JSON.stringify(sel) === JSON.stringify([0, 4, 1, 1]), sel);

  await pick(1, 2);
  await dragHandle(4, 2);
  ok("5. „Pozycja 1” → Pozycja 2, 3, 4", JSON.stringify(await col(2, 2, 4)) === JSON.stringify(["Pozycja 2", "Pozycja 3", "Pozycja 4"]), await col(2, 2, 4));

  await pick(1, 0);
  await dragHandle(3, 0);
  ok("5. jedna liczba = kopia", JSON.stringify(await col(0, 2, 3)) === JSON.stringify(["1", "1"]), await col(0, 2, 3));
  await pick(1, 0);
  await dragHandle(5, 0, { alt: true });
  ok("5. jedna liczba z Alt = seria 2, 3, 4, 5", JSON.stringify(await col(0, 2, 5)) === JSON.stringify(["2", "3", "4", "5"]), await col(0, 2, 5));

  await pick(1, 3);
  await dragHandle(3, 3);
  const kw = await page.evaluate(() => ({ z: cfmtNumFmtAt(2, 3), v: workbook.Sheets.Fill.D4?.v }));
  ok("5. kwota w zł: kopia z formatem „zł”", JSON.stringify(await col(3, 2, 3)) === JSON.stringify(["12,50 zł", "12,50 zł"]) && kw.v === 12.5 && /zł/.test(kw.z), { shown: await col(3, 2, 3), kw });

  await pick(1, 4);
  await page.click(td(2, 4), { modifiers: ["Shift"] });
  await sleep(120);
  await dragHandle(5, 4);
  ok("5. dwie liczby 2, 4 → trend 6, 8, 10", JSON.stringify(await col(4, 3, 5)) === JSON.stringify(["6", "8", "10"]), await col(4, 3, 5));

  await pick(1, 10);
  await page.waitForSelector("#tableWrap .fill-handle", { timeout: 3000 });
  await page.dblclick("#tableWrap .fill-handle");
  await sleep(300);
  ok("5. dwuklik uchwytu: „Kw1” w dół do końca kolumny Nr (kwartały wracają do 1)",
    JSON.stringify(await col(10, 2, 10)) === JSON.stringify(["Kw2", "Kw3", "Kw4", "Kw1", "Kw2", "Kw3", "Kw4", "Kw1", "Kw2"]), await col(10, 2, 10));

  // Ctrl+D na pojedynczej komórce = kopia z komórki NAD nią (tekst zostaje tekstem)
  await pick(2, 5);
  await page.keyboard.press("Control+d");
  await sleep(200);
  const f3 = await page.evaluate(() => workbook.Sheets.Fill.F3 && { t: workbook.Sheets.Fill.F3.t, v: workbook.Sheets.Fill.F3.v });
  ok("5. Ctrl+D: „007” skopiowane jako tekst „007” (nie liczba 7)", f3 && f3.t === "s" && f3.v === "007" && (await text(2, 5)) === "007", f3);
  await pick(2, 6);
  await page.keyboard.press("Control+d");
  await sleep(200);
  const g3 = await page.evaluate(() => workbook.Sheets.Fill.G3?.f || null);
  ok("5. Ctrl+D: formuła z przesuniętym odwołaniem (=A3*2)", g3 === "A3*2", g3);
  await pick(1, 8);
  await page.keyboard.press("Control+r");
  await sleep(200);
  ok("5. Ctrl+R: kopia z komórki po lewej", (await text(1, 8)) === "x", await text(1, 8));

  // 2. ułamek wpisany w komórkę: widok jak w Excelu („1 1/2”, „ 3/4”)
  const typeIn = async (r, c, s) => {
    await page.dblclick(td(r, c));
    await page.waitForSelector(`${td(r, c)} input.cell-editor`, { timeout: 3000 });
    await page.keyboard.press("ControlOrMeta+a");
    await page.keyboard.type(s);
    await page.keyboard.press("Enter");
    await sleep(200);
  };
  await typeIn(3, 8, "1 1/2");
  await typeIn(4, 8, "0 3/4");
  const frCells = await page.evaluate(() => ["I4", "I5"].map((ref) => { const c = workbook.Sheets.Fill[ref]; return c ? `${c.v}|${c.z}` : null; }));
  ok("2. „1 1/2” w komórce → 1,5 [# ?/?], widać „1 1/2”; „0 3/4” → widać „ 3/4”",
    frCells[0] === "1.5|# ?/?" && frCells[1] === "0.75|# ?/?" && (await text(3, 8)) === "1 1/2" && (await text(4, 8)) === " 3/4", { frCells, i4: await text(3, 8), i5: await text(4, 8) });
  await page.evaluate(() => { undoLast(); undoLast(); });
  await sleep(200);

  // 6. Cofnij — całe wypełnienie dwuklikiem to jeden krok? (ostatnie: Ctrl+R; cofamy trzy razy)
  await page.keyboard.press("Escape");
  await page.evaluate(() => { undoLast(); undoLast(); undoLast(); });
  await sleep(250);
  const undone = await page.evaluate(() => ({ i2: workbook.Sheets.Fill.I2?.v ?? null, g3: workbook.Sheets.Fill.G3?.f ?? null, f3: workbook.Sheets.Fill.F3?.v ?? null, k3: workbook.Sheets.Fill.K3?.v ?? null }));
  ok("6. Cofnij ×3 zdejmuje Ctrl+R, oba Ctrl+D, a seria z dwukliku zostaje", undone.i2 === null && undone.g3 === null && undone.f3 === null && undone.k3 === "Kw2", undone);
  await page.evaluate(() => undoLast());
  await sleep(250);
  const undoneSeries = await page.evaluate(() => [3, 4, 5, 6, 7, 8, 9, 10, 11].map((r) => workbook.Sheets.Fill[`K${r}`]?.v ?? null));
  ok("6. jedno Cofnij zdejmuje całą serię z dwukliku", undoneSeries.every((v) => v == null), undoneSeries);

  fs.rmSync(tmp, { recursive: true, force: true });
  await browser.close();
  if (errors.length) failures.push("błędy strony: " + errors.join(" | "));
  if (failures.length) {
    console.error(`❌ excel-round2 (${ENGINE}):\n  - ` + failures.join("\n  - "));
    process.exit(1);
  }
  console.log(`✅ excel-round2 (${ENGINE}): sortowanie typów jak Excel, ułamki, kolory [Red], serie (≈60 przypadków z Excela), uchwyt, Ctrl+D/R, Cofnij`);
}

run().catch((e) => { console.error(e); process.exit(1); });
