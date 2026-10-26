// paste-format-playwright.js — Wklej z formatowaniem / Wklej same wartości (jak w Excelu).
//
// Film Mateusza (2026-10-25): skopiowana data „15 paź 26” wklejała się jako „15-10-26” —
// kopia w aplikacji niosła TEKST z ekranu, wklejenie przepuszczało go przez parser, a format
// zostawał z komórki docelowej. Decyzja 2026-10-26: do wyboru „z formatowaniem” i „bez”.
//
// Co musi być prawdą (kopia z TEJ aplikacji):
//   1. „Wklej” (Ctrl/⌘+V) = wartość + format liczby źródła (kwota „zł”, data jak w źródle),
//   2. wklejana jest PRAWDZIWA wartość, nie zaokrąglony tekst z ekranu (2,345 → nie 2,35),
//   3. „Wklej” niesie też wygląd: pogrubienie, wypełnienie,
//   4. „Wklej same wartości” (menu, Ctrl/⌘+Shift+V) = prawdziwa wartość, wygląd i format
//      komórki docelowej zostają,
//   5. Cofnij przywraca wartość I format sprzed wklejenia,
//   6. format trafia do zapisu (pendingEdits[...].fmt), menu ma obie pozycje.

const playwright = require("playwright");
const path = require("path");

const APP_URL = process.env.APP_URL || "http://127.0.0.1:4175/";
const ENGINE = process.env.ENGINE || "chromium";
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
const FILE = path.join(__dirname, "fixtures", "openpyxl-features.xlsx");

async function run() {
  const browser = await playwright[ENGINE].launch({ headless: true });
  // strefa z przesunięciem względem UTC — tu było widać „01:00” przy datach z pliku
  const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1280, height: 860 }, timezoneId: "Europe/Warsaw" });
  await context.addInitScript(() => localStorage.setItem("introPlayed", "true"));
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push("pageerror: " + e.message));
  page.on("dialog", (d) => d.accept());
  const results = [];
  const check = (name, ok, detail = "") => results.push({ name, ok: !!ok, detail: typeof detail === "string" ? detail : JSON.stringify(detail) });

  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(() => { document.getElementById("heroSplash")?.remove(); return ensureXlsxLibs(false); });
  await page.setInputFiles("#fileInput", FILE);
  await page.waitForFunction(() => !!document.querySelector("#dataTable tbody tr[data-row-key]"), null, { timeout: 20000 });
  await sleep(500);

  // pomocnicze w stronie: komórka po nagłówku i numerze wiersza w pliku (1-based)
  await page.evaluate(() => {
    window.__pf = {
      col: (h) => currentDisplayModel.headers.indexOf(h),
      row: (r1) => currentDisplayModel.rows.find((x) => x.rowIndex0 === r1 - 1),
      focus(r1, h) { setSelectionKind("cell", { repaint: false }); setFocusedCell(getRowSelectionKey(this.row(r1)), this.col(h), { scroll: false }); },
      info(r1, h) {
        const row = this.row(r1), c = this.col(h);
        const ref = XLSX.utils.encode_cell({ r: r1 - 1, c: currentStartCol + c });
        const cell = workbook.Sheets[currentSheetName][ref];
        const st = (row.cellStyles && row.cellStyles[c]) || {};
        const v = row.values[c];
        return { v: v instanceof Date ? v.toISOString().slice(0, 10) : v, shown: getDisplayValue(row, c), numFmt: cfmtNumFmtAt(r1 - 1, c), bold: !!st.bold, fill: st.hasCustomFill ? st.fillColor : null, z: cell && cell.z, pendFmt: !!pendingEdits[currentSheetName]?.[ref]?.fmt };
      },
    };
  });
  const info = (r, h) => page.evaluate(([r, h]) => window.__pf.info(r, h), [r, h]);
  const focus = (r, h) => page.evaluate(([r, h]) => window.__pf.focus(r, h), [r, h]);
  const copy = async (r, h) => { await focus(r, h); await page.evaluate(() => copySelectionToClipboard()); await sleep(120); };
  const paste = async (r, h, mode) => { await focus(r, h); await page.evaluate((m) => pasteClipboardToSelection({ source: "app", mode: m }), mode); await sleep(200); };

  const headers = await page.evaluate(() => currentDisplayModel.headers.slice(0, 12));
  const H = { data: headers[1], kwota: headers[4], udzial: headers[5], opis: headers[6] };

  // 1. kwota „zł” → kolumna Opis (tekst, format Ogólny): Wklej = wartość + format
  const src1 = await info(3, H.kwota);
  const dstBefore = await info(4, H.opis);
  await copy(3, H.kwota);
  await paste(4, H.opis);
  const dst1 = await info(4, H.opis);
  check("1. Wklej: kwota przenosi format liczby („zł”) i wartość", dst1.v === src1.v && dst1.numFmt === src1.numFmt && dst1.shown === src1.shown && dst1.pendFmt, { src1, dst1 });

  // 1b. data → Opis: wygląda jak w źródle (nie „15-10-26”)
  const srcD = await info(3, H.data);
  await copy(3, H.data);
  await paste(5, H.opis);
  const dstD = await info(5, H.opis);
  check("1. Wklej: data wygląda jak w źródle, wartość = data", dstD.v === srcD.v && dstD.shown === srcD.shown && dstD.numFmt === srcD.numFmt, { srcD, dstD });

  // 2. prawdziwa wartość: źródło 2,345 w formacie 0.00 („2,35”)
  await page.evaluate((h) => {
    const r = window.__pf.row(6), c = window.__pf.col(h);
    updateSheetCell(r.rowIndex0, c, { type: "number", value: 2.345 });
    r.values[c] = 2.345;
    window.__pf.focus(6, h);
    cfmtApply({ numFmt: "0.00" });
  }, H.kwota);
  await sleep(150);
  const src2 = await info(6, H.kwota);
  await copy(6, H.kwota);
  await paste(7, H.opis, "values");
  const dst2 = await info(7, H.opis);
  check("2. wklejona prawdziwa wartość 2,345 (nie tekst „2,35” z ekranu)", /2[,.]35/.test(String(src2.shown)) && dst2.v === 2.345, { src2, dst2 });

  // 3. wygląd: pogrubienie + wypełnienie
  await focus(8, H.kwota);
  await page.evaluate(() => cfmtApply({ b: true, fill: "FFFF00" }));
  await sleep(150);
  await copy(8, H.kwota);
  await paste(9, H.opis);
  const dst3 = await info(9, H.opis);
  check("3. Wklej: pogrubienie i wypełnienie źródła", dst3.bold && /FFFF00/i.test(String(dst3.fill)), dst3);

  // 4. same wartości: wygląd celu zostaje
  await focus(10, H.udzial);
  const tgt4 = await info(10, H.udzial);
  await paste(10, H.udzial, "values");
  const dst4 = await info(10, H.udzial);
  const src8 = await info(8, H.kwota);
  check("4. Wklej same wartości: wartość źródła, format i wygląd celu zostają", dst4.v === src8.v && dst4.numFmt === tgt4.numFmt && dst4.bold === tgt4.bold && dst4.fill === tgt4.fill, { tgt4, dst4, src8 });

  // 4b. Ctrl/⌘+Shift+V = same wartości
  await focus(11, H.udzial);
  const tgt4b = await info(11, H.udzial);
  await page.click(`#dataTable tbody tr[data-row-key="wide:10"] td[data-col-index="${headers.indexOf(H.udzial)}"]`);
  await page.keyboard.press(process.platform === "darwin" && ENGINE === "webkit" ? "Meta+Shift+V" : "Control+Shift+V");
  await sleep(300);
  const dst4b = await info(11, H.udzial);
  check("4. Ctrl/⌘+Shift+V wkleja same wartości", dst4b.v === src8.v && dst4b.bold === tgt4b.bold && dst4b.fill === tgt4b.fill && dst4b.numFmt === tgt4b.numFmt, { tgt4b, dst4b });

  // 5. Cofnij po „Wklej” z formatem
  await focus(12, H.opis);
  const before5 = await info(12, H.opis);
  await paste(12, H.opis);
  const mid5 = await info(12, H.opis);
  await page.evaluate(() => undoRedo("undo"));
  await sleep(250);
  const after5 = await info(12, H.opis);
  check("5. Cofnij przywraca wartość i wygląd sprzed wklejenia", mid5.bold && after5.v === before5.v && after5.numFmt === before5.numFmt && after5.bold === before5.bold && after5.fill === before5.fill && after5.pendFmt === before5.pendFmt, { before5, mid5, after5 });

  // 7. daty z pliku = czas „ścienny” (SheetJS czyta je jako północ UTC = 01:00 w Polsce):
  //    pasek formuły bez godziny, formuła porównująca datę = PRAWDA, wklejona data w pliku
  //    jako pełna liczba dni (dawniej 46030,0417 — ukryta godzina 01:00)
  await focus(3, H.data);
  await sleep(150);
  const fx7 = await page.evaluate(() => document.getElementById("formulaInput")?.value || "");
  const eq7 = await page.evaluate(() => {
    const FE = FormulaEngine;
    const v = workbook.Sheets[currentSheetName]["B3"].v;
    const ast = FE.parse(`B3=DATE(${v.getFullYear()},${v.getMonth() + 1},${v.getDate()})`, { lang: "en" });
    return FE.evaluate(ast, fcContext(currentSheetName, 0, 0));
  });
  await copy(3, H.data);
  await paste(13, H.opis);
  const saved7 = await page.evaluate(async () => {
    const bytes = await buildOutputBytes("xlsx");
    const wb = XLSX.read(bytes, {});
    const ref = XLSX.utils.encode_cell({ r: 12, c: currentStartCol + currentDisplayModel.headers.indexOf(currentDisplayModel.headers[6]) });
    return { ref, v: wb.Sheets[currentSheetName][ref]?.v, src: wb.Sheets[currentSheetName]["B3"]?.v };
  });
  check("7. data z pliku na pasku formuły bez „01:00”", /^\d{4}-\d{2}-\d{2}$/.test(fx7.trim()), fx7);
  check("7. =B3=DATA(…) dla daty z pliku daje PRAWDA", eq7 === true, eq7);
  check("7. wklejona data zapisana jako pełna liczba dni, równa źródłu", Number.isInteger(saved7.v) && saved7.v === saved7.src, saved7);

  // 6. menu komórki: obie pozycje
  await page.evaluate(() => document.querySelectorAll(".toast").forEach((x) => x.remove()));
  const td = `#dataTable tbody tr[data-row-key="wide:13"] td[data-col-index="${headers.indexOf(H.opis)}"]`;
  await page.locator(td).scrollIntoViewIfNeeded();
  const box = await page.locator(td).boundingBox();
  await page.mouse.click(box.x + box.width / 2, box.y + box.height / 2, { button: "right" });
  await page.waitForSelector("#cellMenu:not(.hidden)", { timeout: 3000 });
  const menu = await page.evaluate(() => [...document.querySelectorAll("#cellMenu .cell-menu-item")].map((b) => b.textContent));
  check("6. menu komórki: „Wklej …” i „Wklej same wartości”", menu.some((x) => /^Wklej „/.test(x)) && menu.includes("Wklej same wartości"), menu);
  await page.keyboard.press("Escape");

  check("brak błędów strony", errors.length === 0, errors.slice(0, 3).join(" | "));
  await browser.close();

  let failed = 0;
  for (const r of results) {
    console.log(`${r.ok ? "✅" : "❌"} ${r.name}${!r.ok && r.detail ? `  (${r.detail})` : ""}`);
    if (!r.ok) failed += 1;
  }
  if (failed) { console.error(`\n${failed} z ${results.length} nie przeszło`); process.exit(1); }
  console.log(`\n✅ paste-format (${ENGINE}): ${results.length}/${results.length}`);
}

run().catch((e) => { console.error(e); process.exit(1); });
