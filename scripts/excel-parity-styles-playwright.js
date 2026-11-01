// excel-parity-styles-playwright.js — kolory, zawijanie i daty jak w Excelu (scripts/fixtures/excel-parity-styles.xlsx).
//
// Zgłoszenie 2026-11-01 (plik „zadania i przywileje…”, runda 2 porównania z Excelem):
//   1. Wypełnienie gradientowe nie przesuwa kolorów dalszych komórek. SheetJS przy <gradientFill>
//      dokładał dwa puste wpisy do listy wypełnień — żółte FFFF99 wychodziło szare, C00000 znikało.
//   2. Kolor czcionki z palety (color indexed="2") jest czerwony. SheetJS psuł rgb na „ff00”.
//   3. „Zawijaj tekst” w wierszu nad nagłówkiem zawija (Wymiary z Excela).
//   4. Format „[$-415]d/mmm/yyyy;@” wygląda jak w polskim Excelu: „8.wrz.2026” (wyrocznia AppleScript).
//
// Fixture generuje openpyxl (gradient jest pierwszym wypełnieniem w pliku).
// Uruchom z serwerem na APP_URL (domyślnie http://127.0.0.1:4175/), np. `npm run serve:test`.

const { chromium, webkit } = require("playwright");
const path = require("path");

const APP_URL = process.env.APP_URL || "http://127.0.0.1:4175/";
const ENGINE = process.env.ENGINE === "webkit" ? webkit : chromium;
const FILE = path.join(__dirname, "fixtures", "excel-parity-styles.xlsx");

async function run() {
  const browser = await ENGINE.launch({ headless: true });
  const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1280, height: 800 } });
  await context.addInitScript(() => { localStorage.setItem("introPlayed", "true"); });
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push("pageerror: " + e.message));
  page.on("console", (m) => { if (m.type() === "error") errors.push("console: " + m.text()); });
  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(() => { document.getElementById("heroSplash")?.remove(); try { ensureXlsxLibs(false); } catch {} });
  await page.setInputFiles("#fileInput", FILE);
  await page.waitForFunction(() => currentSheetName === "Style" && baseRows.length > 0, null, { timeout: 20000 });
  await page.waitForTimeout(600);

  const failures = [];
  const ok = (name, cond, got) => { if (!cond) failures.push(`${name} — dostałem: ${JSON.stringify(got)}`); };

  const got = await page.evaluate(() => {
    const find = (t) => [...document.querySelectorAll("#dataTable td")].find((td) => td.textContent.trim() === t);
    const cs = (td) => (td ? getComputedStyle(td) : null);
    const pre = [...document.querySelectorAll("#dataTable thead tr.pre-row td.pre-cell")].find((td) => td.textContent.startsWith("Aktualizacja"));
    return {
      header: currentHeaderRow,
      yellow: cs(find("żółte"))?.backgroundColor,
      red: cs(find("czerwone"))?.backgroundColor,
      preColor: cs(pre)?.color,
      preWrap: cs(pre)?.whiteSpace,
      preLines: pre ? Math.round(pre.scrollHeight / parseFloat(cs(pre).lineHeight || "14")) : 0,
      dates: [...document.querySelectorAll('#dataTable tbody td[data-col-index="0"]')].map((td) => td.textContent.trim()).slice(0, 3),
    };
  });
  ok("nagłówek w wierszu 3 (wiersz 1 nad nim)", got.header === 3, got.header);
  ok("żółte FFFF99 po gradiencie", /255, 255, 153/.test(got.yellow || ""), got.yellow);
  ok("czerwone C00000 po gradiencie", /192, 0, 0/.test(got.red || ""), got.red);
  ok("kolor czcionki indexed=2 = czerwony", /255, 0, 0/.test(got.preColor || ""), got.preColor);
  ok("„Zawijaj tekst” nad nagłówkiem zawija", got.preWrap === "normal" && got.preLines >= 2, [got.preWrap, got.preLines]);
  ok("data jak polski Excel (8.wrz.2026)", got.dates[0] === "8.wrz.2026" && got.dates[1] === "15.wrz.2026" && got.dates[2] === "22.wrz.2026", got.dates);

  const real = errors.filter((e) => !/favicon|Failed to load resource/.test(e));
  ok("brak błędów w konsoli", real.length === 0, real.slice(0, 5));
  await browser.close();
  if (failures.length) {
    console.log(`❌ excel-parity-styles (${process.env.ENGINE || "chromium"}): ${failures.length} błędów`);
    failures.forEach((f) => console.log("  ✗", f));
    process.exit(1);
  }
  console.log(`✅ excel-parity-styles (${process.env.ENGINE || "chromium"}): wszystko OK`);
}

run().catch((e) => { console.error(e); process.exit(1); });
