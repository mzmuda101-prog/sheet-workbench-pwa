// excel-parity-playwright.js — wygląd arkusza jak w Excelu (scripts/fixtures/excel-parity.xlsx).
//
// Zgłoszenie 2026-11-01 (plik „zadania i przywileje…”): porównanie zrzutów Excel ↔ apka.
// Co musi być prawdą (przy włączonych „Wymiarach z Excela”, domyślnie):
//   1. Formatowanie warunkowe z LICZ.JEŻELI działa — także na PUSTYCH komórkach (Excel: pusta
//      komórka w kryterium = liczba 0, więc „=LICZ.JEŻELI($G$2:$G$3;C8)=0” jest prawdą).
//   2. Wiersz bez własnej wysokości ma domyślną wysokość arkusza (<sheetFormatPr defaultRowHeight>),
//      a nie wysokość „apkową” (~31 px).
//   3. Nagłówek z „Zawijaj tekst” zawija się (bez dzielenia słów), pusta komórka nagłówka jest pusta.
//   4. Wąska kolumna (Excel 5,66 znaku) ma mały odstęp — „WE:” mieści się bez „…”.
//   5. Arkusz otwiera się przewinięty do <sheetView/pane topLeftCell>, gdy tej komórki nie widać od razu
//      (w pionie tylko przy zablokowanym nagłówku — nagłówki kolumn nie mogą zniknąć); raz na arkusz.
//
// Uruchom z serwerem na APP_URL (domyślnie http://127.0.0.1:4175/), np. `npm run serve:test`.

const { chromium, webkit } = require("playwright");
const path = require("path");

const APP_URL = process.env.APP_URL || "http://127.0.0.1:4175/";
const ENGINE = process.env.ENGINE === "webkit" ? webkit : chromium;
const FILE = path.join(__dirname, "fixtures", "excel-parity.xlsx");

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
  await page.waitForFunction(() => currentSheetName === "Lista" && baseRows.length > 0, null, { timeout: 20000 });
  await page.waitForTimeout(600);

  const failures = [];
  const ok = (name, cond, got) => { if (!cond) failures.push(`${name} — dostałem: ${JSON.stringify(got)}`); };

  const lista = await page.evaluate(() => {
    const cell = (r, c) => document.querySelector(`tbody tr[data-row-index="${r}"] td[data-col-index="${c}"]`);
    const bg = (td) => (td ? getComputedStyle(td).backgroundColor : null);
    const ths = [...document.querySelectorAll("thead tr.header-row th[data-col-index]")];
    const th = (i) => ths.find((x) => x.dataset.colIndex === String(i));
    const we = cell(2, 4);
    return {
      fillMissing: bg(cell(4, 2)), // C5 „Osoba 005” — nie ma jej na liście G2:G3
      fillListed: bg(cell(1, 2)), // C2 „Osoba 002” — jest na liście
      fillBlank: bg(cell(8, 2)), // C9 pusta
      hOwn: cell(2, 0) && cell(2, 0).parentElement.getBoundingClientRect().height, // wiersz 3: ht=15
      hDefault: cell(7, 0) && cell(7, 0).parentElement.getBoundingClientRect().height, // wiersz 8: bez ht
      wrapClass: th(3) && th(3).classList.contains("th-wrap"),
      wrapLines: th(3) && Math.round(th(3).scrollHeight / parseFloat(getComputedStyle(th(3)).lineHeight || "14")),
      emptyHead: th(1) && th(1).textContent,
      emptyHeadName: currentHeaders[1],
      weFits: we && we.scrollWidth <= we.clientWidth + 1,
      wePad: we && getComputedStyle(we).paddingLeft,
    };
  });
  const pink = (c) => /255, 199, 206/.test(c || "");
  ok("CF: wartość spoza listy podświetlona", pink(lista.fillMissing), lista.fillMissing);
  ok("CF: wartość z listy bez podświetlenia", !pink(lista.fillListed), lista.fillListed);
  ok("CF: pusta komórka podświetlona (pusta = 0 jak w Excelu)", pink(lista.fillBlank), lista.fillBlank);
  ok("wiersz bez wysokości = domyślna arkusza (14 pt ≈ 19 px)", lista.hDefault > 0 && lista.hDefault <= 22, lista.hDefault);
  ok("wiersz z wysokością 15 pt ≈ 20 px", lista.hOwn > 0 && lista.hOwn <= 23, lista.hOwn);
  ok("nagłówek z „Zawijaj tekst” zawija się", lista.wrapClass && lista.wrapLines >= 2, [lista.wrapClass, lista.wrapLines]);
  ok("pusta komórka nagłówka pusta (nazwa kolumny zostaje literą)", lista.emptyHead === "" && lista.emptyHeadName === "B", [lista.emptyHead, lista.emptyHeadName]);
  ok("„WE:” mieści się w wąskiej kolumnie", lista.weFits, lista);

  // arkusz zapisany jako przewinięty do H40
  await page.evaluate(() => { sheetSelect.value = "Przewinięty"; loadBtn.click(); });
  await page.waitForFunction(() => currentSheetName === "Przewinięty" && baseRows.length > 0);
  await page.waitForTimeout(800);
  const scrolled = await page.evaluate(() => {
    const wrap = tableWrapEl.getBoundingClientRect();
    const td = document.querySelector('tbody tr[data-row-index="39"] td[data-col-index="7"]');
    const r = td && td.getBoundingClientRect();
    // arkusz ma zablokowany wiersz nagłówka — wiersz 40 ma stać tuż pod nim
    const ref = document.querySelector("thead tr.header-row th").getBoundingClientRect().bottom;
    void ref;
    // wiersz 40 przy górnej krawędzi tabeli (pod nagłówkiem, który zostaje na górze)
    return { scroll: [tableWrapEl.scrollLeft, tableWrapEl.scrollTop], left: r && Math.round(r.left - wrap.left), top: r && Math.round(r.top - wrap.top), under: Math.round(ref - wrap.top), headVisible: document.querySelector("thead tr.header-row th").getBoundingClientRect().bottom > wrap.top };
  });
  ok("widok startowy: przewinięty do H40", scrolled.scroll[0] > 100 && scrolled.scroll[1] > 100 && scrolled.left !== null && scrolled.left < 120 && scrolled.headVisible && Math.abs(scrolled.top - scrolled.under) < 6, scrolled);
  // powrót do arkusza nie przewija drugi raz (zostaje to, co ustawił użytkownik)
  await page.evaluate(() => { tableWrapEl.scrollLeft = 0; tableWrapEl.scrollTop = 0; sheetSelect.value = "Lista"; loadBtn.click(); });
  await page.waitForFunction(() => currentSheetName === "Lista");
  await page.evaluate(() => { sheetSelect.value = "Przewinięty"; loadBtn.click(); });
  await page.waitForFunction(() => currentSheetName === "Przewinięty" && baseRows.length > 0);
  await page.waitForTimeout(600);
  const again = await page.evaluate(() => [tableWrapEl.scrollLeft, tableWrapEl.scrollTop]);
  ok("widok startowy tylko raz na arkusz", again[0] < 50 && again[1] < 50, again);

  const real = errors.filter((e) => !/favicon|Failed to load resource/.test(e));
  ok("brak błędów w konsoli", real.length === 0, real.slice(0, 5));
  await browser.close();
  if (failures.length) {
    console.log(`❌ excel-parity (${process.env.ENGINE || "chromium"}): ${failures.length} błędów`);
    failures.forEach((f) => console.log("  ✗", f));
    process.exit(1);
  }
  console.log(`✅ excel-parity (${process.env.ENGINE || "chromium"}): wszystko OK`);
}

run().catch((e) => { console.error(e); process.exit(1); });
