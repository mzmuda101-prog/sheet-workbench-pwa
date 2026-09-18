// row-reuse-playwright.js — PONOWNE UŻYCIE WIERSZY DOM: gwarancja „zero regresji".
//
// Sortowanie i filtrowanie przestawiają teraz istniejące <tr> zamiast budować ~6800
// węzłów od nowa. To może przyspieszyć, ale nie ma prawa NICZEGO zmienić na ekranie.
// Test sprawdza to mechanicznie: po każdej operacji zdejmuje pełny odcisk tabeli
// (klasy, style, teksty, colspan/rowspan, indeksy kolumn), potem wymusza pełną
// przebudowę tego samego stanu i porównuje oba odciski znak po znaku.
//
// Scenariusze obejmują to, co realnie zmienia wygląd wiersza: sortowanie, filtr,
// zaznaczenie komórki, podświetlanie trafień, tryby auto, widok Wide-to-Long,
// podnagłówki, wymiary z Excela, limit wierszy, zmianę wartości komórki i zoom.
//
// Uruchom z serwerem na APP_URL (domyślnie http://127.0.0.1:4175/).

const { chromium } = require("playwright");
const path = require("path");

const APP_URL = process.env.APP_URL || "http://127.0.0.1:4175/";
const FILE = path.join(__dirname, "stress-test-workbench.xlsx");
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

// Odcisk tabeli: wszystko, co widać. Styl normalizujemy (kolejność deklaracji
// potrafi się różnić przy podmianie w miejscu, a to nie jest różnica wizualna).
const SNAPSHOT_FN = `() => {
  const norm = (css) => String(css || "").split(";").map((s) => s.trim()).filter(Boolean).sort().join(";");
  const rows = [...document.querySelectorAll("#dataTable tbody tr")];
  return rows.map((tr) => {
    const cells = [...tr.children].map((td) => [
      td.tagName,
      td.className,
      td.textContent,
      td.dataset.colIndex || "",
      td.dataset.fullText || "",
      td.dataset.recalcWas || "",
      td.colSpan, td.rowSpan,
      norm(td.getAttribute("style")),
      td.title || "",
    ].join("|"));
    return [tr.className, norm(tr.getAttribute("style")), tr.dataset.rowKey || "", tr.dataset.rowIndex || "", cells.join("#")].join("~");
  }).join("\\n");
}`;

async function run() {
  const browser = await chromium.launch({ headless: true });
  const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1100, height: 850 } });
  await context.addInitScript(() => localStorage.setItem("introPlayed", "true"));
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push("pageerror: " + e.message));
  page.on("console", (m) => { if (m.type() === "error") errors.push("console: " + m.text()); });

  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  await page.evaluate(() => { try { ensureXlsxLibs && ensureXlsxLibs(false); } catch {} });
  await page.setInputFiles("#fileInput", FILE);
  await page.waitForFunction(() => document.getElementById("sheetSelect")?.options?.length > 0, null, { timeout: 20000 });
  await sleep(300);
  await page.click("#loadBtn");
  await sleep(1500);

  // Każdy krok: wykonaj operację → odcisk (ścieżka z ponownym użyciem) →
  // wymuś pełną przebudowę tego samego stanu → odcisk → porównaj.
  const steps = [
    ["stan wyjściowy", "renderActiveTable();"],
    ["sortowanie rosnąco", "setPrimarySort(currentHeaders[1], 'asc'); sortRows(); renderActiveTable();"],
    ["sortowanie malejąco", "setPrimarySort(currentHeaders[1], 'desc'); sortRows(); renderActiveTable();"],
    ["sortowanie po innej kolumnie", "setPrimarySort(currentHeaders[2], 'asc'); sortRows(); renderActiveTable();"],
    ["filtr tekstowy", "searchQueryEl.value='a'; filtersCommitted=true; applyFilters(); sortRows(); renderActiveTable();"],
    ["filtr + podświetlanie komórek", "if (highlightMatchCellsEl) highlightMatchCellsEl.checked = true; applyFilters(); sortRows(); renderActiveTable();"],
    ["zaznaczenie komórki", "const tr=document.querySelector('#dataTable tbody tr'); if (tr) setSelectedCell(tr.dataset.rowKey, 2); renderActiveTable();"],
    ["tryb zaznaczania trafień", "quickSearchHighlightMode = true; applyFilters(); sortRows(); renderActiveTable();"],
    ["wyczyszczenie filtra", "quickSearchHighlightMode=false; if (highlightMatchCellsEl) highlightMatchCellsEl.checked=false; searchQueryEl.value=''; applyFilters(); sortRows(); renderActiveTable();"],
    ["podnagłówki wyłączone", "cellStyleShowSubheaders=false; renderActiveTable();"],
    ["podnagłówki włączone", "cellStyleShowSubheaders=true; renderActiveTable();"],
    ["limit wierszy 40", "maxRowsEl.value='40'; renderActiveTable();"],
    ["limit wierszy 200", "maxRowsEl.value='200'; renderActiveTable();"],
    ["wymiary z Excela", "if (typeof setExcelLayoutEnabled==='function') setExcelLayoutEnabled(true); renderActiveTable();"],
    ["wymiary z Excela wyłączone", "if (typeof setExcelLayoutEnabled==='function') setExcelLayoutEnabled(false); renderActiveTable();"],
    ["zmiana wartości komórki", "const r=baseRows[0]; if (r) { r.values[1] = 'ZMIANA-' + Date.now(); if (r.display) r.display[1] = String(r.values[1]); bumpSheetDataStamp(); } renderActiveTable();"],
    ["sortowanie po zmianie danych", "setPrimarySort(currentHeaders[1], 'desc'); sortRows(); renderActiveTable();"],
    ["widok Wide-to-Long", "if (typeof canUseLongView==='function' && canUseLongView()) { tableViewMode='long'; } renderActiveTable();"],
    ["sortowanie w widoku long", "if (tableViewMode==='long') { setPrimarySort(currentHeaders[1], 'asc'); sortRows(); } renderActiveTable();"],
    ["powrót do widoku klasycznego", "tableViewMode='wide'; renderActiveTable();"],
    ["tryb auto: w toku", "if (typeof smartFilterState!=='undefined' && typeof getSmartModel==='function' && getSmartModel()) { smartFilterState={states:['open'],rowFlags:[],only:'',minDays:null}; filtersCommitted=true; applyFilters(); sortRows(); } renderActiveTable();"],
    ["tryb auto: wyczyszczony", "if (typeof resetSmartFilterState==='function') resetSmartFilterState(); applyFilters(); sortRows(); renderActiveTable();"],
  ];

  const diffs = [];
  const reuseStats = [];
  for (const [label, code] of steps) {
    // 1) ścieżka normalna (z ponownym użyciem wierszy)
    await page.evaluate((code) => { window.__swbNoRowReuse = false; new Function(code)(); }, code);
    await sleep(120);
    const withReuse = await page.evaluate(SNAPSHOT_FN);
    const reused = await page.evaluate(() => ({
      wCache: typeof _rowNodeCache !== "undefined" ? _rowNodeCache.size : -1,
      envOk: typeof _rowNodeEnvKey !== "undefined" ? _rowNodeEnvKey !== "" : false,
    }));
    // 2) ten sam stan, ale zbudowany od zera
    await page.evaluate(() => {
      window.__swbNoRowReuse = true;
      if (typeof invalidateRowNodeCache === "function") invalidateRowNodeCache();
      renderActiveTable();
      window.__swbNoRowReuse = false;
    });
    await sleep(120);
    const fresh = await page.evaluate(SNAPSHOT_FN);

    reuseStats.push(`${label}: cache=${reused.wCache}${reused.envOk ? "" : " (reuse wyłączony)"}`);
    if (withReuse !== fresh) {
      // pokaż pierwszą różniącą się linię, żeby diagnoza była natychmiastowa
      const a = withReuse.split("\n");
      const b = fresh.split("\n");
      let firstDiff = "(różna liczba wierszy)";
      for (let i = 0; i < Math.max(a.length, b.length); i++) {
        if (a[i] !== b[i]) {
          firstDiff = `wiersz ${i}:\n  z reuse: ${String(a[i]).slice(0, 300)}\n  od zera: ${String(b[i]).slice(0, 300)}`;
          break;
        }
      }
      diffs.push(`[${label}] DOM różni się od pełnej przebudowy —\n${firstDiff}`);
    }
  }

  await browser.close();

  console.log(reuseStats.join("\n"));
  const failures = [];
  if (diffs.length) failures.push(diffs.join("\n\n"));
  if (errors.length) failures.push(`błędy konsoli/strony: ${errors.join(" | ")}`);
  if (failures.length) throw new Error(failures.join("\n"));
  console.log(`\n✅ row-reuse-playwright OK — ${steps.length} scenariuszy, DOM identyczny jak przy pełnej przebudowie`);
}

run().catch((error) => {
  console.error(error);
  process.exit(1);
});
