// report-playwright.js — Raport do druku (krok 1: wersja krótka, 1 strona A4).
//
// Sprawdzamy, że raport mówi PRAWDĘ o bieżącym widoku:
//   1. kafelki: wiersze, kolumny z danymi, suma kolumny liczbowej (nie numeru porządkowego!),
//      najczęstsza kategoria, zakres dat,
//   2. wnioski: rozkład statusu, wartość odstająca, duplikaty w „Nr”, braki danych,
//   3. filtr → kartka mówi „X z Y” i jaki to filtr; liczby liczone z przefiltrowanych wierszy,
//   4. styl / kolor / rozmiar zmieniają kartkę i są zapamiętane po przeładowaniu,
//   5. druk: w trybie print widać TYLKO kartkę (bez paska narzędzi i bez apki), bez skali,
//   6. pusty widok → komunikat zamiast pustej kartki; Esc zamyka; brak błędów w konsoli.

const { chromium } = require("playwright");
const APP_URL = process.env.APP_URL || "http://127.0.0.1:4175/";
const SLEEP_SCALE = Math.min(1, Math.max(0.1, Number(process.env.SLEEP_SCALE || 1)));
const sleep = (ms) => new Promise((r) => setTimeout(r, Math.round(ms * SLEEP_SCALE)));

// 10 zleceń: Nr (z jednym duplikatem), Status, Kwota (jedna odstająca), Data, Uwagi (rzadko).
const ROWS = [
  [1, "W toku", 100, "2026-01-05", ""],
  [2, "W toku", 120, "2026-01-20", "pilne"],
  [3, "Zakończone", 90, "2026-02-03", ""],
  [4, "W toku", 110, "2026-02-14", ""],
  [5, "Zakończone", 95, "2026-02-20", ""],
  [6, "W toku", 105, "2026-02-25", ""],
  [7, "Anulowane", 1000, "2026-02-27", "sprawdzić"],
  [7, "W toku", 115, "2026-03-10", ""],
  [9, "Zakończone", 100, "2026-03-15", ""],
  [10, "W toku", 125, "2026-03-28", ""],
];

async function loadSheet(page) {
  await page.evaluate((rows) => {
    const ws = {};
    ["Nr", "Status", "Kwota", "Data", "Uwagi", "Pusta"].forEach((h, i) => { ws[String.fromCharCode(65 + i) + "1"] = { t: "s", v: h }; });
    rows.forEach(([nr, st, kw, dt, uw], i) => {
      const r = i + 2;
      ws["A" + r] = { t: "n", v: nr, w: String(nr) };
      ws["B" + r] = { t: "s", v: st };
      ws["C" + r] = { t: "n", v: kw, w: String(kw) };
      ws["D" + r] = { t: "s", v: dt };
      if (uw) ws["E" + r] = { t: "s", v: uw };
    });
    ws["!ref"] = "A1:F" + (rows.length + 1);
    workbook = { SheetNames: ["Zlecenia"], Sheets: { Zlecenia: ws }, Props: {} };
    currentFileName = "zlecenia.xlsx";
    sheetSelect.replaceChildren();
    const opt = document.createElement("option"); opt.value = "Zlecenia"; opt.textContent = "Zlecenia";
    sheetSelect.appendChild(opt); sheetSelect.value = "Zlecenia";
    document.getElementById("headerRow").value = "1";
    document.getElementById("autoHeaderRow").checked = false;
  }, ROWS);
  await page.evaluate(() => document.getElementById("loadBtn").click());
  await page.waitForFunction((n) => typeof baseRows !== "undefined" && baseRows.length === n, ROWS.length, { timeout: 15000 });
  await sleep(350);
}

const tiles = (page) => page.evaluate(() => Array.from(document.querySelectorAll("#rpPage .rp-tile")).map((el) => ({
  label: el.querySelector(".rp-tile-label").textContent,
  value: el.querySelector(".rp-tile-value").textContent,
  sub: el.querySelector(".rp-tile-sub")?.textContent || "",
})));
const tile = (list, re) => list.find((x) => re.test(x.label));
const pageText = (page) => page.evaluate(() => document.getElementById("rpPage").innerText);

async function run() {
  const browser = await chromium.launch({ headless: true });
  const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1100, height: 820 } });
  await context.addInitScript(() => {
    localStorage.setItem("introPlayed", "true");
    if (!sessionStorage.getItem("rpInit")) {
      sessionStorage.setItem("rpInit", "1");
      localStorage.removeItem("swb-report-prefs");
    }
  });
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push("pageerror: " + e.message));
  page.on("console", (m) => { if (m.type() === "error") errors.push("console: " + m.text()); });

  const results = [];
  const check = (name, cond, got) => {
    results.push({ name, ok: !!cond });
    console.log(`${cond ? "✅" : "❌"} ${name}${cond ? "" : ` — dostałem: ${JSON.stringify(got)}`}`);
  };

  const boot = async () => {
    await page.goto(APP_URL, { waitUntil: "load" });
    await page.evaluate(() => document.getElementById("heroSplash")?.remove());
    await page.evaluate(() => ensureXlsxLibs(false));
    await page.waitForFunction(() => typeof buildRows === "function" && window.__report);
  };
  await boot();

  // ── 0. Pusty arkusz → komunikat, nie pusta kartka ─────────────────────────
  await page.evaluate(() => window.__report.open());
  check("bez danych raport się nie otwiera", !(await page.evaluate(() => window.__report.isOpen())));

  // ── 1. Cały arkusz ───────────────────────────────────────────────────────
  await loadSheet(page);
  await page.evaluate(() => document.getElementById("reportBtn").click());
  await sleep(250);
  check("raport otwarty przyciskiem", await page.evaluate(() => window.__report.isOpen()));
  const t1 = await tiles(page);
  check("kafelek wierszy: 10, cały arkusz", tile(t1, /Wiersze/).value === "10" && /cały arkusz/.test(tile(t1, /Wiersze/).sub), t1);
  check("kolumny z danymi: 5 (pusta pominięta)", tile(t1, /Kolumny/).value === "5" && /1 pustych/.test(tile(t1, /Kolumny/).sub), t1);
  const sumTile = tile(t1, /Suma/);
  check("suma liczy Kwotę, nie Nr", sumTile && /Kwota/.test(sumTile.label) && sumTile.value.replace(/\s| /g, "") === "1960", sumTile);
  const topTile = tile(t1, /Najczęściej/);
  check("najczęstszy status: W toku 60%", topTile && /Status/.test(topTile.label) && topTile.value === "W toku" && /60%/.test(topTile.sub), topTile);
  const dateTile = tile(t1, /Daty/);
  check("zakres dat: styczeń – marzec 2026", dateTile && /sty/.test(dateTile.value) && /mar/.test(dateTile.value) && /2026/.test(dateTile.value), dateTile);

  const f1 = await page.evaluate(() => window.__report.findings());
  const has = (re) => f1.some((f) => re.test(f));
  check("wniosek: rozkład statusu", has(/„Status” najczęściej: W toku — 60% \(6 z 10\)/), f1);
  check("wniosek: wartość odstająca 1000", has(/Największa wartość w „Kwota” \(1[\s ]?000\)/), f1);
  check("wniosek: duplikat w Nr (7 ×2)", has(/„Nr” powtarza się 1 wartości \(7 ×2\)/), f1);
  check("wniosek: braki w Uwagi 80%", has(/„Uwagi” pusta w 80% wierszy/), f1);
  check("wniosek: najwięcej w lutym", has(/luty 2026 \(5\)/), f1);
  const liCount = await page.evaluate(() => document.querySelectorAll("#rpPage .rp-findings li").length);
  check("krótki: na kartce najwyżej 6 wniosków", liCount === Math.min(6, f1.length), { liCount, all: f1.length });

  const text1 = await pageText(page);
  check("zakres: cały arkusz", /Cały arkusz: 10 wierszy/.test(text1), text1.slice(0, 300));
  check("tytuł domyślny z nazwy pliku", /Raport — zlecenia/.test(text1), text1.slice(0, 120));
  check("wykres rozkładu statusu", await page.evaluate(() => document.querySelectorAll("#rpPage .rp-bar-row").length === 3));
  check("mieści się na 1 stronie", /1 stronie/.test(await page.evaluate(() => document.getElementById("rpFitNote").textContent)));

  // ── 1b. Presety i wybór sekcji ───────────────────────────────────────────
  const secs = () => page.evaluate(() => window.__report.sections());
  check("domyślnie krótki: liczby, wnioski, wykres", (await secs()).join() === "tiles,findings,chart", await secs());
  await page.click("#rpContentBtn");
  check("panel zawartości otwarty", await page.evaluate(() => !document.getElementById("rpContentPanel").classList.contains("hidden")));
  await page.click('#rpPresets [data-preset="normal"]');
  await sleep(100);
  check("normalny: + miesiące i statystyki liczb", (await secs()).join() === "tiles,findings,chart,months,numbers", await secs());
  const catRow = await page.evaluate(() => {
    const cb = document.querySelector('#rpSectionList input[value="categories"]');
    return { disabled: cb.disabled, why: cb.closest("label").textContent };
  });
  check("sekcja bez danych wyłączona z powodem", catRow.disabled && /brak w tym arkuszu/.test(catRow.why), catRow);
  const numTable = await page.evaluate(() => Array.from(document.querySelectorAll('#rpPage [data-section="numbers"] tbody tr')).map((tr) => tr.cells[0].textContent));
  check("statystyki liczb: tylko Kwota (bez Nr)", numTable.join() === "Kwota", numTable);
  const monthBars = await page.evaluate(() => Array.from(document.querySelectorAll('#rpPage [data-section="months"] .rp-bar-label')).map((el) => el.textContent));
  check("miesiące po kolei: sty, lut, mar", monthBars.length === 3 && /stycze/.test(monthBars[0]) && /marzec/.test(monthBars[2]), monthBars);
  await page.click('#rpPresets [data-preset="detailed"]');
  await sleep(100);
  check("szczegółowy: + przegląd kolumn i dane", (await secs()).join() === "tiles,findings,chart,months,numbers,columns,data", await secs());
  const detail = await page.evaluate(() => ({
    dataRows: document.querySelectorAll('#rpPage [data-section="data"] tbody tr').length,
    dataCols: document.querySelectorAll('#rpPage [data-section="data"] thead th').length,
    colRows: Array.from(document.querySelectorAll('#rpPage [data-section="columns"] tbody tr')).map((tr) => `${tr.cells[0].textContent}:${tr.cells[1].textContent}`),
  }));
  check("tabela danych: 10 wierszy, 5 kolumn z danymi", detail.dataRows === 10 && detail.dataCols === 5, detail);
  check("przegląd kolumn rozpoznaje rodzaje", detail.colRows.join() === "Nr:liczby,Status:tekst,Kwota:liczby,Data:daty,Uwagi:tekst,Pusta:pusta", detail.colRows);
  await page.evaluate(() => document.querySelector('#rpSectionList input[value="findings"]').click());
  await sleep(100);
  const custom = await page.evaluate(() => ({ preset: window.__report.prefs().preset, btn: document.getElementById("rpContentBtn").textContent, customShown: !document.getElementById("rpPresetCustom").classList.contains("hidden") }));
  check("ręczna zmiana = własny zestaw", custom.preset === "custom" && /własny/.test(custom.btn) && custom.customShown, custom);
  check("wnioski zniknęły z kartki", !(await secs()).includes("findings"), await secs());
  await page.evaluate(() => document.querySelector('#rpSectionList input[value="findings"]').click());
  await sleep(100);
  check("powrót do zestawu presetu = znowu szczegółowy", (await page.evaluate(() => window.__report.prefs().preset)) === "detailed");
  await page.click('#rpPresets [data-preset="short"]');
  await page.keyboard.press("Escape");
  check("Esc najpierw zamyka panel zawartości", await page.evaluate(() => document.getElementById("rpContentPanel").classList.contains("hidden") && window.__report.isOpen()));

  // ── 2. Filtr → kartka mówi „X z Y” i liczy tylko wycinek ─────────────────
  await page.keyboard.press("Escape");
  await sleep(100);
  check("Esc zamyka raport", !(await page.evaluate(() => window.__report.isOpen())));
  await page.evaluate(() => {
    searchQueryEl.value = 'Status:="W toku"';
    filterOperatorsEl.checked = true;
    syncQuickSearchOperatorsControls();
    filtersCommitted = true;
    applyFilters();
    sortRows();
    scheduleViewRefresh({ table: true, immediate: true });
    window.__report.open();
  });
  await sleep(250);
  const t2 = await tiles(page);
  check("filtr: 6 wierszy z 10", tile(t2, /Wiersze/).value === "6" && /z 10/.test(tile(t2, /Wiersze/).sub), t2);
  check("filtr: suma tylko z wycinka (675)", tile(t2, /Suma/).value.replace(/\s| /g, "") === "675", tile(t2, /Suma/));
  const text2 = await pageText(page);
  check("zakres opisuje filtr", /Status = W toku — 6 z 10 wierszy/.test(text2), text2.slice(0, 300));
  check("zakres przefiltrowany wyróżniony", await page.evaluate(() => document.querySelector("#rpPage .rp-scope").classList.contains("is-filtered")));

  // ── 3. Tytuł edytowalny na kartce → trafia do nazwy PDF ──────────────────
  await page.evaluate(() => {
    const h = document.querySelector("#rpPage .rp-title");
    h.textContent = "Zlecenia w toku";
    h.dispatchEvent(new Event("input", { bubbles: true }));
  });

  // ── 4. Styl / kolor / rozmiar ────────────────────────────────────────────
  await page.selectOption("#rpStyle", "classic");
  await page.click('#rpAccents [data-accent="blue"]');
  await page.click("#rpSizeBtn");
  await sleep(100);
  const look = await page.evaluate(() => {
    const p = document.getElementById("rpPage");
    return { style: p.dataset.style, size: p.dataset.size, accent: p.style.getPropertyValue("--rp-accent").trim(), font: getComputedStyle(p).fontFamily };
  });
  check("styl klasyczny = szeryf", look.style === "classic" && /Georgia|Times/.test(look.font), look);
  check("kolor niebieski i duży tekst", look.accent === "#2b5d9b" && look.size === "large", look);
  check("tytuł przeżywa zmianę stylu", /Zlecenia w toku/.test(await pageText(page)));
  await page.selectOption("#rpStyle", "ink");
  await sleep(50);
  check("styl oszczędny wyłącza wybór koloru", await page.evaluate(() => Array.from(document.querySelectorAll("#rpAccents [data-accent]")).every((b) => b.disabled)));

  // ── 5. Druk: tylko kartka ────────────────────────────────────────────────
  await page.evaluate(() => {
    window.__printed = null;
    window.print = () => {
      window.__printed = {
        cls: document.body.classList.contains("rp-printing"),
        title: document.title,
      };
    };
    window.__report.print();
  });
  const printed = await page.evaluate(() => window.__printed);
  check("druk wołany z klasą rp-printing i tytułem raportu", printed && printed.cls && printed.title === "Zlecenia w toku", printed);
  check("po druku tytuł strony wraca", await page.evaluate(() => document.title === "Sheet Workbench" || !/Zlecenia w toku/.test(document.title)));
  await page.evaluate(() => document.body.classList.add("rp-printing"));
  await page.emulateMedia({ media: "print" });
  const printView = await page.evaluate(() => {
    const vis = (el) => !!el && getComputedStyle(el).display !== "none";
    const p = document.getElementById("rpPage");
    return {
      app: vis(document.querySelector(".app")),
      toolbar: vis(document.querySelector(".rp-toolbar")),
      page: vis(p),
      transform: getComputedStyle(p).transform,
      shadow: getComputedStyle(p).boxShadow,
    };
  });
  check("druk: apka ukryta, pasek ukryty, kartka widoczna", !printView.app && !printView.toolbar && printView.page, printView);
  check("druk: kartka bez skali i cienia", printView.transform === "none" && printView.shadow === "none", printView);
  await page.emulateMedia({ media: "screen" });
  await page.evaluate(() => document.body.classList.remove("rp-printing"));
  // Zwykły eksport tabeli (bez klasy raportu) nadal drukuje #printArea.
  await page.emulateMedia({ media: "print" });
  const tablePrint = await page.evaluate(() => getComputedStyle(document.getElementById("reportOverlay")).display);
  check("druk tabeli (bez raportu) chowa nakładkę raportu", tablePrint === "none", tablePrint);
  await page.emulateMedia({ media: "screen" });

  // ── 6. Ustawienia wyglądu zapamiętane ─────────────────────────────────────
  await boot();
  await loadSheet(page);
  await page.evaluate(() => window.__report.open());
  await sleep(200);
  const prefs = await page.evaluate(() => window.__report.prefs());
  check("po przeładowaniu: ink + blue + large", prefs.style === "ink" && prefs.accent === "blue" && prefs.size === "large", prefs);

  // Wąski ekran: kartka przeskalowana, bez poziomego przewijania strony.
  await page.setViewportSize({ width: 390, height: 760 });
  await sleep(200);
  const narrow = await page.evaluate(() => {
    const wrap = document.getElementById("rpSheetWrap");
    const stage = document.getElementById("rpStage");
    return { wrap: wrap.getBoundingClientRect().width, stage: stage.clientWidth, docW: document.documentElement.scrollWidth };
  });
  check("telefon: kartka mieści się w szerokości", narrow.wrap <= narrow.stage && narrow.docW <= 390, narrow);

  check("brak błędów w konsoli", errors.length === 0, errors);
  await browser.close();
  const failed = results.filter((r) => !r.ok);
  if (failed.length) {
    console.log(`\n❌ report: ${failed.length} z ${results.length} nie przeszło`);
    process.exit(1);
  }
  console.log(`\n✅ report: ${results.length}/${results.length}`);
}

run().catch((e) => { console.error(e); process.exit(1); });
