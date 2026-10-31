// desktop-select-playwright.js — ZAZNACZANIE MYSZĄ I OCHRONA DANYCH PRZY WYPEŁNIANIU (film z 2026-10-31).
//   node scripts/desktop-select-playwright.js        ENGINE=webkit node scripts/desktop-select-playwright.js
//
// Na filmie: klik w pustą D9 → podświetlony wiersz + uchwyt bez ramki; przeciągnięcie uchwytu
// (chęć zaznaczenia D9:D11) wyczyściło datę w D10, potem formuły „Długość” w kolumnie E;
// dymek „Przeliczone na dziś…” zasłaniał zaznaczenie; tytuł okna „Sheet Workbench - * Sheet Workbench”.
// Pilnuje:
//   1. Edycja: klik = komórka (ramka, pole nazwy); Podgląd: klik = wiersz (model przeglądania);
//   2. przeciągnięcie myszą po komórkach = zakres, ŻADNA wartość się nie zmienia; klik po nim nie kasuje zakresu;
//   3. wypełnienie, które nadpisuje formuły / czyści dane → ostrzeżenie z podpowiedzią Cofnij;
//   4. przy wciśniętym przycisku myszy dymek komórki się nie pokazuje;
//   5. tytuł: nazwa pliku, „• ” przy niezapisanych zmianach, bez dubla nazwy aplikacji.
// Uruchom z serwerem na APP_URL (domyślnie http://127.0.0.1:4175/).

const playwright = require("playwright");

const APP_URL = process.env.APP_URL || "http://127.0.0.1:4175/";
const ENGINE = process.env.ENGINE || "chromium";
const SLEEP_SCALE = Math.min(1, Math.max(0.1, Number(process.env.SLEEP_SCALE || 1)));
const sleep = (ms) => new Promise((r) => setTimeout(r, Math.round(ms * SLEEP_SCALE)));
const CELL = (r, c) => `#dataTable tbody tr[data-row-key]:nth-child(${r}) td[data-col-index="${c}"]`;

async function run() {
  const browser = await playwright[ENGINE].launch({ headless: true });
  const fails = [];
  const check = (name, ok, detail) => {
    console.log(`${ok ? "✅" : "❌"} ${name}${ok ? "" : ` — ${JSON.stringify(detail)}`}`);
    if (!ok) fails.push(name);
  };
  const ctx = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1280, height: 900 } });
  await ctx.addInitScript(() => {
    if (sessionStorage.getItem("__dsInit")) return;
    sessionStorage.setItem("__dsInit", "1");
    localStorage.clear();
    localStorage.setItem("introPlayed", "true");
  });
  const p = await ctx.newPage();
  const errors = [];
  p.on("pageerror", (e) => errors.push(e.message));
  await p.goto(APP_URL, { waitUntil: "load" });
  await p.evaluate(() => { document.getElementById("heroSplash")?.remove(); return loadSampleFile(); });
  await p.waitForFunction(() => document.querySelector("#dataTable tbody tr[data-row-key]"), null, { timeout: 20000 });
  await p.waitForFunction(() => !document.documentElement.classList.contains("vt-scene"), null, { timeout: 10000 });
  await p.evaluate(() => { sheetEditUi.setMode("edit"); document.querySelectorAll(".toast").forEach((t) => t.remove()); });
  await sleep(400);
  const center = (sel) => p.evaluate((s) => { const b = document.querySelector(s).getBoundingClientRect(); return { x: b.left + b.width / 2, y: b.top + b.height / 2 }; }, sel);
  const values = () => p.evaluate(() => viewRows.slice(0, 6).map((r) => r.values.slice(0, 5).map((v) => (v instanceof Date ? v.getTime() : v))));

  // 5. tytuł
  const t0 = await p.evaluate(() => document.title);
  check("5. tytuł = nazwa pliku + aplikacja (karta przeglądarki)", /Przyklad-obieg-zakresow\.xlsx — /.test(t0) && !/^\*/.test(t0), t0);

  // 1. klik w Edycji = komórka
  await p.click(CELL(2, 3));
  await sleep(200);
  const s1 = await p.evaluate(() => ({ kind: selectionKind, active: document.querySelectorAll("#dataTable tbody td.cell-active").length, name: document.getElementById("nameBox").value }));
  check("1. Edycja: klik zaznacza komórkę (ramka, pole nazwy)", s1.kind === "cell" && s1.active === 1 && s1.name === "D7", s1);

  // 2. przeciągnięcie myszą = zakres, bez zmian w danych
  const before = await values();
  const a = await center(CELL(2, 1));
  const b = await center(CELL(4, 2));
  await p.mouse.move(a.x, a.y);
  await p.mouse.down();
  await p.mouse.move((a.x + b.x) / 2, (a.y + b.y) / 2, { steps: 4 });
  // 4. dymek przy wciśniętym przycisku
  const tipDuring = await p.evaluate(() => { const el = document.querySelector(".cell-tooltip"); return !!el && !el.classList.contains("hidden") && getComputedStyle(el).display !== "none" && getComputedStyle(el).opacity !== "0"; });
  await p.mouse.move(b.x, b.y, { steps: 4 });
  await p.mouse.up();
  await sleep(300);
  const s2 = await p.evaluate(() => { const r = getSelectionRectangle(); return r && { rows: r.rowCount, cols: r.colCount, colMin: r.colMin, kind: selectionKind }; });
  const after = await values();
  check("2. przeciągnięcie myszą zaznacza zakres 3×2", s2 && s2.rows === 3 && s2.cols === 2 && s2.colMin === 1, s2);
  check("2. …i nie zmienia żadnej wartości", JSON.stringify(before) === JSON.stringify(after), null);
  check("4. przy wciśniętym przycisku nie ma dymka", !tipDuring, tipDuring);

  // 3. wypełnienie nadpisujące formułę → ostrzeżenie (kolumna „Długość” = formuły w przykładzie?)
  const formulaCol = await p.evaluate(() => {
    for (let c = 0; c < currentHeaders.length; c++) {
      const ref = XLSX.utils.encode_cell({ r: viewRows[1].rowIndex0, c: currentStartCol + c });
      if (workbook.Sheets[currentSheetName][ref]?.f) return c;
    }
    return -1;
  });
  check("3. (warunek) przykład ma kolumnę z formułami", formulaCol >= 0, formulaCol);
  const warn = await p.evaluate((fc) => {
    document.querySelectorAll(".toast").forEach((t) => t.remove());
    // źródło: komórka z lewej (bez formuły), cel: komórka z formułą w tym samym wierszu
    const view = swbSelView(); // wiersze nad nagłówkiem są na początku widoku zaznaczenia
    const r = swbSelIndexOfKey(view, getRowSelectionKey(viewRows[1]));
    const src = { rowStart: r, rowEnd: r, colMin: fc - 1, colMax: fc - 1 };
    swbFillRange(view, src, "right", fc, { mode: "copy" });
    return [...document.querySelectorAll(".toast")].map((t) => t.textContent);
  }, formulaCol);
  check("3. ostrzeżenie o nadpisanych formułach z podpowiedzią Cofnij", warn.some((x) => /nadpisało formuły w 1 kom\..*Cofnij/.test(x)), warn);
  const t1 = await p.evaluate(() => document.title);
  check("5. niezapisane zmiany: „• ” przed nazwą pliku", /^• Przyklad-obieg-zakresow\.xlsx/.test(t1), t1);
  await p.evaluate(() => undoRedo("undo"));

  // 1. Podgląd: klik = wiersz
  await p.evaluate(() => sheetEditUi.setMode("view"));
  await p.click(CELL(3, 2));
  await sleep(200);
  const s5 = await p.evaluate(() => selectionKind);
  check("1. Podgląd: klik dalej zaznacza wiersz", s5 === "row", s5);

  check("bez błędów strony", errors.length === 0, errors);
  await browser.close();
  if (fails.length) {
    console.error(`\n${fails.length} błędów (${ENGINE})`);
    process.exit(1);
  }
  console.log(`\nOK (${ENGINE})`);
}

run().catch((e) => { console.error(e); process.exit(1); });
