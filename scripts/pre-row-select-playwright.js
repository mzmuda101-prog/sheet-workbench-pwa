// pre-row-select-playwright.js — zaznaczenie komórki w wierszach NAD nagłówkiem (pre-row).
//
// Zgłoszenie z iPhone'a (2026-10-24): stuknięcie komórki nad nagłówkiem (np. C2 w przykładzie)
// pokazywało „Odznacz”, ale stuknięcie w „Odznacz” zostawiało jej obwódkę — komórki pre-row
// mają własne zaznaczenie (klasa pre-active), o którym „Odznacz” i Esc nie wiedziały.
// Od 2026-10-31 wiersze nad nagłówkiem są w TYM SAMYM modelu zaznaczenia co dane (klucz
// „pre:N”, table.js swbSelView) — dawniej miały osobne pre-active i działał tylko dwuklik.
// Pilnuje:
//   1. telefon: stuknięcie pre-cell → obwódka + „Odznacz”; zwykła komórka traci fokus (jedna aktywna),
//   2. telefon: „Odznacz” zdejmuje obwódkę i sam znika; drugie stuknięcie = edycja,
//   3. desktop: Esc zdejmuje obwódkę pre-cell,
//   4. desktop: klik w zwykłą komórkę zdejmuje obwódkę pre-cell,
//   5. desktop: nad nagłówkiem działa to samo co w danych — pole nazwy, Ctrl+B, pisanie,
//      Delete + Cofnij, strzałki przez granicę nagłówka, Shift+strzałka, kopiowanie zakresu,
//      menu (bez „Pokaż/Ukryj wiersze” — to nie rekordy), uchwyt wypełniania.
//
// Uruchom z serwerem na APP_URL (domyślnie http://127.0.0.1:4175/). ENGINE=webkit — Safari.

const playwright = require("playwright");

const APP_URL = process.env.APP_URL || "http://127.0.0.1:4175/";
const ENGINE = process.env.ENGINE || "chromium";
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

async function open(browser, opts) {
  const ctx = await browser.newContext({ serviceWorkers: "block", ...opts });
  await ctx.addInitScript(() => {
    if (sessionStorage.getItem("__prInit")) return;
    sessionStorage.setItem("__prInit", "1");
    localStorage.clear();
    localStorage.setItem("introPlayed", "true");
  });
  const page = await ctx.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(() => { document.getElementById("heroSplash")?.remove(); return loadSampleFile(); });
  await page.waitForFunction(() => document.querySelector("#dataTable thead tr.pre-row td.pre-cell") && document.querySelector("#dataTable tbody tr[data-row-key]"), null, { timeout: 20000 });
  await page.evaluate(() => document.querySelectorAll(".toast").forEach((t) => t.remove()));
  await sleep(500);
  return { ctx, page, errors };
}

const state = (p) => p.evaluate(() => ({
  pre: document.querySelectorAll("#dataTable thead tr.pre-row td.cell-active").length,
  focused: !!focusedCellState && !focusedCellState.rowKey.startsWith("pre:"),
  fab: document.getElementById("clearSelectionFab").classList.contains("is-visible"),
}));
const center = (p, sel, i = 0) => p.evaluate(([s, n]) => {
  const b = document.querySelectorAll(s)[n].getBoundingClientRect();
  return { x: b.left + b.width / 2, y: b.top + b.height / 2 };
}, [sel, i]);
const PRE = '#dataTable thead tr.pre-row td.pre-cell[data-col-index="1"]';
const BODY = '#dataTable tbody tr[data-row-key] td[data-col-index="2"]';

async function run() {
  const browser = await playwright[ENGINE].launch({ headless: true });
  const failures = [];
  const ok = (name, cond, got) => { if (!cond) failures.push(`${name} — dostałem: ${JSON.stringify(got)}`); };

  // ── telefon ──
  const phone = await open(browser, { viewport: { width: 390, height: 844 }, hasTouch: true, isMobile: ENGINE === "chromium" });
  let p = phone.page;
  let pt = await center(p, BODY, 2);
  await p.touchscreen.tap(pt.x, pt.y);
  await sleep(300);
  const s0 = await state(p);
  ok("0. stuknięcie zwykłej komórki daje fokus i „Odznacz”", s0.focused && s0.fab && !s0.pre, s0);
  pt = await center(p, PRE);
  await p.touchscreen.tap(pt.x, pt.y);
  await sleep(300);
  const s1 = await state(p);
  ok("1. pre-cell: obwódka + „Odznacz”, zwykła komórka bez fokusu", s1.pre === 1 && s1.fab && !s1.focused, s1);
  pt = await center(p, "#clearSelectionFab");
  await p.touchscreen.tap(pt.x, pt.y);
  await sleep(300);
  const s2 = await state(p);
  ok("2. „Odznacz” zdejmuje obwódkę pre-cell i znika", s2.pre === 0 && !s2.fab && !s2.focused, s2);
  pt = await center(p, PRE);
  await p.touchscreen.tap(pt.x, pt.y);
  await sleep(300);
  await p.touchscreen.tap(pt.x, pt.y);
  await sleep(300);
  const s2b = await p.evaluate(() => !!document.querySelector("#dataTable thead tr.pre-row input.cell-editor"));
  ok("2b. drugie stuknięcie w pre-cell otwiera edycję", s2b, s2b);
  await p.evaluate(() => { if (activeCellEditor) activeCellEditor.close(); });
  ok("telefon: bez błędów strony", phone.errors.length === 0, phone.errors);
  await phone.ctx.close();

  // ── desktop ──
  const desk = await open(browser, { viewport: { width: 1280, height: 900 } });
  p = desk.page;
  pt = await center(p, PRE);
  await p.mouse.click(pt.x, pt.y);
  await sleep(200);
  const d1 = await state(p);
  ok("3a. klik w pre-cell zaznacza ją", d1.pre === 1, d1);
  await p.keyboard.press("Escape");
  await sleep(200);
  const d2 = await state(p);
  ok("3b. Esc zdejmuje obwódkę pre-cell", d2.pre === 0 && !d2.fab, d2);
  await p.mouse.click(pt.x, pt.y);
  await sleep(200);
  pt = await center(p, BODY, 1);
  await p.mouse.click(pt.x, pt.y);
  await sleep(200);
  const d3 = await state(p);
  ok("4. klik w zwykłą komórkę zdejmuje obwódkę pre-cell", d3.pre === 0 && d3.focused, d3);

  // ── 5. pełna edycja nad nagłówkiem (A2 = „Zakresów łącznie”, B2 = formuła) ──
  await p.evaluate(() => sheetEditUi.setMode("edit"));
  const A2 = '#dataTable thead tr.pre-row[data-row-index="1"] td.pre-cell[data-col-index="0"]';
  const cellV = (ref) => p.evaluate((r) => workbook.Sheets[currentSheetName][r]?.v ?? null, ref);
  const name = () => p.evaluate(() => document.getElementById("nameBox").value);
  await p.click(A2);
  await sleep(200);
  ok("5a. pole nazwy = A2, wstążka aktywna", await name() === "A2" && !(await p.evaluate(() => document.getElementById("formatRibbon").classList.contains("is-disabled"))), await name());
  await p.keyboard.press("Control+b");
  await sleep(200);
  const bold = await p.evaluate(() => ({ fmt: cfmtPendingFmt(currentSheetName, "A2"), w: getComputedStyle(document.querySelector('#dataTable thead tr.pre-row[data-row-index="1"] td.pre-cell[data-col-index="0"]')).fontWeight }));
  ok("5b. Ctrl+B pogrubia komórkę nad nagłówkiem (zapis + widok)", bold.fmt && bold.fmt.b === true && Number(bold.w) >= 600, bold);
  await p.keyboard.type("Tytuł");
  await p.keyboard.press("Enter");
  await sleep(300);
  ok("5c. pisanie od razu + Enter zapisuje i schodzi niżej", await cellV("A2") === "Tytuł" && await name() === "A3", { v: await cellV("A2"), name: await name() });
  await p.keyboard.press("ArrowUp");
  await p.keyboard.press("Delete");
  await sleep(200);
  ok("5d. Delete czyści komórkę nad nagłówkiem", await cellV("A2") === null, await cellV("A2"));
  await p.keyboard.press(process.platform === "darwin" && ENGINE === "webkit" ? "Meta+z" : "Control+z");
  await sleep(300);
  ok("5e. Cofnij przywraca", await cellV("A2") === "Tytuł", await cellV("A2"));
  // granica nagłówka: z ostatniego wiersza nad nim strzałka w dół → 1. wiersz danych i z powrotem
  const lastPre = await p.evaluate(() => [...document.querySelectorAll("#dataTable thead tr.pre-row")].pop().dataset.rowIndex);
  await p.click(`#dataTable thead tr.pre-row[data-row-index="${lastPre}"] td.pre-cell[data-col-index="0"]`);
  await sleep(200);
  await p.keyboard.press("ArrowDown");
  await sleep(200);
  const down = await p.evaluate(() => focusedCellState && focusedCellState.rowKey);
  await p.keyboard.press("ArrowUp");
  await sleep(200);
  const up = await p.evaluate(() => focusedCellState && focusedCellState.rowKey);
  ok("5f. strzałki przechodzą przez wiersz nagłówka w obie strony", /^wide:/.test(down) && up === `pre:${lastPre}`, { down, up });
  await p.click(A2);
  await sleep(200);
  await p.keyboard.press("Shift+ArrowRight");
  await sleep(200);
  const rng = await p.evaluate(() => { const r = getSelectionRectangle(); return { n: r && r.rowCount * r.colCount, cls: document.querySelectorAll("#dataTable thead td.cell-in-range").length }; });
  ok("5g. Shift+strzałka robi zakres nad nagłówkiem (z obwódką)", rng.n === 2 && rng.cls === 2, rng);
  await p.keyboard.press("Control+c");
  await sleep(200);
  const tsv = await p.evaluate(() => internalClipboard && internalClipboard.tsv);
  ok("5h. kopiowanie zakresu nad nagłówkiem", tsv === "Tytuł\t120", tsv);
  await p.click(A2, { button: "right" });
  await sleep(300);
  const menu = await p.evaluate(() => [...document.querySelectorAll("#cellMenu .cell-menu-item")].map((b) => b.className.match(/cmi-[\w-]+/)?.[0]));
  ok("5i. menu komórki działa, bez filtrów rekordów", menu.includes("cmi-copy") && !menu.includes("cmi-show") && !menu.includes("cmi-hide"), menu);
  await p.keyboard.press("Escape");
  await p.click(A2);
  await sleep(300);
  const handle = await p.evaluate(() => { const td = document.querySelector("#dataTable td.has-fill-handle"); return td ? td.parentElement.dataset.rowKey : null; });
  ok("5j. uchwyt wypełniania stoi przy komórce nad nagłówkiem", handle === "pre:1", handle);
  ok("desktop: bez błędów strony", desk.errors.length === 0, desk.errors);
  await desk.ctx.close();

  await browser.close();
  if (failures.length) {
    console.log(`[${ENGINE}] ❌ ${failures.length} problem(y):\n  - ` + failures.join("\n  - "));
    process.exit(1);
  }
  console.log(`[${ENGINE}] ✅ zaznaczenie nad nagłówkiem: Odznacz / Esc / klik w tabelę — OK`);
}

run().catch((e) => { console.error(e); process.exit(1); });
