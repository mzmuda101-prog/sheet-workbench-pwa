// cell-menu-multi-playwright.js — menu komórki (prawy klik) dla zaznaczenia WIELU komórek.
//
// Co musi być prawdą:
//   1. prawy klik W OBRĘBIE zaznaczenia zostawia zaznaczenie i pokazuje menu zaznaczenia,
//   2. prawy klik POZA zaznaczeniem działa jak dawniej (jedna komórka),
//   3. jedna kolumna × kilka wierszy: „Pokaż" zostawia dokładnie wiersze z którąkolwiek
//      z tych wartości (liczone niezależnie na danych), „Ukryj" — dokładnie resztę,
//   4. kilka kolumn × kilka wierszy: to samo dla KOMBINACJI wartości z każdego wiersza,
//      a pokazane + ukryte = wszystkie wiersze (dopełnienie),
//   5. zaznaczenie w jednym wierszu (kilka kolumn) dokleja się do bieżącego szukania przez &&,
//   6. „Kopiuj zaznaczenie" kopiuje cały prostokąt.

const { chromium } = require("playwright");
const path = require("path");

const APP_URL = process.env.APP_URL || "http://127.0.0.1:4175/";
const FILE = path.join(__dirname, "stress-test-workbench.xlsx");
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

async function run() {
  const browser = await chromium.launch({ headless: true });
  const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1280, height: 900 } });
  await context.addInitScript(() => localStorage.setItem("introPlayed", "true"));
  await context.grantPermissions(["clipboard-read", "clipboard-write"]);
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push("pageerror: " + e.message));

  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  await page.evaluate(() => ensureXlsxLibs(false));
  await page.setInputFiles("#fileInput", FILE);
  await page.waitForFunction(() => document.getElementById("sheetSelect")?.options?.length > 0, null, { timeout: 20000 });
  await page.click("#loadBtn");
  await page.waitForFunction(() => document.querySelectorAll("#dataTable tbody tr").length > 5, null, { timeout: 20000 });
  await sleep(1000);
  await page.evaluate(() => setSidebarOpen(false));
  await sleep(400);

  const results = [];
  const check = (name, ok, detail) => results.push({ name, ok: !!ok, detail });

  const cellPos = (rowNth, col) => page.evaluate(([rowNth, col]) => {
    const tr = document.querySelectorAll("#dataTable tbody tr")[rowNth];
    const td = tr.querySelector(`td[data-col-index="${col}"]`);
    td.scrollIntoView({ block: "nearest", inline: "nearest" });
    const r = td.getBoundingClientRect();
    return { x: r.left + r.width / 2, y: r.top + r.height / 2 };
  }, [rowNth, col]);
  const selectRange = async (r0, c0, r1, c1) => {
    const a = await cellPos(r0, c0);
    await page.mouse.click(a.x, a.y);
    await sleep(80);
    const b = await cellPos(r1, c1);
    await page.keyboard.down("Shift");
    await page.mouse.click(b.x, b.y);
    await page.keyboard.up("Shift");
    await sleep(120);
    return page.evaluate(() => { const r = getSelectionRectangle(); return r && { rows: r.rowCount, cols: r.colCount, colMin: r.colMin }; });
  };
  const rightClick = async (rowNth, col) => {
    const p = await cellPos(rowNth, col);
    await page.mouse.click(p.x, p.y, { button: "right" });
    await sleep(120);
    return page.evaluate(() => ({
      open: !document.getElementById("cellMenu").classList.contains("hidden"),
      head: document.querySelector("#cellMenu .cell-menu-head")?.textContent || "",
      items: [...document.querySelectorAll("#cellMenu .cell-menu-item")].map((b) => b.textContent),
      rect: (() => { const r = getSelectionRectangle(); return r && { rows: r.rowCount, cols: r.colCount }; })(),
    }));
  };
  const clearSearch = async () => {
    await page.evaluate(() => { document.getElementById("resetFiltersBtn").click(); });
    await sleep(500);
  };
  const counts = () => page.evaluate(() => ({ ...lastAppliedFilters, query: searchQueryEl.value }));
  // Niezależnie: kombinacje wartości z prostokąta (na danych widoku) i ile wierszy CAŁEGO
  // arkusza (baseRows) ma którąkolwiek z nich.
  const expected = (rowFrom, rowTo, colMin, colMax) => page.evaluate(([rowFrom, rowTo, colMin, colMax]) => {
    const norm = (v) => String(v ?? "").replace(/\s+/g, " ").trim().toLowerCase();
    const key = (row) => { const k = []; for (let c = colMin; c <= colMax; c++) k.push(norm(getDisplayValue(row, c))); return k.join("\u0000"); };
    const combos = new Set();
    for (let i = rowFrom; i <= rowTo; i++) combos.add(key(currentDisplayModel.rows[i]));
    const hit = baseRows.filter((r) => combos.has(key(r))).length;
    return { combos: combos.size, hit, total: baseRows.length };
  }, [rowFrom, rowTo, colMin, colMax]);

  // ── 1. jedna kolumna × 5 wierszy (kolumna 1)
  const rect1 = await selectRange(0, 1, 4, 1);
  check("Shift+klik tworzy zakres 5×1 (warunek testu)", rect1 && rect1.rows === 5 && rect1.cols === 1, JSON.stringify(rect1));
  const exp1 = await expected(0, 4, 1, 1);
  const m1 = await rightClick(2, 1);
  check("prawy klik w zaznaczeniu: menu zaznaczenia", m1.open && m1.items.some((x) => /tymi wartościami/.test(x)), JSON.stringify(m1.items));
  check("…a zaznaczenie zostaje (5×1)", m1.rect && m1.rect.rows === 5 && m1.rect.cols === 1, JSON.stringify(m1.rect));
  check("nagłówek menu mówi, ile komórek i różnych wartości", m1.head.includes("5") && m1.head.includes(String(exp1.combos)), m1.head);
  await page.click("#cellMenu .cmi-show");
  await sleep(600);
  const s1 = await counts();
  check("1 kolumna, „Pokaż”: dokładnie wiersze z tymi wartościami", s1.filtering && s1.matched === exp1.hit, `${s1.matched} vs oczekiwane ${exp1.hit} · ${s1.query}`);
  await clearSearch();
  await selectRange(0, 1, 4, 1);
  await rightClick(2, 1);
  await page.click("#cellMenu .cmi-hide");
  await sleep(600);
  const h1 = await counts();
  check("1 kolumna, „Ukryj”: dokładnie reszta", h1.filtering && h1.matched === exp1.total - exp1.hit, `${h1.matched} vs ${exp1.total - exp1.hit} · ${h1.query}`);
  await clearSearch();

  // ── 2. dwie kolumny × 6 wierszy (kolumny 1–2) — kombinacje
  const rect2 = await selectRange(1, 1, 6, 2);
  check("zakres 6×2 (warunek testu)", rect2 && rect2.rows === 6 && rect2.cols === 2, JSON.stringify(rect2));
  const exp2 = await expected(1, 6, 1, 2);
  await rightClick(3, 2);
  await page.click("#cellMenu .cmi-show");
  await sleep(600);
  const s2 = await counts();
  check("2 kolumny, „Pokaż”: wiersze z którąkolwiek kombinacją", s2.filtering && s2.matched === exp2.hit, `${s2.matched} vs ${exp2.hit} · ${s2.query}`);
  await clearSearch();
  await selectRange(1, 1, 6, 2);
  await rightClick(3, 2);
  await page.click("#cellMenu .cmi-hide");
  await sleep(600);
  const h2 = await counts();
  check("2 kolumny, „Ukryj”: dokładnie reszta", h2.filtering && h2.matched === exp2.total - exp2.hit, `${h2.matched} vs ${exp2.total - exp2.hit} · ${h2.query}`);
  check("pokazane + ukryte = wszystkie", s2.matched + h2.matched === exp2.total, `${s2.matched} + ${h2.matched} vs ${exp2.total}`);
  await clearSearch();

  // ── 3. prawy klik POZA zaznaczeniem = stare menu jednej komórki
  await selectRange(0, 1, 2, 1);
  const m3 = await rightClick(8, 3);
  check("prawy klik poza zaznaczeniem: menu jednej komórki", m3.open && m3.items.includes("Pokaż tylko takie"), JSON.stringify(m3.items));
  await page.keyboard.press("Escape");
  await sleep(150);

  // ── 4. jeden wiersz × 2 kolumny dokleja się przez && do bieżącego szukania
  const first = await cellPos(0, 1);
  await page.mouse.click(first.x, first.y, { button: "right" });
  await sleep(120);
  await page.click("#cellMenu .cmi-show"); // filtr po jednej wartości kolumny 1
  await sleep(600);
  const before = await counts();
  await selectRange(0, 1, 0, 3);
  await rightClick(0, 2);
  await page.click("#cellMenu .cmi-show");
  await sleep(600);
  const after = await counts();
  check("jeden wiersz × kilka kolumn: doklejone przez && (bez zastępowania)", after.query.startsWith(before.query) && after.query.includes("&&") && after.matched <= before.matched
    && after.query.split(before.query).length === 2 /* warunek z bieżącego szukania nie jest dublowany */, `${before.query} → ${after.query}`);
  await clearSearch();

  // ── 5. kopiowanie prostokąta
  await selectRange(0, 1, 2, 2);
  await rightClick(1, 1);
  await page.click("#cellMenu .cmi-copy");
  await sleep(300);
  const clip = await page.evaluate(() => internalClipboard && { rows: internalClipboard.rows, cols: internalClipboard.cols });
  check("„Kopiuj zaznaczenie” kopiuje 3×2", clip && clip.rows === 3 && clip.cols === 2, JSON.stringify(clip));

  check("brak błędów strony", errors.length === 0, errors.slice(0, 3).join(" | "));
  await browser.close();

  let failed = 0;
  for (const r of results) {
    console.log(`${r.ok ? "✅" : "❌"} ${r.name}${r.detail ? `  (${r.detail})` : ""}`);
    if (!r.ok) failed += 1;
  }
  if (failed) { console.error(`\n${failed} z ${results.length} nie przeszło`); process.exit(1); }
  console.log(`\n✅ cell-menu-multi: ${results.length}/${results.length}`);
}

run().catch((e) => { console.error(e); process.exit(1); });
