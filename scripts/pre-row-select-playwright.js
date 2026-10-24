// pre-row-select-playwright.js — zaznaczenie komórki w wierszach NAD nagłówkiem (pre-row).
//
// Zgłoszenie z iPhone'a (2026-10-24): stuknięcie komórki nad nagłówkiem (np. C2 w przykładzie)
// pokazywało „Odznacz”, ale stuknięcie w „Odznacz” zostawiało jej obwódkę — komórki pre-row
// mają własne zaznaczenie (klasa pre-active), o którym „Odznacz” i Esc nie wiedziały.
// Pilnuje:
//   1. telefon: stuknięcie pre-cell → obwódka + „Odznacz”; zwykła komórka traci fokus (jedna aktywna),
//   2. telefon: „Odznacz” zdejmuje obwódkę i sam znika,
//   3. desktop: Esc zdejmuje obwódkę pre-cell,
//   4. desktop: klik w zwykłą komórkę zdejmuje obwódkę pre-cell.
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
  pre: document.querySelectorAll("#dataTable thead td.pre-active").length,
  focused: !!focusedCellState,
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
