// hero-scroll-guard-playwright.js — strażnik: zwijanie nagłówka NIE może zmieniać
// właściwości warstwy przewijanej tabeli (#tableWrap).
//
// Dlaczego: nagłówek zwija się W TRAKCIE przewijania palcem. Na iOS zmiana warstwy
// przewijanego elementu (will-change, transform, filter…) w tej chwili przebudowuje
// natywny scroller pod palcem i iOS zrywa gest — tabela staje, trzeba podnieść palec
// i zacząć od nowa. Tak było w buildzie 20260924-01 (will-change: transform na czas
// animacji): w symulatorze iPhone'a ten sam ruch palcem przewijał ~2 wiersze zamiast ~17.
//
// Co musi być prawdą: w trakcie zwijania i rozwijania (także z blokadą kolumny)
// wszystkie właściwości wpływające na warstwę #tableWrap są IDENTYCZNE jak w spoczynku.

const { chromium, webkit } = require("playwright");
const path = require("path");

const APP_URL = process.env.APP_URL || "http://127.0.0.1:4175/";
const FILE = path.join(__dirname, "stress-test-workbench.xlsx");
const ENGINE = process.env.ENGINE === "webkit" ? webkit : chromium;
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
const PROPS = ["will-change", "transform", "filter", "backdrop-filter", "contain", "position", "overflow-x", "overflow-y", "isolation", "opacity", "backface-visibility", "perspective", "mix-blend-mode", "clip-path", "mask-image", "z-index"];

async function run() {
  const browser = await ENGINE.launch({ headless: true });
  const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 390, height: 844 }, deviceScaleFactor: 2, hasTouch: true, isMobile: ENGINE === chromium });
  await context.addInitScript(() => localStorage.setItem("introPlayed", "true"));
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
  await sleep(800);
  await page.evaluate(() => { setSidebarOpen(false); heroUserCollapsed = false; setHeroCollapsed(false); });
  await sleep(600);

  const results = [];
  const check = (name, ok, detail) => results.push({ name, ok: !!ok, detail });
  const snap = () => page.evaluate((props) => {
    const cs = getComputedStyle(document.getElementById("tableWrap"));
    return Object.fromEntries(props.map((p) => [p, cs.getPropertyValue(p)]));
  }, PROPS);

  for (const freeze of [false, true]) {
    await page.evaluate((on) => { const el = document.getElementById("freezeFirstCol"); if (el.checked !== on) { el.checked = on; el.dispatchEvent(new Event("change", { bubbles: true })); } }, freeze);
    await sleep(400);
    for (const on of [true, false]) {
      const label = `${on ? "zwijanie" : "rozwijanie"}${freeze ? " + blokada kolumny" : ""}`;
      const rest = await snap();
      const during = await page.evaluate((on) => { heroUserCollapsed = on; setHeroCollapsed(on); return document.body.classList.contains("hero-animating"); }, on);
      const mid = await snap();
      const changed = PROPS.filter((p) => rest[p] !== mid[p]).map((p) => `${p}: ${rest[p]} → ${mid[p]}`);
      check(`${label}: animacja trwa (warunek testu)`, during);
      check(`${label}: warstwa przewijanej tabeli bez zmian`, changed.length === 0, changed.join("; ") || "bez zmian");
      await sleep(500);
    }
  }
  check("brak błędów strony", errors.length === 0, errors.slice(0, 3).join(" | "));
  await browser.close();

  let failed = 0;
  for (const r of results) {
    console.log(`${r.ok ? "✅" : "❌"} ${r.name}${r.detail ? `  (${r.detail})` : ""}`);
    if (!r.ok) failed += 1;
  }
  if (failed) { console.error(`\n${failed} z ${results.length} nie przeszło`); process.exit(1); }
  console.log(`\n✅ hero-scroll-guard (${process.env.ENGINE || "chromium"}): ${results.length}/${results.length}`);
}

run().catch((e) => { console.error(e); process.exit(1); });
