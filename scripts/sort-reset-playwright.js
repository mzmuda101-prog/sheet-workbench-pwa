// sort-reset-playwright.js — RESET SORTOWANIA (2026-10-31).
//   node scripts/sort-reset-playwright.js        ENGINE=webkit node scripts/sort-reset-playwright.js
//
// Pilnuje:
//   1. posortowana tabela → w pasku filtrów chip „sortowanie: Kolumna ↓ ×”;
//   2. × na chipie przywraca kolejność z pliku i chip znika;
//   3. Ctrl/⌘+Alt+R robi to samo; bez sortowania mówi, że nic nie ma do zresetowania;
//   4. skrót jest w ściądze.
// Uruchom z serwerem na APP_URL (domyślnie http://127.0.0.1:4175/).

const playwright = require("playwright");

const APP_URL = process.env.APP_URL || "http://127.0.0.1:4175/";
const ENGINE = process.env.ENGINE || "chromium";
const SLEEP_SCALE = Math.min(1, Math.max(0.1, Number(process.env.SLEEP_SCALE || 1)));
const sleep = (ms) => new Promise((r) => setTimeout(r, Math.round(ms * SLEEP_SCALE)));
const MOD = process.platform === "darwin" && ENGINE === "webkit" ? "Meta" : "Control";

async function run() {
  const browser = await playwright[ENGINE].launch({ headless: true });
  const fails = [];
  const check = (name, ok, detail) => {
    console.log(`${ok ? "✅" : "❌"} ${name}${ok ? "" : ` — ${JSON.stringify(detail)}`}`);
    if (!ok) fails.push(name);
  };
  const ctx = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1280, height: 860 } });
  await ctx.addInitScript(() => localStorage.setItem("introPlayed", "true"));
  const p = await ctx.newPage();
  const errors = [];
  p.on("pageerror", (e) => errors.push(e.message));
  await p.goto(APP_URL, { waitUntil: "load" });
  await p.evaluate(() => { document.getElementById("heroSplash")?.remove(); return loadSampleFile(); });
  await p.waitForFunction(() => document.querySelector("#dataTable tbody tr[data-row-key]"), null, { timeout: 20000 });
  await p.waitForFunction(() => !document.documentElement.classList.contains("vt-scene"), null, { timeout: 10000 });
  await p.evaluate(() => setSidebarOpen(false));
  const order = () => p.evaluate(() => viewRows.slice(0, 6).map((r) => r.rowIndex0));
  const fileOrder = await order();
  const sortBy = async () => {
    await p.evaluate(() => { setPrimarySort(currentHeaders[1], "desc"); applyCurrentSort(); });
    await sleep(500);
  };
  const chip = () => p.evaluate(() => { const c = document.querySelector('#activeFilters .af-chip[data-filter="sort"]'); return c ? c.textContent : null; });

  await sortBy();
  const c1 = await chip();
  check("1. posortowane → chip „sortowanie: … ×” w pasku", !!c1 && /sortowanie: .+↓/.test(c1), c1);
  check("1. (warunek) kolejność się zmieniła", JSON.stringify(await order()) !== JSON.stringify(fileOrder), null);
  await p.click('#activeFilters .af-chip[data-filter="sort"] .af-remove');
  await sleep(500);
  check("2. × przywraca kolejność z pliku", JSON.stringify(await order()) === JSON.stringify(fileOrder), await order());
  check("2. chip znika", (await chip()) === null, await chip());

  await sortBy();
  await p.evaluate(() => document.querySelectorAll(".toast").forEach((x) => x.remove()));
  await p.click("#dataTable tbody tr[data-row-key] td[data-col-index='1']");
  await p.keyboard.press(`${MOD}+Alt+KeyR`);
  await sleep(500);
  check("3. Ctrl/⌘+Alt+R przywraca kolejność z pliku", JSON.stringify(await order()) === JSON.stringify(fileOrder), await order());
  await p.keyboard.press(`${MOD}+Alt+KeyR`);
  await sleep(300);
  const toasts = await p.evaluate(() => [...document.querySelectorAll(".toast")].map((x) => x.textContent).join(" | "));
  check("3. bez sortowania: komunikat, że nie ma czego resetować", /nie jest posortowana/.test(toasts), toasts);

  const sheet = await p.evaluate(() => document.querySelector('[data-shortcut-text="resetSortShortcut"]')?.textContent || "");
  check("4. skrót w ściądze", /Reset sortowania/.test(sheet), sheet);

  check("bez błędów strony", errors.length === 0, errors);
  await browser.close();
  if (fails.length) {
    console.error(`\n${fails.length} błędów (${ENGINE})`);
    process.exit(1);
  }
  console.log(`\nOK (${ENGINE})`);
}

run().catch((e) => { console.error(e); process.exit(1); });
