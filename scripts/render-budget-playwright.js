// render-budget-playwright.js — strażnik automatycznego limitu wierszy (app/render-budget.js).
//
// Pomiar kosztu renderu podstawiamy w localStorage (deterministycznie, bez dławienia CPU),
// a sprawdzamy reguły, które NIE mogą się zepsuć:
//   1. bez pomiaru = 200 (jak przed automatem),
//   2. wolne urządzenie nigdy nie schodzi poniżej 200,
//   3. szybkie dostaje więcej, ale w granicach sufitu komórek,
//   4. limit wpisany ręcznie wygrywa z automatem; puste pole wraca do automatu,
//   5. ?rowbudget=off wyłącza automat,
//   6. zwykły render zapisuje próbkę pomiaru.
//
// Uruchom z serwerem na APP_URL (domyślnie http://127.0.0.1:4175/).

const { chromium } = require("playwright");

const APP_URL = process.env.APP_URL || "http://127.0.0.1:4175/";
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

async function openWith(browser, { cost, manual, query = "" }) {
  const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1100, height: 850 } });
  await context.addInitScript(({ cost, manual }) => {
    if (sessionStorage.getItem("__rbInit")) return; // tylko przy pierwszym wejściu
    sessionStorage.setItem("__rbInit", "1");
    localStorage.setItem("introPlayed", "true");
    if (cost) localStorage.setItem("swb-render-cost-v1", JSON.stringify({ c: cost, n: 10 }));
    if (manual) localStorage.setItem("excel-workbench-max-rows", String(manual));
  }, { cost, manual });
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  await page.goto(APP_URL + "?sample=3000" + query, { waitUntil: "load" });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  await page.evaluate(() => document.getElementById("loadSampleBtn")?.click());
  await page.waitForFunction(() => document.querySelectorAll("#dataTable tbody tr").length > 0, null, { timeout: 60000 });
  await sleep(600);
  const info = await page.evaluate(() => ({
    shown: document.querySelectorAll("#dataTable tbody tr").length,
    field: maxRowsEl.value,
    cols: currentHeaders.length,
  }));
  return { context, page, info, errors };
}

async function run() {
  const browser = await chromium.launch({ headless: true });
  const results = [];
  const check = (name, ok, detail = "") => results.push({ name, ok, detail });

  // 1. świeże urządzenie
  let s = await openWith(browser, {});
  check("bez pomiaru → 200 wierszy", s.info.shown === 200, JSON.stringify(s.info));
  const stored = await s.page.evaluate(() => localStorage.getItem("swb-render-cost-v1"));
  check("render zapisał próbkę pomiaru", !!stored && JSON.parse(stored).c > 0, String(stored));
  check("brak błędów strony (świeże)", s.errors.length === 0, s.errors.join("; "));
  await s.context.close();

  // 2. bardzo wolne urządzenie (10 ms/100 komórek ≈ CPU ×6+)
  s = await openWith(browser, { cost: 0.1 });
  check("wolne urządzenie → nie mniej niż 200", s.info.shown === 200, JSON.stringify(s.info));
  await s.context.close();

  // 3. szybkie urządzenie (0,5 ms/100 komórek)
  s = await openWith(browser, { cost: 0.005 });
  const cap = Math.floor(24000 / (s.info.cols + 1));
  check("szybkie urządzenie → więcej niż 200", s.info.shown > 200, JSON.stringify(s.info));
  check("szybkie urządzenie → w granicach sufitu komórek", s.info.shown <= cap && s.info.shown <= 2000, `cap ${cap}, ${JSON.stringify(s.info)}`);
  check("pole limitu pokazuje wartość automatu", s.info.field === String(s.info.shown), JSON.stringify(s.info));

  // 4b. wyczyszczenie ręcznego limitu → powrót do automatu (na tej samej karcie)
  await s.page.evaluate(() => { maxRowsEl.value = "50"; maxRowsEl.dispatchEvent(new Event("change")); });
  await sleep(400);
  const manualShown = await s.page.evaluate(() => document.querySelectorAll("#dataTable tbody tr").length);
  check("ręczne 50 działa od razu", manualShown === 50, String(manualShown));
  await s.page.evaluate(() => { maxRowsEl.value = ""; maxRowsEl.dispatchEvent(new Event("change")); });
  await sleep(400);
  const back = await s.page.evaluate(() => ({
    saved: localStorage.getItem("excel-workbench-max-rows"),
    field: maxRowsEl.value,
    shown: document.querySelectorAll("#dataTable tbody tr").length,
  }));
  check("puste pole → automat (brak zapisu, wartość > 200)", back.saved === null && Number(back.field) > 200 && back.shown === Number(back.field), JSON.stringify(back));
  check("brak błędów strony (szybkie)", s.errors.length === 0, s.errors.join("; "));
  await s.context.close();

  // 4. ręczny limit wygrywa nawet na szybkim urządzeniu
  s = await openWith(browser, { cost: 0.005, manual: 120 });
  check("ręczny limit 120 wygrywa z automatem", s.info.shown === 120 && s.info.field === "120", JSON.stringify(s.info));
  await s.context.close();

  // 5. ?rowbudget=off
  s = await openWith(browser, { cost: 0.005, query: "&rowbudget=off" });
  check("?rowbudget=off → 200 jak dawniej", s.info.shown === 200, JSON.stringify(s.info));
  await s.context.close();

  await browser.close();
  let failed = 0;
  for (const r of results) {
    if (!r.ok) failed += 1;
    console.log(`${r.ok ? "✅" : "❌"} ${r.name}${r.ok ? "" : "  → " + r.detail}`);
  }
  console.log(failed ? `\n${failed} z ${results.length} NIE przeszło` : `\nWszystkie ${results.length} przeszły`);
  process.exit(failed ? 1 : 0);
}

run().catch((e) => { console.error(e); process.exit(1); });
