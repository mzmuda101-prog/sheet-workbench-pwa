// aggregation-search-ops-playwright.js — operatory wyszukiwania w SZUKAJCE WYNIKÓW agregacji.
//
// Co musi być prawdą:
//   1. z WYŁĄCZONYMI operatorami pole działa dokładnie jak wcześniej (zwykłe „zawiera"),
//      także dla fraz, które wyglądają jak operatory — to jest gwarancja „nic nie psujemy",
//   2. z WŁĄCZONYMI działa || (lub), && (i), ! (bez) i {} (grupowanie) — na etykietach grup,
//   3. porównania >> / << filtrują po WARTOŚCI miary, a nie po tekście etykiety,
//   4. tekst nigdy nie wpada w kolumnę wartości (szukanie „1" nie łapie grup po ich liczbach).
//
// Uruchom z serwerem na APP_URL (domyślnie http://127.0.0.1:4175/).

const { chromium } = require("playwright");
const path = require("path");

const APP_URL = process.env.APP_URL || "http://127.0.0.1:4175/";
const FILE = path.join(__dirname, "stress-test-workbench.xlsx");
// SLEEP_SCALE (run-tests.mjs): skraca przerwy „odczekaj po kroku” w szybkim przebiegu;
// powtórka porażki idzie ze skalą 1 (pełne przerwy). sleepFixed = przerwa, która MUSI
// przeczekać timer aplikacji (debounce, opóźnienie podpowiedzi, bezczynność) — bez skali.
const SLEEP_SCALE = Math.min(1, Math.max(0.1, Number(process.env.SLEEP_SCALE || 1)));
const sleep = (ms) => new Promise((r) => setTimeout(r, Math.round(ms * SLEEP_SCALE)));
const sleepFixed = (ms) => new Promise((r) => setTimeout(r, ms)); // eslint-disable-line no-unused-vars

async function run() {
  const browser = await chromium.launch({ headless: true });
  const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1280, height: 900 } });
  await context.addInitScript(() => localStorage.setItem("introPlayed", "true"));
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push("pageerror: " + e.message));
  page.on("console", (m) => { if (m.type() === "error") errors.push("console: " + m.text()); });

  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  await page.evaluate(() => { try { ensureXlsxLibs && ensureXlsxLibs(false); } catch {} });
  await page.setInputFiles("#fileInput", FILE);
  await page.waitForFunction(() => document.getElementById("sheetSelect")?.options?.length > 0, null, { timeout: 15000 });
  await sleep(300);
  await page.click("#loadBtn");
  await sleep(1200);

  // Panel agregacji: grupujemy po „Osoba", miara = liczba wierszy.
  await page.evaluate(() => { document.getElementById("panel-aggregation-workbench").open = true; });
  // Fire-and-forget: NIE zwracamy promisy do Playwrighta — dogrywanie modułu potrafi
  // przeżyć re-render strony, a wtedy await na zniszczonym kontekście wywala test.
  await page.evaluate(() => { if (typeof ensureAnalysisHeavy === "function") ensureAnalysisHeavy(); });
  await page.waitForFunction(() => typeof buildAggregationSearchMatcher === "function", null, { timeout: 15000 });

  const setup = await page.evaluate(() => {
    aggregationWorkbenchState.groupBy = "Osoba";
    aggregationWorkbenchState.measures = ["count_rows"];
    aggregationWorkbenchState.aggregation = "count";
    aggregationWorkbenchState.showCount = 999;
    aggregationWorkbenchState.resultSearch = "";
    aggregationWorkbenchState.resultSearchOperators = false;
    renderAggregationWorkbench();
    return true;
  });
  await sleep(600);

  const search = async (query, ops) => page.evaluate(({ query, ops }) => {
    aggregationWorkbenchState.resultSearch = query;
    aggregationWorkbenchState.resultSearchOperators = ops;
    renderAggregationWorkbench();
    const labels = Array.from(document.querySelectorAll("#aggregationWorkbenchList .duration-person-title")).map((el) => el.textContent);
    const values = Array.from(document.querySelectorAll("#aggregationWorkbenchList .duration-person-value")).map((el) => el.textContent);
    return { labels, values };
  }, { query, ops });

  const failures = [];
  const report = {};
  const ok = (name, cond, got) => { if (!cond) failures.push(`${name} — dostalem: ${JSON.stringify(got)}`); };

  const all = await search("", false);
  report.total = all.labels.length;
  ok("sa wyniki do filtrowania", all.labels.length >= 4, all.labels.length);
  const sample = all.labels[0] || "";
  const firstWord = sample.split(" ")[0];
  const second = all.labels.find((l) => !l.startsWith(firstWord)) || "";
  const secondWord = second.split(" ")[0];

  // 1. Operatory WYŁĄCZONE — stare zachowanie, także dla tekstu z „||"
  const plain = await search(firstWord, false);
  report.plain = { q: firstWord, n: plain.labels.length };
  ok("wylaczone: zwykle zawiera dziala", plain.labels.length > 0 && plain.labels.every((l) => l.toLowerCase().includes(firstWord.toLowerCase())), report.plain);

  const plainOps = await search(`${firstWord} || ${secondWord}`, false);
  report.plainOpsLiteral = { q: `${firstWord} || ${secondWord}`, n: plainOps.labels.length };
  ok("wylaczone: '||' traktowane doslownie (0 trafien)", plainOps.labels.length === 0, report.plainOpsLiteral);

  // 2. Operatory WŁĄCZONE — || i && i !
  const orRes = await search(`${firstWord} || ${secondWord}`, true);
  report.or = { q: `${firstWord} || ${secondWord}`, n: orRes.labels.length };
  ok("wlaczone: || laczy oba warunki", orRes.labels.length >= plain.labels.length + 1, report.or);
  ok("wlaczone: || nie wpuszcza obcych", orRes.labels.every((l) => l.toLowerCase().includes(firstWord.toLowerCase()) || l.toLowerCase().includes(secondWord.toLowerCase())), orRes.labels);

  const notRes = await search(`!${firstWord}`, true);
  report.not = { n: notRes.labels.length };
  ok("wlaczone: ! wyklucza", notRes.labels.length === all.labels.length - plain.labels.length && notRes.labels.every((l) => !l.toLowerCase().includes(firstWord.toLowerCase())), report.not);

  const bracket = await search(`{${firstWord} || ${secondWord}}`, true);
  report.bracket = { n: bracket.labels.length };
  ok("wlaczone: {} grupuje", bracket.labels.length === orRes.labels.length, report.bracket);

  // 3. Porównania na WARTOŚCI miary
  const nums = all.values.map((v) => parseFloat(String(v).replace(/[^\d.,-]/g, "").replace(",", "."))).filter((n) => Number.isFinite(n));
  const median = nums.slice().sort((a, b) => a - b)[Math.floor(nums.length / 2)];
  const gt = await search(`>>${median}`, true);
  const expectedGt = nums.filter((n) => n > median).length;
  report.cmp = { median, got: gt.labels.length, expected: expectedGt };
  ok("wlaczone: >> filtruje po wartosci miary", gt.labels.length === expectedGt, report.cmp);

  const lte = await search(`=<<${median}`, true);
  report.cmpLte = { got: lte.labels.length, expected: nums.filter((n) => n <= median).length };
  ok("wlaczone: =<< dziala jako <=", lte.labels.length === report.cmpLte.expected, report.cmpLte);

  // 4. Tekst nie wpada w kolumnę wartości
  const textOnly = await search(String(median), true);
  report.textNotValue = { q: String(median), n: textOnly.labels.length, byLabel: all.labels.filter((l) => l.includes(String(median))).length };
  ok("wlaczone: tekst szuka tylko w etykiecie", textOnly.labels.length === report.textNotValue.byLabel, report.textNotValue);
  ok("wlaczone: liczba-wartosc nie lapie grupy po wartosci", report.textNotValue.byLabel > 0 || textOnly.labels.length === 0, report.textNotValue);

  console.log(JSON.stringify(report, null, 2));
  if (errors.length) failures.push("bledy strony: " + errors.join(" | "));
  await browser.close();
  if (failures.length) {
    console.error("\n❌ aggregation-search-ops-playwright FAIL:\n" + failures.map((f) => " - " + f).join("\n"));
    process.exit(1);
  }
  console.log("✅ aggregation-search-ops-playwright OK");
}

run().catch((e) => { console.error(e); process.exit(1); });
