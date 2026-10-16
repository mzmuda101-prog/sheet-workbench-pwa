// report-summary-playwright.js — „Podsumowanie” na górze raportu (report.js).
//
// Co musi być prawdą:
//   1. bez AI: akapit z zakresu i pierwszych wniosków — te same liczby co we wnioskach,
//      plus liczba uwag do sprawdzenia,
//   2. AI (atrapa) z liczbami z wniosków → akapit AI + dopisek, że liczby sprawdzone;
//      „Wróć do wersji z reguł” przywraca akapit z reguł,
//   3. AI z wymyśloną liczbą → odrzucone, zostaje wersja z reguł,
//   4. przyciski podsumowania nie trafiają do PDF ani na wydruk (rp-screen-only).
//
// Uruchom z serwerem na APP_URL (domyślnie http://127.0.0.1:4175/).

const { chromium, webkit } = require("playwright");

const APP_URL = process.env.APP_URL || "http://127.0.0.1:4175/";
const ENGINE = process.env.ENGINE === "webkit" ? webkit : chromium;
const results = [];
const check = (name, ok, detail = "") => results.push({ name, ok: !!ok, detail });
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

const ROWS = [
  [1, "W toku", 100], [2, "W toku", 120], [3, "Zakończone", 90], [4, "W toku", 110], [5, "Zakończone", 95],
  [6, "W toku", 105], [7, "Anulowane", 1000], [7, "W toku", 115], [9, "Zakończone", 100], [10, "W toku", 125],
];

async function run() {
  const browser = await ENGINE.launch({ headless: true });
  const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1280, height: 900 } });
  await context.addInitScript(() => localStorage.setItem("introPlayed", "true"));
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push("pageerror: " + e.message));
  page.on("console", (m) => { if (m.type() === "error") errors.push("console: " + m.text()); });
  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  await page.evaluate(() => ensureXlsxLibs(false));
  await page.waitForFunction(() => typeof buildRows === "function" && window.__report);
  await page.evaluate((rows) => {
    const ws = {};
    ["Nr", "Status", "Kwota"].forEach((h, i) => { ws[String.fromCharCode(65 + i) + "1"] = { t: "s", v: h }; });
    rows.forEach(([nr, st, kw], i) => {
      const r = i + 2;
      ws["A" + r] = { t: "n", v: nr, w: String(nr) };
      ws["B" + r] = { t: "s", v: st };
      ws["C" + r] = { t: "n", v: kw, w: String(kw) };
    });
    ws["!ref"] = "A1:C" + (rows.length + 1);
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
  await sleep(400);
  await page.evaluate(() => window.__report.open());
  await page.waitForFunction(() => !!document.querySelector("#rpPage .rp-summary"), null, { timeout: 15000 });

  const rule = await page.evaluate(() => ({
    text: document.querySelector("#rpPage .rp-summary-text")?.textContent || "",
    first: document.querySelector("#rpPage [data-section]")?.dataset.section,
    findings: [...document.querySelectorAll("#rpPage .rp-findings li")].map((li) => li.textContent),
    btn: document.querySelector("#rpPage .rp-summary-ai")?.textContent,
  }));
  check("podsumowanie to pierwsza sekcja kartki", rule.first === "summary", rule.first);
  check("akapit z reguł: zakres (10 wierszy) + pierwszy wniosek", /^Raport obejmuje 10 wierszy\./.test(rule.text) && rule.findings[0] && rule.text.includes(rule.findings[0].replace(/[.]$/, "")), JSON.stringify(rule));
  check("liczba uwag do sprawdzenia, gdy wnioski mają ostrzeżenia", /Do sprawdzenia: \d+/.test(rule.text) === rule.findings.length > 0, rule.text);
  check("przycisk „✦ Napisz płynniej (AI)” na ekranie", rule.btn === "✦ Napisz płynniej (AI)", rule.btn);

  // AI z liczbami z wniosków
  const good = await page.evaluate(async () => {
    const facts = rpSummaryFacts(rpData);
    const nums = (facts.main.join(" ").match(/\d[\d\s ,]*\d|\d/g) || []).slice(0, 2);
    window.__localAiMock = () => `Zestawienie obejmuje 10 wierszy; najważniejsze liczby to ${nums.join(" i ")}.`;
    await rpSummaryWriteAi();
    const out = {
      text: document.querySelector("#rpPage .rp-summary-text")?.textContent,
      note: document.querySelector("#rpPage .rp-summary-note")?.textContent || "",
      btn: document.querySelector("#rpPage .rp-summary-ai")?.textContent,
    };
    document.querySelector("#rpPage .rp-summary-ai").click();
    out.back = document.querySelector("#rpPage .rp-summary-text")?.textContent;
    return out;
  });
  check("AI z liczbami z wniosków → akapit AI + dopisek o sprawdzeniu", /^Zestawienie obejmuje 10 wierszy/.test(good.text || "") && /liczby sprawdzone/.test(good.note), JSON.stringify(good));
  check("„Wróć do wersji z reguł” przywraca akapit z reguł", good.btn === "Wróć do wersji z reguł" && /^Raport obejmuje 10 wierszy/.test(good.back || ""), JSON.stringify(good));

  const bad = await page.evaluate(async () => {
    window.__localAiMock = () => "Obroty wzrosły o 37% do 99 999 zł.";
    await rpSummaryWriteAi();
    delete window.__localAiMock;
    return { text: document.querySelector("#rpPage .rp-summary-text")?.textContent, note: !!document.querySelector("#rpPage .rp-summary-note") };
  });
  check("AI z wymyśloną liczbą → odrzucone, zostaje wersja z reguł", /^Raport obejmuje 10 wierszy/.test(bad.text || "") && !bad.note, JSON.stringify(bad));

  const printCss = await page.evaluate(() => {
    const el = document.querySelector("#rpPage .rp-summary-actions");
    return { cls: el?.classList.contains("rp-screen-only") };
  });
  await page.emulateMedia({ media: "print" });
  const printHidden = await page.evaluate(() => getComputedStyle(document.querySelector("#rpPage .rp-summary-actions")).display === "none");
  await page.emulateMedia({ media: "screen" });
  check("przyciski tylko na ekranie (ukryte na wydruku, pomijane w PDF)", printCss.cls && printHidden, JSON.stringify({ printCss, printHidden }));

  check("brak błędów strony", !errors.length, errors.join(" | "));
  await browser.close();
  let failed = 0;
  results.forEach((r) => { console.log(`${r.ok ? "✅" : "❌"} ${r.name}${r.ok || !r.detail ? "" : ` (${String(r.detail).slice(0, 500)})`}`); if (!r.ok) failed++; });
  if (failed) process.exit(1);
  console.log(`\n✅ report-summary: ${results.length}/${results.length} (${process.env.ENGINE || "chromium"})`);
}
run().catch((e) => { console.error(e); process.exit(1); });
