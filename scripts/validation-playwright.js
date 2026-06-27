// validation-playwright.js — test WALIDACJI listą referencyjną (panel „Walidacja listą").
//
// Symuluje realnego użytkownika kontrolującego jakość danych: wczytuje arkusz, wskazuje
// kolumnę „Status", podaje listę dozwolonych wartości (podzbiór realnych), i sprawdza:
//  (1) podsumowanie + liczba naruszeń zgadza się z ręcznym przeliczeniem,
//  (2) „pokaż tylko niezgodne" zawęża widok DOKŁADNIE do naruszeń (kompozycja z applyFilters),
//  (3) wyłączenie wraca do pełnego widoku,
//  (4) słownik z TEJ SAMEJ kolumny → 0 naruszeń (wartości zawsze w zbiorze).
//
// Uruchom z serwerem na APP_URL (domyślnie http://127.0.0.1:4175/), np. `npm run serve` obok.

const { chromium } = require("playwright");
const path = require("path");

const APP_URL = process.env.APP_URL || "http://127.0.0.1:4175/";
const FILE = path.join(__dirname, "stress-test-workbench.xlsx");
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

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
  await sleep(1000);

  // Wykryj kolumnę Status, wartości i policz oczekiwane naruszenia dla allowed = 2 pierwsze.
  const info = await page.evaluate(() => {
    const idx = currentHeaders.indexOf("Status");
    const counts = new Map();
    baseRows.forEach((r) => { const v = String(getDisplayValue(r, idx) ?? "").trim().toLowerCase(); counts.set(v, (counts.get(v) || 0) + 1); });
    const distinct = [...counts.keys()].filter(Boolean);
    const allowed = distinct.slice(0, 2);
    let expViol = 0;
    baseRows.forEach((r) => { const v = String(getDisplayValue(r, idx) ?? "").trim().toLowerCase(); if (v === "") return; if (!allowed.includes(v)) expViol++; });
    return { idx, distinct, allowed, total: baseRows.length, expViol };
  });

  // Ustaw kolumnę + listę dozwolonych i sprawdź (przycisk w sidebarze → klik przez evaluate).
  await page.evaluate((info) => {
    document.getElementById("panel-validation").open = true;
    document.getElementById("validationColumn").value = String(info.idx);
    const src = document.getElementById("validationSource");
    src.value = "list";
    src.dispatchEvent(new Event("change"));
    document.getElementById("validationAllowed").value = info.allowed.join("\n");
    document.getElementById("validationCheckBtn").click();
  }, info);
  await sleep(300);
  const afterCheck = await page.evaluate(() => ({
    summary: document.getElementById("validationSummary").textContent,
    summaryBad: document.getElementById("validationSummary").classList.contains("bad"),
    badRows: document.querySelectorAll("#validationResults .validation-bad-row").length,
    stateColIdx: validationState.colIdx,
    allowedSize: validationState.allowed ? validationState.allowed.size : null,
    showOnlyWrapHidden: document.getElementById("validationShowOnlyWrap").classList.contains("hidden"),
  }));

  // „Pokaż tylko niezgodne" → widok = liczba naruszeń.
  await page.evaluate(() => { const c = document.getElementById("validationShowOnly"); c.checked = true; c.dispatchEvent(new Event("change")); });
  await sleep(400);
  const showOnlyRows = await page.evaluate(() => viewRows.length);

  // Wyłącz → pełny widok.
  await page.evaluate(() => { const c = document.getElementById("validationShowOnly"); c.checked = false; c.dispatchEvent(new Event("change")); });
  await sleep(300);
  const offRows = await page.evaluate(() => viewRows.length);

  // Słownik z tej samej kolumny → 0 naruszeń.
  await page.evaluate((idx) => {
    const src = document.getElementById("validationSource");
    src.value = "column";
    src.dispatchEvent(new Event("change"));
    document.getElementById("validationDictColumn").value = String(idx);
    document.getElementById("validationCheckBtn").click();
  }, info.idx);
  await sleep(300);
  const dictCheck = await page.evaluate(() => ({
    ok: document.getElementById("validationSummary").classList.contains("ok"),
    showOnly: validationState.showOnly,
  }));

  await browser.close();

  const result = { info, afterCheck, showOnlyRows, offRows, dictCheck, errors };
  const failures = [];
  if (info.idx < 0) failures.push("Status column not found in stress workbook");
  if (afterCheck.stateColIdx !== info.idx) failures.push("validationState.colIdx not set to Status column");
  if (afterCheck.allowedSize !== info.allowed.length) failures.push(`allowed set size ${afterCheck.allowedSize} != ${info.allowed.length}`);
  if (!afterCheck.summaryBad) failures.push("summary should be marked 'bad' when violations exist");
  if (afterCheck.showOnlyWrapHidden) failures.push("'show only non-matching' toggle should be revealed");
  if (showOnlyRows !== info.expViol) failures.push(`show-only rows ${showOnlyRows} != expected violations ${info.expViol}`);
  if (offRows !== info.total) failures.push(`view should restore to ${info.total} rows, got ${offRows}`);
  if (!dictCheck.ok) failures.push("self-dictionary validation should report all valid (0 violations)");
  if (errors.length) failures.push(`console/page errors: ${errors.join(" | ")}`);

  console.log(JSON.stringify(result, null, 2));
  if (failures.length) throw new Error(failures.join("; "));
  console.log("✅ validation-playwright OK");
}

run().catch((error) => {
  console.error(error);
  process.exit(1);
});
