/**
 * test-runner.mjs — Playwright smoke-test dla Excel Workbench PWA
 *
 * Użycie:
 *   node scripts/test-runner.mjs [ścieżka_do_pliku.xlsx] [port]
 *   npm run test                     # skrót (domyślny plik + port 7821)
 *
 * Przykłady:
 *   npm run test                             # stress-test-workbench.xlsx
 *   node scripts/test-runner.mjs ~/moj.xlsx  # Twój własny plik
 *   node scripts/test-runner.mjs ~/moj.xlsx 8080.   # ścieżka do pliku + port
 *                                ^^ nalezy podac pelną ścieżką do pliku
 * Wymaga (jednorazowo):
 *   npm install
 *   npx playwright install chromium
 *
 * Serwer HTTP (osobny terminal przed testem):
 *   npm run serve
 * 
 * DOstępne komendy:
 *   npm run test:stress:file ~/moj.xlsx  # testuje jedną z plików w folderze
 *   npm run test:stress:file ~/moj.xlsx 8080  # ścieżka do pliku + port
 *   npm run test:stress:file ~/moj.xlsx 8080 ~/moj2.xlsx  # ścieżka do pliku + port
 *   npm run test:stress:file ~/moj.xlsx 8080 ~/moj2.xlsx 8081  # ścieżka do pliku + port
 * 
## npm run serve              # serwer lokalny na :7821
## npm run test               # testuje domyślny plik
## npm run test:stress        # testuje domyślny plik
## node scripts/test-runner.mjs ~/twoj-plik.xlsx  # własny plik
 */

import { chromium } from "playwright";
import { resolve, dirname } from "path";
import { existsSync } from "fs";
import { fileURLToPath } from "url";

const __dir = dirname(fileURLToPath(import.meta.url));
const XLSX_ARG = process.argv[2];
const PORT = process.argv[3] || "7821";
const BASE = `http://localhost:${PORT}`;

const XLSX_PATH = XLSX_ARG
  ? resolve(XLSX_ARG)
  : resolve(__dir, "stress-test-workbench.xlsx");

if (!existsSync(XLSX_PATH)) {
  console.error(`✗ Plik nie istnieje: ${XLSX_PATH}`);
  process.exit(1);
}

const log  = (tag, msg) => console.log(`\n[${tag}] ${msg}`);
const ok   = (msg) => console.log(`  ✓ ${msg}`);
const err  = (msg) => console.log(`  ✗ ${msg}`);
const info = (msg) => console.log(`  → ${msg}`);

console.log(`\nExcel Workbench PWA — test runner`);
console.log(`Plik:   ${XLSX_PATH}`);
console.log(`Serwer: ${BASE}\n`);

const browser = await chromium.launch({ headless: true });
const ctx = await browser.newContext();
const page = await ctx.newPage();

const jsErrors = [];
page.on("console", m => { if (m.type() === "error") jsErrors.push(m.text()); });
page.on("pageerror", e => jsErrors.push(e.message));

// ── boot ─────────────────────────────────────────────────────────────────────

await page.goto(BASE, { waitUntil: "networkidle" });
await page.waitForTimeout(1200);
log("BOOT", "Strona załadowana");

// ── wgranie pliku ─────────────────────────────────────────────────────────────

await page.locator("#fileInput").setInputFiles(XLSX_PATH);
log("FILE", `Wgrywam: ${XLSX_PATH.split("/").pop()}`);

await page.waitForFunction(() =>
  document.getElementById("loadingOverlay")?.classList.contains("hidden"),
  { timeout: 20000 }
);
ok("Workbook sparsowany");

await page.click("#loadBtn", { force: true });
await page.waitForFunction(() =>
  document.getElementById("loadingOverlay")?.classList.contains("hidden"),
  { timeout: 20000 }
);
ok("Arkusz załadowany do tabeli");

// ── helpers ───────────────────────────────────────────────────────────────────

const waitLoaded = () => page.waitForFunction(() =>
  document.getElementById("loadingOverlay")?.classList.contains("hidden"),
  { timeout: 15000 }
);

async function loadSheet(name) {
  await page.evaluate((n) => {
    const sel = document.getElementById("sheetSelect");
    if (sel) sel.value = n;
  }, name);
  await page.evaluate(() => document.getElementById("loadBtn")?.click());
  // Poczekaj aż overlay się pojawi (max 1s), potem aż zniknie
  await page.waitForTimeout(300);
  await waitLoaded();
  await page.waitForTimeout(500);
}

async function ensurePanel(id) {
  await page.evaluate((panelId) => {
    const panel = document.getElementById(panelId);
    if (panel && !panel.open) panel.open = true;
  }, id);
  await page.waitForTimeout(200);
}

async function getGroupByOptions() {
  // Czytaj kolumny z nagłówków tabeli (bez kolumny przewodniej A/B/C)
  return page.evaluate(() =>
    Array.from(document.querySelectorAll("#dataTable thead tr:not(.guide-row) th"))
      .map(th => th.textContent.trim())
      .filter(t => t && !t.match(/^[A-Z]+$/) && !t.includes("->"))
  );
}

async function setGroupBy(value) {
  // Kliknij picker button, zaznacz kolumnę, zatwierdź
  const pickBtn = page.locator("[data-aggregation-control='groupby-pick']");
  if (!await pickBtn.count()) return;
  await pickBtn.evaluate(el => el.click());
  await page.waitForTimeout(300);

  // Odznacz wszystkie, zaznacz tylko żądaną
  await page.evaluate((val) => {
    const picker = document.getElementById("columnPicker");
    if (!picker) return;
    picker.querySelectorAll("input[type=checkbox]").forEach(cb => {
      cb.checked = cb.dataset.value === val || cb.value === val ||
        cb.closest("label")?.textContent?.trim() === val;
    });
  }, value);

  // Kliknij Zastosuj
  const applyBtn = page.locator("#applyPickBtn");
  if (await applyBtn.count()) await applyBtn.evaluate(el => el.click());
  await page.waitForTimeout(600);
}

async function aggItemCount() {
  return page.evaluate(() => document.querySelectorAll(".aggregation-item").length);
}

async function aggSummaryText() {
  return page.evaluate(() =>
    document.getElementById("aggregationWorkbenchSummary")?.textContent?.slice(0, 120) || ""
  );
}

async function getFirstBucketValue() {
  return page.evaluate(() => {
    const item = document.querySelector(".aggregation-item");
    if (!item) return null;
    const nums = item.textContent.match(/\d+/g);
    return nums ? nums[nums.length - 1] : null;
  });
}

// ── listy arkuszy ─────────────────────────────────────────────────────────────

const sheetNames = await page.evaluate(() =>
  Array.from(document.querySelectorAll("#sheetSelect option")).map(o => o.value)
);
info(`Arkusze w pliku: ${sheetNames.join(", ")}`);

// ── testy każdego arkusza ─────────────────────────────────────────────────────

for (const sheet of sheetNames) {
  log(`ARK: ${sheet}`, "");
  await loadSheet(sheet);

  const rowCount = await page.evaluate(() =>
    document.querySelectorAll("#dataTable tbody tr").length
  );
  info(`Wierszy widocznych w tabeli: ${rowCount}`);
  if (rowCount > 0) ok(`Tabela wyrenderowana (${rowCount} wierszy)`);
  else err("Tabela pusta!");

  // Wide-to-Long
  const wideLongVisible = await page.locator("#wideLongToggle").isVisible();
  if (wideLongVisible) {
    ok("Wide-to-Long dostępny — powtarzające się bloki wykryte");

    await page.locator("#wideLongToggle").click();
    await page.waitForTimeout(400);
    const longRows = await page.evaluate(() =>
      document.querySelectorAll("#dataTable tbody tr").length
    );
    const longHdrs = await page.evaluate(() =>
      Array.from(document.querySelectorAll("#dataTable thead th:not(.guide-row th)"))
        .map(th => th.textContent.trim()).filter(Boolean).slice(0, 8)
    );
    info(`Long view: ${longRows} wierszy (limit maxRows) | nagłówki: ${longHdrs.join(" | ")}`);

    const hasBlockCol = longHdrs.some(h => /block|blok|nr/i.test(h));
    if (hasBlockCol) ok("Kolumny bloku (Nr bloku / Block #) obecne w Long view");
    else info(`Kolumny bloku niewidoczne w top-8 nagłówków: ${longHdrs.join(", ")}`);

    // Sprawdź licznik wierszy w modelu (nie w DOM — DOM jest ograniczony maxRows)
    const longModelRows = await page.evaluate(() =>
      typeof window.currentDisplayModel !== "undefined" && window.currentDisplayModel?.rows?.length
    );
    if (longModelRows && longModelRows > rowCount) ok(`Long model ma ${longModelRows} wierszy (>${rowCount} wide) — mnożnik działa`);
    else if (longModelRows) info(`Long model: ${longModelRows} wierszy`);

    // Wróć do wide
    await page.locator("#wideLongToggle").click();
    await page.waitForTimeout(300);
  } else {
    info("Wide-to-Long niedostępny — brak wykrytych bloków");
  }

  // Agregacja
  await ensurePanel("panel-aggregation-workbench");
  const groupOpts = await getGroupByOptions();
  info(`Opcje GroupBy (${groupOpts.length}): ${groupOpts.slice(0, 6).join(", ")}${groupOpts.length > 6 ? "..." : ""}`);

  if (groupOpts.length > 0) {
    const firstCol = groupOpts[0];
    await setGroupBy(firstCol);
    const groups = await aggItemCount();
    const summary = await aggSummaryText();
    info(`GroupBy="${firstCol}": ${groups} grup | ${summary.slice(0, 80)}`);
    if (groups > 0) ok(`Agregacja działa — ${groups} grup`);
    else err("Brak wyników agregacji!");

    // Scope test (bez filtrów: filtered = all)
    await page.evaluate(() => {
      const sel = document.querySelector("[data-aggregation-control='scope']");
      if (sel) { sel.value = "filtered"; sel.dispatchEvent(new Event("change", { bubbles: true })); }
    });
    await page.waitForTimeout(400);
    const valFiltered = await getFirstBucketValue();

    await page.evaluate(() => {
      const sel = document.querySelector("[data-aggregation-control='scope']");
      if (sel) { sel.value = "all"; sel.dispatchEvent(new Event("change", { bubbles: true })); }
    });
    await page.waitForTimeout(400);
    const valAll = await getFirstBucketValue();

    if (valFiltered !== null && valAll !== null) {
      if (valFiltered === valAll) ok(`Scope test OK: aktualny widok (${valFiltered}) = cały arkusz (${valAll})`);
      else err(`Scope bez filtrów: aktualny widok=${valFiltered} ≠ cały arkusz=${valAll} — potencjalny bug!`);
    }
  }

  // Wydajność scroll
  const tS = Date.now();
  await page.locator("#tableWrap").evaluate(el => { el.scrollTop = 99999; el.scrollLeft = 99999; });
  await page.waitForTimeout(80);
  await page.locator("#tableWrap").evaluate(el => { el.scrollTop = 0; el.scrollLeft = 0; });
  ok(`Scroll: ${Date.now() - tS}ms`);
}

// ── podsumowanie błędów JS ────────────────────────────────────────────────────

log("DONE", "═══════════════════════════════════");
if (jsErrors.length === 0) {
  ok("Brak błędów JS w konsoli przeglądarki");
} else {
  err(`Błędy JS (${jsErrors.length}):`);
  jsErrors.slice(0, 10).forEach(e => console.log(`    • ${e.slice(0, 150)}`));
}

await browser.close();
