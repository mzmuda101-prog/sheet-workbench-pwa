// export-playwright.js — test ścieżki EKSPORTU (modal wyboru kolumn → CSV + Druk/PDF).
//
// Symuluje realnego użytkownika: wczytuje arkusz, otwiera „Eksport", ogranicza kolumny
// i sprawdza, że (1) CSV zawiera DOKŁADNIE wybrane kolumny, (2) druk buduje czysty
// #printArea tylko z wybranych kolumn i wszystkich wierszy widoku oraz woła window.print().
//
// Uruchom z serwerem na APP_URL (domyślnie http://127.0.0.1:4175/), np. `npm run serve` obok.

const { chromium } = require("playwright");
const fs = require("fs");
const path = require("path");

const APP_URL = process.env.APP_URL || "http://127.0.0.1:4175/";
const FILE = path.join(__dirname, "stress-test-workbench.xlsx");
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

async function run() {
  const browser = await chromium.launch({ headless: true });
  const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1280, height: 900 }, acceptDownloads: true });
  // introPlayed: pomiń splash; window.print: stub, by headless nie wisiał na dialogu druku.
  await context.addInitScript(() => {
    localStorage.setItem("introPlayed", "true");
    window.print = () => { window.__printed = (window.__printed || 0) + 1; };
  });
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
  await page.evaluate(() => { try { setSidebarOpen && setSidebarOpen(false); } catch {} document.documentElement.classList.remove("sidebar-open"); });
  await sleep(150);

  // 1) Otwórz modal eksportu, sprawdź że są wszystkie kolumny i wszystkie zaznaczone.
  await page.click("#exportCsvBtn");
  await sleep(200);
  const modal = await page.evaluate(() => ({
    visible: !document.getElementById("exportModal").classList.contains("hidden"),
    cols: document.querySelectorAll("#exportColumnList input[type=checkbox]").length,
    allChecked: Array.from(document.querySelectorAll("#exportColumnList input[type=checkbox]")).every((c) => c.checked),
    headers: typeof currentHeaders !== "undefined" ? currentHeaders.length : null,
  }));

  // 2) Zostaw tylko pierwsze 3 kolumny → CSV ma mieć 3 kolumny w nagłówku.
  await page.evaluate(() => {
    [...document.querySelectorAll("#exportColumnList input[type=checkbox]")].forEach((c, i) => { c.checked = i < 3; });
  });
  const dlPromise = page.waitForEvent("download");
  await page.click("#exportCsvAction");
  const dl = await dlPromise;
  const csv = fs.readFileSync(await dl.path(), "utf8");
  const firstLine = csv.split("\n")[0];
  const csvCols = firstLine.split(",").length;

  // 3) Druk z 2 kolumnami → #printArea ma 2 nagłówki i tyle wierszy co widok, print() wywołany.
  await page.click("#exportCsvBtn");
  await sleep(150);
  await page.evaluate(() => {
    [...document.querySelectorAll("#exportColumnList input[type=checkbox]")].forEach((c, i) => { c.checked = i < 2; });
  });
  await page.click("#exportPrintAction");
  await sleep(300);
  const printRes = await page.evaluate(() => ({
    printed: window.__printed || 0,
    ths: document.querySelectorAll("#printArea thead th").length,
    bodyRows: document.querySelectorAll("#printArea tbody tr").length,
    viewRows: typeof viewRows !== "undefined" ? viewRows.length : null,
    modalHidden: document.getElementById("exportModal").classList.contains("hidden"),
  }));

  await browser.close();

  const result = { modal, firstLine, csvCols, printRes, errors };
  const failures = [];
  if (!modal.visible) failures.push("export modal should open");
  if (modal.cols !== modal.headers) failures.push(`column list (${modal.cols}) should match header count (${modal.headers})`);
  if (!modal.allChecked) failures.push("all columns should be checked by default");
  if (csvCols !== 3) failures.push(`CSV should have 3 columns, got ${csvCols} (${firstLine})`);
  if (printRes.ths !== 2) failures.push(`print should have 2 header cells, got ${printRes.ths}`);
  if (printRes.bodyRows !== printRes.viewRows) failures.push(`print rows (${printRes.bodyRows}) should equal view rows (${printRes.viewRows})`);
  if (printRes.printed !== 1) failures.push(`window.print should be called once, got ${printRes.printed}`);
  if (!printRes.modalHidden) failures.push("export modal should close after print");
  if (errors.length) failures.push(`console/page errors: ${errors.join(" | ")}`);

  console.log(JSON.stringify(result, null, 2));
  if (failures.length) throw new Error(failures.join("; "));
  console.log("✅ export-playwright OK");
}

run().catch((error) => {
  console.error(error);
  process.exit(1);
});
