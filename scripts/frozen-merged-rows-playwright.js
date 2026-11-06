// frozen-merged-rows-playwright.js — zablokowane kolumny w wierszach ze scaleniami (lista 3.11).
//
// Co musi być prawdą: przy blokadzie N kolumn KAŻDA komórka stoi dokładnie pod literą swojej
// kolumny (rząd liter), także po przewinięciu w bok — w wierszu pod scaleniem pionowym
// (brak komórki A, bo przykrywa ją rowspan z wiersza wyżej), w wierszu ze scaleniem poziomym
// (B:C = jedna komórka) i w nagłówku ze scaleniem. Dawniej przyklejanie szło wg nth-child
// (pozycji komórki w wierszu), więc w takich wierszach przyklejała się zła kolumna.
// Oba silniki tabeli (VIRT=0 stary, domyślnie nowy).

const { chromium, webkit } = require("playwright");
const fs = require("fs");
const os = require("os");
const path = require("path");
const XLSX = require(path.join(__dirname, "..", "lib", "xlsx.full.min.js"));

const APP_URL = process.env.APP_URL || "http://127.0.0.1:4175/";
const ENGINE = process.env.ENGINE === "webkit" ? webkit : chromium;
const VIRT = process.env.VIRT === "0" ? "0" : null;

function fixture() {
  const cols = 14;
  const rows = [Array.from({ length: cols }, (_, c) => (c === 0 ? "Grupa" : `Kolumna ${c + 1} z długą nazwą`))];
  for (let r = 1; r <= 12; r++) rows.push(Array.from({ length: cols }, (_, c) => `w${r + 1} k${c + 1} tekst`));
  const ws = XLSX.utils.aoa_to_sheet(rows);
  ws["!merges"] = [
    { s: { r: 2, c: 0 }, e: { r: 3, c: 0 } }, // A3:A4 — pionowe (wiersz 4 bez komórki A)
    { s: { r: 5, c: 1 }, e: { r: 5, c: 2 } }, // B6:C6 — poziome
    { s: { r: 7, c: 0 }, e: { r: 9, c: 1 } }, // A8:B10 — blok
  ];
  ws["!cols"] = Array.from({ length: cols }, () => ({ wch: 22 }));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Scalenia");
  const file = path.join(fs.mkdtempSync(path.join(os.tmpdir(), "fzm-")), "frozen-merged.xlsx");
  fs.writeFileSync(file, XLSX.write(wb, { type: "buffer", bookType: "xlsx" }));
  return file;
}

(async () => {
  const file = fixture();
  const browser = await ENGINE.launch();
  const ctx = await browser.newContext({ viewport: { width: 1100, height: 800 }, serviceWorkers: "block" });
  await ctx.addInitScript(() => { localStorage.setItem("introPlayed", "true"); });
  const page = await ctx.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  await page.goto(APP_URL + (VIRT ? `?virt=${VIRT}` : ""));
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  await page.setInputFiles("#fileInput", file);
  await page.waitForFunction(() => document.getElementById("sheetSelect")?.options?.length > 0, null, { timeout: 30000 });
  await page.evaluate(() => loadBtn.click());
  await page.waitForFunction(() => Array.isArray(baseRows) && baseRows.length > 0 && document.querySelector("#dataTable tbody td.cell-merged"), null, { timeout: 20000 });
  await page.evaluate(() => { freezeFirstColEl.checked = true; setFrozenColCount(3); applyFreezeFirstColumn(); });
  await page.waitForTimeout(300);

  const measure = () => page.evaluate(() => {
    const guide = [...document.querySelectorAll("#dataTable thead tr.guide-row > th")];
    const bad = [];
    const check = (cell, where) => {
      const k = Number(cell.dataset.colIndex);
      const g = guide[k + 1];
      if (!g) return;
      const d = cell.getBoundingClientRect().left - g.getBoundingClientRect().left;
      if (Math.abs(d) > 1.5) bad.push(`${where} kol.${k} (${cell.textContent.slice(0, 12)}): ${Math.round(d)} px`);
    };
    document.querySelectorAll("#dataTable tbody tr:not(.vr-spacer) > td[data-col-index]").forEach((td) => check(td, `w${td.parentElement.dataset.rowIndex}`));
    document.querySelectorAll("#dataTable thead tr.header-row > th[data-col-index]").forEach((th) => check(th, "nagłówek"));
    return { bad, cells: document.querySelectorAll("#dataTable tbody td[data-col-index]").length, wrap: tableWrapEl.scrollLeft };
  });
  const fails = [];
  const m0 = await measure();
  if (m0.bad.length) fails.push(`bez przewinięcia: ${m0.bad.slice(0, 6).join("; ")}`);
  await page.evaluate(() => { tableWrapEl.scrollLeft = 420; });
  await page.waitForTimeout(500);
  const m1 = await measure();
  if (m1.wrap < 300) fails.push(`tabela nie przewinęła się w bok (${m1.wrap})`);
  if (m1.bad.length) fails.push(`po przewinięciu w bok: ${m1.bad.slice(0, 8).join("; ")}`);
  if (errors.length) fails.push(`błędy strony: ${errors.join(" | ")}`);
  await browser.close();
  const tag = `${process.env.ENGINE || "chromium"}${VIRT ? " virt=0" : ""}`;
  if (fails.length) { console.log(`❌ frozen-merged-rows (${tag}): ${fails.length} błędów`); fails.forEach((f) => console.log("  ✗", f)); process.exit(1); }
  console.log(`✅ frozen-merged-rows (${tag}): wszystko przeszło (${m1.cells} komórek)`);
})().catch((e) => { console.error(e); process.exit(1); });
