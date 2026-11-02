// virt-fast-scroll-playwright.js — nowy silnik przy SZYBKIM przewijaniu (kółko, duże skoki).
//
// Zgłoszenie 2026-11-02 (nagranie z szybkiego Maca): przy szybkim przewijaniu tabela chwilami była
// pusta — przeglądarka przesuwa obraz, zanim skrypt dorysuje wiersze, a stały zapas 1 ekranu nie
// wystarczał na skoki przy 120 Hz. Zapas rośnie teraz z prędkością (z przodu do 4 ekranów).
// Miernik: przy każdym zdarzeniu scroll, ZANIM silnik dorysuje — czy widok mieści się w narysowanych
// wierszach. Przed zmianą: 25 na ~108 zdarzeń, po: 0. Tu: najwyżej 2 (zapas na wahania przeglądarki).
// Arkusz: fixtures/virt-merges.xlsx „Opisy” (400 w.), bez zawijania.

const { chromium, webkit } = require("playwright");
const path = require("path");

const APP_URL = process.env.APP_URL || "http://127.0.0.1:4175/";
const ENGINE = process.env.ENGINE === "webkit" ? webkit : chromium;
const FILE = path.join(__dirname, "fixtures", "virt-merges.xlsx");

async function run() {
  const browser = await ENGINE.launch({ headless: true });
  const ctx = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1400, height: 900 } });
  await ctx.addInitScript(() => localStorage.setItem("introPlayed", "true"));
  const page = await ctx.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  await page.goto(APP_URL + "?virt=1", { waitUntil: "load" });
  await page.evaluate(() => { document.getElementById("heroSplash")?.remove(); try { ensureXlsxLibs(false); } catch {} });
  await page.setInputFiles("#fileInput", FILE);
  await page.waitForFunction(() => typeof baseRows !== "undefined" && baseRows.length > 0, null, { timeout: 30000 });
  await page.evaluate(() => { sheetSelect.value = "Opisy"; loadBtn.click(); });
  await page.waitForFunction(() => currentSheetName === "Opisy" && baseRows.length > 300, null, { timeout: 20000 });
  await page.waitForTimeout(800);
  const box = await page.evaluate(() => { const r = tableWrapEl.getBoundingClientRect(); return { x: r.left + r.width / 2, y: r.top + r.height / 2 }; });
  await page.mouse.move(box.x, box.y);
  await page.evaluate(() => {
    window.__miss = { checks: 0, misses: 0 };
    tableWrapEl.addEventListener("scroll", () => {
      const wb = tableWrapEl.getBoundingClientRect();
      const rows = document.querySelectorAll("#dataTable tbody tr[data-row-key]");
      if (!rows.length) return;
      const first = rows[0].getBoundingClientRect(), last = rows[rows.length - 1].getBoundingClientRect();
      window.__miss.checks++;
      if (first.top > wb.top + 120 || last.bottom < wb.bottom - 2) window.__miss.misses++;
    }, { passive: true, capture: true });
  });
  for (const delta of [400, 900]) {
    for (let i = 0; i < 30; i++) { await page.mouse.wheel(0, i < 15 ? delta : -delta); await page.waitForTimeout(8); }
  }
  const r = await page.evaluate(() => ({ ...window.__miss, virt: swbVirt.isActive() }));
  const failures = [];
  if (!r.virt) failures.push("nowy silnik nie działa");
  if (!(r.checks > 20)) failures.push(`za mało zdarzeń przewijania: ${JSON.stringify(r)}`);
  if (r.misses > 2) failures.push(`widok wyjeżdżał poza narysowane wiersze: ${JSON.stringify(r)}`);
  if (errors.length) failures.push(`błędy strony: ${errors.join("; ")}`);
  await browser.close();
  if (failures.length) {
    console.log(`❌ virt-fast-scroll (${process.env.ENGINE || "chromium"}): ${failures.length} błędów`);
    failures.forEach((f) => console.log("  ✗", f));
    process.exit(1);
  }
  console.log(`✅ virt-fast-scroll (${process.env.ENGINE || "chromium"}): ${r.misses}/${r.checks} wyjazdów poza wiersze`);
}

run().catch((e) => { console.error(e); process.exit(1); });
