// monthly-summary-playwright.js — testy panelu „Podsumowanie miesięczne" (analysis.js).
//
// Po co: panel jest generyczny (dowolny arkusz z datami) i ma sporo logiki — wiele miar
// (wystąpienia/wiersze/suma/śr./min/maks), miara na kolumnie liczbowej/dacie/duracji,
// formatowanie wyniku zależne od typu (data → YYYY-MM-DD, duracja → „Xm Yd"), rozbicie
// per kolumna + tooltip. Łatwo o regresję. Test wstrzykuje syntetyczny model i sprawdza
// renderMonthlySummary() na realnym kodzie aplikacji.
//
// Uruchom z serwerem na APP_URL (domyślnie http://127.0.0.1:4175/).

const { chromium } = require("playwright");

const APP_URL = process.env.APP_URL || "http://127.0.0.1:4175/";

async function run() {
  const browser = await chromium.launch({ headless: true });
  const context = await browser.newContext({ serviceWorkers: "block" });
  await context.addInitScript(() => localStorage.setItem("introPlayed", "true"));
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  page.on("console", (m) => { if (m.type() === "error") errors.push(m.text()); });

  await page.goto(APP_URL, { waitUntil: "load" });
  await page.waitForFunction(() => typeof renderMonthlySummary === "function" && typeof currentDisplayModel !== "undefined");

  const result = await page.evaluate(() => {
    // Syntetyczny model: 3 wiersze, kolumny od/do (daty), Długość (duracja), Kwota (liczba).
    const D = (y, m, d) => new Date(y, m - 1, d);
    const mkRow = (od, dood, dlug, kwota, i) => ({
      values: [od, dood, dlug, kwota],
      display: [`${od.getFullYear()}-${String(od.getMonth() + 1).padStart(2, "0")}-${String(od.getDate()).padStart(2, "0")}`,
        `${dood.getFullYear()}-${String(dood.getMonth() + 1).padStart(2, "0")}-${String(dood.getDate()).padStart(2, "0")}`,
        dlug, String(kwota)],
      rowIndex0: i,
    });
    const model = {
      headers: ["od", "do", "Długość", "Kwota"],
      rows: [
        mkRow(D(2026, 1, 10), D(2026, 2, 15), "1m 5d", 100, 0), // od=sty
        mkRow(D(2026, 1, 20), D(2026, 3, 1), "1m 9d", 200, 1),  // od=sty
        mkRow(D(2026, 2, 5), D(2026, 2, 25), "20d", 50, 2),     // od=lut
      ],
    };
    currentDisplayModel = model;

    const el = document.getElementById("monthlySummary");
    const render = (state) => { monthlySummaryState = Object.assign({ months: 12, anchor: "data", split: true }, state); renderMonthlySummary(); };
    const valuesByMonth = () => {
      const map = {};
      el.querySelectorAll(".monthly-row").forEach((r) => { map[r.querySelector(".monthly-label").textContent] = r.querySelector(".monthly-value").textContent; });
      return map;
    };
    const checks = [];
    const ok = (name, cond, got) => checks.push({ name, ok: !!cond, got });

    // A) wystąpienia, kolumna od → sty=2, lut=1
    render({ dateCols: [0], metric: "occurrences", measureCol: null });
    let v = valuesByMonth();
    ok("occurrences sty=2", v["sty 2026"] === "2", v["sty 2026"]);
    ok("occurrences lut=1", v["lut 2026"] === "1", v["lut 2026"]);

    // B) wiersze unikalne, od → sty=2, lut=1
    render({ dateCols: [0], metric: "rows", measureCol: null });
    v = valuesByMonth();
    ok("rows sty=2", v["sty 2026"] === "2", v["sty 2026"]);

    // C) suma duracji (Długość), od → sty=2m 14d (35+39=74d), lut=20d
    render({ dateCols: [0], metric: "sum", measureCol: 2 });
    v = valuesByMonth();
    ok("sum duration sty=2m 14d", v["sty 2026"] === "2m 14d", v["sty 2026"]);
    ok("sum duration lut=20d", v["lut 2026"] === "20d", v["lut 2026"]);

    // D) średnia daty (do), od → lut=2026-02-25 (1 wiersz), sty pasuje do wzoru daty
    render({ dateCols: [0], metric: "avg", measureCol: 1 });
    v = valuesByMonth();
    ok("avg date lut=2026-02-25", v["lut 2026"] === "2026-02-25", v["lut 2026"]);
    ok("avg date sty = data (YYYY-MM-DD)", /^\d{4}-\d{2}-\d{2}$/.test(v["sty 2026"] || ""), v["sty 2026"]);

    // E) średnia liczby (Kwota), od → sty=150, lut=50
    render({ dateCols: [0], metric: "avg", measureCol: 3 });
    v = valuesByMonth();
    ok("avg number sty=150", v["sty 2026"] === "150", v["sty 2026"]);
    ok("avg number lut=50", v["lut 2026"] === "50", v["lut 2026"]);

    // F) multi-kolumna od+do, wystąpienia → lut=3 (od row2 + do row0,row2), mar=1 (do row1)
    render({ dateCols: [0, 1], metric: "occurrences", measureCol: null });
    v = valuesByMonth();
    ok("multi occ lut=3", v["lut 2026"] === "3", v["lut 2026"]);
    ok("multi occ mar=1", v["mar 2026"] === "1", v["mar 2026"]);
    const hasSplitSeg = el.querySelector(".monthly-bar.split .monthly-seg");
    ok("multi: pojedynczy słupek z segmentami", !!hasSplitSeg && !el.querySelector(".monthly-lanes"));
    const hintEl = [...el.querySelectorAll(".monthly-value-hint")].find((x) => (x.getAttribute("data-hint-pl") || "").includes("·"));
    ok("multi: tooltip per kolumna na wyniku", !!hintEl, hintEl ? hintEl.getAttribute("data-hint-pl") : null);

    // G) suma niedozwolona dla miary datowej (do) → spada do occurrences (brak crashu)
    render({ dateCols: [0], metric: "sum", measureCol: 1 });
    const metricOpts = [...el.querySelectorAll('select[data-monthly-control="metric"] option')].map((o) => o.value);
    ok("data jako miara: brak opcji 'sum'", !metricOpts.includes("sum"), metricOpts.join(","));

    return { checks };
  });

  await browser.close();

  const failures = result.checks.filter((c) => !c.ok);
  console.log("monthly-summary:\n  " + result.checks.map((c) => `${c.ok ? "✅" : "❌"} ${c.name}${c.ok ? "" : ` (got: ${c.got})`}`).join("\n  "));
  if (errors.length) failures.push({ name: "błędy konsoli: " + errors.join(" | ") });
  if (failures.length) {
    console.error("\n❌ monthly-summary FAIL: " + failures.map((f) => f.name).join("; "));
    process.exit(1);
  }
  console.log("\n✅ monthly-summary: wszystkie asercje OK");
}

run().catch((error) => { console.error(error); process.exit(1); });
