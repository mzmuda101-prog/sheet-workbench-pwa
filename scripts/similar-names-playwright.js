// similar-names-playwright.js — „Podobne nazwiska” w agregacji (similar-names.js).
//
// Co musi być prawdą:
//   1. warianty tej samej osoby łączą się: kolejność słów, „Gr1 / gr. 4”, tytuł „mgr”, inicjał,
//      polskie znaki („Łukasz” = „Lukasz”), wielkość liter,
//   2. „Julian Olasz” (jedna litera różnicy) NIE łączy się sam — jest propozycją do sprawdzenia,
//   3. samo „Kowalski” zostaje osobno, gdy w danych są dwaj różni Kowalscy,
//   4. „Połącz” scala propozycję, „Rozdziel” cofa; tryb „Dokładnie” liczy jak dawniej.
//
// Uruchom z serwerem na APP_URL (domyślnie http://127.0.0.1:4175/).

const { chromium, webkit } = require("playwright");
const fs = require("fs");
const os = require("os");
const path = require("path");
const XLSX = require("../lib/xlsx.full.min.js");
const SimilarNames = require("../app/similar-names.js");

const APP_URL = process.env.APP_URL || "http://127.0.0.1:4175/";
const ENGINE = process.env.ENGINE === "webkit" ? webkit : chromium;
const results = [];
const check = (name, ok, detail = "") => results.push({ name, ok: !!ok, detail });

const ROWS = [
  ["Julian Olsz", 5], ["Gr1 Julian Olsz", 3], ["Olsz Julian", 1], ["J. Olsz", 2], ["gr. 4 Julian Olsz", 1],
  ["Julian Olasz", 2],
  ["Anna Nowak", 4], ["mgr Anna Nowak", 1], ["A. Nowak", 1],
  ["Łukasz Kowalski", 2], ["Lukasz Kowalski", 1], ["Piotr Kowalski", 1], ["Kowalski", 1],
];

function unit() {
  const items = ROWS.map(([label, count]) => ({ label, count }));
  const r = SimilarNames.buildSimilarNameGroups(items);
  const g = (l) => r.map.get(l);
  check("unit: warianty Juliana Olsza → jedna grupa", ["Gr1 Julian Olsz", "Olsz Julian", "J. Olsz", "gr. 4 Julian Olsz"].every((l) => g(l) === "Julian Olsz"), JSON.stringify([...r.map]));
  check("unit: „Julian Olasz” osobno + propozycja", g("Julian Olasz") === "Julian Olasz" && r.suggestions.some((s) => [s.a, s.b].includes("Julian Olasz")), JSON.stringify(r.suggestions));
  check("unit: tytuł i inicjał (Anna Nowak)", g("mgr Anna Nowak") === "Anna Nowak" && g("A. Nowak") === "Anna Nowak");
  check("unit: polskie znaki (Łukasz = Lukasz)", g("Lukasz Kowalski") === "Łukasz Kowalski");
  check("unit: samo „Kowalski” przy dwóch Kowalskich zostaje osobno", g("Kowalski") === "Kowalski");
  const approved = SimilarNames.buildSimilarNameGroups(items, { extraPairs: [["Julian Olasz", "Julian Olsz"]] });
  check("unit: zatwierdzona para łączy grupy", approved.map.get("Julian Olasz") === "Julian Olsz");
  check("unit: liczby i puste bez zmian", SimilarNames.buildSimilarNameGroups([{ label: "123", count: 1 }, { label: "", count: 1 }]).map.get("123") === "123");
}

function makeWorkbook() {
  const aoa = [["Osoba", "Kwota"]];
  ROWS.forEach(([name, n]) => { for (let i = 0; i < n; i++) aoa.push([name, 100]); });
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(aoa), "Dane");
  const file = path.join(os.tmpdir(), `similar-names-${process.pid}.xlsx`);
  fs.writeFileSync(file, XLSX.write(wb, { type: "buffer", bookType: "xlsx" }));
  return file;
}

async function run() {
  unit();
  const file = makeWorkbook();
  const browser = await ENGINE.launch({ headless: true });
  const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1280, height: 900 } });
  await context.addInitScript(() => localStorage.setItem("introPlayed", "true"));
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push("pageerror: " + e.message));
  page.on("console", (m) => { if (m.type() === "error") errors.push("console: " + m.text()); });
  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  await page.setInputFiles("#fileInput", file);
  await page.waitForFunction(() => document.getElementById("sheetSelect")?.options?.length > 0, null, { timeout: 15000 });
  await page.evaluate(() => loadBtn.click());
  await page.waitForFunction(() => typeof viewRows !== "undefined" && viewRows.length > 0, null, { timeout: 15000 });
  await page.evaluate(() => { document.getElementById("panel-aggregation-workbench").open = true; if (typeof ensureAnalysisHeavy === "function") ensureAnalysisHeavy(); });
  await page.waitForFunction(() => typeof renderSimilarNamesNote === "function", null, { timeout: 15000 });

  const counts = (mode) => page.evaluate((mode) => {
    Object.assign(aggregationWorkbenchState, { groupBy: "Osoba", groupBy2: "", groupBy3: "", measures: ["count_rows"], aggregation: "count", showCount: 999, groupMode: mode, resultSearch: "" });
    renderAggregationWorkbench();
    const res = buildAggregationWorkbenchResult();
    return Object.fromEntries(res.entries.map((e) => [e.label, e.count]));
  }, mode);

  const exact = await counts("exact");
  check("„Dokładnie” — jak dawniej, każdy wariant osobno", exact["Gr1 Julian Olsz"] === 3 && exact["Julian Olsz"] === 5, JSON.stringify(exact));
  const similar = await counts("similar");
  check("„Podobne nazwiska”: Julian Olsz = 12 (5+3+1+2+1)", similar["Julian Olsz"] === 12, JSON.stringify(similar));
  check("„Julian Olasz” osobno (2)", similar["Julian Olasz"] === 2, JSON.stringify(similar));
  check("Anna Nowak = 6, Łukasz Kowalski = 3, Kowalski = 1", similar["Anna Nowak"] === 6 && similar["Łukasz Kowalski"] === 3 && similar["Kowalski"] === 1, JSON.stringify(similar));
  const note = await page.evaluate(() => ({
    summary: document.querySelector(".similar-names-note > summary")?.textContent || "",
    option: [...document.querySelectorAll('[data-aggregation-control="groupmode"] option')].map((o) => o.textContent),
    mergeBtn: !!document.querySelector("[data-similar-merge]"),
  }));
  check("opcja „Podobne nazwiska” w scalaniu grup", note.option.includes("Podobne nazwiska"), JSON.stringify(note.option));
  check("notka: ile wariantów połączono + 1 do sprawdzenia", /Podobne nazwiska: \d+ wariantów połączonych w 3 grup/.test(note.summary) && /1 do sprawdzenia/.test(note.summary), note.summary);

  await page.evaluate(() => document.querySelector("[data-similar-merge]")?.click());
  const merged = await page.evaluate(() => Object.fromEntries(buildAggregationWorkbenchResult().entries.map((e) => [e.label, e.count])));
  check("„Połącz” → Julian Olsz = 14, brak osobnego Olasza", merged["Julian Olsz"] === 14 && !merged["Julian Olasz"], JSON.stringify(merged));
  await page.evaluate(() => document.querySelector("[data-similar-split]")?.click());
  const split = await page.evaluate(() => Object.fromEntries(buildAggregationWorkbenchResult().entries.map((e) => [e.label, e.count])));
  check("„Rozdziel” cofa połączenie", split["Julian Olsz"] === 12 && split["Julian Olasz"] === 2, JSON.stringify(split));

  check("brak błędów strony", !errors.length, errors.join(" | "));
  await browser.close();
  fs.rmSync(file, { force: true });
  let failed = 0;
  results.forEach((r) => { console.log(`${r.ok ? "✅" : "❌"} ${r.name}${r.ok || !r.detail ? "" : ` (${String(r.detail).slice(0, 500)})`}`); if (!r.ok) failed++; });
  if (failed) process.exit(1);
  console.log(`\n✅ similar-names: ${results.length}/${results.length} (${process.env.ENGINE || "chromium"})`);
}
run().catch((e) => { console.error(e); process.exit(1); });
