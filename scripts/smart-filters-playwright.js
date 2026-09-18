// smart-filters-playwright.js — TRYBY AUTO (filtrowanie po stanie rekordu w cyklach).
//
// Co musi być prawdą (liczone NIEZALEŻNIE od kodu funkcji — wprost z nagłówków
// „od"/„do" i zawartości komórek, więc test nie potwierdza sam siebie):
//   1. model bloków wykrywa cykle i role kolumn (start/koniec/osoba),
//   2. chip „W toku" zostawia DOKŁADNIE wiersze z cyklem zaczętym bez daty końca,
//   3. KORELACJA: „osoba + w toku" to wiersze, gdzie TA osoba nie skończyła
//      (a nie: osoba gdziekolwiek ORAZ cokolwiek otwarte) — to jest sedno funkcji,
//   4. widok Wide-to-Long pokazuje wyłącznie trafione cykle,
//   5. token „@wtoku" w polu szukania daje ten sam wynik co chip,
//   6. „Zakończone" i „Ostatni cykl" zgadzają się z ręcznym przeliczeniem,
//   7. wyczyszczenie trybów wraca do pełnego widoku, bez błędów w konsoli.
//
// Uruchom z serwerem na APP_URL (domyślnie http://127.0.0.1:4175/), np. `npm run serve` obok.

const { chromium } = require("playwright");
const path = require("path");

const APP_URL = process.env.APP_URL || "http://127.0.0.1:4175/";
const FILE = path.join(__dirname, "RODO Obieg terenów 2026 V5.xlsx");
const SHEET = "2025-2026";
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
  await page.waitForFunction(() => document.getElementById("sheetSelect")?.options?.length > 0, null, { timeout: 20000 });
  await page.evaluate((sheet) => {
    const select = document.getElementById("sheetSelect");
    if ([...select.options].some((o) => o.value === sheet)) select.value = sheet;
    select.dispatchEvent(new Event("change", { bubbles: true }));
  }, SHEET);
  await sleep(400);
  // Ten arkusz ma nad nagłówkiem wiersz z etykietami cykli — autodetekcja nagłówka
  // trafia w wiersz 2 (jak u użytkownika po wczytaniu pliku).
  await page.evaluate(() => {
    const auto = document.getElementById("autoHeaderRow");
    if (auto) { auto.checked = true; auto.dispatchEvent(new Event("change", { bubbles: true })); }
    if (typeof applyAutoHeaderRowIfEnabled === "function") applyAutoHeaderRowIfEnabled();
  });
  await sleep(400);
  await page.click("#loadBtn");
  await sleep(1500);

  // ── Niezależne oczekiwania: pary „od"/„do" wprost z nagłówków ────────────────
  const expected = await page.evaluate(() => {
    const norm = (h) => String(h ?? "").toLowerCase().replace(/[\s.]+/g, "").replace(/\d+$/, "");
    const pairs = [];
    currentHeaders.forEach((header, i) => {
      if (norm(header) !== "od") return;
      const endIdx = norm(currentHeaders[i + 1]) === "do" ? i + 1 : -1;
      const entIdx = String(currentHeaders[i - 1] ?? "").toLowerCase().includes("imi") ? i - 1 : -1;
      if (endIdx > 0) pairs.push({ startIdx: i, endIdx, entIdx });
    });
    const text = (row, i) => (i >= 0 ? String(getDisplayValue(row, i) ?? "").trim() : "");
    const openRows = [];
    const closedRows = [];
    const personCount = new Map();
    baseRows.forEach((row) => {
      let hasOpen = false;
      let hasClosed = false;
      pairs.forEach((p) => {
        const from = text(row, p.startIdx);
        const to = text(row, p.endIdx);
        if (from && !to) hasOpen = true;
        if (from && to) hasClosed = true;
        const person = text(row, p.entIdx);
        if (person && from && !to) {
          const entry = personCount.get(person) || { open: 0, rows: new Set() };
          entry.open += 1;
          entry.rows.add(row.rowIndex0);
          personCount.set(person, entry);
        }
      });
      if (hasOpen) openRows.push(row.rowIndex0);
      if (hasClosed) closedRows.push(row.rowIndex0);
    });

    // Osoba do testu korelacji: ma otwarty cykl, ale występuje w wielu wierszach
    // (czyli „samo nazwisko" złapałoby więcej niż „nazwisko + w toku").
    let probe = null;
    for (const [person, entry] of personCount) {
      const anywhere = baseRows.filter((row) => row.values.some((_, i) => String(getDisplayValue(row, i) ?? "").trim() === person)).length;
      if (anywhere > entry.rows.size) { probe = { person, openRows: [...entry.rows], anywhere }; break; }
    }
    return { pairs: pairs.length, openRows, closedRows, probe, total: baseRows.length };
  });

  const model = await page.evaluate(() => {
    const m = getSmartModel();
    return m && { blocks: m.blocks.length, span: m.span, startIdx: m.startIdx, endIdx: m.endIdx, entityIdx: m.entityIdx, hasStateColumns: m.hasStateColumns };
  });

  const applyState = (state, query = "") => page.evaluate(({ state, query }) => {
    searchQueryEl.value = query;
    if (filterOperatorsEl) filterOperatorsEl.checked = query.includes("&&");
    smartFilterState = Object.assign({ states: [], rowFlags: [], only: "", minDays: null }, state);
    filtersCommitted = true;
    applyFilters();
    sortRows();
    const longModel = buildLongViewModelFromRows(viewRows);
    return {
      rows: viewRows.map((r) => r.rowIndex0),
      longRows: longModel.rows.length,
      longBlocks: longModel.rows.map((r) => `${r.sourceRowIndex0}:${r.sourceBlockIndex}`),
    };
  }, { state, query });

  const openRun = await applyState({ states: ["open"] });
  const closedRun = await applyState({ states: ["closed"] });
  const lastRun = await applyState({ only: "last" });
  const probePerson = expected.probe ? expected.probe.person : "";
  const personPlain = probePerson ? await applyState({}, probePerson) : null;
  const personOpen = probePerson ? await applyState({ states: ["open"] }, probePerson) : null;
  const tokenRun = await page.evaluate(() => {
    searchQueryEl.value = "@wtoku";
    smartFilterState = { states: [], rowFlags: [], only: "", minDays: null };
    filtersCommitted = true;
    applyFilters();
    sortRows();
    return { rows: viewRows.map((r) => r.rowIndex0), effective: smartEffectiveState };
  });
  const clearedRun = await applyState({});

  await browser.close();

  const sameSet = (a, b) => a.length === b.length && a.every((v, i) => v === b[i]);
  const sorted = (arr) => arr.slice().sort((x, y) => x - y);

  const failures = [];
  if (!model) failures.push("smart model not detected for a wide cyclical sheet");
  else {
    if (model.blocks < 2) failures.push(`expected >= 2 cycle blocks, got ${model.blocks}`);
    if (!model.hasStateColumns) failures.push("start/end columns not detected (od/do)");
    if (model.startIdx < 0 || model.endIdx < 0) failures.push(`start/end role indexes wrong: ${model.startIdx}/${model.endIdx}`);
    if (model.entityIdx < 0) failures.push("person column inside block not detected");
  }
  if (!expected.pairs) failures.push("test fixture has no od/do pairs — check the sheet");
  if (!sameSet(sorted(openRun.rows), sorted(expected.openRows))) {
    failures.push(`"in progress" rows ${openRun.rows.length} != expected ${expected.openRows.length}`);
  }
  if (!sameSet(sorted(closedRun.rows), sorted(expected.closedRows))) {
    failures.push(`"finished" rows ${closedRun.rows.length} != expected ${expected.closedRows.length}`);
  }
  if (openRun.longRows !== openRun.longBlocks.length || openRun.longRows === 0) {
    failures.push("long view produced no cycle rows for the in-progress filter");
  }
  if (openRun.longRows > openRun.rows.length * 2) {
    failures.push(`long view shows too many cycles (${openRun.longRows}) for ${openRun.rows.length} matching rows`);
  }
  if (lastRun.longRows !== lastRun.rows.length) {
    failures.push(`"last cycle" should yield exactly one cycle per row: ${lastRun.longRows} vs ${lastRun.rows.length}`);
  }
  if (expected.probe) {
    if (!sameSet(sorted(personOpen.rows), sorted(expected.probe.openRows))) {
      failures.push(`correlation broken for "${probePerson}": ${personOpen.rows.length} rows, expected ${expected.probe.openRows.length}`);
    }
    if (personPlain.rows.length <= personOpen.rows.length) {
      failures.push(`probe person "${probePerson}" is not discriminating (plain ${personPlain.rows.length} vs open ${personOpen.rows.length})`);
    }
  }
  if (!sameSet(sorted(tokenRun.rows), sorted(openRun.rows))) {
    failures.push("@wtoku token result differs from the chip result");
  }
  if (clearedRun.rows.length !== expected.total) {
    failures.push(`clearing auto modes should restore ${expected.total} rows, got ${clearedRun.rows.length}`);
  }
  if (errors.length) failures.push(`console/page errors: ${errors.join(" | ")}`);

  console.log(JSON.stringify({
    model,
    expected: { pairs: expected.pairs, total: expected.total, open: expected.openRows.length, closed: expected.closedRows.length, probe: expected.probe && { person: expected.probe.person, openRows: expected.probe.openRows.length, anywhere: expected.probe.anywhere } },
    openRun: { rows: openRun.rows.length, longRows: openRun.longRows },
    closedRun: { rows: closedRun.rows.length },
    lastRun: { rows: lastRun.rows.length, longRows: lastRun.longRows },
    personPlain: personPlain && personPlain.rows.length,
    personOpen: personOpen && personOpen.rows.length,
    tokenRows: tokenRun.rows.length,
    clearedRows: clearedRun.rows.length,
    errors,
  }, null, 2));
  if (failures.length) throw new Error(failures.join("; "));
  console.log("✅ smart-filters-playwright OK");
}

run().catch((error) => {
  console.error(error);
  process.exit(1);
});
