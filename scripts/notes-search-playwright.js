// notes-search-playwright.js — szukanie i filtrowanie po notatkach komórek (app/cell-notes.js).
//
// Prośba Mateusza (2026-10-20): „filtrować / szukać po tym, że komórka ma notatkę”. Co musi być prawdą
// (plik z openpyxl: „Dane”, nagłówek w wierszu 2, notatka K7 z pliku + dwie dopisane w teście):
//   1. panel filtrów: grupa notatek widoczna z licznikiem, na arkuszu bez notatek — ukryta,
//   2. „Tylko wiersze z notatkami”: działa od razu, pigułka w pasku filtrów, licznik filtrów,
//      „Wyczyść” wyłącza,
//   3. tokeny w polu szukania (bez włączania operatorów): @notatka, !@notatka, @notatka:słowo,
//      @notatka:"fraza", połączenie z tekstem (@notatka && Kowalski),
//   4. „Szukaj też w treści notatek”: bez flagi tekst z notatki nie znajduje wiersza, z flagą — tak,
//   5. następna/poprzednia notatka: Ctrl/⌘+Alt+N (+Shift), alias F8 / Shift+F8, z zawijaniem,
//      w kolejności widoku (po sortowaniu),
//   6. okno notatki: strzałki ‹ › przechodzą między notatkami, zmieniona treść zapisuje się po drodze.

const pw = require("playwright");
const path = require("path");

const APP_URL = process.env.APP_URL || "http://127.0.0.1:4175/";
const ENGINE = process.env.ENGINE || "chromium";
const SLEEP_SCALE = Math.min(1, Math.max(0.1, Number(process.env.SLEEP_SCALE || 1)));
const sleep = (ms) => new Promise((r) => setTimeout(r, Math.round(ms * SLEEP_SCALE)));
const FILE = path.join(__dirname, "fixtures", "openpyxl-features.xlsx");

async function run() {
  const browser = await pw[ENGINE].launch({ headless: true });
  const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1280, height: 860 } });
  await context.addInitScript(() => { localStorage.setItem("introPlayed", "true"); });
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  page.on("dialog", (d) => d.accept());
  const failures = [];
  const ok = (name, cond, got) => { if (!cond) failures.push(`${name} — dostałem: ${JSON.stringify(got)}`); };

  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(() => { document.getElementById("heroSplash")?.remove(); return ensureXlsxLibs(false); });
  await page.setInputFiles("#fileInput", FILE);
  await page.waitForFunction(() => currentHeaderRow === 2 && !!document.querySelector("#dataTable tbody tr[data-row-key]"), null, { timeout: 20000 });
  await sleep(400);

  // notatki: K7 (z pliku), A4 i C10 dopisane
  await page.evaluate(() => {
    swbSetNote(3, 0, "Sprawdzić fakturę 12/2026", "T");
    swbSetNote(9, 2, "Dopytać klienta o termin", "T");
    applyFilters(); renderActiveTable();
  });
  await sleep(200);
  const shown = () => page.evaluate(() => viewRows.map((r) => r.rowIndex0).join(","));
  const search = async (q) => {
    await page.evaluate((q) => { searchQueryEl.value = q; filtersCommitted = true; applyFilters(); sortRows(); renderActiveTable(); }, q);
    await sleep(120);
    return shown();
  };
  const total = await page.evaluate(() => baseRows.length);

  // 1. panel
  const g1 = await page.evaluate(() => ({ hidden: document.getElementById("notesFilterGroup").classList.contains("hidden"), label: document.querySelector("#notesOnly").closest("label").textContent.trim() }));
  ok("1. grupa notatek w panelu widoczna z licznikiem (3)", !g1.hidden && /\(3\)/.test(g1.label), g1);

  // 2. tylko wiersze z notatkami
  await page.evaluate(() => { const cb = document.getElementById("notesOnly"); cb.checked = true; cb.dispatchEvent(new Event("change", { bubbles: true })); });
  await sleep(200);
  const s2 = await shown();
  const chip = await page.evaluate(() => ({ chip: !!document.querySelector('#activeFilters [data-chip="notes"], #activeFilters') && /notatk/i.test(document.getElementById("activeFilters").textContent), badge: document.getElementById("filterBadge")?.textContent }));
  ok("2. „Tylko wiersze z notatkami” od razu: A4, K7, C10", s2 === "3,6,9", s2);
  ok("2. pigułka w pasku filtrów i licznik filtrów", chip.chip && Number(chip.badge) >= 1, chip);
  await page.evaluate(() => document.getElementById("resetFiltersBtn").click());
  await sleep(200);
  ok("2. „Wyczyść filtry” wyłącza filtr notatek", (await page.evaluate(() => ({ on: swbNotesOnly, cb: document.getElementById("notesOnly").checked, n: viewRows.length }))).n === total && !(await page.evaluate(() => swbNotesOnly)), null);

  // 3. tokeny
  ok("3. @notatka", (await search("@notatka")) === "3,6,9", await shown());
  const neg = await search("!@notatka");
  ok("3. !@notatka — wszystkie poza trzema", neg.split(",").length === total - 3 && !/\b(3|6|9)\b/.test(neg.split(",").filter((x) => ["3", "6", "9"].includes(x)).join(",")), neg);
  ok("3. @notatka:faktur", (await search("@notatka:faktur")) === "3", await shown());
  ok("3. @komentarz:\"o termin\" (synonim + fraza)", (await search('@komentarz:"o termin"')) === "9", await shown());
  ok("3. @note (EN)", (await search("@note")) === "3,6,9", await shown());
  const kRow = await page.evaluate(() => { const r = baseRows.find((x) => x.rowIndex0 === 9); return String(r.values[2] || "").split(" ")[0]; });
  ok("3. @notatka && tekst z wiersza", (await search(`@notatka && ${kRow}`)) === "9", { kRow, got: await shown() });

  // 4. szukanie w treści notatek
  const without = await search("termin");
  await page.evaluate(() => { const cb = document.getElementById("notesSearch"); cb.checked = true; cb.dispatchEvent(new Event("change", { bubbles: true })); });
  const withFlag = await search("termin");
  ok("4. bez flagi treść notatki nie znajduje wiersza", !without.split(",").includes("9"), without);
  ok("4. z flagą „Szukaj też w treści notatek” — znajduje", withFlag.split(",").includes("9"), withFlag);
  await page.evaluate(() => { const cb = document.getElementById("notesSearch"); cb.checked = false; cb.dispatchEvent(new Event("change", { bubbles: true })); });
  await search("");
  await page.evaluate(() => { filtersCommitted = false; applyFilters(); renderActiveTable(); });

  // 5. nawigacja skrótem
  const focusAt = () => page.evaluate(() => focusedCellState && `${focusedCellState.rowKey}|${focusedCellState.colIndex0}`);
  await page.evaluate(() => { setFocusedCell("wide:2", 0, { scroll: true }); focusGridCell && focusGridCell(); });
  await sleep(100);
  const seq = [];
  for (let k = 0; k < 4; k++) { await page.keyboard.press("ControlOrMeta+Alt+KeyN"); await sleep(120); seq.push(await focusAt()); }
  ok("5. Ctrl/⌘+Alt+N: A4 → K7 → C10 → (zawinięcie) A4", seq.join(" ") === "wide:3|0 wide:6|10 wide:9|2 wide:3|0", seq);
  await page.keyboard.press("ControlOrMeta+Alt+Shift+KeyN");
  await sleep(120);
  ok("5. z Shift — poprzednia (zawinięcie na C10)", (await focusAt()) === "wide:9|2", await focusAt());
  await page.keyboard.press("Shift+F8");
  await sleep(120);
  ok("5. alias Shift+F8 — poprzednia (K7)", (await focusAt()) === "wide:6|10", await focusAt());
  await page.keyboard.press("F8");
  await sleep(120);
  ok("5. alias F8 — następna (C10)", (await focusAt()) === "wide:9|2", await focusAt());
  const tipOk = await page.waitForFunction(() => /termin/.test(document.querySelector("#cellTooltip .cell-tooltip-usernote")?.textContent || ""), null, { timeout: 2000 }).then(() => true).catch(() => false);
  ok("5. po przejściu widać treść notatki", tipOk, await page.evaluate(() => document.querySelector("#cellTooltip .cell-tooltip-usernote")?.textContent || ""));
  // kolejność widoku: sortowanie malejąco po pierwszej kolumnie odwraca porządek
  await page.evaluate(() => { multiSortState = [{ col: currentHeaders[0], dir: "desc" }]; applyFilters(); sortRows(); renderActiveTable(); setFocusedCell(getRowSelectionKey(currentDisplayModel.rows[0]), 0, { scroll: true }); });
  await sleep(150);
  await page.keyboard.press("ControlOrMeta+Alt+KeyN");
  await sleep(120);
  ok("5. kolejność widoku po sortowaniu (pierwsza: C10)", (await focusAt()) === "wide:9|2", await focusAt());
  await page.evaluate(() => { multiSortState = []; sortState = { col: "", dir: "asc" }; applyFilters(); renderActiveTable(); });

  // 6. okno: ‹ › i zapis po drodze
  await page.evaluate(() => { setFocusedCell("wide:3", 0, { scroll: true }); swbNoteModal.open(3, 0, {}); });
  await sleep(150);
  const m1 = await page.evaluate(() => ({ nav: !document.querySelector("#noteModal .note-modal-nav").hidden, pos: document.querySelector("#noteModal .note-modal-pos").textContent }));
  ok("6. okno: nawigacja widoczna, pozycja 1 z 3", m1.nav && /1\D+3/.test(m1.pos), m1);
  await page.fill("#noteModal .note-modal-text", "Sprawdzić fakturę 12/2026 — zmienione");
  await page.click('#noteModal [data-act="next"]');
  await sleep(200);
  const m2 = await page.evaluate(() => ({ title: document.getElementById("noteModalTitle").textContent, pos: document.querySelector("#noteModal .note-modal-pos").textContent, a4: swbNoteOf(currentSheetName, "A4")?.text, focus: focusedCellState && focusedCellState.rowKey }));
  ok("6. ‹ › przechodzi do K7 (2 z 3), a zmiana A4 zapisała się", /K7/.test(m2.title) && /2\D+3/.test(m2.pos) && /zmienione/.test(m2.a4 || "") && m2.focus === "wide:6", m2);
  await page.keyboard.press("ControlOrMeta+Alt+Shift+KeyN");
  await sleep(200);
  ok("6. skrót działa też w oknie (wstecz → A4)", /A4/.test(await page.evaluate(() => document.getElementById("noteModalTitle").textContent)), null);
  await page.keyboard.press("Escape");
  await sleep(100);

  // 1b. arkusz bez notatek — grupa schowana, skrót mówi „brak notatek”
  await page.evaluate(() => { const other = workbook.SheetNames.find((n) => n !== currentSheetName && n !== "Ukryty"); sheetSelect.value = other; loadBtn.click(); });
  await page.waitForFunction(() => currentSheetName !== "Dane", null, { timeout: 10000 });
  const hid = await page.waitForFunction(() => document.getElementById("notesFilterGroup").classList.contains("hidden"), null, { timeout: 5000 }).then(() => true).catch(() => false);
  ok("1. arkusz bez notatek: grupa w panelu ukryta", hid, await page.evaluate(() => ({ s: currentSheetName, n: swbSheetNoteCount(), only: swbNotesOnly, cls: document.getElementById("notesFilterGroup").className })));

  await browser.close();
  if (errors.length) failures.push("błędy strony: " + errors.join(" | "));
  if (failures.length) {
    console.error(`❌ notes-search (${ENGINE}):\n  - ` + failures.join("\n  - "));
    process.exit(1);
  }
  console.log(`✅ notes-search (${ENGINE}): filtr „tylko z notatkami”, @notatka, szukanie w treści, następna/poprzednia (Ctrl/⌘+Alt+N, F8), ‹ › w oknie`);
}

run().catch((e) => { console.error(e); process.exit(1); });
