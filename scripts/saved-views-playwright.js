// saved-views-playwright.js — zapisane widoki (filtry + sortowanie pod nazwą), 2026-10-27.
//
// Co musi być prawdą:
//   1. „☆ Zapisz widok” na pasku filtrów zapisuje filtr + sortowanie pod nazwą (prompt),
//   2. po zapisie pasek pokazuje „★ nazwa” zamiast przycisku,
//   3. po „Wyczyść” puste pole szukania pokazuje sekcję „Zapisane widoki” z opisem,
//   4. klik w widok przywraca tę samą liczbę wierszy i to samo sortowanie,
//   5. samo sortowanie (bez filtra) da się zapisać z listy („＋ Zapisz bieżący widok”),
//   6. widok z kolumną, której nie ma w arkuszu, nie włącza się i NIC nie zmienia (powód w opisie),
//   7. × usuwa widok, a widoki przeżywają przeładowanie strony,
//   8. w konsoli brak błędów.
// ENGINE=webkit (Safari/iPad).

const { chromium, webkit } = require("playwright");
const path = require("path");

const APP_URL = process.env.APP_URL || "http://127.0.0.1:4175/";
const FILE = path.join(__dirname, "stress-test-workbench.xlsx");
const ENGINE = process.env.ENGINE === "webkit" ? webkit : chromium;
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

async function openFile(page) {
  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  await page.setInputFiles("#fileInput", FILE);
  await page.waitForFunction(() => document.getElementById("sheetSelect")?.options?.length > 0, null, { timeout: 15000 });
  await sleep(300);
  await page.evaluate(() => loadBtn.click());
  await sleep(1500);
  await page.evaluate(() => { if (typeof setSidebarOpen === "function") setSidebarOpen(false); });
  await sleep(300);
}

async function run() {
  const browser = await ENGINE.launch({ headless: true });
  const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1280, height: 900 } });
  await context.addInitScript(() => sessionStorage.setItem("introPlayed", "true"));
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push("pageerror: " + e.message));
  page.on("console", (m) => { if (m.type() === "error") errors.push("console: " + m.text()); });
  let promptAnswer = "";
  page.on("dialog", (d) => (d.type() === "prompt" ? d.accept(promptAnswer) : d.accept()));
  await openFile(page);

  const failures = [];
  const ok = (name, cond, got) => { if (!cond) failures.push(`${name} — dostałem: ${JSON.stringify(got)}`); console.log(`${cond ? "✅" : "❌"} ${name}`); };
  const state = () => page.evaluate(() => ({
    shown: viewRows.length,
    total: baseRows.length,
    sort: multiSortState.map((r) => `${r.col}:${r.dir}`).join(","),
    first: viewRows.slice(0, 5).map((r) => r.rowIndex0).join(","),
  }));

  // filtr przez menu komórki („Pokaż tylko takie”) w kolumnie Projekt + sortowanie po Rok malejąco
  const projCol = await page.evaluate(() => currentHeaders.findIndex((h) => String(h).trim() === "Projekt"));
  await page.evaluate((c) => {
    const val = baseRows.map((r) => String(getDisplayValue(r, c) ?? "").trim()).find(Boolean);
    quickSearchEl.value = `Projekt:="${val}"`;
    quickSearchOperatorsEnabled = true;
    [filterOperatorsEl, quickSearchOperatorsEl].forEach((el) => { if (el) el.checked = true; });
    commitQuickSearch();
  }, projCol);
  await sleep(600);
  await page.evaluate(() => { multiSortState = [{ col: "Rok", dir: "desc" }]; applyFilters(); sortRows(); scheduleViewRefresh({ table: true, filterBadge: true, immediate: true }); renderActiveFilters(); });
  await sleep(500);
  const filtered = await state();
  ok("filtr zawęża listę", filtered.shown > 0 && filtered.shown < filtered.total, filtered);

  // 1–2. zapis z paska filtrów
  ok("pasek ma „☆ Zapisz widok”", await page.locator("#activeFilters .af-save-view").isVisible());
  promptAnswer = "Projekt wg roku";
  await page.click("#activeFilters .af-save-view");
  await sleep(300);
  const stored = await page.evaluate(() => JSON.parse(localStorage.getItem("swb-saved-views-v1") || "[]"));
  ok("widok zapisany z filtrem i sortowaniem", stored.length === 1 && stored[0].name === "Projekt wg roku" && stored[0].vs.filter && stored[0].vs.sort[0]?.col === "Rok", stored);
  ok("pasek pokazuje „★ Projekt wg roku”", (await page.locator("#activeFilters .af-view").textContent().catch(() => "")) === "★ Projekt wg roku");

  // 3. wyczyść → lista widoków w pustym polu szukania
  await page.click("#activeFilters .af-clear-all");
  await sleep(600);
  await page.evaluate(() => { multiSortState = []; applyFilters(); sortRows(); scheduleViewRefresh({ table: true, immediate: true }); });
  await sleep(300);
  const cleared = await state();
  ok("po Wyczyść widać wszystko", cleared.shown === cleared.total, cleared);
  await page.click("#quickSearch");
  await sleep(400);
  const listText = await page.evaluate(() => document.getElementById("qsLiveResults")?.textContent || "");
  ok("lista: sekcja „Zapisane widoki” z nazwą i opisem sortowania", /Zapisane widoki/.test(listText) && /Projekt wg roku/.test(listText) && /Rok ↓/.test(listText), listText.slice(0, 200));

  // 4. klik = ten sam widok
  await page.click("#qsLiveResults .qs-view-item");
  await sleep(800);
  const restored = await state();
  ok("widok przywraca te same wiersze i sortowanie", restored.shown === filtered.shown && restored.sort === filtered.sort && restored.first === filtered.first, { restored, filtered });
  ok("pasek znów pokazuje nazwę widoku", (await page.locator("#activeFilters .af-view").textContent().catch(() => "")) === "★ Projekt wg roku");

  // 5. samo sortowanie z listy
  await page.click("#activeFilters .af-clear-all");
  await sleep(500);
  await page.evaluate(() => { multiSortState = [{ col: "Osoba", dir: "asc" }]; applyFilters(); sortRows(); scheduleViewRefresh({ table: true, immediate: true }); quickSearchEl.value = ""; });
  await sleep(300);
  await page.evaluate(() => quickSearchEl.blur());
  await page.click("#quickSearch");
  await sleep(400);
  ok("lista proponuje „＋ Zapisz bieżący widok”", await page.locator("#qsLiveResults .qs-view-save").isVisible());
  promptAnswer = "Alfabetycznie";
  await page.click("#qsLiveResults .qs-view-save");
  await sleep(300);
  const two = await page.evaluate(() => JSON.parse(localStorage.getItem("swb-saved-views-v1") || "[]").map((v) => v.name));
  ok("zapisane samo sortowanie (nowy widok na górze)", JSON.stringify(two) === JSON.stringify(["Alfabetycznie", "Projekt wg roku"]), two);

  // 6. widok z brakującą kolumną nie włącza się i nic nie zmienia
  await page.evaluate(() => {
    const list = JSON.parse(localStorage.getItem("swb-saved-views-v1"));
    list.push({ name: "Obcy", vs: { v: 1, filter: null, sort: [{ col: "Kolumna, której nie ma", dir: "asc" }] }, at: 1 });
    localStorage.setItem("swb-saved-views-v1", JSON.stringify(list));
  });
  const before = await state();
  await page.evaluate(() => quickSearchEl.blur());
  await page.click("#quickSearch");
  await sleep(400);
  const off = page.locator("#qsLiveResults .qs-view-item.is-off");
  ok("obcy widok wyszarzony z powodem", (await off.count()) === 1 && /Kolumna, której nie ma/.test(await off.textContent()), await off.count());
  await off.click();
  await sleep(500);
  const after = await state();
  ok("obcy widok nic nie zmienił", JSON.stringify(after) === JSON.stringify(before), { before, after });

  // 7. × usuwa; widoki przeżywają przeładowanie
  await page.evaluate(() => quickSearchEl.blur());
  await page.click("#quickSearch");
  await sleep(400);
  await page.locator("#qsLiveResults .qs-view-row").filter({ hasText: "Obcy" }).locator(".qs-hist-del").click();
  await sleep(300);
  const left = await page.evaluate(() => JSON.parse(localStorage.getItem("swb-saved-views-v1") || "[]").map((v) => v.name));
  ok("× usuwa widok", JSON.stringify(left) === JSON.stringify(["Alfabetycznie", "Projekt wg roku"]), left);
  await openFile(page);
  const persisted = await page.evaluate(() => swbViewsLoad().map((v) => v.name));
  ok("widoki po przeładowaniu strony", persisted.length === 2, persisted);

  ok("bez błędów w konsoli", !errors.length, errors);
  await browser.close();
  const engine = process.env.ENGINE === "webkit" ? "webkit" : "chromium";
  if (failures.length) { console.log(`\n[${engine}] ${failures.length} nie przeszło:\n - ${failures.join("\n - ")}`); process.exit(1); }
  console.log(`\n[${engine}] wszystkie przeszły`);
}

run().catch((e) => { console.error(e); process.exit(1); });
