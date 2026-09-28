// transcribe-scroll-turbo-playwright.js — testy DOŁOŻONYCH mechanizmów trybu spisywania:
//
//   1. sygnał przewijania: gdy karta nie mieści się na ekranie, pojawia się belka
//      + pigułka „jeszcze N ↓"; po dojechaniu na dół pigułka gaśnie,
//   2. każdy NOWY wiersz zaczyna się od góry (nie dziedziczy przewinięcia poprzedniego),
//   3. przytrzymanie „Spisane i dalej" (spacja / palec) rozpędza odhaczanie,
//      a puszczenie przed progiem daje zwykłe pojedyncze odhaczenie,
//   4. całą serię da się cofnąć jednym przyciskiem — i wraca kursor na wiersz startowy.
//
// Uruchom z serwerem na APP_URL (domyślnie http://127.0.0.1:4175/).

const { chromium } = require("playwright");
const path = require("path");

const APP_URL = process.env.APP_URL || "http://127.0.0.1:4175/";
const FILE = path.join(__dirname, "stress-test-workbench.xlsx");
// SLEEP_SCALE (run-tests.mjs): skraca przerwy „odczekaj po kroku” w szybkim przebiegu;
// powtórka porażki idzie ze skalą 1 (pełne przerwy). sleepFixed = przerwa, która MUSI
// przeczekać timer aplikacji (debounce, opóźnienie podpowiedzi, bezczynność) — bez skali.
// Ten test czeka na timery aplikacji (powiadomienia, bezczynność, schowek) — przy krótszych
// przerwach padał (sprawdzone SLEEP_SCALE=0.5, 2026-09-28), więc zawsze pełne przerwy.
const SLEEP_SCALE = 1;
const sleep = (ms) => new Promise((r) => setTimeout(r, Math.round(ms * SLEEP_SCALE)));
const sleepFixed = (ms) => new Promise((r) => setTimeout(r, ms)); // eslint-disable-line no-unused-vars
const state = (page) => page.evaluate(() => window.__transcribe.state());

async function loadSheet(page) {
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
}

async function run() {
  const browser = await chromium.launch({ headless: true });
  // Mały ekran = karta z 16 polami NA PEWNO się nie mieści (o to w tym teście chodzi).
  const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 900, height: 560 } });
  await context.addInitScript(() => {
    localStorage.setItem("introPlayed", "true");
    localStorage.removeItem("excel-workbench-transcribe");
  });
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push("pageerror: " + e.message));
  page.on("console", (m) => { if (m.type() === "error") errors.push("console: " + m.text()); });

  const failures = [];
  const report = {};
  const ok = (name, cond, got) => { if (!cond) failures.push(`${name} — dostałem: ${JSON.stringify(got)}`); };

  await loadSheet(page);
  await page.click("#transcribeBtn");
  await sleep(300);

  // Wszystkie kolumny na kartę → gwarantowane przepełnienie ekranu.
  await page.evaluate(() => {
    const n = window.__transcribe.state().rows ? 16 : 16;
    window.__transcribe.setFields(Array.from({ length: n }, (_, i) => i));
  });
  await sleep(250);

  // ── 1. Sygnał przewijania ────────────────────────────────────────────────
  const overflow = await state(page);
  report.overflow = { canScroll: overflow.canScroll, ui: overflow.overflowUi, atBottom: overflow.atBottom };
  ok("karta dłuższa niż ekran", overflow.canScroll, overflow.canScroll);
  ok("wskaźnik przewijania włączony", overflow.overflowUi, overflow.overflowUi);
  ok("start: nie jesteśmy na dole", !overflow.atBottom, overflow.atBottom);

  const pill = await page.evaluate(() => {
    const el = document.getElementById("trScrollMore");
    const rail = document.getElementById("trScrollRail");
    const cs = getComputedStyle(el);
    return {
      text: el.textContent.trim(),
      visible: cs.visibility !== "hidden" && Number(cs.opacity) > 0.5,
      railVisible: Number(getComputedStyle(rail).opacity) > 0.5,
      thumbH: document.getElementById("trScrollThumb").style.height,
    };
  });
  report.pill = pill;
  ok("pigulka wiecej widoczna", pill.visible, pill);
  ok("pigulka podaje liczbe pol ponizej", /\d/.test(pill.text), pill.text);
  ok("belka przewijania widoczna", pill.railVisible, pill);
  ok("suwak belki ma wyliczoną wysokość", /px$/.test(pill.thumbH), pill.thumbH);

  // Klik w pigułkę przewija w dół; na samym dole pigułka gaśnie.
  await page.evaluate(() => { const s = document.getElementById("trStage"); s.scrollTop = s.scrollHeight; s.dispatchEvent(new Event("scroll")); });
  await sleep(200);
  const bottom = await state(page);
  const pillBottom = await page.evaluate(() => {
    const cs = getComputedStyle(document.getElementById("trScrollMore"));
    return cs.visibility !== "hidden" && Number(cs.opacity) > 0.5;
  });
  report.bottom = { atBottom: bottom.atBottom, pillVisible: pillBottom, scrollTop: bottom.scrollTop };
  ok("dół: stan at-bottom", bottom.atBottom, bottom);
  ok("dol: pigulka schowana", !pillBottom, pillBottom);

  // ── 2. Nowy wiersz startuje od góry ──────────────────────────────────────
  await page.evaluate(() => window.__transcribe.go(1));
  await sleep(200);
  const afterNext = await state(page);
  report.resetScroll = { scrollTop: afterNext.scrollTop, pos: afterNext.pos };
  ok("po przejściu dalej karta jest przewinięta na górę", afterNext.scrollTop === 0, afterNext.scrollTop);

  await page.evaluate(() => { const s = document.getElementById("trStage"); s.scrollTop = 200; });
  await page.evaluate(() => window.__transcribe.mark());
  await sleep(200);
  const afterMark = await state(page);
  report.resetScrollAfterMark = afterMark.scrollTop;
  ok("po Spisane i dalej tez od gory", afterMark.scrollTop === 0, afterMark.scrollTop);

  // ── 3. Przytrzymanie = szybkie odhaczanie ────────────────────────────────
  await page.evaluate(() => window.__transcribe.reset());
  await sleep(150);
  const beforeTurbo = await state(page);

  // 3a. krótkie naciśnięcie spacji na przycisku → JEDNO odhaczenie, bez turbo
  await page.focus("#trMarkBtn");
  await page.keyboard.down(" ");
  await sleep(120);
  await page.keyboard.up(" ");
  await sleep(150);
  const shortTap = await state(page);
  report.shortTap = { doneBefore: beforeTurbo.done, doneAfter: shortTap.done, pos: shortTap.pos };
  ok("krótka spacja odhacza dokładnie jeden wiersz", shortTap.done === beforeTurbo.done + 1, report.shortTap);
  ok("krótka spacja nie uruchamia turbo", !shortTap.turbo, shortTap.turbo);

  // 3b. przytrzymanie spacji → seria
  await page.keyboard.down(" ");
  await sleep(1500);
  const during = await state(page);
  await page.keyboard.up(" ");
  await sleep(250);
  const afterHold = await state(page);
  report.turbo = { turboDuring: during.turbo, doneDuring: during.done, doneAfter: afterHold.done, burst: afterHold.burst, undo: afterHold.undoVisible };
  ok("w trakcie trzymania turbo jest aktywne", during.turbo, during.turbo);
  ok("turbo odhaczyło wiele wierszy", afterHold.done >= shortTap.done + 5, report.turbo);
  ok("po puszczeniu turbo stoi", !afterHold.turbo, afterHold.turbo);
  ok("pojawił się przycisk cofnięcia serii", afterHold.undoVisible, afterHold.undoVisible);

  // 3c. seria zapisuje stan do localStorage RAZ, po zakończeniu (nie co wiersz)
  const persisted = await page.evaluate(() => {
    const store = JSON.parse(localStorage.getItem("excel-workbench-transcribe") || "{}");
    const scope = Object.values(store.scopes || {})[0] || {};
    return (scope.done || []).length;
  });
  report.persisted = { inStore: persisted, inMemory: afterHold.done };
  ok("po serii stan trafil do pamieci trwalej", persisted === afterHold.done, report.persisted);

  // ── 4. Cofnięcie serii ───────────────────────────────────────────────────
  await page.evaluate(() => window.__transcribe.undoBurst());
  await sleep(200);
  const undone = await state(page);
  report.undo = { done: undone.done, expected: shortTap.done, pos: undone.pos, undoVisible: undone.undoVisible };
  ok("cofnięcie zdejmuje ✓ tylko z serii", undone.done === shortTap.done, report.undo);
  ok("przycisk cofnięcia znika po użyciu", !undone.undoVisible, undone.undoVisible);

  console.log(JSON.stringify(report, null, 2));
  if (errors.length) failures.push("błędy strony: " + errors.join(" | "));
  await browser.close();

  if (failures.length) {
    console.error("\n❌ transcribe-scroll-turbo-playwright FAIL:\n" + failures.map((f) => " - " + f).join("\n"));
    process.exit(1);
  }
  console.log("✅ transcribe-scroll-turbo-playwright OK");
}

run().catch((e) => { console.error(e); process.exit(1); });
