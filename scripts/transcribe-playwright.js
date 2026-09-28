// transcribe-playwright.js — test TRYBU SPISYWANIA (przepisywanie danych na papier).
//
// Sprawdza to, co w tym trybie może się realnie zepsuć:
//   1. karta pokazuje WYBRANE pola bieżącego wiersza (i respektuje kolejność),
//   2. „Spisane i dalej” odhacza i idzie o JEDEN wiersz (nie o dwa),
//   3. „Ukryj spisane” chowa odhaczone, a kursor nie przeskakuje,
//   4. odhaczenia i układ pól przeżywają zamknięcie trybu i przeładowanie strony,
//   5. klawiatura (←/→/Spacja/Esc) działa i NIE przecieka do tabeli pod spodem,
//   6. blokada dotyku faktycznie blokuje kartę, a nie dolny pasek.
//
// Uruchom z serwerem na APP_URL (domyślnie http://127.0.0.1:4175/), np. `npm run serve` obok.

const { chromium } = require("playwright");
const path = require("path");

const APP_URL = process.env.APP_URL || "http://127.0.0.1:4175/";
const FILE = path.join(__dirname, "stress-test-workbench.xlsx");
// SLEEP_SCALE (run-tests.mjs): skraca przerwy „odczekaj po kroku” w szybkim przebiegu;
// powtórka porażki idzie ze skalą 1 (pełne przerwy). sleepFixed = przerwa, która MUSI
// przeczekać timer aplikacji (debounce, opóźnienie podpowiedzi, bezczynność) — bez skali.
const SLEEP_SCALE = Math.min(1, Math.max(0.1, Number(process.env.SLEEP_SCALE || 1)));
const sleep = (ms) => new Promise((r) => setTimeout(r, Math.round(ms * SLEEP_SCALE)));
const sleepFixed = (ms) => new Promise((r) => setTimeout(r, ms)); // eslint-disable-line no-unused-vars

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

const state = (page) => page.evaluate(() => window.__transcribe.state());

async function run() {
  const browser = await chromium.launch({ headless: true });
  const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1100, height: 860 } });
  await context.addInitScript(() => {
    localStorage.setItem("introPlayed", "true");
    // Czyścimy stan trybu TYLKO przy pierwszym wejściu — addInitScript odpala się przy
    // każdej nawigacji, a test świadomie przeładowuje stronę, żeby sprawdzić trwałość.
    if (!sessionStorage.getItem("trTestInit")) {
      localStorage.removeItem("excel-workbench-transcribe");
      sessionStorage.setItem("trTestInit", "1");
    }
  });
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push("pageerror: " + e.message));
  page.on("console", (m) => { if (m.type() === "error") errors.push("console: " + m.text()); });

  const failures = [];
  const report = {};

  await loadSheet(page);

  // ── 1. Otwarcie: nakładka widoczna, domyślne pola, karta = wiersz 1 ────────
  await page.click("#transcribeBtn");
  await sleep(250);
  const opened = await state(page);
  report.opened = opened;
  const cardText = await page.evaluate(() => ({
    visible: !document.getElementById("transcribeOverlay").classList.contains("hidden"),
    fields: [...document.querySelectorAll("#trCard .tr-field")].map((f) => ({
      label: f.querySelector(".tr-field-label").textContent,
      value: f.querySelector(".tr-field-value").textContent,
    })),
    counter: document.getElementById("trCounter").textContent,
    focus: document.activeElement?.id || "",
  }));
  report.cardText = cardText;
  if (!cardText.visible) failures.push("nakładka spisywania powinna być widoczna");
  if (!opened.open) failures.push("stan powinien raportować open=true");
  if (opened.total !== opened.rows) failures.push(`bez „ukryj spisane” widoczne=${opened.total} powinno = wszystkie=${opened.rows}`);
  if (!opened.cols.length || opened.cols.length > 10) failures.push(`domyślnie 1..10 pól, jest ${opened.cols.length}`);
  if (cardText.fields.length !== opened.cols.length) failures.push(`kart pól ${cardText.fields.length} ≠ wybranych ${opened.cols.length}`);
  if (cardText.counter !== `1 / ${opened.rows}`) failures.push(`licznik "${cardText.counter}" ≠ "1 / ${opened.rows}"`);
  if (cardText.focus !== "trMarkBtn") failures.push(`fokus po otwarciu na "${cardText.focus}", oczekiwano trMarkBtn`);

  // Wartości na karcie muszą pochodzić z modelu widoku (a nie np. z surowego arkusza).
  const valuesMatch = cardText.fields.map((f) => f.value).join("|") === opened.values.map((v) => v.trim() || "—").join("|");
  if (!valuesMatch) failures.push(`wartości na karcie ≠ wartości modelu: ${JSON.stringify(cardText.fields.map((f) => f.value))} vs ${JSON.stringify(opened.values)}`);

  // ── 2. „Spisane i dalej” — jeden krok, jedno ✓ ────────────────────────────
  const firstKey = await page.evaluate(() => getRowSelectionKey(currentDisplayModel.rows[0]));
  await page.click("#trMarkBtn");
  await sleep(120);
  const afterMark = await state(page);
  report.afterMark = afterMark;
  if (afterMark.pos !== 1) failures.push(`po „Spisane i dalej” pozycja=${afterMark.pos}, oczekiwano 1`);
  if (afterMark.done !== 1) failures.push(`odhaczonych=${afterMark.done}, oczekiwano 1`);

  // ── 3. Klawiatura: ← wraca, Spacja odhacza, strzałki NIE ruszają tabeli ────
  const cellBefore = await page.evaluate(() => JSON.stringify(focusedCellState));
  await page.keyboard.press("ArrowLeft");
  await sleep(80);
  const afterLeft = await state(page);
  await page.keyboard.press("ArrowRight");
  await page.keyboard.press("ArrowRight");
  await sleep(80);
  const afterRight = await state(page);
  const cellAfter = await page.evaluate(() => JSON.stringify(focusedCellState));
  report.keyboard = { afterLeft: afterLeft.pos, afterRight: afterRight.pos, cellBefore, cellAfter };
  if (afterLeft.pos !== 0) failures.push(`← powinna wrócić na 0, jest ${afterLeft.pos}`);
  if (afterRight.pos !== 2) failures.push(`dwa razy → powinny dać 2, jest ${afterRight.pos}`);
  if (cellBefore !== cellAfter) failures.push("strzałki przeciekły do tabeli pod spodem (zmienił się focusedCellState)");

  // Spacja (fokus poza przyciskiem) = odhacz i dalej.
  await page.evaluate(() => document.getElementById("trCard").focus?.());
  await page.evaluate(() => document.activeElement.blur());
  await page.keyboard.press("Space");
  await sleep(120);
  const afterSpace = await state(page);
  report.afterSpace = afterSpace;
  if (afterSpace.done !== 2) failures.push(`Spacja powinna dać 2 odhaczone, jest ${afterSpace.done}`);
  if (afterSpace.pos !== 3) failures.push(`Spacja powinna przesunąć na 3, jest ${afterSpace.pos}`);

  // ── 4. „Ukryj spisane” — znikają odhaczone, kursor nie ucieka ─────────────
  await page.click("#trHideDone");
  await sleep(150);
  const hidden = await state(page);
  report.hidden = hidden;
  if (hidden.total !== hidden.rows - hidden.done) failures.push(`po „ukryj spisane” widoczne=${hidden.total}, oczekiwano ${hidden.rows - hidden.done}`);

  // Odhaczenie przy włączonym ukrywaniu: pozycja ZOSTAJE (pokazuje kolejny wiersz), nie skacze o dwa.
  const posBefore = hidden.pos;
  const keyBefore = await page.evaluate(() => window.__transcribe.state() && document.querySelector("#trCard .tr-field-value").textContent);
  await page.click("#trMarkBtn");
  await sleep(120);
  const afterHiddenMark = await state(page);
  const keyAfter = await page.evaluate(() => document.querySelector("#trCard .tr-field-value")?.textContent);
  report.afterHiddenMark = { posBefore, pos: afterHiddenMark.pos, total: afterHiddenMark.total, keyBefore, keyAfter };
  if (afterHiddenMark.pos !== posBefore) failures.push(`przy „ukryj spisane” pozycja powinna zostać ${posBefore}, jest ${afterHiddenMark.pos}`);
  if (keyBefore === keyAfter) failures.push("po odhaczeniu z „ukryj spisane” karta powinna pokazać KOLEJNY wiersz");
  if (afterHiddenMark.total !== hidden.total - 1) failures.push(`lista powinna skurczyć się o 1 (${hidden.total} → ${afterHiddenMark.total})`);

  // ── 5. Zmiana i kolejność pól ─────────────────────────────────────────────
  await page.click("#trFieldsBtn");
  await sleep(150);
  const fieldsRows = await page.evaluate(() => document.querySelectorAll("#trFieldsList .tr-field-row").length);
  const headersCount = await page.evaluate(() => currentHeaders.length);
  if (fieldsRows !== headersCount) failures.push(`panel pól ma ${fieldsRows} wierszy, kolumn jest ${headersCount}`);
  // Przesuń trzecią kolumnę na sam początek → ma być pierwszym polem na karcie.
  const labelMoved = await page.evaluate(() => {
    const rows = [...document.querySelectorAll("#trFieldsList .tr-field-row")];
    const name = rows[2].querySelector(".tr-field-name").textContent;
    if (!rows[2].querySelector("input").checked) rows[2].querySelector("input").click();
    window.__transcribe.moveField(2, -1);
    window.__transcribe.moveField(1, -1);
    return name;
  });
  await sleep(120);
  await page.click("#trFieldsDoneBtn");
  await sleep(150);
  const firstLabel = await page.evaluate(() => document.querySelector("#trCard .tr-field-label")?.textContent);
  report.reorder = { labelMoved, firstLabel };
  if (labelMoved !== firstLabel) failures.push(`po przesunięciu pierwszym polem ma być "${labelMoved}", jest "${firstLabel}"`);

  // ── 5b. Automatyczny dobór pól z wiersza ─────────────────────────────────
  // Sedno: w arkuszach z powtarzanymi blokami (Kw1_*, Kw2_*, Kw3_*) dane raz siedzą
  // w jednym bloku, raz w kolejnym. Tryb auto ma pokazać to, co w TYM wierszu
  // naprawdę ma wartość — także kolumny spoza ręcznego zaznaczenia.
  await page.click("#trFieldsBtn");
  await sleep(150);
  await page.click("#trAutoFields");
  await sleep(200);
  await page.click("#trFieldsDoneBtn");
  await sleep(200);
  const auto = await state(page);
  const autoCard = await page.evaluate(() => ({
    values: [...document.querySelectorAll("#trCard .tr-field-value")].map((v) => v.textContent),
    labels: [...document.querySelectorAll("#trCard .tr-field-label")].map((v) => v.textContent),
    skipped: document.querySelector(".tr-card-skipped")?.textContent || "",
    btn: document.getElementById("trFieldsBtn").textContent,
  }));
  report.auto = { cols: auto.cols.length, skipped: auto.skipped, card: autoCard };
  if (!auto.auto) failures.push("stan powinien raportować auto=true");
  if (autoCard.values.some((v) => v === "—")) failures.push(`tryb auto nie powinien pokazywać pustych pól: ${JSON.stringify(autoCard.values)}`);
  if (autoCard.values.length !== auto.cols.length) failures.push(`kart pól ${autoCard.values.length} ≠ pól auto ${auto.cols.length}`);
  if (auto.skipped !== 0 && !autoCard.skipped) failures.push("brak informacji o pominiętych pustych polach");
  if (!/auto/i.test(autoCard.btn)) failures.push(`przycisk „Pola" powinien sygnalizować tryb auto, jest "${autoCard.btn}"`);

  // Auto musi sięgać POZA ręczne zaznaczenie — inaczej nie ratuje przesuniętych bloków.
  const reach = await page.evaluate(() => {
    const total = currentHeaders.length;
    const st = window.__transcribe.state();
    return { cols: st.cols, total, max: Math.max(...st.cols) };
  });
  report.autoReach = reach;
  if (reach.cols.length + (auto.skipped || 0) !== reach.total) {
    failures.push(`auto: pokazane (${reach.cols.length}) + pominięte (${auto.skipped}) powinno dać ${reach.total} kolumn`);
  }

  // Wiersz po wierszu zestaw pól MA się zmieniać — to cały sens trybu.
  const beforeCols = (await state(page)).cols.join(",");
  let changed = false;
  for (let i = 0; i < 12 && !changed; i++) {
    await page.evaluate(() => window.__transcribe.go(1));
    const now = (await state(page)).cols.join(",");
    if (now !== beforeCols) changed = true;
  }
  report.autoVaries = changed;
  if (!changed) failures.push("w trybie auto zestaw pól powinien różnić się między wierszami");

  // Powrót do trybu ręcznego oddaje stałą listę.
  await page.evaluate(() => window.__transcribe.setAutoFields(false));
  await sleep(150);
  const backManual = await state(page);
  report.backManual = { auto: backManual.auto, cols: backManual.cols.length };
  if (backManual.auto) failures.push("wyłączenie trybu auto nie zadziałało");
  if (backManual.cols.length !== 10) failures.push(`po powrocie do ręcznego powinno być 10 pól, jest ${backManual.cols.length}`);
  await page.evaluate(() => window.__transcribe.setAutoFields(true));
  await sleep(120);

  // ── 6. Blokada dotyku ─────────────────────────────────────────────────────
  await page.click("#trLockBtn");
  await sleep(120);
  const locked = await page.evaluate(() => ({
    cls: document.getElementById("transcribeOverlay").classList.contains("is-locked"),
    shield: !document.getElementById("trTouchShield").classList.contains("hidden"),
    state: window.__transcribe.state().locked,
  }));
  report.locked = locked;
  if (!locked.cls || !locked.shield || !locked.state) failures.push(`blokada dotyku nieaktywna: ${JSON.stringify(locked)}`);
  // Dolny pasek MUSI zostać klikalny przy blokadzie — inaczej tryb byłby bez wyjścia.
  const posLocked = (await state(page)).pos;
  await page.click("#trMarkBtn");
  await sleep(120);
  const afterLockedMark = await state(page);
  if (afterLockedMark.pos === posLocked && afterLockedMark.done <= afterHiddenMark.done) {
    failures.push("przy blokadzie dotyku dolny pasek powinien nadal działać");
  }
  await page.click("#trLockBtn");
  await sleep(100);

  // ── 7. Trwałość: Esc zamyka, stan wraca po przeładowaniu ──────────────────
  const beforeClose = await state(page);
  await page.keyboard.press("Escape");
  await sleep(200);
  const closed = await page.evaluate(() => ({
    hidden: document.getElementById("transcribeOverlay").classList.contains("hidden"),
    bodyLock: document.body.classList.contains("tr-active"),
    appInert: document.querySelector(".app").hasAttribute("inert"),
    focus: document.activeElement?.id || "",
  }));
  report.closed = closed;
  if (!closed.hidden) failures.push("Esc powinien zamknąć tryb spisywania");
  if (closed.bodyLock) failures.push("po zamknięciu body.tr-active powinno zniknąć");
  if (closed.appInert) failures.push("po zamknięciu inert powinien zejść z .app");
  if (closed.focus !== "transcribeBtn") failures.push(`fokus powinien wrócić na przycisk, jest "${closed.focus}"`);

  await loadSheet(page);
  await page.click("#transcribeBtn");
  await sleep(300);
  const restored = await state(page);
  const restoredFirstLabel = await page.evaluate(() => document.querySelector("#trCard .tr-field-label")?.textContent);
  report.restored = { done: restored.done, hideDone: restored.hideDone, auto: restored.auto, cols: restored.cols.length, restoredFirstLabel };
  if (restored.done !== beforeClose.done) failures.push(`po przeładowaniu odhaczonych=${restored.done}, oczekiwano ${beforeClose.done}`);
  if (!restored.hideDone) failures.push("po przeładowaniu „ukryj spisane” powinno zostać włączone");
  if (!restored.auto) failures.push("po przeładowaniu tryb auto powinien zostać włączony");
  if (restoredFirstLabel !== labelMoved) failures.push(`po przeładowaniu kolejność pól przepadła: "${restoredFirstLabel}" ≠ "${labelMoved}"`);

  // ── 8. Wyczyszczenie ✓ (dwustopniowe) — mieszka w panelu „Postęp" ─────────
  await page.click("#trProgressBtn");
  await sleep(150);
  await page.click("#trResetBtn");
  await sleep(80);
  const armed = await page.evaluate(() => document.getElementById("trResetBtn").classList.contains("is-armed"));
  await page.click("#trResetBtn");
  await sleep(150);
  const afterReset = await state(page);
  report.reset = { armed, done: afterReset.done };
  if (!armed) failures.push("pierwszy klik „Wyczyść ✓” powinien tylko uzbroić przycisk");
  if (afterReset.done !== 0) failures.push(`po wyczyszczeniu odhaczonych=${afterReset.done}, oczekiwano 0`);

  await browser.close();

  if (errors.length) failures.push(`console/page errors: ${errors.join(" | ")}`);
  console.log(JSON.stringify(report, null, 2));
  if (failures.length) throw new Error(failures.join("; "));
  console.log("✅ transcribe-playwright OK");
}

run().catch((error) => {
  console.error(error);
  process.exit(1);
});
