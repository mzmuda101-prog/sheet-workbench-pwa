// transcribe-inherit-playwright.js — DZIEDZICZENIE Z GÓRY w trybie spisywania.
//
// Układ, o który chodzi (bardzo częsty w arkuszach operacyjnych): nazwisko scalone
// przez kilka wierszy, pod spodem pozycje. W danych wiersze-kontynuacje są PUSTE,
// więc bez dziedziczenia tryb „dobieraj pola z wiersza” ukrywałby pole tożsamości
// dokładnie tam, gdzie jest najbardziej potrzebne.
//
// Sprawdzane:
//   1. domyślnie WYŁĄCZONE — brak regresji względem poprzedniego zachowania,
//   2. kolumny ze scaleniami pionowymi są wykrywane i proponowane same,
//   3. wiersz-kontynuacja dostaje wartość + numer wiersza źródłowego,
//   4. przenoszenie resetuje się na nowym rekordzie (Kowalski → Nowak),
//   5. kolumna NIE wskazana nie dziedziczy (najważniejszy bezpiecznik: żadnych
//      cudzych wartości w rubryce), a wskazana ręcznie — owszem,
//   6. wynik NIE zależy od filtra: po odfiltrowaniu wiersza-kotwicy kontynuacja
//      nadal dziedziczy poprawną wartość (liczone po baseRows, nie po widoku),
//   7. ustawienie przeżywa przeładowanie,
//   8. w widoku Wide-to-Long opcja jest zablokowana (świadome ograniczenie).
//
// Fixture budowany W PRZEGLĄDARCE (XLSX strony) → base64 → setInputFiles bufor.
// Uruchom z serwerem na APP_URL (domyślnie http://127.0.0.1:4175/).

const { chromium } = require("playwright");
const path = require("path");

const APP_URL = process.env.APP_URL || "http://127.0.0.1:4175/";
const STRESS = path.join(__dirname, "stress-test-workbench.xlsx");
// SLEEP_SCALE (run-tests.mjs): skraca przerwy „odczekaj po kroku” w szybkim przebiegu;
// powtórka porażki idzie ze skalą 1 (pełne przerwy). sleepFixed = przerwa, która MUSI
// przeczekać timer aplikacji (debounce, opóźnienie podpowiedzi, bezczynność) — bez skali.
const SLEEP_SCALE = Math.min(1, Math.max(0.1, Number(process.env.SLEEP_SCALE || 1)));
const sleep = (ms) => new Promise((r) => setTimeout(r, Math.round(ms * SLEEP_SCALE)));
const sleepFixed = (ms) => new Promise((r) => setTimeout(r, ms)); // eslint-disable-line no-unused-vars

const state = (page) => page.evaluate(() => window.__transcribe.state());
const fieldByLabel = (st, label) => st.fields.find((f) => f.label === label) || null;

// Skoroszyt: kolumna „Nazwisko” scalona pionowo (A2:A4 i A5:A6), reszta wypełniona
// normalnie. „Uwagi” celowo ma wartość tylko w jednym wierszu — służy do sprawdzenia,
// że kolumna NIE wskazana do dziedziczenia niczego nie przenosi.
async function buildFixture(page) {
  return page.evaluate(() => {
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet([
      ["Nazwisko", "Sprawa", "Kwota", "Uwagi"],
      ["Kowalski", "A1", 100, null],
      [null, "A2", 200, "pilne"],
      [null, "A3", 300, null],
      ["Nowak", "B1", 400, null],
      [null, "B2", 500, null],
      [null, "C1", 600, null], // wiersz 7: POZA jakimkolwiek scaleniem
    ]);
    ws["!merges"] = [
      { s: { r: 1, c: 0 }, e: { r: 3, c: 0 } },
      { s: { r: 4, c: 0 }, e: { r: 5, c: 0 } },
    ];
    XLSX.utils.book_append_sheet(wb, ws, "Sprawy");
    return XLSX.write(wb, { type: "base64", bookType: "xlsx" });
  });
}

async function loadFixture(page) {
  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  await page.evaluate(() => { try { ensureXlsxLibs && ensureXlsxLibs(false); } catch {} });
  await page.waitForFunction(() => typeof XLSX !== "undefined" && XLSX.utils, null, { timeout: 15000 });
  const b64 = await buildFixture(page);
  await page.setInputFiles("#fileInput", {
    name: "scalone-fixture.xlsx",
    mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    buffer: Buffer.from(b64, "base64"),
  });
  await page.waitForFunction(() => document.getElementById("sheetSelect")?.options?.length > 0, null, { timeout: 15000 });
  await sleep(200);
  await page.click("#loadBtn");
  await sleep(700);
  await page.evaluate(() => { try { setSidebarOpen && setSidebarOpen(false); } catch {} document.documentElement.classList.remove("sidebar-open"); });
  await sleep(120);
}

async function run() {
  const browser = await chromium.launch({ headless: true });
  const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1200, height: 900 } });
  await context.addInitScript(() => {
    localStorage.setItem("introPlayed", "true");
    if (!sessionStorage.getItem("trInhTestInit")) {
      localStorage.removeItem("excel-workbench-transcribe");
      sessionStorage.setItem("trInhTestInit", "1");
    }
  });
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push("pageerror: " + e.message));
  page.on("console", (m) => { if (m.type() === "error") errors.push("console: " + m.text()); });

  const failures = [];
  const report = {};

  await loadFixture(page);
  const headers = await page.evaluate(() => currentHeaders.slice());
  report.headers = headers;
  if (headers.join(",") !== "Nazwisko,Sprawa,Kwota,Uwagi") {
    failures.push(`fixture: oczekiwano nagłówków Nazwisko,Sprawa,Kwota,Uwagi — jest ${headers.join(",")}`);
  }

  // ── 1. Domyślnie wyłączone + wykryte scalenia ─────────────────────────────
  await page.click("#transcribeBtn");
  await sleep(300);
  const start = await state(page);
  report.start = { inherit: start.inherit, mergeCols: start.mergeCols, longMode: start.longMode, rows: start.rows };
  if (start.inherit) failures.push("dziedziczenie ma być domyślnie WYŁĄCZONE");
  if (start.mergeCols.join(",") !== "0") failures.push(`scalenia pionowe powinny wskazać kolumnę 0, wskazały [${start.mergeCols}]`);
  if (start.rows !== 6) failures.push(`fixture powinien dać 6 wierszy danych, jest ${start.rows}`);

  // Wiersz 3 arkusza (kontynuacja Kowalskiego) — bez dziedziczenia Nazwisko jest puste.
  await page.evaluate(() => window.__transcribe.go(1));
  await sleep(120);
  const contPlain = await state(page);
  const nazwiskoPlain = fieldByLabel(contPlain, "Nazwisko");
  report.contPlain = { fields: contPlain.fields };
  if (!nazwiskoPlain || nazwiskoPlain.text !== "") failures.push(`bez dziedziczenia Nazwisko ma być puste, jest "${nazwiskoPlain?.text}"`);

  // ── 2. Tryb auto BEZ dziedziczenia gubi pole tożsamości (to naprawiamy) ───
  await page.evaluate(() => window.__transcribe.setAutoFields(true));
  await sleep(150);
  const autoNoInherit = await state(page);
  report.autoNoInherit = autoNoInherit.fields.map((f) => f.label);
  if (fieldByLabel(autoNoInherit, "Nazwisko")) {
    failures.push("kontrola założeń: w trybie auto bez dziedziczenia Nazwisko NIE powinno się pojawić");
  }

  // ── 3. Włączenie dziedziczenia — kolumny ze scaleń same się zaznaczają ────
  await page.evaluate(() => window.__transcribe.setInherit(true));
  await sleep(200);
  const inherited = await state(page);
  const nazwiskoInh = fieldByLabel(inherited, "Nazwisko");
  report.inherited = { inheritCols: inherited.inheritCols, nazwisko: nazwiskoInh };
  if (inherited.inheritCols.join(",") !== "0") failures.push(`po włączeniu powinna zaznaczyć się kolumna 0, jest [${inherited.inheritCols}]`);
  if (!nazwiskoInh) failures.push("w trybie auto z dziedziczeniem Nazwisko powinno WRÓCIĆ na kartę");
  if (nazwiskoInh && nazwiskoInh.text !== "Kowalski") failures.push(`odziedziczono "${nazwiskoInh.text}", oczekiwano "Kowalski"`);
  if (nazwiskoInh && nazwiskoInh.from !== 2) failures.push(`źródłem ma być wiersz 2, wskazano ${nazwiskoInh?.from}`);

  // Znacznik na karcie musi być widoczny — inaczej wartość przepisze się jak własną.
  const badge = await page.evaluate(() => ({
    marked: document.querySelectorAll("#trCard .tr-field.is-inherited").length,
    text: document.querySelector("#trCard .tr-inherited-from")?.textContent || "",
  }));
  report.badge = badge;
  if (badge.marked !== 1) failures.push(`dokładnie jedno pole ma być oznaczone jako odziedziczone, jest ${badge.marked}`);
  if (!/2/.test(badge.text)) failures.push(`plakietka powinna podawać numer wiersza źródłowego, jest "${badge.text}"`);

  // ── 4. Nowy rekord przestawia źródło (Kowalski → Nowak) ───────────────────
  await page.evaluate(() => window.__transcribe.go(3)); // wiersz arkusza 6 = kontynuacja Nowaka
  await sleep(150);
  const nowak = await state(page);
  const nazwiskoNowak = fieldByLabel(nowak, "Nazwisko");
  report.nowak = nazwiskoNowak;
  if (!nazwiskoNowak || nazwiskoNowak.text !== "Nowak") failures.push(`kontynuacja drugiego rekordu ma dziedziczyć "Nowak", jest "${nazwiskoNowak?.text}"`);
  if (nazwiskoNowak && nazwiskoNowak.from !== 5) failures.push(`źródłem ma być wiersz 5, wskazano ${nazwiskoNowak?.from}`);

  // ── 4b. GRANICA SCALENIA: wiersz spoza scaleń nie dziedziczy niczego ──────
  // Wiersz 7 leży poza A2:A4 i A5:A6, więc „Nowak” NIE ma prawa się tam przelać.
  await page.evaluate(() => window.__transcribe.go(1)); // wiersz arkusza 7
  await sleep(150);
  const outside = await state(page);
  report.outsideMerge = { sprawa: fieldByLabel(outside, "Sprawa")?.text, nazwisko: fieldByLabel(outside, "Nazwisko") };
  if (fieldByLabel(outside, "Sprawa")?.text !== "C1") failures.push("test stoi na złym wierszu (oczekiwano C1)");
  if (fieldByLabel(outside, "Nazwisko")) {
    failures.push(`wiersz poza scaleniem odziedziczył wartość: ${JSON.stringify(fieldByLabel(outside, "Nazwisko"))}`);
  }

  // ── 5. BEZPIECZNIK: kolumna niewskazana NIE przenosi wartości ─────────────
  // „Uwagi” ma „pilne” tylko w wierszu 3. Wiersz 4 (ten sam rekord) nie może tego
  // dostać, dopóki kolumna nie zostanie wskazana ręcznie.
  await page.evaluate(() => window.__transcribe.go(-3)); // wiersz arkusza 4
  await sleep(150);
  const noBleed = await state(page);
  if (fieldByLabel(noBleed, "Sprawa")?.text !== "A3") failures.push(`bezpiecznik stoi na złym wierszu: ${fieldByLabel(noBleed, "Sprawa")?.text}`);
  report.noBleed = { fields: noBleed.fields.map((f) => `${f.label}=${f.text}`) };
  if (fieldByLabel(noBleed, "Uwagi")) {
    failures.push(`kolumna spoza dziedziczenia przeniosła wartość: ${JSON.stringify(fieldByLabel(noBleed, "Uwagi"))}`);
  }
  // …ale wskazana ręcznie już tak.
  const uwagiIdx = headers.indexOf("Uwagi");
  await page.evaluate((i) => window.__transcribe.toggleInheritCol(i), uwagiIdx);
  await sleep(150);
  const withUwagi = await state(page);
  const uwagi = fieldByLabel(withUwagi, "Uwagi");
  report.manualCol = uwagi;
  if (!uwagi || uwagi.text !== "pilne") failures.push(`po ręcznym wskazaniu Uwagi powinny dziedziczyć "pilne", jest "${uwagi?.text}"`);
  if (uwagi && uwagi.from !== 3) failures.push(`Uwagi mają pochodzić z wiersza 3, wskazano ${uwagi?.from}`);

  // Kolumna ręczna niesie wartość aż do następnej własnej — to inna reguła niż przy
  // scaleniach i test ma to utrwalić świadomie, a nie przypadkiem.
  await page.evaluate(() => window.__transcribe.go(1)); // wiersz arkusza 5 (nowy rekord)
  await sleep(150);
  const manualCarry = await state(page);
  report.manualCarryRule = {
    sprawa: fieldByLabel(manualCarry, "Sprawa")?.text,
    uwagi: fieldByLabel(manualCarry, "Uwagi"),
    nazwisko: fieldByLabel(manualCarry, "Nazwisko"),
  };
  if (fieldByLabel(manualCarry, "Uwagi")?.text !== "pilne") {
    failures.push("kolumna wskazana ręcznie ma nieść wartość do następnej własnej (reguła świadomie luźniejsza niż scalenia)");
  }
  await page.evaluate(() => window.__transcribe.go(-1));
  await sleep(120);
  await page.evaluate((i) => window.__transcribe.toggleInheritCol(i), uwagiIdx);
  await sleep(120);

  // ── 6. Wynik nie zależy od filtra ani sortowania ──────────────────────────
  // Odfiltrowujemy wiersz-kotwicę („Kowalski”), zostawiając samą kontynuację „A2”.
  // Gdyby dziedziczenie liczyło się po widoku, wartość by zniknęła albo się przesunęła.
  await page.keyboard.press("Escape");
  await sleep(200);
  // Pasek szybkiego szukania jest zawsze widoczny (rama ekranu, 2026-09-27).
  await page.evaluate(() => {
    document.getElementById("quickSearchAction").value = "filter";
    document.getElementById("quickSearch").value = "A2";
  });
  await page.click("#quickSearchBtn");
  await sleep(600);
  const filteredRows = await page.evaluate(() => viewRows.length);
  await page.click("#transcribeBtn");
  await sleep(300);
  const filtered = await state(page);
  const nazwiskoFiltered = fieldByLabel(filtered, "Nazwisko");
  report.filtered = { filteredRows, rows: filtered.rows, nazwisko: nazwiskoFiltered };
  if (filteredRows !== 1) failures.push(`filtr „A2" powinien zostawić 1 wiersz, zostawił ${filteredRows}`);
  if (!nazwiskoFiltered || nazwiskoFiltered.text !== "Kowalski") {
    failures.push(`po odfiltrowaniu wiersza-kotwicy dziedziczenie ma nadal dawać "Kowalski", daje "${nazwiskoFiltered?.text}" — liczone po widoku zamiast po baseRows?`);
  }
  if (nazwiskoFiltered && nazwiskoFiltered.from !== 2) failures.push(`źródło po filtrze ma być wierszem 2, jest ${nazwiskoFiltered?.from}`);

  // ── 7. Trwałość ───────────────────────────────────────────────────────────
  await page.keyboard.press("Escape");
  await sleep(200);
  await loadFixture(page);
  await page.click("#transcribeBtn");
  await sleep(350);
  const restored = await state(page);
  report.restored = { inherit: restored.inherit, inheritCols: restored.inheritCols, auto: restored.auto };
  if (!restored.inherit) failures.push("dziedziczenie powinno przeżyć przeładowanie");
  if (restored.inheritCols.join(",") !== "0") failures.push(`po przeładowaniu kolumny dziedziczące = [${restored.inheritCols}], oczekiwano [0]`);
  await page.keyboard.press("Escape");
  await sleep(200);

  // ── 8. Wide-to-Long: opcja świadomie zablokowana ──────────────────────────
  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  await page.evaluate(() => { try { ensureXlsxLibs && ensureXlsxLibs(false); } catch {} });
  await page.setInputFiles("#fileInput", STRESS);
  await page.waitForFunction(() => document.getElementById("sheetSelect")?.options?.length > 0, null, { timeout: 20000 });
  await sleep(300);
  await page.click("#loadBtn");
  await sleep(1200);
  await page.evaluate(() => { try { setSidebarOpen && setSidebarOpen(false); } catch {} document.documentElement.classList.remove("sidebar-open"); });
  const toggleVisible = await page.evaluate(() => {
    const el = document.getElementById("wideLongToggle");
    return !!el && !el.classList.contains("hidden");
  });
  report.longMode = { toggleVisible };
  if (toggleVisible) {
    await page.click("#wideLongToggle");
    await sleep(900);
    await page.click("#transcribeBtn");
    await sleep(400);
    const longState = await state(page);
    const disabled = await page.evaluate(() => ({
      cb: document.getElementById("trInherit").disabled,
      note: document.getElementById("trInheritNote").textContent,
    }));
    report.longMode = { toggleVisible, mode: longState.longMode, inherit: longState.inherit, disabled };
    if (!longState.longMode) failures.push("po przełączeniu na Wide-to-Long stan powinien raportować longMode=true");
    if (longState.inherit) failures.push("w Wide-to-Long dziedziczenie musi być wyłączone");
    if (!disabled.cb) failures.push("w Wide-to-Long przełącznik dziedziczenia ma być zablokowany");
    if (!disabled.note) failures.push("brak notki wyjaśniającej ograniczenie w Wide-to-Long");
    await page.keyboard.press("Escape");
    await sleep(150);
  }

  await browser.close();

  if (errors.length) failures.push(`console/page errors: ${errors.join(" | ")}`);
  console.log(JSON.stringify(report, null, 2));
  if (failures.length) throw new Error(failures.join("; "));
  console.log("✅ transcribe-inherit-playwright OK");
}

run().catch((error) => {
  console.error(error);
  process.exit(1);
});
