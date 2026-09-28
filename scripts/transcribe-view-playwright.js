// transcribe-view-playwright.js — spisywanie pamięta WIDOK (filtr + sortowanie),
// przy którym odhaczano, i pilnuje, żeby lista wierszy po powrocie była tą samą listą.
//
// Scenariusz z życia (zgłoszony przez Mateusza): spisuje przy filtrze, potem lekko
// poprawia plik i zapisuje go jako „…_V2”. W nowej wersji zapomina ustawić ten sam
// filtr → lista wierszy jest inna, a nic o tym nie mówi.
//
// Sprawdzamy:
//   1. zapis spisywania niesie widok (zapytanie, kolumny zapytania, sortowanie),
//   2. inny widok przy powrocie → baner „wtedy / teraz”, licznik liczy tylko bieżącą listę,
//   3. „Przywróć tamten widok” ustawia filtr i sortowanie w tabeli i otwiera listę na nowo,
//   4. „Zostaw jak jest” → nie pyta drugi raz,
//   5. ✓ poza bieżącą listą są policzone (także w starym zapisie bez widoku — bez przycisku),
//   6. nowa wersja pliku (inna nazwa): propozycja przeniesienia mówi o widoku, po
//      przeniesieniu baner proponuje przywrócenie,
//   7. brak kolumny z filtra → NIC nie zmieniamy i mówimy, której brakuje,
//   8. przywrócony widok dający 0 wierszy → wracamy do poprzedniego.

const { chromium } = require("playwright");
const APP_URL = process.env.APP_URL || "http://127.0.0.1:4175/";
const SLEEP_SCALE = Math.min(1, Math.max(0.1, Number(process.env.SLEEP_SCALE || 1)));
const sleep = (ms) => new Promise((r) => setTimeout(r, Math.round(ms * SLEEP_SCALE)));
const state = (page) => page.evaluate(() => window.__transcribe.state());
const STORE = "excel-workbench-transcribe";

const PEOPLE = [
  ["Ala", "W toku"], ["Bartek", "Zakończone"], ["Cezary", "W toku"],
  ["Dorota", "Zakończone"], ["Edward", "W toku"], ["Franek", "Zakończone"],
];

async function loadSheet(page, { file = "widok.xlsx", statusHeader = "Status" } = {}) {
  await page.evaluate(({ people, file, statusHeader }) => {
    const ws = {};
    ws.A1 = { t: "s", v: "Osoba" };
    ws.B1 = { t: "s", v: statusHeader };
    ws.C1 = { t: "s", v: "Miasto" };
    people.forEach(([n, st], i) => {
      const r = i + 2;
      ws["A" + r] = { t: "s", v: n };
      ws["B" + r] = { t: "s", v: st };
      ws["C" + r] = { t: "s", v: "Miasto " + n };
    });
    ws["!ref"] = "A1:C" + (people.length + 1);
    workbook = { SheetNames: ["Arkusz"], Sheets: { Arkusz: ws }, Props: { ModifiedDate: "2026-09-28T12:00:00.000Z" } };
    currentFileName = file;
    sheetSelect.replaceChildren();
    const opt = document.createElement("option"); opt.value = "Arkusz"; opt.textContent = "Arkusz";
    sheetSelect.appendChild(opt); sheetSelect.value = "Arkusz";
    document.getElementById("headerRow").value = "1";
    document.getElementById("autoHeaderRow").checked = false;
  }, { people: PEOPLE, file, statusHeader });
  await page.evaluate(() => document.getElementById("loadBtn").click());
  await page.waitForFunction((n) => typeof baseRows !== "undefined" && baseRows.length === n && typeof currentFileName !== "undefined", PEOPLE.length, { timeout: 15000 });
  await page.waitForFunction((f) => currentFileName === f, file);
  await sleep(350);
}

// Filtr ustawiony tak, jak robi to użytkownik: szybkie szukanie z operatorami + „Filtruj”,
// do tego sortowanie malejąco po osobie.
async function setFilter(page, query, sortDesc = true) {
  await page.evaluate(({ query, sortDesc }) => {
    resetFilterInputs();
    searchQueryEl.value = query;
    filterOperatorsEl.checked = true;
    syncQuickSearchInputs();
    syncQuickSearchOperatorsControls();
    filtersCommitted = true;
    multiSortState = sortDesc ? [{ col: "Osoba", dir: "desc" }] : [];
    applyFilters();
    sortRows();
    scheduleViewRefresh({ table: true, analyses: true, filterBadge: true, immediate: true });
  }, { query, sortDesc });
}

async function clearFilter(page) {
  await page.evaluate(() => {
    resetFilterInputs();
    multiSortState = [];
    applyFilters();
    sortRows();
    scheduleViewRefresh({ table: true, analyses: true, filterBadge: true, immediate: true });
  });
}

const reopen = async (page) => {
  await page.evaluate(() => window.__transcribe.close());
  await sleep(150);
  await page.evaluate(() => window.__transcribe.open());
  await sleep(250);
};
const banner = (page) => page.evaluate(() => ({
  shown: !document.getElementById("trView").classList.contains("hidden"),
  text: document.getElementById("trViewText").textContent,
  restore: !document.getElementById("trViewRestoreBtn").classList.contains("hidden"),
}));
const tableView = (page) => page.evaluate(() => ({
  filtering: !!(lastAppliedFilters && lastAppliedFilters.filtering),
  query: searchQueryEl.value,
  rows: viewRows.length,
  sort: multiSortState.map((r) => `${r.col}:${r.dir}`).join(","),
}));
const lastToast = (page) => page.evaluate(() => {
  const all = Array.from(document.querySelectorAll(".toast"));
  return all.length ? all[all.length - 1].textContent : "";
});
const cardName = (page) => page.evaluate(() => window.__transcribe.state().values[0] || "");

async function run() {
  const browser = await chromium.launch({ headless: true });
  const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1100, height: 800 } });
  await context.addInitScript(() => {
    localStorage.setItem("introPlayed", "true");
    // czyścimy TYLKO przy pierwszym wejściu — addInitScript odpala się przy każdej nawigacji
    if (!sessionStorage.getItem("trViewInit")) {
      sessionStorage.setItem("trViewInit", "1");
      localStorage.removeItem("excel-workbench-transcribe");
    }
  });
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push("pageerror: " + e.message));
  page.on("console", (m) => { if (m.type() === "error") errors.push("console: " + m.text()); });

  const results = [];
  const check = (name, cond, got) => {
    results.push({ name, ok: !!cond });
    console.log(`${cond ? "✅" : "❌"} ${name}${cond ? "" : ` — dostałem: ${JSON.stringify(got)}`}`);
  };

  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  await page.evaluate(() => ensureXlsxLibs(false));
  await page.waitForFunction(() => typeof buildRows === "function" && window.__transcribe && typeof captureViewState === "function");

  // ── 1. Spisywanie przy filtrze „W toku”, sort malejąco ─────────────────────
  await loadSheet(page);
  await setFilter(page, 'Status:="W toku"');
  const tv1 = await tableView(page);
  check("filtr zawęża tabelę do 3 wierszy", tv1.filtering && tv1.rows === 3, tv1);
  await page.evaluate(() => window.__transcribe.open());
  await sleep(250);
  check("pierwsza karta = Edward (sort malejąco)", (await cardName(page)) === "Edward", await cardName(page));
  check("pierwsze otwarcie: brak banera widoku", !(await banner(page)).shown, await banner(page));
  await page.evaluate(() => { window.__transcribe.mark(); window.__transcribe.mark(); });
  await page.evaluate(() => window.__transcribe.close());
  await sleep(150);
  const rec1 = await page.evaluate((k) => Object.values(JSON.parse(localStorage.getItem(k)).scopes)[0], STORE);
  check("zapis niesie widok z zapytaniem", rec1.view && rec1.view.filter && rec1.view.filter.f1.q === 'Status:="W toku"', rec1.view);
  check("zapis zna kolumnę użytą w zapytaniu", JSON.stringify(rec1.view?.filter?.f1?.qcols) === '["Status"]', rec1.view?.filter?.f1);
  check("zapis niesie sortowanie", JSON.stringify(rec1.view?.sort) === '[{"col":"Osoba","dir":"desc"}]', rec1.view?.sort);

  // ── 2. Powrót bez filtra → baner „wtedy / teraz” ────────────────────────────
  await clearFilter(page);
  await page.evaluate(() => window.__transcribe.open());
  await sleep(250);
  const b2 = await banner(page);
  const s2 = await state(page);
  check("baner widoku widoczny", b2.shown, b2);
  check("baner mówi, jaki był filtr (Status = W toku)", /Status = W toku/.test(b2.text), b2.text);
  check("baner mówi, że teraz bez filtra", /bez filtra/.test(b2.text), b2.text);
  check("baner podaje dawne sortowanie", /Osoba ↓/.test(b2.text), b2.text);
  check("przycisk przywracania widoczny", b2.restore, b2);
  check("lista = cały arkusz (6), ✓ wciąż 2 i obie na liście", s2.rows === 6 && s2.done === 2 && s2.doneInView === 2, s2);

  // ── 3. „Przywróć tamten widok” ───────────────────────────────────────────────
  await page.click("#trViewRestoreBtn");
  await sleep(350);
  const s3 = await state(page);
  const tv3 = await tableView(page);
  check("po przywróceniu spisywanie jest otwarte na 3 wierszach", s3.open && s3.rows === 3, s3);
  check("tabela ma z powrotem filtr i sortowanie", tv3.filtering && tv3.query === 'Status:="W toku"' && tv3.sort === "Osoba:desc", tv3);
  check("✓ zachowane (2)", s3.done === 2 && s3.doneInView === 2, s3);
  check("baner zniknął (widoki zgodne)", !(await banner(page)).shown, await banner(page));
  check("karta na pierwszym nieodhaczonym (Ala)", (await cardName(page)) === "Ala", await cardName(page));
  check("toast o przywróceniu", /Przywrócono widok/.test(await lastToast(page)), await lastToast(page));

  // ── 4. „Zostaw jak jest” → nie pyta drugi raz ────────────────────────────────
  await page.evaluate(() => window.__transcribe.close());
  await clearFilter(page);
  await page.evaluate(() => window.__transcribe.open());
  await sleep(250);
  check("znowu inny widok → baner", (await banner(page)).shown);
  await page.click("#trViewKeepBtn");
  await sleep(100);
  check("„Zostaw” chowa baner", !(await banner(page)).shown);
  await reopen(page);
  check("po „Zostaw” kolejne otwarcie nie pyta", !(await banner(page)).shown, await banner(page));

  // ── 5. ✓ poza bieżącą listą ─────────────────────────────────────────────────
  await page.evaluate(() => window.__transcribe.close());
  await setFilter(page, 'Status:="Zakończone"', false);
  await page.evaluate(() => window.__transcribe.open());
  await sleep(250);
  const b5 = await banner(page);
  const s5 = await state(page);
  check("inny filtr: licznik liczy tylko bieżącą listę (0 z 3)", s5.rows === 3 && s5.doneInView === 0 && s5.outside === 2, s5);
  check("baner mówi o 2 z 2 spisanych poza listą", /2 z 2/.test(b5.text), b5.text);
  const counter = await page.evaluate(() => document.getElementById("trCounter")?.textContent + " | " + (document.getElementById("trDoneCount")?.textContent || ""));
  check("licznik spisanych nie zawyża (0)", !/\b2 \//.test(counter.split("|")[1] || ""), counter);

  // stary zapis (sprzed tej funkcji) — bez widoku: sam komunikat, bez przywracania
  await page.evaluate(() => window.__transcribe.close());
  await page.evaluate((k) => {
    const st = JSON.parse(localStorage.getItem(k));
    Object.values(st.scopes).forEach((r) => { delete r.view; });
    localStorage.setItem(k, JSON.stringify(st));
  }, STORE);
  await page.evaluate(() => window.__transcribe.open());
  await sleep(250);
  const b5b = await banner(page);
  check("stary zapis: baner o ✓ poza listą, z pytaniem o filtr", b5b.shown && /inny filtr/.test(b5b.text), b5b);
  check("stary zapis: bez przycisku przywracania (nie ma czego)", !b5b.restore, b5b);
  await page.click("#trViewKeepBtn");
  await page.evaluate(() => window.__transcribe.close());

  // ── 6. Nowa wersja pliku (…_V2) ─────────────────────────────────────────────
  // Świeży zapis w V1: ✓ przy filtrze W toku.
  await page.evaluate((k) => localStorage.removeItem(k), STORE);
  await setFilter(page, 'Status:="W toku"');
  await page.evaluate(() => window.__transcribe.open());
  await sleep(200);
  await page.evaluate(() => { window.__transcribe.mark(); window.__transcribe.mark(); });
  await page.evaluate(() => window.__transcribe.close());
  await loadSheet(page, { file: "widok_V2.xlsx" });
  await clearFilter(page);
  await page.evaluate(() => window.__transcribe.open());
  await sleep(300);
  const sug = await page.evaluate(() => ({
    shown: !document.getElementById("trSuggest").classList.contains("hidden"),
    text: document.getElementById("trSuggestText").textContent,
  }));
  check("V2: propozycja przeniesienia", sug.shown, sug);
  check("V2: propozycja mówi o tamtym widoku", /Status = W toku/.test(sug.text), sug.text);
  await page.click("#trSuggestYesBtn");
  await sleep(250);
  const b6 = await banner(page);
  check("V2: po przeniesieniu baner proponuje przywrócić widok", b6.shown && b6.restore && /Status = W toku/.test(b6.text), b6);
  await page.click("#trViewRestoreBtn");
  await sleep(350);
  const s6 = await state(page);
  const tv6 = await tableView(page);
  check("V2: po przywróceniu 3 wiersze, 2 ✓ na liście", s6.rows === 3 && s6.done === 2 && s6.doneInView === 2, s6);
  check("V2: tabela przefiltrowana jak w V1", tv6.filtering && tv6.rows === 3, tv6);
  await page.evaluate(() => window.__transcribe.close());

  // ── 7. Brak kolumny z filtra → nic nie zmieniamy ────────────────────────────
  await loadSheet(page, { file: "widok_V3.xlsx", statusHeader: "Stan" });
  await clearFilter(page);
  await page.evaluate(() => window.__transcribe.open());
  await sleep(300);
  // Tu nie polegamy na propozycji (kolumny się różnią) — przenosimy ręcznie z listy.
  await page.evaluate(() => {
    const key = Object.keys(JSON.parse(localStorage.getItem("excel-workbench-transcribe")).scopes)
      .find((k) => k.startsWith("widok_V2.xlsx"));
    window.__transcribe.importFrom(key);
  });
  await sleep(250);
  const b7 = await banner(page);
  check("V3: baner proponuje tamten widok", b7.shown && b7.restore, b7);
  await page.click("#trViewRestoreBtn");
  await sleep(300);
  const tv7 = await tableView(page);
  const t7 = await lastToast(page);
  check("V3: brak kolumny → tabela nietknięta (bez filtra)", !tv7.filtering && tv7.rows === 6 && tv7.query === "", tv7);
  check("V3: komunikat nazywa brakującą kolumnę", /„Status”/.test(t7) && /nic nie zmieniłem/.test(t7), t7);
  check("V3: spisywanie dalej otwarte, baner został", (await state(page)).open && (await banner(page)).shown);
  await page.evaluate(() => window.__transcribe.close());

  // ── 8. Przywrócony widok daje 0 wierszy → wracamy do poprzedniego ───────────
  await loadSheet(page, { file: "widok.xlsx" });
  await clearFilter(page);
  await page.evaluate((k) => {
    const st = JSON.parse(localStorage.getItem(k));
    const key = Object.keys(st.scopes).find((x) => x.startsWith("widok.xlsx"));
    st.scopes[key].view = {
      v: 1,
      filter: { f1: { q: 'Status:="Wstrzymane"', mode: "contains", neg: false, empty: "all", ops: true, cols: [], qcols: ["Status"] } },
      sort: [],
      shown: 2,
      total: 6,
    };
    localStorage.setItem(k, JSON.stringify(st));
  }, STORE);
  await page.evaluate(() => window.__transcribe.open());
  await sleep(250);
  check("pusty widok: baner jest", (await banner(page)).shown);
  await page.click("#trViewRestoreBtn");
  await sleep(300);
  const tv8 = await tableView(page);
  check("pusty widok: tabela wróciła do poprzedniego (6 wierszy, bez filtra)", !tv8.filtering && tv8.rows === 6 && tv8.query === "", tv8);
  check("pusty widok: komunikat ostrzega", /żadnego wiersza/.test(await lastToast(page)), await lastToast(page));
  await page.evaluate(() => window.__transcribe.close());

  check("brak błędów w konsoli", errors.length === 0, errors);
  await browser.close();

  const failed = results.filter((r) => !r.ok);
  if (failed.length) {
    console.log(`\n❌ transcribe-view: ${failed.length} z ${results.length} nie przeszło`);
    process.exit(1);
  }
  console.log(`\n✅ transcribe-view: ${results.length}/${results.length}`);
}

run().catch((e) => { console.error(e); process.exit(1); });
