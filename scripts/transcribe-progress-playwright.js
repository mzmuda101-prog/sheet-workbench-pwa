// transcribe-progress-playwright.js — PAMIĘĆ postępu spisywania: co apka pamięta,
// jak sobie radzi ze ZMIENIONYM plikiem i jak to wyczyścić.
//
// Scenariusz z życia (zgłoszony przez Mateusza): ktoś spisał połowę, zamknął apkę,
// wrócił po dniu — ale plik w międzyczasie się zmienił (doszły wiersze). Klucz ✓ to
// POZYCJA wiersza, więc bez zabezpieczenia wszystkie odhaczenia przesuwają się na
// cudze wiersze i wiersz nieprzepisany wygląda na zrobiony.
//
// Sprawdzamy:
//   1. panel „Postęp" pokazuje spisane / zostało / % i datę ostatniej sesji,
//   2. po WSTAWIENIU wierszy na górze ✓ trafiają na TE SAME wiersze co przed zmianą
//      (dopasowanie po treści), a użytkownik dostaje o tym baner z liczbami,
//   3. „Zacznij od nowa" z banera czyści odhaczenia,
//   4. lista zapamiętanych spisywań pozwala usunąć pojedynczy wpis i wszystkie naraz,
//   5. samo przeliczenie formuł „na dziś" (TODAY) NIE udaje zmiany pliku.

const { chromium } = require("playwright");
const APP_URL = process.env.APP_URL || "http://127.0.0.1:4175/";
// SLEEP_SCALE (run-tests.mjs): skraca przerwy „odczekaj po kroku” w szybkim przebiegu;
// powtórka porażki idzie ze skalą 1 (pełne przerwy). sleepFixed = przerwa, która MUSI
// przeczekać timer aplikacji (debounce, opóźnienie podpowiedzi, bezczynność) — bez skali.
const SLEEP_SCALE = Math.min(1, Math.max(0.1, Number(process.env.SLEEP_SCALE || 1)));
const sleep = (ms) => new Promise((r) => setTimeout(r, Math.round(ms * SLEEP_SCALE)));
const sleepFixed = (ms) => new Promise((r) => setTimeout(r, ms)); // eslint-disable-line no-unused-vars
const state = (page) => page.evaluate(() => window.__transcribe.state());

// Buduje arkusz w pamięci przeglądarki: `extra` wierszy wstawionych NA GÓRZE
// przed pierwotnymi 6 wierszami (Ala..Franek) => każdy stary wiersz zmienia pozycję.
async function loadSheet(page, { extra = 0, volatile: vol = false } = {}) {
  await page.evaluate(({ extra, vol }) => {
    const names = [];
    for (let i = 0; i < extra; i++) names.push("Nowy " + (i + 1));
    names.push("Ala", "Bartek", "Cezary", "Dorota", "Edward", "Franek");
    const ws = {};
    ws.A1 = { t: "s", v: "Osoba" };
    ws.B1 = { t: "s", v: "Miasto" };
    ws.C1 = { t: "s", v: "Dni" };
    names.forEach((n, i) => {
      const r = i + 2;
      ws["A" + r] = { t: "s", v: n };
      ws["B" + r] = { t: "s", v: "Miasto " + n };
      // Kolumna liczona „na dziś" — w wariancie vol formuła z TODAY() ma NIEAKTUALNY wynik
      // w pliku, więc apka ją przeliczy i wartość będzie inna niż zapisana.
      if (vol) ws["C" + r] = { t: "n", v: 1, w: "1", f: "TODAY()-1" };
      // wartość ZWIĄZANA Z WIERSZEM, nie z jego pozycją — inaczej test sprawdzałby
      // co innego, niż deklaruje (wiersz przesunięty to nadal ten sam wiersz)
      else ws["C" + r] = { t: "n", v: n.length, w: String(n.length) };
    });
    ws["!ref"] = "A1:C" + (names.length + 1);
    workbook = { SheetNames: ["Arkusz"], Sheets: { Arkusz: ws }, Props: { ModifiedDate: "2026-06-06T12:00:00.000Z" } };
    currentFileName = "postep.xlsx";
    sheetSelect.replaceChildren();
    const opt = document.createElement("option"); opt.value = "Arkusz"; opt.textContent = "Arkusz";
    sheetSelect.appendChild(opt); sheetSelect.value = "Arkusz";
    document.getElementById("headerRow").value = "1";
    document.getElementById("autoHeaderRow").checked = false;
  }, { extra, vol });
  // klik przez JS: przy węższym oknie sidebar bywa zwinięty, a nam chodzi o logikę, nie o piksele
  await page.evaluate(() => document.getElementById("loadBtn").click());
  await page.waitForFunction((n) => typeof baseRows !== "undefined" && baseRows.length === n, 6 + extra, { timeout: 15000 });
  await sleep(350);
}

// Które NAZWY są odhaczone — jedyna miara, która ma sens po przesunięciu wierszy.
const doneNames = (page) => page.evaluate(() => {
  const st = window.__transcribe;
  return baseRows.filter((r) => st.isDone(getRowSelectionKey(r))).map((r) => String(getDisplayValue(r, 0)));
});

async function run() {
  const browser = await chromium.launch({ headless: true });
  const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1100, height: 800 } });
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
  const ok = (name, cond, got) => { if (!cond) failures.push(`${name} — dostalem: ${JSON.stringify(got)}`); };

  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  await page.evaluate(() => ensureXlsxLibs(false));
  await page.waitForFunction(() => typeof buildRows === "function" && window.__transcribe);

  // ── 1. Spisz połowę (Ala, Bartek, Cezary) ────────────────────────────────
  await loadSheet(page, { extra: 0 });
  await page.evaluate(() => window.__transcribe.open());
  await sleep(250);
  await page.evaluate(() => { window.__transcribe.mark(); window.__transcribe.mark(); window.__transcribe.mark(); });
  await sleep(200);
  const half = await state(page);
  const halfNames = await doneNames(page);
  report.half = { done: half.done, names: halfNames };
  ok("odhaczone 3 pierwsze wiersze", half.done === 3 && halfNames.join(",") === "Ala,Bartek,Cezary", report.half);

  // Panel „Postęp": liczby muszą się zgadzać z rzeczywistością
  await page.evaluate(() => window.__transcribe.openProgress());
  await sleep(200);
  const stats = await page.evaluate(() => ({
    tiles: Array.from(document.querySelectorAll("#trStats .tr-stat")).map((el) => el.querySelector(".tr-stat-value").textContent),
    note: document.getElementById("trScopeNote").textContent,
    store: document.querySelectorAll("#trStoreList .tr-store-item").length,
  }));
  report.stats = stats;
  ok("kafelki: spisane 3, zostalo 3, 50%", stats.tiles.join("|") === "3|3|50%", stats.tiles);
  ok("notka podaje rozmiar arkusza", /6/.test(stats.note), stats.note);
  ok("notka podaje ostatnia sesje", /dziś|today/i.test(stats.note), stats.note);
  ok("lista zapamietanych ma 1 wpis", stats.store === 1, stats.store);
  await page.evaluate(() => window.__transcribe.closeProgress());
  await page.evaluate(() => window.__transcribe.close());
  await sleep(200);

  // ── 2. Plik się zmienił: 2 nowe wiersze NA GÓRZE ─────────────────────────
  await loadSheet(page, { extra: 2 });
  await page.evaluate(() => window.__transcribe.open());
  await sleep(300);
  const afterChange = await state(page);
  const afterNames = await doneNames(page);
  const notice = await page.evaluate(() => ({
    visible: !document.getElementById("trNotice").classList.contains("hidden"),
    text: document.getElementById("trNoticeText").textContent,
  }));
  report.changed = { done: afterChange.done, names: afterNames, notice };
  ok("✓ zostaly przy TYCH SAMYCH osobach", afterNames.join(",") === "Ala,Bartek,Cezary", report.changed);
  ok("liczba ✓ bez zmian", afterChange.done === 3, afterChange.done);
  ok("baner o zmianie jest widoczny", notice.visible, notice);
  ok("baner podaje liczby 6 -> 8 i 3 z 3", /6/.test(notice.text) && /8/.test(notice.text) && /3/.test(notice.text), notice.text);

  // ── 3. „Zacznij od nowa" z banera ────────────────────────────────────────
  await page.evaluate(() => document.getElementById("trNoticeResetBtn").click());
  await sleep(250);
  const afterReset = await state(page);
  const noticeGone = await page.evaluate(() => document.getElementById("trNotice").classList.contains("hidden"));
  report.reset = { done: afterReset.done, noticeGone };
  ok("po Zacznij-od-nowa zero znacznikow", afterReset.done === 0, afterReset.done);
  ok("baner znika po decyzji", noticeGone, noticeGone);

  // ── 4. Czyszczenie pamieci: pojedynczy wpis i wszystko ───────────────────
  await page.evaluate(() => { window.__transcribe.mark(); window.__transcribe.mark(); });
  await page.evaluate(() => window.__transcribe.close());
  await sleep(150);
  // drugi „plik" = drugi wpis w pamieci
  await page.evaluate(() => { currentFileName = "inny.xlsx"; });
  await page.evaluate(() => window.__transcribe.open());
  await sleep(200);
  await page.evaluate(() => window.__transcribe.mark());
  await page.evaluate(() => window.__transcribe.openProgress());
  await sleep(200);
  const twoScopes = await page.evaluate(() => document.querySelectorAll("#trStoreList .tr-store-item").length);
  report.scopes = { count: twoScopes };
  ok("pamiec trzyma osobne wpisy per plik", twoScopes === 2, twoScopes);

  await page.evaluate(() => document.querySelectorAll("#trStoreList .tr-store-del")[1].click());
  await sleep(250);
  const afterDelete = await page.evaluate(() => ({
    items: document.querySelectorAll("#trStoreList .tr-store-item").length,
    store: JSON.parse(localStorage.getItem("excel-workbench-transcribe") || "{}"),
  }));
  report.afterDelete = { items: afterDelete.items, scopes: Object.keys(afterDelete.store.scopes || {}).length };
  ok("usuniecie pojedynczego wpisu dziala", afterDelete.items === 1 && report.afterDelete.scopes === 1, report.afterDelete);

  // „Wyczysc wszystkie" jest dwustopniowe — pierwszy klik tylko uzbraja
  await page.evaluate(() => document.getElementById("trStoreClearAllBtn").click());
  await sleep(120);
  const armed = await page.evaluate(() => document.getElementById("trStoreClearAllBtn").classList.contains("is-armed"));
  await page.evaluate(() => document.getElementById("trStoreClearAllBtn").click());
  await sleep(250);
  const cleared = await page.evaluate(() => ({
    raw: localStorage.getItem("excel-workbench-transcribe"),
    done: window.__transcribe.state().done,
    items: document.querySelectorAll("#trStoreList .tr-store-item").length,
  }));
  report.clearAll = { armed, raw: cleared.raw, done: cleared.done, items: cleared.items };
  ok("pierwszy klik tylko uzbraja", armed, armed);
  ok("po wyczyszczeniu pamiec pusta", !cleared.raw || cleared.raw === "{}" || !Object.keys(JSON.parse(cleared.raw).scopes || {}).length, cleared.raw);
  ok("po wyczyszczeniu zero ✓ w biezacym arkuszu", cleared.done === 0, cleared.done);

  // ── 5. Przeliczanie „na dzis" nie udaje zmiany pliku ─────────────────────
  await page.evaluate(() => window.__transcribe.closeProgress());
  await page.evaluate(() => window.__transcribe.close());
  await page.evaluate(() => { currentFileName = "volatile.xlsx"; });
  await loadSheet(page, { extra: 0, volatile: true });
  await page.evaluate(() => window.__transcribe.open());
  await sleep(250);
  await page.evaluate(() => { window.__transcribe.mark(); window.__transcribe.mark(); });
  await page.evaluate(() => window.__transcribe.close());
  await sleep(150);
  await loadSheet(page, { extra: 0, volatile: true });
  await page.evaluate(() => window.__transcribe.open());
  await sleep(250);
  const volState = await state(page);
  const volNotice = await page.evaluate(() => !document.getElementById("trNotice").classList.contains("hidden"));
  const volCols = await page.evaluate(() => window.__transcribe.state().volatileCols);
  report.volatile = { done: volState.done, notice: volNotice, volCols };
  ok("kolumna z TODAY() rozpoznana jako zmienna", Array.isArray(volCols) && volCols.includes(2), volCols);
  ok("przeliczanie na dzis nie wywoluje ostrzezenia", !volNotice, volNotice);
  ok("✓ przezyly ponowne otwarcie", volState.done === 2, volState.done);

  console.log(JSON.stringify(report, null, 2));
  if (errors.length) failures.push("bledy strony: " + errors.join(" | "));
  await browser.close();
  if (failures.length) {
    console.error("\n❌ transcribe-progress-playwright FAIL:\n" + failures.map((f) => " - " + f).join("\n"));
    process.exit(1);
  }
  console.log("✅ transcribe-progress-playwright OK");
}

run().catch((e) => { console.error(e); process.exit(1); });
