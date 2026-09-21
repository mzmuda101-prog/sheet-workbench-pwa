// transcribe-match-playwright.js — KASKADA dopasowania ✓ po zmianie arkusza.
//
// Zgłoszenie z życia (Mateusz, 2026-09-21): po poprawkach w arkuszach „Przenieś" pokazało
// „81 z 141 przeniesionych, 60 nie znaleziono". Przyczyna: odcisk v1 był jednym hashem
// CAŁEGO wiersza, liczonym po POZYCJACH kolumn — więc poprawka jednej komórki albo dodanie
// kolumny kasowały dopasowanie, choć dla człowieka to oczywiście ten sam wiersz.
//
// Ten test pilnuje, że v2 wytrzymuje realne zmiany pliku i — równie ważne — że NIE
// odhacza cudzych wierszy, gdy dopasowanie jest niejednoznaczne. Przy przepisywaniu na
// papier fałszywe ✓ jest gorsze niż brak ✓: wiersz nieprzepisany wygląda na zrobiony.
//
// Sprawdzamy:
//   1. poprawka komórki w odhaczonym wierszu → ✓ zostaje (tura „po kluczu"),
//   2. DODANA kolumna → ✓ zostają (v1 gubiło tu wszystko),
//   3. PRZESTAWIONE kolumny → ✓ zostają (odcisk idzie po nazwach, nie po numerach),
//   4. usunięty wiersz → ląduje na liście nieodnalezionych z podglądem, a jego ✓ NIE
//      wędruje na inny wiersz,
//   5. „Przenieś" między plikami, gdy źródło ma kolumnę liczoną „na dziś", a cel nie
//      (to był błąd: odciski liczone z różnych zestawów kolumn nie miały prawa trafić).

const { chromium } = require("playwright");
const APP_URL = process.env.APP_URL || "http://127.0.0.1:4175/";
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
const state = (page) => page.evaluate(() => window.__transcribe.state());

// Arkusz: Nr | Osoba | Miasto | Uwagi  (+ opcjonalnie kolumna „na dziś" i zmiany układu)
async function loadSheet(page, opts = {}) {
  await page.evaluate((o) => {
    const people = o.people || ["Ala", "Bartek", "Cezary", "Dorota", "Edward", "Franek"];
    const cols = o.order || ["Nr", "Osoba", "Miasto", "Uwagi"];
    if (o.extraCol) cols.push("Status");
    const letter = (i) => String.fromCharCode(65 + i);
    const ws = {};
    cols.forEach((name, i) => { ws[letter(i) + "1"] = { t: "s", v: name }; });
    people.forEach((person, rowN) => {
      const r = rowN + 2;
      const values = {
        Nr: { t: "n", v: 1000 + (o.nrOf ? o.nrOf[person] : rowN + 1), w: String(1000 + (o.nrOf ? o.nrOf[person] : rowN + 1)) },
        Osoba: { t: "s", v: person },
        Miasto: { t: "s", v: "Miasto " + person },
        Uwagi: { t: "s", v: (o.editedNote && o.editedNote[person]) || ("uwaga " + person) },
        Status: { t: "s", v: "nowy" },
      };
      cols.forEach((name, i) => {
        if (name === "Dni" && o.volatile) ws[letter(i) + r] = { t: "n", v: 1, w: "1", f: "TODAY()-1" };
        else ws[letter(i) + r] = values[name];
      });
    });
    ws["!ref"] = "A1:" + letter(cols.length - 1) + (people.length + 1);
    workbook = { SheetNames: ["Arkusz"], Sheets: { Arkusz: ws }, Props: { ModifiedDate: o.mtime || "2026-06-06T12:00:00.000Z" } };
    currentFileName = o.file || "match.xlsx";
    sheetSelect.replaceChildren();
    const opt = document.createElement("option"); opt.value = "Arkusz"; opt.textContent = "Arkusz";
    sheetSelect.appendChild(opt); sheetSelect.value = "Arkusz";
    document.getElementById("headerRow").value = "1";
    document.getElementById("autoHeaderRow").checked = false;
  }, opts);
  await page.evaluate(() => document.getElementById("loadBtn").click());
  const n = (opts.people || ["Ala", "Bartek", "Cezary", "Dorota", "Edward", "Franek"]).length;
  await page.waitForFunction((k) => typeof baseRows !== "undefined" && baseRows.length === k, n, { timeout: 15000 });
  await sleep(300);
}

// Które OSOBY są odhaczone — jedyna miara odporna na przesuwanie wierszy.
const doneNames = (page) => page.evaluate(() => {
  const st = window.__transcribe;
  const col = (Array.isArray(currentHeaders) ? currentHeaders : []).indexOf("Osoba");
  return baseRows
    .filter((r) => st.isDone(getRowSelectionKey(r)))
    .map((r) => String(getDisplayValue(r, col < 0 ? 1 : col)))
    .sort();
});

async function reopen(page, opts) {
  await page.evaluate(() => window.__transcribe.close());
  await sleep(120);
  await loadSheet(page, opts);
  await page.evaluate(() => window.__transcribe.open());
  await sleep(350);
}

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

  // ── 0. Spisz Alę, Bartka i Cezarego ──────────────────────────────────────
  await loadSheet(page, {});
  await page.evaluate(() => window.__transcribe.open());
  await sleep(250);
  await page.evaluate(() => { window.__transcribe.mark(); window.__transcribe.mark(); window.__transcribe.mark(); });
  await sleep(200);
  report.start = { done: (await state(page)).done, names: await doneNames(page), sigCols: (await state(page)).sigCols };
  ok("start: 3 odhaczone", report.start.done === 3 && report.start.names.join(",") === "Ala,Bartek,Cezary", report.start);
  ok("odcisk obejmuje kolumny arkusza", report.start.sigCols.length >= 3, report.start.sigCols);

  // ── 1. Poprawiona komórka w odhaczonym wierszu ───────────────────────────
  // v1: inny hash całego wiersza → ✓ przepadało. v2: kolumna „Nr" identyfikuje wiersz.
  await reopen(page, { editedNote: { Bartek: "poprawione 2026" }, mtime: "2026-06-07T12:00:00.000Z" });
  const edited = await state(page);
  report.edited = { done: edited.done, names: await doneNames(page), changed: edited.changed };
  ok("poprawka komorki nie gubi ✓", report.edited.names.join(",") === "Ala,Bartek,Cezary", report.edited);
  ok("dopasowanie rozbite na tury", !edited.changed || typeof edited.changed.exact === "number", edited.changed);

  // ── 2. Dodana kolumna ────────────────────────────────────────────────────
  await reopen(page, { extraCol: true, mtime: "2026-06-08T12:00:00.000Z" });
  const added = await state(page);
  report.added = { done: added.done, names: await doneNames(page) };
  ok("dodana kolumna nie gubi ✓", report.added.names.join(",") === "Ala,Bartek,Cezary", report.added);

  // ── 3. Przestawione kolumny ──────────────────────────────────────────────
  await reopen(page, { order: ["Uwagi", "Miasto", "Osoba", "Nr"], mtime: "2026-06-09T12:00:00.000Z" });
  const reordered = await state(page);
  report.reordered = { done: reordered.done, names: await doneNames(page) };
  ok("przestawione kolumny nie gubia ✓", report.reordered.names.join(",") === "Ala,Bartek,Cezary", report.reordered);

  // ── 4. Usunięty wiersz → nieodnaleziony, ale NIE przypięty do cudzego ────
  await reopen(page, {
    people: ["Ala", "Cezary", "Dorota", "Edward", "Franek"],
    nrOf: { Ala: 1, Cezary: 3, Dorota: 4, Edward: 5, Franek: 6 },
    mtime: "2026-06-10T12:00:00.000Z",
  });
  const removed = await state(page);
  report.removed = {
    done: removed.done,
    names: await doneNames(page),
    lost: removed.changed?.lost,
    unmatched: removed.unmatched,
  };
  ok("zostaja tylko istniejace wiersze", report.removed.names.join(",") === "Ala,Cezary", report.removed);
  ok("usuniety wiersz policzony jako nieodnaleziony", removed.changed?.lost === 1, removed.changed);
  ok("nieodnaleziony ma czytelny podglad", /Bartek/.test(JSON.stringify(removed.unmatched)), removed.unmatched);

  // Lista w panelu „Postęp" — od zwinięcia po realne pozycje
  await page.evaluate(() => window.__transcribe.openProgress());
  await sleep(200);
  const listCollapsed = await page.evaluate(() => ({
    visible: !document.getElementById("trUnmatched").classList.contains("hidden"),
    title: document.getElementById("trUnmatchedTitle").textContent,
    items: document.querySelectorAll("#trUnmatchedList .tr-unmatched-item").length,
    note: document.getElementById("trScopeNote").textContent,
  }));
  await page.evaluate(() => document.getElementById("trUnmatchedToggle").click());
  await sleep(150);
  const listOpen = await page.evaluate(() => ({
    items: Array.from(document.querySelectorAll("#trUnmatchedList .tr-unmatched-item")).map((el) => el.querySelector(".tr-unmatched-text").textContent),
  }));
  report.list = { collapsed: listCollapsed, open: listOpen };
  ok("sekcja nieodnalezionych widoczna", listCollapsed.visible, listCollapsed);
  ok("naglowek podaje liczbe", /1/.test(listCollapsed.title), listCollapsed.title);
  ok("lista domyslnie zwinieta", listCollapsed.items === 0, listCollapsed.items);
  ok("notka tlumaczy tury dopasowania", /dokładnie|exactly/i.test(listCollapsed.note), listCollapsed.note);
  ok("po rozwinieciu widac podglad wiersza", listOpen.items.length === 1 && /Bartek/.test(listOpen.items[0]), listOpen);
  await page.evaluate(() => window.__transcribe.closeProgress());

  // ── 5. „Przenieś" z pliku, w którym była kolumna liczona „na dziś" ───────
  // Źródło: kolumna „Dni" = TODAY()-1 (zmienna). Cel: ta sama treść bez kolumny zmiennej.
  await page.evaluate(() => window.__transcribe.close());
  await page.evaluate(() => { localStorage.removeItem("excel-workbench-transcribe"); });
  await loadSheet(page, { file: "zrodlo.xlsx", order: ["Nr", "Osoba", "Miasto", "Dni"], volatile: true });
  await page.evaluate(() => window.__transcribe.open());
  await sleep(300);
  await page.evaluate(() => { window.__transcribe.mark(); window.__transcribe.mark(); });
  await sleep(150);
  await page.evaluate(() => window.__transcribe.close());
  await sleep(150);

  await loadSheet(page, { file: "cel.xlsx", order: ["Nr", "Osoba", "Miasto", "Uwagi"] });
  await page.evaluate(() => window.__transcribe.open());
  await sleep(300);
  await page.evaluate(() => window.__transcribe.openProgress());
  await sleep(200);
  const importBtns = await page.evaluate(() => document.querySelectorAll("#trStoreList .tr-store-import").length);
  await page.evaluate(() => document.querySelector("#trStoreList .tr-store-import").click());
  await sleep(300);
  const imported = await state(page);
  report.import = { buttons: importBtns, done: imported.done, names: await doneNames(page), changed: imported.changed };
  ok("przycisk Przenies jest dostepny", importBtns === 1, importBtns);
  ok("kolumna 'na dzis' w zrodle nie blokuje przeniesienia", imported.done === 2, report.import);
  ok("przeniesione trafily na te same osoby", report.import.names.join(",") === "Ala,Bartek", report.import);

  console.log(JSON.stringify(report, null, 2));
  if (errors.length) failures.push("bledy strony: " + errors.join(" | "));
  await browser.close();
  if (failures.length) {
    console.error("\n❌ transcribe-match-playwright:\n" + failures.map((f) => " - " + f).join("\n"));
    process.exit(1);
  }
  console.log("\n✅ transcribe-match-playwright OK");
}

run().catch((e) => { console.error(e); process.exit(1); });
