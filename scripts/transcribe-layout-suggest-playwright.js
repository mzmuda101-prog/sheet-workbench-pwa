// transcribe-layout-suggest-playwright.js — układ pól po NAZWACH + propozycja przeniesienia
// postępu z innej wersji pliku.
//
// Dziury, które ten test zamyka:
//   A. Układ pól (które rubryki i w jakiej kolejności) był zapisany po NUMERACH kolumn.
//      Wstawiona w Excelu kolumna przesuwała wszystko: karta pokazywała inne kolumny niż
//      wybrane, bez słowa. Przy przepisywaniu na papier = ciche przepisanie złej rubryki.
//   B. Nowa wersja pliku pod INNĄ nazwą startowała od zera; przeniesienie trzeba było
//      znaleźć samemu w „Postęp". Teraz apka proponuje — ale tylko proponuje (pliki
//      o tej samej budowie, np. 2025 i 2026, mają podobne wiersze).
//   C. Przeniesienie zabiera też układ pól, gdy w nowym pliku nikt go jeszcze nie ustawił.
//
// Sprawdzamy:
//   1. wstawiona kolumna na początku → karta pokazuje DOKŁADNIE te same rubryki w tej
//      samej kolejności,
//   2. usunięta kolumna z układu → komunikat z jej nazwą (a nie cicha podmiana),
//   3. nowa nazwa pliku + 2 nowe wiersze → propozycja z liczbami; „Przenieś" daje ✓
//      na właściwych osobach i przenosi układ pól,
//   4. zupełnie inny plik o tej samej budowie → BRAK propozycji,
//   5. „Nie, to inny plik" → nie pyta drugi raz; ręczne „Przenieś" w „Postęp” nadal jest,
//   6. plik, w którym już coś odhaczono → brak propozycji (nie mieszamy postępów).

const { chromium } = require("playwright");
const APP_URL = process.env.APP_URL || "http://127.0.0.1:4175/";
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
const state = (page) => page.evaluate(() => window.__transcribe.state());
const BASE = ["Ala", "Bartek", "Cezary", "Dorota", "Edward", "Franek"];
const NR = { Ala: 1, Bartek: 2, Cezary: 3, Dorota: 4, Edward: 5, Franek: 6, Zenon: 7, Wiktor: 8 };

async function loadSheet(page, opts = {}) {
  await page.evaluate((o) => {
    const people = o.people;
    const cols = o.cols || ["Nr", "Osoba", "Miasto", "Uwagi"];
    const letter = (i) => String.fromCharCode(65 + i);
    const ws = {};
    cols.forEach((name, i) => { ws[letter(i) + "1"] = { t: "s", v: name }; });
    people.forEach((person, rowN) => {
      const r = rowN + 2;
      const nr = (o.nrBase || 1000) + o.nr[person];
      const values = {
        Nr: { t: "n", v: nr, w: String(nr) },
        Osoba: { t: "s", v: person },
        Miasto: { t: "s", v: (o.cityPrefix || "Miasto ") + person },
        Uwagi: { t: "s", v: (o.notePrefix || "uwaga ") + person },
        Status: { t: "s", v: "nowy" },
      };
      cols.forEach((name, i) => { ws[letter(i) + r] = values[name]; });
    });
    ws["!ref"] = "A1:" + letter(cols.length - 1) + (people.length + 1);
    workbook = { SheetNames: ["Arkusz"], Sheets: { Arkusz: ws }, Props: { ModifiedDate: o.mtime || "2026-06-06T12:00:00.000Z" } };
    currentFileName = o.file;
    sheetSelect.replaceChildren();
    const opt = document.createElement("option"); opt.value = "Arkusz"; opt.textContent = "Arkusz";
    sheetSelect.appendChild(opt); sheetSelect.value = "Arkusz";
    document.getElementById("headerRow").value = "1";
    document.getElementById("autoHeaderRow").checked = false;
  }, { ...opts, nr: NR });
  await page.evaluate(() => document.getElementById("loadBtn").click());
  await page.waitForFunction((k) => typeof baseRows !== "undefined" && baseRows.length === k, opts.people.length, { timeout: 15000 });
  await sleep(300);
}

const doneNames = (page) => page.evaluate(() => {
  const st = window.__transcribe;
  const col = currentHeaders.indexOf("Osoba");
  return baseRows.filter((r) => st.isDone(getRowSelectionKey(r))).map((r) => String(getDisplayValue(r, col))).sort();
});
const cardLabels = async (page) => (await state(page)).fields.map((f) => f.label);
const suggestShown = (page) => page.evaluate(() => {
  const el = document.getElementById("trSuggest");
  return { shown: !el.classList.contains("hidden"), text: document.getElementById("trSuggestText").textContent };
});
const lastToasts = (page) => page.evaluate(() => [...document.querySelectorAll(".toast")].map((t) => t.textContent));

async function openFresh(page, opts) {
  await page.evaluate(() => window.__transcribe.state().open && window.__transcribe.close());
  await sleep(120);
  await page.evaluate(() => document.querySelectorAll(".toast").forEach((t) => t.remove()));
  await loadSheet(page, opts);
  await page.evaluate(() => window.__transcribe.open());
  await sleep(700);
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
  const results = [];
  const check = (name, ok, detail) => results.push({ name, ok: !!ok, detail });

  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  await page.evaluate(() => ensureXlsxLibs(false));
  await page.waitForFunction(() => typeof buildRows === "function" && window.__transcribe);

  // ── 0. V5: układ pól = Osoba, Nr (w tej kolejności) + 3 odhaczone
  await openFresh(page, { file: "Obieg V5.xlsx", people: BASE });
  await page.click("#trFieldsBtn");
  await sleep(150);
  await page.click("#trFieldsNoneBtn");
  await page.evaluate(() => {
    const ids = currentHeaders.map((h, i) => [h, i]);
    const idx = (n) => ids.find(([h]) => h === n)[1];
    document.getElementById(`trfield-${idx("Nr")}`).click();
    document.getElementById(`trfield-${idx("Osoba")}`).click();
  });
  // Osoba nad Nr: ▲ na wierszu „Osoba" (drugi na liście)
  await page.evaluate(() => [...document.querySelectorAll(".tr-field-row")].find((r) => r.textContent.includes("Osoba")).querySelector(".tr-move-up").click());
  await page.keyboard.press("Escape");
  await sleep(150);
  const layout0 = await cardLabels(page);
  check("start: karta = Osoba, Nr", layout0.join(",") === "Osoba,Nr", layout0);
  await page.evaluate(() => { window.__transcribe.mark(); window.__transcribe.mark(); window.__transcribe.mark(); });
  await sleep(150);
  check("start: 3 odhaczone", (await doneNames(page)).join(",") === "Ala,Bartek,Cezary", await doneNames(page));

  // ── 1. ta sama nazwa, WSTAWIONA kolumna na początku
  await openFresh(page, { file: "Obieg V5.xlsx", people: BASE, cols: ["Status", "Nr", "Osoba", "Miasto", "Uwagi"], mtime: "2026-06-07T12:00:00.000Z" });
  const layout1 = await cardLabels(page);
  check("wstawiona kolumna: karta nadal Osoba, Nr (nie przesunięta)", layout1.join(",") === "Osoba,Nr", layout1);
  check("…i ✓ na tych samych osobach", (await doneNames(page)).join(",") === "Ala,Bartek,Cezary", await doneNames(page));

  // ── 2. usunięta kolumna z układu → komunikat z nazwą
  await openFresh(page, { file: "Obieg V5.xlsx", people: BASE, cols: ["Status", "Nr", "Miasto", "Uwagi"], mtime: "2026-06-08T12:00:00.000Z" });
  const toasts2 = await lastToasts(page);
  check("usunięta kolumna: komunikat z nazwą „Osoba”", toasts2.some((x) => x.includes("Osoba")), toasts2);
  const layout2 = await cardLabels(page);
  check("…a karta nie pokazuje w jej miejscu innej kolumny", layout2.join(",") === "Nr", layout2);
  // przywróć wersję z Osobą, żeby zapis V5 miał komplet
  await openFresh(page, { file: "Obieg V5.xlsx", people: BASE, cols: ["Status", "Nr", "Osoba", "Miasto", "Uwagi"], mtime: "2026-06-07T12:00:00.000Z" });
  check("(powrót: Osoba wraca na kartę)", (await cardLabels(page)).join(",") === "Osoba,Nr", await cardLabels(page));

  // ── 3. NOWA nazwa pliku, 2 nowe osoby na górze → propozycja
  const V6 = ["Zenon", "Wiktor", ...BASE];
  await openFresh(page, { file: "Obieg V5_edited.xlsx", people: V6 });
  const s3 = await suggestShown(page);
  check("nowa wersja pliku: jest propozycja przeniesienia", s3.shown && s3.text.includes("Obieg V5.xlsx") && s3.text.includes("3"), s3);
  check("…nic nie przeniesione bez zgody", (await state(page)).done === 0, (await state(page)).done);
  check("…karta jeszcze z domyślnymi polami", (await cardLabels(page)).length > 2, await cardLabels(page));
  await page.click("#trSuggestYesBtn");
  await sleep(300);
  check("„Przenieś”: ✓ na właściwych osobach (nie na nowych)", (await doneNames(page)).join(",") === "Ala,Bartek,Cezary", await doneNames(page));
  check("„Przenieś”: układ pól też przeniesiony", (await cardLabels(page)).join(",") === "Osoba,Nr", await cardLabels(page));
  check("baner znika po decyzji", !(await suggestShown(page)).shown);

  // ── 4. zupełnie inny plik o tej samej budowie (inne osoby, inne numery) → brak propozycji
  await openFresh(page, { file: "Obieg 2025.xlsx", people: ["Zenon", "Wiktor"], nrBase: 5000, cityPrefix: "Gród ", notePrefix: "notka " });
  check("inny plik o tej samej budowie: brak propozycji", !(await suggestShown(page)).shown, await suggestShown(page));

  // ── 5. „Nie, to inny plik” → nie pyta drugi raz, ręczny import nadal jest
  await openFresh(page, { file: "Obieg kopia.xlsx", people: BASE });
  check("kopia: propozycja jest", (await suggestShown(page)).shown);
  await page.click("#trSuggestNoBtn");
  await sleep(200);
  check("„Nie”: nic nie przeniesione", (await state(page)).done === 0);
  await openFresh(page, { file: "Obieg kopia.xlsx", people: BASE });
  const s5 = await suggestShown(page);
  // „Nie" odrzuca całą rodzinę pasujących zapisów (V5 i V5_edited), więc nie pyta wcale.
  check("po „Nie”: żadna wersja tego pliku nie wraca w propozycji", !s5.shown, s5);
  await page.click("#trProgressBtn");
  await sleep(200);
  const importBtns = await page.evaluate(() => document.querySelectorAll(".tr-store-import").length);
  check("ręczne „Przenieś” w „Postęp” nadal dostępne", importBtns >= 1, importBtns);
  await page.keyboard.press("Escape");

  // ── 6. plik z własnym postępem → brak propozycji
  await openFresh(page, { file: "Obieg V5_edited.xlsx", people: V6 });
  check("plik z własnymi ✓: brak propozycji", !(await suggestShown(page)).shown && (await state(page)).done === 3, { s: await suggestShown(page), done: (await state(page)).done });

  check("brak błędów strony", errors.length === 0, errors.slice(0, 3).join(" | "));
  await browser.close();

  let failed = 0;
  for (const r of results) {
    console.log(`${r.ok ? "✅" : "❌"} ${r.name}${r.detail !== undefined && !r.ok ? `  (${JSON.stringify(r.detail)})` : ""}`);
    if (!r.ok) failed += 1;
  }
  if (failed) { console.error(`\n${failed} z ${results.length} nie przeszło`); process.exit(1); }
  console.log(`\n✅ transcribe-layout-suggest: ${results.length}/${results.length}`);
}

run().catch((e) => { console.error(e); process.exit(1); });
