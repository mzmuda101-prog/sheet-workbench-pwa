// report-scope-playwright.js — ZAKRES raportu niezależny od tabeli.
//
// Sprawdzamy:
//   1. domyślnie „Jak w tabeli”; bez filtra w tabeli nie ma zdublowanego chipa „Cały arkusz”,
//   2. podpowiedzi: wartości kolumny Status i „ostatnie 30 dni” (świeże daty) — każda ma
//      sensowny wycinek (≥3 wiersze, mniej niż wszystko), a liczba na chipie = wiersze raportu,
//   3. wybór podpowiedzi liczy raport z tych wierszy, TABELA się nie zmienia (viewRows,
//      filtr), panel Agregacje nietknięty, zestawienia liczone z zakresu,
//   4. własne zapytanie (operatory), zapytanie bez wyników = komunikat i stary zakres,
//      tokeny @ = komunikat,
//   5. „Pokaż w tabeli” ustawia filtr tabeli = ten sam zestaw wierszy, zakres wraca do „Jak w tabeli”,
//   6. przy filtrze w tabeli jest „Cały arkusz”, pamięć zakresu po ponownym otwarciu,
//   7. telefon: pasek w jednym rzędzie, bez poziomego przewijania strony.

const { chromium } = require("playwright");
const APP_URL = process.env.APP_URL || "http://127.0.0.1:4175/";
const SLEEP_SCALE = Math.min(1, Math.max(0.1, Number(process.env.SLEEP_SCALE || 1)));
const sleep = (ms) => new Promise((r) => setTimeout(r, Math.round(ms * SLEEP_SCALE)));

// Status, Miasto, Kwota, ile dni temu (data zgłoszenia liczona od DZIŚ → podpowiedź „30 dni”)
const ROWS = [
  ["W toku", "Kraków", 100, 2], ["W toku", "Warszawa", 120, 5], ["W toku", "Kraków", 110, 9],
  ["W toku", "Warszawa", 130, 14], ["W toku", "Kraków", 90, 20],
  ["Zakończone", "Kraków", 400, 40], ["Zakończone", "Warszawa", 380, 55], ["Zakończone", "Kraków", 420, 70],
  ["Zakończone", "Warszawa", 400, 90],
  ["Anulowane", "Kraków", 50, 100], ["Anulowane", "Warszawa", 60, 120], ["Anulowane", "Kraków", 40, 150],
];

async function loadSheet(page) {
  await page.evaluate((rows) => {
    const ws = {};
    ["Nr", "Status", "Miasto", "Kwota", "Data"].forEach((h, i) => { ws[String.fromCharCode(65 + i) + "1"] = { t: "s", v: h }; });
    const now = new Date();
    const iso = (d) => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
    rows.forEach(([st, mi, kw, ago], i) => {
      const r = i + 2;
      ws["A" + r] = { t: "n", v: i + 1, w: String(i + 1) };
      ws["B" + r] = { t: "s", v: st };
      ws["C" + r] = { t: "s", v: mi };
      ws["D" + r] = { t: "n", v: kw, w: String(kw) };
      ws["E" + r] = { t: "s", v: iso(new Date(now.getFullYear(), now.getMonth(), now.getDate() - ago)) };
    });
    ws["!ref"] = "A1:E" + (rows.length + 1);
    workbook = { SheetNames: ["Zlecenia"], Sheets: { Zlecenia: ws }, Props: {} };
    currentFileName = "zakres.xlsx";
    sheetSelect.replaceChildren();
    const opt = document.createElement("option"); opt.value = "Zlecenia"; opt.textContent = "Zlecenia";
    sheetSelect.appendChild(opt); sheetSelect.value = "Zlecenia";
    document.getElementById("headerRow").value = "1";
    document.getElementById("autoHeaderRow").checked = false;
  }, ROWS);
  await page.evaluate(() => document.getElementById("loadBtn").click());
  await page.waitForFunction((n) => typeof baseRows !== "undefined" && baseRows.length === n, ROWS.length, { timeout: 15000 });
  await sleep(350);
}

const chips = (page) => page.evaluate(() => Array.from(document.querySelectorAll("#rpScopeChips [data-scope]")).map((b) => ({
  id: b.dataset.scope,
  label: b.querySelector(".rp-scope-chip-text").textContent,
  n: b.querySelector(".rp-scope-chip-n")?.textContent || "",
  on: b.getAttribute("aria-pressed") === "true",
  hint: b.getAttribute("data-hint") || "",
})));

const state = (page) => page.evaluate(() => ({
  scope: window.__report.scope(),
  rows: window.__report.data().rows,
  viewRows: viewRows.length,
  tableQuery: searchQueryEl.value,
  committed: filtersCommitted,
  scopeText: document.querySelector("#rpPage .rp-scope")?.textContent || "",
  scopeFiltered: !!document.querySelector("#rpPage .rp-scope.is-filtered"),
  status: document.getElementById("rpScopeStatus").textContent,
  row2: !document.getElementById("rpScopeRow2").classList.contains("hidden"),
  applyVisible: !document.getElementById("rpScopeApply").classList.contains("hidden"),
  query: document.getElementById("rpScopeQuery").value,
}));

async function typeQuery(page, q) {
  await page.fill("#rpScopeQuery", q);
  await page.press("#rpScopeQuery", "Enter");
  await sleep(150);
}

async function run() {
  const browser = await chromium.launch({ headless: true });
  const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1100, height: 820 } });
  await context.addInitScript(() => {
    localStorage.setItem("introPlayed", "true");
    localStorage.setItem("swb-report-prefs", JSON.stringify({ preset: "normal" }));
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
  await page.waitForFunction(() => typeof buildRows === "function" && window.__report);
  await loadSheet(page);

  const aggBefore = await page.evaluate(() => window.__report.aggState());
  await page.evaluate(() => window.__report.open());
  await sleep(250);

  // ── 1. Domyślnie ──
  let c = await chips(page);
  let s = await state(page);
  check("domyślnie „Jak w tabeli” (12 wierszy)", s.scope.kind === "view" && s.rows === 12 && c[0].id === "view" && c[0].on, { s, c });
  check("bez filtra w tabeli nie ma drugiego chipa „Cały arkusz”", !c.some((x) => x.id === "all"), c.map((x) => x.id));
  check("drugi rząd (pole) schowany na starcie", !s.row2, s);

  // ── 2. Podpowiedzi ──
  const sug = await page.evaluate(() => window.__report.suggestions());
  check("podpowiedzi: W toku / Zakończone / Anulowane z kolumny Status",
    ["W toku", "Zakończone", "Anulowane"].every((v) => sug.some((x) => x.q === `Status:="${v}"`)), sug.map((x) => x.q));
  const last30 = sug.find((x) => x.kind === "date" && /30/.test(x.label));
  check("podpowiedź „Ostatnie 30 dni” przy świeżych datach (5 wierszy: 2–20 dni temu)", last30 && last30.n === 5, last30);
  check("każda podpowiedź: ≥3 wiersze i mniej niż wszystko", sug.every((x) => x.n >= 3 && x.n < 12), sug.map((x) => [x.label, x.n]));
  check("co najwyżej 5 podpowiedzi", sug.length <= 5, sug.length);
  check("dymek podpowiedzi mówi skąd i pokazuje zapytanie", c.some((x) => /Kolumna „Status”: „W toku” to 42% wierszy\. Zapytanie: Status:="W toku"/.test(x.hint)), c.map((x) => x.hint));

  // ── 3. Wybór podpowiedzi ──
  await page.click('#rpScopeChips [data-scope="q:Status:=\\"W toku\\""]');
  await sleep(200);
  s = await state(page);
  check("„W toku”: raport liczy 5 wierszy", s.scope.kind === "query" && s.rows === 5, s);
  check("tabela nietknięta (12 wierszy, bez filtra)", s.viewRows === 12 && s.tableQuery === "" && !s.committed, s);
  check("kartka: zakres opisuje zapytanie i jest „przefiltrowany”", /W toku/.test(s.scopeText) && /5 z 12/.test(s.scopeText.replace(/ /g, " ")) && s.scopeFiltered, s.scopeText);
  check("pole pokazuje zapytanie (widać składnię)", s.row2 && s.query === 'Status:="W toku"', s);
  check("„Inny niż w tabeli” + przycisk „Pokaż w tabeli”", /Inny niż w tabeli/.test(s.status) && s.applyVisible, s);
  const aggRows = await page.evaluate(() => {
    const box = document.querySelector('#rpPage [data-section="aggAuto"]');
    return box ? box.textContent : "";
  });
  check("zestawienia liczone z zakresu (nie ma Zakończone/Anulowane)", !/Zakończone|Anulowane/.test(aggRows), aggRows.slice(0, 200));
  check("panel Agregacje nietknięty", (await page.evaluate(() => window.__report.aggState())) === aggBefore);
  check("chip „W toku” wciśnięty, liczba na chipie = 5", (await chips(page)).some((x) => x.label === "W toku" && x.on && x.n === "5"));

  // ── 4. Własne zapytanie ──
  await page.click('#rpScopeChips [data-scope="custom"]');
  await typeQuery(page, "Kwota:>>300");
  s = await state(page);
  check("własne „Kwota:>>300” → 4 wiersze, chip „Własny…” wciśnięty", s.scope.kind === "query" && s.rows === 4 && (await chips(page)).find((x) => x.id === "custom").on, s);
  await typeQuery(page, "Kwota:>>99999");
  s = await state(page);
  check("zapytanie bez wyników: komunikat, raport zostaje przy 4 wierszach", s.rows === 4 && /Nic nie pasuje/.test(s.status), s);
  await typeQuery(page, "@wtoku");
  s = await state(page);
  check("token @ → komunikat o trybach auto, zakres bez zmian", s.rows === 4 && /Tryby auto/.test(s.status), s);
  await typeQuery(page, "Kwota:>>300");

  // ── 5. Pokaż w tabeli ──
  await page.click("#rpScopeApply");
  await sleep(250);
  s = await state(page);
  check("„Pokaż w tabeli”: tabela = te same 4 wiersze, filtr z zapytaniem", s.viewRows === 4 && s.committed && s.tableQuery === "Kwota:>>300", s);
  check("zakres wraca do „Jak w tabeli”, raport 4 wiersze, bez „Inny niż w tabeli”", s.scope.kind === "view" && s.rows === 4 && !s.applyVisible, s);

  // ── 6. Przy filtrze w tabeli: „Cały arkusz” ──
  c = await chips(page);
  check("przy filtrze w tabeli jest chip „Cały arkusz” (12)", c.some((x) => x.id === "all" && x.n === "12"), c);
  await page.click('#rpScopeChips [data-scope="all"]');
  await sleep(200);
  s = await state(page);
  check("„Cały arkusz”: raport 12 wierszy, kartka bez bursztynu, tabela dalej 4", s.rows === 12 && !s.scopeFiltered && s.viewRows === 4, s);
  await page.evaluate(() => { window.__report.close(); window.__report.open(); });
  await sleep(200);
  s = await state(page);
  check("po ponownym otwarciu zakres zapamiętany (Cały arkusz)", s.scope.kind === "all" && s.rows === 12, s);

  // Pusta tabela → raport i tak się otwiera, liczy cały arkusz i mówi to wprost.
  await page.evaluate(() => {
    window.__report.close();
    searchQueryEl.value = "nie-ma-takiego";
    filtersCommitted = true;
    applyFilters();
    window.__report.setScope({ kind: "view" });
  });
  await page.evaluate(() => window.__report.open());
  await sleep(200);
  s = await state(page);
  check("pusta tabela: raport otwarty na całym arkuszu + komunikat", s.viewRows === 0 && s.rows === 12 && /nic nie jest widoczne/.test(s.status), s);

  // Druk: pasek zakresu NIE może wejść na kartkę (dokładał stronę — liczba stron ≠ PDF).
  await page.emulateMedia({ media: "print" });
  const printHidden = await page.evaluate(() => {
    document.body.classList.add("rp-printing");
    const hidden = getComputedStyle(document.getElementById("rpScopeBar")).display === "none";
    document.body.classList.remove("rp-printing");
    return hidden;
  });
  await page.emulateMedia({ media: "screen" });
  check("druk: pasek zakresu schowany", printHidden);

  // ── 7. Telefon ──
  await page.evaluate(() => window.__report.close());
  await page.setViewportSize({ width: 375, height: 760 });
  await page.evaluate(() => { searchQueryEl.value = ""; filtersCommitted = false; applyFilters(); window.__report.open(); });
  await sleep(250);
  const phone = await page.evaluate(() => {
    const bar = document.getElementById("rpScopeChips");
    const chipsEls = Array.from(bar.querySelectorAll(".rp-scope-chip"));
    const tops = new Set(chipsEls.map((b) => Math.round(b.getBoundingClientRect().top)));
    return {
      oneRow: tops.size === 1,
      overflow: bar.classList.contains("has-overflow"),
      pageScroll: document.documentElement.scrollWidth > window.innerWidth + 1,
      overlayScroll: document.getElementById("reportOverlay").scrollWidth > window.innerWidth + 1,
    };
  });
  check("telefon: chipy w jednym rzędzie, przewijane w bok z wygaszeniem", phone.oneRow && phone.overflow, phone);
  check("telefon: bez poziomego przewijania strony", !phone.pageScroll && !phone.overlayScroll, phone);
  // Wygaszenie: na starcie tylko z prawej; po wybraniu chipa z końca listy — widoczny i z lewej wygaszenie.
  const fadeStart = await page.evaluate(() => { const el = document.getElementById("rpScopeChips"); return { l: el.classList.contains("fade-left"), r: el.classList.contains("fade-right") }; });
  check("telefon: na starcie wygaszona tylko prawa krawędź", !fadeStart.l && fadeStart.r, fadeStart);
  await page.evaluate(() => { const b = Array.from(document.querySelectorAll('#rpScopeChips [data-scope^="q:"]')).pop(); b.click(); });
  await sleep(200);
  const fadeEnd = await page.evaluate(() => {
    const el = document.getElementById("rpScopeChips");
    const on = el.querySelector('[aria-pressed="true"]').getBoundingClientRect();
    const box = el.getBoundingClientRect();
    return { l: el.classList.contains("fade-left"), visible: on.left >= box.left - 1 && on.right <= box.right + 1 };
  });
  check("telefon: wybrany chip z końca listy jest widoczny, lewa krawędź wygaszona", fadeEnd.l && fadeEnd.visible, fadeEnd);

  check("brak błędów w konsoli", errors.length === 0, errors);
  await browser.close();
  const failed = results.filter((r) => !r.ok);
  if (failed.length) {
    console.log(`\n❌ report-scope: ${failed.length} z ${results.length} nie przeszło`);
    process.exit(1);
  }
  console.log(`\n✅ report-scope: ${results.length}/${results.length}`);
}

run().catch((e) => { console.error(e); process.exit(1); });
