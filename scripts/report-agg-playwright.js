// report-agg-playwright.js — sekcje raportu liczone SILNIKIEM agregacji apki.
//
// Sprawdzamy:
//   1. „Zestawienia”: kwota wg statusu (suma/średnia/udział), czas trwania start→koniec
//      wg statusu (parę kolumn wykrywa silnik), tabela krzyżowa Status × Miasto,
//   2. wnioski z agregacji (udział największej grupy, różnica średnich, najdłuższy czas),
//   3. „Agregacja z panelu” = dokładnie to, co ustawione w panelu Agregacje,
//   4. raport NIE zmienia ustawień panelu Agregacje (pożycza silnik i oddaje stan),
//   5. filtr: zestawienia liczą tylko przefiltrowane wiersze.

const { chromium } = require("playwright");
const APP_URL = process.env.APP_URL || "http://127.0.0.1:4175/";
const SLEEP_SCALE = Math.min(1, Math.max(0.1, Number(process.env.SLEEP_SCALE || 1)));
const sleep = (ms) => new Promise((r) => setTimeout(r, Math.round(ms * SLEEP_SCALE)));

// Status, Miasto, Kwota, dni trwania (start → koniec)
const ROWS = [
  ["W toku", "Kraków", 100, 4], ["W toku", "Warszawa", 120, 5], ["W toku", "Kraków", 110, 6],
  ["W toku", "Warszawa", 130, 5], ["W toku", "Kraków", 90, 5],
  ["Zakończone", "Kraków", 400, 20], ["Zakończone", "Warszawa", 380, 22], ["Zakończone", "Kraków", 420, 18],
  ["Zakończone", "Warszawa", 400, 20],
  ["Anulowane", "Kraków", 50, 2], ["Anulowane", "Warszawa", 60, 1], ["Anulowane", "Kraków", 40, 3],
];

async function loadSheet(page) {
  await page.evaluate((rows) => {
    const ws = {};
    ["Nr", "Status", "Miasto", "Kwota", "Data start", "Data koniec"].forEach((h, i) => { ws[String.fromCharCode(65 + i) + "1"] = { t: "s", v: h }; });
    rows.forEach(([st, mi, kw, days], i) => {
      const r = i + 2;
      const start = new Date(2026, 0, 1 + i * 3);
      const end = new Date(start.getTime() + days * 86400000);
      const iso = (d) => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
      ws["A" + r] = { t: "n", v: i + 1, w: String(i + 1) };
      ws["B" + r] = { t: "s", v: st };
      ws["C" + r] = { t: "s", v: mi };
      ws["D" + r] = { t: "n", v: kw, w: String(kw) };
      ws["E" + r] = { t: "s", v: iso(start) };
      ws["F" + r] = { t: "s", v: iso(end) };
    });
    ws["!ref"] = "A1:F" + (rows.length + 1);
    workbook = { SheetNames: ["Zlecenia"], Sheets: { Zlecenia: ws }, Props: {} };
    currentFileName = "agregacje.xlsx";
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

const aggBoxes = (page) => page.evaluate(() => Array.from(document.querySelectorAll('#rpPage [data-section="aggAuto"] .rp-agg')).map((box) => ({
  title: box.querySelector(".rp-h3").textContent,
  head: Array.from(box.querySelectorAll("thead th")).map((th) => th.textContent),
  rows: Array.from(box.querySelectorAll("tbody tr")).map((tr) => Array.from(tr.cells).map((td) => td.textContent.replace(/ /g, " "))),
})));

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

  // Panel Agregacje ustawiony „po swojemu”: liczba wierszy wg Miasta.
  await page.evaluate(() => {
    Object.assign(aggregationWorkbenchState, {
      groupBy: "Miasto", groupBy2: "", measures: ["count_rows"], aggregation: "count",
      headerRowChoice: "custom", customHeaderRow: currentHeaderRow, showCount: 7, groupMode: "exact",
    });
  });
  const stateBefore = await page.evaluate(() => window.__report.aggState());

  await page.evaluate(() => window.__report.open());
  await sleep(300);
  const secs = await page.evaluate(() => window.__report.sections());
  check("normalny zawiera zestawienia", secs.includes("aggAuto"), secs);
  const boxes = await aggBoxes(page);
  check("3 zestawienia: kwota, czas trwania, krzyżowa", boxes.length === 3, boxes.map((b) => b.title));

  const by = boxes.find((b) => /Kwota/.test(b.title));
  const row = (label) => by && by.rows.find((r) => r[0] === label);
  check("kwota wg statusu: Zakończone suma 1600, śr. 400, udział 70%",
    row("Zakończone") && row("Zakończone")[2].replace(/\s/g, "") === "1600" && row("Zakończone")[3] === "400" && row("Zakończone")[4] === "70%", by);
  check("kwota wg statusu: W toku 5 wierszy, suma 550", row("W toku") && row("W toku")[1] === "5" && row("W toku")[2] === "550", by);
  check("pierwsza grupa = największa suma", by && by.rows[0][0] === "Zakończone", by && by.rows.map((r) => r[0]));

  const dur = boxes.find((b) => /Czas trwania/.test(b.title));
  check("czas trwania z pary Data start → Data koniec", dur && /Data start → Data koniec/.test(dur.title), dur && dur.title);
  check("czas trwania: najdłużej Zakończone (20 dni)", dur && dur.rows[0][0] === "Zakończone" && /20/.test(dur.rows[0][2]), dur);

  const cross = boxes.find((b) => /×/.test(b.title));
  check("krzyżowa: kolumny Kraków, Warszawa, Razem", cross && cross.head.join("|") === "Status|Kraków|Warszawa|Razem", cross && cross.head);
  const crossTotal = cross ? cross.rows.reduce((s, r) => s + Number(r[r.length - 1]), 0) : 0;
  check("krzyżowa: suma = 12 wierszy", crossTotal === 12, cross);
  const cz = cross && cross.rows.find((r) => r[0] === "Anulowane");
  check("krzyżowa: Anulowane = 2 Kraków, 1 Warszawa", cz && cz[1] === "2" && cz[2] === "1", cz);

  const f = await page.evaluate(() => window.__report.findings());
  check("wniosek: udział największej grupy", f.some((x) => /Najwięcej „Kwota” w grupie „Zakończone”: 70%/.test(x)), f);
  check("wniosek: różnica średnich", f.some((x) => /Średnia „Kwota” w „Zakończone” \(400\) jest 8× wyższa niż w „Anulowane” \(50\)/.test(x)), f);
  check("wniosek: najdłuższy czas", f.some((x) => /Najdłużej trwa „Zakończone”/.test(x)), f);

  // ── Agregacja z panelu (preset szczegółowy) ──────────────────────────────
  await page.evaluate(() => window.__report.setPreset("detailed"));
  await sleep(150);
  const panel = await page.evaluate(() => {
    const sec = document.querySelector('#rpPage [data-section="aggPanel"]');
    return sec ? { note: sec.querySelector(".rp-note").textContent, labels: Array.from(sec.querySelectorAll(".rp-bar-label")).map((e) => e.textContent), values: Array.from(sec.querySelectorAll(".rp-bar-value")).map((e) => e.textContent) } : null;
  });
  check("agregacja z panelu jest w szczegółowym", !!panel, panel);
  check("opisuje ustawienia panelu (liczba wg Miasto)", panel && /liczba/.test(panel.note) && /wg Miasto/.test(panel.note), panel && panel.note);
  check("grupy z panelu: Kraków 7, Warszawa 5", panel && panel.labels.join() === "Kraków,Warszawa" && panel.values.join() === "7,5", panel);

  const stateAfter = await page.evaluate(() => window.__report.aggState());
  check("ustawienia panelu Agregacje nietknięte", stateAfter === stateBefore, { before: stateBefore, after: stateAfter });

  // ── Filtr: zestawienia tylko z wycinka ───────────────────────────────────
  await page.evaluate(() => window.__report.close());
  await page.evaluate(() => {
    searchQueryEl.value = 'Miasto:="Kraków"';
    filterOperatorsEl.checked = true;
    syncQuickSearchOperatorsControls();
    filtersCommitted = true;
    applyFilters();
    sortRows();
    scheduleViewRefresh({ table: true, immediate: true });
    window.__report.open();
    window.__report.setPreset("normal");
  });
  await sleep(300);
  const fb = await aggBoxes(page);
  const fcross = fb.find((b) => /×/.test(b.title));
  const fby = fb.find((b) => /Kwota/.test(b.title));
  const fz = fby && fby.rows.find((r) => r[0] === "Zakończone");
  // Po filtrze Miasto ma jedną wartość — nie ma czego krzyżować, więc tabeli ma nie być.
  check("filtr: bez tabeli krzyżowej (Miasto = jedna wartość)", !fcross && fb.length === 2, fb.map((b) => b.title));
  check("filtr: Zakończone w Krakowie suma 820", fz && fz[2].replace(/\s/g, "") === "820", fby);
  check("ustawienia panelu nadal nietknięte", (await page.evaluate(() => window.__report.aggState())) === stateBefore);

  // ── Liczba stron w pasku = liczba stron prawdziwego PDF ─────────────────
  // (symulacja łamania stron musi się zgadzać z tym, co robi przeglądarka przy druku)
  for (const [preset, size, margin] of [["normal", "normal", "normal"], ["detailed", "normal", "narrow"], ["detailed", "large", "wide"]]) {
    await page.evaluate(([preset, size, margin]) => {
      window.__report.close();
      localStorage.setItem("swb-report-prefs", JSON.stringify({ preset, size, margin }));
      window.__report.open();
    }, [preset, size, margin]);
    await sleep(250);
    const note = await page.evaluate(() => document.getElementById("rpFitNote").textContent);
    const bad = await page.evaluate(() => window.__report.marginViolations());
    check(`margines ${preset}/${size}: nic w pasie 15 mm przy krawędzi`, bad.length === 0, bad.slice(0, 3));
    await page.evaluate(() => document.body.classList.add("rp-printing"));
    await page.emulateMedia({ media: "print" });
    // przed pdf(): generowanie PDF odpala „afterprint”, a raport wtedy zdejmuje klasę druku
    const spacers = await page.evaluate(() => Array.from(document.querySelectorAll(".rp-sheet-gap, .rp-gap-row, .rp-head-repeat")).filter((el) => getComputedStyle(el).display !== "none").map((el) => el.className));
    const pdf = await page.pdf({ preferCSSPageSize: true });
    await page.emulateMedia({ media: "screen" });
    await page.evaluate(() => document.body.classList.remove("rp-printing"));
    const real = (pdf.toString("latin1").match(/\/Type\s*\/Page[^s]/g) || []).length;
    const pred = Number((note.match(/(\d+) stron/) || [0, 1])[1]);
    check(`strony ${preset}/${size}: pasek ${pred} = PDF ${real}`, pred === real && real >= 1, { note, real });
    check(`druk ${preset}/${size}: przekładki podglądu niewidoczne`, spacers.length === 0, spacers);
  }

  check("brak błędów w konsoli", errors.length === 0, errors);
  await browser.close();
  const failed = results.filter((r) => !r.ok);
  if (failed.length) {
    console.log(`\n❌ report-agg: ${failed.length} z ${results.length} nie przeszło`);
    process.exit(1);
  }
  console.log(`\n✅ report-agg: ${results.length}/${results.length}`);
}

run().catch((e) => { console.error(e); process.exit(1); });
