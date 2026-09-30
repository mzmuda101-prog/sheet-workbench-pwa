// report-angles-playwright.js — KĄTY raportu: „Stan teraz” i „Porównanie grup”.
//
// Sprawdzamy:
//   1. panel „Zawartość”: chipy kątów, opis wykrycia, sekcja kąta dołożona do presetu
//      (checkbox zablokowany „z wybranego kąta”), etykieta przycisku „Stan teraz · krótki”,
//   2. Stan teraz (kolumna Status): otwarte/zamknięte z NAZW wartości (w tym „niezakończone”),
//      najdłużej otwarte, wiek, po terminie, świeżość; wnioski kąta na początku listy,
//   3. Stan teraz bez statusu: para dat przyjęcia → zamknięcia,
//   4. daty „z kosmosu” (liczby wyświetlane jako 1900) NIE dają kąta,
//   5. Porównanie: grupy kolumny (różnice ≥1,5×, skład, % otwartych), zakres vs reszta
//      arkusza, brak różnic = uczciwe zdanie, wybór „Porównaj” w panelu,
//   6. łamanie stron z nowymi sekcjami = PDF, marginesy, panel Agregacje nietknięty,
//   7. (jeśli jest plik) RODO: stan z cykli Trybów auto + „kto” z otwartego cyklu.

const fs = require("fs");
const { chromium } = require("playwright");
const APP_URL = process.env.APP_URL || "http://127.0.0.1:4175/";
const SLEEP_SCALE = Math.min(1, Math.max(0.1, Number(process.env.SLEEP_SCALE || 1)));
const sleep = (ms) => new Promise((r) => setTimeout(r, Math.round(ms * SLEEP_SCALE)));
const RODO = "scripts/RODO Obieg terenów 2026 V5.xlsx";

// Klient, Status, Miasto, Kwota, przyjęto (dni temu), termin (dni od dziś; ujemne = minął)
const ROWS = [
  ["Alfa", "W toku", "Kraków", 100, 3, 10],
  ["Beta", "W toku", "Kraków", 120, 12, -2],
  ["Gamma", "Nowe", "Kraków", 90, 1, 20],
  ["Delta", "niezakończone", "Kraków", 110, 45, -30],
  ["Epsilon", "W toku", "Kraków", 130, 120, -90],
  ["Zeta", "Zakończone", "Warszawa", 400, 60, 5],
  ["Eta", "Zakończone", "Warszawa", 420, 80, -40],
  ["Theta", "Zakończone", "Warszawa", 380, 90, -50],
  ["Iota", "Anulowane", "Warszawa", 410, 100, -60],
  ["Kappa", "Zakończone", "Warszawa", 390, 70, -20],
];

async function loadRows(page, { headers, rows, file = "katy.xlsx" }) {
  await page.evaluate(({ headers, rows, file }) => {
    const ws = {};
    const col = (i) => String.fromCharCode(65 + i);
    headers.forEach((h, i) => { ws[col(i) + "1"] = { t: "s", v: h }; });
    rows.forEach((r, ri) => r.forEach((v, ci) => {
      if (v === null || v === "") return;
      ws[col(ci) + (ri + 2)] = typeof v === "number" ? { t: "n", v, w: String(v) } : typeof v === "object" ? v : { t: "s", v };
    }));
    ws["!ref"] = `A1:${col(headers.length - 1)}${rows.length + 1}`;
    workbook = { SheetNames: ["Arkusz"], Sheets: { Arkusz: ws }, Props: {} };
    currentFileName = file;
    sheetSelect.replaceChildren();
    const opt = document.createElement("option"); opt.value = "Arkusz"; opt.textContent = "Arkusz";
    sheetSelect.appendChild(opt); sheetSelect.value = "Arkusz";
    document.getElementById("headerRow").value = "1";
    document.getElementById("autoHeaderRow").checked = false;
  }, { headers, rows, file });
  await page.evaluate(() => document.getElementById("loadBtn").click());
  await page.waitForFunction((n) => typeof baseRows !== "undefined" && baseRows.length === n && currentFileName, rows.length, { timeout: 15000 });
  await sleep(300);
}

function isoAgo(days) {
  const d = new Date();
  d.setDate(d.getDate() - days);
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
}

const sectionText = (page, id) => page.evaluate((x) => (document.querySelector(`#rpPage [data-section="${x}"]`)?.innerText || "").replace(/ | /g, " "), id);

async function run() {
  const browser = await chromium.launch({ headless: true });
  const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1200, height: 900 } });
  await context.addInitScript(() => {
    localStorage.setItem("introPlayed", "true");
    localStorage.setItem("swb-report-prefs", JSON.stringify({ preset: "short", angle: "overview" }));
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

  const main = {
    headers: ["Klient", "Status", "Miasto", "Kwota", "Data przyjęcia", "Termin"],
    rows: ROWS.map(([k, st, mi, kw, ago, due]) => [k, st, mi, kw, isoAgo(ago), isoAgo(-due)]),
  };
  await loadRows(page, main);
  const aggBefore = await page.evaluate(() => window.__report.aggState());

  // ── 1. Panel kątów ──
  await page.evaluate(() => window.__report.open());
  await page.click("#rpContentBtn");
  await sleep(150);
  const chips = await page.evaluate(() => Array.from(document.querySelectorAll("#rpAngles [data-angle]")).map((b) => ({ a: b.dataset.angle, on: b.getAttribute("aria-pressed") === "true", dis: b.disabled })));
  check("chipy kątów: Przegląd (wciśnięty), Stan teraz, Porównanie grup, Porównanie okresów — dostępne", chips.length === 4 && chips[0].on && chips.every((c) => !c.dis), chips);
  await page.click('#rpAngles [data-angle="state"]');
  await sleep(200);
  const ui = await page.evaluate(() => ({
    btn: document.getElementById("rpContentBtn").textContent,
    info: document.querySelector("#rpAngles .rp-angle-info")?.textContent || "",
    sections: window.__report.sections(),
    forced: (() => { const cb = document.querySelector('#rpSectionList input[value="state"]'); return cb && cb.checked && cb.disabled; })(),
    forcedNote: document.querySelector('#rpSectionList input[value="state"]')?.parentElement.textContent || "",
  }));
  check("przycisk: „Zawartość: Stan teraz · krótki”", ui.btn === "Zawartość: Stan teraz · krótki", ui.btn);
  check("opis wykrycia w panelu (otwarte = W toku, Nowe, niezakończone; zamknięte = Zakończone, Anulowane)",
    /Status/.test(ui.info) && /otwarte = .*„W toku”.*„Nowe”/.test(ui.info) && /niezakończone/.test(ui.info.split("zamknięte")[0]) && /zamknięte = .*„Zakończone”.*„Anulowane”/.test(ui.info), ui.info);
  check("krótki preset + sekcja kąta (kafelki, wnioski, stan, wykres)", JSON.stringify(ui.sections) === JSON.stringify(["tiles", "findings", "state", "chart"]), ui.sections);
  check("checkbox sekcji kąta zaznaczony i zablokowany „z wybranego kąta”", ui.forced && /z wybranego kąta/.test(ui.forcedNote), ui);

  // Strażnik: po wielu odświeżeniach kartki JEDEN klik = jedno przełączenie panelu (nasłuchy
  // dokładały się przy każdym renderze → parzysta liczba = panel „nie reaguje”).
  const toggles = await page.evaluate(async () => {
    for (let i = 0; i < 5; i++) window.__report.setPreset(i % 2 ? "short" : "normal");
    window.__report.setPreset("short");
    const panel = document.getElementById("rpContentPanel");
    const before = panel.classList.contains("hidden");
    document.getElementById("rpContentBtn").click();
    const after1 = panel.classList.contains("hidden");
    document.getElementById("rpContentBtn").click();
    const after2 = panel.classList.contains("hidden");
    let renders = 0;
    const orig = window.__report.pageCount;
    const obs = new MutationObserver(() => { renders += 1; });
    obs.observe(document.getElementById("rpPage"), { childList: true });
    document.querySelector('#rpPresets [data-preset="short"]').click();
    await new Promise((r) => setTimeout(r, 50));
    obs.disconnect();
    return { before, after1, after2, renders };
  });
  check("po wielu renderach: klik „Zawartość” przełącza raz, preset renderuje raz", toggles.after1 !== toggles.before && toggles.after2 === toggles.before && toggles.renders <= 2, toggles);
  if (toggles.after2) await page.click("#rpContentBtn");

  // ── 2. Stan teraz (Status) ──
  const st = await sectionText(page, "state");
  check("skład: Otwarte 5 (50%), Zamknięte 5 (50%)", /Otwarte: 5 \(50%\)/.test(st) && /Zamknięte: 5 \(50%\)/.test(st), st);
  check("wiek otwartych w przedziałach (do 7 dni: 2, 8–30: 1, 31–90: 1, >90: 1)", /do 7 dni\s*2/.test(st) && /8–30 dni\s*1/.test(st) && /31–90 dni\s*1/.test(st) && /ponad 90 dni\s*1/.test(st), st);
  check("najdłużej otwarte: Epsilon (120 dni) pierwszy, z nazwą klienta", /Najdłużej otwarte[\s\S]*Epsilon[\s\S]*120 dni/.test(st) && st.indexOf("Epsilon") < st.indexOf("Delta"), st);
  check("po terminie: 3 otwarte (Epsilon 90, Delta 30, Beta 2) — zamknięte się nie liczą", /Po terminie \(„Termin”\): 3/.test(st) && /Epsilon[\s\S]*90 dni[\s\S]*Delta[\s\S]*30 dni[\s\S]*Beta[\s\S]*2 dni/.test(st) && !/Eta\s/.test(st.split("Po terminie")[1] || ""), st);
  const f = await page.evaluate(() => window.__report.findings());
  check("wnioski kąta NA POCZĄTKU: otwartych 5 z 10, najdłużej Epsilon", /Otwartych: 5 z 10 \(50%\)\. Najdłużej czeka „Epsilon” — 120 dni/.test(f[0]), f);
  check("wniosek: połowa otwartych czeka dłużej niż 12 dni", f.some((x) => /Połowa otwartych czeka dłużej niż 12 dni/.test(x)), f);
  check("wniosek ostrzega o 3 po terminie", f.some((x) => /3 otwartych po terminie \(„Termin”\); najbardziej „Epsilon” — 90 dni/.test(x)), f);
  check("wniosek o świeżości (ostatni wpis + 30 dni)", f.some((x) => /Ostatni wpis: .*w ostatnich 30 dniach — 3 wierszy/.test(x)), f);
  check("zakres + stan: tylko Kraków → 5 otwartych z 5", await page.evaluate(() => {
    window.__report.setScope({ kind: "query", q: 'Miasto:="Kraków"' });
    const s = window.__report.data().state;
    const ok = s.counts.open === 5 && s.total === 5;
    window.__report.setScope({ kind: "view" });
    return ok;
  }));

  // ── 5. Porównanie ──
  await page.click('#rpAngles [data-angle="compare"]');
  await sleep(200);
  let cmp = await sectionText(page, "compare");
  const pick = await page.evaluate(() => ({ value: document.getElementById("rpCompareBy")?.value, opts: Array.from(document.querySelectorAll("#rpCompareBy option")).map((o) => o.value) }));
  check("porównanie domyślnie wg Status (bez filtra nie ma „zakres vs reszta”)", pick.value === "col:Status" && !pick.opts.includes("scope") && pick.opts.includes("col:Miasto"), pick);
  await page.selectOption("#rpCompareBy", "col:Miasto");
  await sleep(200);
  cmp = await sectionText(page, "compare");
  check("tabela porównania wg Miasto: wiersze, udział, średnia kwota, % otwartych", /Porównanie grup: Miasto/.test(cmp) && /Kraków\s+5\s+50%\s+110\s+100%/.test(cmp) && /Warszawa\s+5\s+50%\s+400\s+0%/.test(cmp), cmp);
  const fc = await page.evaluate(() => window.__report.findings());
  check("wniosek: średnia Kwota Warszawa 400 vs Kraków 110 — 3,6×", /Śr\. Kwota: najwyżej „Warszawa” \(400\), najniżej „Kraków” \(110\) — 3,6× różnicy/.test(fc[0]), fc);
  check("wniosek: otwarte 100% Kraków vs 0% Warszawa", fc.some((x) => /Otwartych najwięcej w „Kraków” \(100%\), najmniej w „Warszawa” \(0%\)/.test(x)), fc);

  // Zakres vs reszta: po wybraniu zakresu pojawia się i jest domyślna, gdy użytkownik nic nie wybrał.
  await page.evaluate(() => { window.__report.setCompareBy(""); window.__report.setScope({ kind: "query", q: 'Status:="W toku"' }); });
  await sleep(200);
  cmp = await sectionText(page, "compare");
  const fs2 = await page.evaluate(() => window.__report.findings());
  check("zakres (W toku) vs reszta arkusza: tytuł + liczby 3 / 7", /Porównanie: zakres na tle reszty arkusza/.test(cmp) && /Zakres\s+3\s+30%/.test(cmp) && /Reszta arkusza\s+7\s+70%/.test(cmp), cmp);
  check("wniosek zakresu: średnia kwota w zakresie 116,67 — mniej niż w reszcie", fs2.some((x) => /Śr\. Kwota w zakresie: 116,67 — 2,7× mniej niż w reszcie arkusza \(314,29\)/.test(x)), fs2);
  await page.evaluate(() => window.__report.setScope({ kind: "view" }));

  // ── 6. Strony = PDF, marginesy, silnik agregacji nietknięty ──
  await page.evaluate(() => { window.__report.setPreset("normal"); window.__report.setAngle("state"); });
  await sleep(300);
  const pages = await page.evaluate(() => ({ note: document.getElementById("rpFitNote").textContent, n: window.__report.pageCount(), bad: window.__report.marginViolations() }));
  await page.evaluate(() => document.body.classList.add("rp-printing"));
  await page.emulateMedia({ media: "print" });
  const pdf = await page.pdf({ format: "A4", preferCSSPageSize: true });
  await page.emulateMedia({ media: "screen" });
  await page.evaluate(() => document.body.classList.remove("rp-printing"));
  const realPages = (pdf.toString("latin1").match(/\/Type\s*\/Page[^s]/g) || []).length;
  check("Stan teraz + normalny: pasek stron = PDF, marginesy OK", pages.n === realPages && pages.bad.length === 0, { ...pages, realPages });
  check("panel Agregacje nietknięty", (await page.evaluate(() => window.__report.aggState())) === aggBefore);

  // ── Brak różnic = uczciwe zdanie ──
  await page.evaluate(() => window.__report.close());
  await loadRows(page, {
    headers: ["Status", "Miasto", "Kwota"],
    rows: Array.from({ length: 12 }, (_, i) => [i % 2 ? "W toku" : "Zakończone", i % 3 ? "A" : "B", 100 + (i % 3)]),
    file: "rowne.xlsx",
  });
  await page.evaluate(() => { window.__report.open(); window.__report.setAngle("compare"); window.__report.setCompareBy("col:Status"); });
  await sleep(200);
  const fn = await page.evaluate(() => window.__report.findings());
  check("brak wyraźnych różnic → jedno uczciwe zdanie, bez naciągania", /Grupy nie różnią się wyraźnie/.test(fn[0]) && !fn.some((x) => /× różnicy/.test(x)), fn);

  // ── 3. Bez statusu: para dat ──
  await page.evaluate(() => window.__report.close());
  await loadRows(page, {
    headers: ["Sprawa", "Data przyjęcia", "Data zamknięcia", "Kwota"],
    rows: [["S1", isoAgo(40), isoAgo(30), 10], ["S2", isoAgo(35), "", 20], ["S3", isoAgo(20), isoAgo(5), 30], ["S4", isoAgo(10), "", 40], ["S5", isoAgo(8), isoAgo(2), 50], ["S6", isoAgo(3), "", 60]],
    file: "daty.xlsx",
  });
  await page.evaluate(() => { window.__report.open(); window.__report.setAngle("state"); });
  await sleep(200);
  const sd = await sectionText(page, "state");
  check("para dat: reguła „otwarte = jest przyjęcie, brak zamknięcia”, 3 otwarte, najdłużej S2 (35 dni)",
    /Otwarte = jest „Data przyjęcia”, brak „Data zamknięcia”/.test(sd) && /Otwarte: 3/.test(sd) && /S2[\s\S]*35 dni/.test(sd), sd);

  // ── 4. Liczby wyświetlane jako daty z 1900 r. → brak kąta ──
  await page.evaluate(() => window.__report.close());
  await loadRows(page, {
    headers: ["Pozycja", "Kwota", "Zapłacono"],
    rows: Array.from({ length: 8 }, (_, i) => [`P${i}`, { t: "n", v: 1200 + i * 300, w: `${String(i + 1).padStart(2, "0")}.0${(i % 9) + 1}.1903` }, { t: "n", v: 900 + i * 100, w: `${String(i + 3).padStart(2, "0")}.0${(i % 9) + 1}.1902` }]),
    file: "budzet.xlsx",
  });
  await page.evaluate(() => { window.__report.open(); document.getElementById("rpContentPanel").classList.remove("hidden"); window.__report.setAngle("state"); });
  await sleep(200);
  const weird = await page.evaluate(() => ({ state: !!window.__report.data().state, dis: document.querySelector('#rpAngles [data-angle="state"]').disabled, hint: document.querySelector('#rpAngles [data-angle="state"]').getAttribute("data-hint"), sec: window.__report.sections() }));
  check("daty z 1900 r. nie udają „od kiedy czeka”: kąt niedostępny + powód w dymku, brak sekcji", !weird.state && weird.dis && /Brak w tym arkuszu kolumny stanu/.test(weird.hint) && !weird.sec.includes("state"), weird);
  await page.evaluate(() => { window.__report.setAngle("overview"); window.__report.close(); });

  // ── 7. RODO: cykle z Trybów auto ──
  if (fs.existsSync(RODO)) {
    await page.setInputFiles("#fileInput", RODO);
    await page.waitForFunction(() => document.getElementById("sheetSelect")?.options?.length > 1, null, { timeout: 30000 });
    await page.evaluate(() => {
      const s = document.getElementById("sheetSelect"); s.value = "2025-2026"; s.dispatchEvent(new Event("change", { bubbles: true }));
      document.getElementById("autoHeaderRow").checked = false; document.getElementById("headerRow").value = "2";
    });
    await sleep(300);
    await page.evaluate(() => document.getElementById("loadBtn").click());
    await page.waitForFunction(() => currentFileName.startsWith("RODO") && baseRows.length > 100 && currentDisplayModel, null, { timeout: 60000 });
    await sleep(600);
    await page.evaluate(() => { window.__report.open(); window.__report.setAngle("state"); });
    await sleep(300);
    const rodo = await page.evaluate(() => ({ s: window.__report.data().state, f: window.__report.findings()[0] }));
    const rs = await sectionText(page, "state");
    check("RODO: stan z cykli (otwarty cykl = „od” bez „do”)", rodo.s && rodo.s.mode === "records" && /Wg cykli/.test(rs) && rodo.s.counts.open > 0, rodo.s && { mode: rodo.s.mode, counts: rodo.s.counts });
    check("RODO: najdłużej otwarte z nazwiskiem z otwartego cyklu", /Najdłużej otwarte\nNr\.\tImię i Nazwisko\tod\tCzeka/.test(rs) && /Najdłużej czeka „\d+” \([^)]+\) — \d+ dni/.test(rodo.f), { rs: rs.slice(0, 400), f: rodo.f });
    await page.evaluate(() => { window.__report.setAngle("overview"); window.__report.close(); });
  }

  check("brak błędów w konsoli", errors.length === 0, errors);
  await browser.close();
  const failed = results.filter((r) => !r.ok);
  if (failed.length) {
    console.log(`\n❌ report-angles: ${failed.length} z ${results.length} nie przeszło`);
    process.exit(1);
  }
  console.log(`\n✅ report-angles: ${results.length}/${results.length}`);
}

run().catch((e) => { console.error(e); process.exit(1); });
