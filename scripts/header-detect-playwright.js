// header-detect-playwright.js — wykrywanie wiersza nagłówka (detectHeaderRowSimple) +
// auto-wykrywanie przy otwarciu pliku (domyślnie WŁĄCZONE od 2026-10-11).
//
// Co musi być prawdą:
//   1. układy syntetyczne: lista samych tekstów → 1 (dawniej ostatni wiersz!), tytuł + KPI
//      nad tabelą → wiersz nagłówka, etykiety grup („Okres A”) nad nagłówkiem → nagłówek,
//      druga tabela obok zaczynająca się niżej → 1, tytuł w 1. i nagłówek w 3. wierszu → 3,
//   2. prawdziwe pliki (jeśli leżą w scripts/): RODO, Kosztorys, stress — znane nagłówki,
//   3. nowy plik z tytułem nad tabelą otwiera się OD RAZU z właściwym nagłówkiem
//      (auto zaznaczone), świadome odznaczenie „Auto” jest pamiętane dla kolejnych plików,
//      a ręczna zmiana wiersza nagłówka wyłącza auto tylko dla bieżącego pliku.
//
// Uruchom z serwerem na APP_URL (domyślnie http://127.0.0.1:4175/).

const { chromium } = require("playwright");
const fs = require("fs");
const path = require("path");

const APP_URL = process.env.APP_URL || "http://127.0.0.1:4175/";
const SLEEP_SCALE = Math.min(1, Math.max(0.1, Number(process.env.SLEEP_SCALE || 1)));
const sleep = (ms) => new Promise((r) => setTimeout(r, Math.round(ms * SLEEP_SCALE)));

async function run() {
  const browser = await chromium.launch({ headless: true });
  const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1280, height: 860 } });
  await context.addInitScript(() => {
    if (sessionStorage.getItem("__hdInit")) return;
    sessionStorage.setItem("__hdInit", "1");
    localStorage.clear();
    localStorage.setItem("introPlayed", "true");
  });
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push("pageerror: " + e.message));
  const failures = [];
  const ok = (name, cond, got) => { if (!cond) failures.push(`${name} — dostałem: ${JSON.stringify(got)}`); };

  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(() => { document.getElementById("heroSplash")?.remove(); return ensureXlsxLibs(false); });
  await page.waitForFunction(() => typeof XLSX !== "undefined" && typeof detectHeaderRowSimple === "function");

  // 1. układy syntetyczne
  const synth = await page.evaluate(() => {
    const det = (aoa) => detectHeaderRowSimple(XLSX.utils.aoa_to_sheet(aoa));
    const names = ["Anna Kowalska", "Jan Nowak", "Piotr Wiśniewski", "Maria Wójcik", "Tomasz Lewandowski", "Katarzyna Zielińska"];
    const allText = [["Imię i nazwisko", "Dział", "Miasto"], ...names.map((n, i) => [n, ["HR", "IT", "Biuro"][i % 3], ["Kraków", "Gdańsk"][i % 2]])];
    const titleKpi = [["Raport sprzedaży 2026"], [], ["Suma", 1234, "Średnia", 56.7], [], ["Data", "Klient", "Produkt", "Kwota", "Status"],
      ...names.map((n, i) => [new Date(2026, 0, i + 1), n, "Produkt " + i, 100 + i, i % 2 ? "Zapłacone" : "Otwarte"])];
    const groups = [[null, null, "Okres A", null, "Okres B", null], ["ID", "Typ", "Start A", "Koniec A", "Start B", "Koniec B"],
      ...names.map((_, i) => [i + 1, "Typ " + (i % 3), 45000 + i, 45010 + i, 45100 + i, 45120 + i])];
    const twoTables = [["Nazwisko", null, "Imię", null, "Skrót"], ...names.map((n, i) => {
      const [im, na] = n.split(" ");
      const row = [na, null, im, null, `${na[0]}. ${im}`];
      if (i === 3) row.push(null, null, "GRUPA 2", null, "GRUPA 3");
      if (i > 3) row.push(null, null, "x", null, "y");
      return row;
    })];
    const title3 = [["Lista zakresów — stan na dziś"], [], ["Nr", "Zakres", "Osoba", "Od"], ...names.map((n, i) => [i + 1, "Zakres " + i, n, new Date(2026, 1, i + 1)])];
    const sampleWb = XLSX.read(buildSampleWorkbookArrayBuffer(), { cellDates: true });
    const sample = detectHeaderRowSimple(sampleWb.Sheets[sampleWb.SheetNames[0]]);
    return { allText: det(allText), titleKpi: det(titleKpi), groups: det(groups), twoTables: det(twoTables), title3: det(title3), sample };
  });
  ok("1. przykładowy plik (KPI nad tabelą) → 5", synth.sample === 5, synth);
  ok("1. lista samych tekstów → 1", synth.allText === 1, synth);
  ok("1. tytuł + KPI nad tabelą → 5", synth.titleKpi === 5, synth);
  ok("1. etykiety grup nad nagłówkiem → 2", synth.groups === 2, synth);
  ok("1. druga tabela obok, niżej → 1", synth.twoTables === 1, synth);
  ok("1. tytuł w 1., nagłówek w 3. → 3", synth.title3 === 3, synth);

  // 2. prawdziwe pliki (prywatne pliki Mateusza mogą nie istnieć — wtedy pomijane)
  const REAL = {
    "RODO Obieg terenów 2026 V5.xlsx": { "2025-2026": 2, "kolumny pomocnicze": 2, "tabela przestawna": 1, "Baza danych": 1, "Analiza Pracowników": 4, "Wykresy": 4 },
    "Kosztorys Renault Laguna 3 .xlsx": { "Budżet podróży służbowej": 5 },
    "stress-test-workbench.xlsx": { WideToLong_Kwartaly: 1, MergedHeaders_Okresy: 2, EdgeCases_Klucze: 1, SuffixCycle_Osoba123: 1, DuzeN_8000wierszy: 1 },
  };
  for (const [file, want] of Object.entries(REAL)) {
    const fp = path.join(__dirname, file);
    if (!fs.existsSync(fp)) continue;
    const got = await page.evaluate(({ b64, names }) => {
      const wb = XLSX.read(Uint8Array.from(atob(b64), (c) => c.charCodeAt(0)), { cellDates: true });
      return Object.fromEntries(names.map((n) => [n, wb.Sheets[n] ? detectHeaderRowSimple(wb.Sheets[n]) : null]));
    }, { b64: fs.readFileSync(fp).toString("base64"), names: Object.keys(want) });
    Object.entries(want).forEach(([sheet, row]) => ok(`2. ${file} › ${sheet} → ${row}`, got[sheet] === row, got[sheet]));
  }

  // 3. przepływ: nowy plik otwiera się od razu z wykrytym nagłówkiem
  const mkFile = async (title) => {
    const b64 = await page.evaluate((title) => {
      const aoa = [[title], [], ["Nr", "Zakres", "Osoba", "Kwota"], ...[1, 2, 3, 4, 5].map((i) => [i, "Zakres " + i, "Osoba " + i, i * 10])];
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(aoa), "Dane");
      const u8 = new Uint8Array(XLSX.write(wb, { bookType: "xlsx", type: "array" }));
      let s = ""; for (let i = 0; i < u8.length; i++) s += String.fromCharCode(u8[i]);
      return btoa(s);
    }, title);
    return { name: `${title}.xlsx`, mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", buffer: Buffer.from(b64, "base64") };
  };
  const openAndWait = async (f) => {
    await page.setInputFiles("#fileInput", f);
    await page.waitForFunction((n) => currentFileName === n && baseRows.length > 0 && !!document.querySelector("#dataTable tbody tr[data-row-key]"), f.name, { timeout: 15000, polling: 100 });
    await sleep(300);
    return page.evaluate(() => ({ header: currentHeaderRow, headers: currentHeaders.slice(0, 4), auto: autoHeaderRowEl.checked, rows: baseRows.length }));
  };
  const s1 = await openAndWait(await mkFile("Plik A"));
  ok("3. nowy plik: nagłówek wykryty sam (wiersz 3), auto zaznaczone", s1.header === 3 && s1.headers.join() === "Nr,Zakres,Osoba,Kwota" && s1.auto && s1.rows === 5, s1);
  // ręczna zmiana wiersza nagłówka = auto WYŁ. tylko dla tego pliku
  await page.evaluate(() => setSidebarOpen(true));
  await page.evaluate(() => { document.getElementById("panel-file-sheet").open = true; });
  await sleep(300);
  await page.fill("#headerRow", "1");
  await sleep(100);
  ok("3. ręczny wiersz nagłówka wyłącza auto (bieżący plik)", (await page.evaluate(() => autoHeaderRowEl.checked)) === false, null);
  const s2 = await openAndWait(await mkFile("Plik B"));
  ok("3. kolejny plik znów z auto (preferencja nietknięta)", s2.auto && s2.header === 3, s2);
  // świadome odznaczenie „Auto” = preferencja na kolejne pliki (panel po wczytaniu mógł się zwinąć)
  await page.evaluate(() => { setSidebarOpen(true); document.getElementById("panel-file-sheet").open = true; });
  await sleep(300);
  await page.click("#autoHeaderRow");
  await sleep(100);
  ok("3. odznaczenie „Auto” zapamiętane", (await page.evaluate(() => localStorage.getItem("swb-auto-header-v1"))) === "0", null);
  await page.fill("#headerRow", "1");
  const s3 = await openAndWait(await mkFile("Plik C"));
  ok("3. przy wyłączonym auto nagłówek z pola (bez wykrywania)", !s3.auto && s3.header === 1, s3);

  await browser.close();
  if (errors.length) failures.push("błędy strony: " + errors.join(" | "));
  if (failures.length) {
    console.error("❌ header-detect:\n  - " + failures.join("\n  - "));
    process.exit(1);
  }
  console.log("✅ header-detect: układy syntetyczne, prawdziwe pliki, otwieranie z wykrytym nagłówkiem");
}

run().catch((e) => { console.error(e); process.exit(1); });
