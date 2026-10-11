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

const pw = require("playwright");
const ENGINE = process.env.ENGINE || "chromium";
const fs = require("fs");
const path = require("path");

const APP_URL = process.env.APP_URL || "http://127.0.0.1:4175/";
const SLEEP_SCALE = Math.min(1, Math.max(0.1, Number(process.env.SLEEP_SCALE || 1)));
const sleep = (ms) => new Promise((r) => setTimeout(r, Math.round(ms * SLEEP_SCALE)));

async function run() {
  const browser = await pw[ENGINE].launch({ headless: true });
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
    await page.evaluate(() => { theadEl.replaceChildren(); tbodyEl.replaceChildren(); });
    await page.setInputFiles("#fileInput", f);
    await page.waitForFunction((n) => currentFileName === n && baseRows.length > 0 && !!document.querySelector("#dataTable thead tr.header-row") && !!document.querySelector("#dataTable tbody tr[data-row-key]"), f.name, { timeout: 15000, polling: 100 });
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
  const autoOff = await page.waitForFunction(() => !autoHeaderRowEl.checked, null, { timeout: 4000 }).then(() => true, () => false);
  ok("3. ręczny wiersz nagłówka wyłącza auto (bieżący plik)", autoOff, null);
  const s2 = await openAndWait(await mkFile("Plik B"));
  ok("3. kolejny plik znów z auto (preferencja nietknięta)", s2.auto && s2.header === 3, s2);
  // świadome odznaczenie „Auto” = preferencja na kolejne pliki (panel po wczytaniu mógł się zwinąć)
  await page.evaluate(() => { setSidebarOpen(true); document.getElementById("panel-file-sheet").open = true; });
  await sleep(300);
  await page.click("#autoHeaderRow");
  const prefSaved = await page.waitForFunction(() => localStorage.getItem("swb-auto-header-v1") === "0", null, { timeout: 4000 }).then(() => true, () => false);
  ok("3. odznaczenie „Auto” zapamiętane", prefSaved, null);
  await page.fill("#headerRow", "1");
  const s3 = await openAndWait(await mkFile("Plik C"));
  ok("3. przy wyłączonym auto nagłówek z pola (bez wykrywania)", !s3.auto && s3.header === 1, s3);

  // 4. wiersze NAD nagłówkiem (punkt A1): arkusz od wiersza 1 jak w Excelu; filtr i sortowanie
  // nie ruszają ich (jak wiersze nad autofiltrem), scalony tytuł = colspan, przełącznik w Widoku.
  {
    const ctx4 = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1280, height: 860 } });
    await ctx4.addInitScript(() => { localStorage.setItem("introPlayed", "true"); });
    const p4 = await ctx4.newPage();
    p4.on("pageerror", (e) => errors.push("pageerror(4): " + e.message));
    p4.on("dialog", (d) => d.accept()); // „niezapisane zmiany” przy wczytaniu kolejnego pliku
    await p4.goto(APP_URL, { waitUntil: "load" });
    await p4.evaluate(() => { document.getElementById("heroSplash")?.remove(); return loadSampleFile(); });
    await p4.waitForFunction(() => document.querySelector("#dataTable tbody tr[data-row-key]"), null, { timeout: 20000 });
    await sleep(400);
    const preState = () => p4.evaluate(() => {
      const pre = Array.from(document.querySelectorAll("#dataTable thead tr.pre-row"));
      return {
        nums: pre.map((tr) => tr.querySelector(".row-head").textContent).join(","),
        title: pre[0] ? pre[0].querySelector(".pre-cell").textContent : "",
        headerBelow: !!document.querySelector("#dataTable thead tr.pre-row + tr.header-row"),
        bodyRows: baseRows.length,
      };
    });
    const a = await preState();
    ok("4. przykładowy plik: wiersze 1–4 nad nagłówkiem, tytuł w A1, nagłówek pod nimi",
      a.nums === "1,2,3,4" && /Raport/.test(a.title) && a.headerBelow && a.bodyRows === 120, a);
    // filtr szybkiego szukania — wiersze nad nagłówkiem zostają
    await p4.evaluate(() => { quickSearchEl.value = "Zakres A-01"; commitQuickSearch(); });
    await p4.waitForFunction(() => viewRows.length > 0 && viewRows.length < 10, null, { timeout: 8000 });
    await sleep(300);
    const b = await preState();
    ok("4. filtr nie rusza wierszy nad nagłówkiem", b.nums === "1,2,3,4", b);
    await p4.evaluate(() => { quickSearchEl.value = ""; commitQuickSearch(); });
    await sleep(400);
    // sortowanie po kolumnie — dalej te same wiersze nad nagłówkiem
    await p4.evaluate(() => document.querySelector("#dataTable thead .header-row th[data-sort-header]").click());
    await sleep(400);
    ok("4. sortowanie nie rusza wierszy nad nagłówkiem", (await preState()).nums === "1,2,3,4", null);
    // przełącznik Widok → „Wiersze nad nagłówkiem”
    await p4.evaluate(() => document.getElementById("rowsAboveHeaderToggle").click());
    await sleep(300);
    const off = await p4.evaluate(() => ({ n: document.querySelectorAll("#dataTable thead tr.pre-row").length, pref: localStorage.getItem("swb-rows-above-header-v1"), pressed: document.getElementById("rowsAboveHeaderToggle").getAttribute("aria-pressed") }));
    ok("4. przełącznik chowa wiersze nad nagłówkiem i zapamiętuje wybór", off.n === 0 && off.pref === "0" && off.pressed === "false", off);
    await p4.evaluate(() => document.getElementById("rowsAboveHeaderToggle").click());
    await sleep(300);
    // 5. edycja komórek nad nagłówkiem (punkt A2): dwuklik → pole, Enter = zapis do arkusza
    // (pendingEdits → zapis pliku), formuła liczy się od razu, Cofnij przywraca, Podgląd blokuje.
    const preTd = (r, c) => `#dataTable thead tr.pre-row[data-row-index="${r}"] td.pre-cell[data-col-index="${c}"]`;
    await p4.dblclick(preTd(0, 0));
    await p4.waitForSelector(`${preTd(0, 0)} input.cell-editor`, { timeout: 3000 });
    await p4.keyboard.press("ControlOrMeta+a");
    await p4.keyboard.type("Raport zmieniony");
    await p4.keyboard.press("Enter");
    await sleep(300);
    const e1 = await p4.evaluate(() => ({
      text: document.querySelector('#dataTable thead tr.pre-row[data-row-index="0"] td.pre-cell[data-col-index="0"]').textContent,
      v: workbook.Sheets[currentSheetName].A1.v,
      pending: !!(pendingEdits[currentSheetName] && pendingEdits[currentSheetName].A1),
      editor: !!document.querySelector("#dataTable input.cell-editor"),
    }));
    ok("5. dwuklik + Enter zmienia komórkę nad nagłówkiem i trafia do zapisu", e1.text === "Raport zmieniony" && e1.v === "Raport zmieniony" && e1.pending && !e1.editor, e1);
    // formuła w wierszu nad nagłówkiem (C3 jest pusta): suma trzech pierwszych „Nr.”
    await p4.dblclick(preTd(2, 2));
    await p4.waitForSelector(`${preTd(2, 2)} input.cell-editor`, { timeout: 3000 });
    await p4.keyboard.type("=SUMA(A6:A8)");
    await p4.keyboard.press("Enter");
    await sleep(300);
    const e2 = await p4.evaluate(() => ({ text: document.querySelector('#dataTable thead tr.pre-row[data-row-index="2"] td.pre-cell[data-col-index="2"]').textContent, f: workbook.Sheets[currentSheetName].C3?.f }));
    ok("5. formuła nad nagłówkiem liczy się od razu", e2.text === "6" && /SUM\(A6:A8\)/i.test(e2.f || ""), e2);
    await p4.evaluate(() => undoRedo("undo"));
    await p4.evaluate(() => undoRedo("undo"));
    await sleep(300);
    const e3 = await p4.evaluate(() => ({ a1: workbook.Sheets[currentSheetName].A1.v, c3: workbook.Sheets[currentSheetName].C3 ? workbook.Sheets[currentSheetName].C3.v : null, text: document.querySelector('#dataTable thead tr.pre-row[data-row-index="0"] td.pre-cell[data-col-index="0"]').textContent }));
    ok("5. Cofnij przywraca komórki nad nagłówkiem", /Raport/.test(e3.a1) && e3.a1 !== "Raport zmieniony" && e3.c3 == null && e3.text === e3.a1, e3);
    await p4.evaluate(() => sheetEditUi.setMode("view"));
    await sleep(200);
    const viewMode = await p4.evaluate(() => typeof swbEditMode !== "undefined" ? swbEditMode : "?");
    if (viewMode === "view") {
      await p4.dblclick(preTd(0, 0));
      await sleep(200);
      ok("5. w Podglądzie komórek nad nagłówkiem nie da się edytować", (await p4.evaluate(() => !document.querySelector("#dataTable thead input.cell-editor"))), null);
    } else {
      failures.push("5. nie udało się przełączyć w Podgląd (" + viewMode + ")");
    }
    await p4.evaluate(() => sheetEditUi.setMode("edit"));
    // scalony tytuł nad tabelą → colspan; plik z nagłówkiem w wierszu 1 → brak wierszy nad
    const mk = async (name, aoa, merges) => {
      const b64 = await p4.evaluate(({ aoa, merges }) => {
        const ws = XLSX.utils.aoa_to_sheet(aoa);
        if (merges) ws["!merges"] = merges.map((m) => XLSX.utils.decode_range(m));
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "Dane");
        const u8 = new Uint8Array(XLSX.write(wb, { bookType: "xlsx", type: "array" }));
        let s = ""; for (let i = 0; i < u8.length; i++) s += String.fromCharCode(u8[i]);
        return btoa(s);
      }, { aoa, merges });
      // stara tabela zostaje w DOM do przerysowania — czyścimy ją, żeby czekać na NOWĄ
      await p4.evaluate(() => { theadEl.replaceChildren(); tbodyEl.replaceChildren(); });
      await p4.setInputFiles("#fileInput", { name, mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", buffer: Buffer.from(b64, "base64") });
      await p4.waitForFunction((n) => currentFileName === n && !!document.querySelector("#dataTable thead tr.header-row") && !!document.querySelector("#dataTable tbody tr[data-row-key]"), name, { timeout: 15000, polling: 100 });
      await sleep(300);
    };
    const data = [1, 2, 3, 4, 5].map((i) => [i, "Zakres " + i, "Osoba " + i, i * 10]);
    await mk("scalony.xlsx", [["Lista zakresów 2026"], [], ["Nr", "Zakres", "Osoba", "Kwota"], ...data], ["A1:D1"]);
    const m = await p4.evaluate(() => {
      const td = document.querySelector("#dataTable thead tr.pre-row .pre-cell");
      return { header: currentHeaderRow, colspan: td ? td.colSpan : 0, text: td ? td.textContent : "", pre: document.querySelectorAll("#dataTable thead tr.pre-row").length };
    });
    ok("4. scalony tytuł nad tabelą rysuje się jako jedna komórka na 4 kolumny", m.header === 3 && m.colspan === 4 && m.text === "Lista zakresów 2026" && m.pre === 2, m);
    // 6. zablokowane okienka z pliku (Excel: Zablokuj okienka, <pane ySplit="3" state="frozen"/>)
    // wygrywają ze zgadywaniem: dwa tekstowe wiersze jeden pod drugim — heurystyka wzięłaby
    // wyższy, autor pliku zablokował 3 wiersze, więc nagłówek = 3.
    const frozenB64 = await p4.evaluate(async () => {
      const aoa = [["Zestawienie"], ["Kod", "Opis", "Grupa", "Uwagi"], ["Nr", "Zakres", "Osoba", "Kwota"],
        ...[1, 2, 3, 4, 5].map((i) => [i, "Zakres " + i, "Osoba " + i, i * 10])];
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(aoa), "Dane");
      const zip = await JSZip.loadAsync(XLSX.write(wb, { bookType: "xlsx", type: "array" }));
      const path = "xl/worksheets/sheet1.xml";
      let xml = await zip.file(path).async("string");
      const view = '<sheetViews><sheetView workbookViewId="0"><pane ySplit="3" topLeftCell="A4" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>';
      xml = /<sheetViews>[\s\S]*?<\/sheetViews>/.test(xml) ? xml.replace(/<sheetViews>[\s\S]*?<\/sheetViews>/, view) : xml.replace(/<sheetFormatPr|<cols>|<sheetData/, (m) => view + m);
      zip.file(path, xml);
      const u8 = await zip.generateAsync({ type: "uint8array" });
      let s = ""; for (let i = 0; i < u8.length; i++) s += String.fromCharCode(u8[i]);
      return btoa(s);
    });
    const heuristicOnly = await p4.evaluate((b64) => detectHeaderRowSimple(XLSX.read(Uint8Array.from(atob(b64), (c) => c.charCodeAt(0))).Sheets.Dane), frozenB64);
    await p4.evaluate(() => { theadEl.replaceChildren(); tbodyEl.replaceChildren(); });
    await p4.setInputFiles("#fileInput", { name: "zablokowane.xlsx", mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", buffer: Buffer.from(frozenB64, "base64") });
    await p4.waitForFunction(() => currentFileName === "zablokowane.xlsx" && !!document.querySelector("#dataTable thead tr.header-row") && !!document.querySelector("#dataTable tbody tr[data-row-key]"), null, { timeout: 15000, polling: 100 });
    await sleep(300);
    const fr = await p4.evaluate(() => ({ header: currentHeaderRow, headers: currentHeaders.join(","), frozen: currentFrozenRowsBySheet.get("Dane") }));
    ok("6. zablokowane okienka z pliku wskazują nagłówek (zamiast zgadywania)", heuristicOnly !== 3 && fr.frozen === 3 && fr.header === 3 && fr.headers === "Nr,Zakres,Osoba,Kwota", { heuristicOnly, ...fr });
    await mk("od-pierwszego.xlsx", [["Nr", "Zakres", "Osoba", "Kwota"], ...data]);
    ok("4. nagłówek w wierszu 1 → nic nad tabelą", (await p4.evaluate(() => document.querySelectorAll("#dataTable thead tr.pre-row").length)) === 0, null);
    await ctx4.close();
  }

  await browser.close();
  if (errors.length) failures.push("błędy strony: " + errors.join(" | "));
  if (failures.length) {
    console.error("❌ header-detect:\n  - " + failures.join("\n  - "));
    process.exit(1);
  }
  console.log(`✅ header-detect (${ENGINE}): układy syntetyczne, prawdziwe pliki, otwieranie z wykrytym nagłówkiem, wiersze nad nagłówkiem + edycja`);
}

run().catch((e) => { console.error(e); process.exit(1); });
