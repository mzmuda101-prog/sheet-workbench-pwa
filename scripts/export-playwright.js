// export-playwright.js — test okna EKSPORT (CSV / Excel / PDF-druk).
//
// Symuluje realnego użytkownika: wczytuje arkusz, otwiera „Eksport” i sprawdza:
//   1. okno mówi, CO eksportujesz (zakres: cały arkusz / filtr + liczba wierszy), licznik kolumn,
//   2. CSV pod polskiego Excela: średnik, BOM UTF-8, CRLF, polskie znaki, tylko wybrane kolumny;
//      przecinek na życzenie — wtedy pola z przecinkiem w cudzysłowie,
//   3. Excel (.xlsx): liczby jako liczby, daty jako daty, autofiltr, arkusz „Info”,
//   4. PDF / druk: otwiera podgląd raportu w trybie tabeli (wybrane kolumny, wszystkie wiersze
//      widoku, szeroka tabela poziomo) — ten sam, co ma marginesy i „Pobierz PDF”,
//   5. wybrany format i opcje CSV są zapamiętane,
//   6. przy filtrze zakres mówi „X z Y”.

const { chromium } = require("playwright");
const fs = require("fs");
const path = require("path");
const XLSX = require(path.join(__dirname, "..", "lib", "xlsx.full.min.js"));

const APP_URL = process.env.APP_URL || "http://127.0.0.1:4175/";
const SLEEP_SCALE = Math.min(1, Math.max(0.1, Number(process.env.SLEEP_SCALE || 1)));
const sleep = (ms) => new Promise((r) => setTimeout(r, Math.round(ms * SLEEP_SCALE)));

// Arkusz z pułapkami: polskie znaki, przecinek w tekście, liczba z częścią dziesiętną, data jako
// liczba seryjna Excela wyświetlana jako data, 8 kolumn (→ PDF poziomo).
async function loadSheet(page) {
  await page.evaluate(() => {
    const ws = {};
    ["Nr", "Klient", "Miasto", "Kwota", "Data", "Status", "Uwagi", "Kod"].forEach((h, i) => { ws[String.fromCharCode(65 + i) + "1"] = { t: "s", v: h }; });
    const rows = [
      [1, "Łukasz Żółć", "Kraków", 1234.5, 46023, "W toku", "pilne, zadzwonić", "A-1"],
      [2, "Anna Ślęzak", "Gdańsk", 99, 46050, "Zakończone", "", "B-2"],
      [3, "Jan Nowak", "Łódź", 10, 46100, "W toku", "—", "C-3"],
    ];
    rows.forEach((r, i) => {
      const n = i + 2;
      ws["A" + n] = { t: "n", v: r[0], w: String(r[0]) };
      ws["B" + n] = { t: "s", v: r[1] };
      ws["C" + n] = { t: "s", v: r[2] };
      ws["D" + n] = { t: "n", v: r[3], w: String(r[3]).replace(".", ",") };
      ws["E" + n] = { t: "n", v: r[4], w: "2026-01-01", z: "yyyy-mm-dd" };
      ws["F" + n] = { t: "s", v: r[5] };
      if (r[6]) ws["G" + n] = { t: "s", v: r[6] };
      ws["H" + n] = { t: "s", v: r[7] };
    });
    ws["!ref"] = "A1:H4";
    workbook = { SheetNames: ["Zlecenia"], Sheets: { Zlecenia: ws }, Props: {} };
    currentFileName = "zlecenia.xlsx";
    sheetSelect.replaceChildren();
    const opt = document.createElement("option"); opt.value = "Zlecenia"; opt.textContent = "Zlecenia";
    sheetSelect.appendChild(opt); sheetSelect.value = "Zlecenia";
    document.getElementById("headerRow").value = "1";
    document.getElementById("autoHeaderRow").checked = false;
  });
  await page.evaluate(() => document.getElementById("loadBtn").click());
  await page.waitForFunction(() => typeof baseRows !== "undefined" && baseRows.length === 3, null, { timeout: 15000 });
  await sleep(300);
}

async function run() {
  const browser = await chromium.launch({ headless: true });
  const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1280, height: 900 }, acceptDownloads: true });
  await context.addInitScript(() => {
    localStorage.setItem("introPlayed", "true");
    if (!sessionStorage.getItem("expInit")) {
      sessionStorage.setItem("expInit", "1");
      localStorage.removeItem("swb-export-prefs");
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

  const boot = async () => {
    await page.goto(APP_URL, { waitUntil: "load" });
    await page.evaluate(() => document.getElementById("heroSplash")?.remove());
    await page.evaluate(() => ensureXlsxLibs(false));
    await page.waitForFunction(() => typeof buildRows === "function" && window.__report);
    await loadSheet(page);
  };
  await boot();

  const open = async () => {
    await page.evaluate(() => document.getElementById("exportCsvBtn").click());
    await sleep(200);
  };
  const pick = (n) => page.evaluate((n) => {
    document.querySelectorAll("#exportColumnList input[type=checkbox]").forEach((c, i) => { c.checked = i < n; });
    document.getElementById("exportColumnList").dispatchEvent(new Event("change", { bubbles: true }));
  }, n);
  const download = async () => {
    const dl = page.waitForEvent("download");
    await page.click("#exportRunAction");
    const d = await dl;
    return { name: d.suggestedFilename(), buf: fs.readFileSync(await d.path()) };
  };

  // ── 1. Okno ──────────────────────────────────────────────────────────────
  await open();
  const modal = await page.evaluate(() => ({
    visible: !document.getElementById("exportModal").classList.contains("hidden"),
    cols: document.querySelectorAll("#exportColumnList input[type=checkbox]").length,
    allChecked: Array.from(document.querySelectorAll("#exportColumnList input[type=checkbox]")).every((c) => c.checked),
    scope: document.getElementById("exportScope").textContent,
    count: document.getElementById("exportColCount").textContent,
    format: document.querySelector('#exportFormats [aria-checked="true"]')?.dataset.format,
    run: document.getElementById("exportRunAction").textContent,
    sepAuto: document.querySelector('#exportCsvSep option[value="auto"]').textContent,
  }));
  check("okno otwarte, 8 kolumn zaznaczonych", modal.visible && modal.cols === 8 && modal.allChecked, modal);
  check("zakres: cały arkusz, 3 wiersze", /cały arkusz: 3 wierszy/.test(modal.scope), modal.scope);
  check("licznik kolumn 8 z 8", /8 z 8/.test(modal.count), modal.count);
  check("domyślnie CSV, auto = średnik", modal.format === "csv" && /Pobierz CSV/.test(modal.run) && /;/.test(modal.sepAuto), modal);

  // ── 2. CSV pod polskiego Excela ──────────────────────────────────────────
  await pick(4);
  const cnt = await page.evaluate(() => document.getElementById("exportColCount").textContent);
  check("licznik reaguje: 4 z 8", /4 z 8/.test(cnt), cnt);
  const csv1 = await download();
  const text1 = csv1.buf.toString("utf8");
  check("CSV: BOM UTF-8 na początku", csv1.buf[0] === 0xef && csv1.buf[1] === 0xbb && csv1.buf[2] === 0xbf, [...csv1.buf.slice(0, 3)]);
  const lines1 = text1.replace(/^﻿/, "").split("\r\n");
  check("CSV: średnik, 4 kolumny, CRLF", lines1[0] === "Nr;Klient;Miasto;Kwota" && lines1.length === 5, lines1.slice(0, 2));
  check("CSV: polskie znaki i przecinek dziesiętny bez cudzysłowu", lines1[1] === "1;Łukasz Żółć;Kraków;1234,5", lines1[1]);
  check("CSV: nazwa pliku", csv1.name === "zlecenia_Zlecenia.csv", csv1.name);

  await open();
  await page.selectOption("#exportCsvSep", "comma");
  await page.evaluate(() => { const c = document.getElementById("exportCsvBom"); c.checked = false; c.dispatchEvent(new Event("change")); });
  const csv2 = await download();
  const lines2 = csv2.buf.toString("utf8").split("\r\n");
  check("CSV przecinek: bez BOM", csv2.buf[0] !== 0xef, [...csv2.buf.slice(0, 3)]);
  check("CSV przecinek: pole z przecinkiem w cudzysłowie", lines2[1].includes('"1234,5"') && lines2[1].includes('"pilne, zadzwonić"'), lines2[1]);

  // ── 3. Excel ─────────────────────────────────────────────────────────────
  await open();
  const rem = await page.evaluate(() => ({ sep: document.getElementById("exportCsvSep").value, bom: document.getElementById("exportCsvBom").checked }));
  check("zapamiętane opcje CSV (przecinek, bez BOM)", rem.sep === "comma" && rem.bom === false, rem);
  await page.click('#exportFormats [data-format="xlsx"]');
  const csvOptsHidden = await page.evaluate(() => document.getElementById("exportCsvOptions").classList.contains("hidden"));
  check("Excel: opcje CSV schowane, przycisk „Pobierz Excel”", csvOptsHidden && /Pobierz Excel/.test(await page.textContent("#exportRunAction")), csvOptsHidden);
  const x = await download();
  const wb = XLSX.read(x.buf, { type: "buffer", cellNF: true });
  const ws = wb.Sheets[wb.SheetNames[0]];
  check("Excel: arkusze Zlecenia + Info", wb.SheetNames.join() === "Zlecenia,Info", wb.SheetNames);
  check("Excel: kwota jako liczba", ws.D2 && ws.D2.t === "n" && ws.D2.v === 1234.5, ws.D2);
  check("Excel: data jako data (seryjna + format)", ws.E2 && ws.E2.t === "n" && ws.E2.v === 46023 && /y/.test(ws.E2.z || ""), ws.E2);
  check("Excel: tekst z polskimi znakami", ws.B2 && ws.B2.v === "Łukasz Żółć", ws.B2);
  check("Excel: autofiltr w nagłówku", ws["!autofilter"] && ws["!autofilter"].ref === "A1:H4", ws["!autofilter"]);
  check("Excel: Info opisuje zakres", /bez filtra/.test(XLSX.utils.sheet_to_csv(wb.Sheets.Info)), XLSX.utils.sheet_to_csv(wb.Sheets.Info));

  // ── 4. PDF / druk → podgląd w trybie tabeli ──────────────────────────────
  await open();
  await page.click('#exportFormats [data-format="pdf"]');
  await page.click("#exportRunAction");
  await sleep(400);
  const pdf = await page.evaluate(() => ({
    open: window.__report.isOpen(),
    mode: window.__report.mode(),
    orient: document.getElementById("rpPage").dataset.orient,
    ths: document.querySelectorAll('#rpPage .rp-data thead th').length,
    rows: document.querySelectorAll('#rpPage .rp-data tbody tr:not(.rp-gap-row):not(.rp-head-repeat)').length,
    contentBtnHidden: document.getElementById("rpContentBtn").classList.contains("hidden"),
    modalHidden: document.getElementById("exportModal").classList.contains("hidden"),
    title: document.querySelector("#rpPage .rp-title").textContent,
  }));
  check("PDF: podgląd otwarty w trybie tabeli", pdf.open && pdf.mode === "table" && pdf.modalHidden, pdf);
  check("PDF: 8 kolumn → strona pozioma", pdf.ths === 8 && pdf.orient === "landscape", pdf);
  check("PDF: wszystkie wiersze widoku (3)", pdf.rows === 3, pdf.rows);
  check("PDF: bez wyboru sekcji, tytuł „Tabela — …”", pdf.contentBtnHidden && /Tabela — zlecenia/.test(pdf.title), pdf);
  const pdfFile = await page.evaluate(async () => { const r = await window.__report.downloadPdf({ deliver: false }); return r ? { pages: r.pages, size: r.blob.size, head: await r.blob.slice(0, 8).text() } : null; });
  check("PDF: „Pobierz PDF” tworzy plik", pdfFile && pdfFile.pages >= 1 && pdfFile.head.startsWith("%PDF-1.4") && pdfFile.size > 10000, pdfFile);
  await page.evaluate(() => window.__report.close());

  // ── 5./6. Zapamiętany format + zakres przy filtrze ───────────────────────
  await page.evaluate(() => {
    searchQueryEl.value = 'Status:="W toku"';
    filterOperatorsEl.checked = true;
    syncQuickSearchOperatorsControls();
    filtersCommitted = true;
    applyFilters();
    sortRows();
    scheduleViewRefresh({ table: true, immediate: true });
  });
  await open();
  const st = await page.evaluate(() => ({
    format: document.querySelector('#exportFormats [aria-checked="true"]')?.dataset.format,
    scope: document.getElementById("exportScope").textContent,
    filtered: document.getElementById("exportScope").classList.contains("is-filtered"),
  }));
  check("zapamiętany format: PDF", st.format === "pdf", st);
  check("zakres przy filtrze: Status = W toku — 2 z 3", /Status = W toku — 2 z 3/.test(st.scope) && st.filtered, st);
  await page.keyboard.press("Escape");

  check("brak błędów w konsoli", errors.length === 0, errors);
  await browser.close();
  const failed = results.filter((r) => !r.ok);
  if (failed.length) {
    console.log(`\n❌ export: ${failed.length} z ${results.length} nie przeszło`);
    process.exit(1);
  }
  console.log(`\n✅ export: ${results.length}/${results.length}`);
}

run().catch((e) => { console.error(e); process.exit(1); });
