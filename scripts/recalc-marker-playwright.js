// recalc-marker-playwright.js — znacznik komórek PRZELICZONYCH na dziś (formuły z TODAY()).
//
// Co musi być prawdą:
//   1. znacznik dostają DOKŁADNIE te komórki, w których przeliczenie zmieniło wyświetlaną
//      wartość względem tej zapisanej w pliku (ani jedna więcej — inaczej przestaje coś znaczyć),
//   2. podpowiedź komórki podaje starą wartość i datę zapisu pliku (Props.ModifiedDate),
//   3. po ODZNACZENIU „Przeliczaj formuły z datą" znaczników nie ma wcale, a komórki
//      wracają do wartości z pliku,
//   4. zwykłe komórki nadal pokazują podpowiedź tylko wtedy, gdy tekst jest ucięty.
//
// Fixture budowany w przeglądarce (nie zależy od prywatnego pliku Mateusza):
// arkusz z kolumną „Długość" = DATEDIF(od; TODAY(); ...) z ZAPISANYM starym wynikiem.

const { chromium } = require("playwright");
const APP_URL = process.env.APP_URL || "http://127.0.0.1:4175/";
// SLEEP_SCALE (run-tests.mjs): skraca przerwy „odczekaj po kroku” w szybkim przebiegu;
// powtórka porażki idzie ze skalą 1 (pełne przerwy). sleepFixed = przerwa, która MUSI
// przeczekać timer aplikacji (debounce, opóźnienie podpowiedzi, bezczynność) — bez skali.
const SLEEP_SCALE = Math.min(1, Math.max(0.1, Number(process.env.SLEEP_SCALE || 1)));
const sleep = (ms) => new Promise((r) => setTimeout(r, Math.round(ms * SLEEP_SCALE)));
const sleepFixed = (ms) => new Promise((r) => setTimeout(r, ms)); // eslint-disable-line no-unused-vars

async function run() {
  const browser = await chromium.launch({ headless: true });
  const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1200, height: 860 } });
  await context.addInitScript(() => localStorage.setItem("introPlayed", "true"));
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push("pageerror: " + e.message));
  page.on("console", (m) => { if (m.type() === "error") errors.push("console: " + m.text()); });

  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  await page.evaluate(() => ensureXlsxLibs(false));
  await page.waitForFunction(() => typeof XLSX !== "undefined" && typeof buildRows === "function");

  // Fixture: 3 wiersze. A=Nr, B=od (data), C=Dni (formuła z TODAY + STARY wynik w pliku),
  // D=Stałe (formuła z TODAY, ale wynik się nie zmienia), E=zwykły tekst.
  await page.evaluate(() => {
    const ws = {};
    const put = (ref, cell) => { ws[ref] = cell; };
    put("A1", { t: "s", v: "Nr" }); put("B1", { t: "s", v: "od" }); put("C1", { t: "s", v: "Dni" });
    put("D1", { t: "s", v: "Stale" }); put("E1", { t: "s", v: "Opis" });
    const serial = (d) => Math.round((d.getTime() - new Date(Date.UTC(1899, 11, 30)).getTime()) / 86400000);
    const today = new Date();
    for (let i = 0; i < 3; i++) {
      const r = i + 2;
      const from = new Date(today.getFullYear(), today.getMonth(), today.getDate() - (10 + i));
      put(`A${r}`, { t: "n", v: i + 1 });
      put(`B${r}`, { t: "n", v: serial(from), w: `${from.getFullYear()}-${from.getMonth() + 1}-${from.getDate()}` });
      // wartość „z pliku" celowo NIEAKTUALNA (99) — przeliczenie musi ją zmienić
      put(`C${r}`, { t: "n", v: 99, w: "99", f: `TODAY()-B${r}` });
      // formuła z TODAY(), której wynik zapisany w pliku jest JUŻ aktualny
      put(`D${r}`, { t: "n", v: 1, w: "1", f: `IF(TODAY()>0,1,0)` });
      put(`E${r}`, { t: "s", v: "tekst " + (i + 1) });
    }
    ws["!ref"] = "A1:E4";
    workbook = { SheetNames: ["Test"], Sheets: { Test: ws }, Props: { ModifiedDate: "2026-06-06T12:00:00.000Z" } };
    currentFileName = "fixture.xlsx";
    sheetSelect.replaceChildren();
    const opt = document.createElement("option"); opt.value = "Test"; opt.textContent = "Test";
    sheetSelect.appendChild(opt);
    sheetSelect.value = "Test";
    document.getElementById("recalcDates").checked = true;
    document.getElementById("headerRow").value = "1";
    document.getElementById("autoHeaderRow").checked = false;
  });
  await page.click("#loadBtn");
  await page.waitForFunction(() => typeof baseRows !== "undefined" && baseRows.length === 3, null, { timeout: 15000 });
  await sleep(400);

  const failures = [];
  const report = {};
  const ok = (name, cond, got) => { if (!cond) failures.push(`${name} — dostalem: ${JSON.stringify(got)}`); };

  const marked = await page.evaluate(() => {
    const cells = Array.from(document.querySelectorAll("#dataTable tbody td.cell-recalced"));
    return {
      count: cells.length,
      cols: [...new Set(cells.map((c) => c.dataset.colIndex))],
      was: [...new Set(cells.map((c) => c.dataset.recalcWas))],
      shown: [...new Set(cells.map((c) => c.textContent.trim()))],
      total: document.querySelectorAll("#dataTable tbody td[data-col-index]").length,
    };
  });
  report.marked = marked;
  ok("znacznik dostaly 3 komorki (kolumna Dni)", marked.count === 3, marked);
  ok("tylko kolumna 2 (Dni)", marked.cols.length === 1 && marked.cols[0] === "2", marked.cols);
  ok("stara wartosc z pliku zapamietana (99)", marked.was.length === 1 && marked.was[0] === "99", marked.was);
  ok("pokazana wartosc jest juz przeliczona (nie 99)", marked.shown.every((v) => v !== "99"), marked.shown);

  // Marker w rogu faktycznie się renderuje (::after ma niezerowy rozmiar)
  const pseudo = await page.evaluate(() => {
    const c = document.querySelector("#dataTable tbody td.cell-recalced");
    const cs = getComputedStyle(c, "::after");
    return { content: cs.content, borderWidth: cs.borderWidth, position: getComputedStyle(c).position };
  });
  report.pseudo = pseudo;
  ok("rog komorki ma narysowany trojkacik", pseudo.borderWidth.includes("7px") && pseudo.position === "relative", pseudo);

  // Podpowiedź: stara wartość + data zapisu pliku
  const tip = await page.evaluate(() => {
    const cell = document.querySelector("#dataTable tbody td.cell-recalced");
    showCellTooltip(cell);
    const el = document.getElementById("cellTooltip");
    return { hidden: el.classList.contains("hidden"), text: el.textContent, hasNote: !!el.querySelector(".cell-tooltip-note") };
  });
  report.tooltip = tip;
  ok("podpowiedz sie pokazuje mimo ze tekst sie miesci", !tip.hidden, tip);
  ok("podpowiedz podaje stara wartosc", tip.text.includes("99"), tip.text);
  ok("podpowiedz podaje date zapisu pliku", /2026/.test(tip.text) && /cze|jun/i.test(tip.text), tip.text);
  ok("notka ma wlasny styl", tip.hasNote, tip);

  // Zwykła, mieszcząca się komórka nadal bez podpowiedzi
  const plainTip = await page.evaluate(() => {
    hideCellTooltip();
    const cell = document.querySelector('#dataTable tbody td[data-col-index="4"]');
    showCellTooltip(cell);
    return document.getElementById("cellTooltip").classList.contains("hidden");
  });
  report.plainTooltip = { hidden: plainTip };
  ok("zwykla komorka bez ucietego tekstu nadal nie pokazuje podpowiedzi", plainTip, plainTip);

  // Odznaczenie przełącznika = brak znaczników i powrót wartości z pliku
  await page.evaluate(() => {
    hideCellTooltip();
    const cb = document.getElementById("recalcDates");
    cb.checked = false;
    cb.dispatchEvent(new Event("change"));
  });
  await sleep(500);
  const off = await page.evaluate(() => ({
    marked: document.querySelectorAll("#dataTable tbody td.cell-recalced").length,
    col2: Array.from(document.querySelectorAll('#dataTable tbody td[data-col-index="2"]')).map((c) => c.textContent.trim()),
  }));
  report.recalcOff = off;
  ok("po odznaczeniu zero znacznikow", off.marked === 0, off);
  ok("po odznaczeniu wartosci wracaja do tych z pliku", off.col2.every((v) => v === "99"), off.col2);

  console.log(JSON.stringify(report, null, 2));
  if (errors.length) failures.push("bledy strony: " + errors.join(" | "));
  await browser.close();
  if (failures.length) {
    console.error("\n❌ recalc-marker-playwright FAIL:\n" + failures.map((f) => " - " + f).join("\n"));
    process.exit(1);
  }
  console.log("✅ recalc-marker-playwright OK");
}

run().catch((e) => { console.error(e); process.exit(1); });
