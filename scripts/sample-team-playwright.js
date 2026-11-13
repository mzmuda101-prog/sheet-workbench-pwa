// sample-team-playwright.js — arkusz „Zespół” z przykładowego pliku (2026-11-13): to, co aplikacja
// liczy jak Excel 365, działa od razu po kliknięciu „Przykład”.
//
// Co musi być prawdą:
//   1. arkusz otwiera się z nagłówkiem tabeli Excela (wiersz 4), dni tygodnia po polsku,
//   2. wpis w „Szukaj osoby” (B2, nad nagłówkiem) → liczba przydziałów obok i podświetlony wiersz osoby,
//   3. nowa osoba z nowym działem w tabeli → kolumny obliczeniowe w nowym wierszu, lista „Działy”
//      (SORTUJ(UNIKATOWE)) i „Bez przydziałów” (FILTRUJ) same się wydłużają,
//   4. plik przykładu przechodzi walidator Office (gdy jest .NET).

const pw = require("playwright");
const { autoAnswerAppDialogs } = require("./app-dialog-auto");
const path = require("path");
const fs = require("fs");
const os = require("os");
let ooxml = null;
try { ooxml = require("./ooxml-validate"); } catch { ooxml = null; }

const APP_URL = process.env.APP_URL || "http://127.0.0.1:4175/";
const ENGINE = process.env.ENGINE || "chromium";
const SLEEP_SCALE = Math.min(1, Math.max(0.1, Number(process.env.SLEEP_SCALE || 1)));
const sleep = (ms) => new Promise((r) => setTimeout(r, Math.round(ms * SLEEP_SCALE)));
const PL_DAYS = ["niedziela", "poniedziałek", "wtorek", "środa", "czwartek", "piątek", "sobota"];

async function run() {
  const browser = await pw[ENGINE].launch({ headless: true });
  const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1400, height: 900 } });
  await context.addInitScript(() => { localStorage.setItem("introPlayed", "true"); });
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  await autoAnswerAppDialogs(page);
  const failures = [];
  const ok = (name, cond, got) => { if (!cond) failures.push(`${name} — dostałem: ${JSON.stringify(got)}`); };

  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(() => { document.getElementById("heroSplash")?.remove(); return ensureXlsxLibs(false); });
  await page.evaluate(() => loadSampleFile());
  await page.waitForFunction(() => typeof workbook !== "undefined" && workbook && workbook.SheetNames.includes("Zespół") && !!document.querySelector("#dataTable tbody tr[data-row-key]"), null, { timeout: 30000 });

  // 4. walidator na pliku przykładu
  if (ooxml && ooxml.available && ooxml.available()) {
    const tmp = fs.mkdtempSync(path.join(os.tmpdir(), "swb-sample-"));
    const out = path.join(tmp, "przyklad.xlsx");
    fs.writeFileSync(out, Buffer.from(await page.evaluate(async () => Array.from(new Uint8Array(await buildSampleWorkbookArrayBuffer())))));
    const [res] = ooxml.validate([out]);
    ok("4. przykład bez błędów schematu Office", res && !res.error && !(res.errors || []).length, (res.errors || []).slice(0, 3).map(ooxml.short));
    fs.rmSync(tmp, { recursive: true, force: true });
  }

  await page.evaluate(() => { sheetSelect.value = "Zespół"; loadBtn.click(); });
  await page.waitForFunction(() => currentSheetName === "Zespół" && !!document.querySelector("#dataTable tbody tr[data-row-key]"), null, { timeout: 30000 });
  await page.evaluate(() => sheetEditUi.setMode("edit"));
  await sleep(400);

  // 1.
  const s1 = await page.evaluate(() => {
    const ws = workbook.Sheets[currentSheetName];
    const days = [5, 6, 7].map((r) => ({ v: ws["D" + r].v, joined: ws["C" + r].v }));
    return { header: currentHeaderRow, headers: currentHeaders.slice(0, 6), days: days.map((d) => d.v), joined: days.map((d) => (d.joined instanceof Date ? d.joined.getDay() : null)) };
  });
  ok("1. nagłówek tabeli w wierszu 4", s1.header === 4 && s1.headers.join() === "Osoba,Dział,Od kiedy,Dzień tygodnia,Staż (dni),Przydziały", s1);
  ok("1. dni tygodnia po polsku (TEKST(...;\"dddd\"))", s1.days.every((d, i) => d === PL_DAYS[s1.joined[i]]), s1);

  // 2. „Szukaj osoby” — B2 nad nagłówkiem, zwykłą drogą edycji
  await page.evaluate(() => {
    const view = swbSelView();
    const pre = view.rows.find((r) => r.rowIndex0 === 1);
    setFocusedCell(getRowSelectionKey(pre), 1, { scroll: true });
  });
  await page.keyboard.press("Enter");
  await page.waitForSelector("input.cell-editor", { timeout: 3000 });
  await page.keyboard.type("Maria Wójcik");
  await page.keyboard.press("Tab");
  await sleep(700);
  await page.keyboard.press("Escape");
  const s2 = await page.evaluate(() => {
    const ws = workbook.Sheets[currentSheetName];
    const rowOf = (name) => Object.keys(ws).find((k) => /^A\d+$/.test(k) && ws[k].v === name);
    const r0 = Number(rowOf("Maria Wójcik").slice(1)) - 1;
    const bg = (r) => document.querySelector(`#dataTable tbody tr[data-row-key="wide:${r}"] td[data-col-index="0"]`)?.style.background || "";
    return { e2: ws.E2 && ws.E2.v, mine: bg(r0), other: bg(r0 - 1) };
  });
  ok("2. obok pojawia się liczba przydziałów osoby", typeof s2.e2 === "number" && s2.e2 > 0, s2);
  ok("2. wiersz osoby podświetlony (formatowanie warunkowe z formułą), inne nie", /255, 230, 153/.test(s2.mine) && !/255, 230, 153/.test(s2.other), s2);

  // 3. nowa osoba z nowym działem
  const before = await page.evaluate(() => { const ws = workbook.Sheets[currentSheetName]; return { h: ws.H5.F, j: ws.J5.F }; });
  await page.click("#addRowBtn"); await sleep(300);
  await page.keyboard.type("Zofia Nowicka"); await page.keyboard.press("Tab"); await sleep(300);
  await page.keyboard.type("Kadry"); await page.keyboard.press("Tab"); await sleep(800);
  await page.keyboard.press("Escape");
  const s3 = await page.evaluate(() => {
    const ws = workbook.Sheets[currentSheetName];
    const col = (L) => { const g = XLSX.utils.decode_range(ws[L + "5"].F); const out = []; for (let r = g.s.r; r <= g.e.r; r++) out.push(ws[L + (r + 1)] && ws[L + (r + 1)].v); return out; };
    const last = baseRows[baseRows.length - 1].rowIndex0 + 1;
    return { h: ws.H5.F, depts: col("H"), idle: col("J"), fNew: ws["F" + last] && ws["F" + last].f, tbl: currentTables.zespol.ref };
  });
  ok("3. nowy wiersz w tabeli z formułą kolumny „Przydziały”", !!s3.fNew && /COUNTIF/.test(s3.fNew) && /A4:F15$/.test(s3.tbl), s3);
  ok("3. „Działy” (SORTUJ(UNIKATOWE)) wydłużają się o „Kadry”, w kolejności", s3.depts.includes("Kadry") && s3.h !== before.h && s3.depts.join() === [...s3.depts].sort((a, b) => a.localeCompare(b, "pl")).join(), s3);
  ok("3. „Bez przydziałów” (FILTRUJ) obejmuje nową osobę", s3.idle.includes("Zofia Nowicka"), s3);

  await browser.close();
  if (errors.length) failures.push("błędy strony: " + errors.join(" | "));
  if (failures.length) {
    console.error(`❌ sample-team (${ENGINE}):\n  - ` + failures.join("\n  - "));
    process.exit(1);
  }
  console.log(`✅ sample-team (${ENGINE}): arkusz „Zespół” — tabela, TEKST po polsku, szukaj z podświetleniem, listy rozlewane`);
}

run().catch((e) => { console.error(e); process.exit(1); });
