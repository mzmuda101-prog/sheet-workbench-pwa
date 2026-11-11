// ghost-rows-playwright.js — „duchy”: blade puste wiersze pod danymi (app/row-grow.js).
//   node scripts/ghost-rows-playwright.js            (Chromium)
//   ENGINE=webkit node scripts/ghost-rows-playwright.js
//
// Pilnuje:
//   1. w Edycji 3 duchy z kolejnymi numerami wierszy Excela; w Podglądzie ich nie ma;
//   2. duchy nie są danymi: liczba wierszy, wiersze widoku i zaznaczenie „wszystko” bez zmian,
//      komórki duchów bez data-col-index (selektory komórek danych ich nie łapią);
//   3. klik w komórkę ducha = nowy wiersz i pisanie w TEJ kolumnie (wygląd i formuły jak przy
//      „＋ Nowy wiersz”), numery duchów przesuwają się w dół;
//   4. kolumny duchów stoją równo z nagłówkami (ta sama <colgroup>), także po przewinięciu w bok;
//   5. Cofnij wpisów w duchach nie zostawia pasa pustych wierszy — duchy wracają tuż pod dane;
//   6. 2 duchy-kolumny z prawej (litery, „＋”): klik w puste pole za wierszem = nowa kolumna i pisanie w nim.

const { chromium, webkit } = require("playwright");

const APP_URL = process.env.APP_URL || "http://127.0.0.1:4175/";
const ENGINE_NAME = process.env.ENGINE === "webkit" ? "webkit" : "chromium";
const ENGINE = ENGINE_NAME === "webkit" ? webkit : chromium;
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

async function run() {
  const fails = [];
  const check = (name, ok, detail) => {
    console.log(`${ok ? "✅" : "❌"} ${name}${ok ? "" : ` — ${JSON.stringify(detail)}`}`);
    if (!ok) fails.push(name);
  };
  const browser = await ENGINE.launch({ headless: true });
  const ctx = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1280, height: 900 } });
  await ctx.addInitScript(() => { localStorage.setItem("introPlayed", "true"); });
  const p = await ctx.newPage();
  const errors = [];
  p.on("pageerror", (e) => errors.push(e.message));
  await p.goto(APP_URL, { waitUntil: "load" });
  await p.evaluate(() => { document.getElementById("heroSplash")?.remove(); return loadSampleFile(); });
  await p.waitForFunction(() => document.querySelector("#dataTable tbody tr[data-row-key]"), null, { timeout: 20000 });
  await sleep(600);
  await p.evaluate(() => { sheetEditUi.setMode("edit"); document.querySelectorAll(".toast").forEach((t) => t.remove()); tableWrapEl.scrollTop = tableWrapEl.scrollHeight; });
  await sleep(400);

  const ghosts = () => p.evaluate(() => {
    const g = document.querySelector("#dataTable tbody.ghost-body");
    return {
      shown: !!g && g.offsetParent !== null,
      nums: g ? [...g.querySelectorAll(".ghost-head")].map((x) => x.textContent) : [],
      colIdx: g ? g.querySelectorAll("td[data-col-index]").length : -1,
      afterData: !!g && g.previousElementSibling === tbodyEl,
    };
  });
  const last = await p.evaluate(() => baseRows.reduce((m, r) => Math.max(m, r.rowIndex0), 0) + 1); // 1-based

  // 1.
  const g1 = await ghosts();
  check("1. w Edycji 3 duchy z kolejnymi numerami, zaraz za danymi", g1.shown && g1.afterData && g1.nums.join() === [last + 1, last + 2, last + 3].join(), { g1, last });
  await p.evaluate(() => sheetEditUi.setMode("view")); await sleep(200);
  check("1. w Podglądzie duchów nie ma", !(await ghosts()).shown);
  await p.evaluate(() => sheetEditUi.setMode("edit")); await sleep(200);

  // 2.
  const s2 = await p.evaluate(() => ({ view: viewRows.length, base: baseRows.length, status: (document.getElementById("status") || {}).textContent || "" }));
  await p.evaluate(() => { const c = document.querySelector("#dataTable thead tr.guide-row th.corner-cell"); c && c.click(); });
  await sleep(200);
  const all = await p.evaluate(() => { const r = getSelectionRectangle(); return r && r.rowCount; });
  check("2. duchy nie są danymi (zaznacz wszystko = tylko wiersze danych, bez data-col-index)",
    g1.colIdx === 0 && all === (await p.evaluate(() => swbSelView().rows.length)), { colIdx: g1.colIdx, all, s2 });

  // 3. klik w ducha (2. wiersz, kolumna D = 3)
  await p.click("#dataTable tbody.ghost-body tr:nth-child(2) td[data-ghost-col='3']");
  await sleep(300);
  const s3 = await p.evaluate(() => ({ focus: focusedCellState && `${focusedCellState.rowKey}:${focusedCellState.colIndex0}`, editor: !!document.querySelector("input.cell-editor") }));
  check("3. klik w ducha = nowy wiersz, pisanie w tej kolumnie", s3.focus === `wide:${last}:3` && s3.editor, { s3, last });
  await p.keyboard.type("Jan Duch"); await p.keyboard.press("Tab"); await sleep(300);
  const s3b = await p.evaluate((r) => { const ws = workbook.Sheets[currentSheetName]; return { D: ws[`D${r}`]?.v, Gf: !!ws[`G${r}`]?.f, base: baseRows.length }; }, last + 1);
  check("3. wpis w nowym wierszu, formuła „Długość” dopisana", s3b.D === "Jan Duch" && s3b.Gf && s3b.base === s2.base + 1, s3b);
  const g3 = await ghosts();
  check("3. duchy przesunęły się o wiersz", g3.nums.join() === [last + 2, last + 3, last + 4].join(), g3);

  // 5. scenariusz z filmu 2026-11-10: wpisy w trzech duchach, potem Cofnij ×3 → puste wiersze NIE
  //    zostają (dawniej pod danymi rósł pas pustych wierszy, a duchy zjeżdżały coraz niżej)
  await p.keyboard.press("Escape");
  for (const txt of ["jdsbad", "snda"]) {
    await p.click("#dataTable tbody.ghost-body tr:nth-child(1) td[data-ghost-col='2']");
    await sleep(250);
    await p.keyboard.type(txt); await p.keyboard.press("Tab"); await sleep(250);
  }
  const before5 = await p.evaluate(() => baseRows.length);
  for (let i = 0; i < 3; i++) { await p.evaluate(() => undoLast()); await sleep(250); }
  await p.evaluate(() => { setFocusedCell(getRowSelectionKey(baseRows[0]), 1, { scroll: false }); renderActiveTable(); });
  await sleep(300);
  const s5 = await p.evaluate(() => ({ base: baseRows.length, view: viewRows.length, nums: [...document.querySelectorAll("tbody.ghost-body .ghost-head")].map((x) => x.textContent) }));
  check("5. po Cofnij puste nowe wiersze znikają, duchy wracają tuż pod dane", before5 === s2.base + 3 && s5.base === s2.base && s5.view === s2.view && s5.nums.join() === [last + 1, last + 2, last + 3].join(), { before5, s5, s2, last });
  for (let i = 0; i < 3; i++) { await p.evaluate(() => redoLast()); await sleep(250); }
  const s5b = await p.evaluate((r) => ({ base: baseRows.length, D: [r, r + 1, r + 2].map((n) => { const row = baseRows.find((x) => x.rowIndex0 === n - 1); return row ? row.values[2] || row.values[3] : null; }) }), last + 1);
  check("5. Ponów przywraca wiersze także w tabeli", s5b.base === s2.base + 3 && s5b.D.every(Boolean), s5b);

  // 4. kolumny równo z nagłówkami (także po przewinięciu w bok)
  const misaligned = async () => p.evaluate(() => {
    let bad = 0;
    const tr = document.querySelector("#dataTable tbody.ghost-body tr");
    if (!tr) return -1;
    tr.querySelectorAll("td[data-ghost-col]").forEach((td) => {
      const th = document.querySelector(`#dataTable thead tr.header-row th[data-col-index="${td.dataset.ghostCol}"]`);
      if (th && Math.abs(th.getBoundingClientRect().left - td.getBoundingClientRect().left) > 1.5) bad++;
    });
    return bad;
  });
  const m0 = await misaligned();
  await p.evaluate(() => { tableWrapEl.scrollLeft = 300; }); await sleep(300);
  const m1 = await misaligned();
  check("4. kolumny duchów równo z nagłówkami (także przewinięte w bok)", m0 === 0 && m1 === 0, { m0, m1 });

  // 6. duchy-kolumny: 2 blade kolumny z prawej, klik w puste pole za wierszem = nowa kolumna i pisanie
  await p.evaluate(() => { document.querySelectorAll(".toast").forEach((t) => t.remove()); tableWrapEl.scrollLeft = tableWrapEl.scrollWidth; });
  await sleep(300);
  const c6 = await p.evaluate(() => {
    const n = currentDisplayModel.headers.length;
    return {
      cols: document.querySelectorAll("#dataTable col.ghost-col").length,
      letters: [...document.querySelectorAll("#dataTable thead tr.guide-row .ghost-col-cell")].map((x) => x.textContent),
      want: [XLSX.utils.encode_col(currentStartCol + n), XLSX.utils.encode_col(currentStartCol + n + 1)],
      plus: document.querySelector("#dataTable thead tr.header-row .ghost-col-cell")?.textContent,
      dataIdx: document.querySelectorAll("#dataTable .ghost-col-cell[data-col-index]").length,
    };
  });
  check("6. 2 duchy-kolumny z kolejnymi literami i „＋”, bez data-col-index", c6.cols === 2 && c6.letters.join() === c6.want.join() && c6.plus === "＋" && c6.dataIdx === 0, c6);
  // wiersz WIDOCZNY w oknie tabeli (wcześniej przewinięto na dół — 7. wiersz DOM bywa poza widokiem)
  const pt = await p.evaluate(() => {
    const wb = tableWrapEl.getBoundingClientRect();
    const head = document.querySelector("#dataTable thead tr.header-row").getBoundingClientRect();
    const top = Math.max(wb.top, head.bottom) + 4; // nagłówek przyklejony albo nie
    // naprawdę widoczny: w środku jego komórki widać TEN wiersz (nie przyklejony nagłówek nad nim)
    const tr = [...tbodyEl.querySelectorAll("tr[data-row-key]")].find((x) => {
      const r = x.getBoundingClientRect();
      if (!(r.top > top && r.bottom < wb.bottom - 8)) return false;
      const c = x.cells[x.cells.length - 1].getBoundingClientRect();
      const el = document.elementFromPoint(c.left + 3, r.top + r.height / 2);
      return !!el && el.closest("tr") === x;
    });
    const last = tr.cells[tr.cells.length - 1].getBoundingClientRect();
    return { x: last.right + 40, y: last.top + last.height / 2, key: tr.dataset.rowKey };
  });
  await p.mouse.click(pt.x, pt.y);
  await p.waitForSelector(".app-dialog-box input", { state: "visible", timeout: 5000 });
  await p.fill(".app-dialog-box input", "Uwagi"); await p.keyboard.press("Enter");
  await p.waitForFunction(() => currentHeaders.includes("Uwagi"), null, { timeout: 20000 });
  await sleep(500);
  const s6 = await p.evaluate(() => ({ focus: focusedCellState && focusedCellState.rowKey, col: focusedCellState && focusedCellState.colIndex0, real: swbRealHeaderCount(), editor: !!document.querySelector("input.cell-editor") }));
  check("6. klik za wierszem = nowa kolumna i pisanie w tym wierszu", s6.focus === pt.key && s6.col === s6.real - 1 && s6.editor, { s6, pt });
  await p.keyboard.type("uwaga"); await p.keyboard.press("Enter"); await sleep(300);
  const s6b = await p.evaluate((key) => { const r = currentDisplayModel.rows.find((x) => getRowSelectionKey(x) === key); return r && r.values[swbRealHeaderCount() - 1]; }, pt.key);
  check("6. wpis trafił do nowej kolumny w klikniętym wierszu", s6b === "uwaga", s6b);
  await p.evaluate(() => sheetEditUi.setMode("view")); await sleep(200);
  check("6. w Podglądzie duchów-kolumn nie ma", await p.evaluate(() => [...document.querySelectorAll("#dataTable .ghost-col-cell")].every((x) => x.offsetParent === null)));

  check("bez błędów strony", errors.length === 0, errors.slice(0, 5));
  await browser.close();
  if (fails.length) { console.log(`❌ ghost-rows (${ENGINE_NAME}): ${fails.length} niezaliczone`); process.exit(1); }
  console.log(`✅ ghost-rows (${ENGINE_NAME}): wszystko przeszło`);
}
run().catch((e) => { console.error(e); process.exit(1); });
