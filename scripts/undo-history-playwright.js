// undo-history-playwright.js — Cofnij/Ponów (undo.js) + historia zapytań (qs-history.js).
//
// Cofnij — co musi być prawdą:
//   1. edycja komórki → Ctrl+Z przywraca DOKŁADNIE stan sprzed: wartość i typ w arkuszu,
//      brak wpisu w pendingEdits (zapis nie dotknie tej komórki), tekst w siatce,
//   2. Ctrl+Shift+Z ponawia, Ctrl+Y też,
//   3. „Wypełnij w dół" na 5 wierszach = JEDEN krok cofania,
//   4. nowa zmiana po cofnięciu czyści „Ponów",
//   5. Ctrl+Z w polu tekstowym NIE cofa arkusza (tam działa cofanie pisania),
//   6. przycisk ↶ w pasku pojawia się po zmianie i działa tak samo jak skrót.
// Historia — co musi być prawdą:
//   7. puste pole szukania pokazuje ostatnie zapytania (najnowsze na górze, z ikonką ≠),
//   8. ☆ przypina (sekcja „Przypięte", trwałe w localStorage), × usuwa,
//   9. klik we wpis stosuje zapytanie RAZEM z ≠ i operatorami (ten sam wynik co ręcznie),
//  10. ↓ z pustego pola wchodzi na listę historii.
//
// Uruchom z serwerem na APP_URL (domyślnie http://127.0.0.1:4175/).

const { chromium, webkit } = require("playwright");
const path = require("path");

const APP_URL = process.env.APP_URL || "http://127.0.0.1:4175/";
const FILE = path.join(__dirname, "stress-test-workbench.xlsx");
const ENGINE = process.env.ENGINE === "webkit" ? webkit : chromium;
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

async function run() {
  const browser = await ENGINE.launch({ headless: true });
  const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1280, height: 900 } });
  await context.addInitScript(() => localStorage.setItem("introPlayed", "true"));
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push("pageerror: " + e.message));
  page.on("console", (m) => { if (m.type() === "error") errors.push("console: " + m.text()); });

  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(() => { document.getElementById("heroSplash")?.remove(); localStorage.removeItem("excel-workbench-qs-history"); });
  await page.evaluate(() => { try { ensureXlsxLibs && ensureXlsxLibs(false); } catch {} });
  await page.setInputFiles("#fileInput", FILE);
  await page.waitForFunction(() => document.getElementById("sheetSelect")?.options?.length > 0, null, { timeout: 15000 });
  await sleep(300);
  await page.click("#loadBtn");
  await sleep(1200);
  await page.evaluate(() => { if (typeof setSidebarOpen === "function") setSidebarOpen(false); });
  await sleep(300);

  const failures = [];
  const report = {};
  const ok = (name, cond, got) => { if (!cond) failures.push(`${name} — dostałem: ${JSON.stringify(got)}`); };

  // Stan komórki: arkusz + pendingEdits + wiersz + DOM.
  const cellState = (rowIndex0, col) => page.evaluate(({ rowIndex0, col }) => {
    const ref = XLSX.utils.encode_cell({ r: rowIndex0, c: currentStartCol + col });
    const c = workbook.Sheets[currentSheetName][ref];
    const row = baseRows.find((r) => r.rowIndex0 === rowIndex0);
    const td = document.querySelector(`tbody tr[data-row-key="wide:${rowIndex0}"] td[data-col-index="${col}"]`);
    return {
      ref,
      v: c ? (c.v instanceof Date ? c.v.toISOString() : c.v) : null,
      t: c ? c.t : null,
      pending: !!(pendingEdits[currentSheetName] && Object.prototype.hasOwnProperty.call(pendingEdits[currentSheetName], ref)),
      rowDisplay: row ? String(getDisplayValue(row, col)) : null,
      td: td ? td.textContent : null,
    };
  }, { rowIndex0, col });

  // Edycja przez prawdziwy edytor komórki (Enter zatwierdza).
  const editCell = async (rowIndex0, col, text) => {
    await page.evaluate(({ rowIndex0, col }) => {
      const td = document.querySelector(`tbody tr[data-row-key="wide:${rowIndex0}"] td[data-col-index="${col}"]`);
      setFocusedCell(`wide:${rowIndex0}`, col, { scroll: true });
      openCellEditor(td);
    }, { rowIndex0, col });
    await page.waitForSelector("input.cell-editor");
    await page.fill("input.cell-editor", text);
    await page.keyboard.press("Enter");
    await sleep(150);
    await page.evaluate(() => { document.activeElement?.blur?.(); giveGridDomFocus && giveGridDomFocus(); });
  };

  const target = await page.evaluate(() => {
    const row = currentDisplayModel.rows.find((r) => String(getDisplayValue(r, 2) ?? "").trim() !== "");
    return { rowIndex0: row.rowIndex0, col: 2 };
  });
  const before = await cellState(target.rowIndex0, target.col);
  report.before = before;

  // 1. edycja → Ctrl+Z
  await editCell(target.rowIndex0, target.col, "ZZZ-test");
  const edited = await cellState(target.rowIndex0, target.col);
  ok("edycja zadziałała (warunek testu)", edited.v === "ZZZ-test" && edited.pending, edited);
  const btnAfterEdit = await page.evaluate(() => !document.getElementById("undoBtn").classList.contains("hidden") && !document.getElementById("undoBtn").disabled);
  ok("6. przycisk ↶ pojawia się po zmianie", btnAfterEdit, btnAfterEdit);

  await page.keyboard.press("Control+z");
  await sleep(250);
  const undone = await cellState(target.rowIndex0, target.col);
  report.undone = undone;
  ok("1. Ctrl+Z przywraca wartość i typ", undone.v === before.v && undone.t === before.t, { undone, before });
  ok("1. Ctrl+Z zdejmuje komórkę z pendingEdits", undone.pending === before.pending, undone);
  ok("1. siatka i wiersz pokazują starą wartość", undone.rowDisplay === before.rowDisplay && undone.td === before.td, { undone, before });

  // 2. ponów
  await page.keyboard.press("Control+Shift+z");
  await sleep(250);
  const redone = await cellState(target.rowIndex0, target.col);
  ok("2. Ctrl+Shift+Z ponawia", redone.v === "ZZZ-test" && redone.pending && redone.td === "ZZZ-test", redone);
  await page.keyboard.press("Control+z");
  await sleep(200);
  await page.keyboard.press("Control+y");
  await sleep(200);
  const redoneY = await cellState(target.rowIndex0, target.col);
  ok("2. Ctrl+Y też ponawia", redoneY.v === "ZZZ-test", redoneY);
  // przycisk ↶ = to samo co skrót
  await page.click("#undoBtn");
  await sleep(250);
  const viaBtn = await cellState(target.rowIndex0, target.col);
  ok("6. przycisk ↶ cofa", viaBtn.v === before.v, viaBtn);

  // 3. Wypełnij w dół na 5 wierszach = jeden krok
  const fill = await page.evaluate(() => {
    const rows = currentDisplayModel.rows.slice(0, 5);
    const col = 3;
    const keys = rows.map((r) => getRowSelectionKey(r));
    const beforeVals = rows.map((r) => String(getDisplayValue(r, col) ?? ""));
    setFocusedCell(keys[0], col, { scroll: false });
    setSelectedCell(keys[4], col, { scroll: false });
    fillSelection("down");
    return { col, idx: rows.map((r) => r.rowIndex0), beforeVals };
  });
  await sleep(250);
  const filled = await page.evaluate(({ idx, col }) => idx.map((i) => String(getDisplayValue(baseRows.find((r) => r.rowIndex0 === i), col) ?? "")), fill);
  ok("wypełnienie zadziałało (warunek testu)", filled.every((v) => v === fill.beforeVals[0]), { filled, fill });
  await page.evaluate(() => giveGridDomFocus && giveGridDomFocus());
  await page.keyboard.press("Control+z");
  await sleep(250);
  const unfilled = await page.evaluate(({ idx, col }) => idx.map((i) => String(getDisplayValue(baseRows.find((r) => r.rowIndex0 === i), col) ?? "")), fill);
  ok("3. „Wypełnij w dół” cofa się JEDNYM Ctrl+Z", JSON.stringify(unfilled) === JSON.stringify(fill.beforeVals), { unfilled, before: fill.beforeVals });

  // 4. nowa zmiana czyści Ponów
  await editCell(target.rowIndex0, target.col, "nowe");
  const redoState = await page.evaluate(() => ({ hidden: document.getElementById("redoBtn").classList.contains("hidden"), stack: redoStack.length }));
  ok("4. nowa zmiana czyści Ponów", redoState.hidden && redoState.stack === 0, redoState);

  // 5. Ctrl+Z w polu tekstowym nie rusza arkusza
  const beforeInput = await cellState(target.rowIndex0, target.col);
  await page.evaluate(() => { const i = document.getElementById("quickSearch"); document.getElementById("quickSearchWrap").classList.remove("hidden"); i.focus(); });
  await page.keyboard.press("Control+z");
  await sleep(200);
  const afterInput = await cellState(target.rowIndex0, target.col);
  ok("5. Ctrl+Z w polu tekstowym nie cofa arkusza", afterInput.v === beforeInput.v && afterInput.v === "nowe", afterInput);
  await page.evaluate(() => { document.activeElement.blur(); document.getElementById("quickSearchWrap").classList.add("hidden"); });

  // 1b. Edytuj → Cofnij → ZAPIS: w pliku wynikowym komórka ma ORYGINAŁ,
  //     a zmiana, której nie cofnięto („nowe"), jest zapisana.
  const other = await page.evaluate(({ skip }) => {
    const row = currentDisplayModel.rows.find((r) => r.rowIndex0 !== skip && String(getDisplayValue(r, 1) ?? "").trim() !== "");
    return { rowIndex0: row.rowIndex0, col: 1 };
  }, { skip: target.rowIndex0 });
  const otherBefore = await cellState(other.rowIndex0, other.col);
  await editCell(other.rowIndex0, other.col, "DO-COFNIECIA");
  await page.evaluate(() => giveGridDomFocus && giveGridDomFocus());
  await page.keyboard.press("Control+z");
  await sleep(250);
  const saved = await page.evaluate(async ({ a, b }) => {
    const bytes = await buildOutputBytes("xlsx");
    const wb = XLSX.read(bytes, { cellDates: true });
    const sh = wb.Sheets[currentSheetName];
    const val = (ref) => { const c = sh[ref]; return c ? (c.v instanceof Date ? c.v.toISOString() : c.v) : null; };
    return { undoneCell: val(a), keptCell: val(b) };
  }, { a: otherBefore.ref, b: target && (await cellState(target.rowIndex0, target.col)).ref });
  report.saved = { ...saved, original: otherBefore.v };
  ok("1b. po Cofnij zapis ma w komórce oryginał, a niecofnięta zmiana jest w pliku",
    saved.undoneCell === otherBefore.v && saved.keptCell === "nowe", report.saved);

  // ── Historia zapytań ──
  const commit = async (q, { neg = false } = {}) => {
    await page.evaluate(({ q, neg }) => {
      closeQuickSearchPopup(); openQuickSearchPopup();
      filterNegateEl.checked = neg; filterModeEl.value = "contains"; syncQuickSearchModeControls();
      document.getElementById("quickSearchPopupOperators").checked = true;
      const input = document.getElementById("quickSearchPopupInput");
      input.value = q; input.focus();
    }, { q, neg });
    await page.keyboard.press("Enter");
    await sleep(350);
    return page.evaluate(() => viewRows.length);
  };
  const nAlpha = await commit("Alpha");
  const nNegBeta = await commit("Beta", { neg: true });
  report.hist = { nAlpha, nNegBeta };

  const openEmpty = async () => {
    await page.evaluate(() => {
      closeQuickSearchPopup(); openQuickSearchPopup();
      const input = document.getElementById("quickSearchPopupInput");
      input.value = ""; input.dispatchEvent(new Event("input", { bubbles: true })); input.blur(); input.focus();
    });
    await sleep(300);
    return page.evaluate(() => ({
      visible: !document.getElementById("qsLiveResultsPopup").classList.contains("hidden"),
      groups: [...document.querySelectorAll("#qsLiveResultsPopup .qs-live-group")].map((g) => g.textContent),
      items: [...document.querySelectorAll("#qsLiveResultsPopup .qs-hist-item")].map((b) => b.textContent),
    }));
  };
  const h1 = await openEmpty();
  report.h1 = h1;
  ok("7. puste pole pokazuje ostatnie zapytania, najnowsze na górze z ≠", h1.visible && h1.items[0] === "≠Beta" && h1.items[1] === "Alpha", h1);

  // 10. ↓ wchodzi na listę
  await page.keyboard.press("ArrowDown");
  const onHist = await page.evaluate(() => document.activeElement?.classList?.contains("qs-hist-item"));
  ok("10. ↓ z pustego pola wchodzi na historię", onHist, onHist);

  // 8. przypnij „Alpha"
  await page.evaluate(() => {
    const rows = [...document.querySelectorAll("#qsLiveResultsPopup .qs-hist-row")];
    rows.find((r) => r.querySelector(".qs-hist-text").textContent === "Alpha").querySelector(".qs-hist-pin").click();
  });
  await sleep(200);
  const h2 = await page.evaluate(() => ({
    groups: [...document.querySelectorAll("#qsLiveResultsPopup .qs-live-group")].map((g) => g.textContent),
    first: document.querySelector("#qsLiveResultsPopup .qs-hist-item")?.textContent,
    stored: JSON.parse(localStorage.getItem("excel-workbench-qs-history") || "{}"),
    stillOpen: !document.getElementById("qsLiveResultsPopup").classList.contains("hidden"),
  }));
  ok("8. ☆ przypina (sekcja Przypięte, zapis w localStorage)", h2.stillOpen && h2.groups.length === 2 && h2.first === "Alpha" && h2.stored.pinned?.[0]?.q === "Alpha", h2);

  // 9. klik we wpis „≠ Beta" → te same wiersze co ręcznie + ≠ przywrócone
  await page.evaluate(() => { filterNegateEl.checked = false; syncQuickSearchModeControls(); });
  await page.evaluate(() => {
    const item = [...document.querySelectorAll("#qsLiveResultsPopup .qs-hist-item")].find((b) => b.textContent === "≠Beta");
    item.click();
  });
  await sleep(400);
  const applied = await page.evaluate(() => ({ n: viewRows.length, q: searchQueryEl.value, neg: filterNegateEl.checked }));
  ok("9. klik we wpis stosuje zapytanie z ≠", applied.n === nNegBeta && applied.q === "Beta" && applied.neg, { applied, nNegBeta });

  // 8b. × usuwa
  await openEmpty();
  await page.evaluate(() => {
    const rows = [...document.querySelectorAll("#qsLiveResultsPopup .qs-hist-row")];
    rows.find((r) => r.querySelector(".qs-hist-text").textContent === "Beta").querySelector(".qs-hist-del").click();
  });
  await sleep(200);
  const h3 = await page.evaluate(() => [...document.querySelectorAll("#qsLiveResultsPopup .qs-hist-item")].map((b) => b.textContent));
  ok("8. × usuwa wpis", !h3.includes("≠Beta") && h3.includes("Alpha"), h3);

  ok("brak błędów w konsoli", errors.length === 0, errors);
  await browser.close();

  console.log(JSON.stringify(report, null, 2));
  if (failures.length) {
    console.log("\n❌ " + failures.length + " błędów:\n - " + failures.join("\n - "));
    process.exit(1);
  }
  console.log("\n✅ undo-history: wszystko OK");
}

run().catch((e) => { console.error(e); process.exit(1); });
