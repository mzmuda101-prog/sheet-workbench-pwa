// mobile-gestures-playwright.js — GESTY EXCELA MOBILNEGO, CZ. 2 (2026-10-31).
//   node scripts/mobile-gestures-playwright.js        ENGINE=webkit node scripts/mobile-gestures-playwright.js
//
// Wg przewodnika Microsoftu („Excel for iPhone touch guide”):
//   1. „pstryknięcie” kółkiem w dół = zaznaczenie do końca danych (jak Ctrl+Shift+↓);
//   2. drugie kółko w lewym górnym rogu zakresu zmienia POCZĄTEK zaznaczenia;
//   3. przenoszenie przeciągnięciem: menu „Przenieś (przeciągnij)” → ramka → upuszczenie w celu:
//      wartość jedzie, źródło puste, jedno Cofnij przywraca; Esc anuluje tryb;
//   4. Ctrl/⌘+↓ skacze do krawędzi danych, z Shiftem rozciąga zaznaczenie (komputer).
// Przykładowy plik: kolumna D („Imię i Nazwisko”) ma dane w wierszach 6–9, pustą komórkę w 10.
// Uruchom z serwerem na APP_URL (domyślnie http://127.0.0.1:4175/).

const playwright = require("playwright");

const APP_URL = process.env.APP_URL || "http://127.0.0.1:4175/";
const ENGINE = process.env.ENGINE || "chromium";
const SLEEP_SCALE = Math.min(1, Math.max(0.1, Number(process.env.SLEEP_SCALE || 1)));
const sleep = (ms) => new Promise((r) => setTimeout(r, Math.round(ms * SLEEP_SCALE)));
const MOD = process.platform === "darwin" && ENGINE === "webkit" ? "Meta" : "Control";

async function open(browser, opts) {
  const ctx = await browser.newContext({ serviceWorkers: "block", ...opts });
  await ctx.addInitScript(() => {
    if (sessionStorage.getItem("__mgInit")) return;
    sessionStorage.setItem("__mgInit", "1");
    localStorage.clear();
    localStorage.setItem("introPlayed", "true");
    localStorage.setItem("swb-touch-range-hint-v1", "1");
  });
  const page = await ctx.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(() => { document.getElementById("heroSplash")?.remove(); return loadSampleFile(); });
  await page.waitForFunction(() => document.querySelector("#dataTable tbody tr[data-row-key]"), null, { timeout: 20000 });
  await page.waitForFunction(() => !document.documentElement.classList.contains("vt-scene"), null, { timeout: 10000 });
  // Bez blokady kolumn: przykład ma w Excelu zablokowane 2 kolumny (domyślnie przejmowane), a wtedy
  // na telefonie kolumna D stoi przy prawej krawędzi i ciągnięcie uchwytu przewija tabelę w bok.
  await page.evaluate(() => { sheetEditUi.setMode("edit"); if (typeof setSheetFrozenCols === "function") setSheetFrozenCols(0); document.querySelectorAll(".toast").forEach((t) => t.remove()); });
  await sleep(400);
  return { ctx, page, errors };
}

// wskaźnik (dotyk/mysz) od punktu a do b, n kroków co ms milisekund. ms = 0: cały ruch jednym
// ciągiem, bez oddawania wątku — „pstryknięcie” to w aplikacji < 250 ms od dotknięcia do puszczenia
// (autofill.js), a przerwy setTimeout pod obciążeniem (5 testów naraz) potrafiły rozciągnąć gest ponad
// ten próg i test widział wolne przeciąganie (D6:F9 zamiast D6:D9, 2026-11-10).
const swipe = (p, a, b, type, n, ms) => p.evaluate(async ([from, to, t, steps, gap]) => {
  const target = document.elementFromPoint(from.x, from.y);
  const mk = (k, pt) => new PointerEvent(k, { bubbles: true, cancelable: true, pointerId: 21, pointerType: t, isPrimary: true, clientX: pt.x, clientY: pt.y, button: 0, buttons: k === "pointerup" ? 0 : 1 });
  target.dispatchEvent(mk("pointerdown", from));
  for (let i = 1; i <= steps; i++) {
    window.dispatchEvent(mk("pointermove", { x: from.x + (to.x - from.x) * i / steps, y: from.y + (to.y - from.y) * i / steps }));
    if (gap > 0) await new Promise((r) => setTimeout(r, gap));
  }
  window.dispatchEvent(mk("pointerup", to));
}, [a, b, type, n, ms]);

async function run() {
  const browser = await playwright[ENGINE].launch({ headless: true });
  const fails = [];
  const check = (name, ok, detail) => {
    console.log(`${ok ? "✅" : "❌"} ${name}${ok ? "" : ` — ${JSON.stringify(detail)}`}`);
    if (!ok) fails.push(name);
  };
  const KEY = (r) => `wide:${r}`; // wiersz arkusza (0-based)
  const TD = (r, c) => `#dataTable tr[data-row-key="wide:${r}"] > td[data-col-index="${c}"]`;

  // ── telefon ──
  const phone = await open(browser, { viewport: { width: 390, height: 844 }, hasTouch: true, isMobile: ENGINE === "chromium" });
  let p = phone.page;
  const center = (sel) => p.evaluate((s) => { const el = document.querySelector(s); if (!el) return null; const b = el.getBoundingClientRect(); return { x: b.left + b.width / 2, y: b.top + b.height / 2 }; }, sel);
  const rect = () => p.evaluate(() => { const r = getSelectionRectangle(); return r ? { r1: r.model.rows[r.rowStart].rowIndex0, r2: r.model.rows[r.rowEnd].rowIndex0, c1: r.colMin, c2: r.colMax } : null; });
  const edgeRow = await p.evaluate(() => {
    // ostatni wiersz ciągłych danych w kolumnie D od pierwszego wiersza danych
    let i = 0;
    while (i + 1 < viewRows.length && viewRows[i + 1].values[3]) i++;
    return viewRows[i].rowIndex0;
  });
  check("(warunek) kolumna D ma przerwę w danych", edgeRow >= 6 && edgeRow < 20, edgeRow);

  // 1. pstryknięcie kółkiem w dół
  await p.evaluate(() => { setSelectionKind("cell"); setFocusedCell("wide:5", 3, { scroll: true }); });
  await sleep(300);
  let k = await center("#tableWrap .fill-handle:not(.sel-start), td .fill-handle:not(.sel-start)");
  await swipe(p, k, { x: k.x, y: k.y + 70 }, "touch", 3, 0);
  await sleep(300);
  const r1 = await rect();
  check("1. pstryknięcie w dół zaznacza do końca danych (D6:D" + (edgeRow + 1) + ")", r1 && r1.r1 === 5 && r1.r2 === edgeRow && r1.c1 === 3 && r1.c2 === 3, r1);

  // wolne przeciągnięcie to dalej zwykłe rozciąganie (bez skoku)
  await p.evaluate(() => { setSelectedCell("", -1); setFocusedCell("wide:5", 3, { scroll: true }); });
  await sleep(300);
  k = await center("#tableWrap .fill-handle:not(.sel-start), td .fill-handle:not(.sel-start)");
  const to = await center(TD(6, 3));
  await swipe(p, k, to, "touch", 8, 60);
  await sleep(300);
  const r1b = await rect();
  check("1. wolne przeciągnięcie = zwykłe rozciąganie (D6:D7)", r1b && r1b.r1 === 5 && r1b.r2 === 6, r1b);

  // 2. drugie kółko (lewy górny róg) zmienia początek
  await p.evaluate(() => { setSelectionKind("cell"); setFocusedCell("wide:6", 2, { scroll: true }); setSelectedCell("wide:8", 3, { scroll: false }); tableWrapEl.scrollLeft = 0; });
  await sleep(300);
  const s = await center("#tableWrap .fill-handle.sel-start, td .fill-handle.sel-start");
  check("2. zakres ma drugie kółko w lewym górnym rogu", !!s, s);
  if (s) {
    const tgt = await center(TD(5, 2));
    await swipe(p, s, tgt, "touch", 8, 60);
    await sleep(300);
    const r2 = await rect();
    check("2. przeciągnięcie górnego kółka w górę: C6:D9 (koniec zostaje)", r2 && r2.r1 === 5 && r2.c1 === 2 && r2.r2 === 8 && r2.c2 === 3, r2);
  }

  // 3. przenoszenie przeciągnięciem (palec)
  const v = (r, c) => p.evaluate(([rr, cc]) => workbook.Sheets[currentSheetName][XLSX.utils.encode_cell({ r: rr, c: currentStartCol + cc })]?.v ?? null, [r, c]);
  const name6 = await v(5, 3);
  const emptyRow = edgeRow + 1; // pusta komórka D
  await p.evaluate(() => { setSelectedCell("", -1); setSelectionKind("cell"); setFocusedCell("wide:5", 3, { scroll: true }); });
  await p.evaluate((td) => { const el = document.querySelector(td); const b = el.getBoundingClientRect(); openCellMenu(el, { x: b.left + 5, y: b.bottom + 5 }); }, TD(5, 3));
  await p.waitForFunction(() => !!document.querySelector("#cellMenu .cmi-move"), null, { timeout: 3000 }).catch(() => {});
  const hasMove = await p.evaluate(() => !!document.querySelector("#cellMenu .cmi-move"));
  check("3. menu ma „Przenieś (przeciągnij)”", hasMove, hasMove);
  await p.evaluate(() => document.querySelector("#cellMenu .cmi-move").click());
  await sleep(300);
  const box = await center(".move-box");
  check("3. ramka przenoszenia nad blokiem", !!box, box);
  const dst = await center(TD(emptyRow, 3));
  await swipe(p, box, dst, "touch", 10, 50);
  await sleep(500);
  const m3 = { from: await v(5, 3), to: await v(emptyRow, 3), active: await p.evaluate(() => swbMoveCells.active) };
  check("3. przeciągnięcie przenosi wartość, źródło puste, tryb się kończy", m3.to === name6 && m3.from == null && !m3.active, { ...m3, name6 });
  await p.evaluate(() => undoRedo("undo"));
  await sleep(400);
  const m3b = { from: await v(5, 3), to: await v(emptyRow, 3) };
  check("3. jedno Cofnij przywraca", m3b.from === name6 && m3b.to == null, m3b);
  await p.evaluate(() => { setSelectedCell("", -1); setFocusedCell("wide:5", 3, { scroll: true }); swbMoveCells.start(); });
  await sleep(200);
  await p.evaluate(() => document.dispatchEvent(new KeyboardEvent("keydown", { key: "Escape", bubbles: true })));
  await sleep(200);
  check("3. Esc kończy tryb przenoszenia", await p.evaluate(() => !swbMoveCells.active && !document.querySelector(".move-box")), null);
  check("telefon: bez błędów strony", phone.errors.length === 0, phone.errors);
  await phone.ctx.close();

  // ── komputer ──
  const desk = await open(browser, { viewport: { width: 1280, height: 900 } });
  p = desk.page;
  await p.click(TD(5, 3));
  await p.keyboard.press(`${MOD}+ArrowDown`);
  await sleep(200);
  const f4 = await p.evaluate(() => focusedCellState && focusedCellState.rowKey);
  check("4. Ctrl/⌘+↓ skacze do krawędzi danych", f4 === KEY(edgeRow), { f4, edgeRow });
  await p.click(TD(5, 3));
  await p.keyboard.press(`${MOD}+Shift+ArrowDown`);
  await sleep(200);
  const r4 = await p.evaluate(() => { const r = getSelectionRectangle(); return r && [r.model.rows[r.rowStart].rowIndex0, r.model.rows[r.rowEnd].rowIndex0]; });
  check("4. Ctrl/⌘+Shift+↓ rozciąga zaznaczenie do krawędzi", r4 && r4[0] === 5 && r4[1] === edgeRow, r4);
  // przenoszenie myszą
  const name6d = await p.evaluate(() => workbook.Sheets[currentSheetName][XLSX.utils.encode_cell({ r: 5, c: currentStartCol + 3 })]?.v);
  await p.click(TD(5, 3));
  await p.evaluate(() => swbMoveCells.start());
  await sleep(200);
  const bx = await p.evaluate(() => { const b = document.querySelector(".move-box").getBoundingClientRect(); return { x: b.left + b.width / 2, y: b.top + b.height / 2 }; });
  const dd = await p.evaluate((s) => { const b = document.querySelector(s).getBoundingClientRect(); return { x: b.left + b.width / 2, y: b.top + b.height / 2 }; }, TD(edgeRow + 1, 3));
  await p.mouse.move(bx.x, bx.y);
  await p.mouse.down();
  await p.mouse.move(dd.x, dd.y, { steps: 8 });
  await p.mouse.up();
  await sleep(400);
  const moved = await p.evaluate((r) => workbook.Sheets[currentSheetName][XLSX.utils.encode_cell({ r, c: currentStartCol + 3 })]?.v, edgeRow + 1);
  check("4. przenoszenie przeciągnięciem działa też myszą", moved === name6d, { moved, name6d });
  check("komputer: bez błędów strony", desk.errors.length === 0, desk.errors);
  await desk.ctx.close();

  await browser.close();
  if (fails.length) {
    console.error(`\n${fails.length} błędów (${ENGINE})`);
    process.exit(1);
  }
  console.log(`\nOK (${ENGINE})`);
}

run().catch((e) => { console.error(e); process.exit(1); });
