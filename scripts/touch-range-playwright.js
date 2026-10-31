// touch-range-playwright.js — ZAZNACZANIE I WYPEŁNIANIE PALCEM + drobiazgi z filmów z iPhone'a (2026-10-31).
//   node scripts/touch-range-playwright.js        ENGINE=webkit node scripts/touch-range-playwright.js
//
// Zgłoszenia: (1) palec złapał róg zaznaczenia przy przesuwaniu tabeli i seria dat nadpisała
// formuły w kolejnych kolumnach; (2) na telefonie nie dało się rozszerzyć zaznaczenia;
// (3) wpisana data „29 Oct 26” obok „17 wrz 26”; (4) toasty przezroczyste i piętrzące się;
// (5) stuknięcie w komórkę w trybie Edycja nie rysowało ramki; (6) komórka nad nagłówkiem:
// ramka na C2, a w polu nazwy szare „A1”.
// Pilnuje:
//   1. stuknięcie palcem w Edycji = poziom komórki (ramka, pole nazwy);
//   2. przeciągnięcie rogu palcem ROZCIĄGA zaznaczenie i nie zmienia ani jednej wartości;
//   3. przytrzymanie zakresu → „Wypełnij serią / w dół”; seria liczy od zaznaczenia sprzed rozciągnięcia;
//   4. mysz: uchwyt dalej wypełnia (jak w Excelu);
//   5. format daty: polskie nazwy miesięcy i dni, [$-409] zostaje po angielsku;
//   6. toasty tego samego rodzaju się scalają, na dotyku mają kryjące tło (i FAB-y też);
//   7. lista „bez rozmycia” w CSS celuje w klasy, które istnieją (dawniej martwe .clear-selection-fab);
//   8. komórka nad nagłówkiem: pole nazwy i pasek formuły pokazują ją, zapis z paska działa.
// Uruchom z serwerem na APP_URL (domyślnie http://127.0.0.1:4175/).

const playwright = require("playwright");
const fs = require("fs");
const path = require("path");

const APP_URL = process.env.APP_URL || "http://127.0.0.1:4175/";
const ENGINE = process.env.ENGINE || "chromium";
const SLEEP_SCALE = Math.min(1, Math.max(0.1, Number(process.env.SLEEP_SCALE || 1)));
const sleep = (ms) => new Promise((r) => setTimeout(r, Math.round(ms * SLEEP_SCALE)));

async function open(browser, opts) {
  const ctx = await browser.newContext({ serviceWorkers: "block", ...opts });
  await ctx.addInitScript(() => {
    if (sessionStorage.getItem("__trInit")) return;
    sessionStorage.setItem("__trInit", "1");
    localStorage.clear();
    localStorage.setItem("introPlayed", "true");
  });
  const page = await ctx.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(() => { document.getElementById("heroSplash")?.remove(); return loadSampleFile(); });
  await page.waitForFunction(() => document.querySelector("#dataTable tbody tr[data-row-key]") && document.querySelector("#dataTable thead tr.pre-row td.pre-cell"), null, { timeout: 20000 });
  await page.evaluate(() => sheetEditUi.setMode("edit"));
  await page.evaluate(() => document.querySelectorAll(".toast").forEach((t) => t.remove()));
  await sleep(500);
  return { ctx, page, errors };
}

const CELL = (r, c) => `#dataTable tbody tr[data-row-key]:nth-child(${r}) td[data-col-index="${c}"]`;
const center = (p, sel) => p.evaluate((s) => {
  const el = document.querySelector(s);
  if (!el) return null;
  const b = el.getBoundingClientRect();
  return { x: b.left + b.width / 2, y: b.top + b.height / 2 };
}, sel);
// Stuknięcie w komórkę, aż zostanie aktywna: tuż po wczytaniu układ jeszcze się przesuwa
// (zwijany nagłówek), a w tłoku równoległych testów pierwszy tap trafiał obok.
async function tapCell(p, sel) {
  for (let i = 0; i < 4; i++) {
    const pt = await center(p, sel);
    await p.touchscreen.tap(pt.x, pt.y);
    await sleep(300);
    const ok = await p.evaluate((s) => {
      const td = document.querySelector(s);
      return !!(td && focusedCellState && td.dataset.colIndex === String(focusedCellState.colIndex0) && td.parentElement.dataset.rowKey === focusedCellState.rowKey);
    }, sel);
    if (ok) return true;
    await sleep(400);
  }
  return false;
}
const colValues = (p, col, n) => p.evaluate(([c, k]) => [...document.querySelectorAll("#dataTable tbody tr[data-row-key]")].slice(0, k)
  .map((tr) => tr.querySelector(`td[data-col-index="${c}"]`)?.textContent), [col, n]);

// przeciągnięcie wskaźnikiem (touch/mouse) od uchwytu do środka komórki
async function dragHandle(p, toSel, pointerType) {
  const a = await center(p, "#tableWrap .fill-handle, td .fill-handle");
  const b = await center(p, toSel);
  await p.evaluate(async ([from, to, type]) => {
    const target = document.elementFromPoint(from.x, from.y);
    const mk = (t, pt) => new PointerEvent(t, { bubbles: true, cancelable: true, pointerId: 11, pointerType: type, isPrimary: true, clientX: pt.x, clientY: pt.y, button: 0, buttons: t === "pointerup" ? 0 : 1 });
    target.dispatchEvent(mk("pointerdown", from));
    for (let i = 1; i <= 8; i++) {
      window.dispatchEvent(mk("pointermove", { x: from.x + (to.x - from.x) * i / 8, y: from.y + (to.y - from.y) * i / 8 }));
      await new Promise((r) => setTimeout(r, 25));
    }
    window.dispatchEvent(mk("pointerup", to));
  }, [a, b, pointerType]);
  await sleep(400);
  return a;
}

async function run() {
  const browser = await playwright[ENGINE].launch({ headless: true });
  const fails = [];
  const check = (name, ok, detail) => {
    console.log(`${ok ? "✅" : "❌"} ${name}${ok ? "" : ` — ${JSON.stringify(detail)}`}`);
    if (!ok) fails.push(name);
  };

  // ── 7. CSS: lista „bez rozmycia” celuje w istniejące klasy ──
  const css = fs.readFileSync(path.join(__dirname, "..", "styles", "app.css"), "utf8");
  const sources = [fs.readFileSync(path.join(__dirname, "..", "index.html"), "utf8"),
    ...fs.readdirSync(path.join(__dirname, "..", "app")).filter((f) => f.endsWith(".js")).map((f) => fs.readFileSync(path.join(__dirname, "..", "app", f), "utf8"))].join("\n");
  const noGlass = [...new Set([...css.matchAll(/html:not\(\.force-glass\) \.([a-z0-9-]+)/g)].map((m) => m[1]))];
  const dead = noGlass.filter((c) => !new RegExp(`["' .]${c}["' .:]`).test(sources));
  check("7. lista „bez rozmycia” nie ma martwych klas", noGlass.length > 5 && dead.length === 0, { dead, noGlass });

  // ── telefon ──
  const phone = await open(browser, { viewport: { width: 390, height: 844 }, hasTouch: true, isMobile: ENGINE === "chromium" });
  let p = phone.page;
  let pt;
  await tapCell(p, CELL(1, 2));
  const s1 = await p.evaluate(() => ({ kind: selectionKind, active: document.querySelectorAll("#dataTable tbody td.cell-active").length, name: document.getElementById("nameBox").value }));
  check("1. stuknięcie palcem w Edycji = komórka z ramką i nazwą", s1.kind === "cell" && s1.active === 1 && s1.name === "C6", s1);

  // (kolumny B–C — na telefonie widać tylko A…C)
  await tapCell(p, CELL(1, 1));
  const before = await colValues(p, 1, 6);
  const beforeD = await colValues(p, 2, 6);
  await dragHandle(p, CELL(4, 2), "touch");
  const s2 = await p.evaluate(() => {
    const r = getSelectionRectangle();
    return { rows: r && r.rowCount, cols: r && r.colCount, filled: [...document.querySelectorAll(".toast")].some((t) => /Wypełniono/.test(t.textContent)) };
  });
  const after = await colValues(p, 1, 6);
  const afterD = await colValues(p, 2, 6);
  check("2. palec na rogu rozciąga zaznaczenie (4×2)", s2.rows === 4 && s2.cols === 2, s2);
  check("2. …i niczego nie wypełnia", !s2.filled && JSON.stringify(before) === JSON.stringify(after) && JSON.stringify(beforeD) === JSON.stringify(afterD), { before, after, beforeD, afterD });

  // zaznaczenie w górę / w lewo od kotwicy też działa
  // stuknięcie W zakres otwiera menu (jak w Excelu) — najpierw zwijamy zakres
  await p.evaluate(() => { closeCellMenu({ restoreFocus: false }); setSelectedCell("", -1); setFocusedCell("", -1); });
  await tapCell(p, CELL(4, 2));
  await dragHandle(p, CELL(2, 1), "touch");
  const s2b = await p.evaluate(() => { const r = getSelectionRectangle(); return r && { rowStart: r.rowStart, rows: r.rowCount, colMin: r.colMin, cols: r.colCount }; });
  check("2. rozciąganie palcem w górę i w lewo", s2b && s2b.rows === 3 && s2b.cols === 2 && s2b.colMin === 1, s2b);

  // ── 3. menu przytrzymania: seria z zaznaczenia sprzed rozciągnięcia ──
  // kolumna „Nr.” (A): 1, 2 zaznaczone → rozciągnięte do 5 wierszy → 1, 2, 3, 4, 5 bez zmian;
  // potem jedna komórka „1” rozciągnięta w dół → seria kopiuje (jak uchwyt w Excelu).
  await tapCell(p, CELL(1, 2));
  await dragHandle(p, CELL(3, 2), "touch");
  const menu = await p.evaluate(() => {
    const td = document.querySelector('#dataTable tbody tr[data-row-key]:nth-child(2) td[data-col-index="2"]');
    const b = td.getBoundingClientRect();
    openCellMenu(td, { x: b.left + 5, y: b.top + 5 });
    return [...document.querySelectorAll("#cellMenu .cell-menu-item")].map((x) => x.className.match(/cmi-[\w-]+/)?.[0]);
  });
  check("3. menu zakresu ma „Wypełnij serią” i „w dół”, bez „w prawo” (1 kolumna)", menu.includes("cmi-fill-series") && menu.includes("cmi-fill-down") && !menu.includes("cmi-fill-right"), menu);
  const src = (await colValues(p, 2, 3))[0];
  await p.evaluate(() => document.querySelector("#cellMenu .cmi-fill-series").click());
  await sleep(500);
  const filled = await colValues(p, 2, 3);
  const keep = await p.evaluate(() => { const r = getSelectionRectangle(); return r && r.rowCount; });
  check("3. „Wypełnij serią” wypełnia zakres od pierwszej komórki", filled.every((v) => v === src), { src, filled });
  check("3. zaznaczenie zostaje po wypełnieniu", keep === 3, keep);
  await p.evaluate(() => { if (typeof swbUndo === "function") swbUndo(); });

  // ── 9. stuknięcia jak w Excelu na telefonie ──
  await p.evaluate(() => { document.querySelectorAll(".toast").forEach((x) => x.remove()); closeCellMenu({ restoreFocus: false }); setSelectedCell("", -1); setFocusedCell("", -1); });
  await tapCell(p, CELL(3, 1));
  await sleep(600); // wolne drugie stuknięcie (nie dwutap)
  pt = await center(p, CELL(3, 1));
  await p.touchscreen.tap(pt.x, pt.y);
  await sleep(300);
  const m9 = await p.evaluate(() => ({ open: !document.getElementById("cellMenu").classList.contains("hidden"), items: [...document.querySelectorAll("#cellMenu .cell-menu-item")].map((b) => b.className.match(/cmi-[\w-]+/)?.[0]), editor: !!document.querySelector("#dataTable input.cell-editor") }));
  check("9. drugie (wolne) stuknięcie w zaznaczoną komórkę = menu z „Edytuj” na górze", m9.open && m9.items[0] === "cmi-edit" && !m9.editor, m9);
  await p.evaluate(() => closeCellMenu({ restoreFocus: false }));
  await sleep(500);
  await p.touchscreen.tap(pt.x, pt.y);
  await sleep(120);
  await p.touchscreen.tap(pt.x, pt.y);
  await sleep(300);
  check("9. podwójne stuknięcie = edycja", await p.evaluate(() => !!document.querySelector("#dataTable input.cell-editor")), null);
  await p.evaluate(() => { if (activeCellEditor) activeCellEditor.close(); });
  // stuknięcie w zakres = menu zakresu, zaznaczenie zostaje
  await p.evaluate(() => { setSelectionKind("cell"); setFocusedCell(getRowSelectionKey(viewRows[1]), 1); setSelectedCell(getRowSelectionKey(viewRows[3]), 2); });
  await sleep(600);
  pt = await center(p, CELL(3, 2));
  await p.touchscreen.tap(pt.x, pt.y);
  await sleep(300);
  const m9b = await p.evaluate(() => { const r = getSelectionRectangle(); return { open: !document.getElementById("cellMenu").classList.contains("hidden"), n: r && r.rowCount * r.colCount, head: document.querySelector("#cellMenu .cell-menu-val")?.textContent }; });
  check("9. stuknięcie w zaznaczony zakres = menu zakresu, zakres zostaje", m9b.open && m9b.n === 6, m9b);
  await p.evaluate(() => closeCellMenu({ restoreFocus: false }));
  // numer wiersza = cały wiersz, litera kolumny = cała kolumna
  await sleep(500);
  pt = await center(p, '#dataTable tbody tr[data-row-key]:nth-child(2) td.row-head');
  await p.touchscreen.tap(pt.x, pt.y);
  await sleep(300);
  const r9 = await p.evaluate(() => { const r = getSelectionRectangle(); return r && { rows: r.rowCount, cols: r.colCount, all: currentDisplayModel.headers.length }; });
  check("9. stuknięcie numeru wiersza = cały wiersz", r9 && r9.rows === 1 && r9.cols === r9.all, r9);
  pt = await center(p, '#dataTable thead tr.guide-row th.guide-cell:nth-child(3)');
  await p.touchscreen.tap(pt.x, pt.y);
  await sleep(300);
  const c9 = await p.evaluate(() => { const r = getSelectionRectangle(); return r && { rows: r.rowCount, cols: r.colCount, colMin: r.colMin, all: swbSelView().rows.length }; });
  check("9. stuknięcie litery kolumny = cała kolumna (B)", c9 && c9.cols === 1 && c9.colMin === 1 && c9.rows === c9.all, c9);
  // Podgląd: palec dalej zaznacza komórkę i ma kółko do rozciągania
  await p.evaluate(() => { sheetEditUi.setMode("view"); setSelectedCell("", -1); setFocusedCell("", -1); });
  await sleep(500);
  await tapCell(p, CELL(2, 1));
  const v9 = await p.evaluate(() => ({ kind: selectionKind, knob: !!document.querySelector("#tableWrap .fill-handle, td .fill-handle") }));
  check("9. Podgląd: stuknięcie = komórka + kółko zaznaczania", v9.kind === "cell" && v9.knob, v9);
  const before9 = await colValues(p, 1, 4);
  await dragHandle(p, CELL(4, 2), "touch");
  const rr9 = await p.evaluate(() => { const r = getSelectionRectangle(); return r && r.rowCount * r.colCount; });
  check("9. Podgląd: rozciąganie działa i nic nie zmienia", rr9 === 6 && JSON.stringify(before9) === JSON.stringify(await colValues(p, 1, 4)), rr9);
  await p.evaluate(() => sheetEditUi.setMode("edit"));

  // ── 6. toasty: scalanie + kryjące tło; FAB kryjący ──
  await p.evaluate(() => { document.querySelectorAll(".toast").forEach((t) => t.remove()); toast("A (1 kom.)", "info"); toast("B (1 kom.)", "info"); toast("C (2 kom.)", "info"); toast("Uwaga", "warning"); });
  await sleep(300);
  const t6 = await p.evaluate(() => [...document.querySelectorAll(".toast")].map((t) => ({ text: t.textContent, bg: getComputedStyle(t).backgroundColor })));
  const opaque = (c) => /^rgb\(/.test(c) || /rgba\([^)]*,\s*1\)$/.test(c);
  check("6. trzy „info” pod rząd = jeden toast, ostrzeżenie osobno", t6.length === 2 && /C \(2 kom\.\)/.test(t6[0].text), t6);
  check("6. toast na dotyku ma kryjące tło", t6.every((x) => opaque(x.bg)), t6);
  const fab = await p.evaluate(() => { const f = document.getElementById("clearSelectionFab"); const cs = getComputedStyle(f); return { vis: f.classList.contains("is-visible"), bgImage: cs.backgroundImage, opacity: cs.opacity, bg: cs.backgroundColor }; });
  check("6. „Odznacz” na dotyku nie prześwituje", fab.vis && fab.opacity === "1" && opaque(fab.bg), fab);

  // ── 8. komórka nad nagłówkiem: pole nazwy + pasek formuły ──
  pt = await center(p, '#dataTable thead tr.pre-row[data-row-index="1"] td.pre-cell[data-col-index="1"]');
  await p.touchscreen.tap(pt.x, pt.y);
  await sleep(300);
  const s8 = await p.evaluate(() => ({ name: document.getElementById("nameBox").value, fx: document.getElementById("formulaInput").value, dis: document.getElementById("formulaInput").disabled, f: workbook.Sheets[currentSheetName].B2?.f }));
  // B2 w przykładzie to formuła — pasek ma pokazać formułę po polsku, nie wynik
  check("8. pole nazwy = B2, pasek formuły = jej formuła", s8.name === "B2" && !s8.dis && !!s8.f && /^=ILE\.NIEPUSTYCH\(/.test(s8.fx), s8);
  await p.evaluate(() => {
    const fx = document.getElementById("formulaInput");
    fx.focus();
    fx.value = "121";
    fx.dispatchEvent(new Event("input", { bubbles: true }));
    fx.dispatchEvent(new KeyboardEvent("keydown", { key: "Enter", bubbles: true }));
  });
  await sleep(500);
  const s8b = await p.evaluate(() => ({ v: workbook.Sheets[currentSheetName].B2?.v, shown: document.querySelector('#dataTable thead tr.pre-row[data-row-index="1"] td.pre-cell[data-col-index="1"]')?.textContent.trim(), name: document.getElementById("nameBox").value }));
  // Enter w pasku formuły schodzi niżej — jak w wierszach danych (i w Excelu)
  check("8. zapis z paska formuły trafia do komórki nad nagłówkiem", s8b.v === 121 && s8b.shown === "121" && s8b.name === "B3", s8b);
  await p.evaluate(() => { setSelectedCell("", -1); setFocusedCell("", -1); });
  const s8c = await p.evaluate(() => document.getElementById("nameBox").value);
  check("8. po odznaczeniu pole nazwy jest puste", s8c === "", s8c);
  check("telefon: bez błędów strony", phone.errors.length === 0, phone.errors);
  await phone.ctx.close();

  // ── desktop ──
  const desk = await open(browser, { viewport: { width: 1280, height: 900 } });
  p = desk.page;
  // 5. polskie nazwy w formatach dat (wyrocznia: polski Excel, 2026-10-31)
  const d5 = await p.evaluate(() => [
    cfmtFormatValue(new Date(2026, 9, 29), "dd mmm yy"),
    cfmtFormatValue(new Date(2026, 9, 29), "d mmmm yyyy"),
    cfmtFormatValue(new Date(2026, 9, 29), "dddd"),
    cfmtFormatValue(new Date(2026, 9, 29), "ddd dd.mm"),
    cfmtFormatValue(new Date(2026, 9, 29), "[$-409]dd mmm yy"),
    cfmtFormatValue(new Date(2026, 4, 1), "mmm yyyy"),
  ]);
  check("5. daty jak polski Excel", JSON.stringify(d5) === JSON.stringify(["29 paź 26", "29 październik 2026", "czwartek", "czw 29.10", "29 Oct 26", "maj 2026"]), d5);

  // 1/4. mysz w Edycji: klik = komórka (od 2026-10-31, jak palec), uchwyt dalej wypełnia
  await p.click(CELL(1, 2));
  await sleep(200);
  const k = await p.evaluate(() => selectionKind);
  check("1. mysz w Edycji: klik zaznacza komórkę", k === "cell", k);
  const d0 = await colValues(p, 2, 3);
  await dragHandle(p, CELL(3, 2), "mouse");
  const d1 = await colValues(p, 2, 3);
  check("4. mysz: uchwyt dalej wypełnia", d1[1] === d0[0] && d1[2] === d0[0] && d0[1] !== d0[0], { d0, d1 });
  check("desktop: bez błędów strony", desk.errors.length === 0, desk.errors);
  await desk.ctx.close();

  await browser.close();
  if (fails.length) {
    console.error(`\n${fails.length} błędów (${ENGINE})`);
    process.exit(1);
  }
  console.log(`\nOK (${ENGINE})`);
}

run().catch((e) => { console.error(e); process.exit(1); });
