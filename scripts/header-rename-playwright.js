// header-rename-playwright.js — ZMIANA NAZWY KOLUMNY (app/header-edit.js), 2026-10-31.
//   node scripts/header-rename-playwright.js        ENGINE=webkit node scripts/header-rename-playwright.js
//
// Nazwa kolumny to klucz dla sortowania, filtrów, szukania („Status:="x"”) i kolumn
// wyliczanych („[Status]”). Pilnuje:
//   1. menu nagłówka (prawy klik; na telefonie przytrzymanie) ma „Zmień nazwę kolumny” i sortowanie;
//      przytrzymanie NIE sortuje kolumny;
//   2. zmiana nazwy: komórka arkusza, nagłówek, lista sortowania, sort i zapytanie idą za nową nazwą,
//      kolumna wyliczana [Status] liczy się dalej;
//   3. Cofnij / Ponów przywracają nazwę RAZEM ze stanem (sort, zapytanie);
//   4. pusta, zdublowana i ze znakami składni („:”) nazwa — odmowa, nic się nie zmienia;
//   5. zapis do pliku i ponowne wczytanie: nowa nazwa jest w pliku;
//   6. kolumna tabeli Excela (plik RODO, arkusz 2025-2026) — zablokowana z wyjaśnieniem.
// Uruchom z serwerem na APP_URL (domyślnie http://127.0.0.1:4175/).

const playwright = require("playwright");
const path = require("path");

const APP_URL = process.env.APP_URL || "http://127.0.0.1:4175/";
const ENGINE = process.env.ENGINE || "chromium";
const SLEEP_SCALE = Math.min(1, Math.max(0.1, Number(process.env.SLEEP_SCALE || 1)));
const sleep = (ms) => new Promise((r) => setTimeout(r, Math.round(ms * SLEEP_SCALE)));
const RODO = path.join(__dirname, "RODO Obieg terenów 2026 V5.xlsx");
const STATUS = '#dataTable thead tr.header-row th[data-col-index="2"]';

async function open(browser, opts) {
  const ctx = await browser.newContext({ serviceWorkers: "block", ...opts });
  await ctx.addInitScript(() => {
    if (sessionStorage.getItem("__hrInit")) return;
    sessionStorage.setItem("__hrInit", "1");
    localStorage.clear();
    localStorage.setItem("introPlayed", "true");
  });
  const page = await ctx.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(() => { document.getElementById("heroSplash")?.remove(); return loadSampleFile(); });
  await page.waitForFunction(() => document.querySelector("#dataTable thead tr.header-row th[data-col-index]"), null, { timeout: 20000 });
  await page.evaluate(() => sheetEditUi.setMode("edit"));
  await sleep(500);
  return { ctx, page, errors };
}

const state = (p) => p.evaluate(() => ({
  h: currentHeaders[2],
  cell: workbook.Sheets[currentSheetName][XLSX.utils.encode_cell({ r: currentHeaderRow - 1, c: currentStartCol + 2 })]?.v,
  th: document.querySelector('#dataTable thead tr.header-row th[data-col-index="2"]')?.dataset.sortHeader,
  sort: sortState.col,
  q: searchQueryEl.value,
  opt: [...document.querySelectorAll("#sortColumnSelect option")].map((o) => o.value),
}));

async function run() {
  const browser = await playwright[ENGINE].launch({ headless: true });
  const fails = [];
  const check = (name, ok, detail) => {
    console.log(`${ok ? "✅" : "❌"} ${name}${ok ? "" : ` — ${JSON.stringify(detail)}`}`);
    if (!ok) fails.push(name);
  };

  // ── desktop ──
  const desk = await open(browser, { viewport: { width: 1280, height: 900 } });
  let p = desk.page;
  await p.evaluate(() => {
    setPrimarySort("Status", "asc");
    sortRows();
    searchQueryEl.value = 'Status:="W trakcie" && Zakres:="Zakres A-01"';
    derivedColumns = [{ id: dcNextId(), name: "Stan krótko", expr: 'LEFT([Status];3)' }];
    applyDerivedColumns({ silent: true });
    renderActiveTable();
  });
  await sleep(300);
  await p.click(STATUS, { button: "right" });
  await sleep(300);
  const menu = await p.evaluate(() => [...document.querySelectorAll("#cellMenu .cell-menu-item")].map((b) => b.className.match(/cmi-[\w-]+/)?.[0]));
  check("1. menu nagłówka: zmiana nazwy + sortowanie", ["cmi-hdr-rename", "cmi-hdr-sort-asc", "cmi-hdr-sort-desc"].every((x) => menu.includes(x)), menu);
  await p.click("#cellMenu .cmi-hdr-rename");
  await sleep(200);
  check("1. pole zmiany nazwy w nagłówku z zaznaczoną nazwą", await p.evaluate(() => {
    const i = document.querySelector("th input.hdr-editor");
    return !!i && document.activeElement === i && i.selectionStart === 0 && i.selectionEnd === i.value.length;
  }), null);
  await p.keyboard.type("Stan zakresu");
  await p.keyboard.press("Enter");
  await sleep(500);
  const s2 = await state(p);
  check("2. nazwa w arkuszu, nagłówku i liście sortowania", s2.h === "Stan zakresu" && s2.cell === "Stan zakresu" && s2.th === "Stan zakresu" && s2.opt.includes("Stan zakresu") && !s2.opt.includes("Status"), s2);
  check("2. sort i zapytanie idą za nową nazwą (wartości w cudzysłowie bez zmian)", s2.sort === "Stan zakresu" && s2.q === 'Stan zakresu:="W trakcie" && Zakres:="Zakres A-01"', s2);
  const dc = await p.evaluate(() => ({ expr: derivedColumns[0].expr, v: viewRows[0] && viewRows[0].display[currentHeaders.indexOf("Stan krótko")] }));
  check("2. kolumna wyliczana [Status] → [Stan zakresu] i liczy się dalej", dc.expr === "LEFT([Stan zakresu];3)" && typeof dc.v === "string" && dc.v.length === 3, dc);

  await p.keyboard.press(ENGINE === "webkit" && process.platform === "darwin" ? "Meta+z" : "Control+z");
  await sleep(500);
  const s3 = await state(p);
  check("3. Cofnij przywraca nazwę razem z sortem i zapytaniem", s3.h === "Status" && s3.cell === "Status" && s3.sort === "Status" && s3.q.startsWith('Status:="W trakcie"'), s3);
  await p.evaluate(() => undoRedo("redo"));
  await sleep(500);
  const s3b = await state(p);
  check("3. Ponów wraca do nowej nazwy", s3b.h === "Stan zakresu" && s3b.sort === "Stan zakresu", s3b);

  const refused = await p.evaluate(() => [swbRenameColumn(2, ""), swbRenameColumn(2, "Zakres"), swbRenameColumn(2, "Stan: nowy"), swbRenameColumn(2, " zakres ")]);
  const s4 = await state(p);
  check("4. pusta / zdublowana / z „:” — odmowa bez zmian", refused.every((x) => x === false) && s4.h === "Stan zakresu", { refused, s4 });

  const saved = await p.evaluate(async () => {
    const out = await buildPatchedXlsx(originalFileBytes, pendingEdits, buildSheetFormulaMaps());
    const wb = XLSX.read(out, { type: "array" });
    const sh = wb.Sheets[currentSheetName];
    return sh[XLSX.utils.encode_cell({ r: currentHeaderRow - 1, c: currentStartCol + 2 })]?.v;
  });
  check("5. nowa nazwa trafia do zapisanego pliku", saved === "Stan zakresu", saved);
  check("desktop: bez błędów strony", desk.errors.length === 0, desk.errors);
  await desk.ctx.close();

  // ── telefon: przytrzymanie nagłówka = menu, bez sortowania ──
  const phone = await open(browser, { viewport: { width: 390, height: 844 }, hasTouch: true, isMobile: ENGINE === "chromium" });
  p = phone.page;
  const before = await p.evaluate(() => sortState.col);
  // pozycja nagłówka musi się ustalić (zwijany nagłówek strony, dorysowanie tabeli) —
  // inaczej przy skróconych przerwach (SLEEP_SCALE) „palec” trafiał w pustkę
  // w trakcie animacji przejścia sceny (html.vt-scene) dotyk trafia w cały dokument
  await p.waitForFunction(() => !document.documentElement.classList.contains("vt-scene"), null, { timeout: 10000 });
  let pt = null;
  for (let i = 0; i < 40; i++) {
    const cur = await p.evaluate((s) => { const b = document.querySelector(s).getBoundingClientRect(); return { x: Math.round(b.left + b.width / 2), y: Math.round(b.top + b.height / 2) }; }, STATUS);
    if (pt && cur.x === pt.x && cur.y === pt.y) break;
    pt = cur;
    await new Promise((r) => setTimeout(r, 150));
  }
  const cdp = ENGINE === "chromium" ? await p.context().newCDPSession(p) : null;
  if (cdp) {
    // trzymamy „palec”, aż menu wyjdzie (w tłoku testów timer przytrzymania się spóźnia)
    await cdp.send("Input.dispatchTouchEvent", { type: "touchStart", touchPoints: [{ x: pt.x, y: pt.y }] });
    await p.waitForFunction(() => !document.getElementById("cellMenu").classList.contains("hidden"), null, { timeout: 5000 }).catch(() => {});
    await sleep(100);
    await cdp.send("Input.dispatchTouchEvent", { type: "touchEnd", touchPoints: [] });
  } else {
    // WebKit w Playwright nie ma długiego dotyku — sprawdzamy tę samą ścieżkę menu
    await p.evaluate((s) => { const th = document.querySelector(s); const b = th.getBoundingClientRect(); swbOpenHeaderMenu(th, { x: b.left + 5, y: b.top + 5 }); }, STATUS);
  }
  await sleep(400);
  const lp = await p.evaluate(() => ({ menu: !document.getElementById("cellMenu").classList.contains("hidden"), sort: sortState.col }));
  check("1. telefon: przytrzymanie nagłówka otwiera menu i nie sortuje", lp.menu && lp.sort === before, { lp, before });
  check("telefon: bez błędów strony", phone.errors.length === 0, phone.errors);
  await phone.ctx.close();

  // ── 6. tabela Excela ──
  const ctx = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1280, height: 900 } });
  await ctx.addInitScript(() => localStorage.setItem("introPlayed", "true"));
  p = await ctx.newPage();
  const errs6 = [];
  p.on("pageerror", (e) => errs6.push(e.message));
  await p.goto(APP_URL, { waitUntil: "load" });
  await p.evaluate(() => document.getElementById("heroSplash")?.remove());
  await p.setInputFiles("#fileInput", RODO);
  await p.waitForFunction(() => document.querySelector("#dataTable thead tr.header-row th[data-col-index]"), null, { timeout: 30000 });
  await sleep(1000);
  await p.evaluate(() => { sheetSelect.value = "2025-2026"; document.getElementById("loadBtn").click(); });
  await p.waitForFunction(() => currentSheetName === "2025-2026" && document.querySelector("#dataTable thead tr.header-row th[data-col-index]"), null, { timeout: 30000 });
  await sleep(800);
  await p.evaluate(() => sheetEditUi.setMode("edit"));
  const t6 = await p.evaluate(() => {
    const before = currentHeaders[0];
    const reason = swbHeaderRenameBlocker(0);
    const ok = swbRenameColumn(0, "Numer");
    return { before, after: currentHeaders[0], reason, ok };
  });
  check("6. kolumna tabeli Excela zablokowana z wyjaśnieniem", !t6.ok && t6.after === t6.before && /opracowanie_terenow/.test(t6.reason || ""), t6);
  await p.click('#dataTable thead tr.header-row th[data-col-index="0"]', { button: "right" });
  await sleep(300);
  const m6 = await p.evaluate(() => ({ dis: document.querySelector("#cellMenu .cmi-hdr-rename")?.getAttribute("aria-disabled"), note: document.querySelector("#cellMenu .cell-menu-note")?.textContent || "" }));
  check("6. menu: pozycja wyszarzona + powód pod nią", m6.dis === "true" && /tabeli Excela/.test(m6.note), m6);
  check("tabela: bez błędów strony", errs6.length === 0, errs6);
  await ctx.close();

  await browser.close();
  if (fails.length) {
    console.error(`\n${fails.length} błędów (${ENGINE})`);
    process.exit(1);
  }
  console.log(`\nOK (${ENGINE})`);
}

run().catch((e) => { console.error(e); process.exit(1); });
