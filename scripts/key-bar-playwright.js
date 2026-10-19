// key-bar-playwright.js — pasek cyfr i operatorów nad klawiaturą telefonu (app/key-bar.js).
//
// W przeglądarce bez okna nie ma klawiatury ekranowej, więc test włącza window.__swbKeyBarForce
// (pasek zachowuje się, jakby klawiatura była otwarta). Sprawdza:
//   1. komórka w edycji: pasek z cyframi, stuknięcia wpisują „12,5” w miejscu kursora,
//      fokus zostaje w polu (edytor nie zatwierdza się), Enter zapisuje LICZBĘ 12,5,
//   2. szybkie szukanie: zestaw operatorów; „>>” włącza „Operatory wyszukiwania”,
//   3. po wyjściu z pola pasek znika,
//   4. komputer (mysz): paska nie ma nigdy.

const pw = require("playwright");

const APP_URL = process.env.APP_URL || "http://127.0.0.1:4175/";
const ENGINE = process.env.ENGINE || "chromium";
const SLEEP_SCALE = Math.min(1, Math.max(0.1, Number(process.env.SLEEP_SCALE || 1)));
const sleep = (ms) => new Promise((r) => setTimeout(r, Math.round(ms * SLEEP_SCALE)));

async function openApp(browser, opts) {
  const ctx = await browser.newContext({ serviceWorkers: "block", ...opts });
  await ctx.addInitScript(() => {
    window.__swbKeyBarForce = true;
    localStorage.setItem("introPlayed", "true");
    localStorage.setItem("excel-workbench-toolbar-collapsed-v2", "0");
  });
  const page = await ctx.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(() => { document.getElementById("heroSplash")?.remove(); return loadSampleFile(); });
  await page.waitForFunction(() => document.querySelector("#dataTable tbody tr[data-row-key]"), null, { timeout: 20000 });
  await sleep(600);
  return { ctx, page, errors };
}

async function run() {
  const browser = await pw[ENGINE].launch({ headless: true });
  const failures = [];
  const ok = (name, cond, got) => { if (!cond) failures.push(`${name} — dostałem: ${JSON.stringify(got)}`); };

  // ── telefon ──
  const phone = await openApp(browser, { viewport: { width: 390, height: 844 }, hasTouch: true, isMobile: ENGINE === "chromium" });
  const p = phone.page;
  const key = (k) => p.locator(`.key-bar .key-bar-key[data-key="${k.replace(/"/g, '\\"')}"]`);

  // 1. komórka w edycji
  const cellSel = '#dataTable tbody tr[data-row-key] td[data-col-index="7"]';
  const target = await p.evaluate((sel) => {
    const td = document.querySelectorAll(sel)[2];
    openCellEditor(td);
    const inp = document.querySelector("input.cell-editor");
    inp.value = "";
    return td.parentElement.dataset.rowKey;
  }, cellSel);
  await sleep(200);
  const bar1 = await p.evaluate(() => ({ visible: swbKeyBar.visible, mode: swbKeyBar.mode, first: document.querySelector(".key-bar .key-bar-key")?.dataset.key }));
  ok("1. edycja komórki na telefonie: pasek z cyframi", bar1.visible && bar1.mode === "cell" && bar1.first === "1", bar1);
  for (const k of ["1", "2", ",", "5"]) { await key(k).tap(); await sleep(60); }
  const typed = await p.evaluate(() => ({ value: document.querySelector("input.cell-editor")?.value, focused: document.activeElement?.classList.contains("cell-editor") }));
  ok("1. stuknięcia wpisują „12,5”, fokus zostaje w polu", typed.value === "12,5" && typed.focused, typed);
  await p.keyboard.press("Enter");
  await sleep(300);
  const saved = await p.evaluate((rowKey) => {
    const row = getRowByKey(rowKey);
    const ref = XLSX.utils.encode_cell({ r: row.rowIndex0, c: currentStartCol + 7 });
    return { v: workbook.Sheets[currentSheetName][ref]?.v, t: workbook.Sheets[currentSheetName][ref]?.t };
  }, target);
  ok("1. Enter zapisuje liczbę 12,5", saved.v === 12.5 && saved.t === "n", saved);

  // 2. szybkie szukanie — operatory
  await p.evaluate(() => {
    document.getElementById("quickSearchOperators").checked = false;
    document.getElementById("quickSearchWrap")?.classList.remove("hidden");
    quickSearchEl.value = "";
    quickSearchEl.focus();
  });
  await sleep(200);
  const bar2 = await p.evaluate(() => ({ visible: swbKeyBar.visible, mode: swbKeyBar.mode, first: document.querySelector(".key-bar .key-bar-key")?.dataset.key }));
  ok("2. szybkie szukanie: pasek z operatorami", bar2.visible && bar2.mode === "search" && bar2.first === ">>", bar2);
  await key(">>").tap();
  await sleep(60);
  for (const k of ["1", "0"]) { await key(k).tap(); await sleep(60); }
  const s2 = await p.evaluate(() => ({ value: quickSearchEl.value, ops: document.getElementById("quickSearchOperators").checked, focused: document.activeElement === quickSearchEl }));
  ok("2. „>>10” wpisane, operatory włączone same, fokus w polu", s2.value === ">>10" && s2.ops && s2.focused, s2);

  // 3. wyjście z pola
  await p.evaluate(() => quickSearchEl.blur());
  await sleep(200);
  ok("3. po wyjściu z pola pasek znika", (await p.evaluate(() => swbKeyBar.visible)) === false, null);
  // 5. Widok → „Pasek cyfr”: przełącznik widać na telefonie; wyłączony = brak paska, zapamiętane
  const tog = await p.evaluate(() => { const b = document.getElementById("keyBarToggle"); return { shown: !!b && !b.hidden && getComputedStyle(b).display !== "none", pressed: b?.getAttribute("aria-pressed") }; });
  ok("5. telefon: przełącznik „Pasek cyfr” w Widoku, włączony", tog.shown && tog.pressed === "true", tog);
  await p.evaluate(() => document.getElementById("keyBarToggle").click());
  await p.evaluate(() => quickSearchEl.focus());
  await sleep(200);
  const off = await p.evaluate(() => ({ visible: swbKeyBar.visible, pref: localStorage.getItem("swb-keybar-v1"), pressed: document.getElementById("keyBarToggle").getAttribute("aria-pressed") }));
  ok("5. wyłączony pasek nie pokazuje się i wybór jest zapamiętany", !off.visible && off.pref === "0" && off.pressed === "false", off);
  await p.evaluate(() => { quickSearchEl.blur(); document.getElementById("keyBarToggle").click(); quickSearchEl.focus(); });
  await sleep(200);
  ok("5. ponowne włączenie przywraca pasek", (await p.evaluate(() => swbKeyBar.visible)) === true, null);
  // 6. „×” na pasku wyłącza go od ręki (fokus zostaje w polu), przełącznik w Widoku to widzi
  await p.evaluate(() => document.querySelector(".key-bar .key-bar-off").click());
  await sleep(200);
  const x = await p.evaluate(() => ({ visible: swbKeyBar.visible, pref: localStorage.getItem("swb-keybar-v1"), pressed: document.getElementById("keyBarToggle").getAttribute("aria-pressed"), focus: document.activeElement === quickSearchEl, toast: [...document.querySelectorAll(".toast")].some((el) => /Widok/.test(el.textContent)) }));
  ok("6. „×” chowa pasek, zapamiętuje wybór i podpowiada, gdzie go włączyć", !x.visible && x.pref === "0" && x.pressed === "false" && x.focus && x.toast, x);
  await p.evaluate(() => { quickSearchEl.blur(); document.getElementById("keyBarToggle").click(); quickSearchEl.focus(); });
  await sleep(200);
  ok("6. włączenie w Widoku przywraca pasek po „×”", (await p.evaluate(() => swbKeyBar.visible)) === true, null);
  ok("brak błędów strony (telefon)", phone.errors.length === 0, phone.errors);
  await phone.ctx.close();

  // ── 4. komputer ──
  const desk = await openApp(browser, { viewport: { width: 1280, height: 800 } });
  await desk.page.evaluate(() => { quickSearchEl.focus(); });
  await sleep(200);
  ok("4. komputer: paska nie ma", (await desk.page.evaluate(() => swbKeyBar.visible)) === false, null);
  ok("4. komputer: przełącznika „Pasek cyfr” nie ma w Widoku", (await desk.page.evaluate(() => getComputedStyle(document.getElementById("keyBarToggle")).display === "none")), null);
  ok("brak błędów strony (komputer)", desk.errors.length === 0, desk.errors);
  await desk.ctx.close();

  await browser.close();
  if (failures.length) {
    console.error(`❌ key-bar (${ENGINE}):\n  - ` + failures.join("\n  - "));
    process.exit(1);
  }
  console.log(`✅ key-bar (${ENGINE}): cyfry w komórce, operatory w szukaniu, fokus zostaje, tylko telefon`);
}

run().catch((e) => { console.error(e); process.exit(1); });
