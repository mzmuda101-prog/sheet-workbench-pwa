// freeze-col-playwright.js — BLOKADA PIERWSZEJ KOLUMNY: poprawność + zabezpieczenie wydajności.
//
// Historia: na iPhonie przewijanie z włączoną blokadą potrafiło zjeść 3× więcej czasu
// na klatkę. Winowajcą był `backdrop-filter: blur(4px)` na komórkach zamrożonej kolumny
// — dostają go DWIE komórki w KAŻDYM wierszu, czyli setki rozmywanych warstw liczonych
// co klatkę. Blur był wprawdzie wyłączany regułą `html.browser-safari`, ale iOS w trybie
// PWA z ekranu głównego nie wysyła w user-agencie tokenu „Safari", więc akurat tam
// obejście nie działało. Dlatego blur zniknął dla wszystkich.
//
// Co musi być prawdą:
//   1. blokada wyłączona → komórki nie są sticky,
//   2. po włączeniu → sticky wchodzi i zamrożona kolumna FAKTYCZNIE stoi w miejscu
//      przy przewijaniu w bok (pozycja na ekranie bez zmian),
//   3. zamrożone komórki NIE mają backdrop-filter (to jest ta regresja, której pilnujemy),
//   4. tło zamrożonych komórek jest nieprzezroczyste (inaczej treść spod spodu prześwituje).
//
// Uruchom z serwerem na APP_URL (domyślnie http://127.0.0.1:4175/).

const { chromium } = require("playwright");
const path = require("path");

const APP_URL = process.env.APP_URL || "http://127.0.0.1:4175/";
const FILE = path.join(__dirname, "stress-test-workbench.xlsx");
// SLEEP_SCALE (run-tests.mjs): skraca przerwy „odczekaj po kroku” w szybkim przebiegu;
// powtórka porażki idzie ze skalą 1 (pełne przerwy). sleepFixed = przerwa, która MUSI
// przeczekać timer aplikacji (debounce, opóźnienie podpowiedzi, bezczynność) — bez skali.
const SLEEP_SCALE = Math.min(1, Math.max(0.1, Number(process.env.SLEEP_SCALE || 1)));
const sleep = (ms) => new Promise((r) => setTimeout(r, Math.round(ms * SLEEP_SCALE)));
const sleepFixed = (ms) => new Promise((r) => setTimeout(r, ms)); // eslint-disable-line no-unused-vars

async function run() {
  const browser = await chromium.launch({ headless: true });
  const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 900, height: 800 } });
  await context.addInitScript(() => localStorage.setItem("introPlayed", "true"));
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push("pageerror: " + e.message));
  page.on("console", (m) => { if (m.type() === "error") errors.push("console: " + m.text()); });

  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  await page.evaluate(() => { try { ensureXlsxLibs && ensureXlsxLibs(false); } catch {} });
  await page.setInputFiles("#fileInput", FILE);
  await page.waitForFunction(() => document.getElementById("sheetSelect")?.options?.length > 0, null, { timeout: 20000 });
  await sleep(300);
  await page.evaluate(() => loadBtn.click());
  await sleep(1200);

  const probe = () => page.evaluate(() => {
    const wrap = document.getElementById("tableWrap");
    const cell = document.querySelector("#dataTable tbody td:nth-child(2)");
    const st = cell ? getComputedStyle(cell) : null;
    return {
      frozen: wrap.classList.contains("freeze-first-col"),
      position: st ? st.position : "",
      backdrop: st ? (st.backdropFilter || st.webkitBackdropFilter || "none") : "",
      background: st ? st.backgroundColor : "",
      left: cell ? Math.round(cell.getBoundingClientRect().left) : -1,
      scrollLeft: Math.round(wrap.scrollLeft),
    };
  });
  const scrollTo = (x) => page.evaluate((x) => {
    const wrap = document.getElementById("tableWrap");
    wrap.scrollLeft = x;
    wrap.dispatchEvent(new Event("scroll"));
    // 350 ms = po spoczynku: w trakcie przewijania prawdziwe komórki oddają miejsce
    // „szybie" (app/freeze-pane.js, test: freeze-pane-playwright.js), a sticky wraca po ~180 ms.
    return new Promise((r) => requestAnimationFrame(() => setTimeout(r, 350)));
  }, x);

  const off = await probe();
  await page.evaluate(() => {
    const el = document.getElementById("freezeFirstCol");
    el.checked = true;
    el.dispatchEvent(new Event("change", { bubbles: true }));
  });
  await sleep(300);
  const onAtZero = await probe();
  await scrollTo(400);
  const onScrolled = await probe();
  await scrollTo(0);
  const backAtZero = await probe();

  // ── N kolumn (2026-11-01): 3 kolumny danych stoją w miejscu i przylegają do siebie,
  // szyba (freeze-pane.js) kopiuje numer wiersza + 3 kolumny, ustawienie pamiętane per arkusz.
  const lefts = () => page.evaluate(() => {
    const wl = document.getElementById("tableWrap").getBoundingClientRect().left;
    return [...document.querySelector("#dataTable tbody tr[data-row-key]").cells].slice(0, 5).map((c) => Math.round(c.getBoundingClientRect().left - wl));
  });
  await page.evaluate(() => { setSheetFrozenCols(3); syncFrozenColOffsets(); });
  const n3AtZero = await lefts();
  await scrollTo(500);
  const n3Scrolled = await lefts();
  const paneCells = await page.evaluate(() => {
    const wrap = document.getElementById("tableWrap");
    wrap.scrollLeft = 520;
    wrap.dispatchEvent(new Event("scroll"));
    const fp = window.__freezePane;
    const on = fp && fp.isActive();
    const row = document.querySelector(".fz-pane tbody tr");
    return { on, cells: row ? row.cells.length : 0, offCount: document.querySelectorAll("#dataTable tbody td.fz-off").length, rows: document.querySelectorAll("#dataTable tbody tr").length };
  });
  await scrollTo(500);
  // przejście klawiaturą do kolumny tuż za zamrożonymi: komórka nie chowa się pod nimi (scroll-padding)
  const reveal = await page.evaluate(async () => {
    const key = document.querySelector("#dataTable tbody tr[data-row-key]").dataset.rowKey;
    document.getElementById("tableWrap").scrollLeft = 900;
    await new Promise((r) => setTimeout(r, 400));
    setFocusedCell(key, 3, { scroll: true });
    const td = findCellElement(focusedCellState);
    const fr = td.parentElement.cells[3].getBoundingClientRect().right;
    return { tdLeft: Math.round(td.getBoundingClientRect().left), frozenRight: Math.round(fr) };
  });
  const n3Mem = await page.evaluate(() => JSON.parse(localStorage.getItem("swb-freeze-by-sheet-v1") || "{}"));
  await page.evaluate(() => setSheetFrozenCols(0));
  await sleep(100);
  const n0 = await probe();

  await browser.close();

  const failures = [];
  if (off.frozen) failures.push("blokada powinna startować wyłączona");
  if (off.position === "sticky") failures.push("bez blokady komórki nie powinny być sticky");
  if (!onAtZero.frozen) failures.push("po włączeniu brakuje klasy freeze-first-col");
  if (onAtZero.position !== "sticky") failures.push(`po włączeniu pozycja komórki = ${onAtZero.position}, oczekiwano sticky`);
  // To jest właściwa regresja, której pilnujemy — rozmycie na setkach komórek zabijało scroll.
  if (onAtZero.backdrop && onAtZero.backdrop !== "none") {
    failures.push(`zamrożona kolumna ma backdrop-filter (${onAtZero.backdrop}) — to kosztuje ~2,5× czasu klatki przy przewijaniu`);
  }
  if (/rgba\(.*0(\.\d+)?\)$/.test(onAtZero.background)) {
    failures.push(`tło zamrożonej kolumny jest przezroczyste (${onAtZero.background}) — treść spod spodu będzie prześwitywać`);
  }
  if (onScrolled.scrollLeft < 100) failures.push(`tabela nie przewinęła się w bok (scrollLeft=${onScrolled.scrollLeft}) — test nic nie sprawdza`);
  if (Math.abs(onScrolled.left - onAtZero.left) > 2) {
    failures.push(`zamrożona kolumna przesunęła się przy przewijaniu: ${onAtZero.left}px → ${onScrolled.left}px`);
  }
  // N kolumn
  if (!n3Scrolled.slice(0, 4).every((x, i) => Math.abs(x - n3AtZero[i]) <= 2)) failures.push(`3 zamrożone kolumny przesunęły się: ${n3AtZero} → ${n3Scrolled}`);
  if (Math.abs(n3Scrolled[4] - n3AtZero[4]) < 100) failures.push(`4. kolumna danych powinna odjechać: ${n3AtZero[4]} → ${n3Scrolled[4]}`);
  if (!(n3AtZero[1] < n3AtZero[2] && n3AtZero[2] < n3AtZero[3])) failures.push(`zamrożone kolumny nie leżą obok siebie: ${n3AtZero}`);
  if (paneCells.on && paneCells.cells !== 4) failures.push(`szyba powinna kopiować 4 komórki w wierszu, ma ${paneCells.cells}`);
  if (paneCells.on && paneCells.offCount !== paneCells.rows * 4) failures.push(`fz-off na ${paneCells.offCount} komórkach, oczekiwano ${paneCells.rows * 4}`);
  if (!Object.values(n3Mem).some((v) => v.c === 3)) failures.push(`liczba kolumn nie zapisana per arkusz: ${JSON.stringify(n3Mem)}`);
  if (reveal.tdLeft < reveal.frozenRight - 1) failures.push(`komórka z fokusem schowana pod zamrożonymi kolumnami: ${JSON.stringify(reveal)}`);
  if (n0.frozen || n0.position === "sticky") failures.push("po „Odblokuj kolumny” komórki nadal przyklejone");
  if (errors.length) failures.push(`błędy konsoli/strony: ${errors.join(" | ")}`);

  console.log(JSON.stringify({ off, onAtZero, onScrolled, backAtZero, n3AtZero, n3Scrolled, paneCells, n0, errors }, null, 2));
  if (failures.length) throw new Error(failures.join("; "));
  console.log("✅ freeze-col-playwright OK");
}

run().catch((error) => {
  console.error(error);
  process.exit(1);
});
