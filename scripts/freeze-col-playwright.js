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
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

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
  await page.click("#loadBtn");
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
    return new Promise((r) => requestAnimationFrame(() => setTimeout(r, 120)));
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
  if (errors.length) failures.push(`błędy konsoli/strony: ${errors.join(" | ")}`);

  console.log(JSON.stringify({ off, onAtZero, onScrolled, backAtZero, errors }, null, 2));
  if (failures.length) throw new Error(failures.join("; "));
  console.log("✅ freeze-col-playwright OK");
}

run().catch((error) => {
  console.error(error);
  process.exit(1);
});
