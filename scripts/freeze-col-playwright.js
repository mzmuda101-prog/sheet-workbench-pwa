// freeze-col-playwright.js — BLOKADA PIERWSZEJ KOLUMNY: poprawność + zabezpieczenie wydajności.
//
// `position: sticky` dostają dwie komórki w KAŻDYM wierszu, więc przy 200 wierszach
// silnik repozycjonuje 400 elementów na każdej klatce przewijania. Dopóki tabela nie
// jest odjechana w bok, te komórki stoją tam, gdzie stałyby normalnie — trzymamy je
// wtedy jako zwykłe (klasa `.freeze-col-active` wchodzi dopiero przy scrollLeft > 0).
//
// Co musi być prawdą:
//   1. blokada wyłączona → brak sticky i brak klasy (stan wyjściowy),
//   2. blokada włączona, scrollLeft = 0 → komórki NIE są sticky (to jest ta oszczędność),
//   3. po przewinięciu w bok → sticky wchodzi, a zamrożona kolumna FAKTYCZNIE stoi
//      w miejscu (jej pozycja na ekranie się nie zmienia) — czyli funkcja dalej działa,
//   4. powrót do lewej krawędzi → sticky znowu schodzi,
//   5. przełączenie klasy nie przesuwa kolumny (brak skoku przy starcie przewijania).
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
    return {
      active: wrap.classList.contains("freeze-col-active"),
      frozen: wrap.classList.contains("freeze-first-col"),
      position: cell ? getComputedStyle(cell).position : "",
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
  if (off.active) failures.push("klasa freeze-col-active nie powinna istnieć bez blokady");
  if (!onAtZero.frozen) failures.push("po włączeniu brakuje klasy freeze-first-col");
  if (onAtZero.active) failures.push("przy scrollLeft=0 sticky ma być ZDJĘTY (to jest oszczędność)");
  if (onAtZero.position !== "static") failures.push(`przy scrollLeft=0 pozycja komórki = ${onAtZero.position}, oczekiwano static`);
  if (!onScrolled.active) failures.push("po przewinięciu w bok brakuje klasy freeze-col-active");
  if (onScrolled.position !== "sticky") failures.push(`po przewinięciu pozycja komórki = ${onScrolled.position}, oczekiwano sticky`);
  if (onScrolled.scrollLeft < 100) failures.push(`tabela nie przewinęła się w bok (scrollLeft=${onScrolled.scrollLeft}) — test nic nie sprawdza`);
  if (Math.abs(onScrolled.left - onAtZero.left) > 2) {
    failures.push(`zamrożona kolumna przesunęła się przy przewijaniu: ${onAtZero.left}px → ${onScrolled.left}px`);
  }
  if (backAtZero.active) failures.push("po powrocie do lewej krawędzi sticky powinien znów zejść");
  if (errors.length) failures.push(`błędy konsoli/strony: ${errors.join(" | ")}`);

  console.log(JSON.stringify({ off, onAtZero, onScrolled, backAtZero, errors }, null, 2));
  if (failures.length) throw new Error(failures.join("; "));
  console.log("✅ freeze-col-playwright OK");
}

run().catch((error) => {
  console.error(error);
  process.exit(1);
});
