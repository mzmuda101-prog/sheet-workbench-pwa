// dv-suggest-position-playwright.js — POZYCJA dropdownu podpowiedzi Data Validation.
//
// Pilnuje, że lista podpowiedzi (.cell-suggest) trzyma się edytowanej komórki także
// wtedy, gdy przewinie się CAŁA STRONA (a nie tylko tabela). Na telefonie to jest
// codzienny przypadek: po tapnięciu komórki iOS sam dosuwa pole nad klawiaturę
// (scroll dokumentu), a wcześniej lista zostawała w starym miejscu — „uciekała"
// kilkaset pikseli od komórki i trzeba było ruszyć palcem po ekranie, żeby wróciła.
//
// Uruchom z serwerem na APP_URL (domyślnie http://127.0.0.1:4175/).
const playwright = require("playwright");
const ENGINE = process.env.ENGINE || "webkit";
const engine = playwright[ENGINE];
const path = require("path");

const APP_URL = process.env.APP_URL || "http://127.0.0.1:4175/";
const FILE = path.join(__dirname, "stress-test-workbench.xlsx");
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

const results = [];
function check(name, pass, detail) { results.push({ name, pass: !!pass, detail }); }

async function run() {
  const browser = await engine.launch({ headless: true });
  const context = await browser.newContext({ serviceWorkers: "block", ...playwright.devices["iPhone 13"] });
  await context.addInitScript(() => localStorage.setItem("introPlayed", "true"));
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push("pageerror: " + e.message));
  page.on("console", (m) => { if (m.type() === "error") errors.push("console: " + m.text()); });

  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  await page.evaluate(() => { try { ensureXlsxLibs && ensureXlsxLibs(false); } catch {} });
  await page.setInputFiles("#fileInput", FILE);
  await page.waitForFunction(() => document.getElementById("sheetSelect")?.options?.length > 0, null, { timeout: 25000 });
  await sleep(300);
  await page.click("#loadBtn");
  await sleep(1500);

  // Reguła DV (lista) na pierwszej kolumnie tekstowej + otwarcie edytora w komórce.
  const opened = await page.evaluate(() => {
    const colIdx = Math.max(0, currentHeaders.indexOf("Status"));
    addManualDvRule(currentSheetName, ["Alfa", "Beta", "Gamma", "Delta"], colIdx, currentStartCol, "warning", "Status");
    const tds = document.querySelectorAll(`#dataTable tbody tr td[data-col-index="${colIdx}"]`);
    const td = tds[2];
    if (!td) return { ok: false };
    td.scrollIntoView({ block: "center" });
    openCellEditor(td);
    return { ok: !!document.querySelector(".cell-suggest:not(.hidden)"), colIdx };
  });
  check("dropdown DV otwarty nad komórką", opened.ok, JSON.stringify(opened));

  const measure = () => page.evaluate(() => {
    const box = document.querySelector(".cell-suggest");
    const input = document.querySelector("input.cell-editor");
    if (!box || !input) return null;
    const b = box.getBoundingClientRect();
    const i = input.getBoundingClientRect();
    const shown = !box.classList.contains("hidden") && !box.classList.contains("cell-suggest-offscreen");
    // Lista siedzi pod komórką ALBO nad nią (flip) — liczy się odległość od tej krawędzi.
    const gap = Math.min(Math.abs(b.top - i.bottom), Math.abs(i.top - b.bottom));
    return { gap: Math.round(gap), dx: Math.round(b.left - i.left), shown, pageY: Math.round(window.scrollY) };
  });

  await sleep(120);
  const before = await measure();
  check("start: lista przy komórce (odstęp < 40px)", before && before.shown && before.gap < 40, JSON.stringify(before));

  // 1) Scroll CAŁEJ STRONY — to robi iOS, gdy wjeżdża klawiatura ekranowa.
  await page.evaluate(() => window.scrollTo(0, document.documentElement.scrollHeight));
  await sleep(250);
  const afterPage = await measure();
  check("po przewinięciu strony lista nadal przy komórce", afterPage && afterPage.shown && afterPage.gap < 40, JSON.stringify(afterPage));
  check("po przewinięciu strony lista nie odjechała w poziomie", afterPage && Math.abs(afterPage.dx) < 40, JSON.stringify(afterPage));

  // 2) scrollIntoView na polu edycji — dokładnie to robi „trzymaj pole nad klawiaturą".
  await page.evaluate(() => document.querySelector("input.cell-editor").scrollIntoView({ block: "center" }));
  await sleep(250);
  const afterIntoView = await measure();
  // Niezmiennik: lista jest ALBO przyklejona do komórki, ALBO schowana. Nigdy
  // „widoczna, ale kilkaset pikseli od komórki" — to był właśnie zgłoszony objaw.
  check("po dosunięciu pola nad klawiaturę: przy komórce albo schowana",
    afterIntoView && (!afterIntoView.shown || afterIntoView.gap < 40), JSON.stringify(afterIntoView));

  // 3) Scroll wewnętrzny tabeli — to działało wcześniej, pilnujemy regresji.
  await page.evaluate(() => { document.getElementById("tableWrap").scrollTop += 60; });
  await sleep(250);
  const afterInner = await measure();
  check("po przewinięciu tabeli lista nadal przy komórce", afterInner && afterInner.shown && afterInner.gap < 40, JSON.stringify(afterInner));

  // 4) Komórka wyjechała poza okno tabeli → lista znika zamiast parkować się na ekranie.
  await page.evaluate(() => { document.getElementById("tableWrap").scrollTop += 1500; });
  await sleep(250);
  const afterFar = await measure();
  check("komórka poza widokiem → lista schowana", afterFar && !afterFar.shown, JSON.stringify(afterFar));

  // 5) Powrót komórki w widok → lista wraca pod komórkę.
  await page.evaluate(() => { document.getElementById("tableWrap").scrollTop -= 1500; });
  await sleep(250);
  const afterBack = await measure();
  check("powrót komórki w widok → lista wraca pod komórkę", afterBack && afterBack.shown && afterBack.gap < 40, JSON.stringify(afterBack));

  // 6) Komórka tuż nad dolną krawędzią widoku (tak wygląda świat, gdy wjedzie
  //    klawiatura ekranowa): lista ma się ZMIEŚCIĆ nad/pod nią, a nie ją zasłonić.
  await page.keyboard.press("Escape");
  await sleep(150);
  const tight = await page.evaluate(() => {
    const colIdx = Math.max(0, currentHeaders.indexOf("Status"));
    const wrap = document.getElementById("tableWrap");
    const wr = wrap.getBoundingClientRect();
    const tds = [...document.querySelectorAll(`#dataTable tbody tr td[data-col-index="${colIdx}"]`)];
    // ostatnia komórka, która jeszcze w całości mieści się w oknie tabeli
    const td = tds.filter((el) => {
      const r = el.getBoundingClientRect();
      return r.top >= wr.top && r.bottom <= wr.bottom;
    }).pop();
    if (!td) return null;
    openCellEditor(td);
    return true;
  });
  await sleep(250);
  const tightRects = tight ? await page.evaluate(() => {
    const box = document.querySelector(".cell-suggest");
    const input = document.querySelector("input.cell-editor");
    if (!box || !input) return null;
    const b = box.getBoundingClientRect(), i = input.getBoundingClientRect();
    const overlap = Math.min(b.bottom, i.bottom) - Math.max(b.top, i.top);
    return { overlap: Math.round(overlap), boxH: Math.round(b.height), shown: !box.classList.contains("hidden") && !box.classList.contains("cell-suggest-offscreen") };
  }) : null;
  check("lista nie zasłania edytowanej komórki", tightRects && (!tightRects.shown || tightRects.overlap <= 0), JSON.stringify(tightRects));

  await browser.close();

  let bad = 0;
  results.forEach((r) => { if (!r.pass) bad++; console.log(`${r.pass ? "OK  " : "FAIL"} ${r.name}${r.detail ? " — " + r.detail : ""}`); });
  if (errors.length) { console.log("Błędy strony:"); errors.forEach((e) => console.log("  " + e)); }
  console.log(bad ? `\n${bad} test(ów) nie przeszło.` : "\nWszystkie testy przeszły.");
  process.exit(bad || errors.length ? 1 : 0);
}

run().catch((e) => { console.error(e); process.exit(1); });
