// hero-pull-expand-playwright.js — „pociągnij, żeby rozwinąć" zwinięty nagłówek (mobile).
//
// Decyzja Mateusza 2026-09-24: samo dojechanie do góry tabeli rozwijało nagłówek
// niechcący. Teraz rozwija go tylko świadomy gest.
//
// Co musi być prawdą:
//   1. przewinięcie w dół zwija nagłówek (bez zmian),
//   2. dojechanie do samej góry go NIE rozwija,
//   3. lekkie pociągnięcie w dół na górze (< progu) też nie — ale uchwyt pokazuje postęp,
//   4. wyraźne pociągnięcie w dół na górze rozwija,
//   5. jeden ciągły ruch „z dołu do góry i jeszcze dalej" rozwija, a ruch kończący się
//      dokładnie na górze — nie,
//   6. kółko/gładzik: pojedyncze przewinięcie w górę na górze nie rozwija, dłuższe tak,
//   7. tap w uchwyt nadal przełącza.
// Dotyk: prawdziwe zdarzenia przez CDP Input.dispatchTouchEvent (Chromium, isMobile).

const { chromium } = require("playwright");
const path = require("path");

const APP_URL = process.env.APP_URL || "http://127.0.0.1:4175/";
const FILE = path.join(__dirname, "stress-test-workbench.xlsx");
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

async function run() {
  const browser = await chromium.launch({ headless: true });
  const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 390, height: 844 }, deviceScaleFactor: 2, hasTouch: true, isMobile: true });
  await context.addInitScript(() => localStorage.setItem("introPlayed", "true"));
  const page = await context.newPage();
  const cdp = await context.newCDPSession(page);
  const errors = [];
  page.on("pageerror", (e) => errors.push("pageerror: " + e.message));

  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  await page.evaluate(() => ensureXlsxLibs(false));
  await page.setInputFiles("#fileInput", FILE);
  await page.waitForFunction(() => document.getElementById("sheetSelect")?.options?.length > 0, null, { timeout: 20000 });
  await page.click("#loadBtn");
  await page.waitForFunction(() => document.querySelectorAll("#dataTable tbody tr").length > 5, null, { timeout: 20000 });
  await sleep(800);
  await page.evaluate(() => setSidebarOpen(false));
  await sleep(600);

  const results = [];
  const check = (name, ok, detail) => results.push({ name, ok: !!ok, detail });
  const collapsed = () => page.evaluate(() => document.body.classList.contains("hero-collapsed"));
  const scrollTop = () => page.evaluate(() => Math.round(document.getElementById("tableWrap").scrollTop));
  const setScroll = async (y) => { await page.evaluate((y) => { document.getElementById("tableWrap").scrollTop = y; }, y); await sleep(450); };
  const center = await page.evaluate(() => { const r = document.getElementById("tableWrap").getBoundingClientRect(); return { x: Math.round(r.left + r.width / 2), top: Math.round(r.top) }; });
  // Palec: ruch pionowy od y0 do y1 w krokach, opcjonalnie odczyt w połowie drogi.
  const drag = async (y0, y1, { steps = 16, midProbe = false } = {}) => {
    const x = center.x;
    await cdp.send("Input.dispatchTouchEvent", { type: "touchStart", touchPoints: [{ x, y: y0 }] });
    let probe = null;
    for (let i = 1; i <= steps; i++) {
      const y = Math.round(y0 + ((y1 - y0) * i) / steps);
      await cdp.send("Input.dispatchTouchEvent", { type: "touchMove", touchPoints: [{ x, y }] });
      await sleep(16);
      if (midProbe && i === Math.round(steps * 0.6)) {
        probe = await page.evaluate(() => { const g = document.getElementById("heroGrip"); return { pulling: g.classList.contains("pulling"), pull: g.style.getPropertyValue("--pull") }; });
      }
    }
    await cdp.send("Input.dispatchTouchEvent", { type: "touchEnd", touchPoints: [] });
    await sleep(450);
    return probe;
  };
  const collapseByScroll = async () => {
    await page.evaluate(() => { heroUserCollapsed = false; setHeroCollapsed(false); });
    await sleep(450);
    await setScroll(0);
    await setScroll(400);
    return collapsed();
  };

  // 1) + 2)
  check("przewinięcie w dół zwija nagłówek", await collapseByScroll());
  await setScroll(0);
  check("dojechanie do samej góry NIE rozwija", await collapsed(), `scrollTop=${await scrollTop()}`);

  // 3) lekkie pociągnięcie na górze
  const y0 = center.top + 120;
  const probe = await drag(y0, y0 + 40, { midProbe: true });
  check("lekkie pociągnięcie (40 px) nie rozwija", await collapsed());
  check("…ale uchwyt pokazuje postęp w trakcie", probe && probe.pulling && +probe.pull > 0, JSON.stringify(probe));
  check("po puszczeniu palca uchwyt wraca", await page.evaluate(() => !document.getElementById("heroGrip").classList.contains("pulling")));

  // 4) wyraźne pociągnięcie
  await drag(y0, y0 + 140);
  check("wyraźne pociągnięcie (140 px) rozwija", !(await collapsed()));

  // 5) ciągły ruch z dołu tabeli: palec w dół przewija do góry i jedzie dalej
  check("(ponowne zwinięcie)", await collapseByScroll());
  await setScroll(120); // niedaleko od góry, żeby jeden ruch palca dojechał i pociągnął dalej
  const startY = center.top + 60;
  await drag(startY, startY + 120 + 20, { steps: 20 });
  check("ruch kończący się ~na górze (bez dociągnięcia) nie rozwija", await collapsed(), `scrollTop=${await scrollTop()}`);
  await setScroll(120);
  await drag(startY, startY + 120 + 150, { steps: 30 });
  check("ciągły ruch do góry i dalej w dół rozwija", !(await collapsed()), `scrollTop=${await scrollTop()}`);

  // 6) kółko
  check("(ponowne zwinięcie)", await collapseByScroll());
  await setScroll(0);
  await page.mouse.move(center.x, center.top + 100);
  await page.mouse.wheel(0, -60);
  await sleep(450);
  check("jedno kółko w górę na górze nie rozwija", await collapsed());
  // (przy dpr 2 jeden „obrót" Playwrighta to −30 px; prawdziwy ząbek kółka ≈ −100 px)
  for (let i = 0; i < 8; i++) { await page.mouse.wheel(0, -60); await sleep(40); }
  await sleep(450);
  check("dłuższe kręcenie w górę na górze rozwija", !(await collapsed()));

  // 7) tap w uchwyt
  await page.tap("#heroGrip");
  await sleep(450);
  check("tap w uchwyt zwija", await collapsed());
  await page.tap("#heroGrip");
  await sleep(450);
  check("tap w uchwyt rozwija", !(await collapsed()));

  check("brak błędów strony", errors.length === 0, errors.slice(0, 3).join(" | "));
  await browser.close();

  let failed = 0;
  for (const r of results) {
    console.log(`${r.ok ? "✅" : "❌"} ${r.name}${r.detail ? `  (${r.detail})` : ""}`);
    if (!r.ok) failed += 1;
  }
  if (failed) { console.error(`\n${failed} z ${results.length} nie przeszło`); process.exit(1); }
  console.log(`\n✅ hero-pull-expand: ${results.length}/${results.length}`);
}

run().catch((e) => { console.error(e); process.exit(1); });
