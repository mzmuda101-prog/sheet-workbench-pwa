// hero-layer-playwright.js — tabela na własnej warstwie w trakcie zwijania nagłówka
// (`body.hero-animating .table-wrap { will-change: transform }` w app.css).
//
// Co musi być prawdą:
//   1. w trakcie animacji warstwa faktycznie jest (reguła działa),
//   2. obraz W POŁOWIE animacji (zwijanie i rozwijanie, z blokadą kolumny i bez)
//      jest taki sam z warstwą i bez niej — optymalizacja nie zmienia wyglądu,
//   3. po animacji warstwa znika (nie trzymamy pamięci GPU bez potrzeby).
// Uruchamiaj też z ENGINE=webkit (Safari).

const { chromium, webkit } = require("playwright");
const path = require("path");

const APP_URL = process.env.APP_URL || "http://127.0.0.1:4175/";
const FILE = path.join(__dirname, "stress-test-workbench.xlsx");
const ENGINE = process.env.ENGINE === "webkit" ? webkit : chromium;
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

async function run() {
  const browser = await ENGINE.launch({ headless: true });
  const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 390, height: 844 }, deviceScaleFactor: 2, hasTouch: true, isMobile: ENGINE === chromium });
  await context.addInitScript(() => localStorage.setItem("introPlayed", "true"));
  const page = await context.newPage();
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
  await page.evaluate(() => { setSidebarOpen(false); });
  await sleep(600);
  await page.evaluate(() => { document.getElementById("tableWrap").scrollTop = 300; });
  await sleep(300);
  // przewinięcie samo zwija nagłówek — wracamy do rozwiniętego, żeby 1. przypadek animował
  await page.evaluate(() => { heroUserCollapsed = false; setHeroCollapsed(false); });
  await sleep(600);

  const results = [];
  const check = (name, ok, detail) => results.push({ name, ok: !!ok, detail });
  await page.addStyleTag({ content: "html.no-hero-layer body.hero-animating .table-wrap { will-change: scroll-position !important; }" });

  // Uruchamia animację, zatrzymuje ją w połowie (jak ją widzi użytkownik) i trzyma
  // stan „w trakcie", żeby dało się zrobić dwa zrzuty: z warstwą i bez niej.
  const freezeMid = (on) => page.evaluate((on) => new Promise((resolve) => {
    heroUserCollapsed = on;
    setHeroCollapsed(on);
    requestAnimationFrame(() => requestAnimationFrame(() => {
      for (const a of document.getAnimations()) { a.pause(); a.currentTime = 150; }
      if (typeof readTableViewportHeight === "function") readTableViewportHeight();
      resolve(getComputedStyle(document.getElementById("tableWrap")).willChange);
    }));
  }), on);
  const holdAnimating = () => page.evaluate(() => { document.body.classList.add("hero-animating"); stopHeroSyncLoop(); });
  const shot = () => page.screenshot({ animations: "allow", caret: "hide" });
  const diff = (a, b) => page.evaluate(async ([a, b]) => {
    const load = (src) => new Promise((res) => { const i = new Image(); i.onload = () => res(i); i.src = src; });
    const [ia, ib] = await Promise.all([load(a), load(b)]);
    const px = (img) => { const c = document.createElement("canvas"); c.width = img.width; c.height = img.height; const x = c.getContext("2d"); x.drawImage(img, 0, 0); return x.getImageData(0, 0, c.width, c.height).data; };
    const da = px(ia), db = px(ib);
    let bad = 0;
    for (let i = 0; i < da.length; i += 4) {
      if (Math.max(Math.abs(da[i] - db[i]), Math.abs(da[i + 1] - db[i + 1]), Math.abs(da[i + 2] - db[i + 2])) > 40) bad += 1;
    }
    return bad;
  }, ["data:image/png;base64," + a.toString("base64"), "data:image/png;base64," + b.toString("base64")]);

  for (const freeze of [false, true]) {
    await page.evaluate((on) => { const el = document.getElementById("freezeFirstCol"); if (el.checked !== on) { el.checked = on; el.dispatchEvent(new Event("change", { bubbles: true })); } }, freeze);
    await sleep(400);
    for (const on of [true, false]) {
      const label = `${on ? "zwijanie" : "rozwijanie"}${freeze ? " + blokada kolumny" : ""}`;
      const wc = await freezeMid(on);
      check(`${label}: tabela ma warstwę w trakcie animacji`, /transform/.test(wc), wc);
      await holdAnimating();
      await sleep(120);
      const withLayer = await shot();
      await page.evaluate(() => document.documentElement.classList.add("no-hero-layer"));
      await sleep(120);
      const without = await shot();
      await page.evaluate(() => document.documentElement.classList.remove("no-hero-layer"));
      const bad = await diff(withLayer, without);
      check(`${label}: obraz w połowie animacji identyczny z warstwą i bez`, bad < 50, `różnych pikseli: ${bad}`);
      // koniec animacji
      await page.evaluate(() => { for (const a of document.getAnimations()) a.finish(); document.body.classList.remove("hero-animating"); });
      await sleep(450);
    }
  }
  const after = await page.evaluate(() => getComputedStyle(document.getElementById("tableWrap")).willChange);
  check("po animacji warstwa znika", !/transform/.test(after), after);
  check("brak błędów strony", errors.length === 0, errors.slice(0, 3).join(" | "));
  await browser.close();

  let failed = 0;
  for (const r of results) {
    console.log(`${r.ok ? "✅" : "❌"} ${r.name}${r.detail != null ? `  (${r.detail})` : ""}`);
    if (!r.ok) failed += 1;
  }
  if (failed) { console.error(`\n${failed} z ${results.length} nie przeszło`); process.exit(1); }
  console.log(`\n✅ hero-layer (${process.env.ENGINE || "chromium"}): ${results.length}/${results.length}`);
}

run().catch((e) => { console.error(e); process.exit(1); });
