// hero-collapse-guard-playwright.js — zwijany nagłówek na telefonie: wysokość, strona, dymki.
//
// Zrzut Mateusza z iPhone'a (2026-10-25): przy zwiniętym nagłówku wisiał dymek „Zwiń lub
// rozwiń nagłówek” nad paskiem szukania i nie dało się go zdjąć. Przy okazji audytu wyszło:
// nagłówek ucinał się po wczytaniu pliku (limit wysokości zmierzony na ekranie startowym),
// a cała strona dawała się przewinąć o 18 px (32 px paddingu vs 14 px w obliczeniu tabeli).
//
// Co musi być prawdą (telefon 390×844, dotyk):
//   1. po wczytaniu pliku nagłówek pokazuje CAŁĄ treść (nic nie ucięte),
//   2. strona nie daje się przewinąć — ani przy rozwiniętym, ani przy zwiniętym nagłówku,
//   3. w trakcie zwijania/rozwijania dół tabeli trzyma się dołu ekranu (każda klatka),
//   4. treść, która urosła przy zwiniętym nagłówku, po rozwinięciu też nie jest ucięta,
//   5. przytrzymanie uchwytu (przed pociągnięciem) nie pokazuje dymka,
//   6. dymek z przytrzymania innego przycisku znika, gdy palec puści poza nim,
//      i po tapnięciu gdzie indziej (bezpiecznik na zgubione pointerup),
//   7. zwinięty uchwyt łapie dotyk też w pasie nad sobą (cel ≥ 28 px),
//   8. niezapisane zmiany: przy zwiniętym nagłówku kropka przy uchwycie.

const playwright = require("playwright");
const path = require("path");

const APP_URL = process.env.APP_URL || "http://127.0.0.1:4175/";
const ENGINE = process.env.ENGINE || "chromium";
const FILE = path.join(__dirname, "stress-test-workbench.xlsx");
const SLEEP_SCALE = Math.min(1, Math.max(0.1, Number(process.env.SLEEP_SCALE || 1)));
const sleep = (ms) => new Promise((r) => setTimeout(r, Math.round(ms * SLEEP_SCALE)));
const sleepFixed = (ms) => new Promise((r) => setTimeout(r, ms));

async function run() {
  const browser = await playwright[ENGINE].launch({ headless: true });
  const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 390, height: 844 }, hasTouch: true, isMobile: ENGINE === "chromium" });
  await context.addInitScript(() => localStorage.setItem("introPlayed", "true"));
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push("pageerror: " + e.message));
  const results = [];
  const check = (name, ok, detail = "") => results.push({ name, ok: !!ok, detail: typeof detail === "string" ? detail : JSON.stringify(detail) });

  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(() => { document.getElementById("heroSplash")?.remove(); return ensureXlsxLibs(false); });
  await page.setInputFiles("#fileInput", FILE);
  await page.waitForFunction(() => !!document.querySelector("#dataTable tbody tr[data-row-key]"), null, { timeout: 30000 });
  await sleep(800);
  await page.evaluate(() => { if (typeof setSidebarOpen === "function") setSidebarOpen(false); document.querySelectorAll(".toast").forEach((t) => t.remove()); });
  await sleep(400);

  const hero = () => page.evaluate(() => {
    const h = document.querySelector(".hero");
    const shown = h.getBoundingClientRect().height;
    const prev = h.style.maxHeight;
    h.style.maxHeight = "none";
    const natural = Math.max(h.getBoundingClientRect().height, h.scrollHeight);
    h.style.maxHeight = prev;
    return { collapsed: document.body.classList.contains("hero-collapsed"), shown: Math.round(shown), natural: Math.round(natural) };
  });
  const pageScroll = () => page.evaluate(() => {
    window.scrollTo(0, 500);
    const y = window.scrollY;
    window.scrollTo(0, 0);
    return { y, doc: document.scrollingElement.scrollHeight, vh: innerHeight };
  });
  const hintVisible = () => page.evaluate(() => !!document.querySelector(".cursor-hint.is-visible"));
  const touch = (type, sel, at) => page.evaluate(({ type, sel, at }) => {
    const el = sel ? document.querySelector(sel) : document.body;
    const r = el.getBoundingClientRect();
    const x = at ? at.x : r.left + r.width / 2;
    const y = at ? at.y : r.top + r.height / 2;
    el.dispatchEvent(new PointerEvent(type, { bubbles: true, cancelable: true, pointerType: "touch", pointerId: 11, isPrimary: true, clientX: x, clientY: y }));
  }, { type, sel, at });

  // 1. pełny nagłówek po wczytaniu pliku
  const h1 = await hero();
  check("1. nagłówek po wczytaniu pliku nie jest ucięty", !h1.collapsed && h1.shown >= h1.natural - 1, h1);

  // 2. strona nie przewija się (rozwinięty)
  const s2a = await pageScroll();
  check("2. strona nie przewija się (nagłówek rozwinięty)", s2a.y === 0 && s2a.doc <= s2a.vh, s2a);

  // 3. animacja: dół tabeli w każdej klatce przy dole ekranu (odczyt PO przeliczeniu w pętli)
  const frames = await page.evaluate(async () => {
    const panel = document.querySelector(".table-panel");
    const orig = window.readTableViewportHeight;
    const bottoms = [];
    window.readTableViewportHeight = function () { orig(); bottoms.push(Math.round(panel.getBoundingClientRect().bottom)); };
    setHeroCollapsed(true);
    await new Promise((r) => setTimeout(r, 500));
    setHeroCollapsed(false);
    await new Promise((r) => setTimeout(r, 500));
    window.readTableViewportHeight = orig;
    return { n: bottoms.length, min: Math.min(...bottoms), max: Math.max(...bottoms), vh: innerHeight };
  });
  check("3. animacja: dół tabeli trzyma się dołu ekranu w każdej klatce", frames.n >= 6 && frames.max - frames.min <= 2 && frames.max <= frames.vh, frames);

  // 2b. strona nie przewija się (zwinięty)
  await page.evaluate(() => setHeroCollapsed(true));
  await sleepFixed(450);
  const s2b = await pageScroll();
  check("2. strona nie przewija się (nagłówek zwinięty)", s2b.y === 0 && s2b.doc <= s2b.vh, s2b);

  // 4. treść urosła przy zwiniętym nagłówku → po rozwinięciu cała widoczna
  await page.evaluate(() => {
    const t = document.getElementById("heroTitle");
    t.dataset.prev = t.textContent;
    t.textContent = "Bardzo-długa-nazwa-pliku-ktora-sie-nie-miesci-w-jednej-linii-na-telefonie-raport-kwartalny.xlsx";
    t.style.whiteSpace = "normal";
    t.style.wordBreak = "break-all";
  });
  await page.evaluate(() => setHeroCollapsed(false));
  await sleepFixed(450);
  const h4 = await hero();
  check("4. treść, która urosła przy zwiniętym nagłówku, po rozwinięciu nie jest ucięta", !h4.collapsed && h4.natural > h1.natural && h4.shown >= h4.natural - 1, { h1, h4 });
  const s4 = await pageScroll();
  check("4. …a tabela pod wyższym nagłówkiem dalej mieści się w ekranie", s4.y === 0 && s4.doc <= s4.vh, s4);
  await page.evaluate(() => { const t = document.getElementById("heroTitle"); t.textContent = t.dataset.prev; t.style.whiteSpace = ""; t.style.wordBreak = ""; });
  await sleep(300);

  // 5. przytrzymanie uchwytu: bez dymka
  await touch("pointerdown", "#heroGrip");
  await sleepFixed(1300);
  const h5 = await hintVisible();
  await touch("pointerup", "#heroGrip");
  check("5. przytrzymanie uchwytu nie pokazuje dymka", !h5, "");

  // 6. dymek z przytrzymania znika po puszczeniu poza elementem / po tapnięciu gdzie indziej
  await touch("pointerdown", "#qsMoreBtn");
  await sleepFixed(1300);
  const h6a = await hintVisible();
  await touch("pointerup", null);
  await sleep(100);
  const h6b = await hintVisible();
  check("6. dymek pokazuje się na przytrzymanie przycisku", h6a, "");
  check("6. puszczenie palca poza przyciskiem chowa dymek", !h6b, "");
  await touch("pointerdown", "#qsMoreBtn");
  await sleepFixed(1300);
  await touch("pointerdown", "#dataTable tbody tr[data-row-key] td[data-col-index='1']");
  await sleep(100);
  check("6. tapnięcie gdzie indziej chowa wiszący dymek", !(await hintVisible()), "");
  await touch("pointerup", "#dataTable tbody tr[data-row-key] td[data-col-index='1']");

  // 7. cel dotyku zwiniętego uchwytu
  await page.evaluate(() => setHeroCollapsed(true));
  await sleepFixed(450);
  const hit = await page.evaluate(() => {
    const g = document.getElementById("heroGrip");
    const r = g.getBoundingClientRect();
    const at = (dy) => g.contains(document.elementFromPoint(r.left + r.width / 2, r.top + dy));
    const above = Math.max(0, Math.round(r.top) - 2);
    return { top: Math.round(r.top), h: Math.round(r.height), above: at(-above), mid: at(r.height / 2), below: at(r.height + 2), size: Math.round(r.height + above + 2) };
  });
  check("7. zwinięty uchwyt łapie dotyk w pasie nad sobą (cel ≥ 28 px)", hit.above && hit.mid && hit.below && hit.size >= 28, hit);
  await page.tap("#heroGrip");
  await sleepFixed(450);
  check("7. tap w uchwyt rozwija", !(await hero()).collapsed, "");

  // 8. kropka niezapisanych zmian przy zwiniętym nagłówku
  const dot = async () => page.evaluate(() => getComputedStyle(document.getElementById("heroGrip"), "::after").display !== "none");
  await page.evaluate(() => setHeroCollapsed(true));
  await sleepFixed(450);
  const d0 = await dot();
  await page.evaluate(async () => {
    updateSheetCell(currentDisplayModel.rows[0].rowIndex0, 1, "HG-1");
    setDirtyState(true);
    await new Promise((r) => setTimeout(r, 80));
  });
  const d1 = await dot();
  check("8. bez zmian: brak kropki przy uchwycie", !d0, "");
  check("8. niezapisane zmiany: kropka przy zwiniętym uchwycie", d1, "");
  await page.evaluate(() => setHeroCollapsed(false));
  await sleepFixed(450);
  check("8. rozwinięty nagłówek: kropki nie ma (widać „Zapisz”)", !(await dot()), "");

  check("brak błędów strony", errors.length === 0, errors.slice(0, 3).join(" | "));
  await browser.close();

  let failed = 0;
  for (const r of results) {
    console.log(`${r.ok ? "✅" : "❌"} ${r.name}${!r.ok && r.detail ? `  (${r.detail})` : ""}`);
    if (!r.ok) failed += 1;
  }
  if (failed) { console.error(`\n${failed} z ${results.length} nie przeszło`); process.exit(1); }
  console.log(`\n✅ hero-collapse-guard (${ENGINE}): ${results.length}/${results.length}`);
}

run().catch((e) => { console.error(e); process.exit(1); });
