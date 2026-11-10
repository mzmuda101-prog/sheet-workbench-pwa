// drawer-tabs-playwright.js — panel narzędzi na telefonie: zakładki grup (app/drawer-tabs.js, audyt QoL S7).
//   node scripts/drawer-tabs-playwright.js            (Chromium)
//   ENGINE=webkit node scripts/drawer-tabs-playwright.js
//
// Pilnuje:
//   1. telefon: „Narzędzia” pokazuje JEDNĄ grupę (za pierwszym razem Filtry), reszta schowana;
//      bez pliku startuje „Plik”;
//   2. klik w zakładkę przełącza grupę, wybór jest pamiętany do kolejnego otwarcia;
//   3. „Aktywne filtry: n” + liczba na zakładce; „Wyczyść filtry” zdejmuje filtr;
//   4. „Znajdź ustawienie…” szuka we wszystkich grupach, po wyczyszczeniu wraca zakładka;
//   5. sekcja otwarta z kodu w innej grupie (Pomoc z menu ⋯) — zakładka idzie za nią i sekcję widać;
//   6. strzałki w pasku zakładek przełączają grupy;
//   7. komputer (pasek ikon, ≥1024 px): zakładek nie ma, panel jak dawniej.

const { chromium, webkit } = require("playwright");

const APP_URL = process.env.APP_URL || "http://127.0.0.1:4175/";
const ENGINE_NAME = process.env.ENGINE === "webkit" ? "webkit" : "chromium";
const ENGINE = ENGINE_NAME === "webkit" ? webkit : chromium;
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

async function run() {
  const fails = [];
  const check = (name, ok, detail) => {
    console.log(`${ok ? "✅" : "❌"} ${name}${ok ? "" : ` — ${JSON.stringify(detail)}`}`);
    if (!ok) fails.push(name);
  };
  const browser = await ENGINE.launch({ headless: true });
  const errors = [];

  const open = async (viewport, extra = {}) => {
    const ctx = await browser.newContext({ serviceWorkers: "block", viewport, ...extra });
    await ctx.addInitScript(() => localStorage.setItem("introPlayed", "true"));
    const page = await ctx.newPage();
    page.on("pageerror", (e) => errors.push(e.message));
    await page.goto(APP_URL, { waitUntil: "load" });
    await page.evaluate(() => document.getElementById("heroSplash")?.remove());
    return { ctx, page };
  };
  const state = (page) => page.evaluate(() => ({
    current: swbDrawerTabs._current(),
    focus: document.documentElement.classList.contains("drawer-focus"),
    visible: [...document.querySelectorAll(".sidebar-scroll > .sidebar-group")].filter((g) => g.offsetParent !== null).map((g) => g.id),
    selected: document.querySelector("#drawerTabs [aria-selected='true']")?.dataset.group || "",
    tabsShown: getComputedStyle(document.getElementById("drawerTabs")).display !== "none",
  }));
  // otwarty panel na telefonie zasłania „Narzędzia” — przełączamy tą samą funkcją, którą woła przycisk
  const toggle = async (page) => { await page.evaluate(() => toggleSidebar()); await sleep(450); };

  // ── telefon ──
  const { ctx, page } = await open({ width: 390, height: 844 }, { hasTouch: ENGINE_NAME === "webkit" });
  await toggle(page);
  const s0 = await state(page);
  check("1. bez pliku: startuje „Plik”, jedna grupa", s0.current === "group-data" && s0.visible.join() === "group-data" && s0.tabsShown, s0);
  await toggle(page);

  await page.evaluate(() => loadSampleFile());
  await page.waitForFunction(() => document.querySelector("#dataTable tbody tr[data-row-key]"), null, { timeout: 20000 });
  await sleep(800);
  await page.evaluate(() => document.querySelectorAll(".toast").forEach((t) => t.remove()));
  await toggle(page);
  const s1 = await state(page);
  check("1. z plikiem, pierwszy raz: „Filtry”, reszta schowana", s1.current === "group-filters" && s1.visible.join() === "group-filters" && s1.selected === "group-filters", s1);

  await page.click('#drawerTabs [data-group="group-edit"]');
  await sleep(200);
  const s2 = await state(page);
  check("2. klik w zakładkę przełącza grupę", s2.visible.join() === "group-edit" && s2.selected === "group-edit", s2);
  await toggle(page); await toggle(page);
  check("2. wybór pamiętany przy kolejnym otwarciu", (await state(page)).current === "group-edit", await state(page));

  // 3. aktywne filtry
  await toggle(page);
  await page.evaluate(() => { quickSearchEl.value = "Nowak"; applyQuickSearch(); });
  await sleep(700);
  await toggle(page);
  const act = await page.evaluate(() => ({ shown: !document.getElementById("drawerActive").hidden, text: document.getElementById("drawerActiveText").textContent, tab: document.querySelector('#drawerTabs [data-group="group-filters"] .drawer-tab-count')?.textContent }));
  check("3. „Aktywne filtry: 1” i liczba na zakładce Filtry", act.shown && act.text === "Aktywne filtry: 1" && act.tab === "1", act);
  await page.click("#drawerActiveClear");
  await sleep(700);
  const cleared = await page.evaluate(() => ({ shown: !document.getElementById("drawerActive").hidden, q: quickSearchEl.value, rows: viewRows.length }));
  check("3. „Wyczyść filtry” zdejmuje filtr i chowa wiersz", !cleared.shown && cleared.q === "", cleared);

  // 4. szukanie ustawienia
  await page.fill("#sidebarFinder", "kolumn");
  await sleep(300);
  const f1 = await state(page);
  check("4. szukanie pokazuje trafienia z wielu grup", !f1.focus && f1.visible.length > 1, f1);
  await page.fill("#sidebarFinder", "");
  await sleep(300);
  const f2 = await state(page);
  check("4. po wyczyszczeniu wraca zakładka", f2.focus && f2.visible.join() === f2.current, f2);

  // 5. Pomoc z kodu
  await page.evaluate(() => swbOpenHelpPanel());
  await sleep(600);
  const h = await page.evaluate(() => ({ current: swbDrawerTabs._current(), visible: !!document.getElementById("panel-shortcuts").offsetParent }));
  check("5. sekcja otwarta z kodu w innej grupie — zakładka za nią", h.current === "group-help" && h.visible, h);

  // 6. strzałki
  await page.focus('#drawerTabs [aria-selected="true"]');
  await page.keyboard.press("ArrowLeft");
  await sleep(150);
  const k = await page.evaluate(() => ({ current: swbDrawerTabs._current(), focused: document.activeElement?.dataset?.group }));
  check("6. ← w pasku zakładek przełącza na poprzednią grupę", k.current === "group-analyze" && k.focused === "group-analyze", k);
  await ctx.close();

  // ── 7. komputer: pasek ikon ──
  const d = await open({ width: 1280, height: 900 });
  await d.page.evaluate(() => loadSampleFile());
  await d.page.waitForFunction(() => document.querySelector("#dataTable tbody tr[data-row-key]"), null, { timeout: 20000 });
  await sleep(600);
  await d.page.evaluate(() => document.querySelectorAll(".toast").forEach((t) => t.remove()));
  await d.page.click('#sidebarRail [data-group="group-view"]');
  await sleep(450);
  const ds = await state(d.page);
  check("7. komputer: bez zakładek, pasek ikon pokazuje swoją grupę", !ds.tabsShown && !ds.focus && ds.visible.join() === "group-view", ds);
  await d.ctx.close();

  await browser.close();
  if (errors.length) { console.log("❌ błędy:", errors.slice(0, 5)); fails.push("błędy strony"); }
  if (fails.length) { console.log(`❌ drawer-tabs (${ENGINE_NAME}): ${fails.length} niezaliczone`); process.exit(1); }
  console.log(`✅ drawer-tabs (${ENGINE_NAME}): wszystko przeszło`);
}
run().catch((e) => { console.error(e); process.exit(1); });
