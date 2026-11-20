// phone-panels-playwright.js — wstążka, etap 7c: panele na telefonie na cały ekran (docs/WSTAZKA.md; makieta
// docs/makiety/wstazka-v2.html, tryb „Telefon”, .fly). Plan etapu 7 zaakceptowany przez Mateusza 2026-11-20.
//
//   1. bez uchwytu „Wysuń”, „Znajdź ustawienie…” i zakładek grup (jak przy wstążce na komputerze),
//   2. Filtry (rząd Szukaj) → panel filtrów na cały ekran, z tytułem; ← zamyka,
//   3. jeden panel naraz: polecenie z karty (Analiza → Przegląd arkusza) pokazuje tylko ten panel,
//   4. Sortowanie i Sprawdzanie na karcie Dane, Formuły na karcie Formuły (na komputerze: szyna),
//   5. Wczytywanie, Skróty wyszukiwania i Dziennik w ⋯ (na komputerze: ekran Plik) — i otwierają panel,
//   6. komputer: przyciski szyny i linki ekranu Plik na swoich miejscach.
//
// Uruchom z serwerem na APP_URL (domyślnie http://127.0.0.1:4175/). ENGINE=webkit — Safari.

const playwright = require("playwright");

const APP_URL = process.env.APP_URL || "http://127.0.0.1:4175/";
const ENGINE = process.env.ENGINE || "chromium";
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

async function open(browser, viewport, extra = {}) {
  const context = await browser.newContext({ serviceWorkers: "block", viewport, ...extra });
  await context.addInitScript(() => {
    if (!sessionStorage.getItem("__pp")) { sessionStorage.setItem("__pp", "1"); localStorage.clear(); localStorage.setItem("introPlayed", "true"); }
  });
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push("pageerror: " + e.message));
  await page.goto(APP_URL + "?sample=300", { waitUntil: "load" });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  await page.evaluate(() => document.getElementById("loadSampleBtn")?.click());
  await page.waitForFunction(() => document.querySelector("#dataTable tbody tr[data-row-key]"), null, { timeout: 30000 });
  await sleep(600);
  await page.evaluate(() => document.querySelectorAll(".toast").forEach((t) => t.remove()));
  return { context, page, errors };
}

const shown = (page, sel) => page.evaluate((s) => { const el = document.querySelector(s); if (!el) return false; const r = el.getBoundingClientRect(); return r.width > 0 && r.height > 0 && getComputedStyle(el).visibility !== "hidden"; }, sel);
const inside = (page, sel, box) => page.evaluate(([s, b]) => !!document.querySelector(s)?.closest(b), [sel, box]);
const pane = (page) => page.evaluate(() => {
  const sb = document.querySelector(".sidebar");
  const r = sb.getBoundingClientRect();
  return {
    open: document.documentElement.classList.contains("sidebar-open"),
    full: r.left <= 1 && r.top <= 1 && r.width >= innerWidth - 1 && r.height >= innerHeight - 1,
    title: document.getElementById("rbPaneTitle")?.hidden ? "" : document.getElementById("rbPaneTitle")?.textContent.trim(),
    visible: [...sb.querySelectorAll("details.panel")].filter((d) => d.getBoundingClientRect().height > 0).map((d) => d.id),
  };
});

async function run() {
  const browser = await playwright[ENGINE].launch({ headless: true });
  const results = [];
  const check = (name, ok, detail = "") => results.push({ name, ok: !!ok, detail: typeof detail === "string" ? detail : JSON.stringify(detail) });
  const touch = { hasTouch: true, isMobile: ENGINE === "chromium" };

  {
    const { context, page, errors } = await open(browser, { width: 390, height: 844 }, touch);
    // 1.
    check("1. bez uchwytu „Wysuń”", !(await shown(page, "#panelHandle")));

    // 2. Filtry
    await page.tap("#panelToggle");
    await sleep(500);
    const f = await pane(page);
    check("2. Filtry: panel otwarty na cały ekran", f.open && f.full, f);
    check("2. Filtry: tytuł i tylko panele filtrów", /Filtr/.test(f.title) && f.visible.length > 0 && f.visible.every((id) => /filter/.test(id)), f);
    check("2. bez „Znajdź ustawienie…” i zakładek grup", !(await shown(page, ".sidebar-finder")) && !(await shown(page, "#drawerTabs")));
    const back = await page.evaluate(() => { const b = document.getElementById("sidebarCloseBtn"); const r = b.getBoundingClientRect(); return { x: Math.round(r.left), h: Math.round(r.height) }; });
    check("2. ← po lewej u góry, cel ≥ 40 px", back.x < 60 && back.h >= 40, back);
    await page.tap("#sidebarCloseBtn");
    await sleep(500);
    check("2. ← zamyka panel", !(await pane(page)).open);
    // licznik działających filtrów przy przycisku Filtry (dawniej też na zakładce grupy — drawer-tabs)
    await page.evaluate(() => { quickSearchEl.value = "Nowak"; applyQuickSearch(); });
    await sleep(700);
    const cnt = await page.evaluate(() => ({ n: document.getElementById("panelToggleCount")?.textContent, hidden: document.getElementById("panelToggleCount")?.hidden }));
    check("2. licznik filtrów przy Filtry („1”)", !cnt.hidden && cnt.n === "1", cnt);
    await page.evaluate(() => { quickSearchEl.value = ""; applyQuickSearch(); });
    await sleep(500);

    // 3. jeden panel z karty
    await page.evaluate(() => swbRibbon.select("an"));
    await sleep(200);
    await page.evaluate(() => document.querySelector('#phBottom [data-rb-panel="panel-workbench-analysis"]').click());
    await sleep(500);
    const a = await pane(page);
    check("3. Analiza → Przegląd arkusza: tylko ten panel, na cały ekran", a.open && a.full && JSON.stringify(a.visible) === '["panel-workbench-analysis"]' && a.title.length > 0, a);
    await page.tap("#sidebarCloseBtn");
    await sleep(500);

    // 4. szyna na kartach
    check("4. Sortowanie i Sprawdzanie na karcie Dane", (await inside(page, '.rb-rail-btn[data-rb-rail="sort"]', "#rbPanel-data")) && (await inside(page, '.rb-rail-btn[data-rb-rail="check"]', "#rbPanel-data")));
    check("4. Formuły na karcie Formuły", await inside(page, '.rb-rail-btn[data-rb-rail="formulas"]', "#rbPanel-fx"));
    await page.evaluate(() => swbRibbon.select("data"));
    await sleep(200);
    await page.evaluate(() => document.querySelector('#phBottom .rb-rail-btn[data-rb-rail="sort"]').scrollIntoView({ inline: "center" }));
    await page.tap('#phBottom .rb-rail-btn[data-rb-rail="sort"]');
    await sleep(500);
    const s = await pane(page);
    check("4. Sortowanie z karty Dane: panel sortowania na cały ekran", s.open && s.full && s.visible.includes("panel-sort-workbench"), s);
    await page.tap("#sidebarCloseBtn");
    await sleep(500);

    // 5. ⋯
    check("5. Wczytywanie, Skróty wyszukiwania, Dziennik w ⋯", await page.evaluate(() => ["panel-file-sheet", "panel-shortcuts", "panel-log"].every((p) => !!document.querySelector(`#appMenu [data-rb-panel="${p}"]`))));
    await page.tap("#appMenuBtn");
    await sleep(300);
    await page.tap('#appMenu [data-rb-panel="panel-log"]');
    await sleep(500);
    const l = await pane(page);
    check("5. Dziennik z ⋯ otwiera panel", l.open && l.full && l.visible.includes("panel-log"), l);
    await page.tap("#sidebarCloseBtn");
    await sleep(400);

    check("brak błędów strony (telefon)", errors.length === 0, errors.join(" | "));
    await context.close();
  }

  {
    const { context, page, errors } = await open(browser, { width: 1280, height: 800 });
    check("6. komputer: przyciski szyny w szynie", await page.evaluate(() => ["sort", "check", "formulas"].every((k) => !!document.querySelector(`.rb-rail-btn[data-rb-rail="${k}"]`)?.closest("#sidebarRail"))));
    check("6. komputer: Dziennik w ekranie Plik", await inside(page, '[data-rb-panel="panel-log"]', "#rbBackstage"));
    check("brak błędów strony (komputer)", errors.length === 0, errors.join(" | "));
    await context.close();
  }

  await browser.close();
  const failed = results.filter((r) => !r.ok);
  results.forEach((r) => console.log(`${r.ok ? "✓" : "✗"} ${r.name}${r.ok ? "" : "  — " + r.detail}`));
  if (failed.length) throw new Error(`${failed.length} z ${results.length} asercji nie przeszło (${ENGINE})`);
  console.log(`\n✅ phone-panels-playwright OK — ${results.length} asercji (${ENGINE})`);
}

run().catch((e) => { console.error(e.message || e); process.exit(1); });
