// qol-phone-playwright.js — poprawki QoL zgłoszone przez Mateusza 2026-11-20:
//
//   1. telefon: sumy na dolnym pasku domyślnie schowane, karty arkuszy na cały rząd; Σ pokazuje / chowa
//      sumy i wybór jest zapamiętany; na komputerze sumy jak dotąd (bez Σ),
//   2. filtr „Zmienione wiersze”: tylko niezapisane / od otwarcia pliku; chip z ✕; licznik filtrów,
//   3. gęstość „Wygodna” na dotyku: niski wiersz z pliku rośnie o 40% do progu z czcionki (≈ 24 px przy 11 pt),
//      „Zwarta” i mysz zostają wierne plikowi.
//
// Uruchom z serwerem na APP_URL (domyślnie http://127.0.0.1:4175/). ENGINE=webkit — Safari.

const playwright = require("playwright");

const APP_URL = process.env.APP_URL || "http://127.0.0.1:4175/";
const ENGINE = process.env.ENGINE || "chromium";
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

async function open(browser, viewport, extra = {}) {
  const context = await browser.newContext({ serviceWorkers: "block", viewport, ...extra });
  await context.addInitScript(() => {
    if (!sessionStorage.getItem("__qp")) { sessionStorage.setItem("__qp", "1"); localStorage.clear(); localStorage.setItem("introPlayed", "true"); }
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
const shown = (page, sel) => page.evaluate((s) => { const el = document.querySelector(s); if (!el) return false; const r = el.getBoundingClientRect(); return r.width > 0 && r.height > 0 && getComputedStyle(el).display !== "none"; }, sel);

async function run() {
  const browser = await playwright[ENGINE].launch({ headless: true });
  const results = [];
  const check = (name, ok, detail = "") => results.push({ name, ok: !!ok, detail: typeof detail === "string" ? detail : JSON.stringify(detail) });

  // ── telefon ──
  {
    const { context, page, errors } = await open(browser, { width: 390, height: 844 }, { hasTouch: true, isMobile: ENGINE === "chromium" });
    // 1. sumy
    check("1. telefon: sumy domyślnie schowane", !(await shown(page, "#cellStatsBar")));
    check("1. telefon: przycisk Σ na dolnym pasku", (await shown(page, "#statsToggle")) && (await page.evaluate(() => !!document.getElementById("statsToggle").closest("#tableFoot"))));
    const tabsW = await page.evaluate(() => { const f = document.getElementById("tableFoot").getBoundingClientRect(); const t = document.getElementById("sheetTabs").getBoundingClientRect(); return { tabs: Math.round(t.width), foot: Math.round(f.width) }; });
    check("1. telefon: karty arkuszy zajmują większość rzędu", tabsW.tabs > tabsW.foot * 0.6, tabsW);
    await page.tap("#statsToggle");
    await sleep(300);
    check("1. Σ pokazuje sumy", await shown(page, "#cellStatsBar"));
    check("1. wybór zapamiętany", await page.evaluate(() => localStorage.getItem("swb-ph-stats-v1") === "1"));
    await page.tap("#statsToggle");
    await sleep(300);
    check("1. Σ drugi raz chowa sumy", !(await shown(page, "#cellStatsBar")));

    // 3. gęstość: wiersze z pliku 9 pt (12 px) — przy Wygodnej co najmniej 29 px
    const dens = await page.evaluate(async () => {
      for (let r = 0; r < 60; r++) currentSheetRowHeights[r] = { hpt: 9 };
      const h = () => { const tr = document.querySelector('#dataTable tbody tr[data-row-key="wide:10"]'); return tr ? Math.round(tr.getBoundingClientRect().height) : 0; };
      renderActiveTable();
      await new Promise((r) => setTimeout(r, 300));
      const comfy = { density: document.documentElement.dataset.density || "comfortable", h: h() };
      document.documentElement.dataset.density = "compact";
      renderActiveTable();
      await new Promise((r) => setTimeout(r, 300));
      const compact = h();
      delete document.documentElement.dataset.density;
      renderActiveTable();
      return { comfy, compact, excel: typeof isExcelLayoutEnabled === "function" && isExcelLayoutEnabled() };
    });
    // reguła dotyku: próg z czcionki pliku (≈ 20–26 px), niższy wiersz +40% do progu — nie stałe 29 px
    check("3. Wygodna na dotyku: niski wiersz z pliku rośnie do progu z czcionki (22–27 px)", dens.comfy.density === "comfortable" && dens.comfy.h >= 22 && dens.comfy.h <= 27, dens);
    check("3. Zwarta: wiersz zostaje niski jak w pliku", dens.compact > 0 && dens.compact < dens.comfy.h, dens);

    check("brak błędów strony (telefon)", errors.length === 0, errors.join(" | "));
    await context.close();
  }

  // ── komputer: sumy jak dotąd; filtr zmienionych wierszy ──
  {
    const { context, page, errors } = await open(browser, { width: 1280, height: 800 });
    check("1. komputer: bez przycisku Σ", !(await shown(page, "#statsToggle")));
    // mysz: Wygodna nie rusza wierszy z pliku (celność palca to sprawa dotyku)
    const mouseH = await page.evaluate(async () => {
      for (let r = 0; r < 60; r++) currentSheetRowHeights[r] = { hpt: 9 };
      delete document.documentElement.dataset.density;
      renderActiveTable();
      await new Promise((r) => setTimeout(r, 300));
      const tr = document.querySelector('#dataTable tbody tr[data-row-key="wide:10"]');
      return tr ? Math.round(tr.getBoundingClientRect().height) : 0;
    });
    check("3. mysz: Wygodna zostawia wiersz z pliku (18 px)", mouseH > 0 && mouseH <= 20, mouseH);
    // 2. zmiana dwóch komórek w różnych wierszach
    const rows = await page.evaluate(() => {
      updateSheetCell(10, 1, { value: "Zmiana A", type: "string" });
      updateSheetCell(20, 1, { value: "Zmiana B", type: "string" });
      if (typeof setDirtyState === "function") setDirtyState(true);
      return Object.keys(pendingEdits[currentSheetName] || {});
    });
    check("2. zmiany zapisane w pendingEdits", rows.length >= 2, rows);
    const view = (mode) => page.evaluate((m) => { swbSetChangedOnly(m); return { n: viewRows.length, rows: viewRows.map((r) => r.rowIndex0), chip: !!document.querySelector('#activeFilters [data-chip="changed"], #activeFilters .af-chip') }; }, mode);
    const u = await view("unsaved");
    check("2. „tylko niezapisane”: dokładnie zmienione wiersze", u.n === 2 && u.rows.includes(10) && u.rows.includes(20), u);
    const cnt = await page.evaluate(() => document.getElementById("panelToggleCount")?.textContent || document.getElementById("filterBadge")?.textContent || "");
    check("2. licznik filtrów liczy ten filtr", /[1-9]/.test(cnt), cnt);
    // po zapisie niezapisanych już nie ma, ale „od otwarcia pliku” je pamięta
    await page.evaluate(() => { pendingEdits = {}; });
    const u2 = await view("unsaved");
    check("2. po zapisie „tylko niezapisane” jest puste", u2.n === 0, u2);
    const a = await view("all");
    check("2. „od otwarcia pliku” pamięta zapisane zmiany", a.n === 2 && a.rows.includes(10) && a.rows.includes(20), a);
    const off = await view("");
    check("2. wyłączony filtr: wszystkie wiersze", off.n > 100, off);
    check("brak błędów strony (komputer)", errors.length === 0, errors.join(" | "));
    await context.close();
  }

  await browser.close();
  const failed = results.filter((r) => !r.ok);
  results.forEach((r) => console.log(`${r.ok ? "✓" : "✗"} ${r.name}${r.ok ? "" : "  — " + r.detail}`));
  if (failed.length) throw new Error(`${failed.length} z ${results.length} asercji nie przeszło (${ENGINE})`);
  console.log(`\n✅ qol-phone-playwright OK — ${results.length} asercji (${ENGINE})`);
}

run().catch((e) => { console.error(e.message || e); process.exit(1); });
