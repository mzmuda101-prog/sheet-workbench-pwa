// app-frame-playwright.js — strażnik „ramy ekranu” (app/app-frame.js, paczka B, 2026-09-27/28).
//
// Pilnuje tego, co widzi użytkownik wokół tabeli:
//   1. po wczytaniu nagłówek pokazuje NAZWĘ PLIKU i arkusz; szukanie jest widoczne bez
//      żadnego przełącznika; „Zapisz” jest, szary,
//   2. strona nie wystaje poza ekran na desktopie (było +24 px → scrollIntoView komórki
//      przesuwał całą stronę i tabela skakała pod kursorem),
//   3. kliknięcie wiersza NIE przesuwa tabeli (status z „• rekord N” w rzędzie z szukaniem
//      zawijał pasek i tabela skakała o ~50 px przy każdym kliknięciu),
//   4. po edycji 2 komórek „Zapisz” jest w akcencie z liczbą 2 i poprawną odmianą,
//   5. menu ⋯: otwiera się w całości na ekranie, Esc zamyka i oddaje fokus, ma linki,
//   6. szukanie trafia do paska „Ostatnie”, a klik w chip wykonuje je ponownie,
//   7. liczba przy „Narzędzia” = liczba pigułek aktywnych filtrów,
//   8. opcja „Pokaż uchwyt panelu” chowa zamknięty uchwyt,
//   9. telefon (375 px): nic nie wystaje w bok, nazwa pliku ma miejsce, rząd przycisków
//      przewija się w bok z wygaszeniem krawędzi (bezpiecznik przepełnienia).
//
// Uruchom z serwerem na APP_URL (domyślnie http://127.0.0.1:4175/). ENGINE=webkit — Safari.

const playwright = require("playwright");

const APP_URL = process.env.APP_URL || "http://127.0.0.1:4175/";
const ENGINE = process.env.ENGINE || "chromium";
// SLEEP_SCALE (run-tests.mjs): skraca przerwy „odczekaj po kroku” w szybkim przebiegu;
// powtórka porażki idzie ze skalą 1 (pełne przerwy). sleepFixed = przerwa, która MUSI
// przeczekać timer aplikacji (debounce, opóźnienie podpowiedzi, bezczynność) — bez skali.
// Ten test czeka na timery aplikacji (powiadomienia, bezczynność, schowek) — przy krótszych
// przerwach padał (sprawdzone SLEEP_SCALE=0.5, 2026-09-28), więc zawsze pełne przerwy.
const SLEEP_SCALE = 1;
const sleep = (ms) => new Promise((r) => setTimeout(r, Math.round(ms * SLEEP_SCALE)));
const sleepFixed = (ms) => new Promise((r) => setTimeout(r, ms)); // eslint-disable-line no-unused-vars

async function open(browser, viewport, extra = {}) {
  const context = await browser.newContext({ serviceWorkers: "block", viewport, ...extra });
  await context.addInitScript(() => {
    if (sessionStorage.getItem("__afInit")) return;
    sessionStorage.setItem("__afInit", "1");
    localStorage.clear();
    localStorage.setItem("introPlayed", "true");
  });
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push("pageerror: " + e.message));
  page.on("console", (m) => { if (m.type() === "error") errors.push("console: " + m.text()); });
  await page.goto(APP_URL + "?sample=300", { waitUntil: "load" });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  await page.evaluate(() => document.getElementById("loadSampleBtn")?.click());
  await page.waitForFunction(() => document.querySelector("#dataTable tbody tr[data-row-key]"), null, { timeout: 30000 });
  await page.evaluate(() => { if (typeof setSidebarOpen === "function") setSidebarOpen(false); document.querySelectorAll(".toast").forEach((t) => t.remove()); });
  await sleep(600);
  return { context, page, errors };
}

async function run() {
  const browser = await playwright[ENGINE].launch({ headless: true });
  const results = [];
  const check = (name, ok, detail = "") => results.push({ name, ok: !!ok, detail });

  // ── desktop ──
  let { context, page, errors } = await open(browser, { width: 1280, height: 900 });

  const s1 = await page.evaluate(() => ({
    title: document.getElementById("heroTitle").textContent,
    file: currentFileName,
    sheet: document.getElementById("heroSheetBtn").hidden ? null : document.getElementById("heroSheetName").textContent,
    current: currentSheetName,
    qsVisible: document.getElementById("quickSearchWrap").offsetParent !== null,
    readingToggle: !!document.getElementById("readingToggle"),
    save: !document.getElementById("heroSaveBtn").hidden,
    saveDirty: document.getElementById("heroSaveBtn").classList.contains("is-dirty"),
  }));
  check("1. nagłówek: nazwa pliku", s1.title === s1.file && !!s1.file, JSON.stringify(s1));
  check("1. nagłówek: bieżący arkusz", s1.sheet === s1.current, JSON.stringify(s1));
  check("1. szukanie widoczne bez przełącznika", s1.qsVisible && !s1.readingToggle, JSON.stringify(s1));
  check("1. Zapisz widoczny, szary", s1.save && !s1.saveDirty, JSON.stringify(s1));

  for (const vp of [{ width: 1280, height: 900 }, { width: 1100, height: 850 }, { width: 900, height: 800 }]) {
    await page.setViewportSize(vp);
    await sleep(400);
    const o = await page.evaluate(() => ({ docH: document.scrollingElement.scrollHeight, vh: innerHeight }));
    check(`2. strona nie wystaje (${vp.width}×${vp.height})`, o.docH <= o.vh + 1, JSON.stringify(o));
  }
  await page.setViewportSize({ width: 1280, height: 900 });
  await sleep(400);

  // 3. klik w wiersz nie przesuwa tabeli — przy kilku szerokościach (na węższych pasek
  //    szukania jest bliżej progu zawinięcia, więc dłuższy status by go przełamał)
  for (const vp of [{ width: 1280, height: 900 }, { width: 1100, height: 850 }, { width: 1000, height: 800 }, { width: 900, height: 800 }]) {
    await page.setViewportSize(vp);
    await sleep(400);
    await page.evaluate(() => { setFocusedCell(null, NaN); tableWrapEl.scrollTop = 0; });
    await sleep(150);
    const before = await page.evaluate(() => Math.round(tableWrapEl.getBoundingClientRect().top));
    const cell = await page.evaluate(() => {
      const td = document.querySelectorAll("#dataTable tbody tr[data-row-key]")[4].querySelector("td[data-col-index='1']");
      const r = td.getBoundingClientRect();
      return { x: r.left + r.width / 2, y: r.top + r.height / 2 };
    });
    await page.mouse.click(cell.x, cell.y);
    await sleep(300);
    await page.mouse.click(cell.x, cell.y + 70, { modifiers: ["Shift"] });
    await sleep(300);
    const after = await page.evaluate(() => ({ top: Math.round(tableWrapEl.getBoundingClientRect().top), status: statusEl.textContent }));
    check(`3. klik/zaznaczenie wiersza nie przesuwa tabeli (${vp.width} px)`, Math.abs(after.top - before) <= 1, JSON.stringify({ before, ...after }));
  }
  await page.setViewportSize({ width: 1280, height: 900 });
  await sleep(300);

  // 4. Zapisz po edycji
  const save = await page.evaluate(async () => {
    updateSheetCell(currentDisplayModel.rows[0].rowIndex0, 1, "AF-1");
    updateSheetCell(currentDisplayModel.rows[1].rowIndex0, 1, "AF-2");
    setDirtyState(true);
    await new Promise((r) => setTimeout(r, 50));
    const b = document.getElementById("heroSaveBtn");
    return { dirty: b.classList.contains("is-dirty"), n: document.getElementById("heroSaveCount").textContent, hidden: document.getElementById("heroSaveCount").hidden, label: b.getAttribute("aria-label") };
  });
  check("4. Zapisz po edycji: akcent + liczba 2", save.dirty && !save.hidden && save.n === "2", JSON.stringify(save));
  check("4. odmiana: „2 niezapisane zmiany”", /2 niezapisane zmiany/.test(save.label), save.label);

  // 5. menu ⋯
  await page.click("#appMenuBtn");
  await sleep(200);
  const menu = await page.evaluate(() => {
    const m = document.getElementById("appMenu");
    const r = m.getBoundingClientRect();
    return {
      open: !m.hidden, left: r.left, right: r.right, top: r.top, bottom: r.bottom, vw: innerWidth, vh: innerHeight,
      links: [...m.querySelectorAll("a[href]")].map((a) => a.href),
      hasLang: !!m.querySelector("#langSwitch"), hasTheme: !!m.querySelector("#themeToggle"), hasRefresh: !!m.querySelector("#brandRefresh"),
    };
  });
  check("5. menu ⋯ otwiera się w całości na ekranie", menu.open && menu.left >= 0 && menu.right <= menu.vw && menu.bottom <= menu.vh, JSON.stringify(menu));
  check("5. menu ⋯ ma język, motyw, odświeżenie i 2 linki", menu.hasLang && menu.hasTheme && menu.hasRefresh && menu.links.length === 2, JSON.stringify(menu));
  await page.keyboard.press("Escape");
  await sleep(150);
  const esc = await page.evaluate(() => ({ hidden: document.getElementById("appMenu").hidden, focus: document.activeElement?.id }));
  // Esc działa, gdy fokus jest w menu — z klawiatury: Enter na przycisku otwiera z fokusem w środku
  await page.focus("#appMenuBtn");
  await page.keyboard.press("Enter");
  await sleep(150);
  const inMenu = await page.evaluate(() => document.getElementById("appMenu").contains(document.activeElement));
  await page.keyboard.press("Escape");
  await sleep(150);
  const esc2 = await page.evaluate(() => ({ hidden: document.getElementById("appMenu").hidden, focus: document.activeElement?.id }));
  check("5. klawiatura: Enter → fokus w menu, Esc → zamyka i wraca na ⋯", inMenu && esc2.hidden && esc2.focus === "appMenuBtn", JSON.stringify({ esc, inMenu, esc2 }));
  if (!esc2.hidden) await page.click("#appMenuBtn");

  // 6. „Ostatnie”
  const recent = await page.evaluate(async () => {
    quickSearchEl.value = "Zakres A";
    commitQuickSearch();
    await new Promise((r) => setTimeout(r, 300));
    quickSearchEl.value = "";
    commitQuickSearch();
    await new Promise((r) => setTimeout(r, 300));
    const chips = [...document.querySelectorAll("#recentStrip .recent-chip")].map((c) => c.textContent);
    const chip = [...document.querySelectorAll("#recentStrip .recent-chip")].find((c) => c.textContent.includes("Zakres A"));
    if (chip) chip.click();
    await new Promise((r) => setTimeout(r, 300));
    return { chips, visible: !document.getElementById("recentStrip").hidden, q: quickSearchEl.value, rows: viewRows.length, total: baseRows.length };
  });
  check("6. zapytanie trafia do „Ostatnie”", recent.visible && recent.chips.some((c) => c.includes("Zakres A")), JSON.stringify(recent));
  check("6. klik w chip szuka ponownie", recent.q === "Zakres A" && recent.rows < recent.total, JSON.stringify(recent));

  // 7. licznik przy „Narzędzia”
  const count = await page.evaluate(() => ({
    badge: document.getElementById("panelToggleCount").hidden ? 0 : Number(document.getElementById("panelToggleCount").textContent),
    chips: document.querySelectorAll("#activeFilters .af-chip").length,
  }));
  check("7. liczba przy „Narzędzia” = liczba aktywnych filtrów", count.badge > 0 && count.badge === count.chips, JSON.stringify(count));

  // 8. opcja uchwytu
  const handle = await page.evaluate(async () => {
    const opt = document.getElementById("showPanelHandle");
    const vis = () => getComputedStyle(document.getElementById("panelHandle")).display !== "none";
    const a = vis();
    opt.checked = false; opt.dispatchEvent(new Event("change", { bubbles: true }));
    const b = vis();
    opt.checked = true; opt.dispatchEvent(new Event("change", { bubbles: true }));
    return { a, b, c: vis() };
  });
  check("8. „Pokaż uchwyt panelu” chowa i przywraca uchwyt", handle.a && !handle.b && handle.c, JSON.stringify(handle));

  check("brak błędów strony (desktop)", errors.length === 0, errors.join(" | "));
  await context.close();

  // ── paczka C: panel obok tabeli (≥1024 px) ──
  ({ context, page, errors } = await open(browser, { width: 1440, height: 860 }));
  const dock = async () => page.evaluate(() => {
    const side = document.querySelector(".sidebar").getBoundingClientRect();
    const main = document.querySelector(".table-panel").getBoundingClientRect();
    return {
      open: isSidebarOpen(), docked: document.documentElement.classList.contains("sidebar-docked"),
      sideRight: Math.round(side.right), mainLeft: Math.round(main.left), mainRight: Math.round(main.right),
      scrim: getComputedStyle(document.getElementById("sidebarScrim")).display,
      docH: document.scrollingElement.scrollHeight, vh: innerHeight, vw: innerWidth,
    };
  });
  await page.evaluate(() => setSidebarOpen(true));
  await sleep(500);
  const d1 = await dock();
  check("C1. ≥1024: panel obok tabeli, bez nachodzenia i bez zasłony", d1.docked && d1.open && d1.sideRight <= d1.mainLeft && d1.scrim === "none", JSON.stringify(d1));
  check("C1. strona nie wystaje z otwartym panelem", d1.docH <= d1.vh + 1 && d1.mainRight <= d1.vw, JSON.stringify(d1));
  // Esc w tabeli nie zamyka panelu obok tabeli
  await page.evaluate(() => { const td = document.querySelector("#dataTable tbody tr[data-row-key] td[data-col-index='1']"); td.click(); });
  await page.keyboard.press("Escape");
  await page.keyboard.press("Escape");
  await sleep(200);
  check("C2. Esc w tabeli nie zamyka panelu obok tabeli", (await dock()).open, "");
  // ‹ w panelu zamyka i zapamiętuje
  await page.click("#sidebarCloseBtn");
  await sleep(500);
  const d2 = await dock();
  const stored = await page.evaluate(() => localStorage.getItem("swb-panel-docked-open-v1"));
  check("C3. ‹ zamyka panel, tabela wraca na całą szerokość", !d2.open && d2.mainLeft < 100, JSON.stringify(d2));
  check("C3. zamknięcie jest zapamiętane", stored === "0", String(stored));
  // szukajka ustawień
  await page.evaluate(() => setSidebarOpen(true));
  await sleep(300);
  await page.fill("#sidebarFinder", "zawijaj");
  await sleep(200);
  const f1 = await page.evaluate(() => ({
    visible: [...document.querySelectorAll(".sidebar details.panel")].filter((d) => !d.hidden).map((d) => d.id),
    hit: !!document.querySelector("#wrapCells")?.closest("label")?.classList.contains("finder-hit"),
  }));
  check("C4. „Znajdź ustawienie”: zostaje sekcja Widok z podświetloną opcją", f1.visible.length === 1 && f1.visible[0] === "panel-view" && f1.hit, JSON.stringify(f1));
  await page.fill("#sidebarFinder", "");
  await sleep(200);
  const f2 = await page.evaluate(() => [...document.querySelectorAll(".sidebar details.panel")].filter((d) => d.hidden).length);
  check("C4. wyczyszczenie przywraca wszystkie sekcje", f2 === 0, String(f2));
  // przegrupowanie: brak „Akcji”, przyciski w nowych miejscach
  const g = await page.evaluate(() => ({
    actions: !!document.getElementById("panel-actions"),
    save: document.getElementById("saveBtn")?.closest("details")?.id,
    apply: !!document.getElementById("applyFilterBtn")?.closest("#group-filters"),
    reset: document.getElementById("resetWidthsBtn")?.closest("details")?.id,
    groups: [...document.querySelectorAll(".sidebar-group-title")].map((e) => e.textContent.trim()),
  }));
  check("C5. nowe grupy i przeniesione przyciski", !g.actions && g.save === "panel-file-sheet" && g.apply && g.reset === "panel-view" && g.groups.join("|") === "Plik|Filtry|Widok|Edycja|Analizy|Pomoc", JSON.stringify(g));
  // < 1024: nakładka z lekką zasłoną
  await page.setViewportSize({ width: 900, height: 800 });
  await sleep(500);
  await page.evaluate(() => setSidebarOpen(true));
  await sleep(400);
  const d3 = await dock();
  const blur = await page.evaluate(() => getComputedStyle(document.getElementById("sidebarScrim")).backdropFilter || getComputedStyle(document.getElementById("sidebarScrim")).webkitBackdropFilter);
  check("C6. <1024: nakładka z lekkim rozmyciem (jak dotąd, słabsze)", !d3.docked && d3.scrim !== "none" && /blur\(2px\)/.test(String(blur)), JSON.stringify({ ...d3, blur }));
  check("brak błędów strony (panel obok tabeli)", errors.length === 0, errors.join(" | "));
  await context.close();

  // ── przełączanie arkusza z nagłówka (plik z kilkoma arkuszami) ──
  {
    const ctx = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1280, height: 860 } });
    await ctx.addInitScript(() => localStorage.setItem("introPlayed", "true"));
    const pg = await ctx.newPage();
    const errs = [];
    pg.on("pageerror", (e) => errs.push(e.message));
    await pg.goto(APP_URL, { waitUntil: "load" });
    await pg.evaluate(() => { document.getElementById("heroSplash")?.remove(); try { ensureXlsxLibs && ensureXlsxLibs(false); } catch (_) {} });
    await pg.setInputFiles("#fileInput", require("path").join(__dirname, "stress-test-workbench.xlsx"));
    await pg.waitForFunction(() => document.getElementById("sheetSelect")?.options?.length > 1, null, { timeout: 20000 });
    await pg.click("#loadBtn");
    await pg.waitForFunction(() => document.querySelector("#dataTable tbody tr[data-row-key]"), null, { timeout: 20000 });
    await sleep(600);
    const before = await pg.evaluate(() => ({ sheet: currentSheetName, names: workbook.SheetNames, multi: document.getElementById("heroSheetBtn").classList.contains("is-multi"), tabsVisible: !!document.getElementById("sheetTabs")?.offsetParent }));
    check("S1. kilka arkuszy: przycisk arkusza klikalny, bez zakładek", before.multi && !before.tabsVisible, JSON.stringify(before));
    await pg.click("#heroSheetBtn");
    await sleep(200);
    const menu = await pg.evaluate(() => {
      const m = document.getElementById("sheetMenu");
      const r = m.getBoundingClientRect();
      return { open: !m.hidden, items: [...m.querySelectorAll(".sheet-menu-item")].map((b) => b.textContent.replace("✓", "").trim()), current: m.querySelector('[aria-selected="true"]')?.textContent.replace("✓", "").trim(), inView: r.left >= 0 && r.right <= innerWidth && r.bottom <= innerHeight };
    });
    check("S2. lista arkuszy pod nazwą: wszystkie arkusze, bieżący zaznaczony, w ekranie", menu.open && menu.items.length === before.names.length && menu.current === before.sheet && menu.inView, JSON.stringify(menu));
    const target = before.names.find((n) => n !== before.sheet);
    await pg.evaluate((name) => [...document.querySelectorAll("#sheetMenu .sheet-menu-item")].find((b) => b.textContent.includes(name)).click(), target);
    await pg.waitForFunction((name) => currentSheetName === name, target, { timeout: 20000 });
    await sleep(500);
    const after = await pg.evaluate(() => ({ sheet: currentSheetName, hero: document.getElementById("heroSheetName").textContent, open: !document.getElementById("sheetMenu").hidden }));
    check("S3. wybór z listy przełącza arkusz i zamyka listę", after.sheet === target && after.hero === target && !after.open, JSON.stringify(after));
    // klawiatura: Enter otwiera z fokusem na bieżącym, Esc zamyka i wraca na przycisk
    await pg.focus("#heroSheetBtn");
    await pg.keyboard.press("Enter");
    await sleep(150);
    const kb1 = await pg.evaluate(() => document.activeElement?.getAttribute("aria-selected"));
    await pg.keyboard.press("Escape");
    await sleep(150);
    const kb2 = await pg.evaluate(() => ({ hidden: document.getElementById("sheetMenu").hidden, focus: document.activeElement?.id }));
    check("S4. klawiatura: Enter → fokus na bieżącym arkuszu, Esc → zamyka", kb1 === "true" && kb2.hidden && kb2.focus === "heroSheetBtn", JSON.stringify({ kb1, kb2 }));
    if (process.env.SHOTS) {
      await pg.click("#heroSheetBtn");
      await sleep(250);
      await pg.screenshot({ path: `${process.env.SHOTS}/sheet-menu.png`, clip: { x: 0, y: 0, width: 700, height: 420 } });
    }
    check("brak błędów strony (arkusze)", errs.length === 0, errs.join(" | "));
    await ctx.close();

    // tablet (dotyk, 1024 px): bez pływającej pigułki arkusza — tylko przycisk w nagłówku
    const tctx = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1024, height: 768 }, hasTouch: true, isMobile: ENGINE === "chromium" });
    await tctx.addInitScript(() => localStorage.setItem("introPlayed", "true"));
    const tp = await tctx.newPage();
    await tp.goto(APP_URL, { waitUntil: "load" });
    await tp.evaluate(() => { document.getElementById("heroSplash")?.remove(); try { ensureXlsxLibs && ensureXlsxLibs(false); } catch (_) {} });
    await tp.setInputFiles("#fileInput", require("path").join(__dirname, "stress-test-workbench.xlsx"));
    await tp.waitForFunction(() => document.getElementById("sheetSelect")?.options?.length > 1, null, { timeout: 20000 });
    await tp.evaluate(() => loadBtn.click());
    await tp.waitForFunction(() => document.querySelector("#dataTable tbody tr[data-row-key]"), null, { timeout: 20000 });
    await sleep(500);
    const tab = await tp.evaluate(() => ({
      coarse: matchMedia("(pointer: coarse)").matches,
      fab: getComputedStyle(document.getElementById("sheetPickerFab")).display,
      hero: !document.getElementById("heroSheetBtn").hidden && document.getElementById("heroSheetBtn").classList.contains("is-multi"),
    }));
    check("S5. tablet: bez pływającej pigułki arkusza, przełącznik w nagłówku", tab.fab === "none" && tab.hero, JSON.stringify(tab));
    await tctx.close();
  }

  // ── telefon ──
  ({ context, page, errors } = await open(browser, { width: 375, height: 812 }, { hasTouch: true, isMobile: ENGINE === "chromium" }));
  const phone = await page.evaluate(async () => {
    const actions = document.getElementById("tableActions");
    actions.scrollLeft = 0;
    await new Promise((r) => setTimeout(r, 100));
    return {
      docW: document.documentElement.scrollWidth, vw: innerWidth,
      leftW: Math.round(document.querySelector(".hero-left").getBoundingClientRect().width),
      title: document.getElementById("heroTitle").textContent,
      overflowing: actions.scrollWidth > actions.clientWidth + 1,
      fadeR: actions.classList.contains("more-r"),
      fadeL: actions.classList.contains("more-l"),
    };
  });
  check("9. telefon: nic nie wystaje w bok", phone.docW <= phone.vw, JSON.stringify(phone));
  check("9. telefon: nazwa pliku ma miejsce w nagłówku", phone.leftW >= 80 && !!phone.title, JSON.stringify(phone));
  check("9. telefon: rząd przycisków wygasa tylko tam, gdzie jest co przewinąć", phone.overflowing ? (phone.fadeR && !phone.fadeL) : (!phone.fadeR && !phone.fadeL), JSON.stringify(phone));
  check("brak błędów strony (telefon)", errors.length === 0, errors.join(" | "));
  await context.close();
  await browser.close();

  const failed = results.filter((r) => !r.ok);
  results.forEach((r) => console.log(`${r.ok ? "✓" : "✗"} ${r.name}${r.ok ? "" : "  — " + r.detail}`));
  if (failed.length) throw new Error(`${failed.length} z ${results.length} asercji nie przeszło (${ENGINE})`);
  console.log(`\n✅ app-frame-playwright OK — ${results.length} asercji (${ENGINE})`);
}

run().catch((e) => { console.error(e); process.exit(1); });
