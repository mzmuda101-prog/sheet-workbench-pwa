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
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

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
