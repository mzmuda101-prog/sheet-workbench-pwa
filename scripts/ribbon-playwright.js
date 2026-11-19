// ribbon-playwright.js — strażnik wstążki (app/ribbon.js; docs/WSTAZKA.md, makieta docs/makiety/wstazka-v2.html).
//
// Etap 1 — rama:
//   1. na komputerze (≥ 769 px) body.rb-on, wstążka widoczna, nagłówek „hero” schowany,
//   2. 7 kart w kolejności z makiety, wybrana „Narzędzia główne”,
//   3. Szukaj, Zapisz, ⋯ i Cofnij to TE SAME elementy, przeniesione do wstążki,
//   4. pasek formuły tuż pod wstążką (≤ 82 px od góry),
//   5. strzałki przełączają karty i przenoszą fokus,
//   6. 768 px: wszystko wraca na swoje miejsca; 1440 px: znowu we wstążce,
//   7. 820 px: nic nie wystaje w bok, rząd kart ma strzałkę „+N”,
//   8. telefon: bez wstążki, układ jak dotąd.
//
// Etap 5 — mini pasek (M5.): po przeciągnięciu myszą nad zakresem, pod nim przy braku miejsca;
// przyciski idą tą samą drogą co wstążka; bez paska przy jednej komórce, klawiaturze i w Podglądzie.
//
// Uruchom z serwerem na APP_URL (domyślnie http://127.0.0.1:4175/). ENGINE=webkit — Safari.

const playwright = require("playwright");

const APP_URL = process.env.APP_URL || "http://127.0.0.1:4175/";
const ENGINE = process.env.ENGINE || "chromium";
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

async function open(browser, viewport, extra = {}) {
  const context = await browser.newContext({ serviceWorkers: "block", viewport, ...extra });
  await context.addInitScript(() => {
    if (sessionStorage.getItem("__rbInit")) return;
    sessionStorage.setItem("__rbInit", "1");
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
  await sleep(400);
  return { context, page, errors };
}

const inRibbon = (page, sel) => page.evaluate((s) => !!document.querySelector(s)?.closest("#ribbon"), sel);
const shown = (page, sel) => page.evaluate((s) => {
  const el = document.querySelector(s);
  if (!el) return false;
  const r = el.getBoundingClientRect();
  return r.width > 0 && r.height > 0 && getComputedStyle(el).visibility !== "hidden";
}, sel);

async function run() {
  const browser = await playwright[ENGINE].launch({ headless: true });
  const results = [];
  const check = (name, ok, detail = "") => results.push({ name, ok: !!ok, detail: String(detail) });

  // ── komputer ────────────────────────────────────────────────────────────
  {
    const { context, page, errors } = await open(browser, { width: 1440, height: 900 });
    check("1. body.rb-on na komputerze", await page.evaluate(() => document.body.classList.contains("rb-on")));
    check("1. wstążka widoczna", await shown(page, "#ribbon"));
    check("1. nagłówek hero schowany", !(await shown(page, ".hero")));
    // dolne rogi wstążki zaokrąglone jak karty aplikacji (.panel, --r-lg) — prośba Mateusza
    const corners = await page.evaluate(() => {
      const r = getComputedStyle(document.querySelector(".rb-r2")), p = getComputedStyle(document.querySelector(".table-panel"));
      return { bl: r.borderBottomLeftRadius, br: r.borderBottomRightRadius, panel: p.borderTopLeftRadius };
    });
    check("1. dolne rogi wstążki jak karty", corners.bl === corners.panel && corners.br === corners.panel && corners.panel !== "0px", JSON.stringify(corners));
    check("1. wstążka zamknięta w sobie (tabela i nagłówek nie są w #ribbon)", await page.evaluate(() => !document.querySelector("#ribbon .workspace, #ribbon .hero")));

    const tabs = await page.evaluate(() => Array.from(document.querySelectorAll('#ribbon [role="tab"][data-rtab]')).filter((t) => !t.hidden).map((t) => t.textContent.trim()));
    check("2. 7 kart w kolejności z makiety", JSON.stringify(tabs) === JSON.stringify(["Narzędzia główne", "Wstaw", "Dane", "Formuły", "Analiza", "Widok", "Asystent"]), tabs.join(" | "));
    const sel = await page.evaluate(() => document.querySelector('#ribbon [role="tab"][aria-selected="true"]')?.dataset.rtab);
    check("2. wybrana karta: Narzędzia główne", sel === "home", sel);

    for (const id of ["#quickSearch", "#heroSaveBtn", "#appMenuBtn", "#undoBtn", "#redoBtn"]) {
      check(`3. ${id} we wstążce`, await inRibbon(page, id));
    }
    check("3. Szukaj widoczne", await shown(page, "#quickSearch"));

    const fbTop = await page.evaluate(() => Math.round(document.getElementById("formulaBar").getBoundingClientRect().top));
    check("4. pasek formuły ≤ 82 px od góry", fbTop <= 82, fbTop + " px");

    await page.focus('#ribbon [role="tab"][aria-selected="true"]');
    await page.keyboard.press("ArrowRight");
    const after = await page.evaluate(() => ({ sel: document.querySelector('#ribbon [role="tab"][aria-selected="true"]')?.dataset.rtab, focus: document.activeElement?.dataset?.rtab }));
    check("5. strzałka w prawo → Wstaw (wybór i fokus)", after.sel === "ins" && after.focus === "ins", JSON.stringify(after));
    const panelOk = await page.evaluate(() => {
      const ps = Array.from(document.querySelectorAll("#ribbon .rb-panel[data-rtab]"));
      return ps.length === 7 && ps.filter((p) => !p.hidden).map((p) => p.dataset.rtab).join() === "ins";
    });
    check("5. widać tylko panel karty Wstaw", panelOk);

    // ── zadanie 2: polecenia kart i menu trybu ──
    const named = await page.evaluate(() => Array.from(document.querySelectorAll("#ribbon .rb-panel[data-rtab]")).map((p) => {
      const btns = Array.from(p.querySelectorAll("button, select"));
      const noName = btns.filter((b) => !(b.getAttribute("aria-label") || b.textContent.trim()));
      return { tab: p.dataset.rtab, n: btns.length, noName: noName.map((b) => b.id || b.className) };
    }));
    check("T2. każda karta ma polecenia z nazwą", named.every((x) => x.n > 0 && !x.noName.length), JSON.stringify(named.filter((x) => !x.n || x.noName.length)));
    // bez dubli (decyzja Mateusza): żadne dwa przyciski wstążki nie otwierają tego samego panelu
    const dupPanels = await page.evaluate(() => {
      const seen = {}; const dup = [];
      document.querySelectorAll("#ribbon [data-rb-panel]").forEach((b) => { const k = b.dataset.rbPanel; if (seen[k]) dup.push(k); seen[k] = 1; });
      return dup;
    });
    check("T2. bez dubli: jeden przycisk na panel", dupPanels.length === 0, dupPanels.join());
    for (const [id, tab] of [["#transcribeBtn", "mode"], ["#entryBtn", "mode"], ["#reportBtn", "an"], ["#resetSortBtn", "data"], ["#wideLongToggle", "data"], ["#aiMenuItem", "ai"]]) {
      check(`T2. ${id} przeniesiony do wstążki (${tab})`, await inRibbon(page, id));
    }
    check("T2. „Więcej ▾” schowane przy wstążce", !(await shown(page, "#moreMenuBtn")));

    // powiększenie na dolnym pasku (nie na karcie Widok); Szukaj rozwija się szeroko, póki w nim pracujesz
    check("Z. powiększenie na dolnym pasku, nie we wstążce", await page.evaluate(() => !!document.getElementById("zoomGroup").closest("#tableFoot")) && !(await inRibbon(page, "#zoomGroup")));
    await page.click("#quickSearch");
    await sleep(300);
    const qsWide = await page.evaluate(() => Math.round(document.getElementById("quickSearchWrap").getBoundingClientRect().width));
    check("Z. Szukaj z kursorem: karta ≥ 600 px", qsWide >= 600, qsWide + " px");
    await page.click("#qsMoreBtn");
    await sleep(200);
    const rowsOk = await page.evaluate(() => {
      const top = (id) => Math.round(document.getElementById(id).getBoundingClientRect().top);
      return top("qsMoreBtn") === top("nlqOpenBtn") && top("qsAllSheets") > top("qsMoreBtn");
    });
    check("Z. ⚙ zostaje w rzędzie pola, opcje w drugim rzędzie", rowsOk);
    await page.mouse.click(5, 500);
    await sleep(300);
    check("Z. klik poza Szukaj zwija kartę", await page.evaluate(() => !document.getElementById("quickSearchWrap").classList.contains("qs-expanded")));

    await page.click("#rbTab-an");
    await page.click('#ribbon [data-rb-panel="panel-aggregation-workbench"]');
    await sleep(300);
    check("T2. Analiza → Agregacje otwiera panel", await page.evaluate(() => document.getElementById("panel-aggregation-workbench").open && document.documentElement.classList.contains("sidebar-open")));
    await page.evaluate(() => { if (typeof setSidebarOpen === "function") setSidebarOpen(false); });

    await page.click("#rbModeBtn");
    check("T2. menu trybu otwarte", await shown(page, "#rbModeMenu"));
    await page.keyboard.press("Escape");
    check("T2. Esc zamyka menu trybu i oddaje fokus", await page.evaluate(() => document.getElementById("rbModeMenu").hidden && document.activeElement?.id === "rbModeBtn"));
    await page.click("#rbModeBtn");
    await page.click('#rbModeMenu [data-rb-mode="view"]');
    await sleep(250);
    const view = await page.evaluate(() => ({
      r2: (() => { const r = document.querySelector(".rb-r2").getBoundingClientRect(); return r.height > 0; })(),
      name: document.getElementById("rbName")?.textContent || "",
      title: document.getElementById("heroTitle")?.textContent || "",
      edit: (() => { const b = document.getElementById("rbEditBtn"); return !!b && b.getBoundingClientRect().width > 0; })(),
      save: (() => { const b = document.getElementById("heroSaveBtn"); return b.getBoundingClientRect().width > 0; })(),
    }));
    check("T2. Podgląd: bez rzędu poleceń", !view.r2, JSON.stringify(view));
    check("T2. Podgląd: nazwa pliku w rzędzie 1", view.title && view.name.includes(view.title), JSON.stringify(view));
    check("T2. Podgląd: „Edytuj” zamiast „Zapisz”", view.edit && !view.save, JSON.stringify(view));
    await page.click("#rbEditBtn");
    await sleep(250);
    const back = await page.evaluate(() => ({ r2: document.querySelector(".rb-r2").getBoundingClientRect().height > 0, tab: document.querySelector('#ribbon [role="tab"][aria-selected="true"]')?.dataset.rtab, edit: typeof swbIsEditMode === "function" && swbIsEditMode() }));
    check("T2. Edytuj: wraca Edycja i karta Analiza", back.r2 && back.tab === "an" && back.edit, JSON.stringify(back));

    await page.click("#rbModeBtn");
    await page.click("#rbModeMenu #transcribeBtn");
    await sleep(500);
    check("T2. Spisywanie z menu trybu otwiera spisywanie", await page.evaluate(() => document.body.classList.contains("tr-active")));
    await page.keyboard.press("Escape");
    await sleep(300);
    await page.evaluate(() => { document.querySelectorAll(".tr-overlay, .tr-root").forEach(() => {}); });

    await page.setViewportSize({ width: 768, height: 900 });
    await page.waitForFunction(() => !document.body.classList.contains("rb-on"), null, { timeout: 5000 }).catch(() => {});
    await sleep(100);
    check("6. 768 px: bez rb-on", !(await page.evaluate(() => document.body.classList.contains("rb-on"))));
    check("6. 768 px: Szukaj wróciło na miejsce", !(await inRibbon(page, "#quickSearch")) && (await page.evaluate(() => !!document.getElementById("quickSearch").closest("#quickSearchWrap")?.closest(".table-panel"))));
    check("6. 768 px: Zapisz wróciło do nagłówka", await page.evaluate(() => !!document.getElementById("heroSaveBtn").closest(".hero")));
    check("6. 768 px: nagłówek widoczny", await shown(page, ".hero"));
    check("Z. 768 px: powiększenie wróciło do paska nad tabelą", await page.evaluate(() => !!document.getElementById("zoomGroup").closest("#toolbarMain")));
    check("T2. 768 px: Spisywanie wróciło do menu Więcej", await page.evaluate(() => !!document.getElementById("transcribeBtn").closest("#moreMenu")));
    await page.setViewportSize({ width: 1440, height: 900 });
    await sleep(300);
    check("6. z powrotem 1440 px: Szukaj we wstążce", await inRibbon(page, "#quickSearch"));

    await page.setViewportSize({ width: 820, height: 900 });
    await sleep(400);
    const ov = await page.evaluate(() => document.documentElement.scrollWidth - innerWidth);
    check("7. 820 px: nic nie wystaje w bok", ov <= 0, ov + " px");
    const nudge = await page.evaluate(() => {
      const n = document.querySelector('#ribbon [data-strip="tabs"] .rb-nudge-r');
      return n ? { vis: n.getBoundingClientRect().width > 0, txt: n.textContent.trim() } : null;
    });
    check("7. 820 px: rząd kart ma strzałkę „+N”", nudge && nudge.vis && /^\+\d+/.test(nudge.txt), JSON.stringify(nudge));
    const lap = await page.evaluate(() => {
      const w = document.getElementById("quickSearchWrap"), q = w.getBoundingClientRect(), m = document.getElementById("rbModeBtn").getBoundingClientRect();
      // najdalej wysunięta WIDOCZNA zawartość pola (tekst, ✦, ⚙), nie sama ramka
      const kids = Array.from(w.querySelectorAll("*")).map((e) => e.getBoundingClientRect()).filter((r) => r.width > 0 && r.height > 0);
      const clipped = getComputedStyle(w).overflow === "hidden"; // ramka przycina — wystająca zawartość jest niewidoczna
      return { qR: Math.round(clipped ? q.right : Math.max(q.right, ...kids.map((r) => r.right))), mL: Math.round(m.left), qW: Math.round(q.width) };
    });
    check("7. 820 px: zwinięte Szukaj nie nachodzi na przycisk trybu", lap.qR <= lap.mL && lap.qW <= 40, JSON.stringify(lap));
    await page.click("#quickSearchWrap");
    await sleep(300);
    const open820 = await page.evaluate(() => ({ focus: document.activeElement?.id, w: Math.round(document.getElementById("quickSearchWrap").getBoundingClientRect().width) }));
    check("7. 820 px: klik w ikonę Szukaj rozwija pole z fokusem", open820.focus === "quickSearch" && open820.w >= 400, JSON.stringify(open820));
    await page.keyboard.press("Escape");

    check("brak błędów strony (komputer)", errors.length === 0, errors.join(" | "));
    await context.close();
  }

  // ── zadanie 3: filtry w rzędzie paska formuły, „Ostatnie” pod Szukaj, arkusze pod tabelą ──
  {
    const { context, page, errors } = await open(browser, { width: 1440, height: 900 });
    await page.fill("#quickSearch", "Anna");
    await page.press("#quickSearch", "Enter");
    await page.waitForFunction(() => { const a = document.getElementById("activeFilters"); return a && !a.classList.contains("hidden") && a.children.length > 0; }, null, { timeout: 10000 }).catch(() => {});
    const row = await page.evaluate(() => {
      const af = document.getElementById("activeFilters"), nb = document.getElementById("nameBox"), fx = document.getElementById("formulaInput");
      const a = af.getBoundingClientRect(), n = nb.getBoundingClientRect();
      return { inRow: !!af.closest("#rbFbarRow") && !!nb.closest("#rbFbarRow"), dTop: Math.abs(Math.round(a.top - n.top)), fxW: Math.round(fx.getBoundingClientRect().width), visible: a.width > 0 };
    });
    check("T3. chipy filtrów w jednym rzędzie z paskiem formuły", row.inRow && row.visible && row.dTop <= 2, JSON.stringify(row));
    check("T3. pole formuły ≥ 160 px", row.fxW >= 160, JSON.stringify(row));
    await page.setViewportSize({ width: 820, height: 900 });
    await page.waitForTimeout(300);
    const narrow = await page.evaluate(() => ({ fxW: Math.round(document.getElementById("formulaInput").getBoundingClientRect().width), ov: document.documentElement.scrollWidth - innerWidth }));
    check("T3. 820 px: pole formuły ≥ 160 px, nic nie wystaje", narrow.fxW >= 160 && narrow.ov <= 0, JSON.stringify(narrow));
    await page.setViewportSize({ width: 1440, height: 900 });
    await page.waitForTimeout(300);
    // Podgląd: pasek formuły znika, chipy zostają
    await page.click("#rbModeBtn");
    await page.click('#rbModeMenu [data-rb-mode="view"]');
    await page.waitForTimeout(250);
    check("T3. Podgląd: chipy filtrów dalej widoczne", await shown(page, "#activeFilters"));
    await page.click("#rbEditBtn");
    // Ostatnie: historia szukania pod polem, tylko gdy jest fokus
    await page.click("#ribbon .rb-tab[data-rtab=home]");
    const closed = await shown(page, "#recentStrip");
    await page.focus("#quickSearch");
    await page.waitForTimeout(250);
    const rec = await page.evaluate(() => { const r = document.getElementById("recentStrip"); const b = r.getBoundingClientRect(); return { inSearch: !!r.closest("#quickSearchWrap"), vis: b.width > 0 && b.height > 0, txt: r.textContent }; });
    check("T3. Ostatnie: bez fokusu schowane, z fokusem pod Szukaj z „Anna”", !closed && rec.inSearch && rec.vis && rec.txt.includes("Anna"), JSON.stringify({ closed, ...rec }));
    // Esc w polu Szukaj: wychodzi z pola i karta Szukaj się zwija (także z rozwiniętymi opcjami ⚙)
    await page.click("#qsMoreBtn");
    await page.focus("#quickSearch");
    await page.keyboard.type("Jan");
    await sleep(250);
    await page.keyboard.press("Escape");
    await sleep(300);
    const esc = await page.evaluate(() => { const w = document.getElementById("quickSearchWrap"); return { inField: w.contains(document.activeElement), w: Math.round(w.getBoundingClientRect().width), expanded: w.classList.contains("qs-expanded"), live: !document.getElementById("qsLiveResults").classList.contains("hidden") }; });
    check("T3. Esc w Szukaj: wyjście z pola, karta zwinięta, bez podglądu", !esc.inField && esc.w <= 240 && !esc.expanded && !esc.live, JSON.stringify(esc));
    await page.fill("#quickSearch", "");
    check("brak błędów strony (T3)", errors.length === 0, errors.join(" | "));
    await context.close();
  }
  {
    const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1440, height: 900 } });
    await context.addInitScript(() => localStorage.setItem("introPlayed", "true"));
    const page = await context.newPage();
    await page.goto(APP_URL, { waitUntil: "load" });
    await page.evaluate(() => document.getElementById("heroSplash")?.remove());
    await page.setInputFiles("#fileInput", require("path").join(__dirname, "stress-test-workbench.xlsx"));
    await page.waitForFunction(() => document.querySelectorAll("#sheetTabs [role=tab]").length > 1, null, { timeout: 30000 });
    await page.waitForTimeout(300);
    // Etap 4 (makieta v2, dolny pasek): karty arkuszy na dole od lewej, obok nich przycisk listy arkuszy
    // (przełączanie + Nowy / Zmień nazwę / Usuń) — przy kilku arkuszach sama ikona, bez dublowania nazwy.
    const sh = await page.evaluate(() => {
      const b = document.getElementById("heroSheetBtn"), st = document.getElementById("sheetTabs"), foot = document.getElementById("tableFoot");
      const tabs = Array.from(st.querySelectorAll("[role=tab]"));
      const name = b.querySelector(".hero-sheet-name");
      return {
        btnInFoot: !!b.closest("#tableFoot"), tabsInFoot: !!st.closest("#tableFoot"), notInRow: !b.closest("#rbFbarRow"),
        tabsVis: st.getBoundingClientRect().width > 0, n: tabs.length, sheets: workbook.SheetNames.length,
        // kolejność na ekranie (CSS order), nie w DOM — row-grow dopisuje swoje przyciski później
        first: (() => { const bl = b.getBoundingClientRect(), tl = st.getBoundingClientRect(); const others = Array.from(foot.children).filter((c) => c !== b && c !== st && c.getBoundingClientRect().width > 0); return bl.right <= tl.left + 1 && others.every((c) => c.getBoundingClientRect().left >= tl.right - 1 || c.getBoundingClientRect().top >= tl.bottom - 1); })(), nameHidden: !name || name.getBoundingClientRect().width === 0,
        label: b.getAttribute("aria-label") || "",
      };
    });
    check("E4. karty arkuszy i przycisk listy na dolnym pasku (od lewej)", sh.btnInFoot && sh.tabsInFoot && sh.notInRow && sh.tabsVis && sh.n === sh.sheets && sh.first, JSON.stringify(sh));
    check("E4. przy kilku arkuszach przycisk listy bez nazwy (nazwa na karcie), z nazwą dostępną", sh.nameHidden && sh.label.length > 3, JSON.stringify(sh));
    const target = await page.evaluate(() => workbook.SheetNames[1]);
    await page.click(`#sheetTabs [data-sheet="${target}"]`);
    await page.waitForFunction((n) => currentSheetName === n, target, { timeout: 30000 }).catch(() => {});
    await page.waitForTimeout(300);
    check("E4. klik w kartę przełącza arkusz", await page.evaluate((n) => currentSheetName === n && document.querySelector(`#sheetTabs [data-sheet="${n}"]`)?.getAttribute("aria-selected") === "true", target));
    await page.click("#heroSheetBtn");
    await page.waitForTimeout(200);
    const menu = await page.evaluate(() => { const m = document.getElementById("sheetMenu"), r = m.getBoundingClientRect(); return { open: !m.hidden, items: m.querySelectorAll(".sheet-menu-item").length, ops: m.querySelectorAll(".sheet-op-item").length, top: Math.round(r.top), bottom: Math.round(r.bottom), vh: innerHeight }; });
    check("E4. lista arkuszy otwiera się w górę, cała na ekranie, z Nowy / Zmień nazwę / Usuń", menu.open && menu.items > 1 && menu.ops >= 3 && menu.top >= 0 && menu.bottom <= menu.vh, JSON.stringify(menu));
    await page.keyboard.press("Escape");
    await page.setViewportSize({ width: 768, height: 900 });
    await page.waitForFunction(() => !document.body.classList.contains("rb-on"), null, { timeout: 5000 }).catch(() => {});
    await page.waitForTimeout(100);
    check("E4. 768 px: przycisk arkusza i karty wracają do nagłówka", await page.evaluate(() => !!document.getElementById("heroSheetBtn").closest(".hero") && !!document.getElementById("sheetTabs").closest(".hero")));
    await context.close();
  }

  // ── poprawki po przeglądzie: fokus, menu przy przejściu progu, pusty stan, skróty ──
  {
    const { context, page, errors } = await open(browser, { width: 1440, height: 900 });
    // F1: fokus przeżywa przejście 769 → 768 (obrót iPada w trakcie szukania)
    await page.focus("#quickSearch");
    await page.keyboard.type("Ann");
    await page.setViewportSize({ width: 768, height: 900 });
    await page.waitForFunction(() => !document.body.classList.contains("rb-on"), null, { timeout: 5000 }).catch(() => {});
    await sleep(150);
    check("F1. 1440→768: fokus zostaje w Szukaj", await page.evaluate(() => document.activeElement?.id === "quickSearch"));
    await page.fill("#quickSearch", "");
    // F2: otwarte „Widok ▾” nie zostaje przyklejone w rogu po przejściu na wstążkę
    await page.click("#viewMenuBtn");
    await sleep(150);
    await page.setViewportSize({ width: 1024, height: 900 });
    await page.waitForFunction(() => document.body.classList.contains("rb-on"), null, { timeout: 5000 }).catch(() => {});
    await sleep(200);
    check("F2. 768→1024: menu Widok zamknięte", await page.evaluate(() => document.getElementById("viewMenu").hidden && document.getElementById("viewMenuBtn").getAttribute("aria-expanded") !== "true"));
    // F3: wybór trybu z klawiatury i „Edytuj” oddają fokus przyciskowi trybu
    await page.focus("#rbModeBtn");
    await page.keyboard.press("Enter");
    await sleep(100);
    await page.keyboard.press("ArrowDown");
    await page.keyboard.press("Enter");
    await sleep(250);
    check("F3. Podgląd z klawiatury: fokus na przycisku trybu", await page.evaluate(() => document.activeElement?.id === "rbModeBtn" && document.body.classList.contains("rb-reading")));
    await page.focus("#rbEditBtn");
    await page.keyboard.press("Enter");
    await sleep(250);
    check("F3. Edytuj: fokus na przycisku trybu", await page.evaluate(() => document.activeElement?.id === "rbModeBtn" && !document.body.classList.contains("rb-reading")));
    // F4: menu nie obiecuje skrótu, którego nie ma (Ctrl+Alt+E = AltGr+E = „ę” na polskim Windowsie)
    check("F4. menu trybu bez nieistniejącego Ctrl+Alt+E", await page.evaluate(() => !document.getElementById("rbModeMenu").textContent.includes("Ctrl+Alt+E")));
    check("brak błędów strony (poprawki)", errors.length === 0, errors.join(" | "));
    await context.close();
  }
  {
    // F5: bez pliku polecenia wymagające pliku są wyszarzone, po wczytaniu działają
    const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1440, height: 900 } });
    await context.addInitScript(() => { localStorage.clear(); localStorage.setItem("introPlayed", "true"); });
    const page = await context.newPage();
    await page.goto(APP_URL, { waitUntil: "load" });
    await page.evaluate(() => document.getElementById("heroSplash")?.remove());
    await sleep(400);
    const empty = await page.evaluate(() => Array.from(document.querySelectorAll("#ribbon [data-rb-act]")).map((b) => ({ a: b.dataset.rbAct, d: b.disabled })));
    check("F5. bez pliku: Wklej/Wytnij/Kopiuj/Notatka/Wyjaśnij wyszarzone", empty.length >= 5 && empty.every((x) => x.d), JSON.stringify(empty));
    await page.evaluate(() => document.getElementById("loadSampleBtn")?.click());
    await page.waitForFunction(() => document.querySelector("#dataTable tbody tr[data-row-key]"), null, { timeout: 30000 });
    await sleep(300);
    check("F5. po wczytaniu: polecenia aktywne", await page.evaluate(() => Array.from(document.querySelectorAll("#ribbon [data-rb-act]")).every((b) => !b.disabled)));
    await context.close();
  }

  // ── etap 2: ekran Plik ─────────────────────────────────────────────────────
  {
    const { context, page, errors } = await open(browser, { width: 1440, height: 900 });
    await page.click("#rbFileBtn");
    await sleep(350);
    const st = await page.evaluate(() => {
      const bs = document.getElementById("rbBackstage"), r = bs?.getBoundingClientRect();
      return {
        vis: !!bs && !bs.hidden && r.width > 0, exp: document.getElementById("rbFileBtn").getAttribute("aria-expanded"),
        focusIn: !!bs && bs.contains(document.activeElement), inert: !!document.querySelector(".workspace")?.inert,
        sec: bs?.querySelector('.rb-bs-nav [aria-current="true"]')?.dataset.bsSec,
        tiles: Array.from(bs?.querySelectorAll('[data-bs-sec-body="new"] .rb-tile') || []).map((x) => x.dataset.bsAct),
      };
    });
    check("E2. Plik otwiera ekran Plik (fokus w środku, reszta inert)", st.vis && st.exp === "true" && st.focusIn && st.inert, JSON.stringify(st));
    check("E2. domyślnie sekcja Nowy z kafelkami Otwórz / Pusty / Przykładowy", st.sec === "new" && st.tiles.join() === "open,new,sample", JSON.stringify(st));
    await page.click('#rbBackstage [data-bs-sec="info"]');
    await sleep(150);
    const info = await page.evaluate(() => { const b = document.querySelector('[data-bs-sec-body="info"]'); return { txt: b?.textContent || "", file: currentFileName, sheetOk: !!b && b.textContent.includes(currentSheetName), vis: !!b && !b.hidden }; });
    check("E2. Informacje: nazwa pliku i arkusz", info.vis && info.file && info.txt.includes(info.file) && info.sheetOk, JSON.stringify(info).slice(0, 300));
    await page.click('#rbBackstage [data-bs-sec="opts"]');
    await sleep(150);
    const opts = await page.evaluate(() => ({ theme: !!document.querySelector('#rbBackstage [data-bs-sec-body="opts"] #themeSeg'), lang: !!document.querySelector('#rbBackstage [data-bs-sec-body="opts"] #langSwitch'), backup: !!document.querySelector("#rbBackstage #settingsBackupItem"), inMenu: !!document.querySelector("#appMenu #themeSeg, #appMenu #langSwitch, #appMenu #settingsBackupItem") }));
    check("E2. Opcje: motyw, język, kopia ustawień (już nie w ⋯)", opts.theme && opts.lang && opts.backup && !opts.inMenu, JSON.stringify(opts));
    const pill = await page.evaluate(() => { const i = document.querySelector("#langSwitch .lang-switch-indicator").getBoundingClientRect(), b = document.querySelector("#langSwitch .lang-button.is-active").getBoundingClientRect(); return { w: Math.round(i.width), bw: Math.round(b.width), dx: Math.round(i.left - b.left) }; });
    check("E2. Opcje: znacznik wybranego języka stoi pod aktywnym przyciskiem", pill.w > 10 && Math.abs(pill.w - pill.bw) <= 2 && Math.abs(pill.dx) <= 2, JSON.stringify(pill));
    await page.click('#rbBackstage [data-bs-sec="help"]');
    await sleep(150);
    check("E2. Pomoc: skróty klawiszowe przeniesione z ⋯", await page.evaluate(() => !!document.querySelector('#rbBackstage [data-bs-sec-body="help"] #shortcutsMenuItem') && !document.querySelector("#appMenu #shortcutsMenuItem")));
    check("E2. bez tymczasowego menu Plik z etapu 1", await page.evaluate(() => !document.getElementById("rbFileMenu")));
    await page.keyboard.press("Escape");
    await sleep(350);
    const closed = await page.evaluate(() => ({ hidden: document.getElementById("rbBackstage").hidden, focus: document.activeElement?.id, inert: !!document.querySelector(".workspace")?.inert, exp: document.getElementById("rbFileBtn").getAttribute("aria-expanded") }));
    check("E2. Esc zamyka, fokus na Plik, aplikacja znów aktywna", closed.hidden && closed.focus === "rbFileBtn" && !closed.inert && closed.exp === "false", JSON.stringify(closed));
    // Eksportuj: ten sam przycisk co dawniej, okno eksportu, ekran Plik się zamyka
    await page.click("#rbFileBtn");
    await sleep(300);
    await page.click('#rbBackstage [data-bs-sec="export"]');
    await page.click("#rbBackstage #exportCsvBtn");
    await sleep(400);
    check("E2. Eksport… z ekranu Plik otwiera okno eksportu i zamyka ekran", await page.evaluate(() => !document.getElementById("exportModal").classList.contains("hidden") && document.getElementById("rbBackstage").hidden));
    await page.keyboard.press("Escape");
    await sleep(200);
    // telefon: przeniesione kontrolki wracają do ⋯
    await page.setViewportSize({ width: 768, height: 900 });
    await page.waitForFunction(() => !document.body.classList.contains("rb-on"), null, { timeout: 5000 }).catch(() => {});
    await sleep(150);
    check("E2. 768 px: motyw, język, skróty, eksport wracają na swoje miejsca", await page.evaluate(() => !!document.querySelector("#appMenu #themeSeg") && !!document.querySelector("#appMenu #langSwitch") && !!document.querySelector("#appMenu #shortcutsMenuItem") && !!document.querySelector("#moreMenu #exportCsvBtn")));
    check("brak błędów strony (E2)", errors.length === 0, errors.join(" | "));
    await context.close();
  }

  // ── etap 3: szyna i jeden panel naraz ───────────────────────────────────────
  {
    const { context, page, errors } = await open(browser, { width: 1440, height: 900 });
    const panelsShown = () => page.evaluate(() => Array.from(document.querySelectorAll(".sidebar details.panel")).filter((d) => d.getBoundingClientRect().height > 30).map((d) => d.id));
    const sideOpen = () => page.evaluate(() => document.documentElement.classList.contains("sidebar-open"));
    const rail = await page.evaluate(() => ({
      vis: document.getElementById("sidebarRail").getClientRects().length > 0,
      btns: Array.from(document.querySelectorAll("#sidebarRail button")).filter((b) => b.getClientRects().length > 0).map((b) => b.dataset.rbRail || b.dataset.group || b.id),
      tool: document.getElementById("panelToggle").getClientRects().length > 0,
      finder: document.getElementById("sidebarFinder").getClientRects().length > 0,
    }));
    check("E3. szyna: Filtry, Sortowanie, Formuły, Sprawdzanie (bez starych grup i 🔧)", rail.vis && rail.btns.join() === "filters,sort,formulas,check" && !rail.tool, JSON.stringify(rail));
    await page.click('#sidebarRail [data-rb-rail="filters"]');
    await sleep(400);
    const f = { open: await sideOpen(), shown: await panelsShown(), title: await page.evaluate(() => document.getElementById("rbPaneTitle")?.textContent.trim()), finder: await page.evaluate(() => document.getElementById("sidebarFinder").getClientRects().length > 0), pressed: await page.getAttribute('#sidebarRail [data-rb-rail="filters"]', "aria-pressed") };
    check("E3. Filtry: jeden panel z tytułem i 3 sekcjami, bez szukajki", f.open && f.title === "Filtry" && f.shown.join() === "panel-text-filters,panel-date-filter,panel-smart-filters" && !f.finder && f.pressed === "true", JSON.stringify(f));
    await page.keyboard.press("Escape");
    await sleep(400);
    check("E3. Esc zamyka panel i oddaje fokus przyciskowi szyny", !(await sideOpen()) && (await page.evaluate(() => document.activeElement?.dataset?.rbRail)) === "filters");
    await page.click('#sidebarRail [data-rb-rail="filters"]');
    await sleep(350);
    await page.click('#sidebarRail [data-rb-rail="filters"]');
    await sleep(400);
    check("E3. drugi klik w Filtry zamyka panel", !(await sideOpen()));
    await page.click("#rbTab-an");
    await page.click('#ribbon [data-rb-panel="panel-aggregation-workbench"]');
    await sleep(400);
    const ag = { shown: await panelsShown(), title: await page.evaluate(() => document.getElementById("rbPaneTitle")?.textContent.trim()) };
    check("E3. Analiza → Agregacje: tylko ten panel, tytuł Agregacje", ag.shown.join() === "panel-aggregation-workbench" && ag.title === "Agregacje", JSON.stringify(ag));
    // szerokie Agregacje nie poszerzają potem wąskiego panelu (schowane otwarte sekcje zwinięte)
    await page.click('#sidebarRail [data-rb-rail="check"]');
    await sleep(450);
    const wCheck = await page.evaluate(() => Math.round(document.querySelector(".sidebar").getBoundingClientRect().width));
    check("E3. po Agregacjach panel Sprawdzanie ma zwykłą szerokość", wCheck <= 380, String(wCheck));
    await page.click("#rbTab-an");
    await page.click('#ribbon [data-rb-panel="panel-aggregation-workbench"]');
    await sleep(400);
    await page.click("#sidebarPinBtn");
    await sleep(500);
    const pin = await page.evaluate(() => ({ pinned: document.documentElement.classList.contains("panel-pinned"), shown: Array.from(document.querySelectorAll(".sidebar details.panel")).filter((d) => d.getClientRects().length > 0).map((d) => d.id), mainLeft: Math.round(document.querySelector(".main").getBoundingClientRect().left), sideRight: Math.round(document.querySelector(".sidebar").getBoundingClientRect().right) }));
    check("E3. pinezka: panel obok tabeli, dalej jeden panel", pin.pinned && pin.shown.join() === "panel-aggregation-workbench" && pin.mainLeft >= pin.sideRight, JSON.stringify(pin));
    await page.click("#sidebarPinBtn");
    await sleep(500);
    // każdy panel ma dokładnie jeden dom; żaden przycisk karty nie dubluje szyny
    const homes = await page.evaluate(() => {
      const railPanels = { filters: ["panel-text-filters", "panel-date-filter", "panel-smart-filters"], sort: ["panel-sort-workbench"], formulas: ["panel-formula-workbench"], check: ["panel-validation"] };
      const onRail = new Set(Object.values(railPanels).flat());
      const btns = Array.from(document.querySelectorAll("#ribbon [data-rb-panel]")).map((b) => b.dataset.rbPanel);
      const all = Array.from(document.querySelectorAll(".sidebar details.panel")).map((d) => d.id);
      const count = (id) => (onRail.has(id) ? 1 : 0) + btns.filter((x) => x === id).length;
      return { none: all.filter((id) => count(id) === 0), many: all.filter((id) => count(id) > 1) };
    });
    check("E3. każdy panel ma dokładnie jeden dom (szyna albo przycisk)", !homes.none.length && !homes.many.length, JSON.stringify(homes));
    // kod aplikacji otwierający konkretną sekcję (np. z menu komórki) — panel pokazuje właśnie ją
    await page.evaluate(() => { setSidebarOpen(true); document.getElementById("panel-edit-tools").open = true; });
    await sleep(400);
    const codeOpen = await page.evaluate(() => ({ title: document.getElementById("rbPaneTitle").textContent.trim(), editVis: document.getElementById("panel-edit-tools").getBoundingClientRect().height > 30 }));
    check("E3. otwarcie sekcji z kodu pokazuje tę sekcję (nie Filtry)", codeOpen.editVis && codeOpen.title !== "Filtry", JSON.stringify(codeOpen));
    await page.evaluate(() => setSidebarOpen(false));
    await sleep(400);
    check("E3. Narzędzia główne: Narzędzia edycji", await page.evaluate(() => !!document.querySelector('#rbPanel-home [data-rb-panel="panel-edit-tools"]')));
    // 900 px: szyna dalej jest, tabela jej nie zasłania
    await page.setViewportSize({ width: 900, height: 900 });
    await sleep(400);
    const narrow = await page.evaluate(() => { const r = document.getElementById("sidebarRail").getBoundingClientRect(), m = document.querySelector(".table-panel").getBoundingClientRect(); return { vis: r.width > 0, railR: Math.round(r.right), mainL: Math.round(m.left), ov: document.documentElement.scrollWidth - innerWidth }; });
    check("E3. 900 px: szyna widoczna, tabela obok niej, nic nie wystaje", narrow.vis && narrow.mainL >= narrow.railR && narrow.ov <= 0, JSON.stringify(narrow));
    await page.click('#sidebarRail [data-rb-rail="sort"]');
    await sleep(400);
    check("E3. 900 px: Sortowanie otwiera jeden panel", (await sideOpen()) && (await panelsShown()).join() === "panel-sort-workbench");
    await page.keyboard.press("Escape");
    await sleep(300);
    await page.setViewportSize({ width: 768, height: 900 });
    await page.waitForFunction(() => !document.body.classList.contains("rb-on"), null, { timeout: 5000 }).catch(() => {});
    await sleep(200);
    const phone = await page.evaluate(() => ({ rbBtns: Array.from(document.querySelectorAll("#sidebarRail .rb-rail-btn")).filter((b) => b.getClientRects().length > 0).length, pane: document.documentElement.classList.contains("rb-pane") }));
    check("E3. 768 px: bez szyny wstążki i bez trybu jednego panelu", phone.rbBtns === 0 && !phone.pane, JSON.stringify(phone));
    check("brak błędów strony (E3)", errors.length === 0, errors.join(" | "));
    await context.close();
  }

  {
    // świeża strona przy 900×900 (tablet w pionie = panel jako szuflada): panel z szyny ma treść
    const { context, page, errors } = await open(browser, { width: 900, height: 900 });
    await page.click('#sidebarRail [data-rb-rail="sort"]');
    await sleep(450);
    const h = await page.evaluate(() => ({ open: document.documentElement.classList.contains("sidebar-open"), h: Math.round(document.getElementById("panel-sort-workbench").getBoundingClientRect().height) }));
    check("E3. 900×900 (szuflada): Sortowanie otwiera panel z treścią", h.open && h.h > 80, JSON.stringify(h));
    check("brak błędów strony (E3 szuflada)", errors.length === 0, errors.join(" | "));
    await context.close();
  }

  // ── etap 6a: pokazywanie wstążki, gęstość wstążki i wierszy ──────────────────
  {
    const { context, page, errors } = await open(browser, { width: 1440, height: 900 });
    const body = () => page.evaluate(() => document.body.className);
    const r2h = () => page.evaluate(() => { const r = document.querySelector(".rb-r2"); return r.getClientRects().length ? Math.round(r.getBoundingClientRect().height) : 0; });
    const rowH = () => page.evaluate(() => Math.round(document.querySelector("#dataTable tbody tr[data-row-key]:nth-child(6)").getBoundingClientRect().height));
    const wrapTop = () => page.evaluate(() => Math.round(document.getElementById("tableWrap").getBoundingClientRect().top));

    // domyślnie na komputerze z myszką: Zawsze, Mysz, wiersze Automatyczna = Standardowa (23 px)
    check("O6. mysz: domyślnie Zawsze + Mysz", /rb-show-pin/.test(await body()) && !/rb-touch/.test(await body()), await body());
    check("O6. mysz: wiersze automatycznie 23 px (Standardowa)", (await rowH()) === 23, await rowH());

    // menu opcji na końcu rzędu poleceń: 3 sposoby pokazywania, zaznaczone Zawsze
    await page.click("#rbOptsBtn");
    await sleep(200);
    const menu = await page.evaluate(() => Array.from(document.querySelectorAll("#rbOptsMenu [data-rb-show]")).map((b) => `${b.dataset.rbShow}:${b.getAttribute("aria-checked")}`));
    check("O6. menu opcji: Zawsze / Chowaj przy przewijaniu / Tylko karty", JSON.stringify(menu) === JSON.stringify(["pin:true", "auto:false", "tabs:false"]), menu.join());

    // Tylko karty: rząd poleceń znika, tabela idzie w górę; klik karty wysuwa polecenia, drugi klik chowa
    const top0 = await wrapTop();
    await page.click('#rbOptsMenu [data-rb-show="tabs"]');
    await sleep(400);
    check("O6. Tylko karty: bez rzędu poleceń", (await r2h()) === 0, await r2h());
    const top1 = await wrapTop();
    check("O6. Tylko karty: tabela wyżej o rząd poleceń", top0 - top1 >= 30, `${top0} → ${top1}`);
    await page.click("#rbTab-data");
    await sleep(250);
    const peek = await page.evaluate(() => ({ h: Math.round(document.querySelector(".rb-r2").getBoundingClientRect().height), data: !document.getElementById("rbPanel-data").hidden, top: Math.round(document.getElementById("tableWrap").getBoundingClientRect().top) }));
    check("O6. Tylko karty: klik karty wysuwa jej polecenia nad tabelą", peek.h > 30 && peek.data && Math.abs(peek.top - top1) <= 1, JSON.stringify(peek));
    await page.click("#rbTab-data");
    await sleep(250);
    check("O6. ta sama karta drugi raz chowa polecenia", (await r2h()) === 0);
    await page.click("#rbTab-home");
    await sleep(200);
    await page.keyboard.press("Escape");
    await sleep(200);
    check("O6. Esc chowa wysunięte polecenia", (await r2h()) === 0);

    // Ctrl+F1 przełącza Tylko karty ↔ Zawsze; wybór zostaje po przeładowaniu
    await page.keyboard.press("Control+F1");
    await sleep(300);
    check("O6. Ctrl+F1: z powrotem Zawsze", /rb-show-pin/.test(await body()) && (await r2h()) > 30);
    await page.keyboard.press("Control+F1");
    await sleep(300);
    await page.reload({ waitUntil: "load" });
    await page.evaluate(() => document.getElementById("heroSplash")?.remove());
    await page.evaluate(() => document.getElementById("loadSampleBtn")?.click());
    await page.waitForFunction(() => document.querySelector("#dataTable tbody tr[data-row-key]"), null, { timeout: 30000 });
    await sleep(400);
    check("O6. Tylko karty zapamiętane po przeładowaniu", /rb-show-tabs/.test(await body()), await body());
    await page.keyboard.press("Control+F1");
    await sleep(300);

    // Chowaj przy przewijaniu: w dół rząd poleceń się chowa, w górę wraca
    await page.click("#rbOptsBtn");
    await page.click('#rbOptsMenu [data-rb-show="auto"]');
    await sleep(200);
    await page.evaluate(() => { const w = document.getElementById("tableWrap"); w.scrollTop = 200; w.dispatchEvent(new Event("scroll")); });
    await sleep(80);
    await page.evaluate(() => { const w = document.getElementById("tableWrap"); w.scrollTop = 600; w.dispatchEvent(new Event("scroll")); });
    await sleep(450);
    check("O6. Chowaj przy przewijaniu: w dół rząd znika", /rb-tucked/.test(await body()) && (await r2h()) <= 1, `${await body()} r2=${await r2h()}`);
    const gap = await page.evaluate(() => { const w = document.getElementById("tableWrap").getBoundingClientRect(); const f = document.getElementById("tableFoot").getBoundingClientRect(); return Math.round(f.top - w.bottom); });
    check("O6. po schowaniu tabela rośnie (bez dziury nad dolnym paskiem)", gap <= 24, gap + " px");
    await page.evaluate(() => { const w = document.getElementById("tableWrap"); w.scrollTop = 400; w.dispatchEvent(new Event("scroll")); });
    await sleep(450);
    check("O6. w górę rząd wraca", !/rb-tucked/.test(await body()) && (await r2h()) > 30);

    // Plik → Opcje: gęstość wstążki Dotyk (36 px) i gęstość wierszy
    await page.click("#rbFileBtn");
    await page.click('[data-bs-sec="opts"]');
    await sleep(300);
    check("O6. Plik → Opcje: pokazywanie, gęstość wstążki, gęstość wierszy", await page.evaluate(() => !!document.querySelector('#rbShowSeg [data-rb-show="auto"][aria-checked="true"]') && !!document.querySelector("#rbDensSeg [data-rb-dens]") && document.querySelectorAll("#rowDensSeg [data-density]").length === 4));
    await page.click('#rbDensSeg [data-rb-dens="touch"]');
    await page.keyboard.press("Escape");
    await sleep(400);
    // gęstość wstążki zmienia wysokość rzędu poleceń — tabela ma się przeliczyć sama (przegląd kodu)
    const gapT = await page.evaluate(() => Math.round(document.getElementById("tableFoot").getBoundingClientRect().top - document.getElementById("tableWrap").getBoundingClientRect().bottom));
    check("O6. Dotyk: tabela przeliczona (bez nachodzenia na dolny pasek)", gapT >= 0 && gapT <= 24, gapT + " px");
    await page.click("#rbFileBtn");
    await page.click('[data-bs-sec="opts"]');
    await sleep(300);
    await page.click('#rowDensSeg [data-density="comfortable"]');
    await page.keyboard.press("Escape");
    await sleep(400);
    const touchBtn = await page.evaluate(() => Math.round(document.querySelector("#rbPanel-home .rb-btn").getBoundingClientRect().height));
    check("O6. Dotyk: przyciski poleceń 36 px", touchBtn === 36, touchBtn);
    check("O6. Wygodna: wiersze 29 px", (await rowH()) === 29, await rowH());
    {
      const cell = (r, c) => page.locator(`#dataTable tbody tr[data-row-key]:nth-child(${r}) td[data-col-index="${c}"]`).boundingBox();
      const a = await cell(8, 1), z = await cell(10, 2);
      await page.mouse.move(a.x + 5, a.y + 5); await page.mouse.down(); await page.mouse.move(z.x + 5, z.y + 5, { steps: 6 }); await page.mouse.up();
      await sleep(300);
      check("O6. Dotyk: bez mini paska", await page.evaluate(() => { const m = document.getElementById("miniBar"); return !m || m.hidden; }));
      await page.keyboard.press("Escape");
    }
    await page.evaluate(() => document.querySelector('#zoomNowBtn')?.click());
    await sleep(200);
    const zoomOpts = await page.evaluate(() => Array.from(document.querySelectorAll("#zoomMenu .density-opt")).length);
    check("O6. menu powiększenia: 4 gęstości (z Automatyczną)", zoomOpts === 4, zoomOpts);
    await page.keyboard.press("Escape");
    check("O6. panel Widok bez drugiego wyboru gęstości", !(await shown(page, "#rowDensity")));
    check("brak błędów strony (O6)", errors.length === 0, errors.join(" | "));
    await context.close();
  }
  {
    // ogranicz ruch: chowanie bez przejścia też przelicza tabelę; Tylko karty z klawiatury — fokus na kartę
    const { context, page, errors } = await open(browser, { width: 1440, height: 900 }, { reducedMotion: "reduce" });
    await page.evaluate(() => swbRibbon.setShow("auto"));
    await sleep(200);
    for (const y of [200, 600]) { await page.evaluate((v) => { const w = document.getElementById("tableWrap"); w.scrollTop = v; w.dispatchEvent(new Event("scroll")); }, y); await sleep(120); }
    await sleep(300);
    const g = await page.evaluate(() => ({ tucked: document.body.classList.contains("rb-tucked"), gap: Math.round(document.getElementById("tableFoot").getBoundingClientRect().top - document.getElementById("tableWrap").getBoundingClientRect().bottom) }));
    check("O6. ogranicz ruch: po schowaniu tabela rośnie", g.tucked && g.gap <= 24, JSON.stringify(g));
    await page.evaluate(() => swbRibbon.setShow("pin"));
    await sleep(200);
    await page.focus("#rbOptsBtn");
    await page.keyboard.press("Enter");
    await sleep(150);
    await page.keyboard.press("ArrowDown");
    await page.keyboard.press("ArrowDown");
    await page.keyboard.press("Enter");
    await sleep(250);
    const f = await page.evaluate(() => ({ tabs: document.body.classList.contains("rb-show-tabs"), focus: document.activeElement?.classList.contains("rb-tab") }));
    check("O6. Tylko karty z klawiatury: fokus na karcie, nie w próżni", f.tabs && f.focus, JSON.stringify(f));
    check("brak błędów strony (O6 ogranicz ruch)", errors.length === 0, errors.join(" | "));
    await context.close();
  }
  {
    // iPad pionowo (sam dotyk): Chowaj przy przewijaniu + Dotyk + wiersze Wygodna; poziomo: Zawsze
    const { context, page, errors } = await open(browser, { width: 820, height: 1180 }, { hasTouch: true });
    const cls = await page.evaluate(() => document.body.className);
    check("O6. iPad pionowo: domyślnie Chowaj przy przewijaniu + Dotyk", /rb-show-auto/.test(cls) && /rb-touch/.test(cls), cls);
    const rh = await page.evaluate(() => Math.round(document.querySelector("#dataTable tbody tr[data-row-key]:nth-child(6)").getBoundingClientRect().height));
    check("O6. iPad: wiersze automatycznie 29 px (Wygodna)", rh === 29, rh);
    await page.setViewportSize({ width: 1180, height: 820 });
    await sleep(400);
    const cls2 = await page.evaluate(() => document.body.className);
    check("O6. iPad poziomo: Zawsze", /rb-show-pin/.test(cls2) && /rb-touch/.test(cls2), cls2);
    check("brak błędów strony (O6 iPad)", errors.length === 0, errors.join(" | "));
    await context.close();
  }
  {
    // iPad poziomo bez paska z boku: dolny pasek w jednej linii
    const { context, page, errors } = await open(browser, { width: 1024, height: 768 });
    const foot = await page.evaluate(() => Math.round(document.getElementById("tableFoot").getBoundingClientRect().height));
    check("O6. 1024 px: dolny pasek w jednej linii (≤ 44 px)", foot <= 44, foot + " px");
    check("brak błędów strony (O6 1024)", errors.length === 0, errors.join(" | "));
    await context.close();
  }

  // ── etap 5: mini pasek nad zaznaczeniem myszą ─────────────────────────────
  {
    const { context, page, errors } = await open(browser, { width: 1440, height: 900 });
    const td = (r, c) => `#dataTable tbody tr[data-row-key]:nth-child(${r}) td[data-col-index="${c}"]`;
    const center = async (sel) => { const b = await page.locator(sel).first().boundingBox(); return { x: b.x + b.width / 2, y: b.y + b.height / 2 }; };
    async function drag(from, to) {
      const a = await center(from), b = await center(to);
      await page.mouse.move(a.x, a.y);
      await page.mouse.down();
      await page.mouse.move((a.x + b.x) / 2, (a.y + b.y) / 2, { steps: 4 });
      await page.mouse.move(b.x, b.y, { steps: 4 });
      await page.mouse.up();
      await sleep(350);
    }
    const bar = () => page.evaluate(() => {
      const m = document.getElementById("miniBar");
      if (!m || m.hidden || m.classList.contains("is-off")) return null;
      const r = m.getBoundingClientRect();
      return { top: r.top, bottom: r.bottom, left: r.left, right: r.right, n: m.querySelectorAll("button, select").length };
    });
    const selBox = (a, b) => page.evaluate(([s1, s2]) => {
      const r1 = document.querySelector(s1).getBoundingClientRect(), r2 = document.querySelector(s2).getBoundingClientRect();
      return { top: Math.min(r1.top, r2.top), bottom: Math.max(r1.bottom, r2.bottom) };
    }, [a, b]);

    // zwykłe kliknięcie jednej komórki — bez paska (decyzja Mateusza)
    await page.click(td(8, 1));
    await sleep(300);
    check("M5. klik jednej komórki: bez mini paska", !(await bar()));

    await drag(td(8, 1), td(10, 2));
    const b1 = await bar();
    const s1 = await selBox(td(8, 1), td(10, 2));
    check("M5. zaznaczenie myszą: mini pasek widoczny", !!b1, JSON.stringify(b1));
    check("M5. pasek nad zaznaczeniem", b1 && b1.bottom <= s1.top + 1, JSON.stringify({ b1, s1 }));
    check("M5. 7 poleceń (B, wypełnienie, format, Σ, Filtruj, Notatka, Wyjaśnij)", b1 && b1.n === 7, b1 && b1.n);
    const labels = await page.evaluate(() => Array.from(document.querySelectorAll("#miniBar button, #miniBar select")).map((b) => b.getAttribute("aria-label") || b.textContent.trim()));
    check("M5. każde polecenie ma nazwę", labels.length === 7 && labels.every(Boolean), labels.join(" | "));

    // przyciski wołają TE SAME polecenia co wstążka (szpieg na oryginałach)
    await page.evaluate(() => {
      window.__mini = [];
      const spy = (sel, name) => document.querySelector(sel)?.addEventListener("click", () => window.__mini.push(name), true);
      spy('#formatRibbon [data-fmt="b"]', "bold");
      spy("#fmtFillBtn", "fill");
      spy("#fmtSumBtn", "sum");
      spy('#ribbon [data-rb-act="note"]', "note");
      spy('#ribbon [data-rb-act="explain"]', "explain");
      document.getElementById("fmtNumFormat").addEventListener("change", (e) => window.__mini.push("num:" + e.target.value));
      // skutków nie sprawdzamy tutaj (mają własne testy) — tylko że idą tą samą drogą
      for (const b of document.querySelectorAll('#formatRibbon [data-fmt="b"], #fmtFillBtn, #fmtSumBtn, #ribbon [data-rb-act="note"], #ribbon [data-rb-act="explain"]')) b.addEventListener("click", (e) => e.stopImmediatePropagation(), true);
    });
    // Σ, Notatka i Wyjaśnij chowają pasek (zmieniają widok) — przed każdym zaznaczamy od nowa
    for (const act of ["bold", "fill", "sum", "note", "explain"]) {
      if (!(await bar())) await drag(td(8, 1), td(10, 2));
      await page.click(`#miniBar [data-mini="${act}"]`);
      await sleep(80);
    }
    check("M5. po Wyjaśnij pasek znika", !(await bar()));
    await drag(td(8, 1), td(10, 2));
    const numVal = await page.evaluate(() => { const s = document.querySelector('#miniBar select[data-mini="num"]'); const o = Array.from(s.options).find((x) => x.value && x.value !== s.value); return o ? o.value : ""; });
    await page.selectOption('#miniBar select[data-mini="num"]', numVal);
    await sleep(80);
    const got = await page.evaluate(() => window.__mini);
    check("M5. B / wypełnienie / Σ / Notatka / Wyjaśnij / format = te same polecenia co wstążka",
      JSON.stringify(got) === JSON.stringify(["bold", "fill", "sum", "note", "explain", "num:" + numVal]), JSON.stringify(got));
    await context.close();
  }
  {
    const { context, page, errors } = await open(browser, { width: 1440, height: 900 });
    const td = (r, c) => `#dataTable tbody tr[data-row-key]:nth-child(${r}) td[data-col-index="${c}"]`;
    const center = async (sel) => { const b = await page.locator(sel).first().boundingBox(); return { x: b.x + b.width / 2, y: b.y + b.height / 2 }; };
    async function drag(from, to) {
      const a = await center(from), b = await center(to);
      await page.mouse.move(a.x, a.y);
      await page.mouse.down();
      await page.mouse.move(b.x, b.y, { steps: 6 });
      await page.mouse.up();
      await sleep(350);
    }
    const shownBar = () => page.evaluate(() => { const m = document.getElementById("miniBar"); return !!m && !m.hidden && !m.classList.contains("is-off"); });

    // Filtruj wg wartości: jak „Pokaż tylko …” z menu komórki — mniej wierszy, pasek znika
    const before = await page.evaluate(() => document.querySelectorAll("#dataTable tbody tr[data-row-key]").length);
    await drag(td(3, 1), td(4, 1));
    await page.click('#miniBar [data-mini="filter"]');
    await sleep(800);
    const fq = await page.evaluate(() => ({ q: document.getElementById("searchQuery")?.value || "", rows: document.querySelectorAll("#dataTable tbody tr[data-row-key]").length }));
    check("M5. Filtruj wg wartości: zapytanie i mniej wierszy", fq.q.trim() && fq.rows < before, JSON.stringify({ ...fq, before }));
    check("M5. po Filtruj pasek znika", !(await shownBar()));
    await page.evaluate(() => document.getElementById("resetFiltersBtn")?.click());
    await sleep(600);

    // brak miejsca nad zaznaczeniem (pierwszy wiersz pod nagłówkiem) — pasek pod nim
    await page.evaluate(() => { const w = document.getElementById("tableWrap"); if (w) w.scrollTop = 0; });
    await sleep(200);
    await drag(td(1, 1), td(2, 2));
    const low = await page.evaluate(() => {
      const m = document.getElementById("miniBar").getBoundingClientRect();
      const c = document.querySelector('#dataTable tbody tr[data-row-key]:nth-child(2) td[data-col-index="2"]').getBoundingClientRect();
      return { barTop: m.top, selBottom: c.bottom, ok: !document.getElementById("miniBar").hidden };
    });
    check("M5. brak miejsca nad zaznaczeniem: pasek pod nim", low.ok && low.barTop >= low.selBottom - 1, JSON.stringify(low));

    // Esc chowa pasek
    await page.keyboard.press("Escape");
    await sleep(200);
    check("M5. Esc chowa pasek", !(await shownBar()));

    // klawiatura (Shift+strzałki) — bez paska
    await page.click(td(6, 1));
    await page.keyboard.press("Shift+ArrowDown");
    await page.keyboard.press("Shift+ArrowRight");
    await sleep(300);
    check("M5. zaznaczenie klawiaturą: bez paska", !(await shownBar()));

    // Podgląd — bez paska
    await page.click("#rbModeBtn");
    await page.click('#rbModeMenu [data-rb-mode="view"]');
    await sleep(400);
    await drag(td(8, 1), td(10, 2));
    check("M5. Podgląd: bez paska", !(await shownBar()));
    check("brak błędów strony (M5)", errors.length === 0, errors.join(" | "));
    await context.close();
  }

  // ── telefon ─────────────────────────────────────────────────────────────
  {
    const { context, page, errors } = await open(browser, { width: 390, height: 844 }, { hasTouch: true, isMobile: ENGINE === "chromium" });
    check("8. telefon: bez rb-on", !(await page.evaluate(() => document.body.classList.contains("rb-on"))));
    check("8. telefon: wstążka ukryta", !(await shown(page, "#ribbon")));
    check("8. telefon: Szukaj na swoim miejscu", !(await inRibbon(page, "#quickSearch")));
    check("brak błędów strony (telefon)", errors.length === 0, errors.join(" | "));
    await context.close();
  }

  await browser.close();
  const failed = results.filter((r) => !r.ok);
  results.forEach((r) => console.log(`${r.ok ? "✓" : "✗"} ${r.name}${r.ok ? "" : "  — " + r.detail}`));
  if (failed.length) throw new Error(`${failed.length} z ${results.length} asercji nie przeszło (${ENGINE})`);
  console.log(`\n✅ ribbon-playwright OK — ${results.length} asercji (${ENGINE})`);
}

run().catch((e) => { console.error(e.message || e); process.exit(1); });
