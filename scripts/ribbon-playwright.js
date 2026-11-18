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
    // Arkusze na komputerze: przycisk z nazwą (lista + Nowy / Zmień nazwę / Usuń), BEZ zakładek —
    // decyzja Mateusza sprzed wstążki (app-frame S1); zakładki u dołu dopiero w etapie 4 (makieta v2).
    const sh = await page.evaluate(() => {
      const b = document.getElementById("heroSheetBtn"), st = document.getElementById("sheetTabs");
      return { inRow: !!b.closest("#rbFbarRow"), vis: b.getBoundingClientRect().width > 0, tabs: !!st.offsetParent };
    });
    check("T3. przycisk arkusza w rzędzie paska formuły, bez zakładek", sh.inRow && sh.vis && !sh.tabs, JSON.stringify(sh));
    await page.click("#heroSheetBtn");
    await page.waitForTimeout(200);
    check("T3. przycisk arkusza otwiera listę arkuszy", await page.evaluate(() => !document.getElementById("sheetMenu").hidden && document.querySelectorAll("#sheetMenu .sheet-menu-item").length > 1));
    await page.keyboard.press("Escape");
    await page.setViewportSize({ width: 768, height: 900 });
    await page.waitForFunction(() => !document.body.classList.contains("rb-on"), null, { timeout: 5000 }).catch(() => {});
    await page.waitForTimeout(100);
    check("T3. 768 px: przycisk arkusza wraca do nagłówka", await page.evaluate(() => !!document.getElementById("heroSheetBtn").closest(".hero")));
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
