// phone-top-playwright.js — wstążka, etap 7a: góra ekranu na telefonie (docs/WSTAZKA.md, makieta
// docs/makiety/wstazka-v2.html, tryb „Telefon”; plan zaakceptowany przez Mateusza 2026-11-20).
//
//   1. z plikiem: pasek u góry ← · Cofnij · nazwa pliku i arkusz · Tryb · ⋯; duża karta „hero” schowana,
//      tabela wyżej niż dotąd (dawniej 188 px od góry),
//   2. cele dotykowe w pasku 40 px (Tryb 36), w każdym rzędzie jedna oś i jedna wysokość; pole Szukaj ma ≥ 16 px tekstu (iPhone nie powiększa strony),
//   3. Zapisz w pasku tylko przy niezapisanych zmianach (z licznikiem),
//   4. pasek formuły tylko podczas edycji komórki — wtedy w miejscu Szukaj; po Enter wraca Szukaj,
//   5. Tryb otwiera menu (Edycja / Podgląd / …), wybór Podglądu działa,
//   6. karty arkuszy pod tabelą (jak na komputerze), bez pływającego przycisku arkusza,
//   7. ← przy niezapisanych zmianach pyta „Zapisz / Nie zapisuj / Anuluj”; Anuluj zostawia plik,
//   8. bez pliku i na komputerze nic się nie zmienia (pasek telefonu ukryty).
//
// Uruchom z serwerem na APP_URL (domyślnie http://127.0.0.1:4175/). ENGINE=webkit — Safari.

const playwright = require("playwright");

const APP_URL = process.env.APP_URL || "http://127.0.0.1:4175/";
const ENGINE = process.env.ENGINE || "chromium";
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

async function open(browser, viewport, extra = {}, { load = true } = {}) {
  const context = await browser.newContext({ serviceWorkers: "block", viewport, ...extra });
  await context.addInitScript(() => {
    if (!sessionStorage.getItem("__pt")) { sessionStorage.setItem("__pt", "1"); localStorage.clear(); localStorage.setItem("introPlayed", "true"); }
  });
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push("pageerror: " + e.message));
  await page.goto(APP_URL + "?sample=300", { waitUntil: "load" });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  if (load) {
    await page.evaluate(() => document.getElementById("loadSampleBtn")?.click());
    await page.waitForFunction(() => document.querySelector("#dataTable tbody tr[data-row-key]"), null, { timeout: 30000 });
  }
  await sleep(600);
  await page.evaluate(() => document.querySelectorAll(".toast").forEach((t) => t.remove()));
  return { context, page, errors };
}

const shown = (page, sel) => page.evaluate((s) => { const el = document.querySelector(s); if (!el) return false; const r = el.getBoundingClientRect(); return r.width > 0 && r.height > 0 && getComputedStyle(el).visibility !== "hidden"; }, sel);
const inTop = (page, sel) => page.evaluate((s) => !!document.querySelector(s)?.closest("#phTop"), sel);

async function run() {
  const browser = await playwright[ENGINE].launch({ headless: true });
  const results = [];
  const check = (name, ok, detail = "") => results.push({ name, ok: !!ok, detail: typeof detail === "string" ? detail : JSON.stringify(detail) });
  const phone = { width: 390, height: 844 };
  const touch = { hasTouch: true, isMobile: ENGINE === "chromium" };

  // ── telefon z plikiem ──
  {
    const { context, page, errors } = await open(browser, phone, touch);
    check("1. pasek telefonu widoczny", await shown(page, "#phTop"));
    check("1. duża karta „hero” schowana", !(await shown(page, ".hero")));
    const order = await page.evaluate(() => {
      const top = document.getElementById("phTop");
      const items = ["#phBackBtn", "#undoBtn", "#phName", "#rbModeBtn", "#appMenuBtn"].map((s) => document.querySelector(s));
      return items.map((el) => (el && top.contains(el) ? Math.round(el.getBoundingClientRect().left) : -1));
    });
    check("1. kolejność: ← · Cofnij · nazwa · Tryb · ⋯", order.every((x) => x >= 0) && order.every((x, i) => i === 0 || x > order[i - 1]), order);
    const name = await page.evaluate(() => ({ file: document.getElementById("phFile")?.textContent, sheet: document.getElementById("phSheetName")?.textContent }));
    check("1. nazwa pliku i arkusza w pasku", /Przyklad|xlsx/.test(name.file || "") && (name.sheet || "").length > 0, name);
    const tableTop = await page.evaluate(() => Math.round(document.getElementById("tableWrap").getBoundingClientRect().top));
    check("1. tabela wyżej niż dotąd (< 160 px od góry)", tableTop < 160, tableTop);

    // przeniesione z hero-collapse-guard (karta „hero” z uchwytem zwijania zastąpiona paskiem w etapie 7a)
    const fit = await page.evaluate(() => ({
      pageScroll: document.documentElement.scrollHeight - innerHeight,
      footBottom: Math.round(document.getElementById("tableFoot").getBoundingClientRect().bottom),
      vh: innerHeight,
    }));
    check("1. strona nie przewija się w pionie (wszystko na ekranie)", fit.pageScroll <= 1, fit);
    check("1. dolny pasek na ekranie (tabela kończy się nad nim)", fit.footBottom <= fit.vh && fit.footBottom > fit.vh - 120, fit);
    const sizes = await page.evaluate(() => ["#phBackBtn", "#undoBtn", "#rbModeBtn", "#appMenuBtn"].map((s) => Math.round(document.querySelector(s)?.getBoundingClientRect().height || 0)));
    check("2. przyciski paska ≥ 40 px (Tryb — pigułka 36 px jak w makiecie)", sizes.every((h, i) => h >= (i === 2 ? 36 : 40)), sizes);
    // jedna oś i jedna wysokość w każdym rzędzie (uwaga Mateusza 2026-11-20: „za duże i niewyrównane”)
    const rows = await page.evaluate(() => {
      const kids = (sel) => [...document.querySelector(sel).children].filter((el) => { const r = el.getBoundingClientRect(); return r.width > 0 && r.height > 0 && getComputedStyle(el).position !== "absolute"; });
      const box = (el) => { const r = el.getBoundingClientRect(); return { id: el.id || el.className, h: Math.round(r.height), mid: Math.round(r.top + r.height / 2) }; };
      return {
        top: [...document.querySelectorAll("#phBackBtn, #undoBtn, #rbModeBtn, #appMenuBtn")].map(box),
        search: kids(".table-toolbar").filter((el) => !el.matches(".table-actions, .recent-strip")).map(box),
        foot: kids("#tableFoot").map(box),
      };
    });
    const oneAxis = (list) => list.length > 1 && Math.max(...list.map((x) => x.mid)) - Math.min(...list.map((x) => x.mid)) <= 1;
    const oneHeight = (list) => list.length > 1 && new Set(list.map((x) => x.h)).size === 1;
    check("2. góra: jedna oś, przyciski 40 px (makieta)", oneAxis(rows.top) && rows.top.every((x) => x.h === 40 || x.id === "rbModeBtn"), rows.top);
    check("2. rząd Szukaj: jedna oś i jedna wysokość", oneAxis(rows.search) && oneHeight(rows.search), rows.search);
    check("2. dolny pasek: jedna oś i jedna wysokość, 40 px (łatwo złapać palcem, iPhone Mateusza)", oneAxis(rows.foot) && oneHeight(rows.foot) && rows.foot[0].h === 40, rows.foot);
    const qsFont = await page.evaluate(() => parseFloat(getComputedStyle(document.getElementById("quickSearch")).fontSize));
    check("2. Szukaj ≥ 16 px (bez powiększania strony na iPhonie)", qsFont >= 16, qsFont);

    check("3. bez zmian: Zapisz schowany", !(await shown(page, "#heroSaveBtn")));
    check("4. przy przeglądaniu: Szukaj widoczne, paska formuły brak", (await shown(page, "#quickSearch")) && !(await shown(page, "#formulaInput")));
    // Szukaj + schowanie klawiatury „✓” (nagranie Mateusza 2026-11-20, SwiftKey): lista „Ostatnie” i poszerzone
    // pole nie mogą zostać nad tabelą, nawet gdy przeglądarka nie zgłosi utraty fokusu — udajemy klawiaturę
    await page.evaluate(() => { try { localStorage.setItem("excel-workbench-qs-history", JSON.stringify({ pinned: [], recent: [{ q: "15.11.26" }, { q: "@notatka" }] })); } catch (_) {} });
    await page.tap("#quickSearch");
    await sleep(400);
    const kb = (h) => page.evaluate((hh) => { Object.defineProperty(visualViewport, "height", { configurable: true, get: () => hh }); visualViewport.dispatchEvent(new Event("resize")); }, h);
    await kb(844 - 320);
    await sleep(100);
    await kb(844);
    await sleep(500);
    const qsAfter = await page.evaluate(() => ({ live: !document.getElementById("qsLiveResults").classList.contains("hidden"), wide: document.querySelector(".table-toolbar").classList.contains("qs-wide"), focus: document.activeElement?.id || "" }));
    check("4b. klawiatura schowana: lista „Ostatnie” i poszerzenie Szukaj zwinięte", !qsAfter.live && !qsAfter.wide && qsAfter.focus !== "quickSearch", qsAfter);
    await page.evaluate(() => { delete visualViewport.height; });
    // „suwaki” w polu Szukaj: opcje (w tym „Operatory wyszukiwania”) muszą się pokazać — w 7a stała wysokość
    // pola je ucinała (Mateusz 2026-11-20: „operator wyszukiwania nie otwiera się”)
    await page.tap("#qsMoreBtn");
    await sleep(400);
    const hit = await page.evaluate(() => ["#quickSearchWrap .qs-operators-toggle", "#qsAllSheets", "#quickSearchColumnsBtn"].map((s) => {
      const el = document.querySelector(s); if (!el) return s + ": brak";
      const r = el.getBoundingClientRect(); if (!(r.width > 0 && r.height > 0)) return s + ": 0";
      const at = document.elementFromPoint(r.left + r.width / 2, r.top + r.height / 2);
      return el.contains(at) ? "ok" : `${s}: zasłania ${at?.id || at?.className}`;
    }));
    check("4c. suwaki w Szukaj: „Operatory wyszukiwania”, Wszystkie arkusze i Kolumny widoczne i do stuknięcia", hit.every((x) => x === "ok"), hit);
    const filt = await page.evaluate(() => { const a = document.getElementById("quickSearchWrap").getBoundingClientRect(), b = document.getElementById("panelToggle").getBoundingClientRect(); return { wrapTop: Math.round(a.top), filtTop: Math.round(b.top), filtRight: Math.round(b.right), vw: innerWidth }; });
    check("4c. rozwinięte opcje: Filtry zostaje obok pola (nie spada do osobnego rzędu)", Math.abs(filt.wrapTop - filt.filtTop) <= 2 && filt.filtRight <= filt.vw, filt);
    // operatory: przełącznik (role=switch) i zapamiętanie na urządzeniu (Mateusz 2026-11-20: „zapamiętuj”)
    check("4d. „Operatory wyszukiwania” to przełącznik (role=switch)", await page.evaluate(() => document.getElementById("quickSearchOperators").getAttribute("role") === "switch"));
    await page.tap("#quickSearchWrap .qs-operators-toggle");
    await sleep(200);
    await page.evaluate(() => typeof resetFilterInputs === "function" && resetFilterInputs()); // jak „Wyczyść filtry” / nowy plik
    await sleep(200);
    const opsKept = await page.evaluate(() => ({ el: document.getElementById("quickSearchOperators").checked, flag: quickSearchOperatorsEnabled, saved: localStorage.getItem("swb-qs-ops-v1") }));
    check("4d. włączone operatory zapamiętane (po wyczyszczeniu filtrów dalej włączone)", opsKept.el && opsKept.flag && opsKept.saved === "1", opsKept);
    await page.tap("#quickSearchWrap .qs-operators-toggle");
    await sleep(200);
    await page.tap("#qsMoreBtn");
    await sleep(400);

    // edycja komórki: podwójne stuknięcie
    const cell = '#dataTable tbody tr[data-row-key="wide:8"] td[data-col-index="1"]';
    await page.tap(cell);
    await sleep(120);
    await page.tap(cell);
    await page.waitForSelector("#dataTable input.cell-editor", { timeout: 5000 }).catch(() => {});
    await sleep(300);
    check("4. podczas edycji: pasek formuły widoczny", await shown(page, "#formulaInput"));
    check("4. podczas edycji: Szukaj schowane", !(await shown(page, "#quickSearch")));
    await page.keyboard.type("Zmiana");
    await page.keyboard.press("Enter");
    await sleep(400);
    check("4. po Enter: znowu Szukaj, bez paska formuły", (await shown(page, "#quickSearch")) && !(await shown(page, "#formulaInput")));
    check("3. po zmianie: Zapisz w pasku telefonu", (await shown(page, "#heroSaveBtn")) && (await inTop(page, "#heroSaveBtn")));

    // Tryb
    await page.tap("#rbModeBtn");
    await sleep(250);
    check("5. Tryb otwiera menu", await shown(page, "#rbModeMenu"));
    await page.tap('#rbModeMenu [data-rb-mode="view"]');
    await sleep(400);
    check("5. Podgląd włączony", await page.evaluate(() => typeof swbIsEditMode === "function" && !swbIsEditMode()));
    await page.tap("#rbModeBtn");
    await sleep(250);
    await page.tap('#rbModeMenu [data-rb-mode="edit"]');
    await sleep(400);

    // karty arkuszy
    check("6. karty arkuszy pod tabelą", (await shown(page, "#sheetTabs")) && (await page.evaluate(() => !!document.getElementById("sheetTabs").closest("#tableFoot"))));
    check("6. bez pływającego przycisku arkusza", !(await shown(page, "#sheetPickerFab")));

    // ← przy niezapisanych zmianach
    await page.tap("#phBackBtn");
    await sleep(400);
    const dlg = await page.evaluate(() => [...document.querySelectorAll(".app-dialog-box button")].map((b) => b.textContent.trim()));
    check("7. ← pyta Zapisz / Nie zapisuj / Anuluj", dlg.length >= 3 && dlg.some((x) => /Nie zapisuj/.test(x)), dlg);
    const cancel = await page.$$(".app-dialog-box button");
    if (cancel.length) await cancel[cancel.length - 1].tap();
    await sleep(400);
    check("7. Anuluj: plik zostaje otwarty", await page.evaluate(() => !!workbook && !!document.querySelector("#dataTable tbody tr[data-row-key]")));

    check("brak błędów strony (telefon)", errors.length === 0, errors.join(" | "));
    await context.close();
  }

  // ── telefon bez pliku ──
  {
    const { context, page, errors } = await open(browser, phone, touch, { load: false });
    check("8. bez pliku: pasek telefonu ukryty, hero jak dotąd", !(await shown(page, "#phTop")) && (await shown(page, ".hero")));
    check("brak błędów strony (bez pliku)", errors.length === 0, errors.join(" | "));
    await context.close();
  }

  // ── komputer ──
  {
    const { context, page, errors } = await open(browser, { width: 1280, height: 800 });
    check("8. komputer: pasek telefonu ukryty, wstążka jest", !(await shown(page, "#phTop")) && (await shown(page, "#ribbon")));
    check("8. komputer: Tryb i ⋯ we wstążce", await page.evaluate(() => !!document.getElementById("rbModeBtn").closest("#ribbon") && !!document.getElementById("appMenuBtn").closest("#ribbon")));
    check("brak błędów strony (komputer)", errors.length === 0, errors.join(" | "));
    await context.close();
  }

  await browser.close();
  const failed = results.filter((r) => !r.ok);
  results.forEach((r) => console.log(`${r.ok ? "✓" : "✗"} ${r.name}${r.ok ? "" : "  — " + r.detail}`));
  if (failed.length) throw new Error(`${failed.length} z ${results.length} asercji nie przeszło (${ENGINE})`);
  console.log(`\n✅ phone-top-playwright OK — ${results.length} asercji (${ENGINE})`);
}

run().catch((e) => { console.error(e.message || e); process.exit(1); });
