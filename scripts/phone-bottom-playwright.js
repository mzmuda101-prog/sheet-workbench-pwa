// phone-bottom-playwright.js — wstążka, etap 7b: dół ekranu na telefonie (docs/WSTAZKA.md; makieta
// docs/makiety/wstazka-v2.html, tryb „Telefon”). Decyzje Mateusza 2026-11-20: te same polecenia co we
// wstążce na komputerze (bez osobnego zestawu), rząd zawsze w Edycji, pasek „Kopiuj / Wklej / Wypełnij” znika.
//
//   1. Edycja: na dole rząd poleceń bieżącej karty — przycisk z nazwą karty + polecenia (te same
//      elementy co we wstążce: #rbCmds), jedna oś i wysokość 40 px,
//   2. rząd Szukaj odchudzony: pole + Filtry (z licznikiem); bez „Aa” i ▾,
//   3. stuknięcie komórki nie pokazuje paska „Kopiuj / Wklej / Wypełnij”; Wypełnij jest w karcie Główne,
//   4. polecenie z rzędu działa (pogrubienie komórki),
//   5. przycisk karty otwiera arkusz: wszystkie karty + polecenia z podpisami; wybór karty i ✕,
//   6. Podgląd: bez rzędu poleceń,
//   7. Ponów u góry, gdy jest co ponawiać,
//   8. nic nie ginie: Spisywanie / Wpisywanie w menu Trybu, Eksport w ⋯, powiększenie na karcie Widok,
//   9. przy otwartej klawiaturze rząd poleceń jedzie razem z paskiem cyfr,
//  10. komputer: polecenia z powrotem we wstążce.
//
// Uruchom z serwerem na APP_URL (domyślnie http://127.0.0.1:4175/). ENGINE=webkit — Safari.

const playwright = require("playwright");

const APP_URL = process.env.APP_URL || "http://127.0.0.1:4175/";
const ENGINE = process.env.ENGINE || "chromium";
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

async function open(browser, viewport, extra = {}) {
  const context = await browser.newContext({ serviceWorkers: "block", viewport, ...extra });
  await context.addInitScript(() => {
    if (!sessionStorage.getItem("__pb")) { sessionStorage.setItem("__pb", "1"); localStorage.clear(); localStorage.setItem("introPlayed", "true"); }
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
const td = (r, c) => `#dataTable tbody tr[data-row-key="wide:${r}"] td[data-col-index="${c}"]`;

async function run() {
  const browser = await playwright[ENGINE].launch({ headless: true });
  const results = [];
  const check = (name, ok, detail = "") => results.push({ name, ok: !!ok, detail: typeof detail === "string" ? detail : JSON.stringify(detail) });
  const touch = { hasTouch: true, isMobile: ENGINE === "chromium" };

  {
    const { context, page, errors } = await open(browser, { width: 390, height: 844 }, touch);
    // 1. rząd poleceń
    check("1. rząd poleceń widoczny na dole", await shown(page, "#phBottom"));
    check("1. polecenia to te same elementy co we wstążce (#rbCmds w rzędzie)", await inside(page, "#rbCmds", "#phBottom"));
    check("1. karta Główne: blok formatu w rzędzie", await inside(page, "#formatRibbon", "#phBottom"));
    const tabTxt = await page.evaluate(() => document.getElementById("phTabBtn")?.textContent.trim());
    check("1. przycisk karty z nazwą „Główne”", /Główne/.test(tabTxt || ""), tabTxt);
    const row = await page.evaluate(() => {
      const box = (el) => { const r = el.getBoundingClientRect(); return { h: Math.round(r.height), mid: Math.round(r.top + r.height / 2) }; };
      const b = document.getElementById("phBottom").getBoundingClientRect();
      const btns = [document.getElementById("phTabBtn"), ...document.querySelectorAll("#phBottom #rbPanel-home > .rb-btn")].filter((el) => el.getBoundingClientRect().width > 0);
      return { list: btns.map(box), bottom: Math.round(b.bottom), vh: innerHeight };
    });
    check("1. rząd: jedna oś, przyciski 40 px", row.list.length > 2 && new Set(row.list.map((x) => x.h)).size === 1 && row.list[0].h === 40 && Math.max(...row.list.map((x) => x.mid)) - Math.min(...row.list.map((x) => x.mid)) <= 1, row.list);
    check("1. rząd na ekranie (strona się nie przewija)", row.bottom <= row.vh && (await page.evaluate(() => document.documentElement.scrollHeight - innerHeight)) <= 1, row);

    // 2. rząd Szukaj
    check("2. bez „Aa” i ▾ w rzędzie Szukaj", !(await shown(page, "#fmtSheetBtn")) && !(await shown(page, "#toolbarToggle")));
    const filt = await page.evaluate(() => document.getElementById("panelToggle")?.getAttribute("aria-label"));
    check("2. przycisk Filtry (dawniej Narzędzia)", (await shown(page, "#panelToggle")) && /Filtry/.test(filt || ""), filt);

    // 3. pasek Kopiuj / Wklej / Wypełnij
    await page.tap(td(6, 2));
    await sleep(400);
    check("3. stuknięcie komórki: bez paska „Kopiuj / Wklej / Wypełnij”", !(await shown(page, "#cellActions")));
    const dis = await page.evaluate(() => [...document.querySelectorAll("#phBottom [data-rb-act]")].filter((b) => b.disabled).map((b) => b.dataset.rbAct));
    check("4. polecenia w rzędzie aktywne przy otwartym pliku (nie wyszarzone)", dis.length === 0, dis);
    check("3. Wypełnij w dół w karcie Główne", await page.evaluate(() => !!document.querySelector('#rbPanel-home [data-rb-act="fillDown"]')));

    // 4. pogrubienie z rzędu
    await page.evaluate(() => document.querySelector('#phBottom [data-fmt="b"]')?.scrollIntoView({ inline: "center" }));
    await sleep(200);
    await page.tap('#phBottom [data-fmt="b"]');
    await sleep(400);
    const bold = await page.evaluate((s) => { const el = document.querySelector(s); return el ? getComputedStyle(el).fontWeight : ""; }, td(6, 2));
    check("4. B z rzędu pogrubia komórkę", Number(bold) >= 600, bold);

    // 5. arkusz kart
    await page.tap("#phTabBtn");
    await sleep(350);
    check("5. arkusz kart otwarty", await shown(page, "#phSheet"));
    const tabs = await page.evaluate(() => [...document.querySelectorAll("#phSheet [role=tab]")].filter((b) => !b.hidden).map((b) => b.textContent.trim()));
    check("5. w arkuszu wszystkie karty", tabs.length >= 7, tabs);
    check("5. polecenia w arkuszu (te same #rbCmds)", await inside(page, "#rbCmds", "#phSheet"));
    await page.tap('#phSheet [role=tab][data-rtab="data"]');
    await sleep(300);
    // „Sortowanie domyślne” chowa się, gdy nie ma czego przywracać — sprawdzamy, że jest w arkuszu
    check("5. karta Dane: Sortowanie domyślne i Z szerokiej na długą w arkuszu", (await inside(page, "#resetSortBtn", "#phSheet #rbPanel-data")) && (await inside(page, "#wideLongToggle", "#phSheet #rbPanel-data")));
    const labelled = await page.evaluate(() => [...document.querySelectorAll("#phSheet #rbPanel-data .rb-btn")].every((b) => b.textContent.trim().length > 0 || getComputedStyle(b, "::after").content.length > 2));
    check("5. polecenia w arkuszu mają podpisy", labelled);
    await page.tap("#phSheetClose");
    await sleep(300);
    check("5. ✕ zamyka arkusz, w rzędzie karta Dane", !(await shown(page, "#phSheet")) && (await inside(page, "#rbCmds", "#phBottom")) && /Dane/.test(await page.evaluate(() => document.getElementById("phTabBtn").textContent)));

    // 8. nic nie ginie
    check("8. Spisywanie i Wpisywanie w menu Trybu", (await inside(page, "#transcribeBtn", "#rbModeMenu")) && (await inside(page, "#entryBtn", "#rbModeMenu")));
    check("8. Eksport w menu ⋯", await inside(page, "#exportCsvBtn", "#appMenu"));
    check("8. powiększenie na karcie Widok", await inside(page, "#zoomGroup", "#rbPanel-view"));
    check("8. Raport na karcie Analiza", await inside(page, "#reportBtn", "#rbPanel-an"));

    // 7. Ponów
    await page.tap(td(7, 2));
    await sleep(120);
    await page.tap(td(7, 2));
    await page.waitForSelector("#dataTable input.cell-editor", { timeout: 5000 }).catch(() => {});
    // 9. klawiatura: test nie ma klawiatury — wymuszamy pasek cyfr (window.__swbKeyBarForce)
    await page.evaluate(() => { window.__swbKeyBarForce = true; window.swbKeyBar?.sync(); });
    await sleep(300);
    check("9. przy klawiaturze rząd poleceń jedzie z paskiem cyfr", await inside(page, "#phBottom", ".key-bar"));
    await page.keyboard.type("X");
    await page.keyboard.press("Enter");
    await page.evaluate(() => { window.__swbKeyBarForce = false; window.swbKeyBar?.sync(); });
    await sleep(300);
    check("9. po edycji rząd wraca na dół", !(await inside(page, "#phBottom", ".key-bar")) && (await shown(page, "#phBottom")));
    await page.tap("#undoBtn");
    await sleep(400);
    check("7. Ponów u góry po Cofnij", (await shown(page, "#redoBtn")) && (await inside(page, "#redoBtn", "#phTop")));

    // 6. Podgląd
    await page.tap("#rbModeBtn");
    await sleep(250);
    await page.tap('#rbModeMenu [data-rb-mode="view"]');
    await sleep(400);
    check("6. Podgląd: bez rzędu poleceń", !(await shown(page, "#phBottom")));

    check("brak błędów strony (telefon)", errors.length === 0, errors.join(" | "));
    await context.close();
  }

  {
    const { context, page, errors } = await open(browser, { width: 1280, height: 800 });
    check("10. komputer: polecenia we wstążce, bez rzędu telefonu", (await inside(page, "#rbCmds", "#ribbon")) && !(await shown(page, "#phBottom")));
    check("10. komputer: blok formatu i powiększenie na swoich miejscach", (await inside(page, "#formatRibbon", "#ribbon")) && (await inside(page, "#zoomGroup", "#tableFoot")));
    check("brak błędów strony (komputer)", errors.length === 0, errors.join(" | "));
    await context.close();
  }

  await browser.close();
  const failed = results.filter((r) => !r.ok);
  results.forEach((r) => console.log(`${r.ok ? "✓" : "✗"} ${r.name}${r.ok ? "" : "  — " + r.detail}`));
  if (failed.length) throw new Error(`${failed.length} z ${results.length} asercji nie przeszło (${ENGINE})`);
  console.log(`\n✅ phone-bottom-playwright OK — ${results.length} asercji (${ENGINE})`);
}

run().catch((e) => { console.error(e.message || e); process.exit(1); });
