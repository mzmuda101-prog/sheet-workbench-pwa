// ux-contract-playwright.js — wspólny kontrakt UX Sheet + Docs (2026-11-07) i audyt QoL S6–S14.
//   node scripts/ux-contract-playwright.js            (Chromium)
//   ENGINE=webkit node scripts/ux-contract-playwright.js
//
// Pilnuje:
//   1. motyw 3-stanowy: Systemowy = brak klucza + śledzi system NA ŻYWO; Jasny/Ciemny zapisują;
//   2. dymek: role status/alert, ✕ przy błędzie, akcja w dymku działa;
//   3. Ctrl/⌘+/ = okno skrótów z wyszukiwaniem; panel Pomoc rysowany z TEGO SAMEGO rejestru;
//      każdy wpis rejestru ma tekst PL i EN; Ctrl/⌘+Alt+/ = szybki motyw;
//   4. wersja widoczna w menu ⋯;
//   5. okno niezapisanej pracy: Zapisz / Nie zapisuj / Anuluj (Anuluj zostawia plik);
//      własne okno nazwy: pusta nazwa = komunikat, okno zostaje;
//   6. wyczyszczenie komórek → „Cofnij” w dymku przywraca wartości;
//   7. kopia ustawień: eksport zbiera preferencje (bez pomiarów urządzenia), import po podglądzie,
//      obce klucze odrzucone;
//   8. ponowne otwarcie tego samego pliku → „Wróć” skacze do ostatniej komórki;
//   9. przykład „Znajdź otwarte cykle” z pustego ekranu filtruje przykład;
//  10. telefon: uchwyt panelu chowa się przy przewijaniu tabeli i wraca;
//  11. zero natywnych confirm()/prompt() w zwykłych przepływach (grep po app/).

const { chromium, webkit } = require("playwright");
const fs = require("fs");
const path = require("path");

const APP_URL = process.env.APP_URL || "http://127.0.0.1:4175/";
const ENGINE_NAME = process.env.ENGINE === "webkit" ? "webkit" : "chromium";
const ENGINE = ENGINE_NAME === "webkit" ? webkit : chromium;
const FILE = path.join(__dirname, "stress-test-workbench.xlsx");
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

async function run() {
  const fails = [];
  const check = (name, ok, detail) => {
    console.log(`${ok ? "✅" : "❌"} ${name}${ok ? "" : ` — ${JSON.stringify(detail)}`}`);
    if (!ok) fails.push(name);
  };

  // 11. grep: natywne okna tylko w narzędziach diagnostycznych
  const appDir = path.join(__dirname, "..", "app");
  const native = fs.readdirSync(appDir).filter((f) => f.endsWith(".js") && !/scroll-diagnostics|ipad-scroll-debug/.test(f))
    .flatMap((f) => fs.readFileSync(path.join(appDir, f), "utf8").split("\n").map((l, i) => ({ f, i: i + 1, l })))
    .filter(({ l }) => /\b(window\.)?(confirm|prompt|alert)\(/.test(l) && !/^\s*\/\//.test(l) && !/appConfirm|appPrompt|zamiast|function /.test(l));
  check("11. brak natywnych confirm()/prompt()/alert() w app/", native.length === 0, native.map((n) => `${n.f}:${n.i}`));

  const browser = await ENGINE.launch({ headless: true });
  const ctx = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1280, height: 900 }, colorScheme: "dark" });
  await ctx.addInitScript(() => sessionStorage.setItem("introPlayed", "true"));
  const page = await ctx.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  page.on("dialog", (d) => { errors.push(`natywne okno: ${d.message()}`); d.dismiss(); });
  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());

  // 1. motyw
  const th0 = await page.evaluate(() => ({ theme: document.documentElement.dataset.theme, key: localStorage.getItem("excel-workbench-theme"), sel: document.querySelector("#themeSeg [aria-checked='true']")?.dataset.themePref }));
  check("1. bez zapisanego wyboru: Systemowy, wygląd = system (ciemny)", th0.theme === "dark" && th0.key === null && th0.sel === "system", th0);
  await page.evaluate(() => document.querySelector("#themeSeg [data-theme-pref='light']").click());
  const th1 = await page.evaluate(() => ({ theme: document.documentElement.dataset.theme, key: localStorage.getItem("excel-workbench-theme") }));
  check("1. Jasny zapisuje „light”", th1.theme === "light" && th1.key === "light", th1);
  await page.evaluate(() => document.querySelector("#themeSeg [data-theme-pref='system']").click());
  await page.emulateMedia({ colorScheme: "light" });
  await sleep(150);
  const th2 = await page.evaluate(() => ({ theme: document.documentElement.dataset.theme, key: localStorage.getItem("excel-workbench-theme") }));
  check("1. Systemowy usuwa klucz i śledzi zmianę systemu na żywo", th2.theme === "light" && th2.key === null, th2);
  await page.reload({ waitUntil: "load" });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());

  // 2. dymek
  const ts = await page.evaluate(() => {
    document.querySelectorAll(".toast").forEach((x) => x.remove());
    const e = toast("Błąd testowy", "error");
    const i = toast("Info testowe", "info");
    let hit = 0;
    const a = toast("Z akcją", "success", { action: "Cofnij", onAction: () => { hit++; } });
    const r = { eRole: e.el.getAttribute("role"), eClose: !!e.el.querySelector(".toast-close"), iRole: i.el.getAttribute("role"), iClose: !!i.el.querySelector(".toast-close"), aBtn: a.el.querySelector(".toast-action")?.textContent, aClose: !!a.el.querySelector(".toast-close") };
    a.el.querySelector(".toast-action").click();
    r.hit = hit;
    return r;
  });
  check("2. błąd: role=alert + ✕; info: role=status bez ✕; akcja działa", ts.eRole === "alert" && ts.eClose && ts.iRole === "status" && !ts.iClose && ts.aBtn === "Cofnij" && ts.aClose && ts.hit === 1, ts);

  // 3. skróty
  await page.keyboard.press(ENGINE_NAME === "webkit" ? "Meta+Slash" : "Control+Slash");
  await page.waitForSelector(".sc-dialog", { timeout: 3000 }).catch(() => {});
  const sc = await page.evaluate(() => {
    const d = document.querySelector(".sc-dialog");
    if (!d) return { open: false };
    const all = d.querySelectorAll(".shortcut-row").length;
    const s = d.querySelector(".sc-search");
    s.value = "zapisz";
    s.dispatchEvent(new Event("input"));
    const rows = [...d.querySelectorAll(".shortcut-row")].map((r) => r.textContent);
    return { open: true, all, total: SWB_SHORTCUTS.length, filtered: rows.length, hasSave: rows.some((r) => /Zapisz/.test(r)) };
  });
  check("3. Ctrl/⌘+/ otwiera okno skrótów z całym rejestrem", sc.open && sc.all === sc.total, sc);
  check("3. wyszukiwanie zawęża listę („zapisz”)", sc.filtered > 0 && sc.filtered < sc.all && sc.hasSave, sc);
  await page.keyboard.press("Escape");
  await sleep(100);
  const reg = await page.evaluate(() => {
    const closed = !document.querySelector(".sc-dialog");
    const panelRows = document.querySelectorAll("#shortcutRegistryList .shortcut-row").length;
    const missing = [];
    for (const lang of ["pl", "en"]) {
      const prev = currentLang; currentLang = lang;
      SWB_SHORTCUTS.forEach((s) => { if (t(s.text) === s.text) missing.push(`${lang}:${s.text}`); });
      SWB_SHORTCUT_CATS.forEach((c) => { if (t(c.label) === c.label) missing.push(`${lang}:${c.label}`); });
      currentLang = prev;
    }
    return { closed, panelRows, total: SWB_SHORTCUTS.length, missing };
  });
  check("3. Esc zamyka okno; panel Pomoc = ten sam rejestr; teksty PL i EN", reg.closed && reg.panelRows === reg.total && !reg.missing.length, reg);
  const before = await page.evaluate(() => document.documentElement.dataset.theme);
  await page.keyboard.press(ENGINE_NAME === "webkit" ? "Meta+Alt+Slash" : "Control+Alt+Slash");
  const after = await page.evaluate(() => document.documentElement.dataset.theme);
  check("3. Ctrl/⌘+Alt+/ przełącza motyw", before !== after, { before, after });
  await page.evaluate(() => { localStorage.removeItem("excel-workbench-theme"); applyThemePref("system", false); });

  // 4. wersja w menu ⋯
  await page.click("#appMenuBtn");
  const ver = await page.evaluate(() => ({ text: document.getElementById("appMenuBuild").textContent, v: APP_BUILD_VERSION, visible: document.getElementById("appMenuVersion").offsetParent !== null }));
  check("4. menu ⋯ pokazuje „Wersja vX.YY.ZZ”", ver.visible && ver.text.includes(`Wersja ${ver.v}`) && /^v\d+\.\d{2}\.\d{2}$/.test(ver.v), ver);
  await page.keyboard.press("Escape");

  // otwarcie pliku z dysku (stały klucz sesji: nazwa+rozmiar+data)
  const openFile = async () => {
    await page.evaluate(() => { document.getElementById("fileInput").value = ""; }); // ten sam plik drugi raz = nowe „change”
    await page.setInputFiles("#fileInput", FILE);
    await page.waitForFunction(() => typeof currentDisplayModel !== "undefined" && currentDisplayModel && currentDisplayModel.rows && currentDisplayModel.rows.length > 5, null, { timeout: 20000 });
    await sleep(400);
  };
  await openFile();
  await page.evaluate(() => { if (typeof setSidebarOpen === "function") setSidebarOpen(false); sheetEditUi.setMode("edit"); document.querySelectorAll(".toast").forEach((x) => x.remove()); });

  // 6. Wyczyść → Cofnij w dymku
  const und = await page.evaluate(async () => {
    const ref = "B8";
    const v0 = workbook.Sheets[currentSheetName][ref]?.v;
    sheetEditUi.goToRef(ref);
    cfmtClearContents();
    await new Promise((r) => setTimeout(r, 50));
    const v1 = workbook.Sheets[currentSheetName][ref]?.v;
    const btn = [...document.querySelectorAll(".toast .toast-action")].pop();
    const label = btn?.textContent;
    btn?.click();
    await new Promise((r) => setTimeout(r, 100));
    const v2 = workbook.Sheets[currentSheetName][ref]?.v;
    const redo = [...document.querySelectorAll(".toast .toast-action")].pop()?.textContent;
    return { v0: String(v0), v1: v1 == null ? "" : String(v1), v2: String(v2), label, redo };
  });
  check("6. „Wyczyść” bez pytania, „Cofnij” w dymku przywraca wartość, potem „Ponów”", und.v0 !== "" && und.v1 === "" && und.v2 === und.v0 && und.label === "Cofnij" && und.redo === "Ponów", und);

  // 5. niezapisane zmiany przy otwieraniu innego pliku
  await page.evaluate(() => { setDirtyState(true); loadSampleFile(); });
  await page.waitForSelector(".app-dialog", { timeout: 4000 }).catch(() => {});
  const un = await page.evaluate(() => {
    const d = document.querySelector(".app-dialog");
    return d ? { buttons: [...d.querySelectorAll(".app-dialog-actions button")].map((b) => `${b.textContent}:${b.dataset.kind}`), modal: d.getAttribute("aria-modal") } : null;
  });
  check("5. okno „Zapisz / Nie zapisuj / Anuluj” (Zapisz = główny)", un && un.modal === "true" && un.buttons.join("|") === "Zapisz:primary|Nie zapisuj:danger|Anuluj:secondary", un);
  await page.click(".app-dialog .app-dialog-actions button:nth-child(3)");
  await sleep(300);
  const kept = await page.evaluate(() => ({ name: currentFileName, dirty: hasUnsavedChanges, dlg: !!document.querySelector(".app-dialog") }));
  check("5. Anuluj zostawia otwarty plik i zmiany", kept.name === "stress-test-workbench.xlsx" && kept.dirty && !kept.dlg, kept);

  const pr = await page.evaluate(async () => {
    const p = appPrompt("Nazwa:", "", { validate: (v) => (v ? "" : t("dlgNameEmpty")) });
    await new Promise((r) => setTimeout(r, 30));
    const d = document.querySelector(".app-dialog");
    d.querySelector(".app-dialog-actions .btn:not(.ghost)").click();
    const err = !d.querySelector(".app-dialog-error").hidden && d.isConnected;
    d.querySelector("input").value = "Mój widok";
    d.querySelector(".app-dialog-actions .btn:not(.ghost)").click();
    return { err, val: await p };
  });
  check("5. własne okno nazwy: pusta = komunikat, potem zwraca wpisany tekst", pr.err && pr.val === "Mój widok", pr);

  // 8. powrót do ostatniej komórki — w NOWEJ sesji (zamknięcie i ponowne uruchomienie apki)
  await page.evaluate(() => { setDirtyState(false); sheetEditUi.goToRef("C20"); });
  await page.evaluate(() => window.dispatchEvent(new Event("pagehide")));
  await page.reload({ waitUntil: "load" });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  await openFile();
  const resumeBtn = await page.waitForFunction(() => [...document.querySelectorAll(".toast")].find((x) => /Wrócić/.test(x.textContent))?.querySelector(".toast-action"), null, { timeout: 5000 }).catch(() => null);
  check("8. ponowne otwarcie pliku: dymek „Wrócić do … C20?”", !!resumeBtn, "brak dymku");
  if (resumeBtn) {
    await page.evaluate(() => [...document.querySelectorAll(".toast")].find((x) => /Wrócić/.test(x.textContent)).querySelector(".toast-action").click());
    await sleep(300);
    const at = await page.evaluate(() => cfmtActiveCell()?.ref);
    check("8. „Wróć” ustawia aktywną komórkę C20", at === "C20", at);
  }

  // 7. kopia ustawień
  const bk = await page.evaluate(async () => {
    localStorage.setItem("swb-saved-views-v1", JSON.stringify([{ name: "Test", vs: {} }]));
    localStorage.setItem("swb-render-cost-v1", "pomiar");
    const data = swbBackupCollect();
    const out = { hasViews: "swb-saved-views-v1" in data, hasCost: "swb-render-cost-v1" in data, hasTheme: Object.keys(data).every((k) => k.startsWith("swb-") || k.startsWith("excel-workbench-") || k === "localAi.prefs.v1") };
    localStorage.removeItem("swb-saved-views-v1");
    const file = new File([JSON.stringify({ app: "sheet-workbench", kind: "settings", version: "v1.00.00", exportedAt: new Date().toISOString(), data: { "swb-saved-views-v1": "[]", "evil-key": "x", "swb-render-cost-v1": "y" } })], "kopia.json", { type: "application/json" });
    const p = swbImportSettings(file);
    await new Promise((r) => setTimeout(r, 50));
    const d = document.querySelector(".app-dialog");
    out.preview = d ? d.querySelector(".app-dialog-msg").textContent : "";
    d?.querySelector(".app-dialog-actions .btn:not(.ghost)")?.click();
    out.ok = await p;
    out.views = localStorage.getItem("swb-saved-views-v1");
    out.evil = localStorage.getItem("evil-key");
    out.cost = localStorage.getItem("swb-render-cost-v1");
    return out;
  });
  check("7. eksport: preferencje tak, pomiary urządzenia nie", bk.hasViews && !bk.hasCost && bk.hasTheme, bk);
  check("7. import po podglądzie; obce klucze i pomiary odrzucone", bk.ok === true && /Kopia z/.test(bk.preview) && bk.views === "[]" && bk.evil === null && bk.cost === "pomiar", bk);
  await page.close();

  // 9 + 10. telefon: przykład z pustego ekranu i uchwyt panelu
  const phone = await browser.newContext({ serviceWorkers: "block", viewport: { width: 390, height: 844 }, hasTouch: true, isMobile: ENGINE_NAME === "chromium" });
  await phone.addInitScript(() => sessionStorage.setItem("introPlayed", "true"));
  const m = await phone.newPage();
  m.on("pageerror", (e) => errors.push(e.message));
  await m.goto(APP_URL, { waitUntil: "load" });
  await m.evaluate(() => document.getElementById("heroSplash")?.remove());
  const exVisible = await m.evaluate(() => document.querySelectorAll("#emptyExamples [data-example]").length);
  check("9. pusty ekran: 3 przykłady", exVisible === 3, exVisible);
  await m.evaluate(() => document.querySelector("#emptyExamples [data-example='open']").click());
  await m.waitForFunction(() => typeof quickSearchEl !== "undefined" && quickSearchEl.value === "@wtoku" && viewRows.length > 0 && viewRows.length < baseRows.length, null, { timeout: 15000 }).catch(() => {});
  const ex = await m.evaluate(() => ({ q: quickSearchEl.value, shown: viewRows.length, total: baseRows.length }));
  check("9. „Znajdź otwarte cykle” otwiera przykład i filtruje wiersze w toku", ex.q === "@wtoku" && ex.shown > 0 && ex.shown < ex.total, ex);
  await m.evaluate(() => { if (typeof setSidebarOpen === "function") setSidebarOpen(false); });
  await sleep(300);
  const h1 = await m.evaluate(async () => {
    const h = document.getElementById("panelHandle");
    tableWrapEl.scrollTop += 120;
    tableWrapEl.dispatchEvent(new Event("scroll"));
    await new Promise((r) => setTimeout(r, 80));
    const quiet = h.classList.contains("is-quiet");
    await new Promise((r) => setTimeout(r, 1500));
    return { quiet, back: !h.classList.contains("is-quiet") };
  });
  check("10. telefon: uchwyt chowa się przy przewijaniu tabeli i wraca", h1.quiet && h1.back, h1);
  await phone.close();

  // „ResizeObserver loop…” to niegroźny komunikat WebKitu przy zmianie układu — nie błąd aplikacji
  const real = errors.filter((e) => !/ResizeObserver loop/.test(e));
  check("bez błędów strony i natywnych okien", real.length === 0, real.slice(0, 4));
  await browser.close();
  if (fails.length) { console.log(`\n❌ ${fails.length} porażek`); process.exit(1); }
  console.log("\n✅ Kontrakt UX: wszystko OK");
}

run().catch((e) => { console.error(e); process.exit(1); });
