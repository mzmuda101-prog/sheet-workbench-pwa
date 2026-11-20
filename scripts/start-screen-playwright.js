// start-screen-playwright.js — etap 8: ekran startowy jak w Documents Workbench (Mateusz 2026-11-20).
//
//   1. komputer bez pliku: u góry logo · Plik · „Sheet Workbench” · ⋯ — bez kart, poleceń, szukania, trybu
//      i szyny; bez dolnego paska; odstęp po szynie nie zostaje,
//   2. środek: kafelek z ikoną, duży tytuł „Otwórz arkusz”, zdanie o prywatności, „Otwórz plik…” (główny)
//      + „Nowy skoroszyt” i „Przykładowy plik”, podpowiedź o przeciąganiu,
//   3. po otwarciu pliku wstążka wraca (karty, polecenia, szyna), nazwa aplikacji znika,
//   4. plik upuszczony w dowolnym miejscu okna otwiera się,
//   5. telefon bez pliku: bez klucza i „Wysuń”; główny przycisk na całą szerokość, dwa pozostałe równe.
//
// Uruchom z serwerem na APP_URL (domyślnie http://127.0.0.1:4175/). ENGINE=webkit — Safari.

const playwright = require("playwright");
const fs = require("fs");
const path = require("path");

const APP_URL = process.env.APP_URL || "http://127.0.0.1:4175/";
const ENGINE = process.env.ENGINE || "chromium";
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

async function open(browser, viewport, extra = {}) {
  const context = await browser.newContext({ serviceWorkers: "block", viewport, ...extra });
  await context.addInitScript(() => {
    if (!sessionStorage.getItem("__st")) { sessionStorage.setItem("__st", "1"); localStorage.clear(); localStorage.setItem("introPlayed", "true"); }
  });
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push("pageerror: " + e.message));
  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  await sleep(700);
  return { context, page, errors };
}
const shown = (page, sel) => page.evaluate((s) => { const el = document.querySelector(s); if (!el) return false; const r = el.getBoundingClientRect(); return r.width > 0 && r.height > 0 && getComputedStyle(el).visibility !== "hidden"; }, sel);

async function run() {
  const browser = await playwright[ENGINE].launch({ headless: true });
  const results = [];
  const check = (name, ok, detail = "") => results.push({ name, ok: !!ok, detail: typeof detail === "string" ? detail : JSON.stringify(detail) });

  {
    const { context, page, errors } = await open(browser, { width: 1280, height: 800 });
    // 1.
    check("1. komputer: nazwa aplikacji obok Plik", (await shown(page, "#rbAppName")) && (await shown(page, "#rbFileBtn")) && (await shown(page, "#appMenuBtn")));
    const hidden = await page.evaluate(() => ["#rbTab-home", "#rbCmds", "#quickSearchWrap", "#rbModeBtn", "#sidebarRail", "#tableFoot"].filter((s) => { const el = document.querySelector(s); return el && el.getBoundingClientRect().width > 0; }));
    check("1. komputer: bez kart, poleceń, szukania, trybu, szyny i dolnego paska", hidden.length === 0, hidden);
    const left = await page.evaluate(() => Math.round(document.getElementById("emptyState").closest(".table-panel, .panel")?.getBoundingClientRect().left ?? -1));
    check("1. komputer: bez pustego odstępu po szynie", left >= 0 && left < 30, left);
    // 2.
    const sc = await page.evaluate(() => ({
      title: document.getElementById("emptyTitle").textContent.trim(),
      size: parseFloat(getComputedStyle(document.getElementById("emptyTitle")).fontSize),
      body: parseFloat(getComputedStyle(document.getElementById("emptySub")).fontSize),
      tile: getComputedStyle(document.querySelector("#emptyState .empty-icon")).backgroundColor,
      btns: [...document.querySelectorAll("#emptyState .empty-actions .btn")].filter((b) => b.offsetParent).map((b) => b.textContent.trim()),
      primary: !document.getElementById("emptyOpenBtn").classList.contains("ghost"),
      sub: document.getElementById("emptySub").textContent,
    }));
    check("2. tytuł „Otwórz arkusz”, co najmniej 1,8× większy niż tekst", sc.title === "Otwórz arkusz" && sc.size >= sc.body * 1.8, sc);
    check("2. kafelek z ikoną (tło akcentu)", sc.tile !== "rgba(0, 0, 0, 0)", sc.tile);
    check("2. przyciski: Otwórz plik… (główny), Nowy skoroszyt, Przykładowy plik", sc.primary && sc.btns.join("|") === "Otwórz plik…|Nowy skoroszyt|Przykładowy plik", sc.btns);
    check("2. zdanie o prywatności", /nic nie jest wysyłane/.test(sc.sub), sc.sub);
    check("2. podpowiedź o przeciąganiu w dowolne miejsce okna", await shown(page, "#emptyDropHint"));
    // 4. upuszczenie pliku w dowolnym miejscu okna (poza strefą w panelu)
    // neutralny plik testowy z repo (bez danych prywatnych)
    const b64 = fs.readFileSync(path.join(__dirname, "excel-theme-colors-2023.xlsx")).toString("base64");
    if (b64) {
      await page.evaluate((data) => {
        const bin = Uint8Array.from(atob(data), (c) => c.charCodeAt(0));
        const file = new File([bin], "upuszczony.xlsx", { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
        const dt = new DataTransfer();
        dt.items.add(file);
        const target = document.querySelector(".ribbon") || document.body;
        target.dispatchEvent(new DragEvent("dragenter", { bubbles: true, cancelable: true, dataTransfer: dt }));
        target.dispatchEvent(new DragEvent("dragover", { bubbles: true, cancelable: true, dataTransfer: dt }));
        target.dispatchEvent(new DragEvent("drop", { bubbles: true, cancelable: true, dataTransfer: dt }));
      }, b64);
      await page.waitForFunction(() => typeof workbook !== "undefined" && workbook && document.querySelector("#dataTable tbody tr"), null, { timeout: 20000 }).catch(() => {});
      check("4. plik upuszczony poza strefą panelu otwiera się", await page.evaluate(() => !!workbook && currentFileName === "upuszczony.xlsx"));
    } else {
      check("4. (brak pliku do upuszczenia — pomijam)", true);
      await page.evaluate(() => document.getElementById("loadSampleBtn").click());
      await page.waitForFunction(() => document.querySelector("#dataTable tbody tr[data-row-key]"), null, { timeout: 30000 });
    }
    await sleep(500);
    // 3.
    check("3. z plikiem: karty, polecenia i szyna wracają, nazwy aplikacji nie ma", (await shown(page, "#rbTab-home")) && (await shown(page, "#rbCmds")) && (await shown(page, "#sidebarRail")) && !(await shown(page, "#rbAppName")));
    check("brak błędów strony (komputer)", errors.length === 0, errors.join(" | "));
    await context.close();
  }

  {
    const { context, page, errors } = await open(browser, { width: 390, height: 844 }, { hasTouch: true, isMobile: ENGINE === "chromium" });
    check("5. telefon: bez klucza i „Wysuń”", !(await shown(page, "#panelToggle")) && !(await shown(page, "#panelHandle")));
    const g = await page.evaluate(() => {
      const w = (s) => Math.round(document.querySelector(s).getBoundingClientRect().width);
      return { open: w("#emptyOpenBtn"), nw: w("#emptyNewBtn"), smp: w("#loadSampleBtn"), row: w("#emptyState .empty-actions") };
    });
    check("5. telefon: Otwórz plik… na całą szerokość, dwa pozostałe równe", g.open >= g.row - 2 && Math.abs(g.nw - g.smp) <= 1 && g.nw < g.open, g);
    check("brak błędów strony (telefon)", errors.length === 0, errors.join(" | "));
    await context.close();
  }

  await browser.close();
  const failed = results.filter((r) => !r.ok);
  results.forEach((r) => console.log(`${r.ok ? "✓" : "✗"} ${r.name}${r.ok ? "" : "  — " + r.detail}`));
  if (failed.length) throw new Error(`${failed.length} z ${results.length} asercji nie przeszło (${ENGINE})`);
  console.log(`\n✅ start-screen-playwright OK — ${results.length} asercji (${ENGINE})`);
}

run().catch((e) => { console.error(e.message || e); process.exit(1); });
