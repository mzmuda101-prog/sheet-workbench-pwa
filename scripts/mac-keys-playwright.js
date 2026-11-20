// mac-keys-playwright.js — skróty na Macu: tylko ⌘, Control wolny (prośba Mateusza 2026-11-20).
//
// Mac (udawany przez navigator.platform): ⌘Z cofa, Control+Z nie; ⌘D wypełnia w dół, Control+D nie;
// ⌘F1 = Tylko karty; podpowiedzi i lista skrótów pokazują ⌘ / ⌥ / ⇧, bez „Ctrl”.
// Windows (udawany): Ctrl+Z i Ctrl+D jak dotąd.
//
// Uruchom z serwerem na APP_URL (domyślnie http://127.0.0.1:4175/). ENGINE=webkit — Safari.

const playwright = require("playwright");

const APP_URL = process.env.APP_URL || "http://127.0.0.1:4175/";
const ENGINE = process.env.ENGINE || "chromium";
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

async function open(browser, platform) {
  const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1440, height: 900 } });
  await context.addInitScript((pf) => {
    Object.defineProperty(Navigator.prototype, "platform", { get: () => pf, configurable: true });
    Object.defineProperty(Navigator.prototype, "userAgentData", { get: () => undefined, configurable: true });
    if (!sessionStorage.getItem("__mk")) { sessionStorage.setItem("__mk", "1"); localStorage.clear(); localStorage.setItem("introPlayed", "true"); }
  }, platform);
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push("pageerror: " + e.message));
  await page.goto(APP_URL + "?sample=300", { waitUntil: "load" });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  await page.evaluate(() => document.getElementById("loadSampleBtn")?.click());
  await page.waitForFunction(() => document.querySelector("#dataTable tbody tr[data-row-key]"), null, { timeout: 30000 });
  await page.evaluate(() => { if (typeof setSidebarOpen === "function") setSidebarOpen(false); document.querySelectorAll(".toast").forEach((t) => t.remove()); });
  await sleep(400);
  return { context, page, errors };
}

const td = (r, c) => `#dataTable tbody tr[data-row-key="wide:${r}"] td[data-col-index="${c}"]`;
const text = (page, r, c) => page.evaluate((s) => document.querySelector(s)?.textContent.trim(), td(r, c));

async function run() {
  const browser = await playwright[ENGINE].launch({ headless: true });
  const results = [];
  const check = (name, ok, detail = "") => results.push({ name, ok: !!ok, detail: typeof detail === "string" ? detail : JSON.stringify(detail) });

  // ── Mac ──
  {
    const { context, page, errors } = await open(browser, "MacIntel");
    // Cofnij: Control+Z nic, ⌘Z cofa
    const orig = await text(page, 20, 3);
    await page.click(td(20, 3));
    await page.keyboard.type("Zmiana");
    await page.keyboard.press("Enter");
    await sleep(250);
    await page.click(td(25, 3));
    await page.keyboard.press("Control+z");
    await sleep(250);
    check("Mac: Control+Z nie cofa (Control wolny)", (await text(page, 20, 3)) === "Zmiana", await text(page, 20, 3));
    await page.keyboard.press("Meta+z");
    await sleep(250);
    check("Mac: ⌘Z cofa", (await text(page, 20, 3)) === orig, await text(page, 20, 3));
    // Wypełnij w dół: zaznaczenie D21:D22 (wiersze 20–21), wartość z D21 do D22
    const top = await text(page, 20, 3);
    const below = await text(page, 21, 3);
    await page.click(td(20, 3));
    await page.click(td(21, 3), { modifiers: ["Shift"] });
    await page.keyboard.press("Control+d");
    await sleep(250);
    check("Mac: Control+D nie wypełnia", (await text(page, 21, 3)) === below, await text(page, 21, 3));
    await page.keyboard.press("Meta+d");
    await sleep(300);
    check("Mac: ⌘D wypełnia w dół (jak Excel na Macu)", (await text(page, 21, 3)) === top, await text(page, 21, 3));
    await page.keyboard.press("Meta+z");
    await sleep(250);
    // ⌘F1 = Tylko karty, Control+F1 nic
    await page.keyboard.press("Control+F1");
    await sleep(200);
    check("Mac: Control+F1 nie przełącza wstążki", !(await page.evaluate(() => document.body.classList.contains("rb-show-tabs"))));
    await page.keyboard.press("Meta+F1");
    await sleep(200);
    check("Mac: ⌘F1 = Tylko karty", await page.evaluate(() => document.body.classList.contains("rb-show-tabs")));
    await page.keyboard.press("Meta+F1");
    await sleep(200);
    // napisy: podpowiedzi i lista skrótów bez „Ctrl”
    const hint = await page.evaluate(() => typeof swbKeysText === "function" ? [swbKeysText("Pogrubienie (Ctrl/⌘+B)"), swbKeysText("Malarz (Ctrl/⌘+Alt+R)"), swbKeysText("Przekreślenie (Ctrl+5)"), swbKeysText("Notatka (Shift+F2)")] : null);
    check("Mac: tekst skrótów ⌘ / ⌥ / ⇧", JSON.stringify(hint) === JSON.stringify(["Pogrubienie (⌘B)", "Malarz (⌘⌥R)", "Przekreślenie (⌘5)", "Notatka (⇧F2)"]), hint);
    const list = await page.evaluate(() => { swbRenderShortcutPanel(); return document.getElementById("shortcutRegistryList")?.textContent || ""; });
    check("Mac: lista skrótów bez „Ctrl”", list.length > 50 && !/Ctrl/.test(list), list.match(/.{0,30}Ctrl.{0,30}/)?.[0] || list.length);
    const kbd = await page.evaluate(() => document.querySelector('#rbOptsMenu [data-rb-show="tabs"] kbd')?.textContent);
    check("Mac: menu opcji pokazuje ⌘F1", kbd === "⌘F1", kbd);
    const boldHint = await page.evaluate(() => document.querySelector('#formatRibbon [data-fmt="b"]')?.dataset.hintPl);
    check("Mac: podpowiedź przy B „Pogrubienie (⌘B)”", boldHint === "Pogrubienie (⌘B)", boldHint);
    // dane użytkownika (tu: tekst dymka z nazwą pliku) nie są przerabiane na symbole
    const raw = await page.evaluate(() => { toast("Ctrl+Z notatki.xlsx", "info"); return [...document.querySelectorAll(".toast")].pop()?.textContent || ""; });
    check("Mac: tekst użytkownika bez zmian („Ctrl+Z notatki.xlsx”)", raw.includes("Ctrl+Z notatki.xlsx"), raw);
    // ⌘R z fokusem poza tabelą (kliknięcie obok): nie przechwytujemy — przeglądarka odświeża stronę
    await page.click(td(21, 3));
    const r = await page.evaluate(() => {
      document.activeElement?.blur();
      const ev = new KeyboardEvent("keydown", { key: "r", code: "KeyR", metaKey: true, bubbles: true, cancelable: true });
      document.body.dispatchEvent(ev);
      return ev.defaultPrevented;
    });
    check("Mac: ⌘R poza tabelą zostaje przeglądarce (bez wypełniania)", r === false);
    check("brak błędów strony (Mac)", errors.length === 0, errors.join(" | "));
    await context.close();
  }

  // ── Windows ──
  {
    const { context, page, errors } = await open(browser, "Win32");
    const orig = await text(page, 20, 3);
    await page.click(td(20, 3));
    await page.keyboard.type("Zmiana");
    await page.keyboard.press("Enter");
    await sleep(250);
    await page.click(td(25, 3));
    await page.keyboard.press("Control+z");
    await sleep(250);
    check("Windows: Ctrl+Z cofa jak dotąd", (await text(page, 20, 3)) === orig, await text(page, 20, 3));
    await page.keyboard.press("Control+y");
    await sleep(250);
    check("Windows: Ctrl+Y ponawia", (await text(page, 20, 3)) === "Zmiana", await text(page, 20, 3));
    const list = await page.evaluate(() => { swbRenderShortcutPanel(); return document.getElementById("shortcutRegistryList")?.textContent || ""; });
    check("Windows: lista skrótów z „Ctrl”", /Ctrl/.test(list) && !/⌘/.test(list));
    check("brak błędów strony (Windows)", errors.length === 0, errors.join(" | "));
    await context.close();
  }

  await browser.close();
  const failed = results.filter((r) => !r.ok);
  results.forEach((r) => console.log(`${r.ok ? "✓" : "✗"} ${r.name}${r.ok ? "" : "  — " + r.detail}`));
  if (failed.length) throw new Error(`${failed.length} z ${results.length} asercji nie przeszło (${ENGINE})`);
  console.log(`\n✅ mac-keys-playwright OK — ${results.length} asercji (${ENGINE})`);
}

run().catch((e) => { console.error(e.message || e); process.exit(1); });
