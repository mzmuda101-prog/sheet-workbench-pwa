// table-undo-playwright.js — „Formatuj jako tabelę”: Cofnij / Ponów i komunikat z zakresem
// (nagranie Mateusza 2026-11-19; decyzja 2026-11-20: zakres zostaje — nagłówek + wszystkie wiersze
// danych w zaznaczonych kolumnach — a komunikat mówi to wprost i daje „Tylko zaznaczenie”).
//
//   1. zaznaczenie H8:I13 → tabela H5:I305; komunikat opisuje zakres i ma „Cofnij” oraz
//      „Tylko zaznaczenie (H8:I13)”,
//   2. Ctrl+Z cofa tabelę (plik jak przed), Ctrl+Y ją przywraca,
//   3. „Tylko zaznaczenie” = tabela dokładnie H8:I13 (pierwszy wiersz zaznaczenia to nagłówek),
//      a jedno Cofnij ją usuwa.
//
// Uruchom z serwerem na APP_URL (domyślnie http://127.0.0.1:4175/). ENGINE=webkit — Safari.

const playwright = require("playwright");

const APP_URL = process.env.APP_URL || "http://127.0.0.1:4175/";
const ENGINE = process.env.ENGINE || "chromium";
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

async function run() {
  const browser = await playwright[ENGINE].launch({ headless: true });
  const results = [];
  const check = (name, ok, detail = "") => results.push({ name, ok: !!ok, detail: typeof detail === "string" ? detail : JSON.stringify(detail) });
  const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1440, height: 900 } });
  await context.addInitScript(() => { if (!sessionStorage.getItem("__tu")) { sessionStorage.setItem("__tu", "1"); localStorage.clear(); localStorage.setItem("introPlayed", "true"); } });
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push("pageerror: " + e.message));
  page.on("console", (m) => { if (m.type() === "error") errors.push("console: " + m.text()); });
  await page.goto(APP_URL + "?sample=300", { waitUntil: "load" });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  await page.evaluate(() => document.getElementById("loadSampleBtn")?.click());
  await page.waitForFunction(() => document.querySelector("#dataTable tbody tr[data-row-key]"), null, { timeout: 30000 });
  await page.evaluate(() => { if (typeof setSidebarOpen === "function") setSidebarOpen(false); });
  await sleep(400);

  const td = (r, c) => `#dataTable tbody tr[data-row-key="wide:${r}"] td[data-col-index="${c}"]`;
  const tables = () => page.evaluate(() => [...new Set(Object.values(currentTables || {}))].filter((x) => x.sheetName === currentSheetName).map((x) => `${x.displayName}:${x.ref}`).sort());
  const idle = () => page.waitForFunction(() => !(typeof swbSheetOpsUi !== "undefined" && swbSheetOpsUi.isBusy()) && document.body.getAttribute("aria-busy") !== "true", null, { timeout: 20000 });
  const dropToasts = () => page.evaluate(() => document.querySelectorAll(".toast").forEach((x) => x.remove()));
  // H8:I13 (wiersze arkusza 8–13 = rowIndex0 7–12, kolumny H–I = indeksy 7–8)
  async function selectH8I13() {
    await page.click(td(7, 7));
    await page.click(td(12, 8), { modifiers: ["Shift"] });
    await sleep(150);
  }
  const before = await tables();

  // 1. tabela z komunikatem
  await selectH8I13();
  await dropToasts();
  await page.click("#fmtTableBtn");
  await idle();
  await sleep(600);
  const t1 = await tables();
  check("1. tabela: nagłówek + wszystkie wiersze danych w kolumnach H–I", t1.length === before.length + 1 && t1.some((x) => /:H5:I\d+$/.test(x)), t1);
  const toast1 = await page.evaluate(() => {
    const el = [...document.querySelectorAll(".toast")].pop();
    return el ? { text: el.textContent, actions: [...el.querySelectorAll(".toast-action")].map((b) => b.textContent.trim()) } : null;
  });
  check("1. komunikat opisuje zakres (wiersz nagłówka i kolumny)", toast1 && /wiersza 5/.test(toast1.text) && /H.I/.test(toast1.text), toast1);
  check("1. komunikat: „Cofnij” i „Tylko zaznaczenie (H8:I13)”", toast1 && toast1.actions.includes("Cofnij") && toast1.actions.some((a) => /Tylko zaznaczenie.*H8:I13/.test(a)), toast1);

  // 2. Ctrl+Z / Ctrl+Y
  await dropToasts();
  await page.click(td(20, 3));
  await page.keyboard.press("ControlOrMeta+z");
  await idle();
  await sleep(500);
  check("2. Ctrl+Z cofa tabelę", JSON.stringify(await tables()) === JSON.stringify(before), await tables());
  check("2. po Cofnij jest Ponów", await page.evaluate(() => !document.getElementById("redoBtn")?.disabled));
  await page.click(td(20, 3));
  await page.keyboard.press("ControlOrMeta+Shift+z");
  await idle();
  await sleep(500);
  check("2. Ctrl+Y przywraca tabelę", JSON.stringify(await tables()) === JSON.stringify(t1), await tables());
  await page.click(td(20, 3));
  await page.keyboard.press("ControlOrMeta+z");
  await idle();
  await sleep(500);

  // 3. Tylko zaznaczenie
  await selectH8I13();
  await dropToasts();
  await page.click("#fmtTableBtn");
  await idle();
  await sleep(600);
  await page.click(".toast .toast-action-2");
  // najpierw Cofnij (wczytanie pliku), potem nowa tabela (drugie wczytanie) — czekamy na wynik
  await page.waitForFunction(() => Object.values(currentTables || {}).some((x) => x.ref === "H8:I13"), null, { timeout: 20000 }).catch(() => {});
  await idle();
  await sleep(300);
  const t3 = await tables();
  check("3. „Tylko zaznaczenie”: tabela dokładnie H8:I13", t3.length === before.length + 1 && t3.some((x) => /:H8:I13$/.test(x)), t3);
  await page.click(td(20, 3));
  await page.keyboard.press("ControlOrMeta+z");
  await idle();
  await sleep(500);
  check("3. jedno Cofnij usuwa tabelę z zaznaczenia", JSON.stringify(await tables()) === JSON.stringify(before), await tables());

  // 4. zmiana komórki po utworzeniu tabeli: „Cofnij” z dymka nie cofa cudzej zmiany (przegląd kodu)
  const cellVal = (r, c) => page.evaluate(([r, c]) => document.querySelector(`#dataTable tbody tr[data-row-key="wide:${r}"] td[data-col-index="${c}"]`)?.textContent.trim(), [r, c]);
  await selectH8I13();
  await dropToasts();
  await page.click("#fmtTableBtn");
  await idle();
  await sleep(600);
  const keepToast = await page.evaluate(() => { const el = [...document.querySelectorAll(".toast")].pop(); if (el) el.dataset.keep = "1"; return !!el; });
  await page.click(td(20, 3));
  await page.keyboard.type("Zmiana po tabeli"); // pisanie na zaznaczonej komórce zastępuje jej treść
  await page.keyboard.press("Enter");
  await sleep(300);
  if (keepToast) await page.click('.toast[data-keep="1"] .toast-action:not(.toast-action-2)');
  await idle();
  await sleep(400);
  check("4. „Cofnij” z dymka po innej zmianie: tabela i zmiana zostają", (await tables()).length === before.length + 1 && (await cellVal(20, 3)) === "Zmiana po tabeli", { tabs: await tables(), v: await cellVal(20, 3) });
  // Ctrl+Z ×2: najpierw zmiana komórki, potem tabela; Ctrl+Y ×2 przywraca obie (Ponów komórek po tabeli zostaje)
  await page.click(td(25, 3));
  await page.keyboard.press("ControlOrMeta+z");
  await sleep(300);
  await page.keyboard.press("ControlOrMeta+z");
  await idle();
  await sleep(500);
  await page.click(td(25, 3));
  await page.keyboard.press("ControlOrMeta+Shift+z");
  await idle();
  await sleep(500);
  await page.keyboard.press("ControlOrMeta+Shift+z");
  await sleep(400);
  check("4. Ponów po cofnięciu tabeli przywraca tabelę i zmianę komórki", (await tables()).length === before.length + 1 && (await cellVal(20, 3)) === "Zmiana po tabeli", { tabs: await tables(), v: await cellVal(20, 3) });

  check("brak błędów strony", errors.length === 0, errors.join(" | "));
  await browser.close();
  const failed = results.filter((r) => !r.ok);
  results.forEach((r) => console.log(`${r.ok ? "✓" : "✗"} ${r.name}${r.ok ? "" : "  — " + r.detail}`));
  if (failed.length) throw new Error(`${failed.length} z ${results.length} asercji nie przeszło (${ENGINE})`);
  console.log(`\n✅ table-undo-playwright OK — ${results.length} asercji (${ENGINE})`);
}

run().catch((e) => { console.error(e.message || e); process.exit(1); });
