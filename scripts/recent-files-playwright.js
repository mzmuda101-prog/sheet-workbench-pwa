// recent-files-playwright.js — „Ostatnio otwierane” na ekranie startowym (app/recent-files.js).
//   node scripts/recent-files-playwright.js            (Chromium)
//   ENGINE=webkit node scripts/recent-files-playwright.js
//
// Pilnuje:
//   1. pusta lista = brak sekcji; przykładowy plik NIE trafia na listę;
//   2. otwarty plik po przeładowaniu jest na liście (nazwa, data, rozmiar) i otwiera się z kopii
//      — tak samo w Safari (WebKit), gdzie nie ma uchwytu pliku;
//   3. ten sam plik otwarty drugi raz nie dubluje wpisu; najwyżej 5 pozycji, najnowsza u góry;
//   4. zapis (pobranie) podmienia kopię na zapisaną wersję — nazwa i treść, bez nowego wpisu;
//   5. ✕ usuwa jedną pozycję; „Wyczyść listę” pyta we własnym oknie i czyści wszystko;
//   6. zmiana języka przerysowuje nagłówek.

const { chromium, webkit } = require("playwright");
const fs = require("fs");
const path = require("path");

const APP_URL = process.env.APP_URL || "http://127.0.0.1:4175/";
const ENGINE_NAME = process.env.ENGINE === "webkit" ? "webkit" : "chromium";
const ENGINE = ENGINE_NAME === "webkit" ? webkit : chromium;
const FILE = path.join(__dirname, "stress-test-workbench.xlsx");
const BYTES = fs.readFileSync(FILE);
const MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

async function run() {
  const fails = [];
  const check = (name, ok, detail) => {
    console.log(`${ok ? "✅" : "❌"} ${name}${ok ? "" : ` — ${JSON.stringify(detail)}`}`);
    if (!ok) fails.push(name);
  };

  const browser = await ENGINE.launch({ headless: true });
  const ctx = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1280, height: 900 }, acceptDownloads: true });
  await ctx.addInitScript(() => { localStorage.setItem("introPlayed", "true"); localStorage.setItem("swb-lang", localStorage.getItem("swb-lang") || "pl"); });
  const page = await ctx.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  page.on("dialog", (d) => { errors.push(`natywne okno: ${d.message()}`); d.dismiss(); });

  const fresh = async () => {
    await page.goto(APP_URL, { waitUntil: "load" });
    await page.evaluate(() => document.getElementById("heroSplash")?.remove());
    await page.waitForFunction(() => typeof swbRecent !== "undefined");
    await sleep(300);
  };
  const items = () => page.evaluate(() => {
    const host = document.getElementById("recentFiles");
    return {
      visible: !!host && !host.hidden && host.offsetParent !== null,
      cap: host?.querySelector(".recent-cap")?.textContent || "",
      rows: [...(host?.querySelectorAll(".recent-item") || [])].map((li) => ({
        name: li.querySelector(".recent-name").textContent,
        meta: li.querySelector(".recent-meta").textContent,
      })),
    };
  });
  const openFile = async (name, buffer = BYTES) => {
    await page.setInputFiles("#fileInput", { name, mimeType: MIME, buffer });
    await page.waitForFunction((n) => typeof workbook !== "undefined" && workbook && currentFileName === n, name, { timeout: 20000 });
    await page.waitForFunction(async () => (await swbRecent._list()).length > 0 && !!swbRecent._current(), null, { timeout: 5000 }).catch(() => {});
    await sleep(200);
  };

  await fresh();
  await page.evaluate(() => indexedDB.deleteDatabase("swb-recent"));
  await fresh();

  // 1. pusto + przykład pomijany
  check("1. pusta lista = brak sekcji", !(await items()).visible, await items());
  await page.evaluate(() => loadSampleFile());
  await page.waitForFunction(() => typeof workbook !== "undefined" && workbook, null, { timeout: 20000 });
  await sleep(500);
  check("1. przykładowy plik nie trafia na listę", (await page.evaluate(() => swbRecent._list())).length === 0);

  // 2. plik na liście po przeładowaniu, otwiera się z kopii
  await fresh();
  await openFile("Rejestr-2026.xlsx");
  await fresh();
  const l2 = await items();
  check("2. po przeładowaniu plik jest na liście", l2.visible && l2.rows.length === 1 && l2.rows[0].name === "Rejestr-2026.xlsx", l2);
  check("2. wiersz: kiedy + rozmiar", /dziś o \d\d:\d\d · \d+(,\d)? (KB|MB)/.test(l2.rows[0]?.meta || ""), l2.rows[0]);
  check("2. nagłówek i notatka po polsku", l2.cap === "Ostatnio otwierane", l2.cap);
  await page.click("#recentFiles .recent-open");
  await page.waitForFunction(() => typeof workbook !== "undefined" && workbook && currentFileName === "Rejestr-2026.xlsx", null, { timeout: 20000 });
  const opened = await page.evaluate(() => ({ sheets: workbook.SheetNames.length, handle: !!currentFileHandle, hidden: document.getElementById("recentFiles").hidden }));
  check("2. klik otwiera plik z kopii (bez uchwytu)", opened.sheets > 0 && !opened.handle, opened);
  check("2. przy otwartym pliku listy nie widać", opened.hidden, opened);
  check("3. ponowne otwarcie nie dubluje wpisu", (await page.evaluate(() => swbRecent._list())).length === 1);

  // 4. zapis (pobranie) = kopia na liście to zapisana wersja
  const before = await page.evaluate(async () => (await swbRecent._list())[0]);
  const [dl] = await Promise.all([
    page.waitForEvent("download", { timeout: 15000 }),
    page.evaluate(() => downloadWorkbook("Rejestr-2026_edited.xlsx", "xlsx")),
  ]);
  const savedPath = await dl.path();
  await sleep(400);
  const after = await page.evaluate(async () => {
    const recs = await swbRecent._list();
    return { n: recs.length, name: recs[0]?.name, size: recs[0]?.size, id: recs[0]?.id };
  });
  const savedSize = savedPath ? fs.statSync(savedPath).size : -1;
  check("4. po zapisie: ta sama pozycja, nowa nazwa i rozmiar zapisanego pliku",
    after.n === 1 && after.id === before.id && after.name === "Rejestr-2026_edited.xlsx" && after.size === savedSize, { before: { id: before.id, size: before.size }, after, savedSize });

  // 3. najwyżej 5, najnowszy u góry
  await fresh();
  for (let i = 1; i <= 6; i++) { await openFile(`plik-${i}.xlsx`); }
  await fresh();
  const l3 = await items();
  check("3. najwyżej 5 pozycji, najnowsza u góry", l3.rows.length === 5 && l3.rows[0].name === "plik-6.xlsx" && !l3.rows.some((r) => r.name === "plik-1.xlsx"), l3.rows.map((r) => r.name));

  // 6. język
  await page.evaluate(() => applyLanguage("en"));
  await sleep(200);
  check("6. zmiana języka przerysowuje nagłówek", (await items()).cap === "Recently opened", (await items()).cap);
  await page.evaluate(() => applyLanguage("pl"));
  await sleep(200);

  // 5. ✕ i „Wyczyść listę”
  await page.click("#recentFiles .recent-item:first-child .recent-remove");
  await sleep(300);
  const l5 = await items();
  check("5. ✕ usuwa jedną pozycję", l5.rows.length === 4 && l5.rows[0].name === "plik-5.xlsx", l5.rows.map((r) => r.name));
  await page.click("#recentFiles .recent-clear");
  // okno z app-dialog.js: nakładka role="dialog" jest position: fixed (offsetParent = null) — sprawdzamy jej pudełko
  await page.waitForSelector(".app-dialog-box", { state: "visible", timeout: 3000 }).catch(() => {});
  const dlg = await page.evaluate(() => [...document.querySelectorAll(".app-dialog-box")].filter((d) => d.offsetParent !== null).map((d) => d.textContent).join(" | "));
  check("5. „Wyczyść listę” pyta we własnym oknie", /Wyczyścić listę ostatnich plików/.test(dlg), dlg);
  await page.evaluate(() => {
    const btn = [...document.querySelectorAll(".app-dialog-box button")].find((b) => b.offsetParent !== null && b.textContent.trim() === "Wyczyść listę");
    btn?.click();
  });
  await sleep(400);
  const l6 = await items();
  check("5. po potwierdzeniu lista pusta i schowana", !l6.visible && (await page.evaluate(() => swbRecent._list())).length === 0, l6);

  await browser.close();
  if (errors.length) { console.log("❌ błędy:", errors.slice(0, 5)); fails.push("błędy strony"); }
  if (fails.length) { console.log(`❌ recent-files (${ENGINE_NAME}): ${fails.length} niezaliczone`); process.exit(1); }
  console.log(`✅ recent-files (${ENGINE_NAME}): wszystko przeszło`);
}
run().catch((e) => { console.error(e); process.exit(1); });
