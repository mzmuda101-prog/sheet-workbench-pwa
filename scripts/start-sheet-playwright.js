// start-sheet-playwright.js — „Ustaw jako arkusz startowy” (Mateusz 2026-11-20): Excel otwiera plik na wybranym
// arkuszu. Wyrocznia Excela (2026-11-20): activeTab w xl/workbook.xml + tabSelected tylko na tym arkuszu →
// Excel pokazuje go po otwarciu.
//
//   1. komputer: prawy klik na karcie arkusza → menu z „Ustaw jako arkusz startowy” (dla TEGO arkusza),
//   2. po wyborze: w pliku activeTab = ten arkusz, tabSelected tylko na nim; w apce zostaje bieżący arkusz,
//   3. drugi raz na tym samym: komunikat „już jest arkuszem startowym”,
//   4. lista arkuszy (menu pod nazwą) też ma tę pozycję,
//   5. telefon: przytrzymanie karty otwiera menu, a puszczenie palca nie przełącza arkusza.
//
// Uruchom z serwerem na APP_URL (domyślnie http://127.0.0.1:4175/). ENGINE=webkit — Safari.
// OUT=plik.xlsx — zapisuje wynik (do sprawdzenia w prawdziwym Excelu).

const playwright = require("playwright");
const fs = require("fs");

const APP_URL = process.env.APP_URL || "http://127.0.0.1:4175/";
const ENGINE = process.env.ENGINE || "chromium";
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

async function open(browser, viewport, extra = {}) {
  const context = await browser.newContext({ serviceWorkers: "block", viewport, ...extra });
  await context.addInitScript(() => {
    if (!sessionStorage.getItem("__ss")) { sessionStorage.setItem("__ss", "1"); localStorage.clear(); localStorage.setItem("introPlayed", "true"); }
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
const idle = (page) => page.waitForFunction(() => !(typeof swbSheetOpsUi !== "undefined" && swbSheetOpsUi.isBusy()) && document.body.getAttribute("aria-busy") !== "true", null, { timeout: 20000 });
// stan pliku: activeTab i zaznaczone karty (z bieżących bajtów, jak przy zapisie)
const fileState = (page) => page.evaluate(async () => {
  const bytes = await buildPatchedXlsx(originalFileBytes, pendingEdits, buildSheetFormulaMaps(), { layout: pendingLayout });
  const zip = await JSZip.loadAsync(bytes);
  const wb = await zip.file("xl/workbook.xml").async("string");
  const active = Number((/activeTab="(\d+)"/.exec(wb) || [0, 0])[1]);
  const names = [...wb.matchAll(/<sheet\b[^>]*\bname="([^"]*)"/g)].map((m) => m[1]);
  const selected = [];
  for (let i = 1; i <= names.length; i++) {
    const x = await zip.file(`xl/worksheets/sheet${i}.xml`)?.async("string");
    if (x && /tabSelected="1"/.test(x)) selected.push(i - 1);
  }
  return { active, names, selected, current: currentSheetName, b64: btoa(String.fromCharCode(...bytes)) };
});

async function run() {
  const browser = await playwright[ENGINE].launch({ headless: true });
  const results = [];
  const check = (name, ok, detail = "") => results.push({ name, ok: !!ok, detail: typeof detail === "string" ? detail : JSON.stringify(detail) });

  {
    const { context, page, errors } = await open(browser, { width: 1280, height: 800 });
    const before = await fileState(page);
    const target = before.names[before.names.length - 1];
    await page.click(`#sheetTabs .sheet-tab[data-sheet="${target}"]`, { button: "right" });
    await page.waitForSelector('#sheetMenu:not([hidden]) [data-sheet-op="start"]', { timeout: 5000 }).catch(() => {});
    const menu = await page.evaluate(() => ({ open: !document.getElementById("sheetMenu").hidden, cap: document.querySelector("#sheetMenu .app-menu-cap")?.textContent, ops: [...document.querySelectorAll("#sheetMenu [data-sheet-op]")].map((b) => b.dataset.sheetOp) }));
    check("1. prawy klik na karcie: menu tego arkusza z „Ustaw jako arkusz startowy”", menu.open && menu.cap === target && menu.ops.includes("start") && menu.ops.includes("rename"), menu);
    await page.click('#sheetMenu [data-sheet-op="start"]');
    await idle(page);
    await sleep(500);
    const after = await fileState(page);
    const ti = after.names.indexOf(target);
    check("2. w pliku activeTab = wybrany arkusz", after.active === ti, { after: { active: after.active, names: after.names }, ti });
    check("2. tabSelected tylko na nim (bez grupowania arkuszy)", JSON.stringify(after.selected) === JSON.stringify([ti]), after.selected);
    check("2. w apce zostaje bieżący arkusz", after.current === before.current, { before: before.current, after: after.current });
    if (process.env.OUT) fs.writeFileSync(process.env.OUT, Buffer.from(after.b64, "base64"));
    await page.evaluate(() => document.querySelectorAll(".toast").forEach((t) => t.remove()));
    await page.click(`#sheetTabs .sheet-tab[data-sheet="${target}"]`, { button: "right" });
    await page.click('#sheetMenu [data-sheet-op="start"]');
    await sleep(600);
    const again = await page.evaluate(() => [...document.querySelectorAll(".toast")].map((x) => x.textContent).join("|"));
    check("3. drugi raz: „już jest arkuszem startowym”", /już jest arkuszem startowym/.test(again), again);
    // 4. lista arkuszy pod nazwą
    await page.click("#heroSheetBtn");
    await sleep(300);
    check("4. lista arkuszy ma „Ustaw jako arkusz startowy”", await page.evaluate(() => !!document.querySelector('#sheetMenu:not([hidden]) [data-sheet-op="start"]')));
    await page.keyboard.press("Escape");
    check("brak błędów strony (komputer)", errors.length === 0, errors.join(" | "));
    await context.close();
  }

  {
    const { context, page, errors } = await open(browser, { width: 390, height: 844 }, { hasTouch: true, isMobile: ENGINE === "chromium" });
    const start = await page.evaluate(() => currentSheetName);
    const tab = await page.evaluate(() => { const t = [...document.querySelectorAll("#sheetTabs .sheet-tab[data-sheet]")].find((x) => x.dataset.sheet !== currentSheetName); t?.scrollIntoView({ inline: "center" }); const r = t?.getBoundingClientRect(); return t ? { name: t.dataset.sheet, x: r.left + r.width / 2, y: r.top + r.height / 2 } : null; });
    // przytrzymanie: pointerdown dotykiem, 0,7 s, puszczenie
    await page.evaluate(async ({ x, y }) => {
      const el = document.elementFromPoint(x, y);
      const opt = { bubbles: true, cancelable: true, pointerType: "touch", clientX: x, clientY: y, isPrimary: true };
      el.dispatchEvent(new PointerEvent("pointerdown", opt));
      await new Promise((r) => setTimeout(r, 700));
      el.dispatchEvent(new PointerEvent("pointerup", opt));
      el.dispatchEvent(new MouseEvent("click", { bubbles: true, cancelable: true, clientX: x, clientY: y }));
    }, tab);
    await sleep(400);
    const st = await page.evaluate(() => ({ open: !document.getElementById("sheetMenu").hidden, cap: document.querySelector("#sheetMenu .app-menu-cap")?.textContent, current: currentSheetName }));
    check("5. telefon: przytrzymanie karty otwiera jej menu", st.open && st.cap === tab.name, st);
    check("5. telefon: puszczenie palca nie przełącza arkusza", st.current === start, { st, start });
    check("brak błędów strony (telefon)", errors.length === 0, errors.join(" | "));
    await context.close();
  }

  await browser.close();
  const failed = results.filter((r) => !r.ok);
  results.forEach((r) => console.log(`${r.ok ? "✓" : "✗"} ${r.name}${r.ok ? "" : "  — " + r.detail}`));
  if (failed.length) throw new Error(`${failed.length} z ${results.length} asercji nie przeszło (${ENGINE})`);
  console.log(`\n✅ start-sheet-playwright OK — ${results.length} asercji (${ENGINE})`);
}

run().catch((e) => { console.error(e.message || e); process.exit(1); });
