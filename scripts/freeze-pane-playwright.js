// freeze-pane-playwright.js — „SZYBA" zamrożonej kolumny na czas przewijania (app/freeze-pane.js).
//
// Po co: przyklejona 1. kolumna to setki osobnych warstw (sticky na 2 komórkach w każdym
// wierszu) — przewijanie z blokadą kosztowało ~3× więcej. Na czas przewijania kładziemy
// JEDNĄ warstwę z kopią kolumny, w spoczynku wraca prawdziwe sticky.
//
// Co musi być prawdą:
//   1. szyba wygląda PIKSEL W PIKSEL jak prawdziwe przyklejone komórki (zrzut ekranu),
//      także z zaznaczonym wierszem i przy zawijaniu tekstu,
//   2. przewijanie włącza szybę, spoczynek ją zdejmuje i nie zostawia śladów (fz-off),
//   3. zmiana klasy zamrożonej komórki w trakcie przewijania trafia do kopii,
//   4. przebudowa tabeli w trakcie przewijania wyłącza szybę,
//   5. tap w szybę w trakcie przewijania nie trafia w kolumnę POD szybą,
//   6. bez blokady / ze scaleniami szyby nie ma.

const { chromium, webkit } = require("playwright");
const path = require("path");

const APP_URL = process.env.APP_URL || "http://127.0.0.1:4175/";
const FILE = path.join(__dirname, "stress-test-workbench.xlsx");
const ENGINE = process.env.ENGINE === "webkit" ? webkit : chromium;
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

async function run() {
  const browser = await ENGINE.launch({ headless: true });
  const MOBILE = process.env.TOUCH === "1";
  const context = await browser.newContext(MOBILE
    ? { serviceWorkers: "block", viewport: { width: 390, height: 844 }, deviceScaleFactor: 3, hasTouch: true, isMobile: ENGINE === chromium }
    : { serviceWorkers: "block", viewport: { width: 900, height: 800 } });
  await context.addInitScript(() => localStorage.setItem("introPlayed", "true"));
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push("pageerror: " + e.message));
  page.on("console", (m) => { if (m.type() === "error") errors.push("console: " + m.text()); });

  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  await page.evaluate(() => { try { ensureXlsxLibs && ensureXlsxLibs(false); } catch {} });
  await page.setInputFiles("#fileInput", FILE);
  await page.waitForFunction(() => document.getElementById("sheetSelect")?.options?.length > 0, null, { timeout: 20000 });
  await sleep(300);
  await page.click("#loadBtn");
  await sleep(1500);
  await page.evaluate(() => { if (typeof setSidebarOpen === "function") setSidebarOpen(false); if (typeof syncSidebarHandle === "function") syncSidebarHandle(); });
  await sleep(500);

  const results = [];
  const check = (name, ok, detail) => { results.push({ name, ok: !!ok, detail }); };
  const setCheckbox = (id, on) => page.evaluate(([id, on]) => {
    const el = document.getElementById(id);
    if (el.checked !== on) { el.checked = on; el.dispatchEvent(new Event("change", { bubbles: true })); }
  }, [id, on]);
  const state = () => page.evaluate(() => ({
    active: window.__freezePane.isActive(),
    pane: !!document.querySelector("#tableWrap > .fz-pane"),
    offCells: document.querySelectorAll("#dataTable td.fz-off").length,
    sticky: getComputedStyle(document.querySelector("#dataTable tbody td:nth-child(2)")).position,
  }));
  // Zrzut obszaru tabeli bez paska przewijania i FAB-ów.
  const shot = async () => {
    const box = await page.evaluate(() => { const r = document.getElementById("tableWrap").getBoundingClientRect(); return { x: r.left + 1, y: r.top + 1, width: Math.min(r.width - 2, 520), height: r.height - 2 }; });
    return page.screenshot({ clip: box, animations: "disabled", caret: "hide" });
  };
  const compare = async (label) => {
    await page.evaluate(() => window.__freezePane.stop());
    await sleep(80);
    const rest = await shot();
    const started = await page.evaluate(() => window.__freezePane.start());
    await sleep(80);
    const withPane = await shot();
    await page.evaluate(() => window.__freezePane.stop());
    if (process.env.SHOTS && !rest.equals(withPane)) { const fs = require("fs"); const n = label.replace(/\W+/g, "_"); fs.writeFileSync(`${process.env.SHOTS}/${n}-rest.png`, rest); fs.writeFileSync(`${process.env.SHOTS}/${n}-pane.png`, withPane); }
    // Porównanie z progiem: przesunięcie, brak tła, podwójny cień dają różnice rzędu
    // 100+ na kanale. Wygładzanie krawędzi i ~8% prześwitu półprzezroczystego
    // podświetlenia wiersza (istniejące zachowanie) to pojedyncze dziesiątki.
    const diff = started ? await page.evaluate(async ([a, b]) => {
      const load = (src) => new Promise((res) => { const i = new Image(); i.onload = () => res(i); i.src = src; });
      const [ia, ib] = await Promise.all([load(a), load(b)]);
      const px = (img) => { const c = document.createElement("canvas"); c.width = img.width; c.height = img.height; const x = c.getContext("2d"); x.drawImage(img, 0, 0); return x.getImageData(0, 0, c.width, c.height).data; };
      const da = px(ia), db = px(ib);
      let big = 0, max = 0;
      for (let i = 0; i < da.length; i += 4) {
        const d = Math.max(Math.abs(da[i] - db[i]), Math.abs(da[i + 1] - db[i + 1]), Math.abs(da[i + 2] - db[i + 2]));
        if (d > max) max = d;
        if (d > 16) big++;
      }
      return { big, max, total: da.length / 4 };
    }, ["data:image/png;base64," + rest.toString("base64"), "data:image/png;base64," + withPane.toString("base64")]) : null;
    const ok = !!diff && diff.max <= 64 && diff.big <= diff.total * 0.002;
    check(`${label}: szyba wygląda jak przyklejone komórki`, ok, diff || "szyba się nie włączyła");
  };

  await page.addStyleTag({ content: ".scroll-top-fab,.clear-sel-fab,.sheet-picker-fab,.cell-actions{visibility:hidden!important}" });
  await page.evaluate(() => { document.body.classList.add("cursor-hint-off"); });

  // 6a. bez blokady szyby nie ma
  await setCheckbox("freezeFirstCol", false);
  await page.evaluate(() => { const w = document.getElementById("tableWrap"); w.scrollTop = 200; w.dispatchEvent(new Event("scroll")); });
  check("bez blokady kolumny przewijanie nie włącza szyby", !(await state()).active);
  await sleep(300);

  await setCheckbox("freezeFirstCol", true);
  await setCheckbox("freezeHeaders", true);
  await sleep(200);
  await page.evaluate(() => { const w = document.getElementById("tableWrap"); w.scrollLeft = 420; w.scrollTop = 900; });
  await sleep(300);

  // 1. wierność pikselowa
  await compare("przewinięte w bok i w dół");
  await page.evaluate(() => { const w = document.getElementById("tableWrap"); w.scrollLeft = 0; });
  await sleep(250);
  await compare("scrollLeft = 0");
  await page.evaluate(() => { const w = document.getElementById("tableWrap"); w.scrollLeft = 420; });
  await sleep(250);
  // zaznaczony wiersz (klik w komórkę danych daleko od zamrożonej kolumny)
  await page.evaluate(() => {
    const td = [...document.querySelectorAll("#dataTable tbody tr")].find((tr) => { const r = tr.getBoundingClientRect(); const w = document.getElementById("tableWrap").getBoundingClientRect(); return r.top > w.top + 120; }).cells[6];
    td.dispatchEvent(new PointerEvent("pointerdown", { bubbles: true, clientX: 1, clientY: 1 }));
    td.dispatchEvent(new MouseEvent("mousedown", { bubbles: true }));
    td.dispatchEvent(new PointerEvent("pointerup", { bubbles: true }));
    td.dispatchEvent(new MouseEvent("mouseup", { bubbles: true }));
    td.click();
  });
  await sleep(250);
  await compare("z zaznaczonym wierszem");
  await setCheckbox("wrapCells", true).catch(() => {});
  await sleep(600);
  await compare("zawijanie tekstu");
  await setCheckbox("wrapCells", false).catch(() => {});
  await sleep(600);

  // 2. cykl: przewijanie włącza, spoczynek zdejmuje
  await page.evaluate(() => { const w = document.getElementById("tableWrap"); w.scrollTop += 60; });
  await sleep(40);
  const during = await state();
  check("przewijanie włącza szybę", during.active && during.pane && during.offCells > 0, during);
  check("w trakcie przewijania prawdziwe komórki nie są sticky", during.sticky === "relative", during.sticky);
  await sleep(400);
  const rest = await state();
  check("spoczynek zdejmuje szybę bez śladów", !rest.active && !rest.pane && rest.offCells === 0 && rest.sticky === "sticky", rest);

  // 3. synchronizacja zmian klas
  const synced = await page.evaluate(async () => {
    window.__freezePane.start();
    const tr = [...document.querySelectorAll("#dataTable tbody tr")][30];
    tr.cells[1].classList.add("cell-selected");
    await new Promise((r) => setTimeout(r, 0));
    const clone = document.querySelector(".fz-pane tbody").rows[30].cells[1];
    const ok = clone.classList.contains("cell-selected");
    tr.cells[1].classList.remove("cell-selected");
    await new Promise((r) => setTimeout(r, 0));
    const ok2 = !document.querySelector(".fz-pane tbody").rows[30].cells[1].classList.contains("cell-selected");
    window.__freezePane.stop();
    return ok && ok2;
  });
  check("zmiana klasy zamrożonej komórki trafia do kopii", synced);

  // 4. przebudowa tabeli wyłącza szybę
  const rebuilt = await page.evaluate(async () => {
    window.__freezePane.start();
    renderActiveTable();
    await new Promise((r) => setTimeout(r, 0));
    return { active: window.__freezePane.isActive(), pane: !!document.querySelector(".fz-pane"), off: document.querySelectorAll("td.fz-off").length };
  });
  check("przebudowa tabeli w trakcie przewijania wyłącza szybę", !rebuilt.active && !rebuilt.pane && rebuilt.off === 0, rebuilt);
  await sleep(300);

  // 5. tap w szybę w trakcie przewijania nie trafia w kolumnę pod spodem
  await page.evaluate(() => { const w = document.getElementById("tableWrap"); w.scrollLeft = 600; });
  await sleep(300);
  const tapped = await page.evaluate(async () => {
    const w = document.getElementById("tableWrap");
    w.scrollTop += 40;
    await new Promise((r) => requestAnimationFrame(() => r()));
    const pane = document.querySelector(".fz-pane");
    if (!pane) return { err: "brak szyby" };
    const row = [...document.querySelectorAll("#dataTable tbody tr")].find((tr) => tr.getBoundingClientRect().top > w.getBoundingClientRect().top + 100);
    const y = row.getBoundingClientRect().top + 8;
    const x = pane.getBoundingClientRect().left + 30;
    const under = document.elementFromPoint(x, y);
    const before = document.querySelectorAll("#dataTable .cell-selected, #dataTable tr.row-selected").length + "|" + (document.querySelector("#dataTable .cell-selected, #dataTable tr.row-selected")?.textContent || "");
    for (const t of ["pointerdown", "mousedown"]) under.dispatchEvent(new (t.startsWith("pointer") ? PointerEvent : MouseEvent)(t, { bubbles: true, cancelable: true, clientX: x, clientY: y }));
    for (const t of ["pointerup", "mouseup", "click"]) under.dispatchEvent(new (t.startsWith("pointer") ? PointerEvent : MouseEvent)(t, { bubbles: true, cancelable: true, clientX: x, clientY: y }));
    const after = document.querySelectorAll("#dataTable .cell-selected, #dataTable tr.row-selected").length + "|" + (document.querySelector("#dataTable .cell-selected, #dataTable tr.row-selected")?.textContent || "");
    return { underCol: under.closest("td")?.cellIndex, before, after };
  });
  check("tap w szybę w trakcie przewijania nie zaznacza kolumny pod szybą", tapped.underCol > 1 && tapped.before === tapped.after, tapped);
  await sleep(500);
  const tapAfter = await page.evaluate(() => {
    const row = [...document.querySelectorAll("#dataTable tbody tr")][12];
    row.cells[1].click();
    return document.querySelector("#dataTable tbody td.cell-selected, #dataTable tbody tr.row-selected, #dataTable tbody tr.selected") ? true : document.activeElement?.tagName;
  });
  check("po spoczynku klik w zamrożoną kolumnę działa jak zawsze", tapAfter === true || tapAfter === "TD", tapAfter);

  // 6b. scalenia → brak szyby
  const merged = await page.evaluate(async () => {
    const td = document.querySelector("#dataTable tbody tr:nth-child(3) td:nth-child(4)");
    td.colSpan = 2;
    const ok = !window.__freezePane.start();
    td.colSpan = 1;
    window.__freezePane.stop();
    return ok;
  });
  check("scalone komórki w danych → bez szyby (zachowanie jak dawniej)", merged);

  await browser.close();
  let failed = 0;
  for (const r of results) {
    console.log(`${r.ok ? "✅" : "❌"} ${r.name}${r.detail !== undefined ? "  " + JSON.stringify(r.detail) : ""}`);
    if (!r.ok) failed++;
  }
  if (errors.length) { console.log("❌ błędy konsoli:", errors.slice(0, 5)); failed++; }
  if (failed) { console.log(`❌ ${failed} niezaliczone`); process.exit(1); }
  console.log("✅ wszystko przeszło");
}
run().catch((e) => { console.error(e); process.exit(1); });
