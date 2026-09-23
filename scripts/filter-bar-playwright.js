// filter-bar-playwright.js — pasek „co teraz filtruje" + menu komórki („Pokaż tylko takie").
//
// Co musi być prawdą (liczby porównujemy z NIEZALEŻNYM przeliczeniem w teście):
//   1. bez filtrów pasek jest ukryty,
//   2. prawy klik na komórce → menu z nazwą kolumny i wartością,
//   3. „Pokaż tylko takie" = wiersze, w których CAŁA komórka tej kolumny równa się wartości,
//      a zapytanie w szybkim szukaniu to czytelne „Kolumna:="wartość"",
//   4. „Ukryj takie" w innej kolumnie DOKLEJA się przez && (zawężanie jak w Accessie),
//   5. pasek pokazuje „Widać X z Y" i pigułkę; × zdejmuje filtr, „Wyczyść" zdejmuje wszystko,
//   6. pusta komórka → „Pokaż tylko puste",
//   7. Shift+F10 otwiera menu z klawiatury (fokus na 1. pozycji), Esc zamyka i oddaje fokus,
//   8. przytrzymanie palcem otwiera menu i NIE otwiera edytora komórki (klik połknięty),
//   9. w konsoli brak błędów.
//
// Uruchom z serwerem na APP_URL (domyślnie http://127.0.0.1:4175/).

const { chromium, webkit } = require("playwright");
const path = require("path");

const APP_URL = process.env.APP_URL || "http://127.0.0.1:4175/";
const FILE = path.join(__dirname, "stress-test-workbench.xlsx");
const ENGINE = process.env.ENGINE === "webkit" ? webkit : chromium;
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

async function run() {
  const browser = await ENGINE.launch({ headless: true });
  const context = await browser.newContext({
    serviceWorkers: "block",
    viewport: { width: 1280, height: 900 },
    hasTouch: true, // potrzebne do symulacji przytrzymania; mysz dalej działa
  });
  await context.addInitScript(() => localStorage.setItem("introPlayed", "true"));
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push("pageerror: " + e.message));
  page.on("console", (m) => { if (m.type() === "error") errors.push("console: " + m.text()); });

  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  await page.evaluate(() => { try { ensureXlsxLibs && ensureXlsxLibs(false); } catch {} });
  await page.setInputFiles("#fileInput", FILE);
  await page.waitForFunction(() => document.getElementById("sheetSelect")?.options?.length > 0, null, { timeout: 15000 });
  await sleep(300);
  await page.click("#loadBtn");
  await sleep(1200);
  await page.evaluate(() => { if (typeof setSidebarOpen === "function") setSidebarOpen(false); });
  await sleep(300);

  const failures = [];
  const report = {};
  const ok = (name, cond, got) => { if (!cond) failures.push(`${name} — dostałem: ${JSON.stringify(got)}`); };

  const stripHidden = () => page.evaluate(() => document.getElementById("activeFilters").classList.contains("hidden"));
  ok("1. bez filtrów pasek ukryty", await stripHidden(), "widoczny");

  // Niezależne liczenie: ile wierszy ma w kolumnie `col` dokładnie `val` (trim, bez wielkości liter).
  const countEq = (conds) => page.evaluate((conds) => baseRows.filter((r) => conds.every(({ col, val, not }) => {
    const idx = currentHeaders.map((h, i) => (String(h).trim().toLowerCase() === String(col).trim().toLowerCase() ? i : -1)).filter((i) => i >= 0);
    const hit = idx.some((i) => String(getDisplayValue(r, i) ?? "").trim().toLowerCase() === String(val).trim().toLowerCase());
    return not ? !hit : hit;
  })).length, conds);

  // Komórka z niepustą wartością w kolumnie 0 (pierwszy widoczny wiersz z danymi).
  const pickCell = (col, wantEmpty = false) => page.evaluate(({ col, wantEmpty }) => {
    const tds = Array.from(document.querySelectorAll(`#tbody tr[data-row-key] td[data-col-index="${col}"], tbody tr[data-row-key] td[data-col-index="${col}"]`));
    for (const td of tds) {
      const tr = td.parentElement;
      const row = currentDisplayModel.rows.find((r) => getRowSelectionKey(r) === tr.dataset.rowKey);
      if (!row) continue;
      const v = String(getDisplayValue(row, col) ?? "").trim();
      if (wantEmpty ? v === "" : v !== "") {
        td.scrollIntoView({ block: "center", inline: "center" });
        const r = td.getBoundingClientRect();
        return { rowKey: tr.dataset.rowKey, header: currentDisplayModel.headers[col], value: v, x: r.left + r.width / 2, y: r.top + r.height / 2 };
      }
    }
    return null;
  }, { col, wantEmpty });

  // 2. prawy klik
  const c0 = await pickCell(0);
  report.c0 = c0;
  ok("jest komórka w kolumnie 0", !!c0, c0);
  if (c0) {
    await page.mouse.click(c0.x, c0.y, { button: "right" });
    await sleep(200);
    const menu = await page.evaluate(() => ({
      visible: !document.getElementById("cellMenu").classList.contains("hidden"),
      head: document.querySelector("#cellMenu .cell-menu-head")?.textContent || "",
      items: [...document.querySelectorAll("#cellMenu .cell-menu-item")].map((b) => b.textContent),
    }));
    report.menu = menu;
    ok("2. prawy klik otwiera menu z kolumną i wartością", menu.visible && menu.head.includes(c0.value.slice(0, 10)) && menu.items.length === 3, menu);

    // 3. Pokaż tylko takie
    await page.click("#cellMenu .cmi-show");
    await sleep(400);
    const after = await page.evaluate(() => ({ n: viewRows.length, q: quickSearchEl.value, ops: filterOperatorsEl.checked }));
    const exp = await countEq([{ col: c0.header, val: c0.value }]);
    report.showOnly = { ...after, expected: exp };
    ok("3. Pokaż tylko takie = dokładna równość w kolumnie", after.n === exp && exp > 0, report.showOnly);
    ok("3. zapytanie jest czytelne: Kolumna:=\"wartość\"", after.q === `${c0.header}:="${c0.value}"` && after.ops, after.q);

    // 5a. pasek widoczny z liczbą
    const strip = await page.evaluate(() => ({
      hidden: document.getElementById("activeFilters").classList.contains("hidden"),
      text: document.getElementById("activeFilters").textContent,
      chips: document.querySelectorAll("#activeFilters .af-chip").length,
    }));
    report.strip = strip;
    ok("5. pasek pokazuje „Widać X z Y” i pigułkę", !strip.hidden && strip.chips === 1 && strip.text.includes(String(exp)), strip);

    // 4. Ukryj takie w kolumnie 1 (wśród już przefiltrowanych) → doklejone &&
    const c1 = await pickCell(1);
    report.c1 = c1;
    if (c1) {
      await page.mouse.click(c1.x, c1.y, { button: "right" });
      await sleep(200);
      await page.click("#cellMenu .cmi-hide");
      await sleep(400);
      const both = await page.evaluate(() => ({ n: viewRows.length, q: quickSearchEl.value }));
      const exp2 = await countEq([{ col: c0.header, val: c0.value }, { col: c1.header, val: c1.value, not: true }]);
      report.combined = { ...both, expected: exp2 };
      ok("4. Ukryj takie dokleja się przez &&", both.n === exp2 && both.q.includes("&&") && both.q.includes(`${c1.header}:!="${c1.value}"`), report.combined);
    }

    // 5b. × zdejmuje filtr
    await page.click("#activeFilters .af-chip .af-remove");
    await sleep(400);
    const cleared = await page.evaluate(() => ({ n: viewRows.length, total: baseRows.length, q: quickSearchEl.value }));
    ok("5. × zdejmuje filtr szukania", cleared.n === cleared.total && cleared.q === "" && await stripHidden(), cleared);
  }

  // 6. pusta komórka
  let emptyCol = -1;
  let ce = null;
  for (let c = 0; c < 12 && !ce; c++) { ce = await pickCell(c, true); if (ce) emptyCol = c; }
  report.empty = ce;
  if (ce) {
    await page.mouse.click(ce.x, ce.y, { button: "right" });
    await sleep(200);
    const label = await page.evaluate(() => document.querySelector("#cellMenu .cmi-show")?.textContent);
    await page.click("#cellMenu .cmi-show");
    await sleep(400);
    const n = await page.evaluate(() => viewRows.length);
    const exp = await countEq([{ col: ce.header, val: "" }]);
    ok("6. pusta komórka → „Pokaż tylko puste”", /pust|empty/i.test(label || "") && n === exp && exp > 0, { label, n, exp, emptyCol });
    // „Wyczyść" = reset wszystkich filtrów
    await page.click("#activeFilters .af-clear-all");
    await sleep(400);
    const all = await page.evaluate(() => viewRows.length === baseRows.length);
    ok("5. „Wyczyść” zdejmuje wszystko", all && await stripHidden(), all);
  }

  // 7. klawiatura: Shift+F10
  const ck = await pickCell(2);
  if (ck) {
    await page.mouse.click(ck.x, ck.y);
    await sleep(200);
    await page.keyboard.press("Shift+F10");
    await sleep(200);
    const kb = await page.evaluate(() => ({
      visible: !document.getElementById("cellMenu").classList.contains("hidden"),
      focused: document.activeElement?.classList?.contains("cmi-show"),
    }));
    await page.keyboard.press("ArrowDown");
    const second = await page.evaluate(() => document.activeElement?.classList?.contains("cmi-hide"));
    await page.keyboard.press("Escape");
    await sleep(150);
    const back = await page.evaluate(() => ({
      closed: document.getElementById("cellMenu").classList.contains("hidden"),
      inGrid: !!document.activeElement?.closest?.("td[data-col-index]"),
    }));
    report.keyboard = { kb, second, back };
    ok("7. Shift+F10 → menu z fokusem, ↓ chodzi, Esc zamyka i oddaje fokus", kb.visible && kb.focused && second && back.closed && back.inGrid, report.keyboard);
  }

  // 8. przytrzymanie palcem (TouchEvent) — menu tak, edytor nie
  const ct = await pickCell(3);
  if (ct) {
    // najpierw zaznacz komórkę, żeby następny tap byłby „drugim tapem" (edycja na dotyku)
    await page.mouse.click(ct.x, ct.y);
    await sleep(150);
    await page.evaluate(({ x, y }) => {
      // Zwykły Event z polem `touches` — desktopowy WebKit nie ma konstruktora Touch,
      // a handler czyta tylko target i touches[0].clientX/Y.
      const el = document.elementFromPoint(x, y);
      const fake = (type, touches) => {
        const ev = new Event(type, { bubbles: true, cancelable: true });
        Object.defineProperty(ev, "touches", { value: touches });
        return ev;
      };
      el.dispatchEvent(fake("touchstart", [{ clientX: x, clientY: y }]));
      window.__lpEl = el; window.__lpFake = fake;
    }, ct);
    await sleep(700);
    const lp = await page.evaluate(() => {
      const el = window.__lpEl;
      el.dispatchEvent(window.__lpFake("touchend", []));
      el.click(); // syntetyczny klik po puszczeniu palca
      return {
        menu: !document.getElementById("cellMenu").classList.contains("hidden"),
        editor: !!document.querySelector("input.cell-editor"),
      };
    });
    report.longPress = lp;
    ok("8. przytrzymanie otwiera menu i nie otwiera edytora", lp.menu && !lp.editor, lp);
    await page.keyboard.press("Escape");
  }

  ok("9. brak błędów w konsoli", errors.length === 0, errors);
  await browser.close();

  console.log(JSON.stringify(report, null, 2));
  if (failures.length) {
    console.log("\n❌ " + failures.length + " błędów:\n - " + failures.join("\n - "));
    process.exit(1);
  }
  console.log("\n✅ filter-bar: wszystko OK");
}

run().catch((e) => { console.error(e); process.exit(1); });
