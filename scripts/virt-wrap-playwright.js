// virt-wrap-playwright.js — nowy silnik tabeli przy „Zawijaj tekst we wszystkich komórkach”.
//
// 2026-11-02, krok 2 planu „nowy silnik domyślnie”: dotąd zawijanie oddawało arkusz staremu silnikowi
// (z limitem wierszy). Teraz silnik zostaje i mierzy wysokości zawiniętych wierszy (pamięć wysokości
// zależy od zawijania i powiększenia — klucz środowiska), a po pomiarze dopełnia okno do pełnego ekranu.
//   1. włączenie zawijania: silnik dalej działa, wiersze mają różne wysokości,
//   2. 41 pozycji przewijania: bez dziur, kolumny równe z nagłówkiem, kolejne wiersze bez przeskoków,
//   3. dół osiągalny (ostatni wiersz przy dolnej krawędzi),
//   4. wyłączenie zawijania: wiersze wracają do jednej linii, dalej bez dziur.
// Fixture: .venv/bin/python scripts/fixtures/make-virt-merges.py (arkusz „Opisy”).

const { chromium, webkit } = require("playwright");
const path = require("path");

const APP_URL = process.env.APP_URL || "http://127.0.0.1:4175/";
const ENGINE = process.env.ENGINE === "webkit" ? webkit : chromium;
const FILE = path.join(__dirname, "fixtures", "virt-merges.xlsx");

async function run() {
  const browser = await ENGINE.launch({ headless: true });
  const ctx = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1200, height: 800 } });
  await ctx.addInitScript(() => localStorage.setItem("introPlayed", "true"));
  const page = await ctx.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  await page.goto(APP_URL + "?virt=1", { waitUntil: "load" });
  await page.evaluate(() => { document.getElementById("heroSplash")?.remove(); try { ensureXlsxLibs(false); } catch {} });
  await page.setInputFiles("#fileInput", FILE);
  await page.waitForFunction(() => typeof baseRows !== "undefined" && baseRows.length > 0, null, { timeout: 30000 });
  await page.evaluate(() => { sheetSelect.value = "Opisy"; loadBtn.click(); });
  await page.waitForFunction(() => currentSheetName === "Opisy" && baseRows.length > 300, null, { timeout: 20000 });
  await page.waitForTimeout(800);

  const failures = [];
  const ok = (name, cond, got) => { if (!cond) failures.push(`${name} — dostałem: ${JSON.stringify(got)}`); };

  const r = await page.evaluate(async () => {
    const fr = () => new Promise((res) => requestAnimationFrame(() => requestAnimationFrame(res)));
    // nagłówki czytane za każdym razem — przełączenie zawijania przerysowuje tabelę
    const colLeft = (i) => { const th = document.querySelector(`thead tr.header-row th[data-col-index="${i}"]`); return th ? th.getBoundingClientRect().left : null; };
    const check = () => {
      const wb = tableWrapEl.getBoundingClientRect();
      const head = document.querySelector("thead tr.header-row th").getBoundingClientRect();
      let holes = 0; const idx = [];
      for (let y = Math.max(wb.top, head.bottom) + 4; y < wb.bottom - 4; y += 8) {
        const el = document.elementFromPoint(wb.left + 60, y);
        const tr = el && el.closest("#dataTable tbody tr[data-row-key]");
        if (!tr) holes++; else idx.push(+tr.dataset.rowIndex);
      }
      const seq = idx.filter((v, i) => i === 0 || v !== idx[i - 1]);
      let jumps = 0;
      for (let i = 1; i < seq.length; i++) if (seq[i] !== seq[i - 1] + 1) jumps++;
      let misaligned = 0;
      for (const td of document.querySelectorAll("#dataTable tbody tr[data-row-key] td[data-col-index]")) {
        const l = colLeft(td.dataset.colIndex);
        if (l != null && Math.abs(td.getBoundingClientRect().left - l) > 1.5) misaligned++;
      }
      return { holes, jumps, misaligned };
    };
    const sweep = async () => {
      const bad = [];
      const max = () => tableWrapEl.scrollHeight - tableWrapEl.clientHeight;
      for (let k = 0; k <= 40; k++) {
        tableWrapEl.scrollTop = Math.round(max() * (k / 40));
        await fr();
        const c = check();
        if (c.holes || c.jumps || c.misaligned) bad.push({ f: k / 40, ...c });
      }
      return bad;
    };
    const heights = () => new Set([...document.querySelectorAll("#dataTable tbody tr[data-row-key]")].slice(0, 40).map((tr) => Math.round(tr.getBoundingClientRect().height))).size;
    const out = {};
    wrapCellsEl.checked = true;
    wrapCellsEl.dispatchEvent(new Event("change", { bubbles: true }));
    await fr(); await fr();
    out.virtWrapped = swbVirt.isActive();
    out.wrapClass = tableEl.classList.contains("wrap-cells");
    out.distinct = heights();
    out.badWrapped = await sweep();
    tableWrapEl.dispatchEvent(new WheelEvent("wheel", { deltaY: 400 }));
    tableWrapEl.scrollTop = tableWrapEl.scrollHeight;
    await new Promise((res) => setTimeout(res, 900)); await fr();
    const rows = [...document.querySelectorAll("#dataTable tbody tr[data-row-key]")];
    const last = rows[rows.length - 1];
    out.last = last && +last.dataset.rowIndex;
    out.expected = currentDisplayModel.rows[currentDisplayModel.rows.length - 1].rowIndex0;
    out.lastAtBottom = !!last && Math.abs(last.getBoundingClientRect().bottom - tableWrapEl.getBoundingClientRect().bottom) < 4;
    wrapCellsEl.checked = false;
    wrapCellsEl.dispatchEvent(new Event("change", { bubbles: true }));
    await fr(); await fr();
    tableWrapEl.scrollTop = 0; await fr();
    out.virtPlain = swbVirt.isActive();
    out.distinctPlain = heights();
    out.badPlain = await sweep();
    return out;
  });
  ok("1. zawijanie: nowy silnik dalej działa", r.virtWrapped && r.wrapClass, r);
  ok("1. zawijanie: różne wysokości wierszy", r.distinct > 2, r);
  ok("2. zawijanie: bez dziur, przeskoków i rozjechanych kolumn", r.badWrapped.length === 0, r.badWrapped.slice(0, 4));
  ok("3. zawijanie: dół osiągalny", r.last === r.expected && r.lastAtBottom, r);
  ok("4. po wyłączeniu: silnik działa, jedna wysokość, bez dziur", r.virtPlain && r.distinctPlain <= 2 && r.badPlain.length === 0, { virt: r.virtPlain, distinct: r.distinctPlain, bad: r.badPlain.slice(0, 4) });
  ok("bez błędów strony", errors.length === 0, errors);
  await browser.close();
  if (failures.length) {
    console.log(`❌ virt-wrap (${process.env.ENGINE || "chromium"}): ${failures.length} błędów`);
    failures.forEach((f) => console.log("  ✗", f));
    process.exit(1);
  }
  console.log(`✅ virt-wrap (${process.env.ENGINE || "chromium"}): wszystko OK`);
}

run().catch((e) => { console.error(e); process.exit(1); });
