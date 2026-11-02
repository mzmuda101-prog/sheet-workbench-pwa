// virt-merges-playwright.js — nowy silnik tabeli (przewijanie wirtualne) na arkuszu ze scaleniami.
//
// 2026-11-02, krok 1 planu „nowy silnik domyślnie”: dotąd scalenia oddawały arkusz staremu silnikowi.
// Teraz okno wierszy nigdy nie przecina scalenia (rozszerza się do całego), więc:
//   1. silnik działa na arkuszu ze scaleniami (400 wierszy, fixtures/virt-merges.xlsx),
//   2. w każdym miejscu przewijania: bez dziur, kolumny równe z nagłówkiem (żadna komórka nie „ucieka”),
//   3. każdy widoczny przykryty wiersz ma w DOM kotwicę swojego scalenia (rowspan sięga do niego),
//   4. dół osiągalny (ostatni wiersz), skok w środek scalenia i z powrotem.
// Uruchom z serwerem na APP_URL (domyślnie http://127.0.0.1:4175/), np. `npm run serve:test`.
// Fixture: .venv/bin/python scripts/fixtures/make-virt-merges.py

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
  await page.waitForFunction(() => currentSheetName === "Grafik" && baseRows.length > 300, null, { timeout: 30000 });
  await page.waitForTimeout(1200);

  const failures = [];
  const ok = (name, cond, got) => { if (!cond) failures.push(`${name} — dostałem: ${JSON.stringify(got)}`); };

  const r = await page.evaluate(async () => {
    const fr = () => new Promise((res) => requestAnimationFrame(() => requestAnimationFrame(res)));
    const settle = async () => { await fr(); await new Promise((res) => setTimeout(res, 400)); await fr(); };
    const heads = [...document.querySelectorAll("thead tr.header-row th[data-col-index]")];
    const colLeft = (i) => { const th = heads.find((t) => t.dataset.colIndex === String(i)); return th ? th.getBoundingClientRect().left : null; };
    const merges = (currentMerges || []).filter((m) => m.e.r > m.s.r);
    const check = () => {
      const wb = tableWrapEl.getBoundingClientRect();
      const head = document.querySelector("thead tr.header-row th").getBoundingClientRect();
      let holes = 0;
      for (let y = Math.max(wb.top, head.bottom) + 4; y < wb.bottom - 4; y += 10) {
        const el = document.elementFromPoint(wb.left + 60, y);
        if (!el || !el.closest("#dataTable tbody tr[data-row-key]")) holes++;
      }
      // kolumny równe z nagłówkiem: każda narysowana komórka stoi pod swoim nagłówkiem
      let misaligned = 0;
      for (const td of document.querySelectorAll("#dataTable tbody tr[data-row-key] td[data-col-index]")) {
        const l = colLeft(td.dataset.colIndex);
        if (l != null && Math.abs(td.getBoundingClientRect().left - l) > 1.5) misaligned++;
      }
      // przykryte wiersze widoczne na ekranie mają kotwicę scalenia w DOM
      let orphan = 0;
      const shown = new Set([...document.querySelectorAll("#dataTable tbody tr[data-row-key]")].map((tr) => +tr.dataset.rowIndex));
      for (const m of merges) {
        let covered = false;
        for (let rr = m.s.r + 1; rr <= m.e.r; rr++) if (shown.has(rr)) covered = true;
        if (!covered) continue;
        const tr = document.querySelector(`#dataTable tbody tr[data-row-index="${m.s.r}"]`);
        const td = tr && tr.querySelector(`td[data-col-index="${m.s.c - currentStartCol}"]`);
        if (!td || td.rowSpan < 2) orphan++;
      }
      return { holes, misaligned, orphan };
    };
    const out = { virt: swbVirt.isActive(), merges: merges.length, points: [] };
    const max = () => tableWrapEl.scrollHeight - tableWrapEl.clientHeight;
    // gęsto (40 pozycji), żeby brzeg okna trafił też w środek pionowych scaleń — przy kilku
    // punktach test przechodził nawet z wyłączonym rozszerzaniem okna (kontrola negatywna)
    for (let k = 0; k <= 40; k++) {
      const f = k / 40;
      tableWrapEl.scrollTop = Math.round(max() * f);
      await fr();
      const c = check();
      if (c.holes || c.misaligned || c.orphan) out.points.push({ f, ...c });
    }
    out.points.push({ f: "wszystkie", holes: 0, misaligned: 0, orphan: 0 });
    // skok dokładnie w środek wysokiego scalenia (kotwica powyżej górnej krawędzi okna)
    const tall = merges.find((m) => m.e.r - m.s.r >= 2 && m.s.r > 150);
    if (tall) {
      const p = currentDisplayModel.rows.findIndex((row) => row.rowIndex0 === tall.s.r + 1);
      tableWrapEl.scrollTop = 0; await settle();
      if (typeof swbVirt.ensureRowRendered === "function") swbVirt.ensureRowRendered(p);
      const tr = document.querySelector(`#dataTable tbody tr[data-row-index="${tall.s.r + 1}"]`);
      if (tr) tr.scrollIntoView({ block: "start" });
      await settle();
      out.jump = check();
    }
    tableWrapEl.dispatchEvent(new WheelEvent("wheel", { deltaY: 400 }));
    tableWrapEl.scrollTop = tableWrapEl.scrollHeight;
    await new Promise((res) => setTimeout(res, 900)); await fr();
    const rows = [...document.querySelectorAll("#dataTable tbody tr[data-row-key]")];
    out.last = rows.length && +rows[rows.length - 1].dataset.rowIndex;
    out.expected = currentDisplayModel.rows[currentDisplayModel.rows.length - 1].rowIndex0;
    out.bottom = check();
    return out;
  });
  ok("1. nowy silnik działa mimo scaleń", r.virt && r.merges > 30, r);
  r.points.forEach((p) => {
    ok(`2. f=${p.f}: bez dziur`, p.holes === 0, p);
    ok(`2. f=${p.f}: kolumny równe z nagłówkiem`, p.misaligned === 0, p);
    ok(`3. f=${p.f}: każde widoczne scalenie ma kotwicę`, p.orphan === 0, p);
  });
  ok("4. skok w środek scalenia: kotwica i kolumny na miejscu", r.jump && !r.jump.orphan && !r.jump.misaligned && !r.jump.holes, r.jump);
  ok("4. dół osiągalny", r.last === r.expected && r.bottom.orphan === 0 && r.bottom.misaligned === 0, r);
  ok("bez błędów strony", errors.length === 0, errors);
  await browser.close();
  if (failures.length) {
    console.log(`❌ virt-merges (${process.env.ENGINE || "chromium"}): ${failures.length} błędów`);
    failures.forEach((f) => console.log("  ✗", f));
    process.exit(1);
  }
  console.log(`✅ virt-merges (${process.env.ENGINE || "chromium"}): wszystko OK`);
}

run().catch((e) => { console.error(e); process.exit(1); });
