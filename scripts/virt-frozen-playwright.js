// virt-frozen-playwright.js — zablokowane kolumny w nowym silniku tabeli (przewijanie wirtualne).
//
// 2026-11-02, krok 3 planu „nowy silnik domyślnie”. W nowym silniku zablokowane kolumny to zwykłe
// przyklejanie (sticky) bez „szyby” (freeze-pane.js): pomiar telefon/CPU×6 — w DOM ~40–70 wierszy, więc
// blokada kosztuje ~10–20% w pionie, a w bok jest taniej niż stary silnik z szybą (149 vs 271 ms).
// Tu pilnujemy POPRAWNOŚCI:
//   1. silnik działa, blokada 2 kolumn włączona,
//   2. po przewinięciu w bok i w dół (wiersze dorysowane w trakcie) zablokowane komórki każdego
//      widocznego wiersza stoją przy lewej krawędzi, jedna za drugą, bez szczeliny i bez nachodzenia,
//   3. pierwsza niezablokowana kolumna jest przewinięta (wychodzi spod blokady, nie stoi w miejscu),
//   4. to samo po skoku na sam dół.
// ENGINE=webkit TOUCH=1 — Safari z dotykiem (telefon).

const { chromium, webkit } = require("playwright");

const APP_URL = process.env.APP_URL || "http://127.0.0.1:4175/";
const ENGINE = process.env.ENGINE === "webkit" ? webkit : chromium;
const TOUCH = process.env.TOUCH === "1";

async function run() {
  const browser = await ENGINE.launch({ headless: true });
  const ctx = await browser.newContext({
    serviceWorkers: "block",
    viewport: TOUCH ? { width: 390, height: 844 } : { width: 1100, height: 800 },
    hasTouch: TOUCH,
    isMobile: TOUCH && process.env.ENGINE !== "webkit",
  });
  await ctx.addInitScript(() => localStorage.setItem("introPlayed", "true"));
  const page = await ctx.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  await page.goto(APP_URL + "?sample=3000&virt=1", { waitUntil: "load" });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  await page.evaluate(() => document.getElementById("loadSampleBtn")?.click());
  await page.waitForFunction(() => document.querySelectorAll("#dataTable tbody tr[data-row-key]").length > 0, null, { timeout: 60000 });
  await page.evaluate(() => { try { heroUserCollapsed = true; } catch (_) {} setSheetFrozenCols(2); });
  await page.waitForTimeout(800);

  const failures = [];
  const ok = (name, cond, got) => { if (!cond) failures.push(`${name} — dostałem: ${JSON.stringify(got)}`); };

  const r = await page.evaluate(async () => {
    const fr = () => new Promise((res) => requestAnimationFrame(() => requestAnimationFrame(res)));
    const settle = async () => { await fr(); await new Promise((res) => setTimeout(res, 300)); await fr(); };
    const check = () => {
      const wb = tableWrapEl.getBoundingClientRect();
      const bad = [];
      for (const tr of document.querySelectorAll("#dataTable tbody tr[data-row-key]")) {
        const rr = tr.getBoundingClientRect();
        if (rr.bottom < wb.top + 40 || rr.top > wb.bottom - 4) continue; // tylko widoczne
        const head = tr.querySelector("td.row-head");
        const c0 = tr.querySelector('td[data-col-index="0"]');
        const c1 = tr.querySelector('td[data-col-index="1"]');
        const c2 = tr.querySelector('td[data-col-index="2"]');
        if (!head || !c0 || !c1 || !c2) continue;
        const h = head.getBoundingClientRect(), a = c0.getBoundingClientRect(), b = c1.getBoundingClientRect(), c = c2.getBoundingClientRect();
        const atEdge = Math.abs(h.left - wb.left) < 2;
        const chain = Math.abs(a.left - h.right) < 1.5 && Math.abs(b.left - a.right) < 1.5;
        const scrolled = c.left < b.right - 5; // kolumna 3 przewinięta pod blokadę
        if (!atEdge || !chain || !scrolled) bad.push({ row: tr.dataset.rowIndex, atEdge, chain, scrolled });
      }
      return bad;
    };
    const out = { virt: swbVirt.isActive(), frozen: getFrozenColCount(), points: [] };
    tableWrapEl.scrollLeft = 400; await settle();
    for (const f of [0.1, 0.33, 0.5, 0.77]) {
      tableWrapEl.scrollTop = Math.round((tableWrapEl.scrollHeight - tableWrapEl.clientHeight) * f);
      await settle();
      const bad = check();
      out.points.push({ f, bad: bad.length, sample: bad.slice(0, 2) });
    }
    tableWrapEl.dispatchEvent(new WheelEvent("wheel", { deltaY: 400 }));
    tableWrapEl.scrollTop = tableWrapEl.scrollHeight;
    await new Promise((res) => setTimeout(res, 900)); await fr();
    const bad = check();
    out.points.push({ f: "dół", bad: bad.length, sample: bad.slice(0, 2) });
    out.scrollLeft = tableWrapEl.scrollLeft;
    return out;
  });
  ok("1. nowy silnik + blokada 2 kolumn", r.virt && r.frozen === 2, r);
  ok("1. przewinięte w bok", r.scrollLeft > 100, r);
  r.points.forEach((p) => ok(`2–4. ${p.f}: zablokowane kolumny na miejscu, reszta przewinięta`, p.bad === 0, p));
  ok("bez błędów strony", errors.length === 0, errors);
  await browser.close();
  const tag = `${process.env.ENGINE || "chromium"}${TOUCH ? " TOUCH" : ""}`;
  if (failures.length) {
    console.log(`❌ virt-frozen (${tag}): ${failures.length} błędów`);
    failures.forEach((f) => console.log("  ✗", f));
    process.exit(1);
  }
  console.log(`✅ virt-frozen (${tag}): wszystko OK`);
}

run().catch((e) => { console.error(e); process.exit(1); });
