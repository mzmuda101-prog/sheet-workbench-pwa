// virt-excel-dims-playwright.js — nowy silnik tabeli (przewijanie wirtualne) z wymiarami z Excela.
//
// Zgłoszenie 2026-11-02: w nowym silniku nie było „prawdziwych wymiarów” (skala tekstu Excela,
// domyślna wysokość wierszy z pliku, „Zawijaj tekst” z pliku). Teraz są — a wysokości dalszych
// wierszy to dla silnika tylko szacunek, więc pierwsze przewinięcie na sam dół lądowało ~60 wierszy
// przed końcem (RODO, arkusz 563 w.: dół → 504 zamiast 564). Po zatrzymaniu silnik dociąga do końca.
//   1. silnik działa, tekst w skali Excela (11 pt = 11 px przy 100%),
//   2. przewinięcie kółkiem na sam dół → po zatrzymaniu widać ostatni wiersz, bez dziur.
// Uruchom z serwerem na APP_URL (domyślnie http://127.0.0.1:4175/), np. `npm run serve:test`.

const { chromium, webkit } = require("playwright");
const path = require("path");

const APP_URL = process.env.APP_URL || "http://127.0.0.1:4175/";
const ENGINE = process.env.ENGINE === "webkit" ? webkit : chromium;
const FILE = path.join(__dirname, "RODO Obieg terenów 2026 V5.xlsx");

async function run() {
  const browser = await ENGINE.launch({ headless: true });
  const ctx = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1300, height: 850 } });
  await ctx.addInitScript(() => localStorage.setItem("introPlayed", "true"));
  const page = await ctx.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  await page.goto(APP_URL + "?virt=1", { waitUntil: "load" });
  await page.evaluate(() => { document.getElementById("heroSplash")?.remove(); try { ensureXlsxLibs(false); } catch {} });
  await page.setInputFiles("#fileInput", FILE);
  await page.waitForFunction(() => typeof baseRows !== "undefined" && baseRows.length > 0, null, { timeout: 30000 });
  // największy arkusz (563 wiersze — silnik przejmuje od 200)
  const big = await page.evaluate(() => workbook.SheetNames.map((n) => [n, XLSX.utils.decode_range(workbook.Sheets[n]["!ref"] || "A1").e.r]).sort((a, b) => b[1] - a[1])[0][0]);
  await page.evaluate((s) => { sheetSelect.value = s; loadBtn.click(); }, big);
  await page.waitForFunction((s) => currentSheetName === s && baseRows.length > 0, big);
  await page.evaluate(() => { if (zoomLevelEl.value !== "1") { zoomLevelEl.value = "1"; zoomLevelEl.dispatchEvent(new Event("change")); } });
  await page.waitForTimeout(1200);

  const failures = [];
  const ok = (name, cond, got) => { if (!cond) failures.push(`${name} — dostałem: ${JSON.stringify(got)}`); };

  const r = await page.evaluate(async () => {
    const fr = () => new Promise((res) => requestAnimationFrame(() => requestAnimationFrame(res)));
    const td = document.querySelector("#dataTable tbody tr[data-row-key] td:nth-child(3)");
    // komórki mają własne rozmiary (10 pt → 10 px) — sprawdzamy bazę skali na tabeli
    const out = { virt: swbVirt.isActive(), fs: tableEl.style.getPropertyValue("--cell-font-base"), base: getWorkbookDefaultFontSize(workbook), cell: td && getComputedStyle(td).fontSize };
    tableWrapEl.dispatchEvent(new WheelEvent("wheel", { deltaY: 400 }));
    tableWrapEl.scrollTop = tableWrapEl.scrollHeight;
    await new Promise((res) => setTimeout(res, 1000)); await fr();
    // „Duchy” (row-grow.js): blade puste wiersze ZA danymi w Edycji — to nie dziura; dół danych = górna
    // krawędź duchów (pod nią, do dołu okna, mają stać same duchy)
    const dataBottom = (bottom) => { const g = document.querySelector("#dataTable tbody.ghost-body"); return g && g.offsetParent ? Math.min(bottom, g.getBoundingClientRect().top) : bottom; };
    const wb = tableWrapEl.getBoundingClientRect();
    const rows = [...document.querySelectorAll("#dataTable tbody tr[data-row-key]")];
    const last = rows[rows.length - 1];
    let holes = 0;
    for (let y = dataBottom(wb.bottom) - 200; y < dataBottom(wb.bottom) - 4; y += 10) {
      const el = document.elementFromPoint(wb.left + 120, y);
      if (!el || !el.closest("#dataTable tbody tr[data-row-key]")) holes++;
    }
    out.lastIdx = last && +last.dataset.rowIndex;
    out.expected = currentDisplayModel.rows[currentDisplayModel.rows.length - 1].rowIndex0;
    out.lastVisible = !!last && Math.abs(last.getBoundingClientRect().bottom - dataBottom(wb.bottom)) < 4;
    // duchy (jeśli są) wypełniają resztę okna do samego dołu — bez pustego pasa pod nimi
    const g = document.querySelector("#dataTable tbody.ghost-body");
    out.ghostsToBottom = !(g && g.offsetParent) || Math.abs(g.getBoundingClientRect().bottom - wb.bottom) < 4;
    out.holes = holes;
    return out;
  });
  ok("1. nowy silnik działa", r.virt, r);
  ok("1. tekst w skali Excela (rozmiar domyślnej czcionki pliku w px)", parseFloat(r.fs) === r.base, r);
  ok("2. dół osiągalny za pierwszym razem (ostatni wiersz przy dolnej krawędzi)", r.lastIdx === r.expected && r.lastVisible, r);
  ok("2. bez dziur przy dole", r.holes === 0 && r.ghostsToBottom, r);
  ok("bez błędów strony", errors.length === 0, errors);
  await browser.close();
  if (failures.length) {
    console.log(`❌ virt-excel-dims (${process.env.ENGINE || "chromium"}): ${failures.length} błędów`);
    failures.forEach((f) => console.log("  ✗", f));
    process.exit(1);
  }
  console.log(`✅ virt-excel-dims (${process.env.ENGINE || "chromium"}): wszystko OK`);
}

run().catch((e) => { console.error(e); process.exit(1); });
