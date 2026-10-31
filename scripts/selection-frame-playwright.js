// selection-frame-playwright.js — STRAŻNIK WYGLĄDU ZAZNACZENIA „JAK W EXCELU" (2026-10-27).
//   node scripts/selection-frame-playwright.js        ENGINE=webkit node scripts/selection-frame-playwright.js
//
// Dawniej zaznaczona komórka była zaokrąglonym „kafelkiem" (radius 6 px, ramka 1,7 px przy
// 62% krycia + poświata), a uchwyt wypełniania siedział W komórce — td ma overflow: hidden,
// więc przycinał go do wnętrza ramki. Co musi być prawdą teraz:
//   1. pojedyncza komórka (tryb komórkowy): ostra ramka 2 px w akcencie, BEZ zaokrąglenia
//      i bez poświaty (jeden cień), kryjące tło;
//   2. podświetlony numer wiersza (td.row-head) i litera kolumny (th.guide-cell) — dokładnie
//      jeden każdy; w trybie wiersza tylko numer wiersza;
//   3. uchwyt leży w #tableWrap (nie w td), środkiem na narożniku ramki, WYSTAJE poza komórkę
//      i nic go nie zasłania (elementFromPoint w jego środku = uchwyt);
//   4. zakres 3×3: obwódka tylko na brzegach (12 krawędzi brzegowych, środek bez krawędzi),
//      kotwica .range-anchor niezabarwiona, 3 numery wierszy + 3 litery kolumn;
//   5. zamrożona 1. kolumna: róg w niej → uchwyt zostaje w komórce (is-inset);
//   6. ciemny motyw: ramka w kolorze --accent motywu.
// Uruchom z serwerem na APP_URL (domyślnie http://127.0.0.1:4175/).

const playwright = require("playwright");
const path = require("path");

const APP_URL = process.env.APP_URL || "http://127.0.0.1:4175/";
const ENGINE = process.env.ENGINE || "chromium";
const FILE = path.join(__dirname, "stress-test-workbench.xlsx");
const SLEEP_SCALE = Math.min(1, Math.max(0.1, Number(process.env.SLEEP_SCALE || 1)));
const sleep = (ms) => new Promise((r) => setTimeout(r, Math.round(ms * SLEEP_SCALE)));

async function run() {
  const browser = await playwright[ENGINE].launch({ headless: true });
  const fails = [];
  const check = (name, ok, detail) => {
    console.log(`${ok ? "✅" : "❌"} ${name}${ok ? "" : ` — ${JSON.stringify(detail)}`}`);
    if (!ok) fails.push(name);
  };

  const ctx = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1280, height: 800 } });
  await ctx.addInitScript(() => localStorage.setItem("introPlayed", "true"));
  const page = await ctx.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  await page.evaluate(() => { try { ensureXlsxLibs && ensureXlsxLibs(false); } catch {} });
  await page.setInputFiles("#fileInput", FILE);
  await page.waitForFunction(() => document.getElementById("sheetSelect")?.options?.length > 0, null, { timeout: 20000 });
  await sleep(300);
  await page.evaluate(() => loadBtn.click());
  await page.waitForFunction(() => !!document.querySelector('#dataTable tbody tr[data-row-key="wide:5"] td[data-col-index="4"]'), null, { timeout: 20000 });
  await sleep(800);

  const td = (r, c) => `#dataTable tbody tr[data-row-key="wide:${r}"] td[data-col-index="${c}"]`;
  const settle = () => page.evaluate(() => new Promise((r) => requestAnimationFrame(() => requestAnimationFrame(r))));

  // Stan zaznaczenia odczytany w JEDNYM przejściu po stronie.
  const probe = (sel) => page.evaluate((sel) => {
    const cell = document.querySelector(sel);
    const st = cell ? getComputedStyle(cell) : null;
    const rootSt = getComputedStyle(document.documentElement);
    const handle = document.querySelector("#tableWrap .fill-handle");
    const hostTd = document.querySelector("#dataTable td.has-fill-handle");
    let h = null;
    if (handle && hostTd) {
      const hb = handle.getBoundingClientRect();
      const tb = hostTd.getBoundingClientRect();
      const cx = hb.left + hb.width / 2;
      const cy = hb.top + hb.height / 2;
      const after = getComputedStyle(handle, "::after");
      const hit = document.elementFromPoint(cx, cy);
      h = {
        inWrap: handle.parentElement === document.getElementById("tableWrap"),
        inset: handle.classList.contains("is-inset"),
        inTd: !!handle.closest("td"),
        host: hostTd.dataset.colIndex + "@" + hostTd.parentElement.dataset.rowKey,
        dx: Math.round(cx - tb.right), dy: Math.round(cy - tb.bottom),
        sticksOut: hb.right > tb.right + 1 && hb.bottom > tb.bottom + 1,
        boxW: parseFloat(after.width) || 0,
        hitSelf: hit === handle,
      };
    }
    // kolor akcentu motywu jako rgb(...) — przez próbnik, bo zmienna to hex
    const probeEl = document.createElement("i");
    probeEl.style.color = rootSt.getPropertyValue("--accent").trim();
    document.body.appendChild(probeEl);
    const accentRgb = getComputedStyle(probeEl).color;
    probeEl.remove();
    return {
      classes: cell ? cell.className : "",
      radius: st ? st.borderTopLeftRadius : "",
      shadow: st ? st.boxShadow : "",
      bg: st ? st.backgroundColor : "",
      accentRgb,
      rowHeads: [...document.querySelectorAll("#dataTable tbody td.row-head.hdr-sel")].map((x) => x.parentElement.dataset.rowKey),
      guides: [...document.querySelectorAll("#dataTable thead tr.guide-row th.guide-cell.hdr-sel")].map((x) => x.dataset.fullText),
      inRange: document.querySelectorAll("#dataTable tbody td.cell-in-range").length,
      handle: h,
    };
  }, sel);
  // poziom komórki: w trybie Edycja daje go już sam klik (2026-10-31); Shift+Spacja tylko,
  // gdyby klik zostawił poziom wiersza (Podgląd) — inaczej przełączyłaby z powrotem na wiersz
  const toCell = async () => { if (await page.evaluate(() => selectionKind) !== "cell") await page.keyboard.press("Shift+Space"); };
  const shadowCount = (s) => (s && s !== "none" ? s.split(/,(?![^(]*\))/).length : 0);

  // 1–3. pojedyncza komórka w trybie komórkowym (klik + Shift+Spacja)
  await page.click(td(3, 2));
  await toCell();
  await settle();
  let p = await probe(td(3, 2));
  check("1. aktywna komórka ma klasę cell-active", /\bcell-active\b/.test(p.classes), p.classes);
  check("1. ramka bez zaokrąglenia (radius 0)", p.radius === "0px", p.radius);
  check("1. jedna ostra ramka inset 2 px (bez poświaty)", shadowCount(p.shadow) === 1 && /inset/.test(p.shadow) && /0px 0px 0px 2px/.test(p.shadow), p.shadow);
  check("1. ramka w kolorze akcentu", p.shadow.includes(p.accentRgb), { shadow: p.shadow, accent: p.accentRgb });
  check("1. tło kryjące (nie przezroczyste)", !/rgba\(.*,\s*0\)|transparent/.test(p.bg) && !/rgba\([^)]*,\s*0?\.\d+\)/.test(p.bg), p.bg);
  check("2. podświetlony dokładnie jeden numer wiersza (wide:3)", JSON.stringify(p.rowHeads) === '["wide:3"]', p.rowHeads);
  check("2. podświetlona dokładnie jedna litera kolumny (C)", JSON.stringify(p.guides) === '["C"]', p.guides);
  check("3. uchwyt w #tableWrap, nie w komórce", !!p.handle && p.handle.inWrap && !p.handle.inTd && !p.handle.inset, p.handle);
  check("3. uchwyt na narożniku komórki C4", !!p.handle && p.handle.host === "2@wide:3" && Math.abs(p.handle.dx) <= 3 && Math.abs(p.handle.dy) <= 3, p.handle);
  check("3. uchwyt wystaje poza komórkę (nie przycięty)", !!p.handle && p.handle.sticksOut && p.handle.boxW >= 6, p.handle);
  check("3. nic nie zasłania uchwytu (elementFromPoint)", !!p.handle && p.handle.hitSelf, p.handle);

  // strzałka w trybie komórki: nagłówki idą za komórką, stare gasną
  await page.keyboard.press("ArrowRight");
  await settle();
  p = await probe(td(3, 3));
  check("2. po strzałce: litera D, wiersz bez zmian, poprzednia litera zgaszona", JSON.stringify(p.guides) === '["D"]' && JSON.stringify(p.rowHeads) === '["wide:3"]', p);

  // 4. zakres 3×3 (Shift+klik)
  await page.click(td(3, 2));
  await page.click(td(5, 4), { modifiers: ["Shift"] });
  await settle();
  const range = await page.evaluate(() => {
    const out = { cells: [], anchor: null };
    document.querySelectorAll("#dataTable tbody td.cell-in-range").forEach((c) => {
      const edges = ["top", "bottom", "left", "right"].filter((e) => c.classList.contains(`range-edge-${e}`));
      out.cells.push({ k: c.parentElement.dataset.rowKey + "/" + c.dataset.colIndex, edges, r: getComputedStyle(c).borderTopLeftRadius });
      if (c.classList.contains("range-anchor")) out.anchor = c.parentElement.dataset.rowKey + "/" + c.dataset.colIndex;
    });
    const a = document.querySelector("#dataTable td.range-anchor");
    const mid = document.querySelector('#dataTable tbody tr[data-row-key="wide:4"] td[data-col-index="3"]');
    out.anchorBg = a ? getComputedStyle(a).backgroundColor : "";
    out.midBg = mid ? getComputedStyle(mid).backgroundColor : "";
    out.midShadow = mid ? getComputedStyle(mid).boxShadow : "";
    return out;
  });
  const edgeCount = range.cells.reduce((n, c) => n + c.edges.length, 0);
  const mid = range.cells.find((c) => c.k === "wide:4/3");
  check("4. zakres = 9 komórek", range.cells.length === 9, range.cells.length);
  check("4. obwódka tylko na brzegach (12 krawędzi, środek bez krawędzi)", edgeCount === 12 && mid && mid.edges.length === 0, range.cells);
  check("4. środek zakresu bez widocznego cienia-ramki", !range.midShadow || range.midShadow === "none" || !/\b2px\b/.test(range.midShadow), range.midShadow);
  check("4. bez zaokrąglonych narożników", range.cells.every((c) => c.r === "0px"), range.cells.map((c) => c.r));
  check("4. kotwica (C4) oznaczona i niezabarwiona jak reszta", range.anchor === "wide:3/2" && range.anchorBg !== range.midBg, range);
  p = await probe(td(5, 4));
  check("4. nagłówki zakresu: 3 numery wierszy, 3 litery kolumn", p.rowHeads.length === 3 && JSON.stringify(p.guides) === '["C","D","E"]', { rows: p.rowHeads, guides: p.guides });
  check("4. uchwyt na narożniku zakresu (E6), wystaje", !!p.handle && p.handle.host === "4@wide:5" && Math.abs(p.handle.dx) <= 3 && Math.abs(p.handle.dy) <= 3 && p.handle.sticksOut && p.handle.hitSelf, p.handle);

  // zwykły klik w trybie Edycja = komórka (jak w Excelu, 2026-10-31): zakres znika,
  // świecą litera B i numer wiersza wide:7, komórka ma ramkę
  await page.click(td(7, 1));
  await settle();
  p = await probe(td(7, 1));
  check("5. po zwykłym kliku: brak zakresu, litera B, numer wiersza wide:7, ramka komórki", p.inRange === 0 && JSON.stringify(p.guides) === '["B"]' && JSON.stringify(p.rowHeads) === '["wide:7"]' && /\bcell-active\b/.test(p.classes), p);
  // …a w Podglądzie klik dalej zaznacza wiersz: bez litery kolumny
  await page.evaluate(() => sheetEditUi.setMode("view"));
  await page.click(td(8, 1));
  await settle();
  p = await probe(td(8, 1));
  check("5. Podgląd: klik = wiersz (brak litery), numer wiersza wide:8", p.inRange === 0 && p.guides.length === 0 && JSON.stringify(p.rowHeads) === '["wide:8"]', p);
  await page.evaluate(() => sheetEditUi.setMode("edit"));

  // 6. zamrożona 1. kolumna — róg w niej: uchwyt zostaje w komórce
  await page.evaluate(() => {
    const el = document.getElementById("freezeFirstCol");
    el.checked = true;
    el.dispatchEvent(new Event("change", { bubbles: true }));
  });
  await sleep(300);
  await page.click(td(4, 0));
  await toCell();
  await settle();
  p = await probe(td(4, 0));
  check("6. zamrożona kolumna: uchwyt w komórce (is-inset), widoczny", !!p.handle && p.handle.inset && p.handle.inTd && p.handle.host === "0@wide:4" && p.handle.hitSelf, p.handle);
  check("6. zamrożona kolumna: ramka komórki bez zaokrąglenia, tło kryjące", p.radius === "0px" && !/rgba\([^)]*,\s*0?\.\d+\)/.test(p.bg), { r: p.radius, bg: p.bg });
  await page.click(td(4, 2));
  await toCell();
  await settle();
  p = await probe(td(4, 2));
  check("6. kolumna niezamrożona: uchwyt wraca do #tableWrap", !!p.handle && p.handle.inWrap && !p.handle.inset, p.handle);
  await page.evaluate(() => {
    const el = document.getElementById("freezeFirstCol");
    el.checked = false;
    el.dispatchEvent(new Event("change", { bubbles: true }));
  });

  // 7. ciemny motyw — ramka w akcencie ciemnego motywu
  await page.evaluate(() => document.documentElement.setAttribute("data-theme", "dark"));
  await settle();
  p = await probe(td(4, 2));
  check("7. ciemny motyw: ramka w kolorze akcentu, radius 0", p.shadow.includes(p.accentRgb) && p.radius === "0px", { shadow: p.shadow, accent: p.accentRgb });

  check("brak błędów strony", errors.length === 0, errors);
  await browser.close();
  if (fails.length) { console.error(`❌ selection-frame (${ENGINE}): ${fails.length} błąd(y)`); process.exit(1); }
  console.log(`✅ selection-frame (${ENGINE}) OK`);
}

run().catch((e) => { console.error(e); process.exit(1); });
