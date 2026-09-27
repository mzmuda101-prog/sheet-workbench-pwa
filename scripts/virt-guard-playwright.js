// virt-guard-playwright.js — STRAŻNIK przed wirtualizacją wierszy (etap F0).
//
// Wirtualizacja = w DOM tylko okno widocznych wierszy + zapas, reszta to wypełniacze.
// Zanim silnik powstanie, ten test opisuje ZACHOWANIE, które użytkownik widzi dziś
// i które nie ma prawa się zmienić. Asercje są celowo „od strony użytkownika”
// (co widać w oknie tabeli, gdzie jest fokus, co trafia do schowka), a nie od strony
// DOM (ile jest <tr>), żeby ten sam test przechodził na obecnym renderze i na
// wirtualnym. Scenariusze pochodzą z przeglądu kodu: miejsca, które dziś szukają
// wiersza w DOM (findCellElement, syncRangeHighlightInDom, focusSection, edytor,
// renderTable z clearMissing), plus to, co zgłaszał Mateusz (szybkie machnięcie).
//
//   1. przewinięcie na sam dół pokazuje ostatni wiersz (i nic po nim),
//   2. brak dziur: w każdym punkcie przewijania (także przy „machnięciu” — wiele
//      skoków scrollTop klatka po klatce) okno tabeli jest pełne kolejnych wierszy,
//   3. strzałka w dół ×40: fokus idzie wiersz po wierszu i zostaje widoczny,
//   4. edycja daleko w arkuszu: Enter zapisuje, edytor znika, fokus schodzi niżej,
//   5. fokus przeżywa przewinięcie daleko + przebudowę tabeli (np. zoom),
//   6. zakres przez wiele ekranów: kopiowanie ma wszystkie wiersze, podświetlenie
//      jest na widocznych komórkach zakresu także po przewinięciu w jego środek,
//   7. skok z panelu analiz (focusSection „scroll-row”) do dalekiego wiersza,
//   8. zmienne wysokości (zawijanie + ręczne wysokości): dół osiągalny, bez dziur,
//   9. blokada 1. kolumny: na dole zamrożona komórka ostatniego wiersza jest przy lewej,
//  10. dokładnie jedna komórka siatki z tabindex=0 (roving) po przewinięciach.
//
// ROWS (domyślnie 600) = limit wierszy ustawiany ręcznie, żeby było co przewijać.
// ENGINE=webkit — te same asercje w silniku Safari. VIRT_QUERY — dopisek do adresu
// (np. "&virt=1"), żeby w kolejnych etapach puścić ten sam test po nowym silniku.
//
// Uruchom z serwerem na APP_URL (domyślnie http://127.0.0.1:4175/).

const playwright = require("playwright");

const APP_URL = process.env.APP_URL || "http://127.0.0.1:4175/";
const ENGINE = process.env.ENGINE || "chromium";
const ROWS = parseInt(process.env.ROWS || "600", 10);
const SAMPLE = parseInt(process.env.SAMPLE || "3000", 10);
const VIRT_QUERY = process.env.VIRT_QUERY || "";
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

// Narzędzia wstrzykiwane do strony (jako window.__vg).
function installHelpers() {
  const frames = (n = 2) => new Promise((res) => {
    const step = () => (n-- <= 0 ? res() : requestAnimationFrame(step));
    requestAnimationFrame(step);
  });
  const limit = () => Math.max(1, parseInt(maxRowsEl.value || "200", 10));
  const reachable = () => Math.min(currentDisplayModel.rows.length, limit());
  const keyAt = (i) => getRowSelectionKey(currentDisplayModel.rows[i]);
  const indexOfKey = (() => {
    let cacheModel = null; let map = null;
    return (key) => {
      if (cacheModel !== currentDisplayModel) {
        cacheModel = currentDisplayModel;
        map = new Map(currentDisplayModel.rows.map((r, i) => [getRowSelectionKey(r), i]));
      }
      return map.has(key) ? map.get(key) : -1;
    };
  })();
  // Pas, w którym widać wiersze danych: od dołu przyklejonego nagłówka do dołu okna.
  const band = () => {
    const wrap = tableWrapEl.getBoundingClientRect();
    const head = theadEl.getBoundingClientRect();
    const top = Math.max(wrap.top, Math.min(wrap.bottom, head.bottom));
    return { top, bottom: wrap.bottom - Math.max(0, wrap.height - tableWrapEl.clientHeight), left: wrap.left, right: wrap.left + tableWrapEl.clientWidth };
  };
  // Co widzi użytkownik w pionie: próbkujemy elementFromPoint co ~12 px w środku
  // pierwszej kolumny danych. Każdy punkt musi trafić w komórkę wiersza danych,
  // a indeksy wierszy muszą iść kolejno (bez luk i powtórek w złej kolejności).
  const sampleColumn = () => {
    const b = band();
    const firstData = theadEl.querySelector(".header-row th:nth-child(2)");
    const fr = firstData ? firstData.getBoundingClientRect() : null;
    let x = fr ? Math.round(fr.left + Math.min(fr.width / 2, 30)) : Math.round(b.left + 80);
    x = Math.max(b.left + 2, Math.min(b.right - 2, x));
    const hits = [];
    const holes = [];
    for (let y = Math.ceil(b.top + 3); y < b.bottom - 3; y += 12) {
      const el = document.elementFromPoint(x, y);
      const tr = el && el.closest ? el.closest("#dataTable tbody tr[data-row-key]") : null;
      if (!tr) { holes.push({ y: Math.round(y - b.top), what: el ? `${el.tagName}.${el.className}` : "null" }); continue; }
      hits.push(indexOfKey(tr.dataset.rowKey));
    }
    const seq = hits.filter((v, i) => i === 0 || v !== hits[i - 1]);
    let gap = null;
    for (let i = 1; i < seq.length; i++) {
      if (seq[i] !== seq[i - 1] + 1) { gap = `${seq[i - 1]}→${seq[i]}`; break; }
    }
    return { holes, gap, first: seq[0], last: seq[seq.length - 1], count: seq.length };
  };
  const isCellVisible = (td) => {
    if (!td) return false;
    const b = band();
    const r = td.getBoundingClientRect();
    const midY = (r.top + r.bottom) / 2;
    return midY > b.top && midY < b.bottom && r.right > b.left && r.left < b.right;
  };
  window.__vg = { frames, limit, reachable, keyAt, indexOfKey, band, sampleColumn, isCellVisible };
}

async function openApp(browser) {
  const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1100, height: 800 } });
  await context.addInitScript((rows) => {
    localStorage.setItem("introPlayed", "true");
    localStorage.setItem("excel-workbench-max-rows", String(rows)); // ręczny limit wygrywa z automatem
  }, ROWS);
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push("pageerror: " + e.message));
  page.on("console", (m) => { if (m.type() === "error") errors.push("console: " + m.text()); });
  await page.goto(`${APP_URL}?sample=${SAMPLE}${VIRT_QUERY}`, { waitUntil: "load" });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  await page.evaluate(() => document.getElementById("loadSampleBtn")?.click());
  await page.waitForFunction(() => document.querySelectorAll("#dataTable tbody tr[data-row-key]").length > 0, null, { timeout: 60000 });
  await page.evaluate(() => { if (typeof setSidebarOpen === "function") setSidebarOpen(false); if (typeof syncSidebarHandle === "function") syncSidebarHandle(); });
  await sleep(500);
  await page.evaluate(installHelpers);
  // Zwinięty nagłówek (hero) na czas testu — jego animacja przy scrollu zmienia wysokość
  // okna tabeli i zafałszowałaby pomiary (patrz notatki przy freeze-pane).
  await page.evaluate(() => { try { if (typeof heroUserCollapsed !== "undefined") heroUserCollapsed = true; } catch (_) {} });
  return { context, page, errors };
}

async function run() {
  const browser = await playwright[ENGINE].launch({ headless: true });
  const results = [];
  const check = (name, ok, detail = "") => results.push({ name, ok: !!ok, detail });

  const { context, page, errors } = await openApp(browser);
  const info = await page.evaluate(() => ({ reachable: __vg.reachable(), total: currentDisplayModel.rows.length, limit: __vg.limit() }));
  check("przygotowanie: jest co przewijać", info.reachable >= Math.min(ROWS, 300), JSON.stringify(info));
  const last = info.reachable - 1;

  // 1. dół
  const bottom = await page.evaluate(async () => {
    tableWrapEl.scrollTop = tableWrapEl.scrollHeight;
    await __vg.frames(3);
    return __vg.sampleColumn();
  });
  check("1. na dole widać ostatni wiersz", bottom.last === last, JSON.stringify(bottom));
  check("1. na dole bez dziur", !bottom.holes.length && !bottom.gap, JSON.stringify(bottom));

  // 2. brak dziur — skoki po całej wysokości + „machnięcie” (skok co klatkę)
  const noHoles = await page.evaluate(async () => {
    const bad = [];
    const max = tableWrapEl.scrollHeight - tableWrapEl.clientHeight;
    for (const f of [0, 0.13, 0.27, 0.41, 0.58, 0.73, 0.88, 1, 0.5, 0.05]) {
      tableWrapEl.scrollTop = Math.round(max * f);
      await __vg.frames(2);
      const s = __vg.sampleColumn();
      if (s.holes.length || s.gap || !s.count) bad.push({ f, ...s, holes: s.holes.slice(0, 3) });
    }
    // machnięcie: 40 klatek po ~1,5 ekranu w dół, sprawdzenie PO każdej klatce
    tableWrapEl.scrollTop = 0;
    await __vg.frames(2);
    const stepPx = Math.round(tableWrapEl.clientHeight * 1.5);
    for (let i = 0; i < 40 && tableWrapEl.scrollTop < max; i++) {
      tableWrapEl.scrollTop += stepPx;
      await __vg.frames(1);
      const s = __vg.sampleColumn();
      if (s.holes.length || s.gap || !s.count) { bad.push({ fling: i, ...s, holes: s.holes.slice(0, 3) }); break; }
    }
    return bad;
  });
  check("2. brak dziur przy skokach i machnięciu", noHoles.length === 0, JSON.stringify(noHoles).slice(0, 600));

  // 3. strzałka w dół ×40 z pierwszej komórki
  await page.evaluate(async () => { tableWrapEl.scrollTop = 0; await __vg.frames(2); });
  await page.click("#dataTable tbody tr[data-row-key] td[data-col-index='1']");
  await sleep(150);
  for (let i = 0; i < 40; i++) await page.keyboard.press("ArrowDown");
  await sleep(250);
  const arrow = await page.evaluate(() => {
    const idx = focusedCellState ? __vg.indexOfKey(focusedCellState.rowKey) : -1;
    const td = findCellElement(focusedCellState);
    return { idx, col: focusedCellState && focusedCellState.colIndex0, visible: __vg.isCellVisible(td || findFocusedRowElement()), active: document.activeElement === td };
  });
  check("3. ↓×40: fokus w wierszu 40", arrow.idx === 40, JSON.stringify(arrow));
  check("3. ↓×40: fokus widoczny i w DOM", arrow.visible && arrow.active, JSON.stringify(arrow));

  // 4. edycja daleko w arkuszu
  const editIdx = Math.max(0, last - 8);
  await page.evaluate(async (i) => {
    setFocusedCell(__vg.keyAt(i), 1, { scroll: true });
    await __vg.frames(3);
  }, editIdx);
  const editTarget = await page.evaluate((i) => {
    const td = findCellElement({ rowKey: __vg.keyAt(i), colIndex0: 1 });
    return td ? { ok: __vg.isCellVisible(td) } : { ok: false };
  }, editIdx);
  check("4. edycja: docelowa komórka przewinięta do widoku", editTarget.ok, JSON.stringify(editTarget));
  const selector = await page.evaluate((i) => `#dataTable tbody tr[data-row-key="${CSS.escape(__vg.keyAt(i))}"] td[data-col-index="1"]`, editIdx);
  // Porażka ma być ZGŁOSZONA (✗), a nie wywalić cały test — stąd krótkie limity i catch.
  const opened = await page.dblclick(selector, { timeout: 3000 })
    .then(() => page.waitForSelector("#dataTable input.cell-editor", { timeout: 3000 }))
    .then(() => true, () => false);
  if (opened) {
    await page.fill("#dataTable input.cell-editor", "VG-EDIT");
    await page.keyboard.press("Enter");
  }
  check("4. edycja: edytor otwiera się na dalekiej komórce", opened);
  await sleep(300);
  const edit = await page.evaluate((i) => {
    const row = currentDisplayModel.rows[i];
    const td = findCellElement({ rowKey: getRowSelectionKey(row), colIndex0: 1 });
    const focusIdx = focusedCellState ? __vg.indexOfKey(focusedCellState.rowKey) : -1;
    return {
      model: String(row.values[1]),
      dom: td ? td.textContent : null,
      editorOpen: !!document.querySelector("#dataTable input.cell-editor"),
      focusIdx,
      focusVisible: __vg.isCellVisible(findCellElement(focusedCellState)),
    };
  }, editIdx);
  check("4. edycja: wartość w danych i w komórce", edit.model === "VG-EDIT" && edit.dom === "VG-EDIT", JSON.stringify(edit));
  check("4. edycja: Enter zamyka edytor i schodzi niżej (widocznie)", !edit.editorOpen && edit.focusIdx === editIdx + 1 && edit.focusVisible, JSON.stringify(edit));

  // 5. fokus przeżywa przewinięcie + przebudowę
  const survive = await page.evaluate(async () => {
    tableWrapEl.scrollTop = 0;
    await __vg.frames(2);
    setFocusedCell(__vg.keyAt(3), 2, { scroll: false });
    await __vg.frames(1);
    tableWrapEl.scrollTop = tableWrapEl.scrollHeight;
    await __vg.frames(2);
    renderActiveTable(); // np. zmiana zoomu / opcji wyświetlania
    await __vg.frames(2);
    const mid = focusedCellState ? __vg.indexOfKey(focusedCellState.rowKey) : -1;
    tableWrapEl.scrollTop = 0;
    await __vg.frames(3);
    const tr = findFocusedRowElement();
    return { stateAfterRender: mid, domRowFocused: !!(tr && tr.classList.contains("row-focused")), idx: tr ? __vg.indexOfKey(tr.dataset.rowKey) : -1 };
  });
  check("5. fokus przeżywa przewinięcie i przebudowę", survive.stateAfterRender === 3 && survive.domRowFocused && survive.idx === 3, JSON.stringify(survive));

  // 6. zakres przez wiele ekranów
  const range = await page.evaluate(async (lastIdx) => {
    tableWrapEl.scrollTop = 0;
    await __vg.frames(2);
    const endIdx = Math.max(10, lastIdx - 2);
    setFocusedCell(__vg.keyAt(2), 1, { scroll: false });
    setSelectionKind("cell", { repaint: false });
    setSelectedCell(__vg.keyAt(endIdx), 2, { scroll: false });
    syncRangeHighlightInDom();
    const rect = getSelectionRectangle();
    const tsv = rect ? serializeRangeToTsv(rect) : "";
    const lines = tsv.replace(/\n$/, "").split("\n").length;
    // przewiń w środek zakresu i sprawdź, że widoczne komórki kolumn 1–2 są podświetlone
    const midKey = __vg.keyAt(Math.round((2 + endIdx) / 2));
    const midTr0 = document.querySelector(`#dataTable tbody tr[data-row-key="${CSS.escape(midKey)}"]`);
    if (midTr0) midTr0.scrollIntoView({ block: "center" });
    else tableWrapEl.scrollTop = Math.round((tableWrapEl.scrollHeight - tableWrapEl.clientHeight) / 2);
    await __vg.frames(3);
    const visibleRows = [...document.querySelectorAll("#dataTable tbody tr[data-row-key]")].filter((tr) => {
      const i = __vg.indexOfKey(tr.dataset.rowKey);
      return i > 2 && i < endIdx && __vg.isCellVisible(tr.querySelector("td[data-col-index='1']"));
    });
    const unlit = visibleRows.filter((tr) => !tr.querySelector("td[data-col-index='1']").classList.contains("cell-in-range")
      || !tr.querySelector("td[data-col-index='2']").classList.contains("cell-in-range"));
    const out = { expected: endIdx - 2 + 1, lines, visibleChecked: visibleRows.length, unlit: unlit.length };
    // sprzątanie
    setSelectedCell(null, NaN); setFocusedCell(null, NaN); setSelectionKind("row", { repaint: true });
    return out;
  }, last);
  check("6. zakres: kopiowanie ma wszystkie wiersze", range.lines === range.expected, JSON.stringify(range));
  check("6. zakres: widoczne komórki w środku są podświetlone", range.visibleChecked > 5 && range.unlit === 0, JSON.stringify(range));

  // 7. skok z panelu analiz do dalekiego wiersza
  const jumpIdx = Math.max(0, last - 20);
  await page.evaluate(async (i) => {
    tableWrapEl.scrollTop = 0;
    await __vg.frames(2);
    focusSection({ action: "scroll-row", rowIndex0: currentDisplayModel.rows[i].rowIndex0 });
  }, jumpIdx);
  // płynne przewinięcie o kilkanaście tysięcy px trwa ~1,5 s — czekamy, aż stanie
  await page.evaluate(async () => {
    let prev = -1;
    for (let k = 0; k < 40; k++) {
      await new Promise((r) => setTimeout(r, 150));
      if (tableWrapEl.scrollTop === prev) break;
      prev = tableWrapEl.scrollTop;
    }
  });
  const jump = await page.evaluate((i) => {
    const tr = document.querySelector(`#dataTable tbody tr[data-row-index="${currentDisplayModel.rows[i].rowIndex0}"]`);
    const td = tr && tr.querySelector("td[data-col-index='0']");
    const r = td ? td.getBoundingClientRect() : null;
    return { found: !!tr, visible: __vg.isCellVisible(td), top: r && Math.round(r.top), band: __vg.band(), scrollTop: tableWrapEl.scrollTop, docScroll: document.scrollingElement.scrollTop };
  }, jumpIdx);
  check("7. skok z analiz: wiersz widoczny", jump.found && jump.visible, JSON.stringify(jump));

  // 8. zmienne wysokości: zawijanie tekstu + ręcznie wyższe co 7. wiersze (próbka ma
  //    krótkie teksty, więc samo zawijanie nie dałoby różnych wysokości)
  const wrap = await page.evaluate(async () => {
    if (!wrapCellsEl) return { skipped: true };
    const tall = currentDisplayModel.rows.filter((_, i) => i % 7 === 3).map((r) => r.rowIndex0);
    tall.forEach((ri) => { manualRowHeights[ri] = 64; });
    wrapCellsEl.checked = true;
    wrapCellsEl.dispatchEvent(new Event("change", { bubbles: true }));
    renderActiveTable();
    await __vg.frames(3);
    const heights = new Set([...document.querySelectorAll("#dataTable tbody tr[data-row-key]")].slice(0, 80).map((tr) => Math.round(tr.getBoundingClientRect().height)));
    tableWrapEl.scrollTop = tableWrapEl.scrollHeight;
    await __vg.frames(3);
    const s = __vg.sampleColumn();
    const mid = [];
    const max = tableWrapEl.scrollHeight - tableWrapEl.clientHeight;
    for (const f of [0.2, 0.6, 0.9]) {
      tableWrapEl.scrollTop = Math.round(max * f);
      await __vg.frames(2);
      const m = __vg.sampleColumn();
      if (m.holes.length || m.gap) mid.push({ f, ...m });
    }
    wrapCellsEl.checked = false;
    tall.forEach((ri) => { delete manualRowHeights[ri]; });
    wrapCellsEl.dispatchEvent(new Event("change", { bubbles: true }));
    renderActiveTable();
    await __vg.frames(2);
    return { distinctHeights: heights.size, bottomLast: s.last, bottomHoles: s.holes.length, gap: s.gap, mid };
  });
  check("8. zmienne wysokości: dół osiągalny, bez dziur", wrap.skipped || (wrap.distinctHeights > 1 && wrap.bottomLast === last && !wrap.bottomHoles && !wrap.gap && !wrap.mid.length), JSON.stringify(wrap));

  // 9. blokada 1. kolumny
  const freeze = await page.evaluate(async () => {
    if (!freezeFirstColEl) return { skipped: true };
    freezeFirstColEl.checked = true;
    freezeFirstColEl.dispatchEvent(new Event("change", { bubbles: true }));
    await __vg.frames(2);
    tableWrapEl.scrollLeft = 400;
    tableWrapEl.scrollTop = tableWrapEl.scrollHeight;
    await new Promise((r) => setTimeout(r, 400)); // szyba (freeze-pane) wraca do spoczynku po 180 ms
    await __vg.frames(2);
    const lastKey = __vg.keyAt(__vg.reachable() - 1);
    const td = document.querySelector(`#dataTable tbody tr[data-row-key="${CSS.escape(lastKey)}"] td[data-col-index="0"]`);
    const b = __vg.band();
    const r = td ? td.getBoundingClientRect() : null;
    const out = { found: !!td, leftOffset: r ? Math.round(r.left - b.left) : null, visible: __vg.isCellVisible(td) };
    tableWrapEl.scrollLeft = 0;
    freezeFirstColEl.checked = false;
    freezeFirstColEl.dispatchEvent(new Event("change", { bubbles: true }));
    await __vg.frames(2);
    return out;
  });
  check("9. blokada kolumny: zamrożona komórka ostatniego wiersza przy lewej", freeze.skipped || (freeze.found && freeze.visible && freeze.leftOffset < 120), JSON.stringify(freeze));

  // 10. roving tabindex
  const roving = await page.evaluate(async () => {
    tableWrapEl.scrollTop = 0;
    await __vg.frames(2);
    setFocusedCell(__vg.keyAt(5), 1, { scroll: false });
    tableWrapEl.scrollTop = Math.round(tableWrapEl.scrollHeight / 2);
    await __vg.frames(2);
    renderActiveTable();
    await __vg.frames(2);
    tableWrapEl.scrollTop = 0;
    await __vg.frames(2);
    return document.querySelectorAll("#dataTable tbody td[tabindex='0']").length;
  });
  check("10. dokładnie jedna komórka z tabindex=0", roving === 1, String(roving));

  check("brak błędów strony/konsoli", errors.length === 0, errors.join(" | "));
  await context.close();
  await browser.close();

  const failed = results.filter((r) => !r.ok);
  results.forEach((r) => console.log(`${r.ok ? "✓" : "✗"} ${r.name}${r.ok ? "" : "  — " + r.detail}`));
  if (failed.length) throw new Error(`${failed.length} z ${results.length} asercji nie przeszło (${ENGINE}${VIRT_QUERY ? " " + VIRT_QUERY : ""})`);
  console.log(`\n✅ virt-guard-playwright OK — ${results.length} asercji (${ENGINE}, ${info.reachable} wierszy${VIRT_QUERY ? ", " + VIRT_QUERY : ""})`);
}

run().catch((error) => {
  console.error(error);
  process.exit(1);
});
