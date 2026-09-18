// keyboard-nav-playwright.js — obsługa BEZ MYSZY (tablet + klawiatura).
//
// Pilnuje, że aplikacji da się używać z samej klawiatury:
//  (1) roving tabindex — do siatki wchodzi się Tabem, dokładnie jednym skokiem,
//  (2) strzałki ruszają kursor komórki bez wcześniejszego kliknięcia myszą,
//  (3) zwinięty sidebar znika z kolejności Tab (inert),
//  (4) Ctrl+Alt+1/2/3 i F6 skaczą między panelem / paskiem / tabelą,
//  (5) globalny handler nie zjada strzałek <select>om trybu/akcji szukania,
// plus regresje na tym, co działało wcześniej: wyjście Tabem z siatki, edytor
// komórki (Enter/Esc), utrzymanie fokusu w polu szukania przy przebudowie tabeli.
//
// Uruchom z serwerem na APP_URL (domyślnie http://127.0.0.1:4175/).

// ENGINE=webkit pozwala przepuscic te same asercje przez silnik Safari/iPada —
// tam :focus-visible i obsluga fokusu zachowuja sie inaczej niz w Chromium.
const playwright = require("playwright");
const ENGINE = process.env.ENGINE || "chromium";
// TOUCH=1 emuluje iPada (pointer: coarse) — tam <select> ustepuje miejsca segmentom.
const TOUCH = process.env.TOUCH === "1";
const chromium = playwright[ENGINE];
const path = require("path");

const APP_URL = process.env.APP_URL || "http://127.0.0.1:4175/";
const FILE = path.join(__dirname, "stress-test-workbench.xlsx");
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

const results = [];
function check(name, pass, detail) {
  results.push({ name, pass: !!pass, detail });
}

async function run() {
  const browser = await chromium.launch({ headless: true });
  const context = await browser.newContext(TOUCH
    ? { serviceWorkers: "block", ...playwright.devices["iPad Pro 11"] }
    : { serviceWorkers: "block", viewport: { width: 1280, height: 900 } });
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
  await sleep(1200);

  // ── 1. Roving tabindex: DOKŁADNIE jedna komórka jest punktem wejścia ──────────
  const roving = await page.evaluate(() => document.querySelectorAll('#dataTable tbody td[tabindex="0"]').length);
  check("roving tabindex = dokładnie 1 komórka wejściowa", roving === 1, `znaleziono ${roving}`);

  // ── 2. Skok Ctrl+Alt+3 do tabeli + ruch strzałkami BEZ kliknięcia myszą ───────
  await page.evaluate(() => document.body.focus());
  await page.keyboard.press("Control+Alt+Digit3");
  await sleep(200);
  const afterJump = await page.evaluate(() => ({
    tag: document.activeElement?.tagName,
    col: document.activeElement?.dataset?.colIndex,
    state: focusedCellState ? { ...focusedCellState } : null,
  }));
  check("Ctrl+Alt+3 ustawia fokus na <td>", afterJump.tag === "TD", JSON.stringify(afterJump));
  check("…i ustawia model kursora komórki", !!afterJump.state, JSON.stringify(afterJump.state));

  // Model gestów: domyślnie zaznaczony jest WIERSZ, a wtedy → przewija w bok.
  // Ten test sprawdza ruch kursora, więc najpierw schodzimy na poziom komórki.
  await page.evaluate(() => setSelectionKind("cell"));
  await page.keyboard.press("ArrowRight");
  await page.keyboard.press("ArrowDown");
  await sleep(200);
  const afterArrows = await page.evaluate(() => ({
    tag: document.activeElement?.tagName,
    col: document.activeElement?.dataset?.colIndex,
    state: focusedCellState ? { ...focusedCellState } : null,
    rovingCount: document.querySelectorAll('#dataTable tbody td[tabindex="0"]').length,
  }));
  check("strzałki ruszają kursor bez klikania (kolumna 0→1)",
    afterArrows.col === "1" && afterArrows.state?.colIndex0 === 1, JSON.stringify(afterArrows));
  check("fokus DOM jedzie za kursorem", afterArrows.tag === "TD", afterArrows.tag);
  check("roving nadal = 1 po ruchu", afterArrows.rovingCount === 1, `${afterArrows.rovingCount}`);
  check("wiersz się zmienił po ArrowDown",
    afterArrows.state?.rowKey !== afterJump.state?.rowKey, `${afterJump.state?.rowKey} → ${afterArrows.state?.rowKey}`);

  // ── 3. REGRESJA: Tab z siatki wychodzi jednym skokiem (nie przez tysiące komórek)
  await page.keyboard.press("Tab");
  await sleep(150);
  const afterTabOut = await page.evaluate(() => ({
    tag: document.activeElement?.tagName,
    inGrid: document.getElementById("dataTable")?.contains(document.activeElement),
  }));
  check("Tab wychodzi z siatki (nie skacze po komórkach)", !afterTabOut.inGrid, JSON.stringify(afterTabOut));

  // ── 4. inert: zwinięty sidebar znika z kolejności Tab ─────────────────────────
  const inertState = await page.evaluate(() => {
    const sb = document.querySelector(".sidebar");
    setSidebarOpen(false);
    const closed = sb.hasAttribute("inert");
    setSidebarOpen(true);
    const opened = sb.hasAttribute("inert");
    return { closed, opened };
  });
  check("zwinięty sidebar dostaje inert", inertState.closed === true, JSON.stringify(inertState));
  check("rozwinięty sidebar traci inert", inertState.opened === false, JSON.stringify(inertState));

  // Realny test: przy zwiniętym panelu Tab NIE wchodzi w jego kontrolki.
  const tabWalk = await page.evaluate(async () => {
    setSidebarOpen(false);
    document.body.focus();
    return null;
  });
  let hitsInSidebar = 0;
  for (let i = 0; i < 12; i++) {
    await page.keyboard.press("Tab");
    const inSb = await page.evaluate(() => document.querySelector(".sidebar")?.contains(document.activeElement));
    if (inSb) hitsInSidebar++;
  }
  check("12× Tab przy zwiniętym panelu nie trafia w jego kontrolki", hitsInSidebar === 0, `trafień: ${hitsInSidebar}`);
  await page.evaluate(() => setSidebarOpen(true));

  // ── 5. Strzałki nie kradną <select> trybu w oknie szukania ────────────────────
  // Asercja idzie na defaultPrevented, NIE na zmianę wartości: headless Chromium
  // nie przestawia zamkniętego <select> strzałką, więc wartość nic by nie mówiła.
  // Istotne jest, że nasz globalny handler przestał zjadać to zdarzenie.
  // Termin musi mieć >= 2 znaki (próg live-podglądu) i realnie występować w danych —
  // szukamy pierwszej dostatecznie długiej wartości tekstowej w kilku pierwszych wierszach.
  const term = await page.evaluate(() => {
    for (const row of baseRows.slice(0, 40)) {
      for (let c = 0; c < currentHeaders.length; c++) {
        const v = String(getDisplayValue(row, c) ?? "").trim();
        if (v.length >= 4 && /[a-zA-Z\u00c0-\u017f]{4}/.test(v)) return v.slice(0, 4);
      }
    }
    return null;
  });
  if (!term) throw new Error("nie znalazłem tekstowej wartości do wyszukania w pliku testowym");
  await page.evaluate(() => {
    document.getElementById("quickSearchPopup").classList.remove("hidden");
    const input = document.getElementById("quickSearchPopupInput");
    input.value = "";
    input.focus();
    window.__prevented = null;
    document.addEventListener("keydown", (e) => { window.__prevented = e.defaultPrevented; }, false);
  });
  await page.keyboard.type(term);
  await sleep(700);
  const liveInfo = await page.evaluate(() => ({
    visible: !document.getElementById("qsLiveResultsPopup")?.classList.contains("hidden"),
    hits: document.querySelectorAll("#qsLiveResultsPopup .qs-live-item").length,
  }));
  check("live-lista ma realne trafienia (warunek testu)",
    liveInfo.visible && liveInfo.hits > 0, JSON.stringify(liveInfo) + ` term="${term}"`);

  // Na dotyku <select> jest ukryty (zastąpiony segmentami), więc ta asercja
  // dotyczy wyłącznie profilu desktopowego.
  const selectWidoczny = await page.evaluate(() =>
    getComputedStyle(document.getElementById("quickSearchPopupMode")).display !== "none");
  if (selectWidoczny) {
    await page.evaluate(() => {
      window.__prevented = null;
      document.getElementById("quickSearchPopupMode").focus();
    });
    await page.keyboard.press("ArrowDown");
    await sleep(200);
    const selectEvt = await page.evaluate(() => ({
      prevented: window.__prevented,
      stillFocused: document.activeElement?.id === "quickSearchPopupMode",
    }));
    check("↓ na <select> Tryb NIE jest przechwytywane (select dostaje swoje zdarzenie)",
      selectEvt.prevented === false && selectEvt.stillFocused, JSON.stringify(selectEvt));
  }

  // ── 6. REGRESJA: ↓ z POLA szukania nadal wchodzi na listę wyników ─────────────
  await page.evaluate(() => document.getElementById("quickSearchPopupInput").focus());
  await page.keyboard.press("ArrowDown");
  await sleep(200);
  const onLive = await page.evaluate(() =>
    !!document.activeElement?.classList?.contains("qs-live-item"));
  check("↓ z pola szukania nadal wchodzi na listę trafień", onLive, `${onLive}`);

  await page.keyboard.press("Escape");
  await sleep(200);

  // ── 7. REGRESJA: render nie kradnie fokusu z pola szukania ───────────────────
  // Inline'owy #quickSearch żyje tylko w trybie czytania (display:none inaczej),
  // więc testujemy na polu okna szukania — tam realnie się pisze przy szukaniu.
  await page.evaluate(() => {
    document.getElementById("quickSearchPopup").classList.remove("hidden");
    document.getElementById("quickSearchPopupInput").focus();
  });
  const focusBeforeRender = await page.evaluate(() => document.activeElement?.id);
  await page.evaluate(() => { renderActiveTable(); });
  await sleep(300);
  const focusKept = await page.evaluate(() => document.activeElement?.id);
  check("render tabeli nie wyrywa fokusu z pola szukania",
    focusBeforeRender === "quickSearchPopupInput" && focusKept === "quickSearchPopupInput",
    `${focusBeforeRender} → ${focusKept}`);
  await page.evaluate(() => document.getElementById("quickSearchPopup").classList.add("hidden"));

  // ── 8. REGRESJA: sortowanie zachowuje fokus w siatce ──────────────────────────
  await page.evaluate(() => {
    const cell = document.querySelector('#dataTable tbody td[data-col-index="0"]');
    syncGridRovingTabindex(cell);
    cell.focus();
  });
  await sleep(150);
  const beforeSort = await page.evaluate(() => document.activeElement?.tagName);
  await page.evaluate(() => { renderActiveTable(); });
  await sleep(300);
  const afterSort = await page.evaluate(() => ({
    tag: document.activeElement?.tagName,
    inGrid: document.getElementById("dataTable")?.contains(document.activeElement),
  }));
  check("po przebudowie tabeli fokus wraca do siatki",
    beforeSort === "TD" && afterSort.inGrid, `${beforeSort} → ${JSON.stringify(afterSort)}`);

  // ── 9. REGRESJA: edytor komórki (Enter otwiera, Esc anuluje) ──────────────────
  await page.evaluate(() => {
    const cell = document.querySelector('#dataTable tbody tr[data-row-key] td[data-col-index="0"]');
    syncGridRovingTabindex(cell);
    cell.focus();
  });
  await sleep(150);
  await page.keyboard.press("Enter");
  await sleep(300);
  const editorOpen = await page.evaluate(() => ({
    isInput: document.activeElement?.classList?.contains("cell-editor"),
  }));
  check("Enter na komórce nadal otwiera edytor", editorOpen.isInput, JSON.stringify(editorOpen));
  await page.keyboard.press("Escape");
  await sleep(250);
  const editorClosed = await page.evaluate(() => !document.querySelector(".cell-editor"));
  check("Esc zamyka edytor", editorClosed, `${editorClosed}`);

  // ── 10. F6 jako alias cyklu regionów ─────────────────────────────────────────
  await page.evaluate(() => { setSidebarOpen(true); document.body.focus(); });
  await page.keyboard.press("F6");
  await sleep(200);
  const f6 = await page.evaluate(() => ({
    inSidebar: document.querySelector(".sidebar")?.contains(document.activeElement),
    tag: document.activeElement?.tagName,
  }));
  check("F6 wchodzi do panelu", f6.inSidebar, JSON.stringify(f6));

  await page.keyboard.press("F6");
  await sleep(200);
  const f6b = await page.evaluate(() => ({
    inSidebar: document.querySelector(".sidebar")?.contains(document.activeElement),
    region: currentAppRegion(),
    tag: document.activeElement?.tagName,
  }));
  check("kolejne F6 opuszcza panel (idzie dalej w cyklu)", !f6b.inSidebar, JSON.stringify(f6b));
  check("cykl F6 nie grzęźnie na pustym regionie", f6b.region !== null && f6b.region !== "panel", JSON.stringify(f6b));

  // ── 11. Skip-link: pierwszy przystanek Tab prowadzi wprost na komórkę ────────
  const skip = await page.evaluate(() => {
    const all = [...document.querySelectorAll('a[href],button,input,select,textarea,[tabindex]:not([tabindex^="-"])')]
      .filter((el) => !el.disabled && !el.closest(".hidden")
        && (el.offsetParent !== null || getComputedStyle(el).position === "fixed"));
    return { first: all[0]?.id, opacity: getComputedStyle(document.getElementById("skipToTable")).opacity };
  });
  check("skip-link jest pierwszym przystankiem Tab", skip.first === "skipToTable", JSON.stringify(skip));
  check("skip-link niewidoczny dopóki nie dostanie fokusu", skip.opacity === "0", `opacity: ${skip.opacity}`);

  await page.evaluate(() => document.getElementById("skipToTable").focus());
  await sleep(600); // pozwól dobiec animacji zjazdu (transition na transform+opacity)
  const skipVisible = await page.evaluate(() =>
    getComputedStyle(document.getElementById("skipToTable")).opacity);
  await page.keyboard.press("Enter");
  await sleep(300);
  const afterSkip = await page.evaluate(() => ({
    tag: document.activeElement?.tagName, state: focusedCellState,
  }));
  check("skip-link po ofokusowaniu staje się widoczny", parseFloat(skipVisible) > 0.9, `opacity: ${skipVisible}`);
  check("Enter na skip-linku ląduje na aktywnej komórce",
    afterSkip.tag === "TD" && !!afterSkip.state, JSON.stringify(afterSkip));

  // ── 12. Model gestów: tryb wynika z tego, JAK zaznaczyłeś ───────────────────
  const gest = async (fn) => page.evaluate(fn);
  const state = () => page.evaluate(() => ({
    kind: selectionKind,
    rows: document.querySelectorAll("tbody tr.row-focused").length,
    cells: document.querySelectorAll("tbody td.cell-active").length,
    range: document.querySelectorAll("tbody td.cell-in-range").length,
  }));

  // (a) zwykły klik → cały wiersz
  // Sidebar musi być zwinięty — wcześniejsze testy F6 go otworzyły, a rozwinięty
  // przykrywa tabelę i przechwytuje kliknięcia.
  await page.evaluate(() => { setSidebarOpen(false); setSelectionKind("row"); });
  await sleep(400);
  await page.click('#dataTable tbody tr[data-row-key] td[data-col-index="1"]');
  await sleep(200);
  let st = await state();
  check("zwykły klik zaznacza cały wiersz", st.kind === "row" && st.rows === 1 && st.cells === 0, JSON.stringify(st));

  // (b) strzałka w dół rusza WIERSZEM, nie komórką
  const rowBefore = await page.evaluate(() => focusedCellState.rowKey);
  await page.keyboard.press("ArrowDown");
  await sleep(200);
  st = await state();
  const rowAfter = await page.evaluate(() => focusedCellState.rowKey);
  check("↓ przy zaznaczonym wierszu przesuwa wiersz",
    st.kind === "row" && st.rows === 1 && rowAfter !== rowBefore, JSON.stringify(st));

  // (c) ←/→ przewijają w bok, zaznaczenie wiersza NIETKNIĘTE
  const scrollBefore = await page.evaluate(() => {
    const w = document.getElementById("tableWrap");
    w.scrollLeft = 0;
    return { left: w.scrollLeft, scrollable: w.scrollWidth - w.clientWidth };
  });
  await page.keyboard.press("ArrowRight");
  await sleep(250);
  const afterScroll = await page.evaluate(() => ({
    left: document.getElementById("tableWrap").scrollLeft,
    kind: selectionKind,
    rows: document.querySelectorAll("tbody tr.row-focused").length,
    col: focusedCellState?.colIndex0,
  }));
  check("→ przy zaznaczonym wierszu przewija tabelę w bok",
    scrollBefore.scrollable > 0 && afterScroll.left > scrollBefore.left,
    JSON.stringify({ scrollBefore, afterScroll }));
  check("…i NIE zmienia zaznaczenia wiersza ani kolumny",
    afterScroll.kind === "row" && afterScroll.rows === 1, JSON.stringify(afterScroll));

  // (d) Shift+Spacja → schodzi do pojedynczej komórki
  await page.evaluate(() => {
    const cell = document.querySelector('#dataTable tbody td[tabindex="0"]');
    if (cell) cell.focus();
  });
  await page.keyboard.press("Shift+Space");
  await sleep(250);
  st = await state();
  check("Shift+Spacja schodzi z wiersza do pojedynczej komórki",
    st.kind === "cell" && st.cells === 1 && st.rows === 0, JSON.stringify(st));

  // (e) teraz strzałki ruszają KOMÓRKĄ, także w bok
  const colBefore = await page.evaluate(() => focusedCellState.colIndex0);
  await page.keyboard.press("ArrowRight");
  await sleep(200);
  st = await state();
  const colAfter = await page.evaluate(() => focusedCellState.colIndex0);
  check("→ w trybie komórki przesuwa komórkę (a nie przewija)",
    colAfter === colBefore + 1 && st.cells === 1 && st.rows === 0,
    JSON.stringify({ colBefore, colAfter, ...st }));

  // (f) Shift+Spacja z powrotem → cały wiersz, zakres zwinięty
  await page.keyboard.press("Shift+Space");
  await sleep(250);
  st = await state();
  check("Shift+Spacja wraca do całego wiersza",
    st.kind === "row" && st.rows === 1 && st.cells === 0 && st.range === 0, JSON.stringify(st));

  // (g) Shift+strzałka zawsze schodzi na poziom komórki i buduje zakres
  await page.keyboard.press("Shift+ArrowDown");
  await sleep(250);
  st = await state();
  check("Shift+strzałka z wiersza schodzi do zakresu komórek",
    st.kind === "cell" && st.range > 0, JSON.stringify(st));

  // (h) Shift+klik ustawia tryb komórki (gest myszy nadal działa)
  await page.evaluate(() => { setSelectionKind("row"); setSelectedCell("", -1); });
  await page.click('#dataTable tbody tr[data-row-key] td[data-col-index="2"]', { modifiers: ["Shift"] });
  await sleep(250);
  st = await state();
  check("Shift+klik ustawia poziom komórki", st.kind === "cell" && st.rows === 0, JSON.stringify(st));

  // (i) zwykły klik wraca do trybu domyślnego
  await page.click('#dataTable tbody tr[data-row-key] td[data-col-index="1"]');
  await sleep(200);
  st = await state();
  check("zwykły klik wraca do trybu domyślnego (wiersz)", st.kind === "row" && st.rows === 1, JSON.stringify(st));

  // (j) klik wraca do wiersza ZAWSZE — także wprost z trybu komórki (nie ma już
  // ustawienia „trybu domyślnego", więc gest jest jedynym źródłem prawdy)
  await page.evaluate(() => setSelectionKind("cell"));
  await page.click('#dataTable tbody tr[data-row-key] td[data-col-index="1"]');
  await sleep(200);
  st = await state();
  check("klik z trybu komórki też wraca do wiersza (bez wyjątków)",
    st.kind === "row" && st.rows === 1 && st.cells === 0, JSON.stringify(st));
  const noSetting = await page.evaluate(() => ({
    field: !!document.getElementById("selectionMode"),
    fn: typeof window.setSelectionMode,
  }));
  check("usunięte ustawienie nie zostawiło po sobie kontrolki ani funkcji",
    noSetting.field === false && noSetting.fn === "undefined", JSON.stringify(noSetting));

  // (k) REGRESJA: zwykła Spacja nadal otwiera edytor (nic nie zabraliśmy)
  await page.evaluate(() => { setSelectionKind("row"); });
  await page.evaluate(() => {
    const cell = document.querySelector('#dataTable tbody td[tabindex="0"]');
    if (cell) cell.focus();
  });
  await page.keyboard.press("Space");
  await sleep(350);
  const spaceEditor = await page.evaluate(() => ({
    open: !!document.querySelector(".cell-editor"),
    value: document.querySelector(".cell-editor")?.value,
  }));
  check("zwykła Spacja nadal otwiera edytor (Shift+Spacja nic nie zabrała)",
    spaceEditor.open === true, JSON.stringify(spaceEditor));
  await page.keyboard.press("Escape");
  await sleep(200);
  await page.evaluate(() => setSelectionKind("row"));

  // ── 13. Okno szybkiego szukania: obsługa z samej klawiatury ─────────────────
  await page.evaluate(() => { closeQuickSearchPopup(); setSidebarOpen(false); });
  await sleep(300);
  await page.evaluate(() => { openQuickSearchPopup(); });
  await sleep(200);
  await page.keyboard.type(term);
  await sleep(700);
  const hits = await page.evaluate(() =>
    document.querySelectorAll("#qsLiveResultsPopup .qs-live-item").length);
  check("okno szukania ma trafienia na liście (warunek testu)", hits > 0, `${hits}`);

  // Lista wyników poza kolejnością Tab — inaczej 8 trafień = 8 przystanków
  // wciśniętych między kontrolki okna.
  const liveTabbable = await page.evaluate(() =>
    [...document.querySelectorAll("#qsLiveResultsPopup .qs-live-item")].filter((el) => el.tabIndex >= 0).length);
  check("pozycje listy wyników są poza kolejnością Tab", liveTabbable === 0, `tabbowalnych: ${liveTabbable}`);

  // Pełny obieg Tab po oknie musi wrócić do pola, nie uciec na <body>
  await page.evaluate(() => document.getElementById("quickSearchPopupInput").focus());
  let escapes = 0;
  const visited = [];
  for (let i = 0; i < 20; i++) {
    await page.keyboard.press("Tab");
    const info = await page.evaluate(() => ({
      out: !document.getElementById("quickSearchPopup").contains(document.activeElement),
      id: document.activeElement?.id || document.activeElement?.className,
    }));
    if (info.out) escapes++;
    // 9 przystanków: tyle ma okno z paskiem chipów „Tryby auto"; bez niego pętla
    // i tak wraca do pola wcześniej (asercja szuka pola w odwiedzonych).
    if (i < 9) visited.push(info.id);
  }
  check("Tab nie ucieka z okna szukania (pułapka fokusu)", escapes === 0, `ucieczek: ${escapes}`);
  check("obieg Tab wraca do pola szukania", visited.includes("quickSearchPopupInput"), visited.join(" → "));

  // „Operatory wyszukiwania": Spacja przełącza I widać, że się na tym stoi
  const ops = await page.evaluate(() => {
    const cb = document.getElementById("quickSearchPopupOperators");
    cb.checked = false;
    cb.focus();
    const pill = cb.closest(".qs-operators-toggle");
    return { przed: cb.checked, ring: getComputedStyle(pill).boxShadow !== "none" };
  });
  await page.keyboard.press("Space");
  await sleep(200);
  const opsAfter = await page.evaluate(() => document.getElementById("quickSearchPopupOperators").checked);
  check("Spacja przełącza „Operatory wyszukiwania”", ops.przed === false && opsAfter === true,
    `${ops.przed} → ${opsAfter}`);
  check("pigułka „Operatory” pokazuje fokus klawiatury", ops.ring === true, `ring: ${ops.ring}`);

  // Esc zamyka okno i oddaje fokus TAM, skąd je otwarto — nawet gdy w tabeli
  // jest zaznaczona komórka (kiedyś Esc odznaczał ją i okno zostawało otwarte)
  await page.evaluate(() => { closeQuickSearchPopup(); });
  await sleep(200);
  await page.evaluate(() => {
    const cell = document.querySelector('#dataTable tbody tr[data-row-key] td[data-col-index="0"]');
    syncGridRovingTabindex(cell); cell.focus();
  });
  const fromCell = await page.evaluate(() => ({
    tag: document.activeElement?.tagName, col: document.activeElement?.dataset?.colIndex,
  }));
  await page.keyboard.press("Control+Shift+KeyF");
  await sleep(400);
  const inPopup = await page.evaluate(() => document.activeElement?.id);
  await page.keyboard.press("Escape");
  await sleep(300);
  const backTo = await page.evaluate(() => ({
    tag: document.activeElement?.tagName, col: document.activeElement?.dataset?.colIndex,
    popupClosed: document.getElementById("quickSearchPopup").classList.contains("hidden"),
  }));
  check("Esc zamyka okno szukania nawet przy zaznaczonej komórce", backTo.popupClosed === true,
    JSON.stringify(backTo));
  check("…i oddaje fokus tam, skąd otwarto okno",
    backTo.tag === fromCell.tag && backTo.col === fromCell.col,
    `${JSON.stringify(fromCell)} → ${inPopup} → ${JSON.stringify(backTo)}`);

  // ── 14. Zasięg klawiatury w Safari/WebKit ───────────────────────────────────
  // Safari prowadzi Tab tylko po polach tekstowych i listach wyboru — przyciski
  // i pola wyboru wypadają z nawigacji, jeśli nie mają JAWNEGO tabindex.
  const reach = await page.evaluate(() => {
    const brak = [...document.querySelectorAll('button, input[type="checkbox"], input[type="radio"]')]
      .filter((el) => !el.hasAttribute("tabindex") && !el.closest(".hidden") && el.offsetParent !== null);
    return { bezTabindex: brak.length, przyklady: brak.slice(0, 3).map((el) => el.id || el.className) };
  });
  check("przyciski i pola wyboru mają jawny tabindex (inaczej Safari je pomija)",
    reach.bezTabindex === 0, JSON.stringify(reach));

  // Pełny obieg Tab w oknie szukania musi być IDENTYCZNY w każdym silniku
  await page.evaluate(() => { openQuickSearchPopup(); });
  await sleep(300);
  // Pasek chipów „Tryby auto" pojawia się w oknie tylko dla arkuszy z powtarzalnymi
  // blokami kolumn — obieg liczymy więc pod faktyczny stan okna, nie pod stałą listę.
  const chipBarVisible = await page.evaluate(() => {
    const el = document.getElementById("quickSearchSmartChips");
    return !!el && !el.classList.contains("hidden") && el.querySelectorAll("button").length > 0;
  });
  const obieg = [];
  for (let i = 0; i < (chipBarVisible ? 8 : 7); i++) {
    await page.keyboard.press("Tab");
    obieg.push(await page.evaluate(() => {
      const a = document.activeElement;
      // Na dotyku dwa pierwsze przystanki to grupy segmentów zamiast <select>
      if (a?.getAttribute && a.getAttribute("role") === "radio") {
        return a.closest(".seg-group")?.previousElementSibling?.id || "seg";
      }
      // Chipy trybów auto: cały pasek to JEDEN przystanek (roving tabindex, strzałki)
      const chipBar = a?.closest ? a.closest("#quickSearchSmartChips") : null;
      if (chipBar) return chipBar.id;
      return a?.id || "‹POZA›";
    }));
  }
  const oczekiwany = ["quickSearchPopupMode", "quickSearchPopupAction", "quickSearchPopupOperators",
    "qsAllSheetsPopup", "quickSearchPopupColumnsBtn", "quickSearchPopupBtn",
    ...(chipBarVisible ? ["quickSearchSmartChips"] : []),
    "quickSearchPopupInput"];
  check("obieg Tab w oknie szukania jest taki sam w każdym silniku i na dotyku",
    obieg.join(",") === oczekiwany.join(","), obieg.join(" → "));
  await page.evaluate(() => closeQuickSearchPopup());

  // ── 15. Segmenty zamiast <select> na dotyku (iPad) ──────────────────────────
  // Na iPadzie fokus na <select> otwiera systemowy picker i nie da sie tego
  // przechwycic — wiec na dotyku podmieniamy sama powierzchnie wejscia.
  await page.evaluate(() => { closeQuickSearchPopup(); openQuickSearchPopup(); });
  await sleep(300);
  const seg = await page.evaluate(() => {
    const sel = document.getElementById("quickSearchPopupAction");
    const grp = sel.nextElementSibling;
    return {
      coarse: matchMedia("(pointer: coarse)").matches,
      selectWidoczny: getComputedStyle(sel).display !== "none",
      segmentyWidoczne: !!grp && grp.classList.contains("seg-group") && getComputedStyle(grp).display !== "none",
      opcje: grp ? grp.querySelectorAll('[role="radio"]').length : 0,
      etykiety: [...document.querySelectorAll("#quickSearchPopup .seg-group")].map((g) => g.getAttribute("aria-label")),
    };
  });
  if (seg.coarse) {
    check("na dotyku <select> ustępuje miejsca segmentom",
      seg.segmentyWidoczne === true && seg.selectWidoczny === false, JSON.stringify(seg));
    check("grupy segmentów mają własne, różne etykiety",
      seg.etykiety.length === 2 && seg.etykiety[0] !== seg.etykiety[1], JSON.stringify(seg.etykiety));
    // Strzałka zmienia wartość ukrytego <select> — reszta kodu czyta ją bez zmian
    const przed = await page.evaluate(() => {
      const grp = document.getElementById("quickSearchPopupAction").nextElementSibling;
      grp.querySelector('[role="radio"][tabindex="0"]').focus();
      return document.getElementById("quickSearchPopupAction").value;
    });
    await page.keyboard.press("ArrowRight");
    await sleep(200);
    const po = await page.evaluate(() => document.getElementById("quickSearchPopupAction").value);
    check("strzałka na segmentach zmienia wartość źródłowego <select>", przed !== po, `${przed} → ${po}`);
    // Cała grupa to JEDEN przystanek Tab — inaczej lek byłby gorszy od choroby
    await page.evaluate(() => document.getElementById("quickSearchPopupInput").focus());
    const stops = [];
    for (let i = 0; i < 3; i++) {
      await page.keyboard.press("Tab");
      stops.push(await page.evaluate(() => {
        const a = document.activeElement;
        return a.getAttribute && a.getAttribute("role") === "radio"
          ? "seg:" + a.closest(".seg-group").getAttribute("aria-label") : (a.id || a.tagName);
      }));
    }
    check("każda grupa segmentów to jeden przystanek Tab",
      new Set(stops).size === stops.length, stops.join(" → "));
  } else {
    check("na laptopie zostają listy wyboru (segmenty ukryte)",
      seg.selectWidoczny === true && seg.segmentyWidoczne === false, JSON.stringify(seg));
  }
  await page.evaluate(() => closeQuickSearchPopup());

  // ── 16. Dotknięcie komórki oddaje siatce fokus DOM ──────────────────────────
  // Bez tego po tapnięciu JAKIEGOKOLWIEK przycisku fokus zostawał na nim, a komórka
  // wyglądała na zaznaczoną, choć pisanie i strzałki nad tabelą były martwe.
  // Objaw zgłoszony przez użytkownika: „zaznaczam Shift+tapem i nie mogę zacząć pisać".
  await page.evaluate(() => { closeQuickSearchPopup(); setSidebarOpen(false); });
  await sleep(300);
  const przyciskId = "excelLayoutToggle";
  if (TOUCH) await page.tap("#" + przyciskId); else await page.click("#" + przyciskId);
  await sleep(250);
  const naPrzycisku = await page.evaluate((id) => document.activeElement?.id === id, przyciskId);

  await page.keyboard.down("Shift");
  if (TOUCH) await page.tap('#dataTable tbody tr[data-row-key] td[data-col-index="2"]');
  else await page.click('#dataTable tbody tr[data-row-key] td[data-col-index="2"]', { modifiers: ["Shift"] });
  await page.keyboard.up("Shift");
  await sleep(300);
  const poKliku = await page.evaluate(() => ({
    tag: document.activeElement?.tagName,
    blokuje: shouldIgnoreTableArrowNavigation(),
    zaznaczona: !!focusedCellState,
  }));
  check("klik/tap w komórkę oddaje siatce fokus DOM",
    poKliku.tag === "TD" && poKliku.blokuje === false && poKliku.zaznaczona,
    `fokus przedtem na przycisku: ${naPrzycisku} → ${JSON.stringify(poKliku)}`);

  await page.keyboard.press("x");
  await sleep(350);
  const odRazuPisanie = await page.evaluate(() => ({
    edytor: !!document.querySelector(".cell-editor"),
    wartosc: document.querySelector(".cell-editor")?.value ?? null,
  }));
  check("po zaznaczeniu komórki można pisać OD RAZU (bez kilku prób)",
    odRazuPisanie.edytor === true && odRazuPisanie.wartosc === "x", JSON.stringify(odRazuPisanie));
  await page.keyboard.press("Escape");
  await sleep(200);

  const przedRuchem = await page.evaluate(() => focusedCellState && { ...focusedCellState });
  await page.keyboard.press("ArrowDown");
  await sleep(250);
  const poRuchu = await page.evaluate(() => focusedCellState && { ...focusedCellState });
  check("…i strzałki działają od razu po kliknięciu w komórkę",
    JSON.stringify(przedRuchem) !== JSON.stringify(poRuchu),
    `${JSON.stringify(przedRuchem)} → ${JSON.stringify(poRuchu)}`);

  // ── 17. Wiersz to wiersz, komórka to komórka ────────────────────────────────
  // Przy zaznaczonym WIERSZU żadna pojedyncza komórka nie może być dodatkowo
  // wyróżniona — wcześniej po ruchu strzałką kotwica dostawała obwódkę fokusu
  // i wyglądało to jak dwie konkurencyjne selekcje naraz.
  const wyglad = async () => page.evaluate(() => {
    const kotwica = document.querySelector(
      `tr[data-row-key="${CSS.escape(focusedCellState.rowKey)}"] td[data-col-index="${focusedCellState.colIndex0}"]`);
    const sasiad = [...kotwica.parentElement.querySelectorAll("td[data-col-index]")]
      .find((td) => td !== kotwica);
    return {
      kotwica: getComputedStyle(kotwica).boxShadow,
      sasiad: sasiad ? getComputedStyle(sasiad).boxShadow : null,
      tloKotwicy: getComputedStyle(kotwica).backgroundColor,
      tloSasiada: sasiad ? getComputedStyle(sasiad).backgroundColor : null,
    };
  });

  await page.evaluate(() => { setSidebarOpen(false); setSelectionKind("row"); });
  await sleep(200);
  if (TOUCH) await page.tap('#dataTable tbody tr[data-row-key] td[data-col-index="2"]');
  else await page.click('#dataTable tbody tr[data-row-key] td[data-col-index="2"]');
  await sleep(250);
  await page.keyboard.press("ArrowDown");
  await sleep(300);
  const wWierszu = await wyglad();
  check("w trybie wiersza kotwica wygląda jak sąsiednie komórki (brak drugiej selekcji)",
    wWierszu.kotwica === wWierszu.sasiad && wWierszu.tloKotwicy === wWierszu.tloSasiada,
    JSON.stringify(wWierszu).slice(0, 150));

  // …a w trybie komórki MUSI się wyróżniać, inaczej nie wiadomo, gdzie się stoi
  await page.keyboard.press("Shift+Space");
  await sleep(300);
  const wKomorce = await wyglad();
  check("w trybie komórki kotwica wyróżnia się od sąsiadów",
    wKomorce.kotwica !== wKomorce.sasiad, JSON.stringify(wKomorce).slice(0, 150));
  await page.evaluate(() => setSelectionKind("row"));

  console.log(`\n─── WYNIKI (${ENGINE}) ───`);
  let failed = 0;
  for (const r of results) {
    if (!r.pass) failed++;
    console.log(`${r.pass ? "✅" : "❌"} ${r.name}${r.detail ? `  [${r.detail}]` : ""}`);
  }
  console.log(`\nbłędy konsoli/strony: ${errors.length}`);
  errors.slice(0, 10).forEach((e) => console.log("   " + e));

  await browser.close();
  if (failed || errors.length) {
    console.log(`\n❌ ${failed} nieudanych asercji, ${errors.length} błędów`);
    process.exit(1);
  }
  console.log("\n✅ wszystko przeszło");
}

run().catch((e) => { console.error(e); process.exit(1); });
