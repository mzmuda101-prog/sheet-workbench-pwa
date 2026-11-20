// ribbon.js — wstążka jak w Excelu (docs/WSTAZKA.md; makieta docs/makiety/wstazka-v2.html).
//
// Od 769 px nagłówek „hero” i paski nad tabelą zamieniają się w dwa rzędy:
//   1. logo · Plik · karty · Szukaj · tryb · Zapisz · ⋯
//   2. Cofnij/Ponów + polecenia wybranej karty w jednej linii.
// Kontrolki to TE SAME elementy (te same id) — przenoszą się do gniazd [data-rb] wstążki, a na
// telefonie wracają na swoje miejsca (przez komentarz-znacznik), więc moduły i testy ich nie gubią.
// Mechanika skopiowana z Documents Workbench (app/ribbon.js tam) — ta sama rodzina aplikacji.
//
// Rząd, który się nie mieści, przewija się w bok (zasada Mateusza: nic nie chowa się pod „⋯”):
// krawędź wygasa, strzałka pokazuje, ile jest dalej („+3”), a za pierwszym razem w sesji rząd
// lekko drgnie w bok i wróci (bez ruchu przy „ogranicz ruch”).

const swbRibbon = (() => {
  const root = document.getElementById("ribbon");
  if (!root) return null;
  const deskMq = matchMedia("(min-width: 769px)");
  const reduceMq = matchMedia("(prefers-reduced-motion: reduce)");
  const HINT_KEY = "swb-rb-hinted-v1";
  const tr = (k) => (typeof t === "function" ? t(k) : k);

  // ── przenoszenie kontrolek ────────────────────────────────────────────────
  const moves = Array.from(root.querySelectorAll("[data-rb]")).map((slot) => {
    const el = document.querySelector(slot.dataset.rb);
    return el ? { el, slot, mark: document.createComment("swb-rb") } : null;
  }).filter(Boolean);
  // Poza wstążką (makieta v2): pasek formuły i chipy filtrów w JEDNYM rzędzie (chipy po prawej,
  // przy braku miejsca schodzą pod formułę), a „Ostatnie” zapytania — pod polem Szukaj.
  const editBars = document.getElementById("editBars");
  const fbarRow = document.createElement("div");
  fbarRow.id = "rbFbarRow";
  fbarRow.className = "rb-fbar-row";
  const extra = [
    // Etap 4 (dolny pasek, makieta v2): karty arkuszy na dole od lewej, przed nimi przycisk listy arkuszy
    // (przełączanie + Nowy / Zmień nazwę / Usuń, app-frame.js). „start” = na początek paska; kolejność
    // wstawiania odwrotna, więc najpierw karty, potem przycisk (ląduje przed nimi).
    [document.getElementById("sheetTabs"), document.getElementById("tableFoot"), "start"],
    [document.getElementById("heroSheetBtn"), document.getElementById("tableFoot"), "start"],
    [editBars, fbarRow],
    [document.getElementById("activeFilters"), fbarRow],
    [document.getElementById("recentStrip"), document.getElementById("quickSearchWrap")],
    // powiększenie na dolnym pasku, po prawej (jak w Excelu; WSTAZKA.md, etap 4)
    [document.getElementById("zoomGroup"), document.getElementById("tableFoot")],
  ].filter(([el, into]) => el && into).map(([el, slot, at]) => ({ el, slot, at, mark: document.createComment("swb-rb") }));
  moves.push(...extra);

  // ── telefon z plikiem (etap 7a, makieta v2 „Telefon”): pasek u góry zamiast dużej karty „hero” ──
  // Te same elementy co we wstążce — Tryb (z menu), Cofnij, Zapisz, ⋯ — a karty arkuszy jak na komputerze
  // pod tabelą. Pasek formuły tylko podczas edycji (w miejscu Szukaj; klasa ph-editing, CSS).
  const phTop = document.getElementById("phTop");
  const heroEl = document.querySelector(".hero");
  const foot = document.getElementById("tableFoot");
  const phMoves = !phTop ? [] : [
    ...Array.from(phTop.querySelectorAll("[data-ph]")).map((slot) => [document.querySelector(slot.dataset.ph), slot]),
    [document.getElementById("rbModeMenu"), phTop],
    [document.getElementById("sheetTabs"), foot, "start"],
    [document.getElementById("heroSheetBtn"), foot, "start"],
    // etap 7b: polecenia kart na dole — kontrolki z id wchodzą do gniazd kart (jak na komputerze),
    // a cały rząd poleceń (#rbCmds) do rzędu nad dolnym paskiem. Spisywanie / Wpisywanie do menu Trybu,
    // powiększenie na kartę Widok, Eksport do ⋯ (ekranu Plik na telefonie nie ma).
    ...Array.from(root.querySelectorAll(".rb-panel:not(.rb-cl) [data-rb], #rbModeMenu [data-rb]")).map((slot) => [document.querySelector(slot.dataset.rb), slot]),
    [document.getElementById("zoomGroup"), document.getElementById("rbPanel-view")],
    [document.getElementById("exportCsvBtn"), document.getElementById("appMenu")],
    [document.getElementById("rbCmds"), document.getElementById("phCmdsSlot")],
    // etap 7c: szyna paneli (komputer) na telefonie stoi na kartach — Sortowanie i Sprawdzanie na Danych,
    // Formuły na Formułach (Filtry ma przycisk w rzędzie Szukaj); linki z ekranu Plik do ⋯
    [document.querySelector('.rb-rail-btn[data-rb-rail="sort"]'), document.getElementById("rbPanel-data")],
    [document.querySelector('.rb-rail-btn[data-rb-rail="check"]'), document.getElementById("rbPanel-data")],
    [document.querySelector('.rb-rail-btn[data-rb-rail="formulas"]'), document.getElementById("rbPanel-fx")],
    ...["panel-file-sheet", "panel-shortcuts", "panel-log"].map((p) => [root.querySelector(`.rb-bs-link[data-rb-panel="${p}"]`), document.getElementById("appMenu")]),
  ].filter(([el, slot]) => el && slot).map(([el, slot, at]) => ({ el, slot, at, mark: document.createComment("swb-ph") }));
  let phOn = false;
  const hasFile = () => !!heroEl?.classList.contains("has-file");
  function placePhone(want) {
    if (!phTop || want === phOn) return;
    if (!want) { closePhSheet(); rideKeyBar(false); }
    phOn = want;
    closeMenus();
    for (const m of phMoves) {
      if (want) { m.el.before(m.mark); if (m.at === "start") m.slot.prepend(m.el); else m.slot.append(m.el); }
      else if (m.mark.parentNode) m.mark.replaceWith(m.el);
    }
    phTop.hidden = !want;
    if (phBottom) phBottom.hidden = !want;
    document.body.classList.toggle("ph-on", want);
    if (!want) document.body.classList.remove("ph-editing", "ph-reading");
    const pt = document.getElementById("panelToggle");
    if (pt) { if (want) pt.setAttribute("aria-label", tr("phFilters")); else pt.removeAttribute("aria-label"); }
    syncPhName();
    syncPhTab();
    syncKeyRide();
    document.dispatchEvent(new CustomEvent("swb:phone", { detail: { on: want } }));
    requestAnimationFrame(() => {
      if (typeof onViewportChange === "function") onViewportChange();
      window.dispatchEvent(new Event("resize"));
    });
  }
  function syncPhName() {
    if (!phOn) return;
    const f = document.getElementById("phFile"), s = document.getElementById("phSheetName");
    if (f) f.textContent = document.getElementById("heroTitle")?.textContent.trim() || "";
    if (s) s.textContent = typeof currentSheetName === "string" ? currentSheetName : "";
  }
  // ← „Wróć do plików”: to samo pytanie co przy „Aktualizuj” (Zapisz / Nie zapisuj / Anuluj), potem
  // ekran startowy z ostatnimi plikami. Przeładowanie zamiast ręcznego czyszczenia stanu wszystkich modułów.
  document.getElementById("phBackBtn")?.addEventListener("click", async () => {
    if (typeof confirmReloadWithUnsaved === "function" && !(await confirmReloadWithUnsaved())) return;
    window.location.reload();
  });
  // edycja komórki (pole w komórce albo pasek formuły) → pasek formuły w miejscu Szukaj
  const EDIT_SEL = "input.cell-editor, #formulaInput, #nameBox";
  const syncEditing = () => {
    const a = document.activeElement;
    const was = document.body.classList.contains("ph-editing");
    const now = phOn && !!a?.matches?.(EDIT_SEL);
    syncKeyRide();
    if (now === was) return;
    document.body.classList.toggle("ph-editing", now);
    if (typeof onViewportChange === "function") requestAnimationFrame(() => onViewportChange());
  };
  document.addEventListener("focusin", syncEditing);
  // pasek cyfr (key-bar.js) powstaje i pokazuje się chwilę po fokusie — sprawdź jeszcze raz; dalej pilnuje obserwator
  document.addEventListener("focusin", () => setTimeout(syncKeyRide, 50));
  document.addEventListener("focusout", () => setTimeout(syncEditing, 0));
  // pole w komórce znika z DOM po Enter / Esc — wtedy przeglądarka nie zgłasza utraty fokusu
  const later = () => { if (document.body.classList.contains("ph-editing")) setTimeout(syncEditing, 0); };
  document.addEventListener("keydown", later, true);
  document.addEventListener("pointerdown", later, true);
  // ── telefon, etap 7b: rząd poleceń, arkusz kart, jazda z paskiem cyfr ──
  const phBottom = document.getElementById("phBottom");
  const phTabBtn = document.getElementById("phTabBtn");
  const phSheet = document.getElementById("phSheet");
  const phSheetTabs = document.getElementById("phSheetTabs");
  const phSheetBody = document.getElementById("phSheetBody");
  const phCmdsSlot = document.getElementById("phCmdsSlot");
  const tabsList = root.querySelector(".rb-tabs");
  const tabsMark = document.createComment("swb-ph-tabs");
  const bottomMark = document.createComment("swb-ph-bottom");
  function syncPhTab() {
    const txt = phTabBtn?.querySelector(".ph-tab-txt");
    if (!txt) return;
    const tab = root.querySelector(`.rb-tab[data-rtab="${current}"]`) || phSheet?.querySelector(`.rb-tab[data-rtab="${current}"]`);
    txt.textContent = current === "home" ? tr("phTabHome") : (tab?.textContent.trim() || "");
    phSheet?.setAttribute("aria-label", tr("phSheetAria"));
  }
  // arkusz kart „w miejscu klawiatury”: klawiatura się chowa, karty jako pigułki, polecenia z podpisami
  function openPhSheet() {
    if (!phOn || !phSheet || !phSheet.hidden) return;
    if (document.activeElement && document.activeElement !== document.body) document.activeElement.blur();
    rideKeyBar(false);
    if (tabsList) { tabsList.before(tabsMark); phSheetTabs.append(tabsList); }
    if (rbCmds) phSheetBody.append(rbCmds);
    phSheet.hidden = false;
    document.body.classList.add("ph-sheet-open");
    phTabBtn?.setAttribute("aria-expanded", "true");
    phSheet.querySelector('.rb-tab[aria-selected="true"]')?.scrollIntoView({ inline: "center", block: "nearest" });
    if (typeof onViewportChange === "function") requestAnimationFrame(() => onViewportChange());
  }
  function closePhSheet() {
    if (!phSheet || phSheet.hidden) return;
    if (tabsMark.parentNode) tabsMark.replaceWith(tabsList);
    if (rbCmds && phCmdsSlot) phCmdsSlot.append(rbCmds);
    phSheet.hidden = true;
    document.body.classList.remove("ph-sheet-open");
    phTabBtn?.setAttribute("aria-expanded", "false");
    syncPhTab();
    if (typeof onViewportChange === "function") requestAnimationFrame(() => onViewportChange());
  }
  phTabBtn?.addEventListener("click", () => (phSheet?.hidden ? openPhSheet() : closePhSheet()));
  document.getElementById("phSheetClose")?.addEventListener("click", () => { closePhSheet(); phTabBtn?.focus(); });
  // przy otwartej klawiaturze (pasek cyfr widoczny, edycja komórki) rząd poleceń jedzie razem z paskiem cyfr
  function rideKeyBar(want) {
    const kb = document.querySelector(".key-bar");
    if (!phBottom) return;
    if (want && kb) { if (!bottomMark.parentNode) phBottom.before(bottomMark); kb.prepend(phBottom); }
    else if (bottomMark.parentNode) bottomMark.replaceWith(phBottom);
  }
  let kbObserved = null;
  function syncKeyRide() {
    const kb = document.querySelector(".key-bar");
    if (kb && kbObserved !== kb) { kbObserved = kb; new MutationObserver(syncKeyRide).observe(kb, { attributes: true, attributeFilter: ["hidden"] }); }
    const a = document.activeElement;
    rideKeyBar(phOn && !!kb && !kb.hidden && !!a?.matches?.(EDIT_SEL) && (!phSheet || phSheet.hidden));
  }

  // telefon: sumy na dolnym pasku domyślnie schowane — karty arkuszy dostają cały rząd; Σ pokazuje / chowa
  // (Mateusz 2026-11-20), wybór zapamiętany na urządzeniu
  const STATS_KEY = "swb-ph-stats-v1";
  const statsBtn = document.getElementById("statsToggle");
  const statsOn = () => { try { return localStorage.getItem(STATS_KEY) === "1"; } catch (_) { return false; } };
  function syncStats() {
    const v = statsOn();
    document.body.classList.toggle("ph-stats-on", v);
    if (statsBtn) {
      statsBtn.setAttribute("aria-pressed", String(v));
      statsBtn.setAttribute("aria-label", tr(v ? "phStatsHide" : "phStatsShow"));
      statsBtn.dataset.hint = "";
      statsBtn.dataset.hintPl = statsBtn.dataset.hintEn = tr(v ? "phStatsHide" : "phStatsShow");
    }
  }
  statsBtn?.addEventListener("click", () => {
    try { localStorage.setItem(STATS_KEY, statsOn() ? "0" : "1"); } catch (_) { /* tryb prywatny */ }
    syncStats();
  });
  syncStats();

  function syncDevice() {
    const desk = deskMq.matches;
    if (desk) placePhone(false); // najpierw oddaj kontrolki z paska telefonu, potem wstążka bierze swoje
    place(desk);
    if (!desk) placePhone(hasFile());
  }
  // bez pliku (etap 8): ekran startowy jak w Documents — wstążka i paski chowają polecenia (CSS: body.swb-nofile)
  const syncNoFile = () => {
    const was = document.body.classList.contains("swb-nofile");
    const now = !hasFile();
    if (was === now) return;
    document.body.classList.toggle("swb-nofile", now);
    // wstążka zmienia wysokość (rząd poleceń znika / wraca) — przypięty panel i tabela liczą się od nowa
    requestAnimationFrame(() => {
      if (typeof onViewportChange === "function") onViewportChange();
      window.dispatchEvent(new Event("resize"));
    });
  };
  syncNoFile();
  if (heroEl) new MutationObserver(() => { syncNoFile(); if (!deskMq.matches) placePhone(hasFile()); }).observe(heroEl, { attributes: true, attributeFilter: ["class"] });
  const sheetNameEl = document.getElementById("heroSheetName");
  if (sheetNameEl) new MutationObserver(syncPhName).observe(sheetNameEl, { childList: true, characterData: true, subtree: true });

  let on = false;
  function place(desk) {
    if (desk === on) return;
    on = desk;
    // przeniesienie elementu z fokusem go rozmywa (obrót iPada w trakcie pisania) — oddajemy fokus
    const had = document.activeElement;
    if (desk && editBars) editBars.before(fbarRow);
    for (const m of moves) {
      if (desk) { m.el.before(m.mark); if (m.at === "start") m.slot.prepend(m.el); else m.slot.append(m.el); }
      else if (m.mark.parentNode) { m.mark.replaceWith(m.el); }
    }
    if (!desk) fbarRow.remove();
    classicNow = false; // na telefonie kontrolki wróciły do siebie (znaczniki), Klasyczna zaczyna od nowa
    applyLayout(); // na telefonie zdejmuje też body.rb-classic
    // kontrolki, które dopiero weszły do wstążki (blok formatu z podpisami grup), dostają napisy w bieżącym języku
    if (desk) labels();
    root.hidden = !desk;
    document.body.classList.toggle("rb-on", desk);
    // dopiero po zmianie klasy: wcześniej pole w chowanym pasku nie ma jeszcze rozmiaru
    if (had && had !== document.body && had.isConnected && document.activeElement !== had && had.getClientRects().length) had.focus({ preventScroll: true });
    if (desk) requestAnimationFrame(() => strips.forEach((s) => s.sync(true)));
    // wysokość tabeli liczy się od jej górnej krawędzi — po zmianie ramy przelicz
    requestAnimationFrame(() => {
      if (typeof onViewportChange === "function") onViewportChange();
      window.dispatchEvent(new Event("resize"));
    });
    document.dispatchEvent(new CustomEvent("swb:ribbon", { detail: { on: desk } }));
  }

  // ── układ Klasyczna / Uproszczona (etap 6b; ribbon-classic.js buduje panele Klasycznej) ──
  // Kontrolki z id stoją w Klasycznej w gniazdach [data-rbc] — przy przełączeniu przenoszą się
  // między gniazdami obu układów; na telefonie wracają na swoje miejsce (znaczniki z place()).
  const clSlots = (typeof swbRibbonClassic !== "undefined" && swbRibbonClassic?.slots) || [];
  let classicNow = false;
  function applyLayout() {
    const want = on && prefs().layout === "classic" && clSlots.length > 0;
    document.body.classList.toggle("rb-classic", want);
    if (want !== classicNow) {
      classicNow = want;
      for (const s of clSlots) {
        const el = document.querySelector(s.dataset.rbc);
        if (!el) continue;
        if (want) s.append(el);
        else if (el.parentNode === s) root.querySelector(`.rb-panel:not(.rb-cl) [data-rb="${CSS.escape(s.dataset.rbc)}"]`)?.append(el);
      }
    }
    select(current);
  }

  // ── karty ─────────────────────────────────────────────────────────────────
  const tabs = Array.from(root.querySelectorAll(".rb-tab[data-rtab]"));
  const panels = Array.from(root.querySelectorAll(".rb-panel[data-rtab]"));
  let current = "home";
  function select(id, { focus = false } = {}) {
    const tab = tabs.find((x) => x.dataset.rtab === id && !x.hidden);
    if (!tab) return;
    current = id;
    tabs.forEach((x) => {
      const sel = x === tab;
      x.setAttribute("aria-selected", String(sel));
      x.tabIndex = sel ? 0 : -1;
    });
    panels.forEach((p) => { p.hidden = p.dataset.rtab !== id || p.classList.contains("rb-cl") !== classicNow; });
    // karta wskazuje panel bieżącego układu (czytnik ekranu: „przejdź do sterowanego elementu”)
    tabs.forEach((x) => x.setAttribute("aria-controls", `${classicNow ? "rbPanelCl" : "rbPanel"}-${x.dataset.rtab}`));
    const cmds = strips.find((s) => s.key === "cmds");
    if (cmds) { cmds.sc.scrollLeft = 0; cmds.sync(); }
    tabStrip?.reveal(tab);
    if (focus) tab.focus();
    syncPhTab();
  }
  tabs.forEach((x) => x.addEventListener("click", () => onTabClick(x.dataset.rtab)));
  // strzałki bez animacji — przełączanie kart to ruch powtarzany setki razy
  root.querySelector('[role="tablist"]')?.addEventListener("keydown", (e) => {
    if (!["ArrowLeft", "ArrowRight", "Home", "End"].includes(e.key)) return;
    const vis = tabs.filter((x) => !x.hidden);
    const i = vis.indexOf(document.activeElement);
    if (i < 0) return;
    e.preventDefault();
    const n = e.key === "Home" ? 0 : e.key === "End" ? vis.length - 1 : (i + (e.key === "ArrowRight" ? 1 : -1) + vis.length) % vis.length;
    select(vis[n].dataset.rtab, { focus: true });
  });

  // ── przewijany rząd z podpowiedzią ───────────────────────────────────────
  function hinted() { try { return JSON.parse(sessionStorage.getItem(HINT_KEY) || "[]"); } catch (_) { return []; } }
  function markHinted(key) { try { sessionStorage.setItem(HINT_KEY, JSON.stringify([...hinted(), key])); } catch (_) {} }
  function makeStrip(el) {
    const sc = el.querySelector(".rb-sc");
    const left = el.querySelector(".rb-nudge-l");
    const right = el.querySelector(".rb-nudge-r");
    const cntR = right?.querySelector(".rb-cnt");
    const cntL = left?.querySelector(".rb-cnt");
    const key = el.dataset.strip;
    const items = () => Array.from(sc.querySelectorAll(key === "tabs"
      ? ".rb-tab:not([hidden])"
      : ".rb-panel:not([hidden]) > *:not(.rb-slot), .rb-panel:not([hidden]) > .rb-slot > * > *, .rb-panel:not([hidden]) > .rb-slot > *:not(.format-ribbon)"))
      .filter((c) => c.offsetParent && !c.matches(".fmt-sep, .rb-sep, .fmt-gcap"));
    function sync(first) {
      if (!on) return;
      const max = sc.scrollWidth - sc.clientWidth;
      const L = max > 1 && sc.scrollLeft > 2;
      const R = max > 1 && sc.scrollLeft < max - 2;
      el.classList.toggle("has-l", L);
      el.classList.toggle("has-r", R);
      const S = sc.getBoundingClientRect();
      const all = (L || R) ? items() : [];
      const nR = R ? all.filter((c) => c.getBoundingClientRect().right > S.right + 1).length : 0;
      const nL = L ? all.filter((c) => c.getBoundingClientRect().left < S.left - 1).length : 0;
      if (cntR) cntR.textContent = nR ? `+${nR}` : "";
      if (cntL) cntL.textContent = nL ? `+${nL}` : "";
      if (first && R && !hinted().includes(key)) {
        markHinted(key);
        right.classList.add("is-hint");
        setTimeout(() => right.classList.remove("is-hint"), 1100);
        if (!reduceMq.matches) {
          sc.scrollBy({ left: 36, behavior: "smooth" });
          setTimeout(() => sc.scrollBy({ left: -36, behavior: "smooth" }), 420);
        }
      }
    }
    // z prostokątów, nie offsetLeft: kontrolka w gnieździe .rb-slot (display: contents) ma offsetLeft 0
    function reveal(node) {
      if (!node) return;
      const r = node.getBoundingClientRect(), S = sc.getBoundingClientRect();
      if (!r.width) return;
      if (r.left < S.left) sc.scrollLeft += r.left - S.left - 24;
      else if (r.right > S.right) sc.scrollLeft += r.right - S.right + 40;
    }
    sc.addEventListener("scroll", () => sync(), { passive: true });
    // pionowe kółko przewija rząd w bok (Shift+kółko działa natywnie)
    sc.addEventListener("wheel", (e) => {
      if (Math.abs(e.deltaY) <= Math.abs(e.deltaX) || sc.scrollWidth <= sc.clientWidth + 1) return;
      sc.scrollLeft += e.deltaY;
      e.preventDefault();
    }, { passive: false });
    left?.addEventListener("click", () => sc.scrollBy({ left: -sc.clientWidth * 0.7, behavior: reduceMq.matches ? "auto" : "smooth" }));
    right?.addEventListener("click", () => sc.scrollBy({ left: sc.clientWidth * 0.7, behavior: reduceMq.matches ? "auto" : "smooth" }));
    // fokus z klawiatury na schowanym przycisku — rząd sam do niego dojeżdża
    sc.addEventListener("focusin", (e) => { reveal(e.target.closest(".rb-tab") || e.target); });
    let rq = 0;
    const later = () => { cancelAnimationFrame(rq); rq = requestAnimationFrame(() => sync()); };
    if (typeof ResizeObserver === "function") new ResizeObserver(later).observe(sc);
    return { key, sc, sync, reveal };
  }
  const strips = Array.from(root.querySelectorAll(".rb-strip")).map(makeStrip);
  const tabStrip = strips.find((s) => s.key === "tabs");

  // ── polecenia kart ───────────────────────────────────────────────────────
  // Każdy przycisk woła TO SAMO, co dotąd robił skrót, menu komórki albo panel (jedna logika).
  const ACTS = {
    paste: () => typeof pasteClipboardToSelection === "function" && pasteClipboardToSelection(),
    cut: () => typeof cutSelectionToClipboard === "function" && cutSelectionToClipboard(),
    copy: () => typeof copySelectionToClipboard === "function" && copySelectionToClipboard(),
    // Shift+F2 (cell-notes.js) — ta sama droga co skrót, na bieżącej komórce
    note: () => document.dispatchEvent(new KeyboardEvent("keydown", { key: "F2", shiftKey: true, bubbles: true, cancelable: true })),
    // jak „Wyjaśnij formułę” w menu komórki (filter-bar.js), dla komórki z kursorem
    // ten sam fill co dawny pasek dotykowy i Ctrl/⌘+D / R (table.js fillSelection)
    fillDown: () => typeof fillSelection === "function" && fillSelection("down"),
    fillRight: () => typeof fillSelection === "function" && fillSelection("right"),
    explain: () => {
      const row = typeof focusedCellState !== "undefined" && focusedCellState && typeof getRowByKey === "function" ? getRowByKey(focusedCellState.rowKey) : null;
      const item = row && typeof cellMenuExplainItem === "function" ? cellMenuExplainItem({ row, col: focusedCellState.colIndex0 }) : null;
      if (item) item.run();
      else if (typeof toast === "function") toast(tr("rbNoFormula"), "info");
    },
  };
  // polecenia kart (#rbCmds) mają własną obsługę — na telefonie (etap 7b) stoją poza wstążką
  const rbCmds = document.getElementById("rbCmds");
  function onCmd(e) {
    const b = e.target.closest("[data-rb-act], [data-rb-panel]");
    if (!b) return;
    if (b.dataset.rbAct) ACTS[b.dataset.rbAct]?.();
    else { closeMenus(); closePhSheet(); openPanel(b.dataset.rbPanel); }
  }
  rbCmds?.addEventListener("click", onCmd);
  root.addEventListener("click", (e) => { if (!rbCmds?.contains(e.target)) onCmd(e); });

  // ── menu: tryb i Plik ────────────────────────────────────────────────────
  const modeBtn = document.getElementById("rbModeBtn");
  const modeMenu = document.getElementById("rbModeMenu");
  const editBtn = document.getElementById("rbEditBtn");
  const nameEl = document.getElementById("rbName");
  const modeSwitch = document.getElementById("sheetModeSwitch");
  const optsBtn = document.getElementById("rbOptsBtn");
  const optsMenu = document.getElementById("rbOptsMenu");
  const MENUS = [[modeBtn, modeMenu, "right"], [optsBtn, optsMenu, "right"]].filter(([x, y]) => x && y);
  const menuItems = (m) => Array.from(m.querySelectorAll(".rb-mi, .rb-slot > button")).filter((x) => !x.hidden);
  function closeMenus(refocus) {
    for (const [btn, menu] of MENUS) {
      if (menu.hidden) continue;
      menu.hidden = true;
      btn.setAttribute("aria-expanded", "false");
      if (refocus) btn.focus();
    }
  }
  function openMenu(btn, menu, side) {
    closeMenus();
    menu.hidden = false;
    btn.setAttribute("aria-expanded", "true");
    const r = btn.getBoundingClientRect(), R = root.getBoundingClientRect();
    menu.style.top = `${Math.round(r.bottom - R.top + 6)}px`;
    if (side === "right") { menu.style.right = `${Math.max(8, Math.round(R.right - r.right))}px`; menu.style.left = ""; menu.style.transformOrigin = "top right"; }
    else { menu.style.left = `${Math.max(8, Math.round(r.left - R.left))}px`; menu.style.right = ""; menu.style.transformOrigin = "top left"; }
    // menu opcji: fokus na bieżącym sposobie pokazywania (najczęstszy wybór), nie na układzie
    (menu.querySelector('[data-rb-show][aria-checked="true"]') || menu.querySelector('[aria-checked="true"]') || menuItems(menu)[0])?.focus();
  }
  for (const [btn, menu, side] of MENUS) {
    btn.addEventListener("click", () => (menu.hidden ? openMenu(btn, menu, side) : closeMenus()));
    // po wyborze fokus wraca na przycisk menu — chyba że akcja sama go gdzieś postawiła (Spisywanie, okno)
    menu.addEventListener("click", (e) => {
      if (!e.target.closest("button")) return;
      setTimeout(() => {
        const a = document.activeElement;
        const lost = !a || a === document.body || menu.contains(a);
        closeMenus(false);
        if (lost && btn.getClientRects().length) btn.focus();
      }, 0);
    });
    menu.addEventListener("keydown", (e) => {
      const items = menuItems(menu);
      const i = items.indexOf(document.activeElement);
      if (e.key === "Escape") { e.preventDefault(); e.stopPropagation(); closeMenus(true); }
      else if (e.key === "ArrowDown") { e.preventDefault(); items[(i + 1) % items.length]?.focus(); }
      else if (e.key === "ArrowUp") { e.preventDefault(); items[(i - 1 + items.length) % items.length]?.focus(); }
      else if (e.key === "Tab") closeMenus();
    });
  }
  document.addEventListener("pointerdown", (e) => {
    if (MENUS.some(([btn, menu]) => !menu.hidden && (menu.contains(e.target) || btn.contains(e.target)))) return;
    closeMenus();
  }, true);

  // Edycja / Podgląd: klikamy TE SAME przyciski #sheetModeSwitch (sheet-edit-ui.js trzyma stan)
  const MODE_ICONS = {
    edit: '<path d="M12 20h9"/><path d="M16.5 3.5a2.1 2.1 0 0 1 3 3L7 19l-4 1 1-4Z"/>',
    view: '<path d="M2 12s3.5-7 10-7 10 7 10 7-3.5 7-10 7S2 12 2 12Z"/><circle cx="12" cy="12" r="3"/>',
  };
  const isView = () => typeof swbIsEditMode === "function" && !swbIsEditMode();
  function setMode(m) { modeSwitch?.querySelector(`.mode-btn[data-mode="${m}"]`)?.click(); }
  modeMenu?.querySelectorAll("[data-rb-mode]").forEach((b) => b.addEventListener("click", () => setMode(b.dataset.rbMode)));
  // „Edytuj” chowa się po kliknięciu — fokus na przycisk trybu, nie w próżnię
  editBtn?.addEventListener("click", () => { setMode("edit"); setTimeout(() => modeBtn?.focus(), 0); });
  // polecenia działające na komórkach — bez pliku wyszarzone (komunikat „tylko w trybie szerokim” mylił)
  function syncFile() {
    const has = typeof workbook !== "undefined" && !!workbook;
    // #rbCmds na telefonie stoi poza wstążką (etap 7b) — jego polecenia też
    document.querySelectorAll("#ribbon [data-rb-act], #rbCmds [data-rb-act]").forEach((b) => { b.disabled = !has; });
  }
  function syncMode() {
    syncFile();
    const m = isView() ? "view" : "edit";
    document.body.classList.toggle("rb-reading", on && m === "view");
    document.body.classList.toggle("ph-reading", phOn && m === "view");
    if (phOn && m === "view") closePhSheet();
    modeMenu?.querySelectorAll("[data-rb-mode]").forEach((b) => b.setAttribute("aria-checked", String(b.dataset.rbMode === m)));
    if (modeBtn) {
      modeBtn.dataset.mode = m;
      modeBtn.querySelector(".rb-mode-ico").innerHTML = `<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">${MODE_ICONS[m]}</svg>`;
      const label = tr(m === "view" ? "rbModeView" : "rbModeEdit");
      modeBtn.querySelector(".rb-mode-txt").textContent = label;
      modeBtn.setAttribute("aria-label", `${tr("rbModeAria")}: ${label}`);
    }
    if (editBtn) editBtn.hidden = m !== "view";
    if (nameEl) {
      nameEl.hidden = m !== "view";
      const file = document.getElementById("heroTitle")?.textContent.trim() || "";
      const sheet = typeof currentSheetName === "string" ? currentSheetName : "";
      nameEl.textContent = [file, sheet, tr("rbReadOnly")].filter(Boolean).join(" · ");
    }
    strips.forEach((x) => x.sync());
    syncPhName();
  }
  if (modeSwitch) new MutationObserver(syncMode).observe(modeSwitch, { subtree: true, attributes: true, attributeFilter: ["class", "aria-pressed"] });
  const titleEl = document.getElementById("heroTitle");
  if (titleEl) new MutationObserver(syncMode).observe(titleEl, { childList: true, characterData: true, subtree: true });

  // ── teksty (PL/EN) ────────────────────────────────────────────────────────
  function labels() {
    root.querySelectorAll("[data-rb-label]").forEach((el) => {
      const txt = tr(el.dataset.rbLabel);
      if (el.matches("button")) {
        el.setAttribute("aria-label", txt);
        const sp = el.querySelector(":scope > span:not([class])");
        if (sp) sp.textContent = txt;
        el.dataset.hint = "";
        el.dataset.hintPl = el.dataset.hintEn = el.dataset.rbHint ? `${txt} (${el.dataset.rbHint})` : txt;
      } else el.textContent = txt;
    });
    modeMenu?.setAttribute("aria-label", tr("rbModeMenuAria"));
    optsMenu?.setAttribute("aria-label", tr("rbOpts"));
    const ed = editBtn?.querySelector("span");
    if (ed) ed.textContent = tr("rbEdit");
    syncMode();
    const file = document.getElementById("rbFileBtn");
    if (file) file.textContent = tr("rbFile");
    tabs.forEach((x) => { x.textContent = tr(`rbTab_${x.dataset.rtab}`); });
    root.querySelector('[role="tablist"]')?.setAttribute("aria-label", tr("rbTabsAria"));
    root.querySelectorAll(".rb-nudge-l").forEach((b) => b.setAttribute("aria-label", tr("rbScrollLeft")));
    root.querySelectorAll(".rb-nudge-r").forEach((b) => b.setAttribute("aria-label", tr("rbScrollRight")));
    if (typeof swbRibbonClassic !== "undefined") swbRibbonClassic?.labels();
    strips.forEach((s) => s.sync());
  }

  // Szukaj zwinięte do lupy (wąskie okno): klik gdziekolwiek w nim = fokus w polu → karta się rozwija
  const searchSlot = root.querySelector(".rb-search");
  // click, nie pointerdown: mousedown na ikonie i tak zabiera fokus — dajemy go po nim
  searchSlot?.addEventListener("click", () => {
    const qs = document.getElementById("quickSearch");
    if (!qs || searchSlot.contains(document.activeElement)) return;
    if ((qs.closest(".qs-field") || qs).getBoundingClientRect().width > 4) return; // pole widać — zwykły klik
    qs.focus();
  });
  // karta Szukaj rozwinięta przez ⚙ zwija się po kliknięciu poza nią (szeroka tylko, gdy w niej pracujesz);
  // okna i menu otwarte z karty (np. Kolumny) jej nie zwijają
  document.addEventListener("pointerdown", (e) => {
    const wrap = document.getElementById("quickSearchWrap");
    const more = document.getElementById("qsMoreBtn");
    if (!on || !wrap?.classList.contains("qs-expanded") || wrap.contains(e.target)) return;
    if (e.target.closest?.('[role="dialog"], [role="menu"], [role="listbox"], .app-menu')) return;
    wrap.classList.remove("qs-expanded");
    more?.setAttribute("aria-expanded", "false");
  }, true);

  // Esc w polu Szukaj (wstążka): wyjście z pola — karta Szukaj się zwija, także z opcjami ⚙.
  // Podpowiedzi tokenów i inne nakładki, które same łapią Esc (stopPropagation), zamykają się pierwsze.
  const qsWrap = document.getElementById("quickSearchWrap");
  qsWrap?.addEventListener("keydown", (e) => {
    if (e.key !== "Escape" || !on || e.defaultPrevented && e.target.id !== "quickSearch") return;
    if (qsWrap.classList.contains("qs-expanded")) document.getElementById("qsMoreBtn")?.click();
    if (qsWrap.contains(document.activeElement)) document.activeElement.blur();
  });

  // ── panel boczny: otwórz sekcję (do etapu 3 — potem szyna z wysuwanym panelem) ──
  // telefon (etap 7c): Filtry w rzędzie Szukaj otwiera panel filtrów (jak przycisk szyny), nie cały panel boczny
  document.getElementById("panelToggle")?.addEventListener("click", (e) => {
    if (!phOn || typeof swbRail === "undefined" || !swbRail?.showKey) return;
    e.stopImmediatePropagation();
    e.preventDefault();
    swbRail.showKey("filters");
  }, true);
  // linki do paneli poza wstążką i rzędem poleceń (na telefonie: Wczytywanie / Skróty / Dziennik w ⋯)
  document.addEventListener("click", (e) => {
    const b = e.target.closest?.("[data-rb-panel]");
    if (!b || root.contains(b) || rbCmds?.contains(b)) return;
    const menu = b.closest("#appMenu");
    if (menu) { menu.hidden = true; document.getElementById("appMenuBtn")?.setAttribute("aria-expanded", "false"); }
    openPanel(b.dataset.rbPanel);
  });

  function openPanel(detailsId) {
    const d = document.getElementById(detailsId);
    if (!d) return;
    // etap 3: jeden panel naraz, z tytułem (rail.js); bez rail.js — jak dotąd
    if (typeof swbRail !== "undefined" && swbRail && (on || phOn)) { swbRail.show([detailsId]); return; }
    if (typeof setSidebarOpen === "function") setSidebarOpen(true);
    d.open = true;
    requestAnimationFrame(() => d.scrollIntoView({ block: "nearest" }));
  }

  // ── opcje wyświetlania (etap 6a): Zawsze / Chowaj przy przewijaniu / Tylko karty, gęstość ──
  // Menu na końcu rzędu poleceń i Plik → Opcje; zapamiętane na urządzeniu. Bez zapisanego wyboru
  // domyślne zależą od urządzenia (decyzja Mateusza 2026-11-19): sam dotyk (iPad bez myszki) =
  // Dotyk, a trzymany pionowo także Chowaj przy przewijaniu; mysz = Zawsze + Mysz.
  const PREFS_KEY = "swb-rb-prefs-v1";
  const SHOWS = ["pin", "auto", "tabs"];
  const fineMq = matchMedia("(any-pointer: fine)");
  const portraitMq = matchMedia("(orientation: portrait)");
  const saved = (() => {
    try { const p = JSON.parse(localStorage.getItem(PREFS_KEY) || "{}"); return { show: SHOWS.includes(p.show) ? p.show : null, dens: ["mouse", "touch"].includes(p.dens) ? p.dens : null, layout: p.layout === "classic" ? "classic" : null }; }
    catch (_) { return { show: null, dens: null, layout: null }; }
  })();
  const touchOnly = () => !fineMq.matches;
  const prefs = () => ({
    show: saved.show || (touchOnly() && portraitMq.matches ? "auto" : "pin"),
    dens: saved.dens || (touchOnly() ? "touch" : "mouse"),
    layout: saved.layout || "simple",
  });
  const SHOW_ICONS = {};
  optsMenu?.querySelectorAll("[data-rb-show]").forEach((b) => { SHOW_ICONS[b.dataset.rbShow] = b.querySelector("svg")?.innerHTML || ""; });
  const r2 = root.querySelector(".rb-r2");
  let peek = false;
  let lastLayout = "";
  // zwinięcie/rozwinięcie rzędu poleceń: wysokość tabeli przeliczamy po końcu przejścia (transitionend),
  // a bez przejścia (ogranicz ruch) — od razu, bo transitionend wtedy nie przychodzi
  function tuck(v) {
    const b = document.body;
    if (b.classList.contains("rb-tucked") === v) return;
    b.classList.toggle("rb-tucked", v);
    if (reduceMq.matches) relayout();
  }
  function syncPrefs() {
    const p = prefs();
    const b = document.body;
    SHOWS.forEach((v) => b.classList.toggle(`rb-show-${v}`, p.show === v));
    b.classList.toggle("rb-touch", p.dens === "touch");
    if (p.show !== "tabs") peek = false;
    b.classList.toggle("rb-peek", peek);
    if (p.show !== "auto") tuck(false);
    const svg = optsBtn?.querySelector("svg");
    if (svg) svg.innerHTML = SHOW_ICONS[p.show] || "";
    const mark = (sel, key, val) => document.querySelectorAll(sel).forEach((x) => x.setAttribute("aria-checked", String(x.dataset[key] === val)));
    mark("#rbOptsMenu [data-rb-show], #rbShowSeg [data-rb-show]", "rbShow", p.show);
    mark("#rbDensSeg [data-rb-dens]", "rbDens", p.dens);
    mark("#rbOptsMenu [data-rb-layout], #rbLayoutSeg [data-rb-layout]", "rbLayout", p.layout);
    // wysokość tabeli liczy się od jej górnej krawędzi — po zmianie rzędu poleceń przelicz
    // wysokość rzędu poleceń zależy od pokazywania i gęstości — tabela liczy się od jej górnej krawędzi
    if (`${p.show}|${p.dens}|${p.layout}` !== lastLayout) { lastLayout = `${p.show}|${p.dens}|${p.layout}`; relayout(); }
    requestAnimationFrame(() => strips.forEach((st) => st.sync()));
    document.dispatchEvent(new CustomEvent("swb:rbprefs", { detail: p }));
  }
  function relayout() {
    requestAnimationFrame(() => {
      if (typeof onViewportChange === "function") onViewportChange();
      window.dispatchEvent(new Event("resize"));
    });
  }
  function save(k, v) {
    saved[k] = v;
    try { localStorage.setItem(PREFS_KEY, JSON.stringify(saved)); } catch (_) {}
    syncPrefs();
  }
  const setShow = (v) => { peek = false; save("show", v); };
  const setDens = (v) => save("dens", v);
  const setLayout = (v) => { saved.layout = v === "classic" ? "classic" : null; save("layout", saved.layout); applyLayout(); };
  function setPeek(v) {
    if (peek === v) return;
    peek = v;
    syncPrefs();
    if (v) requestAnimationFrame(() => strips.find((x) => x.key === "cmds")?.sync());
  }
  function onTabClick(id) {
    // Tylko karty: klik karty wysuwa polecenia, ta sama karta drugi raz je chowa
    if (prefs().show === "tabs") setPeek(!(peek && current === id));
    tuck(false);
    select(id);
  }
  optsMenu?.addEventListener("click", (e) => { const b = e.target.closest("[data-rb-layout]"); if (b) setLayout(b.dataset.rbLayout); });
  document.getElementById("rbLayoutSeg")?.addEventListener("click", (e) => { const b = e.target.closest("[data-rb-layout]"); if (b) setLayout(b.dataset.rbLayout); });
  optsMenu?.addEventListener("click", (e) => {
    const b = e.target.closest("[data-rb-show]");
    if (!b) return;
    setShow(b.dataset.rbShow);
    // Tylko karty chowa rząd z przyciskiem menu — fokus na wybraną kartę, nie w próżnię (po powrocie menu)
    if (b.dataset.rbShow === "tabs") setTimeout(() => { const a = document.activeElement; if (!optsBtn?.getClientRects().length && (!a || a === document.body || !a.getClientRects().length)) tabs.find((x) => x.getAttribute("aria-selected") === "true")?.focus(); }, 0);
  });
  document.getElementById("rbShowSeg")?.addEventListener("click", (e) => { const b = e.target.closest("[data-rb-show]"); if (b) setShow(b.dataset.rbShow); });
  document.getElementById("rbDensSeg")?.addEventListener("click", (e) => { const b = e.target.closest("[data-rb-dens]"); if (b) setDens(b.dataset.rbDens); });
  // Ctrl/⌘+F1: Tylko karty ↔ Zawsze (jak w Excelu); Esc chowa wysunięte polecenia
  document.addEventListener("keydown", (e) => {
    if (!on) return;
    if (e.key === "F1" && swbMod(e) && !e.altKey && !e.shiftKey) {
      e.preventDefault();
      setShow(prefs().show === "tabs" ? "pin" : "tabs");
    } else if (e.key === "Escape" && peek) setPeek(false);
  });
  // wysunięte polecenia chowa klik poza wstążką (nie w menu i podpowiedzi, które z niej wyszły)
  document.addEventListener("pointerdown", (e) => {
    if (peek && !e.target.closest?.("#ribbon, .app-menu, .fmt-pop, .cursor-hint, [role='menu'], [role='dialog']")) setPeek(false);
  }, true);
  // Chowaj przy przewijaniu: tabela w dół — rząd poleceń się zwija, w górę — wraca. Po zmianie
  // wysokości tabela rośnie/maleje, co samo przesuwa scrollTop przy dole — przez chwilę to ignorujemy.
  const wrapEl = document.getElementById("tableWrap");
  let lastTop = 0, quietUntil = 0;
  wrapEl?.addEventListener("scroll", () => {
    const top = wrapEl.scrollTop, d = top - lastTop;
    lastTop = top;
    if (!on || prefs().show !== "auto" || Date.now() < quietUntil) return;
    const b = document.body;
    const tucked = b.classList.contains("rb-tucked");
    // nie chowamy rzędu z widocznym fokusem z klawiatury (WCAG 2.4.11); fokus po kliknięciu myszą nie blokuje
    if ((d > 6 && top > 40 && !tucked && !r2?.querySelector(":focus-visible")) || ((d < -6 || top < 10) && tucked)) {
      tuck(!tucked);
      quietUntil = Date.now() + 350;
    }
  }, { passive: true });
  // po zwinięciu/rozwinięciu (koniec przejścia) tabela dostaje nową wysokość
  r2?.addEventListener("transitionend", (e) => { if (e.target === r2 && e.propertyName === "block-size") relayout(); });
  r2?.addEventListener("focusin", () => tuck(false)); // Tab do poleceń
  const onDevice = () => { if (!saved.show || !saved.dens) syncPrefs(); };
  for (const mq of [fineMq, portraitMq]) { if (mq.addEventListener) mq.addEventListener("change", onDevice); else mq.addListener(onDevice); }

  const onMq = () => {
    closeMenus();
    // menu paska (frame-compact.js) przypięte do przycisku, który właśnie zmienia miejsce
    if (typeof frameCompact !== "undefined" && frameCompact?.closeMenu) frameCompact.closeMenu(false);
    syncDevice();
    syncMode();
  };
  if (deskMq.addEventListener) deskMq.addEventListener("change", onMq); else deskMq.addListener(onMq);
  syncPrefs();
  syncDevice();
  labels();
  document.addEventListener("keydown", (e) => { if (e.key === "Escape" && MENUS.some(([, m]) => !m.hidden)) { e.preventDefault(); closeMenus(true); } });

  return { isOn: () => on, select, current: () => current, openPanel, labels, closeMenus, prefs: () => ({ ...prefs(), peek }), setShow, setDens, setLayout };
})();
