// frame-compact.js — więcej miejsca na tabelę (makieta zaakceptowana 2026-10-10).
//
// Zmierzone przed zmianą: laptop 1280×720 widział 6 wierszy (tabela od 346 px), iPad
// poziomo 8, telefon 11. Winne: 3–4 rzędy pasków nad tabelą, wysoki nagłówek i panel
// zabierający ~330 px szerokości. Tu:
//  1. Menu „Widok ▾” (Wide-to-Long, Wymiary z Excela, Domyślne sortowanie) i „Więcej ▾”
//     (Spisywanie, Raport, Eksport) — te same przyciski co dawniej w osobnym rzędzie.
//  2. Powiększenie − 100% + na pasku, a pod procentem gęstość wierszy (Wygodna = jak dotąd,
//     Standardowa, Kompaktowa). Gęstość jest OPCJĄ (decyzja Mateusza), pamiętana na urządzeniu.
//  3. Panel jako pasek ikon (≥1024 px). Klik w ikonę wysuwa TĘ grupę nad tabelą (tabela się
//     nie przesuwa); ten sam klik / Esc / klik w tabelę — chowa. 📌 przypina panel obok
//     tabeli ze wszystkimi grupami (dawne zachowanie). Najechanie nic nie wysuwa — tylko
//     podpowiedź (auto-wysuwanie łapało kursor jadący do tabeli).
//  4. Telefon: „Aa” wysuwa formatowanie od dołu (tryb Edycja); status pod tabelą.

const SWB_PIN_KEY = "swb-panel-pinned-v1";
const SWB_DENSITY_KEY = "swb-row-density-v1";

var frameCompact = (() => {
  const root = document.documentElement;
  const readLS = (k) => { try { return localStorage.getItem(k); } catch { return null; } };
  const writeLS = (k, v) => { try { localStorage.setItem(k, v); } catch { /* prywatne okno */ } };

  // ── menu paska (Widok ▾, Więcej ▾, powiększenie) ───────────────────────────
  const menus = [
    { btn: document.getElementById("viewMenuBtn"), el: document.getElementById("viewMenu"), align: "left" },
    { btn: document.getElementById("moreMenuBtn"), el: document.getElementById("moreMenu"), align: "left" },
    { btn: document.getElementById("zoomNowBtn"), el: document.getElementById("zoomMenu"), align: "center", build: buildZoomMenu },
  ].filter((m) => m.btn && m.el);
  let openMenu = null;

  function menuItems(m) {
    return Array.from(m.el.querySelectorAll("button:not(.hidden):not([hidden]), select")).filter((b) => !b.disabled && b.offsetParent !== null);
  }
  function placeMenu(m) {
    const r = m.btn.getBoundingClientRect();
    const vw = document.documentElement.clientWidth;
    const w = m.el.offsetWidth || 240;
    let left = m.align === "center" ? r.left + r.width / 2 - w / 2 : r.left;
    left = Math.max(8, Math.min(vw - w - 8, left));
    m.el.style.left = `${Math.round(left)}px`;
    m.el.style.top = `${Math.round(r.bottom + 6)}px`;
    m.el.style.maxHeight = `${Math.max(160, window.innerHeight - r.bottom - 20)}px`;
  }
  function closeMenu(returnFocus) {
    const m = openMenu;
    if (!m) return;
    m.el.hidden = true;
    m.btn.setAttribute("aria-expanded", "false");
    openMenu = null;
    if (returnFocus) m.btn.focus();
  }
  function showMenu(m, focusFirst) {
    closeMenu(false);
    if (m.build) m.build(m.el);
    m.el.hidden = false;
    m.btn.setAttribute("aria-expanded", "true");
    placeMenu(m);
    openMenu = m;
    if (focusFirst) { const f = menuItems(m)[0]; if (f) f.focus(); }
  }
  menus.forEach((m) => {
    if (m.el.parentElement !== document.body) document.body.appendChild(m.el);
    m.btn.addEventListener("click", (e) => {
      e.stopPropagation();
      if (openMenu === m) closeMenu(false); else showMenu(m, e.detail === 0);
    });
    // Klik w pozycję menu robi swoje (handler przycisku) i zamyka menu. Wyjątek: dymek
    // powiększenia — tam można klikać kilka razy (−, +, gęstość) i obserwować tabelę.
    m.el.addEventListener("click", (e) => {
      const b = e.target.closest("button");
      if (!b || m.el.id === "zoomMenu") return;
      setTimeout(() => closeMenu(false), 0);
    });
    m.el.addEventListener("keydown", (e) => {
      if (e.key !== "ArrowDown" && e.key !== "ArrowUp") return;
      const items = menuItems(m);
      const i = items.indexOf(document.activeElement);
      const next = items[(i + (e.key === "ArrowDown" ? 1 : -1) + items.length) % items.length];
      if (next) { e.preventDefault(); next.focus(); }
    });
  });
  document.addEventListener("pointerdown", (e) => {
    if (!openMenu) return;
    if (openMenu.el.contains(e.target) || openMenu.btn.contains(e.target)) return;
    closeMenu(false);
  }, true);
  document.addEventListener("keydown", (e) => {
    if (e.key === "Escape" && openMenu) { e.preventDefault(); e.stopPropagation(); closeMenu(true); }
  }, true);
  window.addEventListener("resize", () => { if (openMenu) placeMenu(openMenu); }, { passive: true });

  // ── powiększenie + gęstość ─────────────────────────────────────────────────
  const zoomSel = typeof zoomLevelEl !== "undefined" ? zoomLevelEl : document.getElementById("zoomLevel");
  const zoomText = document.getElementById("zoomNowText");
  const densitySel = document.getElementById("rowDensity");
  const DENSITIES = ["", "standard", "compact"];

  function zoomOptions() { return zoomSel ? Array.from(zoomSel.options).map((o) => o.value) : []; }
  function setZoom(value) {
    if (!zoomSel || zoomSel.value === value) return;
    zoomSel.value = value;
    zoomSel.dispatchEvent(new Event("change", { bubbles: true }));
    syncZoomText();
  }
  function stepZoom(dir) {
    const opts = zoomOptions();
    const i = opts.indexOf(zoomSel.value);
    const next = opts[Math.max(0, Math.min(opts.length - 1, (i < 0 ? opts.indexOf("1") : i) + dir))];
    if (next) setZoom(next);
  }
  function syncZoomText() {
    if (zoomText && zoomSel) zoomText.textContent = `${Math.round(parseFloat(zoomSel.value || "1") * 100)}%`;
  }
  const zoomOut = document.getElementById("zoomOutBtn");
  const zoomIn = document.getElementById("zoomInBtn");
  if (zoomOut) zoomOut.addEventListener("click", () => stepZoom(-1));
  if (zoomIn) zoomIn.addEventListener("click", () => stepZoom(1));
  if (zoomSel) zoomSel.addEventListener("change", syncZoomText);

  function setDensity(value, { save = true } = {}) {
    const v = DENSITIES.includes(value) ? value : "";
    if (v) root.dataset.density = v; else delete root.dataset.density;
    if (densitySel && densitySel.value !== v) densitySel.value = v;
    if (save) writeLS(SWB_DENSITY_KEY, v);
    if (typeof syncFrozenHeaderMetrics === "function") syncFrozenHeaderMetrics();
    if (typeof syncTableViewportHeight === "function") syncTableViewportHeight();
    if (save && typeof renderActiveTable === "function" && typeof currentHeaders !== "undefined" && currentHeaders.length) renderActiveTable();
    if (openMenu && openMenu.el.id === "zoomMenu") openMenu.build(openMenu.el);
  }
  if (densitySel) densitySel.addEventListener("change", () => setDensity(densitySel.value));

  function buildZoomMenu(el) {
    el.setAttribute("aria-label", t("zoomMenuAria"));
    const head = document.createElement("div");
    head.className = "app-menu-cap";
    head.textContent = t("zoomMenuZoom");
    const presets = document.createElement("div");
    presets.className = "zoom-presets";
    ["0.75", "0.9", "1", "1.25", "1.5"].forEach((v) => {
      const b = document.createElement("button");
      b.type = "button";
      b.className = "zoom-preset";
      b.textContent = `${Math.round(parseFloat(v) * 100)}%`;
      b.setAttribute("aria-pressed", String(zoomSel && zoomSel.value === v));
      b.addEventListener("click", () => { setZoom(v); buildZoomMenu(el); });
      presets.appendChild(b);
    });
    const cap2 = document.createElement("div");
    cap2.className = "app-menu-cap";
    cap2.textContent = t("densityLabel");
    const cur = root.dataset.density || "";
    const opts = DENSITIES.map((d) => {
      const b = document.createElement("button");
      b.type = "button";
      b.className = "app-menu-item density-opt";
      b.setAttribute("role", "menuitemradio");
      b.setAttribute("aria-checked", String(d === cur));
      const name = document.createElement("span");
      name.className = "app-menu-item-text";
      name.textContent = t(`density_${d || "comfortable"}`);
      const lines = document.createElement("span");
      lines.className = `density-glyph density-glyph-${d || "comfortable"}`;
      lines.setAttribute("aria-hidden", "true");
      b.append(lines, name);
      b.addEventListener("click", () => setDensity(d));
      return b;
    });
    el.replaceChildren(head, presets, cap2, ...opts);
  }

  // ── pasek ikon panelu (≥1024 px) ───────────────────────────────────────────
  const rail = document.getElementById("sidebarRail");
  const railBtns = rail ? Array.from(rail.querySelectorAll(".rail-btn[data-group]")) : [];
  const railCount = document.getElementById("railFilterCount");
  const pinBtn = document.getElementById("sidebarPinBtn");
  const closeBtn = document.getElementById("sidebarCloseBtn");
  const sidebar = document.querySelector(".sidebar");
  const scrollBox = document.getElementById("sidebarScroll");
  const finder = document.getElementById("sidebarFinder");
  let pinned = readLS(SWB_PIN_KEY) === "1";
  let currentGroup = "";

  const railMode = () => typeof appFrame !== "undefined" && appFrame.isDocked();
  function syncRailClasses() {
    const on = railMode();
    root.classList.toggle("rail-on", on);
    root.classList.toggle("panel-pinned", on && pinned);
    if (pinBtn) {
      pinBtn.setAttribute("aria-pressed", String(pinned));
      pinBtn.setAttribute("aria-label", t(pinned ? "panelUnpin" : "panelPin"));
    }
    railBtns.forEach((b) => {
      const active = on && typeof isSidebarOpen === "function" && isSidebarOpen() && (pinned ? false : b.dataset.group === currentGroup);
      b.setAttribute("aria-pressed", String(active));
    });
  }
  function focusGroup(id) {
    currentGroup = id || "";
    const flyout = railMode() && !pinned && !!id;
    root.classList.toggle("rail-focus", flyout);
    document.querySelectorAll(".sidebar .sidebar-group").forEach((g) => g.classList.toggle("rail-current", g.id === id));
  }
  function openGroup(id) {
    if (finder && finder.value) { finder.value = ""; if (typeof appFrame !== "undefined") appFrame.runFinder(); }
    focusGroup(id);
    if (typeof setSidebarOpen === "function") setSidebarOpen(true);
    const g = document.getElementById(id);
    if (g) {
      // grupa bez rozwiniętej sekcji — rozwiń pierwszą (inaczej po kliku widać same tytuły)
      if (!g.querySelector("details.panel[open]")) { const d = g.querySelector("details.panel"); if (d) d.open = true; }
      if (pinned) g.scrollIntoView({ block: "start", behavior: "smooth" });
      else if (scrollBox) scrollBox.scrollTop = 0;
    }
    syncRailClasses();
  }
  function closeFlyout({ returnFocus = false } = {}) {
    if (typeof setSidebarOpen === "function") setSidebarOpen(false);
    focusGroup("");
    syncRailClasses();
    if (returnFocus) {
      const b = railBtns.find((x) => x.dataset.group === currentGroup) || railBtns[0];
      if (b) b.focus();
    }
  }
  railBtns.forEach((b) => b.addEventListener("click", () => {
    const id = b.dataset.group;
    const open = typeof isSidebarOpen === "function" && isSidebarOpen();
    if (!pinned && open && currentGroup === id) { closeFlyout(); return; }
    openGroup(id);
  }));
  function setPinned(next) {
    pinned = !!next;
    writeLS(SWB_PIN_KEY, pinned ? "1" : "0");
    if (pinned) {
      focusGroup("");
      if (typeof setSidebarOpen === "function") setSidebarOpen(true);
      try { localStorage.setItem("swb-panel-docked-open-v1", "1"); } catch { /* */ }
    } else {
      closeFlyout();
    }
    syncRailClasses();
    // szerokość tabeli zmienia się bez zmiany okna — przelicz metryki po przejściu
    setTimeout(() => {
      if (typeof syncHorizontalScrollbar === "function") syncHorizontalScrollbar();
      if (typeof syncFrozenHeaderMetrics === "function") syncFrozenHeaderMetrics();
      if (typeof swbVirt !== "undefined" && swbVirt.isActive()) swbVirt.update(true);
    }, 340);
  }
  if (pinBtn) pinBtn.addEventListener("click", () => setPinned(!pinned));
  // ‹ w panelu: przy pasku ikon = schowaj (i odepnij — przypięty panel bez ‹ nie dałby się zwinąć)
  if (closeBtn) closeBtn.addEventListener("click", () => {
    if (!railMode()) return;
    if (pinned) { pinned = false; writeLS(SWB_PIN_KEY, "0"); }
    focusGroup("");
    syncRailClasses();
  }, true);
  // Szukanie ustawienia przeszukuje WSZYSTKIE grupy — zdejmij zawężenie do jednej.
  if (finder) finder.addEventListener("input", () => { if (finder.value) { root.classList.remove("rail-focus"); } else if (currentGroup && !pinned) focusGroup(currentGroup); });

  // Wysunięty (nieprzypięty) panel chowa Esc i klik poza nim — jak menu.
  const flyoutOpen = () => railMode() && !pinned && typeof isSidebarOpen === "function" && isSidebarOpen();
  document.addEventListener("keydown", (e) => {
    if (e.key !== "Escape" || !flyoutOpen()) return;
    if (document.querySelector(".modal:not(.hidden)") || openMenu) return;
    if (finder && document.activeElement === finder && finder.value) return; // najpierw czyści szukajkę
    e.preventDefault();
    e.stopPropagation();
    closeFlyout({ returnFocus: !!(sidebar && sidebar.contains(document.activeElement)) });
  }, true);
  document.addEventListener("pointerdown", (e) => {
    if (!flyoutOpen()) return;
    const t0 = e.target;
    if (!t0 || !t0.closest) return;
    if (t0.closest(".sidebar, .sidebar-rail, .app-menu, .modal, .cell-menu, .fx-suggest, .toast-container, .qs-live, .sheet-picker-overlay, #cellTooltip")) return;
    closeFlyout();
  }, true);

  // licznik filtrów na ikonie (ten sam co przy 🔧)
  function syncRailCount() {
    const src = document.getElementById("panelToggleCount");
    if (!railCount || !src) return;
    railCount.hidden = src.hidden;
    railCount.textContent = src.textContent;
  }
  const countSrc = document.getElementById("panelToggleCount");
  if (countSrc && typeof MutationObserver === "function") {
    new MutationObserver(syncRailCount).observe(countSrc, { attributes: true, childList: true, characterData: true, subtree: true });
  }
  // stan panelu zmienia się też z innych miejsc (🔧 na węższym ekranie, start aplikacji)
  new MutationObserver(syncRailClasses).observe(root, { attributes: true, attributeFilter: ["class"] });
  const dockMq = typeof matchMedia === "function" ? matchMedia("(min-width: 1024px)") : null;
  if (dockMq && typeof dockMq.addEventListener === "function") dockMq.addEventListener("change", () => {
    if (railMode() && !pinned && typeof isSidebarOpen === "function" && isSidebarOpen() && !currentGroup) closeFlyout();
    syncRailClasses();
  });

  // ── telefon: „Aa” = formatowanie od dołu ───────────────────────────────────
  const sheetBtn = document.getElementById("fmtSheetBtn");
  const ribbon = document.getElementById("formatRibbon");
  const phoneMq = typeof matchMedia === "function" ? matchMedia("(max-width: 768px)") : null;
  let sheetDone = null;
  const ribbonHome = ribbon ? ribbon.parentElement : null; // rząd 2 (#tableActions)
  const sheetHost = document.createElement("div");
  sheetHost.className = "fmt-sheet-host";
  document.body.appendChild(sheetHost);
  // Na telefonie rząd 2 bywa zwinięty (▾) — fixed w ukrytym rodzicu by się nie narysował,
  // więc na czas arkusza pasek formatu przeprowadza się do <body> i wraca na miejsce.
  function setFmtSheet(open) {
    root.classList.toggle("fmt-sheet-open", !!open);
    if (sheetBtn) sheetBtn.setAttribute("aria-expanded", String(!!open));
    if (!ribbon) return;
    if (open && !sheetDone) {
      sheetDone = document.createElement("button");
      sheetDone.type = "button";
      sheetDone.className = "btn fmt-sheet-done";
      sheetDone.addEventListener("click", () => setFmtSheet(false));
      ribbon.prepend(sheetDone);
    }
    if (sheetDone) sheetDone.textContent = t("fmtSheetDone");
    if (open && ribbon.parentElement !== sheetHost) sheetHost.appendChild(ribbon);
    if (!open && ribbonHome && ribbon.parentElement !== ribbonHome) ribbonHome.prepend(ribbon);
  }
  if (sheetBtn) sheetBtn.addEventListener("click", () => setFmtSheet(!root.classList.contains("fmt-sheet-open")));
  function syncSheetBtn() {
    if (!sheetBtn) return;
    const show = !!(phoneMq && phoneMq.matches) && typeof swbIsEditMode === "function" && swbIsEditMode()
      && typeof currentHeaders !== "undefined" && currentHeaders.length > 0;
    sheetBtn.hidden = !show;
    if (!show && root.classList.contains("fmt-sheet-open")) setFmtSheet(false);
  }
  if (phoneMq && typeof phoneMq.addEventListener === "function") phoneMq.addEventListener("change", syncSheetBtn);

  // ── napisy ─────────────────────────────────────────────────────────────────
  let labelsLang = "";
  function applyLabels() {
    if (labelsLang === currentLang) return;
    labelsLang = currentLang;
    const set = (id, key) => { const el = document.getElementById(id); if (el && el.firstChild) el.firstChild.nodeValue = `${t(key)} `; };
    set("viewMenuBtn", "viewMenuLabel");
    set("moreMenuBtn", "moreMenuLabel");
    const vm = document.getElementById("viewMenu"); if (vm) vm.setAttribute("aria-label", t("viewMenuLabel"));
    const mm = document.getElementById("moreMenu"); if (mm) mm.setAttribute("aria-label", t("moreMenuLabel"));
    if (rail) rail.setAttribute("aria-label", t("railAria"));
    const dl = document.getElementById("rowDensityLabel"); if (dl) dl.textContent = t("densityLabel");
    if (densitySel) Array.from(densitySel.options).forEach((o) => { o.textContent = t(`density_${o.value || "comfortable"}`); });
    if (zoomOut) zoomOut.setAttribute("aria-label", t("zoomOutAria"));
    if (zoomIn) zoomIn.setAttribute("aria-label", t("zoomInAria"));
    if (sheetBtn) sheetBtn.setAttribute("aria-label", t("fmtSheetAria"));
    railBtns.forEach((b) => { const s = b.querySelector(".visually-hidden"); if (s) s.textContent = t(`rail_${b.dataset.group}`); });
    syncRailClasses();
  }

  // Rząd 2 bez treści (Podgląd i brak „Ostatnich”) znika — nie zostawia pustego pasa.
  const actions = document.getElementById("tableActions");
  const recent = document.getElementById("recentStrip");
  function syncActionsEmpty() {
    if (!actions) return;
    const ribbonShown = ribbon && !ribbon.hidden && ribbon.parentElement === actions;
    const recentShown = recent && !recent.hidden;
    actions.classList.toggle("is-empty", !ribbonShown && !recentShown);
  }
  if (typeof MutationObserver === "function") {
    [ribbon, recent].forEach((el) => el && new MutationObserver(syncActionsEmpty).observe(el, { attributes: true, attributeFilter: ["hidden"] }));
  }

  function sync() {
    applyLabels();
    syncZoomText();
    syncSheetBtn();
    syncRailCount();
    syncActionsEmpty();
  }

  const statsBar = document.getElementById("cellStatsBar");
  if (statsBar && typeof attachOverflowFade === "function") attachOverflowFade(statsBar);

  // start
  setDensity(readLS(SWB_DENSITY_KEY) || "", { save: false });
  syncRailClasses();
  sync();
  return { sync, setDensity, setPinned, openGroup, closeFlyout, closeMenu, get pinned() { return pinned; } };
})();
