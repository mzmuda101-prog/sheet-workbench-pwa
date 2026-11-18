// rail.js — szyna paneli przy wstążce (etap 3; docs/WSTAZKA.md, plan 2026-11-18-wstazka-etap3.md).
//
// Przy wstążce (body.rb-on) pasek ikon #sidebarRail ma 4 przyciski: Filtry (z licznikiem), Sortowanie,
// Formuły, Sprawdzanie wartości — i każdy przycisk (szyny, karty, ekranu Plik) pokazuje w panelu bocznym
// JEDEN panel: z tytułem, pinezką (≥ 1024 px) i ×. Wysuwanie nad tabelę, przypinanie, Esc i klik obok
// robi dalej frame-compact.js / app-frame.js — tu tylko zawężamy, co widać.
// Bez wstążki (telefon) panel działa jak dotąd: grupy, zakładki, „Znajdź ustawienie…”.

const swbRail = (() => {
  const root = document.documentElement;
  const rail = document.getElementById("sidebarRail");
  const sidebar = document.querySelector(".sidebar");
  const titleEl = document.getElementById("rbPaneTitle");
  if (!rail || !sidebar || !titleEl) return null;
  const tr = (k) => (typeof t === "function" ? t(k) : k);
  const RAIL = {
    filters: { panels: ["panel-text-filters", "panel-date-filter", "panel-smart-filters"], title: "rbRail_filters" },
    sort: { panels: ["panel-sort-workbench"], title: "rbRail_sort" },
    formulas: { panels: ["panel-formula-workbench"], title: "rbRail_formulas" },
    check: { panels: ["panel-validation"], title: "rbRail_check" },
  };
  const btns = Array.from(rail.querySelectorAll(".rb-rail-btn[data-rb-rail]"));
  let current = null; // { key, panels, title }
  let releaseTimer = 0;

  const ribbonOn = () => document.body.classList.contains("rb-on");
  const isOpen = () => root.classList.contains("sidebar-open");
  const docked = () => typeof appFrame !== "undefined" && appFrame.isDocked && appFrame.isDocked();
  const panelTitle = (id) => document.getElementById(id)?.querySelector("summary .panel-title")?.textContent.trim() || "";

  function apply() {
    const on = !!current && ribbonOn();
    root.classList.toggle("rb-pane", on);
    root.classList.toggle("rb-pane-single", on && current.panels.length === 1);
    const set = new Set(on ? current.panels : []);
    sidebar.querySelectorAll("details.panel").forEach((d) => d.classList.toggle("rb-pane-on", set.has(d.id)));
    sidebar.querySelectorAll(".sidebar-group").forEach((g) => g.classList.toggle("rb-pane-group", !!g.querySelector("details.panel.rb-pane-on")));
    titleEl.hidden = !on;
    titleEl.textContent = on ? (current.title ? tr(current.title) : panelTitle(current.panels[0])) : "";
    btns.forEach((b) => b.setAttribute("aria-pressed", String(on && isOpen() && current.key === b.dataset.rbRail)));
  }

  // Jeden panel (albo komplet z szyny) — otwiera panel boczny tą samą drogą co pasek ikon.
  function show(panels, { key = "", title = "" } = {}) {
    clearTimeout(releaseTimer);
    current = { key, panels, title };
    panels.forEach((id) => { const d = document.getElementById(id); if (d) d.open = true; });
    const group = document.getElementById(panels[0])?.closest(".sidebar-group")?.id;
    if (docked() && typeof frameCompact !== "undefined" && group) frameCompact.openGroup(group);
    else if (typeof setSidebarOpen === "function") setSidebarOpen(true);
    apply();
    const scroll = document.getElementById("sidebarScroll");
    if (scroll) scroll.scrollTop = 0;
  }

  function close() {
    if (docked() && typeof frameCompact !== "undefined") frameCompact.closeFlyout();
    else if (typeof setSidebarOpen === "function") setSidebarOpen(false);
  }

  btns.forEach((b) => b.addEventListener("click", () => {
    const k = b.dataset.rbRail;
    if (isOpen() && current?.key === k && !root.classList.contains("panel-pinned")) { close(); return; }
    show(RAIL[k].panels, { key: k, title: RAIL[k].title });
  }));

  // Zamknięty panel: tryb jednego panelu zdejmujemy po zjechaniu (inaczej panel zmienia treść w ruchu);
  // fokus wraca na przycisk szyny, jeśli był w panelu.
  let wasOpen = isOpen();
  new MutationObserver(() => {
    const open = isOpen();
    if (open === wasOpen) { apply(); return; }
    wasOpen = open;
    if (open) { apply(); return; }
    const key = current?.key;
    if (key && (sidebar.contains(document.activeElement) || document.activeElement === document.body)) {
      btns.find((b) => b.dataset.rbRail === key)?.focus();
    }
    btns.forEach((b) => b.setAttribute("aria-pressed", "false"));
    clearTimeout(releaseTimer);
    releaseTimer = setTimeout(() => { if (!isOpen()) { current = null; apply(); } }, 320);
  }).observe(root, { attributes: true, attributeFilter: ["class"] });

  // licznik filtrów — ten sam co przy 🔧 / starym pasku ikon
  const countSrc = document.getElementById("panelToggleCount");
  const countEl = document.getElementById("rbRailFilterCount");
  function syncCount() {
    if (!countSrc || !countEl) return;
    countEl.hidden = countSrc.hidden;
    countEl.textContent = countSrc.textContent;
    const b = btns.find((x) => x.dataset.rbRail === "filters");
    if (b) b.setAttribute("aria-label", countSrc.hidden ? tr("rbRail_filters") : `${tr("rbRail_filters")}, ${countSrc.textContent}`);
  }
  if (countSrc) new MutationObserver(syncCount).observe(countSrc, { attributes: true, childList: true, characterData: true, subtree: true });

  function labels() {
    btns.forEach((b) => {
      const txt = tr(RAIL[b.dataset.rbRail].title);
      b.setAttribute("aria-label", txt);
      b.dataset.hintPl = b.dataset.hintEn = txt;
    });
    syncCount();
    apply();
  }
  // 769–1023 px: szyna zaczyna się na wysokości tabeli (≥ 1024 liczy to app-frame przez --dock-top)
  function placeTop() {
    const tp = document.querySelector(".table-panel");
    if (tp && ribbonOn()) root.style.setProperty("--rb-rail-top", `${Math.round(tp.getBoundingClientRect().top)}px`);
  }
  window.addEventListener("resize", () => requestAnimationFrame(placeTop), { passive: true });
  requestAnimationFrame(placeTop);
  // wstążka znika (telefon) → panel wraca do zwykłego układu
  document.addEventListener("swb:ribbon", () => { if (!ribbonOn()) current = null; apply(); requestAnimationFrame(placeTop); });
  labels();

  return { show, close, labels, current: () => current };
})();
