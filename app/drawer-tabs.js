// drawer-tabs.js — panel narzędzi na telefonie i tablecie: zakładki grup (audyt QoL S7, 2026-11-10).
//
// Poniżej 1024 px (bez paska ikon, frame-compact.js) „Narzędzia” wysuwały WSZYSTKIE grupy
// naraz: ~2100 px przewijania, a domyślnie rozwinięty „Widok” sam miał ~1000 px — filtry
// i edycja ginęły gdzieś niżej. Tu, jak przy pasku ikon na komputerze, widać jedną grupę:
//   • zakładki Plik · Filtry · Widok · Edycja · Analizy · Pomoc w stałym pasku nad przewijaną
//     częścią; zapamiętana ostatnia (bez pliku zawsze „Plik”, za pierwszym razem „Filtry”);
//   • stały wiersz „Aktywne filtry: n · Wyczyść filtry”, gdy coś filtruje;
//   • „Znajdź ustawienie…” dalej szuka we wszystkich grupach (zawężenie zdjęte na czas szukania);
//   • gdy kod sam otworzy sekcję z innej grupy (np. „Pomoc” z menu ⋯), zakładka idzie za nią.
// Nic nie znika z panelu — zmienia się tylko to, co widać naraz.

const SWB_DRAWER_GROUP_KEY = "swb-drawer-group-v1";

const swbDrawerTabs = (() => {
  const root = document.documentElement;
  const tabsEl = document.getElementById("drawerTabs");
  const activeEl = document.getElementById("drawerActive");
  const activeText = document.getElementById("drawerActiveText");
  const activeClear = document.getElementById("drawerActiveClear");
  const scrollBox = document.getElementById("sidebarScroll");
  const finder = document.getElementById("sidebarFinder");
  const countSrc = document.getElementById("panelToggleCount");
  let current = "";

  const groups = () => Array.from(document.querySelectorAll(".sidebar-scroll > .sidebar-group"));
  const tabs = () => (tabsEl ? Array.from(tabsEl.querySelectorAll("[data-group]")) : []);
  // tryb szuflady = bez paska ikon (ten ma własne zawężenie: rail-focus)
  const drawerMode = () => !root.classList.contains("rail-on");
  const finding = () => !!(finder && finder.value);
  const filterCount = () => (countSrc && !countSrc.hidden ? Number(countSrc.textContent) || 0 : 0);
  const readLS = (k) => { try { return localStorage.getItem(k); } catch { return null; } };
  const writeLS = (k, v) => { try { localStorage.setItem(k, v); } catch { /* bez pamięci */ } };

  function apply() {
    const on = drawerMode() && !!current && !finding();
    root.classList.toggle("drawer-focus", on);
    groups().forEach((g) => g.classList.toggle("drawer-current", g.id === current));
    tabs().forEach((b) => {
      const sel = b.dataset.group === current;
      b.setAttribute("aria-selected", String(sel));
      b.tabIndex = sel ? 0 : -1;
    });
    syncEdges();
  }

  // Na 390 px sześć zakładek się nie mieści — zanikająca krawędź mówi, że pasek przewija się w bok.
  function syncEdges() {
    if (!tabsEl) return;
    const max = tabsEl.scrollWidth - tabsEl.clientWidth;
    tabsEl.classList.toggle("more-right", max > 1 && tabsEl.scrollLeft < max - 1);
    tabsEl.classList.toggle("more-left", max > 1 && tabsEl.scrollLeft > 1);
  }
  function revealTab() {
    const b = tabs().find((x) => x.dataset.group === current);
    if (!b || !tabsEl) return;
    const l = b.offsetLeft - tabsEl.offsetLeft;
    if (l < tabsEl.scrollLeft) tabsEl.scrollLeft = l - 8;
    else if (l + b.offsetWidth > tabsEl.scrollLeft + tabsEl.clientWidth) tabsEl.scrollLeft = l + b.offsetWidth - tabsEl.clientWidth + 8;
    syncEdges();
  }

  function select(id, { remember = true, scroll = true } = {}) {
    if (!document.getElementById(id)) return;
    current = id;
    if (remember) writeLS(SWB_DRAWER_GROUP_KEY, id);
    const g = document.getElementById(id);
    // grupa bez rozwiniętej sekcji — rozwiń pierwszą (inaczej widać same tytuły), jak przy pasku ikon
    if (!g.querySelector("details.panel[open]")) { const d = g.querySelector("details.panel"); if (d) d.open = true; }
    apply();
    revealTab();
    if (scroll && scrollBox) scrollBox.scrollTop = 0;
  }

  function startGroup() {
    if (typeof workbook === "undefined" || !workbook) return "group-data";
    const saved = readLS(SWB_DRAWER_GROUP_KEY);
    return saved && document.getElementById(saved) ? saved : "group-filters";
  }

  function syncActive() {
    const n = filterCount();
    if (activeEl) activeEl.hidden = n === 0;
    if (activeText) activeText.textContent = t("drawerActiveFilters", { count: n });
    tabs().forEach((b) => {
      const c = b.querySelector(".drawer-tab-count");
      if (!c) return;
      const show = b.dataset.group === "group-filters" && n > 0;
      c.hidden = !show;
      c.textContent = show ? String(n) : "";
    });
  }

  function labels() {
    if (tabsEl) tabsEl.setAttribute("aria-label", t("drawerTabsAria"));
    tabs().forEach((b) => { const s = b.querySelector(".drawer-tab-label"); if (s) s.textContent = t(`rail_${b.dataset.group}`); });
    if (activeClear) activeClear.textContent = t("drawerActiveClear");
    syncActive();
  }

  function init() {
    if (!tabsEl || !scrollBox) return;
    tabs().forEach((b) => b.addEventListener("click", () => {
      if (finding()) { finder.value = ""; if (typeof appFrame !== "undefined" && appFrame.runFinder) appFrame.runFinder(); }
      select(b.dataset.group);
    }));
    // strzałki w pasku zakładek (wzorzec tablist)
    tabsEl.addEventListener("keydown", (e) => {
      if (!["ArrowLeft", "ArrowRight", "Home", "End"].includes(e.key)) return;
      const list = tabs();
      const i = list.findIndex((b) => b.dataset.group === current);
      const next = e.key === "Home" ? 0 : e.key === "End" ? list.length - 1
        : (i + (e.key === "ArrowRight" ? 1 : -1) + list.length) % list.length;
      e.preventDefault();
      select(list[next].dataset.group);
      list[next].focus();
    });
    if (activeClear) activeClear.addEventListener("click", () => {
      const reset = document.getElementById("resetFiltersBtn");
      if (reset) reset.click();
    });

    // Otwarcie panelu w trybie szuflady → grupa startowa. Zamknięcie niczego nie zmienia.
    let wasOpen = root.classList.contains("sidebar-open");
    new MutationObserver(() => {
      const open = root.classList.contains("sidebar-open");
      if (open && !wasOpen && drawerMode()) select(startGroup(), { remember: false });
      else apply(); // np. przejście przez 1024 px
      wasOpen = open;
    }).observe(root, { attributes: true, attributeFilter: ["class"] });

    // Kod otworzył sekcję w innej grupie (Pomoc z menu ⋯, agregacja z zapytania…) — zakładka za nią.
    scrollBox.addEventListener("toggle", (e) => {
      const d = e.target;
      if (!d.open || !d.matches || !d.matches("details.panel")) return;
      if (!root.classList.contains("drawer-focus")) return;
      const g = d.closest(".sidebar-group");
      if (!g || g.id === current) return;
      select(g.id, { scroll: false });
      requestAnimationFrame(() => d.scrollIntoView({ block: "start" }));
    }, true);

    tabsEl.addEventListener("scroll", syncEdges, { passive: true });
    window.addEventListener("resize", syncEdges, { passive: true });

    // szukanie ustawienia: wszystkie grupy; po wyczyszczeniu wraca zakładka
    if (finder) finder.addEventListener("input", apply);

    if (countSrc && typeof MutationObserver === "function") {
      new MutationObserver(syncActive).observe(countSrc, { attributes: true, childList: true, characterData: true, subtree: true });
    }
    labels();
    apply();
  }

  return { init, labels, select, _current: () => current };
})();

document.addEventListener("DOMContentLoaded", () => swbDrawerTabs.init());
