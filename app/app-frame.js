// app-frame.js — „rama ekranu” wokół tabeli (paczka B, makieta zaakceptowana 2026-09-27).
//
// Co tu jest i dlaczego:
// - Nagłówek pokazuje, CO jest otwarte: nazwę pliku i bieżący arkusz (na telefonie
//   przycisk z listą arkuszy, na desktopie zakładki przeniesione z dołu tabeli).
// - „Zapisz” na wierzchu: szary bez zmian, akcent + liczba zmian po edycji. Klik robi
//   dokładnie to, co „Zapisz” w panelu (albo „Zapisz jako…”, gdy nadpisanie oryginału
//   w tej przeglądarce jest niedostępne).
// - Menu ⋯: język, motyw, odświeżenie, inne aplikacje, stan PWA — rzadko używane rzeczy,
//   które dotąd zajmowały 7 miejsc w nagłówku.
// - „Narzędzia” (dawny „Filtry”) przy polu szukania, z liczbą działających filtrów.
// - Szukanie zawsze widoczne (dawny przełącznik „Szybkie szukanie” zniknął).
// - Pasek „Ostatnie” z historii zapytań (qs-history.js): klik = szukaj ponownie.
// - Bezpiecznik przepełnienia: rzędy o zmiennej długości (zakładki, pasek przycisków,
//   „Ostatnie”) nigdy nie wychodzą poza ekran — przewijają się w bok, a krawędź miękko
//   wygasa TYLKO wtedy, gdy faktycznie jest co przewinąć (zasada Mateusza).
// - Opcja „Pokaż uchwyt panelu” (Widok) — wąski uchwyt przy krawędzi zostaje domyślnie.

const PANEL_HANDLE_KEY = "swb-panel-handle-v1";
const RECENT_STRIP_MAX = 12;

// ── bezpiecznik przepełnienia ─────────────────────────────────────────────────
function attachOverflowFade(el) {
  if (!el || el._swbFade) return;
  el._swbFade = true;
  el.classList.add("overflow-fade");
  const sync = () => {
    const max = el.scrollWidth - el.clientWidth;
    el.classList.toggle("more-l", max > 1 && el.scrollLeft > 2);
    el.classList.toggle("more-r", max > 1 && el.scrollLeft < max - 2);
  };
  // Pionowe kółko myszy przewija rząd w bok (Shift+kółko działa natywnie).
  el.addEventListener("wheel", (e) => {
    if (Math.abs(e.deltaY) <= Math.abs(e.deltaX)) return;
    if (el.scrollWidth <= el.clientWidth + 1) return;
    el.scrollLeft += e.deltaY;
    e.preventDefault();
  }, { passive: false });
  el.addEventListener("scroll", sync, { passive: true });
  if (typeof ResizeObserver === "function") new ResizeObserver(sync).observe(el);
  if (typeof MutationObserver === "function") new MutationObserver(sync).observe(el, { childList: true, subtree: true, characterData: true });
  el._swbFadeSync = sync;
  sync();
}

const appFrame = (() => {
  const heroEl = document.querySelector(".hero");
  const heroTitleEl = document.getElementById("heroTitle");
  const heroSheetBtn = document.getElementById("heroSheetBtn");
  const heroSheetName = document.getElementById("heroSheetName");
  const heroTabsSlot = document.getElementById("heroTabsSlot");
  const heroSaveBtn = document.getElementById("heroSaveBtn");
  const heroSaveCount = document.getElementById("heroSaveCount");
  const menuBtn = document.getElementById("appMenuBtn");
  const menuEl = document.getElementById("appMenu");
  const buildEl = document.getElementById("appMenuBuild");
  const panelCountEl = document.getElementById("panelToggleCount");
  const recentStrip = document.getElementById("recentStrip");
  const recentSep = document.querySelector(".recent-sep");
  const handleOpt = document.getElementById("showPanelHandle");
  const tableActions = document.getElementById("tableActions");

  // ── nagłówek: plik + arkusz ────────────────────────────────────────────────
  const tablePanel = document.querySelector(".table-panel");
  function syncFile() {
    const hasFile = !!(typeof workbook !== "undefined" && workbook && currentFileName);
    if (heroEl) heroEl.classList.toggle("has-file", hasFile);
    // bez pliku pole szukania nie ma czego szukać — chowamy je (zostaje „Narzędzia”)
    if (tablePanel) tablePanel.classList.toggle("no-data", !hasFile);
    if (heroTitleEl) {
      heroTitleEl.textContent = hasFile ? currentFileName : "Sheet Workbench";
      if (hasFile) heroTitleEl.setAttribute("title", currentFileName); else heroTitleEl.removeAttribute("title");
    }
    const sheet = typeof currentSheetName === "string" ? currentSheetName : "";
    const multi = hasFile && workbook.SheetNames && workbook.SheetNames.length > 1;
    if (heroSheetBtn) {
      heroSheetBtn.hidden = !(hasFile && sheet);
      heroSheetBtn.disabled = !multi;
      heroSheetBtn.classList.toggle("is-multi", !!multi);
      if (heroSheetName) heroSheetName.textContent = sheet;
      heroSheetBtn.setAttribute("aria-label", multi ? `${t("heroSheetSwitch")}: ${sheet}` : sheet);
    }
    syncSave();
  }

  // Zakładki arkuszy z dołu tabeli → do nagłówka (desktop; na telefonie slot jest
  // ukryty CSS-em, a arkusz zmienia przycisk z nazwą). Element zostaje ten sam, więc
  // cała obsługa (klik, klawiatura, aria-selected) działa bez zmian.
  function mountSheetTabs() {
    if (typeof sheetTabsEl === "undefined" || !sheetTabsEl || !heroTabsSlot) return;
    if (sheetTabsEl.parentElement !== heroTabsSlot) heroTabsSlot.appendChild(sheetTabsEl);
    attachOverflowFade(sheetTabsEl);
  }

  if (heroSheetBtn) {
    heroSheetBtn.addEventListener("click", () => {
      if (heroSheetBtn.disabled) return;
      if (typeof openSheetPicker === "function") openSheetPicker();
    });
  }

  // ── Zapisz ─────────────────────────────────────────────────────────────────
  function pendingCount() {
    if (typeof pendingEdits === "undefined" || !pendingEdits) return 0;
    let n = 0;
    Object.values(pendingEdits).forEach((m) => { if (m) n += Object.keys(m).length; });
    return n;
  }
  function syncSave() {
    if (!heroSaveBtn) return;
    const hasFile = !!(typeof workbook !== "undefined" && workbook);
    heroSaveBtn.hidden = !hasFile;
    const dirty = hasFile && typeof hasUnsavedChanges !== "undefined" && !!hasUnsavedChanges;
    const n = dirty ? pendingCount() : 0;
    heroSaveBtn.classList.toggle("is-dirty", dirty);
    if (heroSaveCount) {
      heroSaveCount.hidden = !(dirty && n > 0);
      heroSaveCount.textContent = n > 99 ? "99+" : String(n);
    }
    heroSaveBtn.setAttribute("aria-label", dirty && n ? t("heroSaveDirty", { changes: formatUnsavedCount(n) }) : t("heroSave"));
  }
  if (heroSaveBtn) {
    heroSaveBtn.addEventListener("click", () => {
      const save = typeof saveBtn !== "undefined" ? saveBtn : null;
      const saveAs = typeof saveAsBtn !== "undefined" ? saveAsBtn : null;
      if (save && !save.disabled) save.click();
      else if (saveAs && !saveAs.disabled) saveAs.click();
    });
  }

  // ── menu ⋯ ─────────────────────────────────────────────────────────────────
  function menuItems() {
    return Array.from(menuEl.querySelectorAll("button, a[href]")).filter((el) => !el.disabled && el.offsetParent !== null);
  }
  // Menu mieszka w <body>, nie w nagłówku: nagłówek ma backdrop-filter (to tworzy
  // własny układ odniesienia dla position:fixed) i na telefonie overflow:hidden
  // (zwijanie) — menu w środku byłoby przycięte. Pozycję liczymy z przycisku.
  if (menuEl && menuEl.parentElement !== document.body) document.body.appendChild(menuEl);
  function placeMenu() {
    const r = menuBtn.getBoundingClientRect();
    const vw = document.documentElement.clientWidth;
    const width = Math.min(260, vw - 16);
    menuEl.style.width = `${width}px`;
    menuEl.style.top = `${Math.round(r.bottom + 8)}px`;
    menuEl.style.left = `${Math.round(Math.max(8, Math.min(vw - width - 8, r.right - width)))}px`;
    menuEl.style.maxHeight = `${Math.max(160, window.innerHeight - r.bottom - 20)}px`;
  }
  function setMenuOpen(open, opts = {}) {
    if (!menuEl || !menuBtn) return;
    menuEl.hidden = !open;
    menuBtn.setAttribute("aria-expanded", String(open));
    if (open) {
      placeMenu();
      // wskaźnik PL/EN liczy się z offsetLeft przycisków — w ukrytym menu było 0
      if (typeof updateLangSwitchIndicator === "function") updateLangSwitchIndicator();
      if (buildEl && typeof APP_BUILD_VERSION !== "undefined") buildEl.textContent = APP_BUILD_VERSION;
      if (opts.focusFirst) { const first = menuItems()[0]; if (first) first.focus(); }
    } else if (opts.returnFocus) {
      menuBtn.focus();
    }
  }
  if (menuBtn && menuEl) {
    menuBtn.addEventListener("click", (e) => {
      e.stopPropagation();
      setMenuOpen(menuEl.hidden, { focusFirst: e.detail === 0 }); // klawiatura → fokus w menu
    });
    document.addEventListener("pointerdown", (e) => {
      if (menuEl.hidden) return;
      if (menuEl.contains(e.target) || menuBtn.contains(e.target)) return;
      setMenuOpen(false);
    });
    // Esc zamyka menu ZAWSZE, gdy jest otwarte — także gdy fokus został na przycisku ⋯
    // (otwarcie myszką). Faza capture: wcześniej niż globalny Esc apki (odznaczanie itd.).
    document.addEventListener("keydown", (e) => {
      if (e.key !== "Escape" || menuEl.hidden) return;
      e.preventDefault();
      e.stopPropagation();
      setMenuOpen(false, { returnFocus: menuEl.contains(document.activeElement) || document.activeElement === menuBtn });
    }, true);
    menuEl.addEventListener("keydown", (e) => {
      if (e.key !== "ArrowDown" && e.key !== "ArrowUp") return;
      const items = menuItems();
      const i = items.indexOf(document.activeElement);
      const next = items[(i + (e.key === "ArrowDown" ? 1 : -1) + items.length) % items.length];
      if (next) { e.preventDefault(); next.focus(); }
    });
    window.addEventListener("resize", () => { if (!menuEl.hidden) placeMenu(); }, { passive: true });
    // Zmiana języka w menu → odśwież napisy tworzone tutaj.
    menuEl.querySelectorAll(".lang-button").forEach((b) => b.addEventListener("click", () => setTimeout(() => { renderRecent(); syncFile(); }, 0)));
    // Linki do innych aplikacji otwierają nową kartę — menu niech się zamknie.
    menuEl.querySelectorAll("a[href]").forEach((a) => a.addEventListener("click", () => setMenuOpen(false)));
  }

  // ── „Narzędzia”: liczba działających filtrów ───────────────────────────────
  function syncPanelCount() {
    if (!panelCountEl) return;
    const snap = typeof lastAppliedFilters !== "undefined" ? lastAppliedFilters : null;
    const n = (snap && (snap.filtering || snap.marking) && typeof afCollectChips === "function" && currentHeaders.length)
      ? afCollectChips().length
      : 0;
    panelCountEl.hidden = n === 0;
    panelCountEl.textContent = String(n);
  }

  // Pasek aktywnych filtrów (filter-bar.js) przelicza się po każdej zmianie filtrów —
  // tam też odświeżamy licznik przy „Narzędzia” i pasek „Ostatnie” (zależy od danych).
  if (typeof window.renderActiveFilters === "function") {
    const origRender = window.renderActiveFilters;
    window.renderActiveFilters = function renderActiveFiltersAndCount(...args) {
      const r = origRender.apply(this, args);
      syncPanelCount();
      if (recentStrip && recentStrip.hidden === (typeof currentHeaders !== "undefined" && currentHeaders.length > 0)) renderRecent();
      return r;
    };
  }

  // ── „Ostatnie” zapytania ───────────────────────────────────────────────────
  function renderRecent() {
    if (!recentStrip || typeof qsHistoryLoad !== "function") return;
    const hasData = typeof currentHeaders !== "undefined" && currentHeaders.length > 0;
    const h = qsHistoryLoad();
    const pinnedKeys = new Set(h.pinned.map(qsHistoryKey));
    const list = [
      ...h.pinned.map((e) => ({ e, pinned: true })),
      ...h.recent.filter((e) => !pinnedKeys.has(qsHistoryKey(e))).map((e) => ({ e, pinned: false })),
    ].slice(0, RECENT_STRIP_MAX);
    const show = hasData && list.length > 0;
    recentStrip.hidden = !show;
    if (recentSep) recentSep.hidden = !show;
    if (!show) { recentStrip.replaceChildren(); return; }
    const frag = document.createDocumentFragment();
    const label = document.createElement("span");
    label.className = "recent-label";
    label.textContent = t("recentLabel");
    frag.appendChild(label);
    list.forEach(({ e, pinned }) => {
      const chip = document.createElement("button");
      chip.type = "button";
      chip.className = pinned ? "recent-chip is-pinned" : "recent-chip";
      const glyph = e.neg ? "≠ " : (e.mode && e.mode !== "contains" && typeof QS_FLAG_GLYPH !== "undefined" ? `${QS_FLAG_GLYPH[e.mode] || ""} ` : "");
      chip.textContent = `${glyph}${e.q}`;
      chip.title = e.q;
      chip.addEventListener("click", () => {
        if (typeof qsHistoryApply !== "function") return;
        qsHistoryApply({ inputEl: quickSearchEl, liveEl: document.getElementById("qsLiveResults"), wrapEl: quickSearchWrap }, e);
      });
      frag.appendChild(chip);
    });
    recentStrip.replaceChildren(frag);
    attachOverflowFade(recentStrip);
  }
  // Każda zmiana historii (nowe zapytanie, przypięcie, usunięcie, czyszczenie) przechodzi
  // przez qsHistorySave() w qs-history.js — owijamy ten jeden punkt, żeby pasek był świeży.
  if (typeof window.qsHistorySave === "function") {
    const origSave = window.qsHistorySave;
    window.qsHistorySave = function qsHistorySaveAndRefresh(...args) {
      const r = origSave.apply(this, args);
      queueMicrotask(renderRecent);
      return r;
    };
  }

  // ── uchwyt panelu (opcja) ──────────────────────────────────────────────────
  function applyHandleOpt(show) {
    rootEl.classList.toggle("no-panel-handle", !show);
  }
  if (handleOpt) {
    let show = true;
    try { show = localStorage.getItem(PANEL_HANDLE_KEY) !== "0"; } catch (_) {}
    handleOpt.checked = show;
    applyHandleOpt(show);
    handleOpt.addEventListener("change", () => {
      applyHandleOpt(handleOpt.checked);
      try { localStorage.setItem(PANEL_HANDLE_KEY, handleOpt.checked ? "1" : "0"); } catch (_) {}
    });
  }

  // ── start ──────────────────────────────────────────────────────────────────
  // Szukanie zawsze widoczne: dawny „tryb szybkiego szukania” jest teraz stanem stałym.
  if (typeof setReadingMode === "function") setReadingMode(true);
  mountSheetTabs();
  attachOverflowFade(tableActions);
  syncFile();
  renderRecent();

  return { syncFile, syncSave, syncPanelCount, renderRecent, setMenuOpen };
})();
