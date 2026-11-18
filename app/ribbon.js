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
  let on = false;
  function place(desk) {
    if (desk === on) return;
    on = desk;
    for (const m of moves) {
      if (desk) { m.el.before(m.mark); m.slot.append(m.el); }
      else if (m.mark.parentNode) { m.mark.replaceWith(m.el); }
    }
    root.hidden = !desk;
    document.body.classList.toggle("rb-on", desk);
    if (desk) requestAnimationFrame(() => strips.forEach((s) => s.sync(true)));
    // wysokość tabeli liczy się od jej górnej krawędzi — po zmianie ramy przelicz
    requestAnimationFrame(() => {
      if (typeof onViewportChange === "function") onViewportChange();
      window.dispatchEvent(new Event("resize"));
    });
    document.dispatchEvent(new CustomEvent("swb:ribbon", { detail: { on: desk } }));
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
    panels.forEach((p) => { p.hidden = p.dataset.rtab !== id; });
    const cmds = strips.find((s) => s.key === "cmds");
    if (cmds) { cmds.sc.scrollLeft = 0; cmds.sync(); }
    tabStrip?.reveal(tab);
    if (focus) tab.focus();
  }
  tabs.forEach((x) => x.addEventListener("click", () => select(x.dataset.rtab)));
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
      .filter((c) => c.offsetParent && !c.matches(".fmt-sep, .rb-sep"));
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
    sheet: () => typeof swbSheetOpsUi !== "undefined" && swbSheetOpsUi.addSheet(),
    // jak „Wyjaśnij formułę” w menu komórki (filter-bar.js), dla komórki z kursorem
    explain: () => {
      const row = typeof focusedCellState !== "undefined" && focusedCellState && typeof getRowByKey === "function" ? getRowByKey(focusedCellState.rowKey) : null;
      const item = row && typeof cellMenuExplainItem === "function" ? cellMenuExplainItem({ row, col: focusedCellState.colIndex0 }) : null;
      if (item) item.run();
      else if (typeof toast === "function") toast(tr("rbNoFormula"), "info");
    },
  };
  root.addEventListener("click", (e) => {
    const b = e.target.closest("[data-rb-act], [data-rb-panel], [data-rb-panel-menu]");
    if (!b || !root.contains(b)) return;
    if (b.dataset.rbAct) ACTS[b.dataset.rbAct]?.();
    else { closeMenus(); openPanel(b.dataset.rbPanel || b.dataset.rbPanelMenu); }
  });

  // ── menu: tryb i Plik ────────────────────────────────────────────────────
  const modeBtn = document.getElementById("rbModeBtn");
  const modeMenu = document.getElementById("rbModeMenu");
  const fileBtn = document.getElementById("rbFileBtn");
  const fileMenu = document.getElementById("rbFileMenu");
  const editBtn = document.getElementById("rbEditBtn");
  const nameEl = document.getElementById("rbName");
  const modeSwitch = document.getElementById("sheetModeSwitch");
  const MENUS = [[modeBtn, modeMenu, "right"], [fileBtn, fileMenu, "left"]].filter(([x, y]) => x && y);
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
    (menu.querySelector('[aria-checked="true"]') || menuItems(menu)[0])?.focus();
  }
  for (const [btn, menu, side] of MENUS) {
    btn.addEventListener("click", () => (menu.hidden ? openMenu(btn, menu, side) : closeMenus()));
    menu.addEventListener("click", (e) => { if (e.target.closest("button")) setTimeout(() => closeMenus(false), 0); });
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
  editBtn?.addEventListener("click", () => setMode("edit"));
  function syncMode() {
    const m = isView() ? "view" : "edit";
    document.body.classList.toggle("rb-reading", on && m === "view");
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
    fileMenu?.setAttribute("aria-label", tr("rbFile"));
    const ed = editBtn?.querySelector("span");
    if (ed) ed.textContent = tr("rbEdit");
    syncMode();
    const file = document.getElementById("rbFileBtn");
    if (file) file.textContent = tr("rbFile");
    tabs.forEach((x) => { x.textContent = tr(`rbTab_${x.dataset.rtab}`); });
    root.querySelector('[role="tablist"]')?.setAttribute("aria-label", tr("rbTabsAria"));
    root.querySelectorAll(".rb-nudge-l").forEach((b) => b.setAttribute("aria-label", tr("rbScrollLeft")));
    root.querySelectorAll(".rb-nudge-r").forEach((b) => b.setAttribute("aria-label", tr("rbScrollRight")));
    strips.forEach((s) => s.sync());
  }

  // ── panel boczny: otwórz sekcję (do etapu 3 — potem szyna z wysuwanym panelem) ──
  function openPanel(detailsId) {
    const d = document.getElementById(detailsId);
    if (!d) return;
    if (typeof setSidebarOpen === "function") setSidebarOpen(true);
    d.open = true;
    requestAnimationFrame(() => d.scrollIntoView({ block: "nearest" }));
  }

  const onMq = () => { closeMenus(); place(deskMq.matches); syncMode(); };
  if (deskMq.addEventListener) deskMq.addEventListener("change", onMq); else deskMq.addListener(onMq);
  place(deskMq.matches);
  labels();
  document.addEventListener("keydown", (e) => { if (e.key === "Escape" && MENUS.some(([, m]) => !m.hidden)) { e.preventDefault(); closeMenus(true); } });

  return { isOn: () => on, select, current: () => current, openPanel, labels };
})();
