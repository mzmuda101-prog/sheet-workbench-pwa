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

  // ── teksty (PL/EN) ────────────────────────────────────────────────────────
  function labels() {
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

  const onMq = () => place(deskMq.matches);
  if (deskMq.addEventListener) deskMq.addEventListener("change", onMq); else deskMq.addListener(onMq);
  place(deskMq.matches);
  labels();

  return { isOn: () => on, select, current: () => current, openPanel, labels };
})();
