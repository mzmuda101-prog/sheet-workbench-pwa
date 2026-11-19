// @ts-check
// ribbon-classic.js — wstążka Klasyczna (etap 6b; docs/WSTAZKA.md, makieta v2 .r2.classic).
//
// Duże przyciski (ikona nad podpisem), kolumny małych z podpisami i nazwy grup pod spodem. Buduje
// panele .rb-panel.rb-cl obok paneli Uproszczonej; ribbon.js pokazuje jedne albo drugie.
// Wzór: Documents Workbench (app/ribbon-classic.js tam). Bez kopii logiki:
//   - przyciski z data-rb-act / data-rb-panel to KLONY przycisków Uproszczonej (te same atrybuty,
//     więc tę samą akcję wykonuje obsługa kliknięć w ribbon.js; ikona i podpis z oryginału),
//   - kontrolki z id (blok formatu, Raport, Widok ▾…) to TE SAME elementy: w Klasycznej przenoszą
//     się do gniazd [data-rbc] (ribbon.js applyLayout). Blok formatu zostaje w całości — jego
//     obsługa w sheet-edit-ui.js słucha na całym bloku — a w grupy układa go sam CSS.
// Tylko polecenia, które aplikacja już ma (decyzja Mateusza 2026-11-19); brakujące z makiety
// są spisane w docs/WSTAZKA.md.

/* global t */

const swbRibbonClassic = (() => {
  const root = document.getElementById("ribbon");
  const host = root?.querySelector('[data-strip="cmds"] .rb-sc');
  if (!root || !host) return null;
  const tr = (/** @type {string} */ k) => (typeof t === "function" ? t(k) : k);

  // ── opis kart ─────────────────────────────────────────────────────────────
  // B = duży przycisk, S = mały z podpisem, L = gniazdo na prawdziwą kontrolkę, ST = kolumna
  const B = (src) => ({ k: "big", src });
  const S = (src) => ({ k: "sm", src });
  const L = (sel) => ({ k: "slot", sel });
  const ST = (...items) => ({ k: "stack", items });
  const G = (key, ...items) => ({ key, items });
  const act = (a) => `[data-rb-act="${a}"]`;
  const pan = (p) => `[data-rb-panel="${p}"]`;
  const SPEC = {
    home: [
      G("rbG_clip", B(act("paste")), ST(S(act("cut")), S(act("copy")))),
      { slot: "#formatRibbon" }, // Czcionka · Wyrównanie · Liczba · Edycja — grupy z CSS
      G("rbG_tools", B(pan("panel-edit-tools"))),
    ],
    ins: [
      G("rbG_data", B(pan("panel-derived-columns"))),
      G("rbG_cell", B(act("note"))),
    ],
    data: [
      G("rbG_sort", L("#resetSortBtn")),
      G("rbG_dataTools", L("#wideLongToggle")),
      G("rbG_protect", B(pan("panel-anonymize"))),
    ],
    fx: [
      G("rbG_audit", B(act("explain"))),
    ],
    an: [
      G("rbG_sheet", B(pan("panel-workbench-analysis")), ST(S(pan("panel-sheet-inspector")), S(pan("panel-kpi-extractor")))),
      G("rbG_summary", B(pan("panel-aggregation-workbench")), ST(S(pan("panel-monthly-summary")))),
      G("rbG_result", L("#reportBtn")),
    ],
    view: [
      G("rbG_view", L("#viewMenuBtn"), B(pan("panel-view"))),
    ],
    ai: [
      G("rbG_ask", L("#aiMenuItem")),
    ],
  };

  // ── budowa ────────────────────────────────────────────────────────────────
  const simple = (sel) => /** @type {HTMLElement|null} */ (root.querySelector(`.rb-panel:not(.rb-cl) ${sel}`));
  /** @type {HTMLElement[]} */
  const slots = [];
  function button(it) {
    const src = simple(it.src);
    if (!src) return null;
    const b = /** @type {HTMLElement} */ (src.cloneNode(true));
    b.removeAttribute("id");
    b.className = `rb-btn ${it.k === "big" ? "rb-big" : "rb-sm"}`;
    // podpis zawsze widoczny (w Uproszczonej część przycisków to same ikony); ribbon.js labels() go tłumaczy
    if (!b.querySelector(":scope > span:not([class])")) {
      const sp = document.createElement("span");
      sp.textContent = src.getAttribute("aria-label") || "";
      b.append(sp);
    }
    return b;
  }
  function slot(sel) {
    const s = document.createElement("span");
    s.className = "rb-slot";
    s.dataset.rbc = sel;
    slots.push(s);
    return s;
  }
  function node(it) {
    if (it.k === "stack") {
      const d = document.createElement("div");
      d.className = "rb-cst";
      d.append(...it.items.map(node).filter(Boolean));
      return d;
    }
    if (it.k === "slot") return slot(it.sel);
    return button(it);
  }
  function group(g) {
    if (g.slot) return slot(g.slot);
    const d = document.createElement("div");
    d.className = "rb-cg";
    d.setAttribute("role", "group");
    const items = document.createElement("div");
    items.className = "rb-gi";
    items.append(...g.items.map(node).filter(Boolean));
    const cap = document.createElement("div");
    cap.className = "rb-gcap";
    cap.dataset.rbLabel = g.key; // ribbon.js labels(): tłumaczenie
    cap.textContent = tr(g.key);
    d.setAttribute("aria-label", tr(g.key));
    d.append(items, cap);
    return d;
  }
  for (const [tab, groups] of Object.entries(SPEC)) {
    const p = document.createElement("div");
    p.className = "rb-panel rb-cl";
    p.setAttribute("role", "tabpanel");
    p.id = `rbPanelCl-${tab}`;
    p.dataset.rtab = tab;
    p.setAttribute("aria-labelledby", `rbTab-${tab}`);
    p.hidden = true;
    p.append(...groups.map(group));
    host.append(p);
  }
  // nazwy grup dla czytnika ekranu nadążają za językiem (podpis widoczny tłumaczy labels())
  function labels() {
    host.querySelectorAll(".rb-cl .rb-cg").forEach((g) => g.setAttribute("aria-label", g.querySelector(".rb-gcap")?.textContent || ""));
  }

  return { slots, labels };
})();
