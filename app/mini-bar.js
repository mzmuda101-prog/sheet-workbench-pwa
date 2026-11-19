// @ts-check
// mini-bar.js — mini pasek nad zaznaczeniem (etap 5 wstążki; docs/WSTAZKA.md, makieta v2 .mini).
//
// Po zaznaczeniu MYSZĄ co najmniej dwóch komórek (przeciągnięcie albo Shift+klik) nad zakresem
// pojawia się mały pasek: pogrubienie, wypełnienie, format liczb · Σ, Filtruj wg wartości ·
// Notatka, Wyjaśnij. To celowy skrót przy zaznaczeniu (decyzja Mateusza 2026-11-19) — przyciski
// klikają TE SAME kontrolki co wstążka i menu komórki, nic nie jest kopią logiki.
// Jedna komórka, klawiatura, dotyk, Podgląd i telefon (bez wstążki) — bez paska.
//
// Miejsce: nad zakresem, a gdy tam brak miejsca (nagłówek kolumn) — pod nim. Przy przewijaniu
// gaśnie i wraca po zatrzymaniu. Znika po kliknięciu obok, po klawiszu (też Esc) i po poleceniu,
// które zmienia widok (Σ, Filtruj, Notatka, Wyjaśnij); po B / wypełnieniu / formacie zostaje.

/* global t, tableEl, tableWrapEl, getSelectionRectangle,
   collectSelectionCombos, applyCellFilter, cellMenuTarget, giveGridDomFocus, swbIsEditMode,
   currentDisplayModel, dragSelEndedAt */

const swbMiniBar = (() => {
  if (typeof tableEl === "undefined" || !tableEl) return null;
  const $ = (/** @type {string} */ s) => /** @type {HTMLElement|null} */ (document.querySelector(s));
  const tr = (/** @type {string} */ k) => (typeof t === "function" ? t(k) : k);
  const GAP = 8;
  const SRC = {
    bold: '#formatRibbon [data-fmt="b"]',
    fill: "#fmtFillBtn",
    num: "#fmtNumFormat",
    sum: "#fmtSumBtn",
    note: '#ribbon [data-rb-act="note"]',
    explain: '#ribbon [data-rb-act="explain"]',
  };
  const FILTER_ICON = '<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M3 4h18l-7 8.5V19l-4 2v-8.5L3 4z"/></svg>';

  const bar = document.createElement("div");
  bar.id = "miniBar";
  bar.className = "mini-bar";
  bar.hidden = true;
  bar.setAttribute("role", "toolbar");

  // ikona (i podpis) — klon z oryginału, bez id
  function cloneInto(/** @type {HTMLElement} */ b, /** @type {string} */ srcSel) {
    const src = $(srcSel);
    if (!src) return;
    for (const n of Array.from(src.childNodes)) {
      // bez podpisu karty (goły <span>); litery B / Σ (.fmt-glyph) i pasek koloru zostają
      if (n.nodeType !== 1 || (n.nodeName === "SPAN" && !(/** @type {HTMLElement} */ (n).className))) continue;
      const c = /** @type {HTMLElement} */ (n.cloneNode(true));
      c.removeAttribute("id");
      b.append(c);
    }
  }
  /** @type {Array<{b: HTMLElement, srcSel: string}>} */
  const mirrors = [];
  function button(/** @type {string} */ key, /** @type {() => void} */ act, { label = false, cls = "", hide = false } = {}) {
    const b = document.createElement("button");
    b.type = "button";
    b.className = `rb-btn mini-btn ${label ? "" : "rb-ic"} ${cls}`.trim();
    b.dataset.mini = key;
    if (SRC[key]) { cloneInto(b, SRC[key]); mirrors.push({ b, srcSel: SRC[key] }); }
    b.addEventListener("mousedown", (e) => e.preventDefault()); // fokus zostaje w tabeli
    b.addEventListener("click", () => { act(); if (hide) hideBar(); else setTimeout(sync, 60); });
    return b;
  }
  const sep = () => { const s = document.createElement("span"); s.className = "rb-sep"; s.setAttribute("aria-hidden", "true"); return s; };
  const click = (/** @type {string} */ key) => () => $(SRC[key])?.click();

  // format liczb: lustro listy ze wstążki, wybór przechodzi do oryginału (jego obsługa robi resztę)
  const num = document.createElement("select");
  num.className = "fmt-select mini-num";
  num.dataset.mini = "num";
  num.addEventListener("change", () => {
    const src = /** @type {HTMLSelectElement|null} */ ($(SRC.num));
    if (!src) return;
    src.value = num.value;
    src.dispatchEvent(new Event("change", { bubbles: true }));
    giveGridDomFocus();
    setTimeout(sync, 60);
  });

  // Filtruj wg wartości = „Pokaż tylko wiersze z tą wartością” z menu komórki, na całym zakresie
  function filterByValue() {
    const rect = getSelectionRectangle();
    if (!rect) return;
    const { headers, combos } = collectSelectionCombos(rect);
    // eslint-disable-next-line no-global-assign
    cellMenuTarget = { multi: true, rect, headers, combos, td: null };
    applyCellFilter(false);
    cellMenuTarget = null;
  }
  const filterBtn = button("filter", filterByValue, { label: true, hide: true });
  filterBtn.innerHTML = FILTER_ICON;
  const filterTxt = document.createElement("span");
  filterBtn.append(filterTxt);
  const explainBtn = button("explain", click("explain"), { label: true, hide: true });
  const explainTxt = document.createElement("span");
  explainBtn.append(explainTxt);

  bar.append(
    button("bold", click("bold")),
    button("fill", click("fill"), { cls: "mini-fill" }),
    num,
    sep(),
    button("sum", click("sum"), { hide: true }),
    filterBtn,
    sep(),
    button("note", click("note"), { hide: true }),
    explainBtn,
  );
  document.body.append(bar);

  // napisy, podpowiedzi i stan — z oryginałów (język, „wciśnięte”, kolor wypełnienia)
  function sync() {
    bar.setAttribute("aria-label", tr("miniBarAria"));
    for (const { b, srcSel } of mirrors) {
      const src = $(srcSel);
      if (!src) continue;
      b.setAttribute("aria-label", src.getAttribute("aria-label") || "");
      for (const k of ["hint", "hintPl", "hintEn"]) if (src.dataset[k] != null) b.dataset[k] = src.dataset[k];
      if (src.hasAttribute("aria-pressed")) b.setAttribute("aria-pressed", src.getAttribute("aria-pressed") || "false");
    }
    const sw = $("#fmtFillSwatch"), dst = /** @type {HTMLElement|null} */ (bar.querySelector(".mini-fill .fmt-swatch"));
    if (sw && dst) dst.style.background = sw.style.background;
    const src = /** @type {HTMLSelectElement|null} */ ($(SRC.num));
    if (src) {
      num.replaceChildren(...Array.from(src.children).map((c) => c.cloneNode(true)));
      num.value = src.value;
      num.setAttribute("aria-label", src.getAttribute("aria-label") || "");
      for (const k of ["hint", "hintPl", "hintEn"]) if (src.dataset[k] != null) num.dataset[k] = src.dataset[k];
    }
    filterTxt.textContent = tr("miniFilter");
    filterBtn.setAttribute("aria-label", tr("miniFilter"));
    filterBtn.dataset.hint = "";
    filterBtn.dataset.hintPl = filterBtn.dataset.hintEn = tr("miniFilterHint");
    explainTxt.textContent = tr("miniExplain");
    explainBtn.setAttribute("aria-label", $(SRC.explain)?.getAttribute("aria-label") || tr("miniExplain"));
  }

  // ── kiedy i gdzie ───────────────────────────────────────────────────────
  let wanted = false;
  const on = () => document.body.classList.contains("rb-on") && swbIsEditMode() && currentDisplayModel?.mode === "wide";
  function selRect() {
    const r = getSelectionRectangle();
    return r && r.rowCount * r.colCount >= 2 ? r : null;
  }
  function hideBar() {
    wanted = false;
    bar.hidden = true;
    bar.classList.remove("is-off");
  }
  function place() {
    const r = wanted && on() ? selRect() : null;
    if (!r) { hideBar(); return; }
    // rogi zakresu, które są w DOM (wirtualizacja może schować część wierszy)
    const tds = [];
    for (const row of tableEl.querySelectorAll("tr[data-row-key]")) {
      if (!r.rowKeys.has(/** @type {HTMLElement} */ (row).dataset.rowKey)) continue;
      for (const c of [r.colMin, r.colMax]) { const td = row.querySelector(`td[data-col-index="${c}"]`); if (td) tds.push(td.getBoundingClientRect()); }
    }
    const V = tableWrapEl.getBoundingClientRect();
    const head = tableEl.querySelector("thead");
    const topLimit = Math.max(V.top, head ? head.getBoundingClientRect().bottom : V.top);
    if (!tds.length) { bar.classList.add("is-off"); return; }
    const top = Math.min(...tds.map((x) => x.top)), bottom = Math.max(...tds.map((x) => x.bottom));
    const left = Math.min(...tds.map((x) => x.left));
    if (bottom < topLimit + 4 || top > V.bottom - 4) { bar.classList.add("is-off"); return; }
    const fresh = bar.hidden;
    if (fresh) { sync(); bar.hidden = false; }
    const w = bar.offsetWidth, h = bar.offsetHeight;
    let y = top - h - GAP;
    let below = false;
    if (y < topLimit) { y = bottom + GAP; below = true; } // nad zakresem brak miejsca — pod nim
    const x = Math.max(V.left + GAP, Math.min(left, V.right - w - GAP));
    bar.style.top = `${Math.round(y)}px`;
    bar.style.left = `${Math.round(x)}px`;
    bar.classList.toggle("is-below", below);
    bar.classList.remove("is-off");
  }

  // pokazanie: puszczenie myszy po przeciągnięciu po komórkach (ui-controls.js) albo Shift+klik
  window.addEventListener("pointerup", (e) => {
    if (e.pointerType !== "mouse" || e.button !== 0) return;
    const drag = Date.now() - dragSelEndedAt < 120;
    if (!drag && !(e.shiftKey && tableEl.contains(/** @type {Node} */ (e.target)))) return;
    setTimeout(() => {
      if (!on() || !selRect()) return;
      wanted = true;
      place();
    }, 0);
  });
  document.addEventListener("pointerdown", (e) => { if (wanted && !bar.contains(/** @type {Node} */ (e.target))) hideBar(); }, true);
  document.addEventListener("keydown", (e) => {
    if (!wanted || bar.contains(/** @type {Node} */ (e.target)) || ["Shift", "Control", "Alt", "Meta"].includes(e.key)) return;
    hideBar();
  }, true);
  let st = 0;
  tableWrapEl.addEventListener("scroll", () => {
    if (!wanted) return;
    bar.classList.add("is-off");
    clearTimeout(st);
    st = setTimeout(place, 160);
  }, { passive: true });
  window.addEventListener("resize", () => { if (wanted) place(); }, { passive: true });
  document.getElementById("sheetModeSwitch")?.addEventListener("click", hideBar);

  return { hide: hideBar, place, el: bar };
})();
