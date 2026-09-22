// freeze-pane.js — „SZYBA" ZAMROŻONEJ KOLUMNY NA CZAS PRZEWIJANIA.
//
// PROBLEM: blokada 1. kolumny to `position: sticky` na DWÓCH komórkach w KAŻDYM wierszu
// (numer wiersza + 1. kolumna danych). Przeglądarka robi z każdej przyklejonej komórki
// OSOBNĄ warstwę — przy 200 wierszach to ~400 warstw, które w każdej klatce przewijania
// trzeba przeliczyć, złożyć i często przemalować. Zmierzone (Chromium, CPU ×6, 200 w.):
// przewijanie w bok 185 ms pracy bez blokady → 604 ms z blokadą; w pionie 884 → 1195 ms.
// Excel nie ma tego problemu, bo zamrożony obszar to dla niego JEDNA osobna „szyba".
//
// ROZWIĄZANIE (hybryda, żeby nie ruszać interakcji):
//   • W SPOCZYNKU nic się nie zmienia — prawdziwe przyklejone komórki, pełna obsługa
//     kliknięć, zaznaczania, edycji, klawiatury, dokładnie jak wcześniej.
//   • NA CZAS PRZEWIJANIA przyklejenie z komórek zdejmujemy (klasa `fz-off` na nich),
//     a nad tabelą kładziemy jedną przyklejoną warstwę z kopią dwóch pierwszych
//     kolumn. Jedna warstwa zamiast setek → przewijanie kosztuje tyle, co bez blokady.
//   • Po ~180 ms bez ruchu szyba znika i wracają prawdziwe przyklejone komórki.
//
// BEZPIECZNIKI (wtedy zachowanie = jak przed zmianą, bez szyby):
//   • scalone komórki w danych (kopia nie odwzorowałaby wiernie nth-child),
//   • otwarty edytor komórki albo fokus na zamrożonej komórce (obwódka fokusu),
//   • przebudowa wierszy tabeli w trakcie przewijania (np. filtr) — szyba znika
//     i nie wraca aż do spoczynku; zmiany klas/treści w zamrożonych komórkach
//     (np. zaznaczenie strzałkami) są na bieżąco przenoszone do kopii.
// Kopia NIE przyjmuje kliknięć (pointer-events: none) — szczegóły przy obsłudze gestów.

(function () {
  if (typeof tableWrapEl === "undefined" || !tableWrapEl || !tableEl || !tbodyEl) return;
  // ?nopane = stare zachowanie (samo sticky) — do porównania A/B na telefonie.
  try { if (new URLSearchParams(location.search).has("nopane")) return; } catch (_) {}

  const IDLE_MS = 180;
  let active = false;
  let suppressed = false;   // po przebudowie tabeli w trakcie przewijania — do spoczynku
  let idleTimer = 0;
  let pane = null;
  let cloneTable = null;
  let cloneBody = null;
  let observer = null;
  let offCells = [];        // prawdziwe zamrożone komórki z chwilowo zdjętym sticky

  // Fokus na zamrożonej komórce (chodzenie strzałkami): kopia nie ma :focus, więc
  // obwódka fokusu zniknęłaby na czas przewijania — wtedy zostajemy przy sticky.
  const focusInFrozen = () => {
    const a = document.activeElement;
    return !!(a && a.tagName === "TD" && a.cellIndex < 2 && a.parentElement && a.parentElement.parentElement === tbodyEl);
  };

  const eligible = () =>
    tableWrapEl.classList.contains("freeze-first-col")
    && !focusInFrozen()
    && !tableWrapEl.classList.contains("hidden")
    && tbodyEl.rows.length > 0
    && !tbodyEl.querySelector("input.cell-editor, td[colspan], td[rowspan]");

  function cloneCell(td) {
    const c = td.cloneNode(true);
    c.removeAttribute("id");
    c.removeAttribute("tabindex");
    c.querySelectorAll("[id]").forEach((el) => el.removeAttribute("id"));
    c.querySelectorAll("[tabindex]").forEach((el) => el.removeAttribute("tabindex"));
    return c;
  }

  function build() {
    const rows = tbodyEl.rows;
    const first = rows[0];
    if (!first || first.cells.length < 2) return false;
    // Szerokości z faktycznego układu (table-layout: fixed potrafi rozciągnąć kolumny).
    const w0 = first.cells[0].getBoundingClientRect().width;
    const w1 = first.cells[1].getBoundingClientRect().width;
    const heights = new Array(rows.length);
    for (let i = 0; i < rows.length; i++) heights[i] = rows[i].getBoundingClientRect().height;

    pane = document.createElement("div");
    pane.className = "fz-pane";
    pane.setAttribute("aria-hidden", "true");
    pane.style.width = `${w0 + w1}px`;

    cloneTable = document.createElement("table");
    cloneTable.className = `${tableEl.className} fz-clone`.trim();
    cloneTable.style.width = `${w0 + w1}px`;
    cloneTable.style.top = `${tbodyEl.offsetTop}px`;
    const colgroup = document.createElement("colgroup");
    [w0, w1].forEach((w) => {
      const col = document.createElement("col");
      col.style.width = `${w}px`;
      colgroup.appendChild(col);
    });
    cloneBody = document.createElement("tbody");
    cloneBody.className = tbodyEl.className;
    const frag = document.createDocumentFragment();
    for (let i = 0; i < rows.length; i++) {
      const tr = rows[i];
      const ctr = document.createElement("tr");
      ctr.className = tr.className;
      const st = tr.getAttribute("style");
      if (st) ctr.setAttribute("style", st);
      ctr.style.height = `${heights[i]}px`;
      if (tr.cells[0]) ctr.appendChild(cloneCell(tr.cells[0]));
      if (tr.cells[1]) ctr.appendChild(cloneCell(tr.cells[1]));
      frag.appendChild(ctr);
    }
    cloneBody.appendChild(frag);
    cloneTable.append(colgroup, cloneBody);
    pane.appendChild(cloneTable);
    return true;
  }

  // Zmiany w trakcie przewijania: klasy/treść zamrożonych komórek przenosimy do kopii
  // (np. zaznaczenie przesuwane strzałkami). Przebudowa wierszy = koniec szyby.
  function onMutations(records) {
    if (!active) return;
    for (const r of records) {
      if (r.type === "childList" && (r.target === tbodyEl || r.target === tableEl)) { stop(true); return; }
      let node = r.target.nodeType === 1 ? r.target : r.target.parentElement;
      if (!node) continue;
      if (node === tbodyEl) { if (cloneBody) cloneBody.className = tbodyEl.className; continue; }
      if (node === tableEl) { if (cloneTable) cloneTable.className = `${tableEl.className} fz-clone`.trim(); continue; }
      if (node.tagName === "TR" && node.parentElement === tbodyEl) {
        const ctr = cloneBody && cloneBody.rows[node.sectionRowIndex];
        if (ctr) ctr.className = node.className;
        continue;
      }
      const td = node.closest("td");
      if (!td || td.parentElement.parentElement !== tbodyEl) continue;
      if (td.querySelector("input.cell-editor")) { stop(true); return; }
      if (td.cellIndex > 1) continue;
      const ctr = cloneBody && cloneBody.rows[td.parentElement.sectionRowIndex];
      const old = ctr && ctr.cells[td.cellIndex];
      if (old) old.replaceWith(cloneCell(td));
    }
  }

  function start() {
    if (!eligible() || !build()) return;
    active = true;
    tableWrapEl.insertBefore(pane, tableEl);
    // Klasa na SAMYCH zamrożonych komórkach, nie na kontenerze: przełączenie klasy na
    // .table-wrap kazało przeliczyć style wszystkich ~4000 komórek (zmierzone: ~50 ms
    // stylu przy CPU ×6 na starcie każdego przewijania), tu tylko ~400.
    const rows = tbodyEl.rows;
    offCells = [];
    for (let i = 0; i < rows.length; i++) {
      const c = rows[i].cells;
      if (c[0]) offCells.push(c[0]);
      if (c[1]) offCells.push(c[1]);
    }
    for (let i = 0; i < offCells.length; i++) offCells[i].classList.add("fz-off");
    observer = observer || new MutationObserver(onMutations);
    observer.observe(tableEl, { subtree: true, childList: true, attributes: true, characterData: true });
  }

  function stop(suppress) {
    clearTimeout(idleTimer);
    idleTimer = 0;
    if (suppress) { suppressed = true; idleTimer = setTimeout(onIdle, IDLE_MS); }
    if (!active) return;
    active = false;
    if (observer) { onMutations(observer.takeRecords()); observer.disconnect(); }
    for (let i = 0; i < offCells.length; i++) offCells[i].classList.remove("fz-off");
    offCells = [];
    if (pane) pane.remove();
    pane = cloneTable = cloneBody = null;
  }

  function onIdle() {
    idleTimer = 0;
    suppressed = false;
    stop(false);
  }

  tableWrapEl.addEventListener("scroll", () => {
    if (!active && !suppressed) start();
    clearTimeout(idleTimer);
    idleTimer = setTimeout(onIdle, IDLE_MS);
  }, { passive: true });

  tbodyEl.addEventListener("focusin", () => { if (active && focusInFrozen()) stop(true); });

  // Szyba jest przezroczysta dla wskaźnika (pointer-events: none), więc przewijanie
  // palcem zaczęte na zamrożonej kolumnie działa normalnie. Ale gdy w trybie szyby
  // przewinięto w bok, POD szybą leżą INNE kolumny — tap w szybę trafiłby w nie.
  // Dlatego gest zaczęty na szybie połykamy (tylko obsługa tabeli go nie widzi;
  // natywne przewijanie działa dalej). Tak jak tap w trakcie rozpędu na iOS: tylko
  // zatrzymuje ruch. Kolejny tap — już w spoczynku — trafia w prawdziwą komórkę.
  let swallowGesture = false;
  const inPane = (e) => {
    if (!active || !pane) return false;
    const pt = e.touches && e.touches[0] ? e.touches[0] : e;
    const r = pane.getBoundingClientRect();
    return pt.clientX >= r.left && pt.clientX < r.left + pane.offsetWidth
      && pt.clientY >= tableWrapEl.getBoundingClientRect().top;
  };
  const onDown = (e) => {
    if (e.type === "pointerdown" || e.type === "touchstart") swallowGesture = inPane(e);
    else if (e.type === "mousedown" && !swallowGesture) swallowGesture = inPane(e);
    if (swallowGesture) e.stopPropagation();
  };
  const onRest = (e) => {
    if (!swallowGesture) return;
    e.stopPropagation();
    if (e.type === "pointerup" || e.type === "pointercancel" || e.type === "touchend" || e.type === "touchcancel") {
      setTimeout(() => { swallowGesture = false; }, 350); // click przychodzi zaraz po pointerup
    }
    if (e.type === "click" || e.type === "dblclick") {
      e.preventDefault();
      if (e.type === "click") setTimeout(() => { swallowGesture = false; }, 0);
    }
  };
  ["pointerdown", "mousedown", "touchstart"].forEach((t) =>
    tableWrapEl.addEventListener(t, onDown, { capture: true, passive: true }));
  ["pointermove", "pointerup", "pointercancel", "mouseup", "touchmove", "touchend", "touchcancel", "click", "dblclick"].forEach((t) =>
    tableWrapEl.addEventListener(t, onRest, { capture: true, passive: t !== "click" && t !== "dblclick" }));

  // Dla testów i diagnostyki.
  window.__freezePane = { isActive: () => active, start: () => { if (!active) start(); return active; }, stop: () => stop(false) };
})();
