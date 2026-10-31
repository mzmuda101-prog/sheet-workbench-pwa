// app/move-cells.js — PRZENOSZENIE KOMÓREK PRZECIĄGNIĘCIEM (2026-10-31).
//
// Excel na telefonie: przytrzymaj zaznaczenie, aż zaczną „biegać” kropki, i przeciągnij je w nowe
// miejsce. U nas przytrzymanie otwiera menu, a pozycja „Przenieś (przeciągnij)” nakłada na blok
// ramkę z touch-action: none — przeciągasz ją palcem albo myszą, puszczasz w celu i blok się
// przenosi (swbMoveBlock w table.js: wartości/formuły + format, źródło puste, formuły zależne
// przepięte, jeden krok Cofnij — ta sama droga co Wytnij → Wklej).
// Dlaczego nie „przytrzymaj i od razu ciągnij”: w przeglądarce ruch palca po siatce to
// przewijanie tabeli, a zablokować je w połowie gestu da się tylko globalnym touchmove
// z passive:false — to właśnie psuło płynne przewijanie na iPadzie (ui-controls.js, resize).
// Ramka ma touch-action ustawione PRZED gestem, więc przewijanie i przenoszenie się nie gryzą.

const swbMoveCells = (() => {
  if (typeof document === "undefined" || typeof tableWrapEl === "undefined" || !tableWrapEl) return { start() {}, cancel() {}, get active() { return false; } };
  const frame = document.createElement("div");
  frame.className = "move-box";
  frame.setAttribute("aria-hidden", "true");
  // podgląd celu: JEDEN obrys całego bloku (jak w Excelu), nie ramka na każdej komórce
  const ghost = document.createElement("div");
  ghost.className = "move-ghost";
  ghost.setAttribute("aria-hidden", "true");
  let mode = null; // { block, drag: null | { grab, target, id, last, raf } }

  const keyRow = (key) => {
    const m = /^(?:wide|pre):(\d+)$/.exec(String(key || ""));
    return m ? Number(m[1]) : null;
  };
  const tdAt = (r, c) => tableEl.querySelector(`tr[data-row-key="wide:${r}"] > td[data-col-index="${c}"], tr[data-row-key="pre:${r}"] > td[data-col-index="${c}"]`);
  const cellAtPoint = (x, y) => {
    const el = document.elementFromPoint(x, y);
    const td = el && el.closest ? el.closest("#dataTable tr[data-row-key] > td[data-col-index]") : null;
    if (!td || td.classList.contains("row-head")) return null;
    const r = keyRow(td.parentElement.dataset.rowKey);
    const c = parseInt(td.dataset.colIndex, 10);
    return r == null || !Number.isFinite(c) ? null : { r, c };
  };

  // Ramka nad blokiem komórek (r1..r2, c1..c2) w układzie #tableWrap; false = nie narysowany.
  function boxOver(el, r1, c1, r2, c2) {
    const tl = tdAt(r1, c1);
    const br = tdAt(r2, c2) || tl;
    if (!tl) { el.style.display = "none"; return false; }
    el.style.display = "";
    const wrap = tableWrapEl.getBoundingClientRect();
    const a = tl.getBoundingClientRect();
    const z = br.getBoundingClientRect();
    const x = a.left - wrap.left - tableWrapEl.clientLeft + tableWrapEl.scrollLeft;
    const y = a.top - wrap.top - tableWrapEl.clientTop + tableWrapEl.scrollTop;
    el.style.transform = `translate(${Math.round(x)}px, ${Math.round(y)}px)`;
    el.style.width = `${Math.round(z.right - a.left)}px`;
    el.style.height = `${Math.round(z.bottom - a.top)}px`;
    return true;
  }
  function place() {
    if (!mode) return;
    const b = mode.block;
    boxOver(frame, b.r1, b.c1, b.r2, b.c2);
  }

  function clearPreview() {
    ghost.remove();
  }
  function preview(target) {
    if (!mode || !target) { clearPreview(); return; }
    const b = mode.block;
    if (ghost.parentElement !== tableWrapEl) tableWrapEl.appendChild(ghost);
    if (!boxOver(ghost, target.r, target.c, target.r + (b.r2 - b.r1), target.c + (b.c2 - b.c1))) clearPreview();
  }

  function start() {
    if (mode) cancel();
    const block = typeof swbSelectionSheetBlock === "function" ? swbSelectionSheetBlock() : null;
    if (!block) return false;
    if (typeof swbCancelCut === "function") swbCancelCut();
    mode = { block: { r1: block.r1, r2: block.r2, c1: block.c1, c2: block.c2 }, drag: null };
    frame.style.pointerEvents = "";
    tableWrapEl.appendChild(frame);
    place();
    document.documentElement.classList.add("swb-moving");
    toast(t("moveHint"), "info");
    return true;
  }

  function cancel() {
    if (!mode) return false;
    if (mode.drag) {
      cancelAnimationFrame(mode.drag.raf);
      window.removeEventListener("pointermove", onMove, true);
      window.removeEventListener("pointerup", onUp, true);
      window.removeEventListener("pointercancel", onCancel, true);
    }
    mode = null;
    clearPreview();
    frame.remove();
    document.documentElement.classList.remove("swb-moving");
    return true;
  }

  function autoScroll() {
    if (!mode || !mode.drag) return;
    const box = tableWrapEl.getBoundingClientRect();
    const { x, y } = mode.drag.last;
    let dx = 0;
    let dy = 0;
    if (y > box.bottom - 28) dy = 16;
    else if (y < box.top + 40) dy = -16;
    if (x > box.right - 28) dx = 16;
    else if (x < box.left + 28) dx = -16;
    if (dx || dy) {
      tableWrapEl.scrollBy(dx, dy);
      update(Math.min(Math.max(x, box.left + 4), box.right - 4), Math.min(Math.max(y, box.top + 2), box.bottom - 4));
    }
    mode.drag.raf = requestAnimationFrame(autoScroll);
  }

  function update(x, y) {
    const hit = cellAtPoint(x, y);
    if (!hit || !mode || !mode.drag) return;
    const target = { r: hit.r - mode.drag.grab.dr, c: hit.c - mode.drag.grab.dc };
    if (mode.drag.target && mode.drag.target.r === target.r && mode.drag.target.c === target.c) return;
    mode.drag.target = target;
    preview(target);
  }

  frame.addEventListener("pointerdown", (e) => {
    if (!mode || e.button !== 0) return;
    e.preventDefault();
    e.stopPropagation();
    // trafienie komórki pod ramką: ramka na czas przeciągania przepuszcza wskaźnik
    frame.style.pointerEvents = "none";
    const hit = cellAtPoint(e.clientX, e.clientY);
    const b = mode.block;
    const grab = hit && hit.r >= b.r1 && hit.r <= b.r2 && hit.c >= b.c1 && hit.c <= b.c2
      ? { dr: hit.r - b.r1, dc: hit.c - b.c1 } : { dr: 0, dc: 0 };
    mode.drag = { grab, target: null, id: e.pointerId, last: { x: e.clientX, y: e.clientY }, raf: 0 };
    window.addEventListener("pointermove", onMove, true);
    window.addEventListener("pointerup", onUp, true);
    window.addEventListener("pointercancel", onCancel, true);
    mode.drag.raf = requestAnimationFrame(autoScroll);
  });
  function onMove(e) {
    if (!mode || !mode.drag || e.pointerId !== mode.drag.id) return;
    e.preventDefault();
    mode.drag.last = { x: e.clientX, y: e.clientY };
    update(e.clientX, e.clientY);
  }
  function endDrag() {
    const d = mode && mode.drag;
    if (!d) return null;
    cancelAnimationFrame(d.raf);
    window.removeEventListener("pointermove", onMove, true);
    window.removeEventListener("pointerup", onUp, true);
    window.removeEventListener("pointercancel", onCancel, true);
    mode.drag = null;
    frame.style.pointerEvents = "";
    clearPreview();
    return d;
  }
  function onUp(e) {
    if (!mode || !mode.drag || e.pointerId !== mode.drag.id) return;
    const d = endDrag();
    const b = mode.block;
    if (!d.target || (d.target.r === b.r1 && d.target.c === b.c1)) { place(); return; } // bez ruchu — tryb trwa
    const moved = swbMoveBlock(b, d.target.r - b.r1, d.target.c - b.c1);
    if (moved) cancel();
    else place();
  }
  function onCancel(e) {
    if (!mode || !mode.drag || e.pointerId !== mode.drag.id) return;
    endDrag();
    place();
  }

  // Esc albo stuknięcie / klik poza ramką kończy tryb (jak „odłożenie” w Excelu)
  document.addEventListener("keydown", (e) => {
    if (e.key === "Escape" && mode) { e.preventDefault(); e.stopPropagation(); cancel(); }
  }, true);
  document.addEventListener("pointerdown", (e) => {
    if (mode && !mode.drag && e.target !== frame) cancel();
  }, true);
  if (typeof ResizeObserver === "function" && typeof tableEl !== "undefined" && tableEl) {
    new ResizeObserver(() => { if (mode && !mode.drag) place(); }).observe(tableEl);
  }

  return { start, cancel, place, get active() { return !!mode; } };
})();
