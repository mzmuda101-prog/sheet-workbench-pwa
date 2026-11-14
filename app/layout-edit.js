// layout-edit.js — przeciągnięta szerokość kolumny / wysokość wiersza trafia do pliku (docs/OD-ZERA.md, poz. 1).
//
// Dawniej przeciąganie zmieniało tylko widok. Teraz w Edycji (jak w Excelu):
//   • zmiana idzie do pendingLayout[arkusz] = { cols: { kolumna0: szerokość }, rows: { wiersz0: pkt } }
//     i do metadanych arkusza SheetJS (!cols / !rows) — po przełączeniu arkusza i powrocie zostaje;
//   • zapis dopisuje ją do pliku (xlsx-layout.js, wołane z buildPatchedXlsx);
//   • Cofnij / Ponów jak każda zmiana; plik „niezapisany”, szkic (drafts.js) ją pamięta.
// W Podglądzie przeciąganie dalej zmienia tylko widok (Podgląd nie zmienia pliku).

const swbLayoutEdit = (() => {
  const zoom = () => (typeof zoomLevelEl !== "undefined" && parseFloat(zoomLevelEl.value)) || 1;
  const sheetMeta = (ws, kind) => { const k = kind === "cols" ? "!cols" : "!rows"; if (!Array.isArray(ws[k])) ws[k] = []; return ws[k]; };
  const clone = (v) => (v && typeof v === "object" ? { ...v } : v);

  // Stan jednej pozycji (do Cofnij): ręczny rozmiar w widoku, metadane SheetJS, wpis do zapisu.
  function capture(sheet, kind, key) {
    const ws = workbook.Sheets[sheet];
    const own = sheet === currentSheetName;
    const i = kind === "cols" ? key - currentStartCol : key;
    const manual = own ? (kind === "cols" ? manualColumnWidths[i] : manualRowHeights[i]) : undefined;
    const pend = pendingLayout[sheet] && pendingLayout[sheet][kind];
    return { manual, meta: clone(sheetMeta(ws, kind)[key]), has: !!pend && key in pend, pending: pend ? pend[key] : undefined };
  }
  function restore(sheet, kind, key, st) {
    const ws = workbook.Sheets[sheet];
    const own = sheet === currentSheetName;
    const i = kind === "cols" ? key - currentStartCol : key;
    const meta = sheetMeta(ws, kind);
    if (st.meta) meta[key] = clone(st.meta); else delete meta[key];
    if (!pendingLayout[sheet]) pendingLayout[sheet] = { cols: {}, rows: {} };
    if (st.has) pendingLayout[sheet][kind][key] = st.pending; else delete pendingLayout[sheet][kind][key];
    if (!own) return;
    if (typeof invalidateColumnWidthCache === "function") invalidateColumnWidthCache(); // klucz pamięci nie widzi zmiany szerokości z pliku
    const map = kind === "cols" ? manualColumnWidths : manualRowHeights;
    if (st.manual) map[i] = st.manual; else delete map[i];
    if (kind === "cols") { if (Array.isArray(currentSheetColWidths)) currentSheetColWidths[i] = meta[key] || null; }
    else if (currentSheetRowHeights) { if (meta[key]) currentSheetRowHeights[key] = meta[key]; else delete currentSheetRowHeights[key]; }
  }

  // Koniec przeciągania. kind "cols": i = kolumna widoku, px = szerokość na ekranie; "rows": i = wiersz0.
  // prevManual: ręczny rozmiar sprzed przeciągania (w trakcie mapa ma już nowy) — do Cofnij
  function commit(kind, i, px, prevManual) {
    if (typeof workbook === "undefined" || !workbook || !currentSheetName || !(px > 0)) return false;
    if (typeof swbIsEditMode === "function" && !swbIsEditMode()) return false; // Podgląd: tylko widok
    if (kind === "cols" && (typeof swbRealHeaderCount === "function" ? i >= swbRealHeaderCount() : false)) return false; // kolumna wyliczana — nie ma jej w pliku
    const sheet = currentSheetName;
    const key = kind === "cols" ? currentStartCol + i : i;
    const before = capture(sheet, kind, key);
    before.manual = prevManual;
    const base = px / zoom();
    const ws = workbook.Sheets[sheet];
    if (!pendingLayout[sheet]) pendingLayout[sheet] = { cols: {}, rows: {} };
    if (kind === "cols") {
      const width = swbLayout.pxToWidth(base);
      pendingLayout[sheet].cols[key] = width;
      sheetMeta(ws, "cols")[key] = { ...(before.meta || {}), width, wpx: swbLayout.widthToPx(width), wch: Math.max(0, Math.round((width * 6 - 5) / 6)), MDW: 6, customwidth: "1" };
      if (Array.isArray(currentSheetColWidths)) currentSheetColWidths[i] = sheetMeta(ws, "cols")[key];
    } else {
      const pt = swbLayout.pxToPt(base);
      pendingLayout[sheet].rows[key] = pt;
      sheetMeta(ws, "rows")[key] = { ...(before.meta || {}), hpt: pt, hpx: pt };
      if (currentSheetRowHeights) currentSheetRowHeights[key] = sheetMeta(ws, "rows")[key];
    }
    undoStack.push({ workbook, sheet, changes: [], layout: { kind, key, prev: before } });
    while (undoStack.length > UNDO_LIMIT) undoStack.shift();
    redoStack = [];
    if (typeof updateUndoButtons === "function") updateUndoButtons();
    setDirtyState(true);
    return true;
  }

  // Krok Cofnij/Ponów z undo.js: przywraca stan sprzed zmiany i zwraca krok odwrotny.
  function applyStep(step) {
    const { kind, key, prev } = step.layout;
    if (!workbook.Sheets[step.sheet]) return null;
    const now = capture(step.sheet, kind, key);
    restore(step.sheet, kind, key, prev);
    return { workbook: step.workbook, sheet: step.sheet, changes: [], layout: { kind, key, prev: now } };
  }

  function count() {
    let n = 0;
    Object.values(pendingLayout || {}).forEach((l) => { if (l) n += Object.keys(l.cols || {}).length + Object.keys(l.rows || {}).length; });
    return n;
  }

  // Szkic (drafts.js): układ z odzyskanego szkicu z powrotem do zapisu i do arkuszy.
  function restoreAll(layout) {
    pendingLayout = {};
    Object.entries(layout || {}).forEach(([sheet, l]) => {
      const ws = workbook && workbook.Sheets[sheet];
      if (!ws || !l) return;
      pendingLayout[sheet] = { cols: { ...(l.cols || {}) }, rows: { ...(l.rows || {}) } };
      Object.entries(l.cols || {}).forEach(([c, width]) => { sheetMeta(ws, "cols")[c] = { ...(sheetMeta(ws, "cols")[c] || {}), width, wpx: swbLayout.widthToPx(width), MDW: 6, customwidth: "1" }; });
      Object.entries(l.rows || {}).forEach(([r, pt]) => { sheetMeta(ws, "rows")[r] = { ...(sheetMeta(ws, "rows")[r] || {}), hpt: pt, hpx: pt }; });
    });
  }

  return { commit, applyStep, count, restoreAll };
})();
