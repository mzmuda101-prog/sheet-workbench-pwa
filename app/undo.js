// undo.js — Cofnij / Ponów dla edycji komórek (Ctrl/⌘+Z, Ctrl/⌘+Shift+Z, Ctrl+Y).
//
// Jak to działa: KAŻDA zmiana wartości komórki przechodzi przez updateSheetCell()
// (edytor, wklejanie, „Wypełnij w dół", narzędzia edycji). Tam, tuż przed zmianą,
// woła się undoCaptureCell() — zapisujemy stan „przed": obiekt komórki arkusza,
// wpis w pendingEdits (to, co trafi do pliku przy zapisie) i wartość w wierszu tabeli.
//
// Krok = wszystko, co zmieniło się w JEDNYM zadaniu JS (zamykamy go mikro-zadaniem).
// Dzięki temu wklejenie 200 komórek albo „Znajdź i zamień" to jedno Cofnij, a żaden
// z tych modułów nie musiał się o tym dowiadywać.
//
// Cofnięcie przywraca dokładnie zapisany stan — łącznie z tym, że komórka wraca do
// „nieedytowanej" (znika z pendingEdits), więc zapis nie dotknie jej w pliku, a formuła
// albo styl z oryginału zostają nietknięte. Ponów działa tak samo w drugą stronę.

const UNDO_LIMIT = 100;
let undoStack = [];
let redoStack = [];
let undoOpenStep = null;
let undoApplying = false;

const undoBtnEl = document.getElementById("undoBtn");
const redoBtnEl = document.getElementById("redoBtn");

function undoRowMap() {
  const map = new Map();
  for (const row of baseRows) {
    if (row && !row.isLongViewRow && Number.isFinite(row.rowIndex0) && !map.has(row.rowIndex0)) map.set(row.rowIndex0, row);
  }
  return map;
}

function undoSnapshot(sheetName, cellRef, rowIndex0, col, rowMap) {
  const sheet = workbook?.Sheets?.[sheetName];
  const cell = sheet && sheet[cellRef] ? { ...sheet[cellRef] } : null;
  const pend = pendingEdits[sheetName];
  const pendingHas = !!pend && Object.prototype.hasOwnProperty.call(pend, cellRef);
  const snap = { cell, pendingHas, pending: pendingHas ? pend[cellRef] : undefined, rowVals: null };
  const row = rowMap ? rowMap.get(rowIndex0) : null;
  if (row) {
    snap.rowVals = {
      v: Array.isArray(row.values) ? row.values[col] : undefined,
      r: Array.isArray(row.rawValues) ? row.rawValues[col] : undefined,
      d: Array.isArray(row.display) ? row.display[col] : undefined,
    };
  }
  return snap;
}

function undoValueKey(cell) {
  if (!cell) return "∅";
  const v = cell.v instanceof Date ? `d${cell.v.getTime()}` : String(cell.v);
  return `${cell.t}|${v}|${cell.f || ""}`;
}

// Wołane z updateSheetCell() PRZED zmianą.
function undoCaptureCell(rowIndex0, colIndex0) {
  if (undoApplying || !workbook || !currentSheetName) return;
  const cellRef = XLSX.utils.encode_cell({ r: rowIndex0, c: currentStartCol + colIndex0 });
  if (!undoOpenStep) {
    undoOpenStep = { workbook, sheet: currentSheetName, changes: [], seen: new Set(), rowMap: null };
    queueMicrotask(undoCloseStep);
  }
  const step = undoOpenStep;
  if (step.sheet !== currentSheetName || step.seen.has(cellRef)) return; // liczy się PIERWSZY stan „przed"
  step.seen.add(cellRef);
  if (!step.rowMap) step.rowMap = undoRowMap();
  step.changes.push({ cellRef, rowIndex0, col: colIndex0, snap: undoSnapshot(step.sheet, cellRef, rowIndex0, colIndex0, step.rowMap) });
}

function undoCloseStep() {
  const step = undoOpenStep;
  undoOpenStep = null;
  if (!step || !step.changes.length) return;
  // Odrzuć komórki, które po całej operacji wyglądają jak przedtem (np. formuła odrzucona)
  // — ale TYLKO gdy nie przybył im też wpis w pendingEdits. Wklejenie tej samej wartości
  // nie zmienia komórki, a jednak oznacza ją do zapisu; bez tego warunku Cofnij zostawiało
  // takie „ciche" wpisy i zapis przepisywał komórkę, której user już nie chciał ruszać.
  const sheet = workbook?.Sheets?.[step.sheet];
  const pendNow = pendingEdits[step.sheet];
  step.changes = step.changes.filter((ch) =>
    undoValueKey(ch.snap.cell) !== undoValueKey(sheet ? sheet[ch.cellRef] : null)
    || ch.snap.pendingHas !== (!!pendNow && Object.prototype.hasOwnProperty.call(pendNow, ch.cellRef)));
  delete step.seen;
  delete step.rowMap;
  if (!step.changes.length) return;
  undoStack.push(step);
  if (undoStack.length > UNDO_LIMIT) undoStack.shift();
  redoStack = [];
  updateUndoButtons();
}

// Przywraca stan zapisany w kroku; zwraca krok odwrotny (do drugiego stosu).
function undoApplyStep(step) {
  const sheet = workbook?.Sheets?.[step.sheet];
  if (!sheet) return null;
  const rowMap = step.sheet === currentSheetName ? undoRowMap() : null;
  const inverse = { workbook: step.workbook, sheet: step.sheet, changes: [] };
  undoApplying = true;
  try {
    for (let i = step.changes.length - 1; i >= 0; i--) {
      const ch = step.changes[i];
      inverse.changes.unshift({ ...ch, snap: undoSnapshot(step.sheet, ch.cellRef, ch.rowIndex0, ch.col, rowMap) });
      if (ch.snap.cell) sheet[ch.cellRef] = { ...ch.snap.cell };
      else delete sheet[ch.cellRef];
      if (!pendingEdits[step.sheet]) pendingEdits[step.sheet] = {};
      if (ch.snap.pendingHas) pendingEdits[step.sheet][ch.cellRef] = ch.snap.pending;
      else delete pendingEdits[step.sheet][ch.cellRef];
      const row = rowMap ? rowMap.get(ch.rowIndex0) : null;
      if (row) {
        // Wartości wiersza: dokładnie te sprzed zmiany; gdy ich nie mamy (krok nagrany
        // na innym arkuszu), odtwarzamy z komórki.
        const rv = ch.snap.rowVals;
        const v = rv ? rv.v : (ch.snap.cell ? ch.snap.cell.v : null);
        if (Array.isArray(row.values)) row.values[ch.col] = v;
        if (Array.isArray(row.rawValues)) row.rawValues[ch.col] = rv ? rv.r : v;
        if (Array.isArray(row.display)) row.display[ch.col] = rv ? rv.d : (v == null ? "" : toDisplay(v));
      }
    }
  } finally {
    undoApplying = false;
  }
  bumpSheetDataStamp();
  return inverse;
}

function undoDropStale() {
  undoStack = undoStack.filter((s) => s.workbook === workbook);
  redoStack = redoStack.filter((s) => s.workbook === workbook);
}

function undoRedo(direction) {
  undoCloseStep(); // domknij krok z bieżącego zadania, jeśli jest
  undoDropStale(); // inny plik → stara historia nie ma sensu
  const from = direction === "undo" ? undoStack : redoStack;
  const to = direction === "undo" ? redoStack : undoStack;
  const step = from.pop();
  if (!step) {
    toast(t(direction === "undo" ? "undoNothing" : "redoNothing"), "info");
    updateUndoButtons();
    return false;
  }
  const inverse = undoApplyStep(step);
  if (inverse) to.push(inverse);
  setDirtyState(true);

  const count = step.changes.length;
  if (step.sheet === currentSheetName) {
    renderActiveTable();
    scheduleViewRefresh({ analyses: true, filterBadge: true });
    // Pokaż, co się zmieniło: pierwsza komórka kroku (o ile jest w widoku).
    const first = step.changes[0];
    const rowKey = `wide:${first.rowIndex0}`;
    if (currentDisplayModel?.mode === "wide" && currentDisplayModel.rows.some((r) => getRowSelectionKey(r) === rowKey)) {
      setFocusedCell(rowKey, first.col, { scroll: true });
    }
    toast(t(direction === "undo" ? "undoDone" : "redoDone", { count }), "info");
  } else {
    toast(t(direction === "undo" ? "undoDoneSheet" : "redoDoneSheet", { count, sheet: step.sheet }), "info");
  }
  updateUndoButtons();
  return true;
}

function undoLast() { return undoRedo("undo"); }
function redoLast() { return undoRedo("redo"); }

function updateUndoButtons() {
  const canUndo = undoStack.some((s) => s.workbook === workbook);
  const canRedo = redoStack.some((s) => s.workbook === workbook);
  if (undoBtnEl) {
    undoBtnEl.classList.toggle("hidden", !canUndo && !canRedo);
    undoBtnEl.disabled = !canUndo;
  }
  if (redoBtnEl) {
    redoBtnEl.classList.toggle("hidden", !canRedo);
    redoBtnEl.disabled = !canRedo;
  }
}

if (undoBtnEl) undoBtnEl.addEventListener("click", undoLast);
if (redoBtnEl) redoBtnEl.addEventListener("click", redoLast);

// Skróty: w polach tekstowych (edytor komórki, szukanie, panel) zostaje natywne
// cofanie PISANIA — tu obsługujemy tylko cofanie zmian w arkuszu.
document.addEventListener("keydown", (e) => {
  if (!(e.ctrlKey || e.metaKey) || e.altKey) return;
  const isZ = e.code === "KeyZ" || String(e.key).toLowerCase() === "z";
  const isY = (e.code === "KeyY" || String(e.key).toLowerCase() === "y") && !e.metaKey;
  if (!isZ && !isY) return;
  const a = document.activeElement;
  const tag = String(a?.tagName || "").toLowerCase();
  if (a && (a.isContentEditable || tag === "textarea" || (tag === "input" && !["checkbox", "radio", "button"].includes(a.type)))) return;
  const trOverlay = document.getElementById("transcribeOverlay");
  if (trOverlay && !trOverlay.classList.contains("hidden")) return; // Spisywanie ma własne cofanie
  if (!workbook) return;
  e.preventDefault();
  if (isY || e.shiftKey) redoLast();
  else undoLast();
});
