// entry.js — tryb „Wpisywanie”: przepisywanie fizycznych kartek / kart pracy DO arkusza.
//
// Odwrotność Spisywania (transcribe.js). Jedna kartka = jedna karta: cel (istniejący wiersz albo
// nowy) → pola → podgląd „było → będzie” → zatwierdzenie. Projekt:
// docs/superpowers/specs/2026-11-15-wpisywanie-design.md.
//
// Zasady, których pilnuje ten plik:
//   • zapis WYŁĄCZNIE przez writeCellFromInput (ta sama droga co wklejanie: parsowanie PL, styl,
//     pendingEdits, Cofnij), nowy wiersz WYŁĄCZNIE przez swbRowGrow.addDraft (styl i formuły
//     z wiersza wyżej, tabele Excela rosną przy zapisie);
//   • najpierw sprawdzamy CAŁĄ kartkę, dopiero potem piszemy — błąd w jednym polu = zero zapisów;
//   • puste pole nigdy nie czyści komórki; niepusta komórka zmienia się tylko z jawną zgodą
//     (overwrite); kolumny z formułą i wyliczane są tylko do odczytu;
//   • cel to numer wiersza ARKUSZA (rowIndex0), nigdy pozycja w widoku (sortowanie, filtr);
//   • arkusz zmieniony między podglądem a zatwierdzeniem → odmowa „stale”, nic nie piszemy.
// Wejście programowe swbEntry.fill({ target, fields }) — pod przyszłe skanowanie kart (OCR).

const swbEntry = (() => {
  const isBlank = (v) => v === null || v === undefined || (typeof v === "string" && v.trim() === "");
  let lastPlan = null;    // { key, before: {col: text} } — ostatni podgląd (do wykrycia „stale”)
  let lastCard = null;    // { top, steps } — ostatnia zatwierdzona kartka (Cofnij kartki)
  let filled = null;      // stan karty z fill()

  const ready = () => typeof workbook !== "undefined" && !!workbook && !!currentSheetName
    && typeof currentDisplayModel !== "undefined" && !!currentDisplayModel && currentDisplayModel.mode === "wide"
    && Array.isArray(currentDisplayModel.headers);

  const sheetCell = (rowIndex0, col) => {
    const ws = workbook.Sheets[currentSheetName];
    return ws ? ws[XLSX.utils.encode_cell({ r: rowIndex0, c: currentStartCol + col })] : null;
  };

  // Wiersze danych (bez wierszy widoku long i tego, co nad/w nagłówku), w kolejności arkusza.
  function candidates() {
    if (!ready() || !Array.isArray(baseRows)) return [];
    return baseRows
      .filter((r) => r && !r.isLongViewRow && Number.isFinite(r.rowIndex0) && r.rowIndex0 >= currentHeaderRow
        && Array.isArray(r.values) && r.values.some((v) => !isBlank(v)))
      .slice().sort((a, b) => a.rowIndex0 - b.rowIndex0);
  }
  const rowAt = (rowIndex0) => candidates().find((r) => r.rowIndex0 === rowIndex0) || null;
  const lastDataRow = () => { const c = candidates(); return c.length ? c[c.length - 1] : null; };

  const sameValue = (a, b) => {
    if (a instanceof Date && b instanceof Date) return a.getTime() === b.getTime();
    if (isBlank(a) && isBlank(b)) return true;
    if (typeof a === "string" && typeof b === "string") return a.trim() === b.trim();
    return a === b;
  };

  const keyOf = (target) => (target && target.newRow ? "new" : `r${target && target.rowIndex0}`);

  // Podgląd kartki: co się zmieni, co blokuje, przed czym ostrzec. Niczego nie zapisuje.
  function plan(input) {
    const out = { ok: false, changes: [], errors: [], warnings: [] };
    if (!ready() || !input || !input.target) return out;
    const isNew = !!input.target.newRow;
    const row = isNew ? null : rowAt(input.target.rowIndex0);
    if (!isNew && !row) { out.errors.push({ col: -1, reason: "no-target" }); return out; }
    // nowy wiersz: format, lista wyboru i formuły jak w ostatnim wierszu danych (tak rośnie tabela w Excelu)
    const ref = isNew ? lastDataRow() : row;
    const refIdx = ref ? ref.rowIndex0 : currentHeaderRow;
    const overwrite = new Set((input.overwrite || []).map(Number));
    const n = currentDisplayModel.headers.length;
    Object.keys(input.fields || {}).map(Number).sort((a, b) => a - b).forEach((col) => {
      const raw = input.fields[col];
      if (!Number.isInteger(col) || col < 0 || col >= n) return;
      if (isBlank(raw)) return; // puste pole nigdy nie czyści komórki
      const text = String(raw).trim();
      if (typeof cfmtIsRealCol === "function" && !cfmtIsRealCol(col)) { out.errors.push({ col, reason: "not-real-col" }); return; }
      const fcell = ref ? sheetCell(refIdx, col) : null;
      if (fcell && fcell.f) { out.errors.push({ col, reason: "formula" }); return; }
      const parsed = parseCellInput(text, refIdx, col);
      const current = row ? row.values[col] : null;
      if (row && parsed && sameValue(parsed.value, current)) return;
      const before = row ? String(getDisplayValue(row, col) ?? "") : "";
      const isOverwrite = !!row && !isBlank(current);
      const after = parsed && parsed.type === "formula" ? text : (parsed && parsed.value != null ? toDisplay(parsed.value) : text);
      if (isOverwrite && !overwrite.has(col)) out.errors.push({ col, reason: "overwrite-not-allowed" });
      let dv = null;
      try { dv = typeof getCellDataValidation === "function" ? getCellDataValidation(currentSheetName, refIdx, currentStartCol + col) : null; } catch { dv = null; }
      if (dv && !(parsed && parsed.type === "formula")) {
        const norm = String(parsed ? parsed.value : text).trim().toLowerCase();
        const hit = dv.values.some((v) => String(v).trim().toLowerCase() === norm);
        if (!hit && dv.mode === "stop") out.errors.push({ col, reason: "dv-stop" });
        else if (!hit && dv.mode === "warning") out.warnings.push({ col, reason: "dv-warning" });
      }
      out.changes.push({ col, raw: text, before, after, parsedType: parsed ? parsed.type : "empty", overwrite: isOverwrite });
    });
    out.ok = out.errors.length === 0;
    lastPlan = { key: keyOf(input.target), before: Object.fromEntries(out.changes.map((c) => [c.col, c.before])) };
    return out;
  }

  const undoTop = () => (typeof undoStack !== "undefined" && undoStack.length ? undoStack[undoStack.length - 1] : null);
  const nextTask = () => new Promise((r) => setTimeout(r, 0));

  // Zatwierdzenie kartki. Asynchroniczne tylko po to, żeby policzyć kroki Cofnij — krok
  // domyka się mikro-zadaniem po zapisie (undo.js); same zapisy idą w JEDNYM zadaniu.
  async function commit(input) {
    if (!ready()) return { ok: false, reason: "guard" };
    if (typeof swbIsEditMode === "function" && !swbIsEditMode()) return { ok: false, reason: "guard" };
    const isNew = !!(input && input.target && input.target.newRow);
    if (isNew && (typeof swbRowGrow === "undefined" || swbRowGrow.filtered())) return { ok: false, reason: "filtered" };
    const expect = input && input.expect ? input.expect : (lastPlan && lastPlan.key === keyOf(input && input.target) ? lastPlan.before : null);
    const p = plan(input);
    if (p.errors.some((e) => e.reason === "no-target")) return { ok: false, reason: "stale", plan: p };
    if (!isNew && expect) {
      const row = rowAt(input.target.rowIndex0);
      const moved = Object.keys(expect).some((col) => String(getDisplayValue(row, Number(col)) ?? "") !== expect[col]);
      if (moved) return { ok: false, reason: "stale", plan: p };
    }
    if (!p.ok) return { ok: false, reason: "invalid", plan: p };
    if (!p.changes.length) return { ok: false, reason: "empty", plan: p };

    const topBefore = undoTop();
    const depthBefore = typeof undoStack !== "undefined" ? undoStack.length : 0;
    let row;
    if (isNew) {
      const key = swbRowGrow.addDraft({ col: p.changes[0].col });
      row = key ? getRowByKey(key) : null;
      if (!row) return { ok: false, reason: "filtered", plan: p };
    } else {
      row = rowAt(input.target.rowIndex0);
    }
    const failed = [];
    p.changes.forEach((c) => {
      const res = writeCellFromInput(currentDisplayModel, row, c.col, c.raw);
      if (res !== "ok") failed.push({ col: c.col, reason: res });
      else if (typeof rememberColumnValue === "function") {
        const v = row.values[c.col];
        rememberColumnValue(c.col, v instanceof Date && typeof formatDateForEdit === "function" ? formatDateForEdit(v) : v);
      }
    });
    if (typeof setDirtyState === "function") setDirtyState(true);
    if (typeof renderActiveTable === "function") renderActiveTable();
    await nextTask();
    const top = undoTop();
    let steps = 0;
    if (typeof undoStack !== "undefined" && top && top !== topBefore) {
      steps = Math.max(1, Math.min(2, undoStack.length - depthBefore));
      if (undoStack.length - depthBefore < 1) steps = top.auto ? 2 : 1; // stos ucięty limitem
    }
    lastCard = steps ? { top, steps } : null;
    lastPlan = null;
    if (failed.length) return { ok: false, reason: "write-failed", failed, rowIndex0: row.rowIndex0, undoSteps: steps };
    return { ok: true, rowIndex0: row.rowIndex0, undoSteps: steps };
  }

  // Cofnij ostatnią kartkę — tylko gdy na wierzchu historii nadal leży jej krok.
  function undoCard() {
    if (!lastCard || undoTop() !== lastCard.top) return false;
    for (let i = 0; i < lastCard.steps; i++) undoLast();
    lastCard = null;
    return true;
  }

  // Wejście programowe (ręczna karta i w przyszłości skan): ustawia cel i pola, zwraca podgląd.
  function fill(input) {
    filled = { target: input && input.target ? { ...input.target } : null, fields: { ...((input && input.fields) || {}) } };
    return filled.target ? plan({ target: filled.target, fields: filled.fields }) : null;
  }

  return { candidates, plan, commit, undoCard, fill, _filled: () => filled };
})();

window.__entry = swbEntry;
