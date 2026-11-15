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
  // Z pamięci, dopóki dane się nie zmienią (sheetDataStamp rośnie przy każdym zapisie komórki,
  // baseRows podmienia się przy przebudowie) — pole celu woła to przy każdym klawiszu.
  let memo = null; // { rows, base, len, stamp, sheet, byIndex, foldCache }
  function memoNow() {
    if (!ready() || !Array.isArray(baseRows)) return null;
    const stamp = typeof sheetDataStamp === "number" ? sheetDataStamp : 0;
    if (memo && memo.base === baseRows && memo.len === baseRows.length && memo.stamp === stamp && memo.sheet === currentSheetName) return memo;
    const rows = baseRows
      .filter((r) => r && !r.isLongViewRow && Number.isFinite(r.rowIndex0) && r.rowIndex0 >= currentHeaderRow
        && Array.isArray(r.values) && r.values.some((v) => !isBlank(v)))
      .sort((a, b) => a.rowIndex0 - b.rowIndex0);
    memo = { rows, base: baseRows, len: baseRows.length, stamp, sheet: currentSheetName, byIndex: new Map(rows.map((r) => [r.rowIndex0, r])), foldCache: new Map() };
    return memo;
  }
  const candidates = () => { const m = memoNow(); return m ? m.rows : []; };
  const rowAt = (rowIndex0) => { const m = memoNow(); return (m && m.byIndex.get(rowIndex0)) || null; };
  const foldCache = () => { const m = memoNow(); return m ? m.foldCache : null; };
  // Wiersze do szukania celu: klucz jako TEKST Z KOMÓRKI (to, co widać i co jest na kartce) —
  // liczba 32 w formacie 000 to dla człowieka „032”; surowa wartość dawała „32” i brak trafienia.
  const searchRows = (keyCol) => {
    const m = memoNow();
    if (!m) return [];
    if (!m.search) m.search = new Map();
    if (!m.search.has(keyCol)) {
      m.search.set(keyCol, m.rows.map((r) => {
        const v = []; // tablica — silnik czyta values[keyCol] tylko z tablicy
        v[keyCol] = String(getDisplayValue(r, keyCol) ?? "");
        return { rowIndex0: r.rowIndex0, values: v };
      }));
    }
    return m.search.get(keyCol);
  };
  const lastDataRow = () => { const c = candidates(); return c.length ? c[c.length - 1] : null; };

  const sameValue = (a, b) => {
    if (a instanceof Date && b instanceof Date) return a.getTime() === b.getTime();
    if (isBlank(a) && isBlank(b)) return true;
    if (typeof a === "string" && typeof b === "string") return a.trim() === b.trim();
    return a === b;
  };

  // Tekst, który POKAŻE komórka po zapisie — w jej formacie liczbowym (jak cfmtDisplayFor),
  // a nie surowa wartość: „12.03” w kolumnie dat to liczba, którą komórka pokaże jako 12.01.1900.
  function previewText(parsed, refIdx, col, fallback) {
    if (!parsed || parsed.value == null) return fallback;
    const v = parsed.value;
    const fileFmt = typeof cfmtNumFmtAt === "function" ? cfmtNumFmtAt(refIdx, col) : "General";
    const code = parsed.numFmt && fileFmt === "General" ? parsed.numFmt : fileFmt;
    if (code && code !== "General" && !(v instanceof Date) && typeof cfmtFormatValue === "function") {
      const out = cfmtFormatValue(v, code);
      if (out != null) return out;
    }
    return typeof swbGeneralText === "function" ? swbGeneralText(v) : toDisplay(v);
  }

  const noteAt = (rowIndex0, col) => (typeof swbNoteOf === "function"
    ? swbNoteOf(currentSheetName, XLSX.utils.encode_cell({ r: rowIndex0, c: currentStartCol + col })) : null);
  // „było” każdej zmiany: wartość pod kluczem kolumny, notatka pod „n<kolumna>” (wykrycie „stale”)
  const expectOf = (p) => Object.fromEntries(p.changes.map((c) => [c.kind === "note" ? `n${c.col}` : String(c.col), c.before]));

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
      const after = parsed && parsed.type === "formula" ? text : previewText(parsed, refIdx, col, text);
      if (isOverwrite && !overwrite.has(col)) out.errors.push({ col, reason: "overwrite-not-allowed" });
      let dv = null;
      try { dv = typeof getCellDataValidation === "function" ? getCellDataValidation(currentSheetName, refIdx, currentStartCol + col) : null; } catch { dv = null; }
      if (dv && !(parsed && parsed.type === "formula")) {
        const norm = String(parsed ? parsed.value : text).trim().toLowerCase();
        const hit = dv.values.some((v) => String(v).trim().toLowerCase() === norm);
        if (!hit && dv.mode === "stop") out.errors.push({ col, reason: "dv-stop" });
        else if (!hit && dv.mode === "warning") out.warnings.push({ col, reason: "dv-warning" });
      }
      out.changes.push({ col, kind: "value", raw: text, before, after, parsedType: parsed ? parsed.type : "empty", overwrite: isOverwrite });
    });
    // Notatki komórek (cell-notes.js): te same zasady co wartości — puste pole nic nie usuwa,
    // zmiana istniejącej notatki tylko ze zgodą; wolno też w kolumnie z formułą (jak w Excelu).
    const overwriteNotes = new Set((input.overwriteNotes || []).map(Number));
    Object.keys(input.notes || {}).map(Number).sort((a, b) => a - b).forEach((col) => {
      const raw = input.notes[col];
      if (!Number.isInteger(col) || col < 0 || col >= n) return;
      if (isBlank(raw)) return;
      const text = String(raw).replace(/\r\n?/g, "\n").replace(/\s+$/, "");
      if (typeof cfmtIsRealCol === "function" && !cfmtIsRealCol(col)) { out.errors.push({ col, reason: "not-real-col", kind: "note" }); return; }
      const cur = row ? noteAt(row.rowIndex0, col) : null;
      if (cur && cur.threaded) { out.errors.push({ col, reason: "note-threaded", kind: "note" }); return; }
      if (cur && cur.text === text) return;
      if (cur && !overwriteNotes.has(col)) out.errors.push({ col, reason: "note-overwrite-not-allowed", kind: "note" });
      out.changes.push({ col, kind: "note", raw: text, before: cur ? cur.text : "", after: text, parsedType: "note", overwrite: !!cur });
    });
    // nowy wiersz z samymi notatkami nie jest wierszem danych — tabela by nie urosła, a notatka wisiałaby pod nią
    if (isNew && out.changes.length && !out.changes.some((c) => c.kind === "value")) out.errors.push({ col: -1, reason: "new-row-needs-value" });
    out.ok = out.errors.length === 0;
    lastPlan = { key: keyOf(input.target), before: expectOf(out) };
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
      const moved = Object.keys(expect).some((k) => {
        if (k[0] === "n") { const cur = noteAt(row.rowIndex0, Number(k.slice(1))); return (cur ? cur.text : "") !== expect[k]; }
        return String(getDisplayValue(row, Number(k)) ?? "") !== expect[k];
      });
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
    p.changes.filter((c) => c.kind === "value").forEach((c) => {
      const res = writeCellFromInput(currentDisplayModel, row, c.col, c.raw);
      if (res !== "ok") failed.push({ col: c.col, reason: res });
      else if (typeof rememberColumnValue === "function") {
        const v = row.values[c.col];
        rememberColumnValue(c.col, v instanceof Date && typeof formatDateForEdit === "function" ? formatDateForEdit(v) : v);
      }
    });
    // notatki po wartościach (nowy wiersz już istnieje), w tym samym zadaniu = ten sam krok Cofnij
    p.changes.filter((c) => c.kind === "note").forEach((c) => {
      const okNote = typeof swbSetNote === "function" && swbSetNote(row.rowIndex0, c.col, c.raw, typeof swbNoteAuthor === "function" ? swbNoteAuthor() : "");
      if (!okNote) failed.push({ col: c.col, reason: "note" });
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

  const api = { candidates, plan, commit, undoCard, fill, sheetCell, rowAt, lastDataRow, ready, previewText, foldCache, expectOf, searchRows, _filled: () => filled };
  return api;
})();

// ── Nakładka: cel → pola → podgląd ─────────────────────────────────────────────
// Stan karty żyje tu (card), DOM rysujemy od nowa przy zmianie kroku; w kroku 2 pola
// aktualizują tylko swoje podpowiedzi (bez przebudowy — fokus i klawiatura zostają).
(() => {
  const C = typeof SwbEntryCore !== "undefined" ? SwbEntryCore : null;
  const overlay = document.getElementById("entryOverlay");
  const btn = document.getElementById("entryBtn");
  if (!overlay || !C) return;

  const DEFAULT_OPTIONS = { fuzzyTarget: true, typoHints: true, typeHints: true, suggestions: true, duplicateGuard: true, confirmStep: true, queryOperators: true };
  const ui = {
    open: false, step: 1, keyCol: 0, query: "", targets: [], active: -1,
    target: null, fields: {}, overwrite: new Set(), pinned: new Set(), font: 2,
    options: { ...DEFAULT_OPTIONS }, profiles: new Map(), lastPlan: null, sessionLog: [], returnFocus: null,
    fieldOrder: [], selected: new Set(), panel: null, draft: null, reminded: false,
    notes: {}, overwriteNotes: new Set(), queryMode: false, queryTotal: 0,
  };

  // ── Pamięć (localStorage, tylko na tym urządzeniu) ──
  // { options, scopes: { "plik\tarkusz": { keyCol, fieldOrder, selected, pinned: {kol: wartość}, draft, at } } }
  // Opcje są wspólne dla wszystkich plików; reszta per arkusz. Brak localStorage = działa bez pamięci.
  const STORE_KEY = "excel-workbench-entry";
  const MAX_SCOPES = 30;
  const loadStore = () => { try { const x = JSON.parse(localStorage.getItem(STORE_KEY) || "{}"); return x && typeof x === "object" ? x : {}; } catch { return {}; } };
  function saveStore(store) {
    try {
      const scopes = store.scopes || {};
      const keys = Object.keys(scopes);
      if (keys.length > MAX_SCOPES) keys.sort((a, b) => (scopes[a]?.at || 0) - (scopes[b]?.at || 0)).slice(0, keys.length - MAX_SCOPES).forEach((k) => delete scopes[k]);
      localStorage.setItem(STORE_KEY, JSON.stringify(store));
    } catch { /* brak miejsca / tryb prywatny — karta działa dalej, bez pamięci */ }
  }
  const scopeKey = () => `${typeof currentFileName !== "undefined" && currentFileName ? currentFileName : ""}\t${currentSheetName || ""}`;
  function cleanFields(obj) {
    const out = {};
    Object.keys(obj || {}).forEach((k) => { const v = obj[k]; if (v != null && String(v) !== "") out[k] = String(v); });
    return out;
  }
  function persist({ draft = true } = {}) {
    if (!currentSheetName) return;
    const store = loadStore();
    store.options = { ...ui.options };
    store.scopes = store.scopes || {};
    const pinned = {};
    ui.pinned.forEach((c) => { if (ui.fields[c] != null && ui.fields[c] !== "") pinned[c] = String(ui.fields[c]); else pinned[c] = ""; });
    const rec = { keyCol: ui.keyCol, fieldOrder: ui.fieldOrder.slice(), selected: [...ui.selected], pinned, at: Date.now() };
    const own = Object.keys(cleanFields(ui.fields)).filter((c) => !ui.pinned.has(Number(c))).length + Object.keys(cleanFields(ui.notes)).length;
    rec.draft = draft && ui.target && own ? { target: { ...ui.target }, fields: cleanFields(ui.fields), notes: cleanFields(ui.notes), query: ui.query, savedAt: Date.now() } : null;
    store.scopes[scopeKey()] = rec;
    saveStore(store);
  }
  let persistTimer = null;
  const persistSoon = () => { clearTimeout(persistTimer); persistTimer = setTimeout(() => persist(), 300); };
  const persistNow = () => { clearTimeout(persistTimer); persistTimer = null; persist(); };
  function restore() {
    const store = loadStore();
    ui.options = { ...DEFAULT_OPTIONS, ...(store.options || {}) };
    const rec = (store.scopes || {})[scopeKey()] || null;
    const n = headers().length;
    const valid = (c) => Number.isInteger(c) && c >= 0 && c < n;
    const cols = realCols();
    const order = rec && Array.isArray(rec.fieldOrder) ? rec.fieldOrder.filter((c) => valid(c) && cols.includes(c)) : [];
    cols.forEach((c) => { if (!order.includes(c)) order.push(c); });
    ui.fieldOrder = order;
    ui.selected = new Set(rec && Array.isArray(rec.selected) ? rec.selected.filter(valid) : cols);
    cols.forEach((c) => { if (rec && Array.isArray(rec.fieldOrder) && !rec.fieldOrder.includes(c)) ui.selected.add(c); }); // nowa kolumna = na karcie
    ui.keyColSaved = rec && valid(rec.keyCol) ? rec.keyCol : null;
    ui.pinned = new Set();
    const pinnedVals = {};
    if (rec && rec.pinned) Object.keys(rec.pinned).forEach((k) => { const c = Number(k); if (valid(c)) { ui.pinned.add(c); if (rec.pinned[k]) pinnedVals[c] = rec.pinned[k]; } });
    if (!ui.target) ui.fields = { ...pinnedVals };
    ui.draft = rec && rec.draft && rec.draft.target ? rec.draft : null;
    if (ui.draft && !ui.draft.target.newRow && !swbEntry.rowAt(ui.draft.target.rowIndex0)) ui.draft = null;
  }
  const suggests = [];
  let wakeLock = null;

  const esc = (s) => String(s ?? "").replace(/[&<>"']/g, (ch) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[ch]));
  const headers = () => currentDisplayModel.headers;
  const realCols = () => headers().map((_, i) => i).filter((i) => typeof cfmtIsRealCol !== "function" || cfmtIsRealCol(i));
  const columnValues = (col) => swbEntry.candidates().map((r) => r.values[col]);
  const profile = (col) => {
    if (!ui.profiles.has(col)) ui.profiles.set(col, C.profileColumn(columnValues(col)));
    return ui.profiles.get(col);
  };
  const isFormulaCol = (col) => {
    const ref = ui.target && !ui.target.newRow ? swbEntry.rowAt(ui.target.rowIndex0) : swbEntry.lastDataRow();
    const cell = ref ? swbEntry.sheetCell(ref.rowIndex0, col) : null;
    return !!(cell && cell.f);
  };
  const excelRow = (rowIndex0) => rowIndex0 + 1;
  const keyText = (row) => (row ? String(getDisplayValue(row, ui.keyCol) ?? "") : "");

  const ICON_PIN = '<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><line x1="12" y1="17" x2="12" y2="22"/><path d="M5 17h14v-1.76a2 2 0 0 0-1.11-1.79l-1.78-.9A2 2 0 0 1 15 10.76V6h1a2 2 0 0 0 0-4H8a2 2 0 0 0 0 4h1v4.76a2 2 0 0 1-1.11 1.79l-1.78.9A2 2 0 0 0 5 15.24Z"/></svg>';
  const ICON_NOTE = '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M21 15a2 2 0 0 1-2 2H7l-4 4V5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2z"/></svg>';
  const ICON_X = '<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>';

  function destroySuggests() { while (suggests.length) { try { suggests.pop().destroy(); } catch { /* już zamknięte */ } } }

  function shell() {
    const src = `${esc(typeof currentFileName !== "undefined" && currentFileName ? currentFileName : "")} · ${esc(currentSheetName)}`;
    overlay.innerHTML = `
      <div class="tr-head">
        <div class="tr-head-main">
          <div class="tr-title" id="entryTitle">${esc(t("entryTitle"))}</div>
          <div class="tr-source">${src}</div>
        </div>
        <div class="tr-head-tools">
          <button class="btn btn-sm ghost" id="enLogBtn" type="button" aria-expanded="false" aria-controls="enLog">${esc(t("entryLog", { n: ui.sessionLog.length }))}</button>
          <button class="btn btn-sm ghost" id="enSettingsBtn" type="button" aria-expanded="false" aria-controls="enSettings">${esc(t("entrySettings"))}</button>
          <button class="btn btn-sm ghost" id="enFontBtn" type="button" aria-label="${esc(t("entryFont"))}">A</button>
          <button class="modal-close" id="enCloseBtn" type="button" aria-label="${esc(t("entryClose"))}">${ICON_X}</button>
        </div>
      </div>
      <div class="en-live visually-hidden" id="enLive" aria-live="polite"></div>
      <div class="tr-stage-wrap en-stage-wrap"><div class="en-panel-slot" id="enPanelSlot"></div><div class="tr-stage en-stage" id="enStage"></div></div>
      <div class="tr-bar en-bar" id="enBar"></div>`;
    overlay.querySelector("#enCloseBtn").addEventListener("click", close);
    overlay.querySelector("#enSettingsBtn").addEventListener("click", () => togglePanel("settings"));
    overlay.querySelector("#enLogBtn").addEventListener("click", () => togglePanel("log"));
    overlay.querySelector("#enFontBtn").addEventListener("click", () => {
      ui.font = ui.font % 3 + 1;
      overlay.dataset.font = String(ui.font);
    });
  }
  function togglePanel(which) {
    ui.panel = ui.panel === which ? null : which;
    drawPanel();
    if (!ui.panel) render();
  }

  function drawPanel() {
    const slot = overlay.querySelector("#enPanelSlot");
    if (!slot) return;
    overlay.querySelector("#enSettingsBtn")?.setAttribute("aria-expanded", String(ui.panel === "settings"));
    overlay.querySelector("#enLogBtn")?.setAttribute("aria-expanded", String(ui.panel === "log"));
    const logBtn = overlay.querySelector("#enLogBtn");
    if (logBtn) logBtn.textContent = t("entryLog", { n: ui.sessionLog.length });
    if (ui.panel === "settings") {
      const opts = [["fuzzyTarget", "entryOptFuzzy"], ["typoHints", "entryOptTypo"], ["typeHints", "entryOptType"], ["suggestions", "entryOptSuggest"], ["duplicateGuard", "entryOptDup"], ["confirmStep", "entryOptConfirm"], ["queryOperators", "entryOptQuery"]];
      const textCols = realCols();
      slot.innerHTML = `<section class="tr-card en-card en-panel" id="enSettings" aria-label="${esc(t("entrySettings"))}">
        <label class="en-label" for="enKeyCol">${esc(t("entryKeyColLabel"))}</label>
        <select class="en-select" id="enKeyCol">${textCols.map((c) => `<option value="${c}" ${c === ui.keyCol ? "selected" : ""}>${esc(headers()[c])}</option>`).join("")}</select>
        <h2 class="en-panel-title">${esc(t("entryOptionsTitle"))}</h2>
        ${opts.map(([k, label]) => `<label class="en-check"><input type="checkbox" data-en-opt="${k}" ${ui.options[k] ? "checked" : ""}> ${esc(t(label))}</label>`).join("")}
        <p class="en-hint">${esc(t("entryOptionsNote"))}</p>
        <h2 class="en-panel-title">${esc(t("entryFieldsTitle"))}</h2>
        <ul class="en-fieldlist">${ui.fieldOrder.map((c, i) => `<li class="en-fieldrow" data-en-fieldrow="${c}">
          <label class="en-check"><input type="checkbox" data-en-field-on ${ui.selected.has(c) ? "checked" : ""}> ${esc(headers()[c])}</label>
          <span class="en-move">
            <button class="btn btn-xs ghost" type="button" data-en-move="up" ${i === 0 ? "disabled" : ""} aria-label="${esc(t("entryMoveUp", { name: headers()[c] }))}">↑</button>
            <button class="btn btn-xs ghost" type="button" data-en-move="down" ${i === ui.fieldOrder.length - 1 ? "disabled" : ""} aria-label="${esc(t("entryMoveDown", { name: headers()[c] }))}">↓</button>
          </span></li>`).join("")}</ul>
      </section>`;
      slot.querySelector("#enKeyCol").addEventListener("change", (e) => { ui.keyCol = Number(e.target.value); ui.keyColSaved = ui.keyCol; persistNow(); });
      slot.querySelectorAll("[data-en-opt]").forEach((cb) => cb.addEventListener("change", () => { ui.options[cb.dataset.enOpt] = cb.checked; persistNow(); }));
      slot.querySelectorAll("[data-en-fieldrow]").forEach((li) => {
        const c = Number(li.dataset.enFieldrow);
        li.querySelector("[data-en-field-on]").addEventListener("change", (e) => { if (e.target.checked) ui.selected.add(c); else ui.selected.delete(c); persistNow(); });
        li.querySelectorAll("[data-en-move]").forEach((b) => b.addEventListener("click", () => {
          const i = ui.fieldOrder.indexOf(c);
          const j = b.dataset.enMove === "up" ? i - 1 : i + 1;
          if (j < 0 || j >= ui.fieldOrder.length) return;
          [ui.fieldOrder[i], ui.fieldOrder[j]] = [ui.fieldOrder[j], ui.fieldOrder[i]];
          persistNow();
          drawPanel();
          slot.querySelector(`[data-en-fieldrow="${c}"] [data-en-move="${b.dataset.enMove}"]`)?.focus();
        }));
      });
    } else if (ui.panel === "log") {
      const fmt = (ms) => new Date(ms).toLocaleTimeString(currentLang === "en" ? "en-GB" : "pl-PL", { hour: "2-digit", minute: "2-digit" });
      slot.innerHTML = `<section class="tr-card en-card en-panel" id="enLog" aria-label="${esc(t("entryLogTitle"))}">
        <h2 class="en-panel-title">${esc(t("entryLogTitle"))}</h2>
        ${ui.sessionLog.length ? `<ul class="en-fieldlist">${ui.sessionLog.map((x) => `<li class="en-fieldrow">
          <span class="en-log-key">${esc(x.key)} · ${esc(t("entryRowN", { n: x.rowIndex0 + 1 }))} · ${esc(fmt(x.at))}</span>
          <button class="btn btn-xs ghost" type="button" data-en-log-row="${x.rowIndex0}">${esc(t("entryDoneShow"))}</button></li>`).join("")}</ul>` : `<p class="en-hint">${esc(t("entryLogEmpty"))}</p>`}
      </section>`;
      slot.querySelectorAll("[data-en-log-row]").forEach((b) => b.addEventListener("click", () => showInTable(Number(b.dataset.enLogRow))));
    } else {
      slot.innerHTML = "";
    }
  }

  function showInTable(rowIndex0) {
    close();
    setFocusedCell(`wide:${rowIndex0}`, ui.keyCol, { scroll: true });
  }

  const live = (msg) => { const el = overlay.querySelector("#enLive"); if (el) el.textContent = msg; };

  function render() {
    destroySuggests();
    const stage = overlay.querySelector("#enStage");
    const bar = overlay.querySelector("#enBar");
    if (!stage || !bar) return;
    if (ui.step === 1) renderTarget(stage, bar);
    else if (ui.step === 2) renderFields(stage, bar);
    else renderReview(stage, bar);
  }

  // ── Krok 1: cel ──
  function renderTarget(stage, bar) {
    stage.innerHTML = `
      <div class="tr-card en-card">
        <label class="en-label" for="enTargetInput">${esc(t("entryTargetLabel"))}</label>
        <input class="en-input en-target-input" id="enTargetInput" type="text" autocomplete="off" spellcheck="false"
          enterkeyhint="go" role="combobox" aria-expanded="false" aria-controls="enCandidates" aria-autocomplete="list"
          placeholder="${esc(t("entryTargetPlaceholder"))}" value="${esc(ui.query)}">
        <div class="en-hint" id="enKeyHint">${esc(t("entryTargetHint", { col: headers()[ui.keyCol] || "" }))}</div>
        <div class="en-candidates" id="enCandidates" role="listbox" aria-label="${esc(t("entryTargetLabel"))}"></div>
      </div>`;
    bar.innerHTML = `<span class="en-keys">${esc(t("entryKeysHint"))}</span>`;
    if (ui.draft) {
      const d = ui.draft;
      const row = d.target.newRow ? null : swbEntry.rowAt(d.target.rowIndex0);
      const own = Object.keys(d.fields || {}).filter((c) => !ui.pinned.has(Number(c))).length + Object.keys(d.notes || {}).length;
      const note = document.createElement("div");
      note.className = "tr-notice en-draft";
      note.id = "enDraft";
      note.setAttribute("role", "status");
      note.innerHTML = `<div class="tr-notice-text">${esc(t("entryDraft", { key: row ? keyText(row) : (d.query || t("entryDraftNew")), n: own }))}</div>
        <div class="tr-notice-actions"><button class="btn btn-xs" id="enDraftResume" type="button">${esc(t("entryDraftResume"))}</button>
        <button class="btn btn-xs ghost" id="enDraftDiscard" type="button">${esc(t("entryDraftDiscard"))}</button></div>`;
      stage.prepend(note);
      note.querySelector("#enDraftResume").addEventListener("click", () => {
        const dd = ui.draft;
        ui.draft = null;
        ui.query = dd.query || "";
        ui.target = { ...dd.target };
        ui.fields = { ...pinnedFields(), ...dd.fields };
        ui.notes = { ...(dd.notes || {}) };
        ui.overwrite = new Set();
        ui.overwriteNotes = new Set();
        ui.step = 2;
        render();
      });
      note.querySelector("#enDraftDiscard").addEventListener("click", () => {
        ui.draft = null;
        persist({ draft: false });
        note.remove();
        overlay.querySelector("#enTargetInput")?.focus();
      });
    }
    const input = stage.querySelector("#enTargetInput");
    input.addEventListener("input", () => { ui.query = input.value; search(); });
    input.addEventListener("keydown", onTargetKey);
    search();
    input.focus();
  }

  function search() {
    const rows = swbEntry.candidates();
    const q = C.trimQuery(ui.query);
    ui.queryMode = !!ui.options.queryOperators && C.looksLikeQuery(q, headers());
    ui.queryCrit = null;
    if (ui.queryMode) {
      // ta sama składnia i ten sam silnik co szybkie szukanie (workbook.js) — po wszystkich kolumnach pliku
      const criterion = { query: q, mode: "contains", indexes: realCols(), headers: headers(), operatorsEnabled: true, negated: false, emptyMode: "all" };
      ui.queryCrit = criterion;
      let hits = [];
      try { hits = rows.filter((r) => rowMatchesTextFilter(r, [criterion], false)); } catch { hits = []; }
      ui.queryTotal = hits.length;
      ui.targets = hits.slice(0, 6).map((r) => ({ rowIndex0: r.rowIndex0, score: 3, exact: false }));
      ui.active = hits.length === 1 ? 0 : -1; // zapytanie jest precyzyjne — jedno trafienie = wybrane
    } else {
      ui.queryTotal = 0;
      ui.targets = C.findTargets(swbEntry.searchRows(ui.keyCol), ui.keyCol, q, { fuzzy: ui.options.fuzzyTarget, foldCache: swbEntry.foldCache() });
      // jeden dokładny kandydat = podświetlony (Enter go wybiera); kilka albo literówki — wybiera człowiek
      const exact = ui.targets.filter((x) => x.exact);
      ui.active = exact.length === 1 && ui.targets[0] && ui.targets[0].exact ? 0 : -1;
    }
    drawCandidates();
  }

  function drawCandidates() {
    const box = overlay.querySelector("#enCandidates");
    const input = overlay.querySelector("#enTargetInput");
    if (!box || !input) return;
    const rows = swbEntry.candidates();
    const cols = C.distinguishingCols(headers(), headers().map((_, i) => rows.map((r) => r.values[i])), ui.keyCol, 2);
    const q = C.trimQuery(ui.query);
    const items = ui.targets.map((tg, i) => {
      const row = swbEntry.rowAt(tg.rowIndex0);
      const extra = cols.map((c) => `${esc(headers()[c])}: ${esc(getDisplayValue(row, c) || "—")}`).join(" · ");
      // przy zapytaniu: DLACZEGO ten wiersz pasuje (np. „032” w innej kolumnie niż numer)
      let why = "";
      if (ui.queryCrit && typeof collectMatchingCellsForRow === "function") {
        let hitCols = [];
        try { hitCols = [...collectMatchingCellsForRow(row, [ui.queryCrit], null)].filter((c) => c !== ui.keyCol).sort((a, b) => a - b).slice(0, 2); } catch { hitCols = []; }
        if (hitCols.length) why = `<span class="en-cand-why">${esc(t("entryWhy", { where: hitCols.map((c) => `${headers()[c]}: ${getDisplayValue(row, c)}`).join(" · ") }))}</span>`;
      }
      return `<div class="en-cand" role="option" id="enCand${i}" data-en-row="${tg.rowIndex0}" aria-selected="${i === ui.active}">
        <span class="en-cand-key">${esc(keyText(row))}</span>
        <span class="en-cand-extra">${extra}</span>${why}
        <span class="en-cand-row">${esc(t("entryRowN", { n: excelRow(tg.rowIndex0) }))}</span></div>`;
    });
    const newIdx = ui.targets.length;
    const newLabel = ui.queryMode ? t("entryNewRow") : t("entryNewRowFor", { text: q });
    if (q) items.push(`<div class="en-cand en-cand-new" role="option" id="enCand${newIdx}" data-en-new aria-selected="${ui.active === newIdx}">${esc(newLabel)}</div>`);
    const keyHint = overlay.querySelector("#enKeyHint");
    if (keyHint) keyHint.hidden = ui.queryMode; // zapytanie szuka we wszystkich kolumnach, nie tylko w kluczu
    box.innerHTML = (ui.queryMode ? `<div class="en-hint en-query-hint" id="enQueryHint">${esc(t("entryQueryHint", { n: ui.queryTotal }))}</div>` : "") + items.join("");
    input.setAttribute("aria-expanded", items.length ? "true" : "false");
    if (ui.active >= 0) input.setAttribute("aria-activedescendant", `enCand${ui.active}`); else input.removeAttribute("aria-activedescendant");
    box.querySelectorAll(".en-cand").forEach((el, i) => {
      // indeks = pozycja wśród opcji (podpowiedź o zapytaniu nie jest opcją)
      el.addEventListener("pointerdown", (e) => e.preventDefault()); // fokus zostaje w polu
      el.addEventListener("click", () => { ui.active = i; choose(); });
    });
    if (q && !ui.targets.length) live(t(ui.queryMode ? "entryQueryNone" : "entryNoMatch"));
    else if (ui.targets.length > 1 && ui.active < 0) live(t("entryPickHint"));
  }

  function onTargetKey(e) {
    const total = ui.targets.length + (ui.query.trim() ? 1 : 0);
    if (e.key === "ArrowDown" || e.key === "ArrowUp") {
      e.preventDefault();
      if (!total) return;
      const d = e.key === "ArrowDown" ? 1 : -1;
      ui.active = ui.active < 0 ? (d > 0 ? 0 : total - 1) : (ui.active + d + total) % total;
      drawCandidates();
    } else if (e.key === "Enter" && !e.ctrlKey && !e.metaKey) {
      e.preventDefault();
      if (ui.active >= 0) choose();
      else if (ui.query.trim() && !ui.targets.length && !ui.queryMode) { ui.active = 0; choose(); }
      else if (ui.queryMode && !ui.targets.length) live(t("entryQueryNone"));
      else if (ui.targets.length) live(t("entryPickHint"));
    }
  }

  async function choose() {
    const isNew = ui.active === ui.targets.length;
    if (!isNew) {
      const tg = ui.targets[ui.active];
      if (!tg) return;
      startCard({ rowIndex0: tg.rowIndex0 });
      return;
    }
    if (!ui.queryMode && ui.options.duplicateGuard && C.isDuplicateRisk(ui.targets)) {
      const best = ui.targets[0];
      const row = swbEntry.rowAt(best.rowIndex0);
      const n = excelRow(best.rowIndex0);
      const ans = await appDialog({
        title: t("entryDupTitle"),
        message: t("entryDupMsg", { key: keyText(row), n }),
        actions: [{ label: t("entryDupFill", { n }), value: "fill", kind: "primary" }, { label: t("entryDupNew"), value: "new", kind: "secondary" }],
      });
      if (ans === "fill") { startCard({ rowIndex0: best.rowIndex0 }); return; }
      if (ans !== "new") { overlay.querySelector("#enTargetInput")?.focus(); return; }
    }
    startCard({ newRow: true });
  }

  function pinnedFields() {
    const out = {};
    ui.pinned.forEach((c) => { if (ui.fields[c] != null && ui.fields[c] !== "") out[c] = ui.fields[c]; });
    return out;
  }

  function startCard(target) {
    const keep = pinnedFields();
    ui.target = target;
    ui.fields = { ...keep };
    if (target.newRow && C.trimQuery(ui.query) && !ui.queryMode) ui.fields[ui.keyCol] = C.trimQuery(ui.query);
    ui.overwrite = new Set();
    ui.notes = {};
    ui.overwriteNotes = new Set();
    ui.step = 2;
    ui.draft = null;
    render();
    persistNow();
  }

  // ── Krok 2: pola ──
  function renderFields(stage, bar) {
    const row = ui.target.newRow ? null : swbEntry.rowAt(ui.target.rowIndex0);
    const head = ui.target.newRow ? t("entryNewRow") : t("entryTargetExisting", { n: excelRow(ui.target.rowIndex0), key: keyText(row) });
    const cardCols = ui.fieldOrder.filter((c) => ui.selected.has(c));
    const fieldsHtml = cardCols.map((col) => {
      const formula = isFormulaCol(col);
      const now = row ? String(getDisplayValue(row, col) ?? "") : "";
      const val = formula ? now : (ui.fields[col] ?? "");
      const pinned = ui.pinned.has(col);
      const curNote = row && typeof swbNoteOf === "function" ? swbNoteOf(currentSheetName, XLSX.utils.encode_cell({ r: row.rowIndex0, c: currentStartCol + col })) : null;
      const noteVal = ui.notes[col] ?? (curNote ? curNote.text : "");
      const noteOpen = ui.notes[col] != null && ui.notes[col] !== "";
      const noteBtnLabel = curNote ? t("entryNoteEdit", { name: headers()[col] }) : t("entryNoteAdd", { name: headers()[col] });
      const noteHtml = `<button class="btn btn-xs ghost en-note-btn${curNote ? " has-note" : ""}" type="button" aria-expanded="${noteOpen}" aria-controls="enN${col}" aria-label="${esc(noteBtnLabel)}">${ICON_NOTE}<span>${esc(t("entryNote"))}</span></button>`;
      const noteArea = `<textarea class="en-input en-note" id="enN${col}" rows="2" ${noteOpen ? "" : "hidden"} ${curNote && curNote.threaded ? "readonly" : ""}
          aria-label="${esc(noteBtnLabel)}" placeholder="${esc(t("entryNotePlaceholder"))}">${esc(noteVal)}</textarea>`;
      return `<div class="tr-field en-field${formula ? " is-formula" : ""}" data-en-col="${col}">
        <div class="en-field-head">
          <label class="tr-field-label" for="enF${col}">${esc(headers()[col])}</label>
          <span class="en-field-tools">${noteHtml}
          ${formula ? `<span class="en-tag">${esc(t("entryFormulaField"))}</span>` : `<button class="en-pin" type="button" aria-pressed="${pinned}" aria-label="${esc(pinned ? t("entryUnpin") : t("entryPin"))}" title="${esc(pinned ? t("entryUnpin") : t("entryPin"))}">${ICON_PIN}</button>`}</span>
        </div>
        <input class="en-input" id="enF${col}" type="text" autocomplete="off" spellcheck="false" enterkeyhint="next"
          ${formula ? "readonly tabindex=\"-1\"" : ""} value="${esc(val)}" placeholder="${esc(!formula && now ? now : "")}"
          aria-describedby="enP${col}">
        <div class="en-preview" id="enP${col}" aria-live="polite"></div>
        ${noteArea}
      </div>`;
    }).join("");
    stage.innerHTML = `<div class="tr-card en-card">
      <div class="en-target-line"><span class="en-target-name">${esc(head)}</span>
        <button class="btn btn-xs ghost" id="enChangeTarget" type="button">${esc(t("entryChangeTarget"))}</button></div>
      ${fieldsHtml}</div>`;
    bar.innerHTML = `<span class="en-keys">${esc(t("entryKeysHint"))}</span>
      <button class="btn" id="enReviewBtn" type="button">${esc(ui.options.confirmStep ? t("entryReview") : t("entryCommit"))}</button>`;
    stage.querySelector("#enChangeTarget").addEventListener("click", () => {
      ui.target = null; ui.fields = pinnedFields(); ui.notes = {}; ui.overwriteNotes = new Set(); ui.step = 1; render(); persistNow();
    });
    bar.querySelector("#enReviewBtn").addEventListener("click", toReview);
    stage.querySelectorAll(".en-field").forEach((f) => {
      const col = Number(f.dataset.enCol);
      const btn = f.querySelector(".en-note-btn");
      const area = f.querySelector(".en-note");
      if (!btn || !area) return;
      btn.addEventListener("click", () => {
        const open = area.hidden;
        area.hidden = !open;
        btn.setAttribute("aria-expanded", String(open));
        if (open) area.focus();
      });
      area.addEventListener("input", () => { ui.notes[col] = area.value; persistSoon(); });
      // Enter w notatce = nowa linia; Ctrl/⌘+Enter idzie do podglądu (obsługa globalna)
    });
    const inputs = [...stage.querySelectorAll(".en-field:not(.is-formula) input")];
    inputs.forEach((input, i) => {
      const col = Number(input.closest("[data-en-col]").dataset.enCol);
      // Lista podpowiedzi tylko dla pola z fokusem: inaczej otwarta lista z pliku zasłaniała
      // sąsiednie pola (tap w pinezkę wybierał pozycję listy), a każda lista kręciła własną pętlę rAF.
      let sug = null;
      input.addEventListener("focus", () => {
        if (sug || !(ui.options.suggestions || fieldDv(col))) return;
        sug = attachSuggest(input, col, () => focusField(inputs, i + 1));
      });
      input.addEventListener("blur", () => setTimeout(() => {
        if (!sug || document.activeElement === input || (sug.isInteracting && sug.isInteracting())) return;
        const k = suggests.indexOf(sug);
        if (k >= 0) suggests.splice(k, 1);
        try { sug.destroy(); } catch { /* już zamknięta */ }
        sug = null;
      }, 0));
      input.addEventListener("input", () => { ui.fields[col] = input.value; hint(col); persistSoon(); });
      input.addEventListener("keydown", (e) => {
        if (e.key !== "Enter" || e.ctrlKey || e.metaKey) return;
        e.preventDefault();
        if (e.shiftKey) focusField(inputs, i - 1);
        else if (i === inputs.length - 1) toReview();
        else focusField(inputs, i + 1);
      });
      input.closest("[data-en-col]").querySelector(".en-pin")?.addEventListener("click", (e) => {
        const b = e.currentTarget;
        if (ui.pinned.has(col)) ui.pinned.delete(col); else ui.pinned.add(col);
        const on = ui.pinned.has(col);
        b.setAttribute("aria-pressed", String(on));
        b.setAttribute("aria-label", on ? t("entryUnpin") : t("entryPin"));
        b.title = on ? t("entryUnpin") : t("entryPin");
        persistNow();
      });
      hint(col);
    });
    const firstEmpty = inputs.find((x) => !x.value) || inputs[0];
    if (firstEmpty) firstEmpty.focus();
  }

  const focusField = (inputs, i) => { if (inputs[i]) { inputs[i].focus(); inputs[i].select?.(); } };

  function fieldDv(col) {
    const ref = ui.target && !ui.target.newRow ? ui.target.rowIndex0 : (swbEntry.lastDataRow() || {}).rowIndex0;
    if (!Number.isFinite(ref)) return null;
    try { return getCellDataValidation(currentSheetName, ref, currentStartCol + col); } catch { return null; }
  }

  function attachSuggest(input, col, next) {
    if (typeof createCellSuggestions !== "function") return null;
    const dv = fieldDv(col);
    const values = dv && dv.values.length ? dv.values : (ui.options.suggestions && typeof buildColumnSuggestions === "function" ? buildColumnSuggestions(col) : []);
    if (!values.length) return null;
    const api = createCellSuggestions(input, values, (v) => { input.value = v; ui.fields[col] = v; hint(col); persistSoon(); next(); },
      dv && dv.values.length ? {} : { variant: "soft", minChars: 1 });
    suggests.push(api);
    return api;
  }

  // Wpis innego typu niż reszta kolumny (daty / liczby): ostrzeżenie, nie blokada — pod polem i w podglądzie.
  function typeWarn(col, parsedType) {
    if (!ui.options.typeHints || parsedType === "formula") return [];
    const p = profile(col);
    if (p.type === "date" && parsedType !== "date") return ["entryTypeWarnDate"];
    if (p.type === "number" && parsedType !== "number") return ["entryTypeWarnNumber"];
    return [];
  }

  // Podpowiedź pod polem: jak się zapisze, ostrzeżenie o typie, literówka, nadpisanie.
  function hint(col) {
    const box = overlay.querySelector(`[data-en-col="${col}"] .en-preview`);
    if (!box) return;
    const raw = String(ui.fields[col] ?? "").trim();
    if (!raw) { box.innerHTML = ""; return; }
    const row = ui.target && !ui.target.newRow ? swbEntry.rowAt(ui.target.rowIndex0) : null;
    const refIdx = row ? row.rowIndex0 : ((swbEntry.lastDataRow() || {}).rowIndex0 ?? currentHeaderRow);
    const parsed = parseCellInput(raw, refIdx, col);
    const parts = [];
    if (ui.options.typeHints && parsed && parsed.type !== "formula") {
      parts.push(`<span class="en-will">${esc(t("entryWillSave", { v: swbEntry.previewText(parsed, refIdx, col, raw) }))}</span>`);
      if (typeWarn(col, parsed.type).includes("entryTypeWarnDate")) {
        // „12.03” z kartki to w kolumnie dat prawie na pewno dzień i miesiąc — proponujemy pełną datę
        const m = /^(\d{1,2})[.,](\d{1,2})$/.exec(raw);
        const full = m ? `${m[1]}.${m[2]}.${new Date().getFullYear()}` : null;
        const fp = full ? parseCellInput(full, refIdx, col) : null;
        const fix = fp && fp.type === "date" ? ` <button class="btn btn-xs ghost en-date-fix" type="button" data-v="${esc(full)}">${esc(t("entryTypoFix", { v: full }))}</button>` : "";
        parts.push(`<span class="en-warn">${esc(t("entryTypeWarnDate"))}${fix}</span>`);
      }
      if (typeWarn(col, parsed.type).includes("entryTypeWarnNumber")) parts.push(`<span class="en-warn">${esc(t("entryTypeWarnNumber"))}</span>`);
    }
    if (ui.options.typoHints) {
      const p = profile(col);
      const sug = p.dictionary ? C.typoSuggest(raw, p.known) : null;
      if (sug) parts.push(`<span class="en-warn">${esc(t("entryTypo", { v: sug }))} <button class="btn btn-xs ghost en-typo-fix" type="button" data-v="${esc(sug)}">${esc(t("entryTypoFix", { v: sug }))}</button></span>`);
    }
    if (row) {
      const now = String(getDisplayValue(row, col) ?? "");
      if (now && now !== (parsed && parsed.value != null ? toDisplay(parsed.value) : raw)) parts.push(`<span class="en-over">${esc(t("entryOverwriteMark", { v: now }))}</span>`);
    }
    box.innerHTML = parts.join(" ");
    box.querySelectorAll(".en-typo-fix, .en-date-fix").forEach((b) => b.addEventListener("click", (e) => {
      const v = e.currentTarget.dataset.v;
      const input = overlay.querySelector(`#enF${col}`);
      ui.fields[col] = v;
      if (input) { input.value = v; input.focus(); }
      hint(col);
    }));
  }

  function planInput() {
    return { target: ui.target, fields: { ...ui.fields }, overwrite: [...ui.overwrite], notes: { ...ui.notes }, overwriteNotes: [...ui.overwriteNotes] };
  }

  function toReview() {
    ui.lastPlan = swbEntry.plan(planInput());
    const needsLook = ui.lastPlan.changes.some((c) => c.overwrite) || !ui.lastPlan.ok || !ui.lastPlan.changes.length;
    if (!ui.options.confirmStep && !needsLook) { doCommit(); return; }
    ui.step = 3;
    render();
  }

  // ── Krok 3: podgląd ──
  const ERR = { formula: "entryErrFormula", "dv-stop": "entryErrDvStop", "overwrite-not-allowed": "entryErrOverwrite", "not-real-col": "entryErrNotReal",
    "note-threaded": "entryErrNoteThreaded", "new-row-needs-value": "entryErrNewRowNeedsValue" };
  const GATE = new Set(["overwrite-not-allowed", "note-overwrite-not-allowed"]); // te pokazuje checkbox, nie komunikat
  function renderReview(stage, bar) {
    const p = ui.lastPlan = swbEntry.plan(planInput());
    const row = ui.target.newRow ? null : swbEntry.rowAt(ui.target.rowIndex0);
    const head = ui.target.newRow ? t("entryNewRow") : t("entryTargetExisting", { n: excelRow(ui.target.rowIndex0), key: keyText(row) });
    const kindOf = (e) => e.kind || "value";
    const errsOf = (c) => p.errors.filter((e) => e.col === c.col && kindOf(e) === c.kind && !GATE.has(e.reason));
    const items = p.changes.map((c) => {
      const isNote = c.kind === "note";
      const errs = errsOf(c).map((e) => `<div class="en-err" role="alert">${esc(t(ERR[e.reason] || "entryErrNotReal"))}</div>`).join("");
      const warns = isNote ? "" : p.warnings.filter((w) => w.col === c.col).map(() => `<div class="en-warn">${esc(t("entryWarnDv"))}</div>`).join("")
        + typeWarn(c.col, c.parsedType).map((k) => `<div class="en-warn">${esc(t(k))}</div>`).join("");
      const gateAttr = isNote ? "data-en-overwrite-note" : "data-en-overwrite";
      const gateOn = isNote ? ui.overwriteNotes.has(c.col) : ui.overwrite.has(c.col);
      const ow = c.overwrite ? `<label class="en-ow"><input type="checkbox" ${gateAttr}="${c.col}" ${gateOn ? "checked" : ""}> ${esc(t("entryOverwriteOk"))}</label>` : "";
      return `<li class="en-change${c.overwrite ? " is-overwrite" : ""}${isNote ? " is-note" : ""}">
        <div class="tr-field-label">${esc(headers()[c.col])}${isNote ? ` · ${esc(t("entryNote"))}` : ""}</div>
        <div class="en-diff"><span class="en-before${c.before ? "" : " is-empty"}">${esc(c.before || t("entryEmptyVal"))}</span><span aria-hidden="true">→</span><span class="en-after">${esc(c.after)}</span></div>
        ${ow}${errs}${warns}</li>`;
    }).join("");
    const blocked = p.errors.filter((e) => !GATE.has(e.reason) && !p.changes.some((c) => c.col === e.col && c.kind === kindOf(e)));
    const extra = blocked.map((e) => `<li class="en-change"><div class="tr-field-label">${esc(headers()[e.col] || "")}</div><div class="en-err" role="alert">${esc(t(ERR[e.reason] || "entryErrNotReal"))}</div></li>`).join("");
    stage.innerHTML = `<div class="tr-card en-card">
      <div class="en-target-line"><span class="en-target-name">${esc(head)}</span></div>
      <h2 class="en-review-title">${esc(t("entryReviewTitle"))}</h2>
      ${p.changes.length || blocked.length ? `<ul class="en-changes">${items}${extra}</ul>` : `<p class="en-nothing">${esc(t("entryNothing"))}</p>`}
    </div>`;
    bar.innerHTML = `<button class="btn ghost" id="enBackBtn" type="button">${esc(t("entryBack"))}</button>
      <button class="btn" id="enCommitBtn" type="button" ${p.ok && p.changes.length ? "" : "disabled"}>${esc(t("entryCommit"))}</button>`;
    bar.querySelector("#enBackBtn").addEventListener("click", () => { ui.step = 2; render(); });
    bar.querySelector("#enCommitBtn").addEventListener("click", doCommit);
    stage.querySelectorAll("[data-en-overwrite], [data-en-overwrite-note]").forEach((cb) => cb.addEventListener("change", () => {
      const isNote = cb.hasAttribute("data-en-overwrite-note");
      const col = Number(isNote ? cb.dataset.enOverwriteNote : cb.dataset.enOverwrite);
      const set = isNote ? ui.overwriteNotes : ui.overwrite;
      if (cb.checked) set.add(col); else set.delete(col);
      const q = swbEntry.plan(planInput());
      ui.lastPlan = q;
      bar.querySelector("#enCommitBtn").disabled = !(q.ok && q.changes.length);
    }));
    (bar.querySelector("#enCommitBtn:not([disabled])") || bar.querySelector("#enBackBtn")).focus();
  }

  let committing = false;
  async function doCommit() {
    if (committing) return;
    committing = true;
    try {
      const input = planInput();
      if (ui.lastPlan) input.expect = swbEntry.expectOf(ui.lastPlan);
      const res = await swbEntry.commit(input);
      if (!res.ok) {
        if (res.reason === "stale") toast(t("entryStale"), "warning");
        else if (res.reason === "filtered") toast(t("entryFilteredNow"), "warning");
        else if (res.reason === "guard") toast(t("entryNeedEdit"), "warning");
        else if (res.reason === "write-failed") toast(t("entryWriteFailed", { n: excelRow(res.rowIndex0) }), "error");
        if (res.reason === "write-failed") { afterCommit(res); return; }
        ui.step = 3;
        render();
        return;
      }
      afterCommit(res);
    } finally { committing = false; }
  }

  function afterCommit(res) {
    const row = swbEntry.rowAt(res.rowIndex0);
    ui.sessionLog.unshift({ key: keyText(row), rowIndex0: res.rowIndex0, at: Date.now() });
    ui.profiles.clear();
    toast(t("entryDone", { n: excelRow(res.rowIndex0) }), "success", {
      action: t("entryDoneUndo"),
      onAction: () => { if (swbEntry.undoCard()) toast(t("entryUndone"), "info"); else toast(t("entryUndoGone"), "info"); },
      action2: t("entryDoneShow"),
      onAction2: () => { close(); setFocusedCell(`wide:${res.rowIndex0}`, ui.keyCol, { scroll: true }); },
    });
    ui.fields = pinnedFields();
    ui.notes = {};
    ui.overwriteNotes = new Set();
    ui.target = null;
    ui.query = "";
    ui.overwrite = new Set();
    ui.step = 1;
    persist({ draft: false });
    if (ui.open) { render(); drawPanel(); }
  }

  // ── Otwieranie / zamykanie ──
  function setInert(on) {
    document.querySelectorAll(".app, .hero-overlay").forEach((el) => { if (on) el.setAttribute("inert", ""); else el.removeAttribute("inert"); });
  }
  async function requestWake() {
    try { if ("wakeLock" in navigator) { wakeLock = await navigator.wakeLock.request("screen"); wakeLock.addEventListener("release", () => { wakeLock = null; }); } } catch { /* ekran może zgasnąć */ }
  }

  async function open() {
    if (ui.open) return true;
    if (typeof workbook === "undefined" || !workbook || !currentSheetName || typeof currentDisplayModel === "undefined" || !currentDisplayModel) { toast(t("entryNoData"), "warning"); return false; }
    if (typeof swbIsEditMode === "function" && !swbIsEditMode()) { toast(t("entryNeedEdit"), "warning"); return false; }
    if (currentDisplayModel.mode !== "wide") { toast(t("entryNeedWide"), "warning"); return false; }
    if (typeof swbRowGrow !== "undefined" && swbRowGrow.filtered()) {
      const ans = await appDialog({
        title: t("entryFilterTitle"), message: t("entryFilterMsg"),
        actions: [{ label: t("entryFilterClear"), value: "clear", kind: "primary" }, { label: t("entryCancel"), value: null, kind: "secondary" }],
      });
      if (ans !== "clear") return false;
      resetFilterInputs();
      if (typeof quickSearchEl !== "undefined" && quickSearchEl) quickSearchEl.value = "";
      applyQuickSearch();
      if (swbRowGrow.filtered()) { toast(t("entryFilterStill"), "warning"); return false; }
    }
    ui.returnFocus = document.activeElement;
    ui.profiles.clear();
    if (ui.scope !== scopeKey()) { ui.target = null; ui.fields = {}; ui.step = 1; ui.sessionLog = ui.scope ? [] : ui.sessionLog; ui.scope = scopeKey(); }
    restore();
    const rows = swbEntry.candidates();
    ui.keyCol = Number.isInteger(ui.keyColSaved) ? ui.keyColSaved : C.pickKeyColumn(headers(), headers().map((_, i) => rows.map((r) => r.values[i])));
    ui.open = true;
    ui.step = ui.target ? ui.step : 1;
    overlay.dataset.font = String(ui.font);
    shell();
    overlay.classList.remove("hidden");
    document.body.classList.add("en-active");
    setInert(true);
    requestWake();
    ui.panel = null;
    render();
    return true;
  }

  function close() {
    if (!ui.open) return;
    destroySuggests();
    ui.open = false;
    overlay.classList.add("hidden");
    document.body.classList.remove("en-active");
    setInert(false);
    try { wakeLock?.release?.(); } catch { /* ignore */ }
    wakeLock = null;
    const back = ui.returnFocus;
    ui.returnFocus = null;
    persistNow();
    if (back && document.contains(back) && back.getClientRects().length) back.focus();
    else document.getElementById("moreMenuBtn")?.focus();
    // raz na sesję: kartki są w aplikacji, plik trzeba jeszcze zapisać
    if (!ui.reminded && ui.sessionLog.length && typeof hasUnsavedChanges !== "undefined" && hasUnsavedChanges) {
      ui.reminded = true;
      toast(t("entryUnsaved"), "info");
    }
  }

  // Klawisze: Esc i Ctrl/⌘+Enter tu; reszta zostaje w karcie (nie dochodzi do tabeli i Cofnij).
  document.addEventListener("keydown", (e) => {
    if (!ui.open) return;
    if (document.querySelector(".modal.app-dialog:not(.hidden)")) return; // okno nad kartą obsługuje się samo
    if (e.key === "Escape") {
      e.preventDefault(); e.stopPropagation();
      // najpierw zamyka się lista podpowiedzi pod polem, dopiero drugie Esc — karta
      const open = document.querySelectorAll(".cell-suggest:not(.hidden)");
      if (open.length) { open.forEach((b) => b.classList.add("hidden")); return; }
      if (ui.panel) { togglePanel(ui.panel); overlay.querySelector("#enSettingsBtn")?.focus(); return; }
      close(); return;
    }
    if (e.key === "Enter" && (e.ctrlKey || e.metaKey)) {
      e.preventDefault(); e.stopPropagation();
      if (ui.step === 2) toReview();
      else if (ui.step === 3) { const b = overlay.querySelector("#enCommitBtn"); if (b && !b.disabled) doCommit(); }
    }
  }, true);
  overlay.addEventListener("keydown", (e) => e.stopPropagation());
  document.addEventListener("visibilitychange", () => {
    if (!ui.open) return;
    if (document.visibilityState === "hidden") persistNow(); // iOS potrafi zamknąć aplikację w tle bez ostrzeżenia
    else if (!wakeLock) requestWake();
  });
  window.addEventListener("pagehide", () => { if (ui.open) persistNow(); });
  if (btn) btn.addEventListener("click", () => {
    open();
  });

  // Stan i haki dla Task 4 (pamięć) i testów.
  swbEntry.open = open;
  swbEntry.close = close;
  swbEntry.ui = ui;

  swbEntry.state = () => ({
    open: ui.open, step: ui.step, target: ui.target ? { ...ui.target } : null, fields: { ...ui.fields },
    pinned: [...ui.pinned], keyCol: ui.keyCol, options: { ...ui.options }, sessionLog: ui.sessionLog.map((x) => ({ ...x })),
    fieldOrder: ui.fieldOrder.slice(), selected: [...ui.selected], notes: { ...ui.notes },
  });
  // fill() z silnika + przejście do kroku 2 (wejście programowe, np. przyszły skan karty)
  const engineFill = swbEntry.fill;
  swbEntry.fill = (input) => {
    const res = engineFill(input);
    if (input && input.target) {
      ui.target = { ...input.target };
      ui.fields = { ...((input && input.fields) || {}) };
      ui.overwrite = new Set();
      ui.step = 2;
      if (ui.open) render();
    }
    return res;
  };
})();

window.__entry = swbEntry;
