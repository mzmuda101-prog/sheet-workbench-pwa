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
    if (code && code !== "General" && typeof cfmtFormatValue === "function") { // także data: „d mmm yy” → „15 lis 26”, jak komórka
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
    if (!isNew && expect && isStale(rowAt(input.target.rowIndex0), expect)) return { ok: false, reason: "stale", plan: p };
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
    const failed = writeChanges(row, p.changes);
    const steps = await finishWrite(topBefore, depthBefore);
    if (failed.length) return { ok: false, reason: "write-failed", failed, rowIndex0: row.rowIndex0, undoSteps: steps };
    return { ok: true, rowIndex0: row.rowIndex0, undoSteps: steps };
  }

  function snapshotOf(row, fields, notes) {
    const out = {};
    if (!row) return out;
    Object.keys(fields || {}).forEach((k) => { if (!isBlank(fields[k])) out[k] = String(getDisplayValue(row, Number(k)) ?? ""); });
    Object.keys(notes || {}).forEach((k) => { if (!isBlank(notes[k])) { const cur = noteAt(row.rowIndex0, Number(k)); out[`n${k}`] = cur ? cur.text : ""; } });
    return out;
  }

  // Arkusz zmieniony od podglądu? expect = { "kol": tekst wartości, "n<kol>": tekst notatki }.
  function isStale(row, expect) {
    if (!row) return true;
    return Object.keys(expect || {}).some((k) => {
      if (k[0] === "n") { const cur = noteAt(row.rowIndex0, Number(k.slice(1))); return (cur ? cur.text : "") !== expect[k]; }
      return String(getDisplayValue(row, Number(k)) ?? "") !== expect[k];
    });
  }

  // Zapis zmian jednego wiersza (wartości, potem notatki). Wołane w tym samym zadaniu JS dla całej
  // kartki albo całej serii — undo.js składa to w jeden krok Cofnij. Zwraca listę nieudanych.
  function writeChanges(row, changes) {
    const failed = [];
    changes.filter((c) => c.kind === "value").forEach((c) => {
      const res = writeCellFromInput(currentDisplayModel, row, c.col, c.raw);
      if (res !== "ok") failed.push({ rowIndex0: row.rowIndex0, col: c.col, reason: res });
      else if (typeof rememberColumnValue === "function") {
        const v = row.values[c.col];
        rememberColumnValue(c.col, v instanceof Date && typeof formatDateForEdit === "function" ? formatDateForEdit(v) : v);
      }
    });
    // notatki po wartościach (nowy wiersz już istnieje), w tym samym zadaniu = ten sam krok Cofnij
    changes.filter((c) => c.kind === "note").forEach((c) => {
      const okNote = typeof swbSetNote === "function" && swbSetNote(row.rowIndex0, c.col, c.raw, typeof swbNoteAuthor === "function" ? swbNoteAuthor() : "");
      if (!okNote) failed.push({ rowIndex0: row.rowIndex0, col: c.col, reason: "note" });
    });
    return failed;
  }

  // Po zapisie: odśwież tabelę, policz kroki Cofnij (domykają się mikro-zadaniem) i zapamiętaj je.
  async function finishWrite(topBefore, depthBefore) {
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
    return steps;
  }

  // ── Kolejka „wszystkim naraz”: te same pola w wielu ISTNIEJĄCYCH wierszach ──
  // overwriteAll = jedna zgoda na nadpisania we wszystkich wierszach (pokazana z liczbą w podglądzie).
  // items = [{ rowIndex0, fields, notes, overwriteNotes }] — każdy wiersz z WŁASNYMI polami (tryb smart:
  // data w kolumnie swojego cyklu); bez items: te same fields/notes w każdym z targets.
  function planMany(input) {
    const items = Array.isArray(input && input.items) ? input.items
      : (Array.isArray(input && input.targets) ? input.targets : []).map((r) => ({ rowIndex0: r, fields: input.fields || {}, notes: input.notes || {} }));
    const all = currentDisplayModel && Array.isArray(currentDisplayModel.headers) ? currentDisplayModel.headers.map((_, i) => i) : [];
    const rows = items.map((it) => ({
      rowIndex0: it.rowIndex0,
      plan: plan({ target: { rowIndex0: it.rowIndex0 }, fields: it.fields || {}, notes: it.notes || {}, overwrite: input.overwriteAll ? all : [], overwriteNotes: input.overwriteAll ? all : (it.overwriteNotes || []) }),
    }));
    const ow = rows.reduce((n, r) => n + (r.plan.changes.some((c) => c.overwrite) ? 1 : 0), 0);
    const changes = rows.reduce((n, r) => n + r.plan.changes.length, 0);
    lastPlan = null;
    return {
      ok: rows.length > 0 && rows.every((r) => r.plan.ok) && changes > 0,
      rows, overwriteCount: ow, changes,
      // „było” dla KAŻDEGO pola z formularza, także bez zmiany w podglądzie — inaczej komórka zmieniona
      // w międzyczasie (np. Cofnij w tabeli) zostałaby nadpisana zgodą „we wszystkich” bez pokazania
      expect: Object.fromEntries(items.map((it) => [it.rowIndex0, snapshotOf(rowAt(it.rowIndex0), it.fields, it.notes)])),
    };
  }

  async function commitMany(input) {
    if (!ready()) return { ok: false, reason: "guard" };
    if (typeof swbIsEditMode === "function" && !swbIsEditMode()) return { ok: false, reason: "guard" };
    const p = planMany(input);
    if (!p.rows.length) return { ok: false, reason: "empty", plan: p };
    if (p.rows.some((r) => r.plan.errors.some((e) => e.reason === "no-target"))) return { ok: false, reason: "stale", plan: p };
    const expect = input && input.expect;
    if (expect && p.rows.some((r) => expect[r.rowIndex0] && isStale(rowAt(r.rowIndex0), expect[r.rowIndex0]))) return { ok: false, reason: "stale", plan: p };
    if (!p.rows.every((r) => r.plan.ok)) return { ok: false, reason: "invalid", plan: p };
    if (!p.changes) return { ok: false, reason: "empty", plan: p };
    const topBefore = undoTop();
    const depthBefore = typeof undoStack !== "undefined" ? undoStack.length : 0;
    const failed = [];
    const written = [];
    p.rows.forEach((r) => {
      if (!r.plan.changes.length) return;
      failed.push(...writeChanges(rowAt(r.rowIndex0), r.plan.changes));
      written.push({ rowIndex0: r.rowIndex0, changes: r.plan.changes.length });
    });
    const steps = await finishWrite(topBefore, depthBefore);
    if (failed.length) return { ok: false, reason: "write-failed", failed, rows: written, undoSteps: steps };
    return { ok: true, rows: written, undoSteps: steps };
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

  // ── Tryb smart: źródło cykli dla bieżącego arkusza (z pamięci, jak lista wierszy) ──
  // roles = ręczna poprawka z Ustawień (null = zgadywanie). Formuła = kolumna z formułą w którymś z pierwszych wierszy.
  function cycleSource(roles = null) {
    const m = memoNow();
    if (!m || typeof SwbEntryCore === "undefined") return null;
    const key = JSON.stringify(roles || null);
    if (!m.cycles) m.cycles = new Map();
    if (!m.cycles.has(key)) {
      const hs = currentDisplayModel.headers;
      const probe = m.rows.slice(0, 20);
      const info = hs.map((_, c) => ({
        type: SwbEntryCore.profileColumn(m.rows.map((r) => r.values[c])).type,
        formula: probe.some((r) => { const x = sheetCell(r.rowIndex0, c); return !!(x && x.f); }),
      }));
      m.cycles.set(key, SwbEntryCore.detectCycleColumns(hs, info, { roles }));
    }
    return m.cycles.get(key);
  }
  // Decyzja smart dla wiersza: planCycleAction + stan cykli (do paska stanu i okna „poprzedni otwarty”).
  function smartPlan(rowIndex0, action, opts = {}) {
    const src = cycleSource(opts.roles || null);
    const row = rowAt(rowIndex0);
    if (!src || !row) return { status: "no-source" };
    const cycles = SwbEntryCore.cyclesOfRow(row.values, src);
    return { ...SwbEntryCore.planCycleAction(cycles, action, opts), state: SwbEntryCore.cycleState(cycles), cycles };
  }

  const api = { candidates, plan, commit, planMany, commitMany, undoCard, fill, sheetCell, rowAt, lastDataRow, ready, previewText, foldCache, expectOf, searchRows, cycleSource, smartPlan, _filled: () => filled };
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

  const DEFAULT_OPTIONS = { fuzzyTarget: true, typoHints: true, typeHints: true, suggestions: true, duplicateGuard: true, confirmStep: true, queryOperators: true,
    queueLists: true, queueAutoNext: true, queueMissing: "skip", orNumbersAsList: false }; // || domyślnie = „lub” jak w szybkim szukaniu (decyzja Mateusza 2026-11-16)
  const ui = {
    open: false, step: 1, keyCol: 0, query: "", targets: [], active: -1,
    target: null, fields: {}, overwrite: new Set(), pinned: new Set(), font: 2,
    options: { ...DEFAULT_OPTIONS }, profiles: new Map(), lastPlan: null, sessionLog: [], returnFocus: null,
    fieldOrder: [], selected: new Set(), panel: null, draft: null, reminded: false,
    notes: {}, overwriteNotes: new Set(), queryMode: false, queryTotal: 0,
    list: null, queue: null, queueDraft: null,
    smartCfg: null, smart: null, smartLast: null,
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
    rec.queue = ui.queue ? JSON.parse(JSON.stringify(ui.queue)) : null;
    rec.smart = { ...ui.smartCfg };
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
    ui.smartCfg = { ...DEFAULT_SMART, ...((rec && rec.smart) || {}) };
    if (!ui.smart) ui.smart = newSmart();
    ui.pinned = new Set();
    const pinnedVals = {};
    if (rec && rec.pinned) Object.keys(rec.pinned).forEach((k) => { const c = Number(k); if (valid(c)) { ui.pinned.add(c); if (rec.pinned[k]) pinnedVals[c] = rec.pinned[k]; } });
    if (!ui.target) ui.fields = { ...pinnedVals };
    ui.draft = rec && rec.draft && rec.draft.target ? rec.draft : null;
    // niedokończona kolejka: wiersze mogły zniknąć (inny plik pod tą nazwą) — wtedy pozycja wraca do wyboru
    ui.queueDraft = null;
    if (!ui.queue && rec && rec.queue && Array.isArray(rec.queue.items) && rec.queue.items.length) {
      const q = rec.queue;
      q.items.forEach((it) => { if (Number.isInteger(it.choice) && !swbEntry.rowAt(it.choice)) { it.choice = null; it.status = "missing"; } });
      ui.queueDraft = q;
    }
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
  // Odmiana do komunikatów z liczbą: „1 pozycja / 3 pozycje / 5 pozycji”, „do 1 wiersza”, „w 2 wierszach”.
  function countWords(n) {
    if (typeof currentLang !== "undefined" && currentLang === "en") return { items: n === 1 ? "item" : "items", rowsGen: n === 1 ? "row" : "rows", rowsLoc: n === 1 ? "row" : "rows" };
    const f = typeof plPluralForm === "function" ? plPluralForm(n) : (n === 1 ? "one" : "many");
    return { items: f === "one" ? "pozycja" : f === "few" ? "pozycje" : "pozycji", rowsGen: f === "one" ? "wiersza" : "wierszy", rowsLoc: f === "one" ? "wierszu" : "wierszach" };
  }
  const tn = (key, n, extra = {}) => t(key, { n, ...countWords(n), ...extra });
  const keyText = (row) => (row ? String(getDisplayValue(row, ui.keyCol) ?? "") : "");

  // ── Tryb smart (cykle od/do): karta pokazuje akcję, a silnik sam wybiera kolumny ──
  // Nazwy akcji i słowo na pozycję są per arkusz (Ustawienia), bo z kart korzystają różne osoby z różnymi arkuszami.
  const DEFAULT_SMART = { on: true, issueName: "", closeName: "", item: "", dateDefault: "today", anomalies: "ask", roles: null };
  const cfg = () => ui.smartCfg || DEFAULT_SMART;
  const smartSrc = () => (cfg().on && swbEntry.ready() ? swbEntry.cycleSource(cfg().roles) : null);
  const smartName = (a) => cfg()[a === "issue" ? "issueName" : "closeName"] || t(a === "issue" ? "entrySmartIssue" : "entrySmartClose");
  const itemWord = () => cfg().item || t("entrySmartItem");
  const smartDefaultDate = () => (cfg().dateDefault === "empty" ? "" : C.todayText());
  function newSmart(over = {}) { return { off: false, action: null, date: null, data: {}, resolve: null, note: "", ...over }; }
  const smartActive = () => !!(ui.target && !ui.target.newRow && ui.smart && !ui.smart.off && smartSrc());
  const smartOpts = (s, resolve = s.resolve) => ({ date: s.date ?? smartDefaultDate(), data: s.data, resolve, roles: cfg().roles });
  const cellOf = (row, col) => String(getDisplayValue(row, col) ?? "");
  function cycleStateText(row, p) {
    if (!row || !p || !p.state) return "";
    const st = p.state, cy = p.cycles;
    const who = (c) => c.data.map((d) => cellOf(row, d.col)).filter(Boolean).join(" · ");
    const parts = [];
    if (st.issues.length) parts.push(t("entrySmartCheck", { why: st.issues.map((x) => t(x === "odd" ? "entrySmartOdd" : "entrySmartManyOpen")).join(", ") }));
    st.open.forEach((i) => { const c = cy[i]; parts.push([t("entrySmartOpenSince", { n: c.n, date: cellOf(row, c.start.col) }), who(c)].filter(Boolean).join(": ")); });
    if (!st.issues.length && !st.open.length) parts.push(st.free >= 0 ? t("entrySmartFree", { n: st.free + 1 }) : t("entrySmartFull"));
    return parts.join(" · ");
  }
  // Nota „Tylko notatka”: w komórce klucza; istniejąca notatka zostaje, nowa treść dochodzi pod nią.
  function smartNote(rowIndex0, text) {
    const cur = typeof swbNoteOf === "function" ? swbNoteOf(currentSheetName, XLSX.utils.encode_cell({ r: rowIndex0, c: currentStartCol + ui.keyCol })) : null;
    return { notes: { [ui.keyCol]: cur && cur.text ? `${cur.text}\n${text}` : text }, overwriteNotes: cur ? [ui.keyCol] : [] };
  }
  // główna kolumna decyzji (do pokazania „→ do2”): koniec przy Zdaj, start przy Wydaj
  const smartCol = (p, action) => (p.cycles && p.cycle ? p.cycles[p.cycle - 1][action === "close" ? "end" : "start"].col : null);

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
        <h2 class="en-panel-title">${esc(t("entryQueueTitle"))}</h2>
        ${[["queueLists", "entryOptQueueLists"], ["orNumbersAsList", "entryOptOrList"], ["queueAutoNext", "entryOptQueueAutoNext"]].map(([k, label]) => `<label class="en-check"><input type="checkbox" data-en-opt="${k}" ${ui.options[k] ? "checked" : ""}> ${esc(t(label))}</label>`).join("")}
        <label class="en-label" for="enQueueMissing">${esc(t("entryQueueMissingLabel"))}</label>
        <select class="en-select" id="enQueueMissing">
          <option value="skip" ${ui.options.queueMissing !== "new" ? "selected" : ""}>${esc(t("entryQSkipOpt"))}</option>
          <option value="new" ${ui.options.queueMissing === "new" ? "selected" : ""}>${esc(t("entryQNewOpt"))}</option>
        </select>
        ${smartSettingsHtml()}
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
      slot.querySelector("#enQueueMissing").addEventListener("change", (e) => { ui.options.queueMissing = e.target.value === "new" ? "new" : "skip"; persistNow(); });
      wireSmartSettings(slot);
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

  // Ustawienia → Tryb smart (per arkusz). Sekcja tylko, gdy arkusz ma cykle (wykrycie automatyczne).
  function smartSettingsHtml() {
    const auto = swbEntry.ready() ? swbEntry.cycleSource(null) : null;
    if (!auto) return "";
    const c = cfg();
    const src = swbEntry.cycleSource(c.roles) || auto;
    const roleOpt = (p, r) => `<option value="${r}" ${src.roles[p] === r ? "selected" : ""}>${esc(t({ data: "entrySmartRoleData", start: "entrySmartRoleStart", end: "entrySmartRoleEnd", skip: "entrySmartRoleSkip" }[r]))}</option>`;
    const sel = (id, label, val, opts) => `<label class="en-label" for="${id}">${esc(t(label))}</label>
      <select class="en-select" id="${id}">${opts.map(([v, k]) => `<option value="${v}" ${val === v ? "selected" : ""}>${esc(t(k))}</option>`).join("")}</select>`;
    const txt = (id, label, val, ph) => `<label class="en-label" for="${id}">${esc(t(label))}</label>
      <input class="en-input" id="${id}" type="text" autocomplete="off" value="${esc(val)}" placeholder="${esc(ph)}">`;
    return `<h2 class="en-panel-title">${esc(t("entrySmartTitle"))}</h2>
      <p class="en-hint">${esc(t("entrySmartDetected", { count: src.count, span: src.span, labels: src.labels.join(", ") }))}</p>
      <label class="en-check"><input type="checkbox" id="enSmartOnOpt" ${c.on ? "checked" : ""}> ${esc(t("entrySmartOptOn"))}</label>
      ${txt("enSmartIssueName", "entrySmartIssueNameLabel", c.issueName, t("entrySmartIssue"))}
      ${txt("enSmartCloseName", "entrySmartCloseNameLabel", c.closeName, t("entrySmartClose"))}
      ${txt("enSmartItemWord", "entrySmartItemLabel", c.item, t("entrySmartItem"))}
      ${sel("enSmartDateDefault", "entrySmartDateLabel", c.dateDefault, [["today", "entrySmartDateToday"], ["empty", "entrySmartDateEmpty"]])}
      ${sel("enSmartAnomalies", "entrySmartAnomLabel", c.anomalies, [["ask", "entrySmartAnomAsk"], ["skip", "entrySmartAnomSkip"]])}
      <div class="en-label">${esc(t("entrySmartRolesLabel"))}</div>
      <ul class="en-fieldlist">${src.labels.map((l, p) => `<li class="en-fieldrow"><label class="en-role-label" for="enRole${p}">${esc(l)}</label>
        <select class="en-select" id="enRole${p}" data-en-role="${p}">${["data", "start", "end", "skip"].map((r) => roleOpt(p, r)).join("")}</select></li>`).join("")}</ul>
      <p class="en-warn" id="enSmartRolesBad" hidden>${esc(t("entrySmartRolesBad"))}</p>`;
  }
  function wireSmartSettings(slot) {
    if (!slot.querySelector("#enSmartOnOpt")) return;
    const c = ui.smartCfg;
    const save = () => { persistNow(); };
    slot.querySelector("#enSmartOnOpt").addEventListener("change", (e) => { c.on = e.target.checked; save(); });
    [["enSmartIssueName", "issueName"], ["enSmartCloseName", "closeName"], ["enSmartItemWord", "item"]].forEach(([id, k]) => {
      const el = slot.querySelector(`#${id}`);
      const set = () => { c[k] = el.value.trim(); save(); };
      el.addEventListener("input", set); el.addEventListener("change", set);
    });
    slot.querySelector("#enSmartDateDefault").addEventListener("change", (e) => { c.dateDefault = e.target.value === "empty" ? "empty" : "today"; save(); });
    slot.querySelector("#enSmartAnomalies").addEventListener("change", (e) => { c.anomalies = e.target.value === "skip" ? "skip" : "ask"; save(); });
    slot.querySelectorAll("[data-en-role]").forEach((el) => el.addEventListener("change", () => {
      const roles = [...slot.querySelectorAll("[data-en-role]")].map((x) => x.value);
      const good = roles.filter((r) => r === "start").length === 1 && roles.filter((r) => r === "end").length === 1;
      slot.querySelector("#enSmartRolesBad").hidden = good;
      c.roles = good ? roles : null;
      save();
    }));
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
    else if (ui.step === "queue") renderQueue(stage, bar);
    else if (ui.step === "queueReview") renderQueueReview(stage, bar);
    else renderReview(stage, bar);
  }

  // ── Krok 1: cel ──
  function renderTarget(stage, bar) {
    stage.innerHTML = `
      <div class="tr-card en-card">
        <div class="en-target-line"><label class="en-label" for="enTargetInput">${esc(t("entryTargetLabel"))}</label>
          <button class="btn btn-xs ghost" id="enHelpBtn" type="button" aria-expanded="${!!ui.helpOpen}" aria-controls="enHelp">${esc(t("entryHelpBtn"))}</button></div>
        <dl class="en-help" id="enHelp" ${ui.helpOpen ? "" : "hidden"}>
          <dt>${esc(t("entryHelpOne"))}</dt><dd>${esc(t("entryHelpOneEx"))}</dd>
          <dt>${esc(t("entryHelpList"))}</dt><dd>${esc(t("entryHelpListEx"))}</dd>
          <dt>${esc(t("entryHelpQuery"))}</dt><dd>${esc(t("entryHelpQueryEx"))}</dd>
          <dt>${esc(t("entryHelpNew"))}</dt><dd>${esc(t("entryHelpNewEx"))}</dd>
        </dl>
        <input class="en-input en-target-input" id="enTargetInput" type="text" autocomplete="off" spellcheck="false"
          enterkeyhint="go" role="combobox" aria-expanded="false" aria-controls="enCandidates" aria-autocomplete="list"
          placeholder="${esc(t("entryTargetPlaceholder"))}" value="${esc(ui.query)}">
        <div class="en-hint" id="enKeyHint">${esc(t("entryTargetHint", { col: headers()[ui.keyCol] || "" }))}</div>
        <div class="en-candidates" id="enCandidates" role="listbox" aria-label="${esc(t("entryTargetLabel"))}"></div>
      </div>`;
    bar.innerHTML = `<span class="en-keys">${esc(t("entryKeysHint"))}</span>`;
    if (ui.queueDraft && !ui.queue) {
      const qd = ui.queueDraft;
      const note = document.createElement("div");
      note.className = "tr-notice en-draft";
      note.id = "enQueueDraft";
      note.setAttribute("role", "status");
      note.innerHTML = `<div class="tr-notice-text">${esc(t("entryQDraft", { done: qd.done || 0, n: qd.items.length }))}</div>
        <div class="tr-notice-actions"><button class="btn btn-xs" id="enQueueDraftResume" type="button">${esc(t("entryDraftResume"))}</button>
        <button class="btn btn-xs ghost" id="enQueueDraftDiscard" type="button">${esc(t("entryDraftDiscard"))}</button></div>`;
      stage.prepend(note);
      note.querySelector("#enQueueDraftResume").addEventListener("click", () => {
        ui.queue = ui.queueDraft; ui.queueDraft = null; ui.queue.mode = null; ui.target = null; ui.step = "queue"; render();
      });
      note.querySelector("#enQueueDraftDiscard").addEventListener("click", () => {
        ui.queueDraft = null; persistNow(); note.remove(); overlay.querySelector("#enTargetInput")?.focus();
      });
    }
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
        ui.smart = newSmart({ off: true }); // szkic to wpisane pola — wraca zwykła karta
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
    stage.querySelector("#enHelpBtn").addEventListener("click", (e) => {
      ui.helpOpen = !ui.helpOpen;
      e.currentTarget.setAttribute("aria-expanded", String(ui.helpOpen));
      stage.querySelector("#enHelp").hidden = !ui.helpOpen;
    });
    const input = stage.querySelector("#enTargetInput");
    input.addEventListener("input", () => { ui.query = input.value; search(); });
    input.addEventListener("keydown", onTargetKey);
    search();
    input.focus();
  }

  function search() {
    const rows = swbEntry.candidates();
    const q = C.trimQuery(ui.query);
    // lista celów („031, 032, 045-048”) = propozycja kolejki zamiast szukania jednego wiersza
    ui.list = ui.options.queueLists ? C.parseTargetList(q, headers(), { orList: ui.options.orNumbersAsList }) : null;
    if (ui.list) { ui.queryMode = false; ui.queryCrit = null; ui.targets = []; ui.active = 0; drawCandidates(); return; }
    ui.queryMode = !!ui.options.queryOperators && C.looksLikeQuery(q, headers());
    ui.queryCrit = null;
    if (ui.queryMode) {
      // ta sama składnia i ten sam silnik co szybkie szukanie (workbook.js) — po wszystkich kolumnach pliku
      const criterion = { query: q, mode: "contains", indexes: realCols(), headers: headers(), operatorsEnabled: true, negated: false, emptyMode: "all" };
      ui.queryCrit = criterion;
      let hits = [];
      try { hits = rows.filter((r) => rowMatchesTextFilter(r, [criterion], false)); } catch { hits = []; }
      ui.queryTotal = hits.length;
      ui.queryHits = hits.map((r) => r.rowIndex0);
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
    if (ui.list) {
      const dup = ui.list.dupes.length ? `<div class="en-hint">${esc(t("entryQueueDupes", { list: ui.list.dupes.join(", ") }))}</div>` : "";
      box.innerHTML = `<div class="en-cand en-cand-new en-cand-queue" role="option" id="enCand0" data-en-queue aria-selected="true">${esc(tn("entryQueueOpen", ui.list.items.length))}</div>${dup}`;
      input.setAttribute("aria-expanded", "true");
      input.setAttribute("aria-activedescendant", "enCand0");
      const opt = box.querySelector("[data-en-queue]");
      opt.addEventListener("pointerdown", (e) => e.preventDefault());
      opt.addEventListener("click", () => openQueue(ui.list));
      const keyHint = overlay.querySelector("#enKeyHint");
      if (keyHint) keyHint.hidden = false;
      return;
    }
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
    // przy kilku wynikach zapytania: wszystkie naraz do kolejki (np. „Status:=wzięta” → przejdź po każdym)
    const allBtn = ui.queryMode && ui.queryTotal >= 2
      ? `<button class="btn btn-xs en-queue-all" type="button" data-en-queue-all>${esc(t("entryQueueAll", { n: ui.queryTotal }))}</button>` : "";
    box.innerHTML = (ui.queryMode ? `<div class="en-hint en-query-hint" id="enQueryHint">${esc(t("entryQueryHint", { n: ui.queryTotal }))}</div>${allBtn}` : "") + items.join("");
    box.querySelector("[data-en-queue-all]")?.addEventListener("click", () => openQueueRows(ui.queryHits || []));
    input.setAttribute("aria-expanded", items.length ? "true" : "false");
    if (ui.active >= 0) input.setAttribute("aria-activedescendant", `enCand${ui.active}`); else input.removeAttribute("aria-activedescendant");
    box.querySelectorAll(".en-cand").forEach((el, i) => { // przycisk „wszystkie do kolejki” nie jest opcją listy
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
      if (ui.list) { openQueue(ui.list); return; }
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
    ui.smart = newSmart();
    ui.smartLast = null;
    ui.step = 2;
    ui.draft = null;
    render();
    persistNow();
  }

  // ── Krok 2: pola ──
  function renderFields(stage, bar) {
    if (smartActive()) { renderSmart(stage, bar); return; }
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
          ${formula ? "readonly tabindex=\"-1\"" : ""} value="${esc(val)}" placeholder="${esc(!formula && now ? now : "")}"${!formula && now ? ' data-en-now=""' : ""}
          aria-describedby="enP${col}">
        <div class="en-preview" id="enP${col}" aria-live="polite"></div>
        ${noteArea}
      </div>`;
    }).join("");
    const walking = !!(ui.queue && ui.queue.mode === "walk");
    const fixing = !!(ui.queue && ui.queue.mode === "fix");
    const targetTools = fixing ? `<button class="btn btn-xs ghost" id="enQFixBack" type="button">${esc(t("entryQBack"))}</button>`
      : walking
      ? `<span class="en-field-tools"><button class="btn btn-xs ghost" id="enQBack" type="button">${esc(t("entryQBack"))}</button>
         <button class="btn btn-xs ghost" id="enQSkip" type="button">${esc(t("entryQSkip"))}</button></span>`
      : `<button class="btn btn-xs ghost" id="enChangeTarget" type="button">${esc(t("entryChangeTarget"))}</button>`;
    const smartBack = fixing ? "" : !ui.target.newRow && ui.smart && ui.smart.off && smartSrc()
      ? `<button class="btn btn-xs ghost en-smart-toggle" id="enSmartOn" type="button">${esc(t("entrySmartBack"))}</button>` : "";
    stage.innerHTML = `<div class="tr-card en-card">
      ${walking ? queueProgressHtml() : ""}
      <div class="en-target-line"><span class="en-target-name">${esc(head)}</span>${targetTools}</div>
      ${smartBack}${fieldsHtml}</div>`;
    stage.querySelector("#enSmartOn")?.addEventListener("click", () => { ui.smart.off = false; render(); });
    stage.querySelector("#enQFixBack")?.addEventListener("click", endQueueFix);
    bar.innerHTML = `<span class="en-keys">${esc(t("entryKeysHint"))}</span>
      <button class="btn" id="enReviewBtn" type="button">${esc(ui.options.confirmStep ? t("entryReview") : (walking ? t("entryQCommitNext") : t("entryCommit")))}</button>`;
    stage.querySelector("#enChangeTarget")?.addEventListener("click", () => {
      ui.target = null; ui.fields = pinnedFields(); ui.notes = {}; ui.overwriteNotes = new Set(); ui.step = 1; render(); persistNow();
    });
    if (walking) wireQueueWalk(stage);
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

  // ── Karta smart jednego wiersza ──
  function renderSmart(stage, bar) {
    const src = smartSrc();
    const row = swbEntry.rowAt(ui.target.rowIndex0);
    const s = ui.smart;
    if (s.date === null) s.date = smartDefaultDate();
    const now = swbEntry.smartPlan(row.rowIndex0, "close", smartOpts(s, null));
    if (!s.action) s.action = now.state && now.state.open.length ? "close" : "issue";
    const walking = !!(ui.queue && ui.queue.mode === "walk");
    const head = t("entryTargetExisting", { n: excelRow(row.rowIndex0), key: keyText(row) });
    const targetTools = walking
      ? `<span class="en-field-tools"><button class="btn btn-xs ghost" id="enQBack" type="button">${esc(t("entryQBack"))}</button>
         <button class="btn btn-xs ghost" id="enQSkip" type="button">${esc(t("entryQSkip"))}</button></span>`
      : `<button class="btn btn-xs ghost" id="enChangeTarget" type="button">${esc(t("entryChangeTarget"))}</button>`;
    const dataPos = src.roles.map((r, p) => (r === "data" ? p : -1)).filter((p) => p >= 0);
    const dateLabel = src.labels[src.roles.indexOf(s.action === "close" ? "end" : "start")];
    const dataHtml = s.action !== "issue" ? "" : dataPos.map((p, i) => `<div class="tr-field en-field">
        <label class="tr-field-label" for="enSD${i}">${esc(src.labels[p])}</label>
        <input class="en-input" id="enSD${i}" data-en-smart-data="${i}" data-col="${src.first + p}" type="text" autocomplete="off" spellcheck="false" enterkeyhint="next" value="${esc(s.data[i] ?? "")}"></div>`).join("");
    stage.innerHTML = `<div class="tr-card en-card" id="enSmart">
      ${walking ? queueProgressHtml() : ""}
      <div class="en-target-line"><span class="en-target-name">${esc(head)}</span>${targetTools}</div>
      <p class="en-smart-state" id="enSmartState">${esc(`${itemWord()} ${keyText(row)} · ${cycleStateText(row, now)}`)}</p>
      <div class="en-seg" role="group" aria-label="${esc(t("entrySmartAction"))}">
        ${["close", "issue"].map((a) => `<button class="en-seg-btn" type="button" data-en-action="${a}" aria-pressed="${s.action === a}">${esc(smartName(a))}</button>`).join("")}
      </div>
      ${dataHtml}
      <div class="tr-field en-field">
        <label class="tr-field-label" for="enSmartDate">${esc(t("entrySmartDate", { col: dateLabel }))}</label>
        <input class="en-input" id="enSmartDate" type="text" autocomplete="off" spellcheck="false" enterkeyhint="go" value="${esc(s.date)}">
      </div>
      <p class="en-warn" id="enSmartMsg" role="status" hidden></p>
      <button class="btn btn-xs ghost en-smart-toggle" id="enSmartOff" type="button">${esc(t("entrySmartAllFields"))}</button>
    </div>`;
    bar.innerHTML = `<span class="en-keys">${esc(t("entryKeysHint"))}</span>
      <button class="btn" id="enReviewBtn" type="button">${esc(walking && !ui.options.confirmStep ? t("entryQCommitNext") : (ui.options.confirmStep ? t("entryReview") : t("entryCommit")))}</button>`;
    stage.querySelector("#enChangeTarget")?.addEventListener("click", () => {
      ui.target = null; ui.fields = pinnedFields(); ui.notes = {}; ui.overwriteNotes = new Set(); ui.step = 1; render(); persistNow();
    });
    if (walking) wireQueueWalk(stage);
    bar.querySelector("#enReviewBtn").addEventListener("click", toReview);
    stage.querySelectorAll("[data-en-action]").forEach((b) => b.addEventListener("click", () => {
      s.action = b.dataset.enAction; render(); overlay.querySelector(`[data-en-action="${s.action}"]`)?.focus();
    }));
    stage.querySelector("#enSmartOff").addEventListener("click", () => { s.off = true; render(); });
    const inputs = [...stage.querySelectorAll("[data-en-smart-data], #enSmartDate")];
    inputs.forEach((input, i) => {
      const isDate = input.id === "enSmartDate";
      const k = isDate ? -1 : Number(input.dataset.enSmartData);
      const set = (v) => { if (isDate) s.date = v; else s.data[k] = v; const m = stage.querySelector("#enSmartMsg"); if (m) m.hidden = true; };
      if (!isDate) {
        let sug = null;
        input.addEventListener("focus", () => {
          if (sug || !(ui.options.suggestions || fieldDv(Number(input.dataset.col)))) return;
          sug = attachSuggest(input, Number(input.dataset.col), () => focusField(inputs, i + 1), set);
        });
        input.addEventListener("blur", () => setTimeout(() => {
          if (!sug || document.activeElement === input || (sug.isInteracting && sug.isInteracting())) return;
          const j = suggests.indexOf(sug);
          if (j >= 0) suggests.splice(j, 1);
          try { sug.destroy(); } catch { /* już zamknięta */ }
          sug = null;
        }, 0));
      }
      input.addEventListener("input", () => set(input.value));
      input.addEventListener("keydown", (e) => {
        if (e.key !== "Enter" || e.ctrlKey || e.metaKey) return;
        e.preventDefault();
        if (e.shiftKey) focusField(inputs, i - 1);
        else if (i === inputs.length - 1) toReview();
        else focusField(inputs, i + 1);
      });
    });
    // Zdaj: data już jest (dziś) — fokus na przycisku, bez klawiatury na telefonie; Wydaj: pierwsze pole danych
    const first = s.action === "issue" ? inputs[0] : (s.date ? bar.querySelector("#enReviewBtn") : stage.querySelector("#enSmartDate"));
    first?.focus();
  }

  function smartMsg(text) {
    const m = overlay.querySelector("#enSmartMsg");
    if (m) { m.textContent = text; m.hidden = !text; }
    live(text);
  }

  // Okno przy nietypowym wierszu (decyzja Mateusza 3): stan + wybór, nigdy automat.
  async function smartAsk(p, row) {
    const key = keyText(row);
    const state = cycleStateText(row, p);
    const cancel = { label: t("entryCancel"), value: null, kind: "secondary" };
    if (p.status === "open-prev") {
      const issue = smartName("issue");
      return appDialog({
        title: t("entrySmartPrevTitle"), message: t("entrySmartPrevMsg", { key, state }),
        actions: [
          { label: t("entrySmartCloseIssue", { date: ui.smart.date || "—", issue: issue.toLocaleLowerCase(currentLang === "en" ? "en" : "pl") }), value: "close-issue", kind: "primary" },
          { label: t("entrySmartIssueAnyway", { issue }), value: "issue-anyway", kind: "secondary" },
          { label: t("entrySmartNoteOnly"), value: "note", kind: "secondary" }, cancel],
      });
    }
    return appDialog({
      title: t("entrySmartCheckTitle"), message: t("entrySmartCheckMsg", { key, state }),
      actions: [{ label: t("entrySmartNoteOnly"), value: "note", kind: "primary" }, { label: t("entrySmartAllFields"), value: "fields", kind: "secondary" }, cancel],
    });
  }

  async function smartToReview() {
    const s = ui.smart;
    const row = swbEntry.rowAt(ui.target.rowIndex0);
    if (!row) return;
    s.resolve = null; // pytamy przy każdym podejściu — stan wiersza mógł się zmienić
    let p = swbEntry.smartPlan(row.rowIndex0, s.action, smartOpts(s));
    if (p.status === "open-prev" || p.status === "check") {
      const ans = await smartAsk(p, row);
      if (!ans) { overlay.querySelector("#enReviewBtn")?.focus(); return; }
      if (ans === "fields") { s.off = true; render(); return; }
      if (ans === "note") {
        const text = await appDialog({
          title: t("entrySmartNoteTitle"), message: t("entrySmartNoteMsg", { col: headers()[ui.keyCol], key: keyText(row) }),
          input: { label: t("entrySmartNoteLabel"), value: s.note || "" },
          actions: [{ label: t("entrySmartNoteSave"), value: true, kind: "primary" }, { label: t("entryCancel"), value: null, kind: "secondary" }],
        });
        if (!text || !String(text).trim()) { overlay.querySelector("#enReviewBtn")?.focus(); return; }
        s.note = String(text).trim();
        const n = smartNote(row.rowIndex0, s.note);
        ui.fields = {}; ui.notes = n.notes; ui.overwriteNotes = new Set(n.overwriteNotes); ui.overwrite = new Set();
        ui.smartLast = { action: s.action, note: true };
        reviewNow();
        return;
      }
      s.resolve = ans;
      p = swbEntry.smartPlan(row.rowIndex0, s.action, smartOpts(s));
    }
    if (p.status !== "ok") {
      const key = keyText(row);
      smartMsg(p.status === "not-out" ? t("entrySmartNotOut", { key, close: smartName("close") })
        : p.status === "full" ? t("entrySmartFullMsg", { key }) : p.status === "need-date" ? t("entrySmartNeedDate") : t("entrySmartCheckMsg", { key, state: cycleStateText(row, p) }));
      return;
    }
    ui.fields = { ...p.writes }; ui.notes = {}; ui.overwrite = new Set(); ui.overwriteNotes = new Set();
    ui.smartLast = { action: s.action, opts: smartOpts(s), writes: p.writes, cycle: p.cycle, closeCycle: p.closeCycle || null, note: false };
    reviewNow();
  }

  const focusField = (inputs, i) => { if (inputs[i]) { inputs[i].focus(); inputs[i].select?.(); } };

  function fieldDv(col) {
    const ref = ui.target && !ui.target.newRow ? ui.target.rowIndex0 : (swbEntry.lastDataRow() || {}).rowIndex0;
    if (!Number.isFinite(ref)) return null;
    try { return getCellDataValidation(currentSheetName, ref, currentStartCol + col); } catch { return null; }
  }

  function attachSuggest(input, col, next, onValue) {
    if (typeof createCellSuggestions !== "function") return null;
    const dv = fieldDv(col);
    const values = dv && dv.values.length ? dv.values : (ui.options.suggestions && typeof buildColumnSuggestions === "function" ? buildColumnSuggestions(col) : []);
    if (!values.length) return null;
    const api = createCellSuggestions(input, values, (v) => {
      input.value = v;
      if (onValue) onValue(v); else { ui.fields[col] = v; hint(col); }
      persistSoon(); next();
    },
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
    if (smartActive()) { smartToReview(); return; }
    ui.smartLast = null;
    reviewNow();
  }
  function reviewNow() {
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
    const sl = smartActive() && ui.smartLast;
    const smartLine = !sl ? "" : `<p class="en-smart-state" id="enSmartCycle">${esc(sl.note ? t("entrySmartNoteLine")
      : t("entrySmartCycleLine", { action: smartName(sl.action), n: sl.cycle }) + (sl.closeCycle ? t("entrySmartCloseLine", { n: sl.closeCycle }) : ""))}</p>`;
    stage.innerHTML = `<div class="tr-card en-card">
      <div class="en-target-line"><span class="en-target-name">${esc(head)}</span></div>
      ${smartLine}
      <h2 class="en-review-title">${esc(t("entryReviewTitle"))}</h2>
      ${p.changes.length || blocked.length ? `<ul class="en-changes">${items}${extra}</ul>` : `<p class="en-nothing">${esc(t("entryNothing"))}</p>`}
    </div>`;
    bar.innerHTML = `<button class="btn ghost" id="enBackBtn" type="button">${esc(t("entryBack"))}</button>
      <button class="btn" id="enCommitBtn" type="button" ${p.ok && p.changes.length ? "" : "disabled"}>${esc(ui.queue && ui.queue.mode === "walk" ? t("entryQCommitNext") : t("entryCommit"))}</button>`;
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
      const sl = smartActive() && ui.smartLast;
      if (sl && !sl.note) {
        const again = swbEntry.smartPlan(ui.target.rowIndex0, sl.action, sl.opts);
        if (again.status !== "ok" || JSON.stringify(again.writes) !== JSON.stringify(sl.writes)) { toast(t("entryStale"), "warning"); ui.step = 2; render(); return; }
      }
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
    if (ui.queue && ui.queue.mode === "fix") { endQueueFix(); drawPanel(); return; }
    if (ui.queue && ui.queue.mode === "walk") {
      const it = ui.queue.items[ui.queue.pos];
      if (it) { it.state = "done"; it.rowIndex0 = res.rowIndex0; }
      ui.queue.done = (ui.queue.done || 0) + 1;
      ui.notes = {}; ui.overwriteNotes = new Set(); ui.overwrite = new Set(); ui.target = null;
      if (ui.options.queueAutoNext) nextQueueItem(ui.queue.pos + 1);
      else { ui.step = "queue"; ui.queue.mode = null; render(); }
      persist({ draft: false });
      drawPanel();
      return;
    }
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

  // ── Kolejka ────────────────────────────────────────────────────────────────
  // Jedna kolejka, dwa sposoby: pola wypełnione na ekranie kolejki są WSPÓLNE (idą do wszystkich),
  // reszta jest OSOBNA dla każdej pozycji. „Naraz” zapisuje same wspólne w istniejących wierszach
  // (jedno Cofnij), „po kolei” prowadzi kartą za kartą ze wspólnymi już wpisanymi.
  function openQueue(list) {
    const resolved = C.resolveList(list.items, swbEntry.searchRows(ui.keyCol), ui.keyCol, { foldCache: swbEntry.foldCache() });
    ui.queue = {
      items: resolved.map((x) => ({
        text: x.text, status: x.status, rowIndex0: Number.isInteger(x.rowIndex0) ? x.rowIndex0 : null, options: x.options || null,
        choice: x.status === "ok" ? x.rowIndex0 : x.status === "missing" ? ui.options.queueMissing : x.status === "same-row" ? "skip" : null,
        state: "pending",
      })),
      pos: -1, mode: null, common: (() => { const c = pinnedFields(); delete c[ui.keyCol]; return c; })(), dupes: list.dupes || [], done: 0, ow: false,
      action: "fields", smart: { date: null, data: {}, note: "" },
    };
    ui.list = null; ui.query = ""; ui.target = null; ui.queueDraft = null;
    ui.step = "queue";
    render();
    persistNow();
  }

  // Kolejka z wyników zapytania: wiersze już znalezione — każda pozycja od razu „ok”.
  function openQueueRows(rowIdx) {
    const rows = rowIdx.map((r) => swbEntry.rowAt(r)).filter(Boolean);
    if (!rows.length) return;
    ui.queue = {
      items: rows.map((r) => ({ text: keyText(r) || t("entryRowN", { n: excelRow(r.rowIndex0) }), status: "ok", rowIndex0: r.rowIndex0, options: null, choice: r.rowIndex0, state: "pending" })),
      pos: -1, mode: null, common: (() => { const c = pinnedFields(); delete c[ui.keyCol]; return c; })(), dupes: [], done: 0, ow: false,
      action: "fields", smart: { date: null, data: {}, note: "" },
    };
    ui.list = null; ui.query = ""; ui.target = null; ui.queueDraft = null; ui.queryHits = null;
    ui.step = "queue";
    render();
    persistNow();
  }

  const activeItems = () => (ui.queue ? ui.queue.items.map((it, i) => ({ it, i })).filter(({ it }) => it.choice !== "skip" && it.choice !== null) : []);
  const commonFilled = () => !!ui.queue && Object.values(ui.queue.common || {}).some((v) => String(v ?? "").trim());

  function rowLabel(rowIndex0) {
    const row = swbEntry.rowAt(rowIndex0);
    const rows = swbEntry.candidates();
    const cols = C.distinguishingCols(headers(), headers().map((_, i) => rows.map((r) => r.values[i])), ui.keyCol, 1);
    const extra = row && cols.length ? ` · ${getDisplayValue(row, cols[0]) || "—"}` : "";
    return `${t("entryRowN", { n: excelRow(rowIndex0) })}${extra}`;
  }

  // ── Kolejka z akcją smart: każda pozycja dostaje decyzję dla SWOJEGO wiersza ──
  const qSmartOn = () => !!(ui.queue && ui.queue.action && ui.queue.action !== "fields" && smartSrc());
  function qSmart() { const q = ui.queue; q.smart = q.smart || { date: null, data: {}, note: "" }; if (q.smart.date === null) q.smart.date = smartDefaultDate(); return q.smart; }
  // { p, base, needs } dla pozycji z wierszem; needs = wymaga wyboru (poprzedni otwarty / do sprawdzenia)
  function queueDecision(it) {
    const q = ui.queue, sm = qSmart();
    if (!Number.isInteger(it.choice) || it.state === "done") return null;
    const base = swbEntry.smartPlan(it.choice, q.action, { date: sm.date, data: sm.data, roles: cfg().roles });
    const needs = base.status === "open-prev" || base.status === "check";
    // brak otwartego / wolnego cyklu: domyślnie pomijamy, ale można dopisać notatkę albo poprawić wiersz ręcznie
    const optional = base.status === "not-out" || base.status === "full";
    if (needs && it.resolve === undefined) it.resolve = base.status === "check" && cfg().anomalies === "skip" ? "skip" : null;
    if (optional && (it.resolve === undefined || it.resolve === null)) it.resolve = "skip";
    if (optional) return { p: base, base, needs: false, optional };
    if (!needs) return { p: base, base, needs };
    const r = it.resolve;
    const p = r && r !== "skip" && r !== "note" ? swbEntry.smartPlan(it.choice, q.action, { date: sm.date, data: sm.data, resolve: r, roles: cfg().roles }) : base;
    return { p, base, needs };
  }
  // pozycje do zapisu „naraz” (pola per wiersz) + czy coś czeka na decyzję
  function queueSmartItems() {
    const q = ui.queue, sm = qSmart();
    const items = [], meta = {};
    let unresolved = false, needNote = false;
    q.items.forEach((it) => {
      const d = queueDecision(it);
      if (!d) return;
      if (d.needs && it.resolve === null) { unresolved = true; return; }
      if ((d.needs || d.optional) && (it.resolve === "skip" || it.resolve === "fix")) return;
      if ((d.needs || d.optional) && it.resolve === "note") {
        if (!String(sm.note || "").trim()) { needNote = true; return; }
        const n = smartNote(it.choice, String(sm.note).trim());
        items.push({ rowIndex0: it.choice, fields: {}, notes: n.notes, overwriteNotes: n.overwriteNotes });
        meta[it.choice] = { note: true };
        return;
      }
      if (d.p.status !== "ok") return;
      items.push({ rowIndex0: it.choice, fields: { ...d.p.writes }, notes: {} });
      meta[it.choice] = { cycle: d.p.cycle, closeCycle: d.p.closeCycle || null, writes: d.p.writes };
    });
    return { items, meta, unresolved, needNote };
  }

  function queueSmartStatus(it, i) {
    const d = queueDecision(it);
    if (!d) return null;
    const row = swbEntry.rowAt(it.choice);
    const rowN = esc(t("entryRowN", { n: excelRow(it.choice) }));
    if (d.needs || d.optional) {
      const opts = d.base.status === "open-prev" ? ["close-issue", "issue-anyway", "note", "fix", "skip"] : ["note", "fix", "skip"];
      const label = (v) => (v === "close-issue" ? t("entrySmartCloseIssue", { date: qSmart().date || "—", issue: smartName("issue").toLocaleLowerCase(currentLang === "en" ? "en" : "pl") })
        : v === "issue-anyway" ? t("entrySmartIssueAnyway", { issue: smartName("issue") }) : v === "note" ? t("entrySmartNoteOnly")
          : v === "fix" ? t("entrySmartFix") : t("entryQSkipOpt"));
      const why = d.optional ? ` · ${t(d.base.status === "not-out" ? "entrySmartQNotOutShort" : "entrySmartQFullShort")}` : "";
      const after = it.resolve && it.resolve !== "skip" && it.resolve !== "note" && d.p.status === "ok" ? ` ${esc(t("entrySmartQOk", { col: headers()[smartCol(d.p, ui.queue.action)], n: d.p.cycle }))}` : "";
      return `<label class="en-q-st ${d.optional ? "is-miss" : "is-warn"}"><span data-en-q-smart="${i}">${rowN} · ${esc(cycleStateText(row, d.base))}${esc(why)}${after}</span>
        <select class="en-select" data-en-q-resolve="${i}">${d.optional ? "" : `<option value="" ${it.resolve === null ? "selected" : ""}>${esc(t("entrySmartQChoose"))}</option>`}
        ${opts.map((v) => `<option value="${v}" ${it.resolve === v ? "selected" : ""}>${esc(label(v))}</option>`).join("")}</select></label>`;
    }
    const p = d.p;
    if (p.status === "ok") return `<span class="en-q-st is-ok" data-en-q-smart="${i}">${rowN} ${esc(t("entrySmartQOk", { col: headers()[smartCol(p, ui.queue.action)], n: p.cycle }))}</span>`;
    const why = p.status === "not-out" ? t("entrySmartQNotOut") : p.status === "full" ? t("entrySmartQFull") : t("entrySmartQNeedDate");
    return `<span class="en-q-st is-miss" data-en-q-smart="${i}">${rowN} · ${esc(why)}</span>`;
  }

  // lista pozycji (rysowana też po zmianie daty/akcji — bez przebudowy pól, fokus zostaje)
  function drawQueueList() {
    const q = ui.queue;
    const ul = overlay.querySelector("#enQList");
    if (!q || !ul) return;
    const smart = qSmartOn();
    ul.innerHTML = q.items.map((it, i) => {
      let ctl = smart ? queueSmartStatus(it, i) : null;
      if (ctl) return `<li class="en-q-item" data-en-q="${i}"><span class="en-q-text">${esc(it.text)}</span>${ctl}</li>`;
      if (it.status === "ok") ctl = `<span class="en-q-st is-ok">${esc(t("entryQOk", { n: excelRow(it.rowIndex0) }))}</span>`;
      else if (it.status === "ambiguous") ctl = `<label class="en-q-st is-warn"><span class="visually-hidden">${esc(it.text)}: </span>${esc(t("entryQAmbiguous"))}
        <select class="en-select" data-en-q-choose="${i}"><option value="">${esc(t("entryQChoose"))}</option>
        ${(it.options || []).map((r) => `<option value="${r}" ${it.choice === r ? "selected" : ""}>${esc(rowLabel(r))}</option>`).join("")}
        <option value="skip" ${it.choice === "skip" ? "selected" : ""}>${esc(t("entryQSkipOpt"))}</option></select></label>`;
      else if (it.status === "missing") ctl = `<label class="en-q-st is-miss">${esc(t("entryQMissing"))}
        <select class="en-select" data-en-q-missing="${i}"><option value="skip" ${it.choice !== "new" ? "selected" : ""}>${esc(t("entryQSkipOpt"))}</option>
        <option value="new" ${it.choice === "new" ? "selected" : ""}>${esc(t("entryQNewOpt"))}</option></select></label>`;
      else ctl = `<span class="en-q-st is-miss">${esc(t("entryQSameRow"))}</span>`;
      return `<li class="en-q-item" data-en-q="${i}"><span class="en-q-text">${esc(it.text)}</span>${ctl}</li>`;
    }).join("");
    ul.querySelectorAll("[data-en-q-choose]").forEach((sel) => sel.addEventListener("change", () => {
      const it = q.items[Number(sel.dataset.enQChoose)];
      it.choice = sel.value === "" ? null : sel.value === "skip" ? "skip" : Number(sel.value);
      it.resolve = undefined;
      if (qSmartOn()) drawQueueList();
      updateQueueButtons(); persistSoon();
    }));
    ul.querySelectorAll("[data-en-q-missing]").forEach((sel) => sel.addEventListener("change", () => {
      q.items[Number(sel.dataset.enQMissing)].choice = sel.value === "new" ? "new" : "skip";
      updateQueueButtons(); persistSoon();
    }));
    ul.querySelectorAll("[data-en-q-resolve]").forEach((sel) => sel.addEventListener("change", () => {
      const i = Number(sel.dataset.enQResolve);
      if (sel.value === "fix") { startQueueFix(i); return; }
      q.items[i].resolve = sel.value === "" ? null : sel.value;
      drawQueueList();
      ul.querySelector(`[data-en-q-resolve="${i}"]`)?.focus();
      const note = overlay.querySelector("#enQNoteField");
      if (note) note.hidden = !q.items.some((it) => it.resolve === "note");
      updateQueueButtons(); persistSoon();
    }));
  }

  // „Popraw ręcznie”: zwykła karta wiersza (wszystkie pola), po zapisie albo „Wróć” — z powrotem do kolejki,
  // a pozycja liczy się od nowa (np. dopisana data „od” = cykl otwarty = gotowa do Zdaj)
  function startQueueFix(i) {
    const q = ui.queue, it = q.items[i];
    if (!it || !Number.isInteger(it.choice)) return;
    q.pos = i; q.mode = "fix";
    ui.target = { rowIndex0: it.choice };
    ui.fields = {}; ui.notes = {}; ui.overwrite = new Set(); ui.overwriteNotes = new Set();
    ui.smart = newSmart({ off: true }); ui.smartLast = null;
    ui.step = 2;
    render(); persistNow();
  }
  function endQueueFix() {
    const q = ui.queue;
    const it = q && q.items[q.pos];
    if (it) it.resolve = undefined;
    q.mode = null; ui.target = null; ui.fields = {}; ui.notes = {}; ui.overwrite = new Set(); ui.overwriteNotes = new Set();
    ui.step = "queue";
    render(); persistNow();
  }

  function renderQueue(stage, bar) {
    const q = ui.queue;
    const src = smartSrc();
    if (!src) q.action = "fields";
    const smart = qSmartOn();
    const seg = !src ? "" : `<div class="en-seg en-seg-3" role="group" aria-label="${esc(t("entrySmartAction"))}">
        ${["fields", "close", "issue"].map((a) => `<button class="en-seg-btn" type="button" data-en-q-action="${a}" aria-pressed="${q.action === a}">${esc(a === "fields" ? t("entrySmartFieldsAction") : smartName(a))}</button>`).join("")}</div>`;
    let commonHtml;
    if (smart) {
      const sm = qSmart();
      const dataPos = src.roles.map((r, p) => (r === "data" ? p : -1)).filter((p) => p >= 0);
      const dateLabel = src.labels[src.roles.indexOf(q.action === "close" ? "end" : "start")];
      commonHtml = (q.action === "issue" ? dataPos.map((p, i) => `<div class="tr-field en-field">
          <label class="tr-field-label" for="enQSD${i}">${esc(src.labels[p])}</label>
          <input class="en-input" id="enQSD${i}" data-en-q-smart-data="${i}" data-col="${src.first + p}" type="text" autocomplete="off" spellcheck="false" enterkeyhint="next" value="${esc(sm.data[i] ?? "")}"></div>`).join("") : "")
        + `<div class="tr-field en-field"><label class="tr-field-label" for="enQSmartDate">${esc(t("entrySmartDate", { col: dateLabel }))}</label>
          <input class="en-input" id="enQSmartDate" data-en-q-smart-date type="text" autocomplete="off" spellcheck="false" enterkeyhint="done" value="${esc(sm.date)}"></div>
        <div class="tr-field en-field" id="enQNoteField" ${q.items.some((it) => it.resolve === "note") ? "" : "hidden"}><label class="tr-field-label" for="enQNote">${esc(t("entrySmartQNote"))}</label>
          <input class="en-input" id="enQNote" type="text" autocomplete="off" value="${esc(sm.note || "")}"></div>`;
    } else {
      // bez kolumny-klucza: wspólny klucz nadałby wszystkim wierszom ten sam numer/nazwisko
      const cols = ui.fieldOrder.filter((c) => ui.selected.has(c) && !isFormulaCol(c) && c !== ui.keyCol);
      commonHtml = cols.map((c) => `<div class="tr-field en-field" data-en-qcol="${c}">
        <label class="tr-field-label" for="enQF${c}">${esc(headers()[c])}</label>
        <input class="en-input" id="enQF${c}" data-en-qf="${c}" type="text" autocomplete="off" spellcheck="false" enterkeyhint="next" value="${esc(q.common[c] ?? "")}">
      </div>`).join("");
    }
    stage.innerHTML = `<div class="tr-card en-card">
      <h2 class="en-review-title">${esc(tn("entryQTitle", q.items.length))}</h2>
      ${q.dupes && q.dupes.length ? `<p class="en-hint">${esc(t("entryQueueDupes", { list: q.dupes.join(", ") }))}</p>` : ""}
      ${seg}
      <ul class="en-q-list" id="enQList"></ul>
      <p class="en-warn" id="enQHint" hidden>${esc(t("entryQNeedChoice"))}</p>
      <h2 class="en-panel-title">${esc(t("entryQCommonTitle"))}</h2>
      ${commonHtml}
    </div>`;
    drawQueueList();
    stage.querySelectorAll("[data-en-q-action]").forEach((b) => b.addEventListener("click", () => {
      q.action = b.dataset.enQAction;
      q.items.forEach((it) => { it.resolve = undefined; });
      render(); persistSoon();
      overlay.querySelector(`[data-en-q-action="${q.action}"]`)?.focus();
    }));
    stage.querySelectorAll("[data-en-q-smart-data], #enQSmartDate, #enQNote").forEach((input) => {
      input.addEventListener("input", () => {
        const sm = qSmart();
        if (input.id === "enQSmartDate") sm.date = input.value;
        else if (input.id === "enQNote") sm.note = input.value;
        else sm.data[Number(input.dataset.enQSmartData)] = input.value;
        if (input.id !== "enQNote") drawQueueList();
        updateQueueButtons(); persistSoon();
      });
    });
    bar.innerHTML = `<button class="btn ghost" id="enQCancel" type="button">${esc(t("entryQCancel"))}</button>
      <button class="btn ghost en-q-walk" id="enQWalk" type="button"></button>
      <button class="btn" id="enQAll" type="button"></button>`;
    const inputs = [...stage.querySelectorAll("[data-en-qf]")];
    inputs.forEach((input, i) => {
      const col = Number(input.dataset.enQf);
      let sug = null;
      input.addEventListener("focus", () => {
        if (sug || !(ui.options.suggestions || fieldDv(col))) return;
        sug = attachSuggest(input, col, () => focusField(inputs, i + 1), (v) => { q.common[col] = v; updateQueueButtons(); });
      });
      input.addEventListener("blur", () => setTimeout(() => {
        if (!sug || document.activeElement === input || (sug.isInteracting && sug.isInteracting())) return;
        const k = suggests.indexOf(sug);
        if (k >= 0) suggests.splice(k, 1);
        try { sug.destroy(); } catch { /* już zamknięta */ }
        sug = null;
      }, 0));
      input.addEventListener("input", () => { q.common[col] = input.value; updateQueueButtons(); persistSoon(); });
      input.addEventListener("keydown", (e) => { if (e.key === "Enter" && !e.ctrlKey && !e.metaKey) { e.preventDefault(); focusField(inputs, e.shiftKey ? i - 1 : i + 1); } });
    });
    bar.querySelector("#enQCancel").addEventListener("click", () => finishQueue());
    bar.querySelector("#enQWalk").addEventListener("click", () => { q.mode = "walk"; nextQueueItem(0); });
    bar.querySelector("#enQAll").addEventListener("click", () => { q.ow = false; ui.step = "queueReview"; render(); });
    updateQueueButtons();
  }

  function updateQueueButtons() {
    const q = ui.queue;
    const walk = overlay.querySelector("#enQWalk");
    const all = overlay.querySelector("#enQAll");
    if (!q || !walk || !all) return;
    const smart = qSmartOn();
    const si = smart ? queueSmartItems() : null;
    const unresolved = q.items.some((it) => it.choice === null) || !!(si && si.unresolved);
    const k = activeItems().filter(({ it }) => it.state !== "done").length;
    const m = si ? si.items.length : q.items.filter((it) => Number.isInteger(it.choice) && it.state !== "done").length;
    walk.textContent = t("entryQWalk", { n: k });
    all.textContent = t("entryQAll", { n: m });
    walk.disabled = unresolved || k === 0;
    all.disabled = unresolved || m === 0 || (!smart && !commonFilled()) || !!(si && si.needNote);
    const hint = overlay.querySelector("#enQHint");
    const why = unresolved ? t(q.items.some((it) => it.choice === null) ? "entryQNeedChoice" : "entrySmartQNeedChoice") : si && si.needNote ? t("entrySmartQNeedNote") : (!smart && m > 0 && !commonFilled() ? t("entryQNeedCommon") : "");
    if (hint) { hint.textContent = why; hint.hidden = !why; }
  }

  function queueProgressHtml() {
    const q = ui.queue;
    const act = activeItems();
    const at = act.findIndex(({ i }) => i === q.pos);
    const chips = q.items.map((it, i) => {
      const cls = i === q.pos ? " is-current" : it.state === "done" ? " is-done" : it.choice === "skip" || it.state === "skipped" ? " is-skip" : "";
      return `<button class="en-q-chip${cls}" type="button" data-en-q-jump="${i}" ${it.choice === "skip" || it.choice === null ? "disabled" : ""} aria-current="${i === q.pos ? "step" : "false"}">${esc(it.text)}</button>`;
    }).join("");
    return `<div class="en-q-progress"><span class="en-q-pos" id="enQProgress">${esc(t("entryQPos", { i: at + 1, n: act.length }))}</span>
      <div class="en-q-chips" role="list">${chips}</div></div>`;
  }

  function wireQueueWalk(stage) {
    stage.querySelector("#enQBack")?.addEventListener("click", () => { ui.queue.mode = null; ui.target = null; ui.step = "queue"; render(); persistNow(); });
    stage.querySelector("#enQSkip")?.addEventListener("click", () => {
      const it = ui.queue.items[ui.queue.pos];
      if (it && it.state !== "done") it.state = "skipped";
      nextQueueItem(ui.queue.pos + 1);
    });
    stage.querySelectorAll("[data-en-q-jump]").forEach((b) => b.addEventListener("click", () => startQueueItem(Number(b.dataset.enQJump))));
    stage.querySelector(".en-q-chip.is-current")?.scrollIntoView?.({ block: "nearest", inline: "center" });
  }

  function nextQueueItem(from) {
    const q = ui.queue;
    for (let i = from; i < q.items.length; i++) {
      const it = q.items[i];
      if (it.choice !== "skip" && it.choice !== null && it.state === "pending") { startQueueItem(i); return; }
    }
    finishQueue();
  }

  function startQueueItem(i) {
    const q = ui.queue;
    const it = q.items[i];
    if (!it) return;
    q.pos = i;
    q.mode = "walk";
    ui.target = it.choice === "new" ? { newRow: true } : { rowIndex0: it.choice };
    ui.fields = { ...q.common };
    if (it.choice === "new") ui.fields[ui.keyCol] = it.text;
    ui.notes = {}; ui.overwrite = new Set(); ui.overwriteNotes = new Set();
    const sm = q.smart || {};
    ui.smart = q.action && q.action !== "fields" ? newSmart({ action: q.action, date: sm.date ?? null, data: { ...(sm.data || {}) } }) : newSmart({ off: true });
    ui.smartLast = null;
    ui.step = 2;
    render();
    persistNow();
  }

  function finishQueue() {
    const q = ui.queue;
    if (q && (q.done || 0) > 0) toast(t("entryQSummary", { done: q.done, skipped: q.items.length - q.done }), "success");
    ui.queue = null; ui.queueDraft = null; ui.target = null; ui.fields = pinnedFields(); ui.notes = {}; ui.query = "";
    ui.step = 1;
    render();
    persist({ draft: false });
  }

  function queueTargets() { return ui.queue.items.filter((it) => Number.isInteger(it.choice) && it.state !== "done").map((it) => it.choice); }

  // wejście planMany/commitMany kolejki: przy akcji smart pola per wiersz, inaczej wspólne
  function queueInput() {
    const q = ui.queue;
    if (!qSmartOn()) return { input: { targets: queueTargets(), fields: q.common, overwriteAll: q.ow }, meta: null };
    const si = queueSmartItems();
    return { input: { items: si.items }, meta: si.meta };
  }
  function renderQueueReview(stage, bar) {
    const q = ui.queue;
    const qi = queueInput();
    const targets = qi.input.items ? qi.input.items.map((x) => x.rowIndex0) : qi.input.targets;
    const p = swbEntry.planMany(qi.input);
    ui.queuePlan = p;
    ui.queueMeta = qi.meta;
    const rowsHtml = p.rows.map((r) => {
      const row = swbEntry.rowAt(r.rowIndex0);
      const changes = r.plan.changes.map((c) => `<div class="en-diff en-q-diff"><span class="tr-field-label">${esc(headers()[c.col])}</span>
        <span class="en-before${c.before ? "" : " is-empty"}">${esc(c.before || t("entryEmptyVal"))}</span><span aria-hidden="true">→</span><span class="en-after">${esc(c.after)}</span></div>`).join("");
      const errs = r.plan.errors.filter((e) => !GATE.has(e.reason)).map((e) => `<div class="en-err" role="alert">${esc(headers()[e.col] || "")}: ${esc(t(ERR[e.reason] || "entryErrNotReal"))}</div>`).join("");
      const ow = r.plan.changes.some((c) => c.overwrite);
      const mt = qi.meta && qi.meta[r.rowIndex0];
      const cyc = !mt ? "" : `<div class="en-smart-state">${esc(mt.note ? t("entrySmartNoteLine") : t("entrySmartCycleLine", { action: smartName(q.action), n: mt.cycle }) + (mt.closeCycle ? t("entrySmartCloseLine", { n: mt.closeCycle }) : ""))}</div>`;
      return `<li class="en-change en-q-row${ow ? " is-overwrite" : ""}"><div class="en-target-name">${esc(t("entryTargetExisting", { n: excelRow(r.rowIndex0), key: keyText(row) }))}</div>${cyc}
        ${changes || `<div class="en-hint">${esc(t("entryQNoChange"))}</div>`}${errs}</li>`;
    }).join("");
    const inSet = new Set(targets);
    const skipped = q.items.filter((it) => it.state !== "done" && !(Number.isInteger(it.choice) && inSet.has(it.choice))).map((it) => it.text);
    const owCount = p.overwriteCount;
    stage.innerHTML = `<div class="tr-card en-card">
      <h2 class="en-review-title">${esc(tn("entryQReviewTitle", targets.length))}</h2>
      ${skipped.length ? `<p class="en-hint">${esc(t("entryQSkippedNote", { list: skipped.join(", ") }))}</p>` : ""}
      ${owCount ? `<label class="en-ow"><input type="checkbox" id="enQOverwrite" ${q.ow ? "checked" : ""}> ${esc(tn("entryQOverwrite", owCount))}</label>` : ""}
      <ul class="en-changes">${rowsHtml}</ul>
    </div>`;
    bar.innerHTML = `<button class="btn ghost" id="enBackBtn" type="button">${esc(t("entryBack"))}</button>
      <button class="btn" id="enCommitBtn" type="button" ${p.ok ? "" : "disabled"}>${esc(t("entryCommit"))}</button>`;
    bar.querySelector("#enBackBtn").addEventListener("click", () => { ui.step = "queue"; render(); });
    bar.querySelector("#enCommitBtn").addEventListener("click", doCommitQueue);
    stage.querySelector("#enQOverwrite")?.addEventListener("change", (e) => {
      q.ow = e.target.checked;
      const pp = swbEntry.planMany({ ...queueInput().input, overwriteAll: q.ow });
      bar.querySelector("#enCommitBtn").disabled = !pp.ok;
    });
    (bar.querySelector("#enCommitBtn:not([disabled])") || stage.querySelector("#enQOverwrite") || bar.querySelector("#enBackBtn")).focus();
  }

  async function doCommitQueue() {
    if (committing) return;
    committing = true;
    try {
      const q = ui.queue;
      const qi = queueInput();
      // decyzje smart liczone jeszcze raz na świeżych danych — inne niż w podglądzie = „stale”, nic nie piszemy
      if (qi.meta && JSON.stringify(qi.meta) !== JSON.stringify(ui.queueMeta)) { toast(t("entryStale"), "warning"); render(); return; }
      const res = await swbEntry.commitMany({ ...qi.input, overwriteAll: q.ow, expect: ui.queuePlan ? ui.queuePlan.expect : null });
      if (!res.ok && res.reason !== "write-failed") {
        if (res.reason === "stale") toast(t("entryStale"), "warning");
        else if (res.reason === "guard") toast(t("entryNeedEdit"), "warning");
        render();
        return;
      }
      if (res.reason === "write-failed") toast(t("entryWriteFailed", { n: res.failed.map((f) => excelRow(f.rowIndex0)).join(", ") }), "error");
      res.rows.forEach((r) => ui.sessionLog.unshift({ key: keyText(swbEntry.rowAt(r.rowIndex0)), rowIndex0: r.rowIndex0, at: Date.now() }));
      ui.profiles.clear();
      const same = (qi.input.items ? qi.input.items.length : queueTargets().length) - res.rows.length; // wiersze, które już miały te wartości
      toast(same > 0 ? tn("entryQDoneAllSame", res.rows.length, { same }) : tn("entryQDoneAll", res.rows.length), "success", {
        action: t("entryDoneUndo"),
        onAction: () => { if (swbEntry.undoCard()) toast(t("entryUndone"), "info"); else toast(t("entryUndoGone"), "info"); },
      });
      ui.queue = null; ui.queueDraft = null; ui.target = null; ui.fields = pinnedFields(); ui.query = "";
      ui.step = 1;
      persist({ draft: false });
      render();
      drawPanel();
    } finally { committing = false; }
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
    if (ui.scope !== scopeKey()) { ui.target = null; ui.fields = {}; ui.step = 1; ui.queue = null; ui.sessionLog = ui.scope ? [] : ui.sessionLog; ui.scope = scopeKey(); }
    restore();
    const rows = swbEntry.candidates();
    // arkusz z cyklami: klucz to kolumna PRZED cyklami (Nr.), nie osoba z któregoś cyklu
    const cyc = smartSrc();
    const keyCols = headers().slice(0, cyc && cyc.first > 0 ? cyc.first : headers().length);
    ui.keyCol = Number.isInteger(ui.keyColSaved) ? ui.keyColSaved : C.pickKeyColumn(keyCols, keyCols.map((_, i) => rows.map((r) => r.values[i])));
    ui.open = true;
    ui.step = ui.queue ? (ui.queue.mode === "walk" && ui.target ? ui.step : "queue") : (ui.target ? ui.step : 1);
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
      else if (ui.step === "queueReview") { const b = overlay.querySelector("#enCommitBtn"); if (b && !b.disabled) doCommitQueue(); }
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

  swbEntry.countWords = countWords;
  swbEntry.state = () => ({
    queue: ui.queue ? { items: ui.queue.items.map((x) => ({ text: x.text, status: x.status, choice: x.choice, state: x.state })), pos: ui.queue.pos, mode: ui.queue.mode, common: { ...ui.queue.common } } : null,
    open: ui.open, step: ui.step, target: ui.target ? { ...ui.target } : null, fields: { ...ui.fields },
    pinned: [...ui.pinned], keyCol: ui.keyCol, options: { ...ui.options }, sessionLog: ui.sessionLog.map((x) => ({ ...x })),
    fieldOrder: ui.fieldOrder.slice(), selected: [...ui.selected], notes: { ...ui.notes },
  });
  // Wejście programowe kolejki (przyszły skan stosu kart): lista tekstów → ekran kolejki.
  swbEntry.fillQueue = (texts) => {
    const seen = new Set();
    const items = [];
    const dupes = [];
    (texts || []).map((x) => String(x ?? "").trim()).filter(Boolean).forEach((x) => { const k = C.fold(x); if (seen.has(k)) dupes.push(x); else { seen.add(k); items.push(x); } });
    if (!items.length || !ui.open) return false;
    openQueue({ items, dupes });
    return true;
  };
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
