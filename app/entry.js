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

  const api = { candidates, plan, commit, undoCard, fill, sheetCell, rowAt, lastDataRow, ready, previewText, _filled: () => filled };
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

  const DEFAULT_OPTIONS = { fuzzyTarget: true, typoHints: true, typeHints: true, suggestions: true, duplicateGuard: true, confirmStep: true };
  const ui = {
    open: false, step: 1, keyCol: 0, query: "", targets: [], active: -1,
    target: null, fields: {}, overwrite: new Set(), pinned: new Set(), font: 2,
    options: { ...DEFAULT_OPTIONS }, profiles: new Map(), lastPlan: null, sessionLog: [], returnFocus: null,
  };
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
          <button class="btn btn-sm ghost" id="enFontBtn" type="button" aria-label="${esc(t("entryFont"))}">A</button>
          <button class="modal-close" id="enCloseBtn" type="button" aria-label="${esc(t("entryClose"))}">${ICON_X}</button>
        </div>
      </div>
      <div class="en-live visually-hidden" id="enLive" aria-live="polite"></div>
      <div class="tr-stage-wrap en-stage-wrap"><div class="tr-stage en-stage" id="enStage"></div></div>
      <div class="tr-bar en-bar" id="enBar"></div>`;
    overlay.querySelector("#enCloseBtn").addEventListener("click", close);
    overlay.querySelector("#enFontBtn").addEventListener("click", () => {
      ui.font = ui.font % 3 + 1;
      overlay.dataset.font = String(ui.font);
    });
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
        <div class="en-hint">${esc(t("entryTargetHint", { col: headers()[ui.keyCol] || "" }))}</div>
        <div class="en-candidates" id="enCandidates" role="listbox" aria-label="${esc(t("entryTargetLabel"))}"></div>
      </div>`;
    bar.innerHTML = `<span class="en-keys">${esc(t("entryKeysHint"))}</span>`;
    const input = stage.querySelector("#enTargetInput");
    input.addEventListener("input", () => { ui.query = input.value; search(); });
    input.addEventListener("keydown", onTargetKey);
    search();
    input.focus();
  }

  function search() {
    const rows = swbEntry.candidates();
    ui.targets = C.findTargets(rows, ui.keyCol, ui.query, { fuzzy: ui.options.fuzzyTarget });
    // jeden dokładny kandydat = podświetlony (Enter go wybiera); kilka albo literówki — wybiera człowiek
    const exact = ui.targets.filter((x) => x.exact);
    ui.active = exact.length === 1 && ui.targets[0] && ui.targets[0].exact ? 0 : -1;
    drawCandidates();
  }

  function drawCandidates() {
    const box = overlay.querySelector("#enCandidates");
    const input = overlay.querySelector("#enTargetInput");
    if (!box || !input) return;
    const rows = swbEntry.candidates();
    const cols = C.distinguishingCols(headers(), headers().map((_, i) => rows.map((r) => r.values[i])), ui.keyCol, 2);
    const q = ui.query.trim();
    const items = ui.targets.map((tg, i) => {
      const row = swbEntry.rowAt(tg.rowIndex0);
      const extra = cols.map((c) => `${esc(headers()[c])}: ${esc(getDisplayValue(row, c) || "—")}`).join(" · ");
      return `<div class="en-cand" role="option" id="enCand${i}" data-en-row="${tg.rowIndex0}" aria-selected="${i === ui.active}">
        <span class="en-cand-key">${esc(keyText(row))}</span>
        <span class="en-cand-extra">${extra}</span>
        <span class="en-cand-row">${esc(t("entryRowN", { n: excelRow(tg.rowIndex0) }))}</span></div>`;
    });
    const newIdx = ui.targets.length;
    if (q) items.push(`<div class="en-cand en-cand-new" role="option" id="enCand${newIdx}" data-en-new aria-selected="${ui.active === newIdx}">${esc(t("entryNewRowFor", { text: q }))}</div>`);
    box.innerHTML = items.join("");
    input.setAttribute("aria-expanded", items.length ? "true" : "false");
    if (ui.active >= 0) input.setAttribute("aria-activedescendant", `enCand${ui.active}`); else input.removeAttribute("aria-activedescendant");
    box.querySelectorAll(".en-cand").forEach((el, i) => {
      el.addEventListener("pointerdown", (e) => e.preventDefault()); // fokus zostaje w polu
      el.addEventListener("click", () => { ui.active = i; choose(); });
    });
    if (q && !ui.targets.length) live(t("entryNoMatch"));
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
      else if (ui.query.trim() && !ui.targets.length) { ui.active = 0; choose(); }
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
    if (ui.options.duplicateGuard && C.isDuplicateRisk(ui.targets)) {
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
    if (target.newRow && ui.query.trim()) ui.fields[ui.keyCol] = ui.query.trim();
    ui.overwrite = new Set();
    ui.step = 2;
    render();
  }

  // ── Krok 2: pola ──
  function renderFields(stage, bar) {
    const row = ui.target.newRow ? null : swbEntry.rowAt(ui.target.rowIndex0);
    const head = ui.target.newRow ? t("entryNewRow") : t("entryTargetExisting", { n: excelRow(ui.target.rowIndex0), key: keyText(row) });
    const fieldsHtml = realCols().map((col) => {
      const formula = isFormulaCol(col);
      const now = row ? String(getDisplayValue(row, col) ?? "") : "";
      const val = formula ? now : (ui.fields[col] ?? "");
      const pinned = ui.pinned.has(col);
      return `<div class="tr-field en-field${formula ? " is-formula" : ""}" data-en-col="${col}">
        <div class="en-field-head">
          <label class="tr-field-label" for="enF${col}">${esc(headers()[col])}</label>
          ${formula ? `<span class="en-tag">${esc(t("entryFormulaField"))}</span>` : `<button class="en-pin" type="button" aria-pressed="${pinned}" aria-label="${esc(pinned ? t("entryUnpin") : t("entryPin"))}" title="${esc(pinned ? t("entryUnpin") : t("entryPin"))}">${ICON_PIN}</button>`}
        </div>
        <input class="en-input" id="enF${col}" type="text" autocomplete="off" spellcheck="false" enterkeyhint="next"
          ${formula ? "readonly tabindex=\"-1\"" : ""} value="${esc(val)}" placeholder="${esc(!formula && now ? now : "")}"
          aria-describedby="enP${col}">
        <div class="en-preview" id="enP${col}" aria-live="polite"></div>
      </div>`;
    }).join("");
    stage.innerHTML = `<div class="tr-card en-card">
      <div class="en-target-line"><span class="en-target-name">${esc(head)}</span>
        <button class="btn btn-xs ghost" id="enChangeTarget" type="button">${esc(t("entryChangeTarget"))}</button></div>
      ${fieldsHtml}</div>`;
    bar.innerHTML = `<span class="en-keys">${esc(t("entryKeysHint"))}</span>
      <button class="btn" id="enReviewBtn" type="button">${esc(ui.options.confirmStep ? t("entryReview") : t("entryCommit"))}</button>`;
    stage.querySelector("#enChangeTarget").addEventListener("click", () => { ui.step = 1; render(); });
    bar.querySelector("#enReviewBtn").addEventListener("click", toReview);
    const inputs = [...stage.querySelectorAll(".en-field:not(.is-formula) input")];
    inputs.forEach((input, i) => {
      const col = Number(input.closest("[data-en-col]").dataset.enCol);
      if (ui.options.suggestions || (fieldDv(col))) attachSuggest(input, col, () => focusField(inputs, i + 1));
      input.addEventListener("input", () => { ui.fields[col] = input.value; hint(col); });
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
    if (typeof createCellSuggestions !== "function") return;
    const dv = fieldDv(col);
    const values = dv && dv.values.length ? dv.values : (ui.options.suggestions && typeof buildColumnSuggestions === "function" ? buildColumnSuggestions(col) : []);
    if (!values.length) return;
    suggests.push(createCellSuggestions(input, values, (v) => { input.value = v; ui.fields[col] = v; hint(col); next(); },
      dv && dv.values.length ? {} : { variant: "soft", minChars: 1 }));
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
      const p = profile(col);
      if (p.type === "date" && parsed.type !== "date") {
        // „12.03” z kartki to w kolumnie dat prawie na pewno dzień i miesiąc — proponujemy pełną datę
        const m = /^(\d{1,2})[.,](\d{1,2})$/.exec(raw);
        const full = m ? `${m[1]}.${m[2]}.${new Date().getFullYear()}` : null;
        const fp = full ? parseCellInput(full, refIdx, col) : null;
        const fix = fp && fp.type === "date" ? ` <button class="btn btn-xs ghost en-date-fix" type="button" data-v="${esc(full)}">${esc(t("entryTypoFix", { v: full }))}</button>` : "";
        parts.push(`<span class="en-warn">${esc(t("entryTypeWarnDate"))}${fix}</span>`);
      }
      if (p.type === "number" && parsed.type !== "number") parts.push(`<span class="en-warn">${esc(t("entryTypeWarnNumber"))}</span>`);
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
    return { target: ui.target, fields: { ...ui.fields }, overwrite: [...ui.overwrite] };
  }

  function toReview() {
    ui.lastPlan = swbEntry.plan(planInput());
    const needsLook = ui.lastPlan.changes.some((c) => c.overwrite) || !ui.lastPlan.ok || !ui.lastPlan.changes.length;
    if (!ui.options.confirmStep && !needsLook) { doCommit(); return; }
    ui.step = 3;
    render();
  }

  // ── Krok 3: podgląd ──
  const ERR = { formula: "entryErrFormula", "dv-stop": "entryErrDvStop", "overwrite-not-allowed": "entryErrOverwrite", "not-real-col": "entryErrNotReal" };
  function renderReview(stage, bar) {
    const p = ui.lastPlan = swbEntry.plan(planInput());
    const row = ui.target.newRow ? null : swbEntry.rowAt(ui.target.rowIndex0);
    const head = ui.target.newRow ? t("entryNewRow") : t("entryTargetExisting", { n: excelRow(ui.target.rowIndex0), key: keyText(row) });
    const errsOf = (col) => p.errors.filter((e) => e.col === col && e.reason !== "overwrite-not-allowed");
    const items = p.changes.map((c) => {
      const errs = errsOf(c.col).map((e) => `<div class="en-err" role="alert">${esc(t(ERR[e.reason] || "entryErrNotReal"))}</div>`).join("");
      const warns = p.warnings.filter((w) => w.col === c.col).map(() => `<div class="en-warn">${esc(t("entryWarnDv"))}</div>`).join("");
      const ow = c.overwrite ? `<label class="en-ow"><input type="checkbox" data-en-overwrite="${c.col}" ${ui.overwrite.has(c.col) ? "checked" : ""}> ${esc(t("entryOverwriteOk"))}</label>` : "";
      return `<li class="en-change${c.overwrite ? " is-overwrite" : ""}">
        <div class="tr-field-label">${esc(headers()[c.col])}</div>
        <div class="en-diff"><span class="en-before">${esc(c.before || t("entryEmptyVal"))}</span><span aria-hidden="true">→</span><span class="en-after">${esc(c.after)}</span></div>
        ${ow}${errs}${warns}</li>`;
    }).join("");
    const blocked = p.errors.filter((e) => e.reason !== "overwrite-not-allowed" && !p.changes.some((c) => c.col === e.col));
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
    stage.querySelectorAll("[data-en-overwrite]").forEach((cb) => cb.addEventListener("change", () => {
      const col = Number(cb.dataset.enOverwrite);
      if (cb.checked) ui.overwrite.add(col); else ui.overwrite.delete(col);
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
      if (ui.lastPlan) input.expect = Object.fromEntries(ui.lastPlan.changes.map((c) => [c.col, c.before]));
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
    ui.target = null;
    ui.query = "";
    ui.overwrite = new Set();
    ui.step = 1;
    if (ui.open) render();
    if (typeof onCardDone === "function") onCardDone();
  }
  let onCardDone = null;

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
    const rows = swbEntry.candidates();
    if (!Number.isInteger(ui.keyColOverride)) ui.keyCol = C.pickKeyColumn(headers(), headers().map((_, i) => rows.map((r) => r.values[i])));
    ui.open = true;
    ui.step = ui.target ? ui.step : 1;
    overlay.dataset.font = String(ui.font);
    shell();
    overlay.classList.remove("hidden");
    document.body.classList.add("en-active");
    setInert(true);
    requestWake();
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
    if (typeof onClose === "function") onClose();
    if (back && document.contains(back) && back.getClientRects().length) back.focus();
    else document.getElementById("moreMenuBtn")?.focus();
  }
  let onClose = null;

  // Klawisze: Esc i Ctrl/⌘+Enter tu; reszta zostaje w karcie (nie dochodzi do tabeli i Cofnij).
  document.addEventListener("keydown", (e) => {
    if (!ui.open) return;
    if (document.querySelector(".modal.app-dialog:not(.hidden)")) return; // okno nad kartą obsługuje się samo
    if (e.key === "Escape") {
      e.preventDefault(); e.stopPropagation();
      // najpierw zamyka się lista podpowiedzi pod polem, dopiero drugie Esc — karta
      const open = document.querySelectorAll(".cell-suggest:not(.hidden)");
      if (open.length) { open.forEach((b) => b.classList.add("hidden")); return; }
      close(); return;
    }
    if (e.key === "Enter" && (e.ctrlKey || e.metaKey)) {
      e.preventDefault(); e.stopPropagation();
      if (ui.step === 2) toReview();
      else if (ui.step === 3) { const b = overlay.querySelector("#enCommitBtn"); if (b && !b.disabled) doCommit(); }
    }
  }, true);
  overlay.addEventListener("keydown", (e) => e.stopPropagation());
  document.addEventListener("visibilitychange", () => { if (ui.open && document.visibilityState === "visible" && !wakeLock) requestWake(); });
  if (btn) btn.addEventListener("click", () => {
    open();
  });

  // Stan i haki dla Task 4 (pamięć) i testów.
  swbEntry.open = open;
  swbEntry.close = close;
  swbEntry.ui = ui;
  swbEntry.hooks = {
    set onCardDone(fn) { onCardDone = fn; },
    set onClose(fn) { onClose = fn; },
    render: () => { if (ui.open) render(); },
  };
  swbEntry.state = () => ({
    open: ui.open, step: ui.step, target: ui.target ? { ...ui.target } : null, fields: { ...ui.fields },
    pinned: [...ui.pinned], keyCol: ui.keyCol, options: { ...ui.options }, sessionLog: ui.sessionLog.map((x) => ({ ...x })),
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
