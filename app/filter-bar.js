// filter-bar.js — dwie rzeczy, które skracają drogę do filtra:
//
//  1. Pasek „co teraz filtruje" (#activeFilters): pod paskiem narzędzi jedna linia
//     pigułek — każdy aktywny filtr z ×, liczba „Widać X z Y" i „Wyczyść". Bez niego
//     przy 12 wierszach z 3000 trzeba było otwierać panel, żeby zrozumieć, co je ukrywa.
//     Czyta STAN KONTROLEK w chwili applyFilters (migawka lastAppliedFilters), więc
//     pokazuje to, co naprawdę działa, a nie to, co ktoś wpisał w panelu i nie zatwierdził.
//
//  2. Menu komórki (#cellMenu): prawy klik, przytrzymanie palcem albo klawisz menu /
//     Shift+F10 na komórce → „Pokaż tylko takie" / „Ukryj takie" / „Kopiuj".
//     Filtr trafia do szybkiego szukania jako zwykła składnia („Status:="Anulowana"",
//     „Status:!="Anulowana""), więc widać go w polu, da się go poprawić ręcznie,
//     a kolejne kliknięcia DOKLEJAJĄ się przez && (jak „filtr według zaznaczenia" w Accessie).

// ── 1. Pasek aktywnych filtrów ──────────────────────────────────────────────────

const activeFiltersEl = document.getElementById("activeFilters");

function afSelectText(select) {
  const opt = select && select.selectedOptions && select.selectedOptions[0];
  return opt ? String(opt.textContent || "").trim() : "";
}

function afShort(text, max = 40) {
  const s = String(text || "").replace(/\s+/g, " ").trim();
  return s.length > max ? `${s.slice(0, max - 1)}…` : s;
}

// Zapytanie po ludzku: „Osoba:="Anna" && Projekt:!="Beta"" → „Osoba = Anna · Projekt ≠ Beta".
// Tłumaczy tylko proste warunki kolumnowe (=, !=, puste); resztę zostawia jak w polu.
// null = nie ma czego upiększać (zwykły tekst zostaje w cudzysłowie „…”).
function afPrettyQuery(query) {
  if (!/:!?=/.test(query) || /[{}]|\|\|/.test(query)) return null;
  const parts = query.split("&&").map((p) => p.trim()).filter(Boolean);
  let changed = false;
  const out = parts.map((part) => {
    const m = part.match(/^(.+?):(!?)=(?!<<|>>)\s*(.*)$/);
    if (!m) return part;
    changed = true;
    const val = unquoteTerm(m[3]);
    const shown = val ? afShort(val, 28) : t("cellMenuEmptyValue");
    return `${m[1].trim()} ${m[2] ? "≠" : "="} ${shown}`;
  });
  return changed ? out.join(" · ") : null;
}

// Opis jednego filtra tekstowego: „„faktura”” + NIE / tryb / puste / kolumny.
function afTextFilterChip(queryEl, modeEl, negateEl, emptyEl, colSet) {
  const query = String(queryEl?.value || "").trim();
  const emptyMode = emptyEl ? getNormalizedSelectValue(emptyEl) : "all";
  if (!query && emptyMode === "all") return null;
  const parts = [];
  if (negateEl?.checked) parts.push({ text: t("activeFiltersNot"), tag: "not" });
  if (query) {
    const pretty = afPrettyQuery(query);
    parts.push({ text: pretty || `„${afShort(query)}”`, tag: "query" });
  }
  const mode = modeEl ? getNormalizedSelectValue(modeEl) : "contains";
  if (query && mode && mode !== "contains") parts.push({ text: afSelectText(modeEl), tag: "meta" });
  if (emptyMode !== "all") parts.push({ text: afSelectText(emptyEl), tag: "meta" });
  if (colSet && colSet.size) parts.push({ text: t("activeFiltersCols", { n: colSet.size }), tag: "meta" });
  return parts;
}

function afDateChip() {
  const mode = getNormalizedSelectValue(dateModeEl);
  const emptyMode = getNormalizedSelectValue(dateEmptyModeEl);
  const parts = [];
  if (dateNegateEl.checked) parts.push({ text: t("activeFiltersNot"), tag: "not" });
  if (mode === "last_n_days") {
    const n = Math.max(1, parseInt(lastDaysEl.value || "30", 10));
    parts.push({ text: t("activeFiltersLastDays", { n }), tag: "query" });
  } else {
    const from = dateFromEl.value.trim();
    const to = dateToEl.value.trim();
    if (from && mode !== "before") parts.push({ text: t("activeFiltersFrom", { d: from }), tag: "query" });
    if (to && mode !== "after") parts.push({ text: t("activeFiltersTo", { d: to }), tag: "query" });
  }
  if (emptyMode !== "all") parts.push({ text: afSelectText(dateEmptyModeEl), tag: "meta" });
  const hasRange = parts.some((p) => p.tag === "query");
  if (!hasRange && emptyMode === "all") return null;
  return parts;
}

// Zdjęcie jednego filtra = wyzerowanie JEGO kontrolek + ten sam przebieg co „Filtruj".
function afRerun() {
  filtersCommitted = true;
  applyFilters();
  sortRows();
  scheduleViewRefresh({ table: true, analyses: true, filterBadge: true });
}

function afClearSearch1() {
  searchQueryEl.value = "";
  filterNegateEl.checked = false;
  filterEmptyModeEl.value = "all";
  quickSearchHighlightMode = false;
  quickSearchCellsMode = false;
  quickSearchFilterCellsMode = false;
  syncQuickSearchInputs();
  syncQuickSearchModeControls();
  afRerun();
}
function afClearSearch2() {
  searchQuery2El.value = "";
  filterNegate2El.checked = false;
  filterEmptyMode2El.value = "all";
  afRerun();
}
function afClearDates() {
  dateModeEl.value = "between";
  dateFromEl.value = "";
  dateToEl.value = "";
  lastDaysEl.value = "";
  dateEmptyModeEl.value = "all";
  dateNegateEl.checked = false;
  updateDateChipsActive();
  afRerun();
}
function afClearOnlyData() {
  onlyNonEmptyEl.checked = false;
  afRerun();
}
function afClearSmart() {
  resetSmartFilterState({ silent: false });
  afRerun();
}
function afClearValidation() {
  validationState.showOnly = false;
  const cb = document.getElementById("validationShowOnly");
  if (cb) cb.checked = false;
  if (typeof refreshValidationView === "function") refreshValidationView();
  else afRerun();
}

function afCollectChips() {
  const chips = [];
  const s1 = afTextFilterChip(searchQueryEl, filterModeEl, filterNegateEl, filterEmptyModeEl, columnSelections.filter1);
  if (s1) chips.push({ id: "search1", parts: s1, clear: afClearSearch1 });
  const s2 = afTextFilterChip(searchQuery2El, filterMode2El, filterNegate2El, filterEmptyMode2El, columnSelections.filter2);
  if (s2) chips.push({ id: "search2", parts: [{ text: t("activeFiltersFilter2"), tag: "label" }, ...s2], clear: afClearSearch2 });
  const d = afDateChip();
  if (d) chips.push({ id: "dates", parts: [{ text: t("activeFiltersDates"), tag: "label" }, ...d], clear: afClearDates });
  if (onlyNonEmptyEl.checked) chips.push({ id: "onlydata", parts: [{ text: t("activeFiltersOnlyData"), tag: "label" }], clear: afClearOnlyData });
  const smartN = typeof smartActiveCount === "function" ? smartActiveCount() : 0;
  if (smartN) chips.push({ id: "smart", parts: [{ text: t("activeFiltersSmart", { n: smartN }), tag: "label" }], clear: afClearSmart });
  if (validationState && validationState.showOnly) chips.push({ id: "validation", parts: [{ text: t("activeFiltersValidation"), tag: "label" }], clear: afClearValidation });
  return chips;
}

function renderActiveFilters() {
  if (!activeFiltersEl) return;
  const snap = lastAppliedFilters;
  const chips = (snap && currentHeaders.length) ? afCollectChips() : [];
  // Nic nie filtruje ani nie wyróżnia → pasek znika (zero miejsca na ekranie).
  if (!chips.length || !snap || (!snap.filtering && !snap.marking)) {
    activeFiltersEl.classList.add("hidden");
    activeFiltersEl.replaceChildren();
    return;
  }
  activeFiltersEl.setAttribute("aria-label", t("activeFiltersAria"));
  const frag = document.createDocumentFragment();

  const count = document.createElement("span");
  count.className = "af-count";
  const loc = typeof currentLang !== "undefined" && currentLang === "en" ? "en-US" : "pl-PL";
  const fmt = (n) => Number(n).toLocaleString(loc);
  count.textContent = snap.filtering
    ? t("activeFiltersShown", { shown: fmt(snap.matched), total: fmt(snap.total) })
    : t("activeFiltersMarked", { count: fmt(snap.matched), total: fmt(snap.total) });
  frag.appendChild(count);

  chips.forEach((chip) => {
    const el = document.createElement("span");
    el.className = "af-chip";
    el.dataset.filter = chip.id;
    const label = chip.parts.map((p) => p.text).join(" ");
    chip.parts.forEach((p) => {
      const part = document.createElement("span");
      part.className = `af-part af-${p.tag}`;
      part.textContent = p.text;
      el.appendChild(part);
    });
    const x = document.createElement("button");
    x.type = "button";
    x.className = "af-remove";
    x.textContent = "×";
    x.setAttribute("aria-label", t("activeFiltersRemove", { label }));
    x.addEventListener("click", chip.clear);
    el.appendChild(x);
    frag.appendChild(el);
  });

  const clearAll = document.createElement("button");
  clearAll.type = "button";
  clearAll.className = "af-clear-all";
  clearAll.textContent = t("activeFiltersClearAll");
  clearAll.setAttribute("aria-label", t("activeFiltersClearAllAria"));
  clearAll.addEventListener("click", () => { if (resetFiltersBtn) resetFiltersBtn.click(); });
  frag.appendChild(clearAll);

  activeFiltersEl.replaceChildren(frag);
  activeFiltersEl.classList.remove("hidden");
}

// ── 2. Menu komórki: „Pokaż tylko takie / Ukryj takie" ─────────────────────────

const cellMenuEl = document.getElementById("cellMenu");
let cellMenuTarget = null;         // { header, value, td }
let cellMenuReturnFocus = null;
let suppressCellClickUntil = 0;    // po przytrzymaniu palcem: połknij klik (nie otwieraj edytora)

// Wartość komórki w składni szybkiego szukania. null = nie da się jej wyrazić
// (&&, ||, { } rozbiłyby zapytanie na części).
function buildCellFilterTerm(header, value, exclude) {
  const h = String(header ?? "").replace(/\s+/g, " ").trim();
  const v = String(value ?? "").replace(/\s+/g, " ").trim();
  if (/&&|\|\||[{}]/.test(h) || /&&|\|\||[{}]/.test(v)) return null;
  return `${h}:${exclude ? "!=" : "="}"${v}"`;
}

// Doklejenie do istniejącego zapytania przez && — tylko gdy znaczenie zostaje to samo.
// Zwraca { query, replaced }.
function combineWithCurrentQuery(term) {
  const prev = String(searchQueryEl.value || "").trim();
  const committed = !!(lastAppliedFilters && (lastAppliedFilters.filtering || lastAppliedFilters.marking));
  if (!prev || !committed) return { query: term, replaced: false };
  if (prev.toLowerCase().includes(term.toLowerCase())) return { query: prev, replaced: false };
  // Odwrócone szukanie (≠) odwróciłoby też doklejony warunek — tu nie da się połączyć.
  if (filterNegateEl.checked) return { query: term, replaced: true };
  // Poprzednie zapytanie bez operatorów było dosłownym tekstem; włączając operatory
  // zmienilibyśmy jego sens, jeśli ma znaki specjalne.
  if (!filterOperatorsEl?.checked && /&&|\|\||[{}!:]|>>|<</.test(prev)) return { query: term, replaced: true };
  if (prev.includes("||")) {
    if (/[{}]/.test(prev)) return { query: term, replaced: true }; // nawiasów nie zagnieżdżamy
    return { query: `{${prev}} && ${term}`, replaced: false };
  }
  return { query: `${prev} && ${term}`, replaced: false };
}

// ── Menu dla zaznaczenia WIELU komórek ───────────────────────────────────────
// Wiersz zaznaczenia = kombinacja wartości w zaznaczonych kolumnach. „Pokaż" = wiersze,
// które mają KTÓRĄKOLWIEK z tych kombinacji; „Ukryj" = wiersze, które nie mają żadnej.
// Parser szukania zna ||, && oraz grupy {…} jednego poziomu (także !{…}), więc:
//   jedna kolumna:  pokaż {K:="a" || K:="b"}          ukryj K:!="a" && K:!="b"
//   kilka kolumn:   pokaż {A:="x" && B:="y"} || {…}   ukryj !{A:="x" && B:="y"} && !{…}
const CELL_MENU_MULTI_MAX = 40; // więcej różnych kombinacji = nieczytelne zapytanie

// Prostokąt zaznaczenia, jeśli komórka pod menu leży W NIM i ma on > 1 komórkę.
function cellMenuSelectionRect(td) {
  const rect = typeof getSelectionRectangle === "function" ? getSelectionRectangle() : null;
  if (!rect || rect.rowCount * rect.colCount < 2) return null;
  const rowKey = td.parentElement?.dataset.rowKey || "";
  const col = parseInt(td.dataset.colIndex || "", 10);
  if (!rect.rowKeys.has(rowKey) || !(col >= rect.colMin && col <= rect.colMax)) return null;
  return rect;
}

function collectSelectionCombos(rect) {
  const headers = [];
  for (let c = rect.colMin; c <= rect.colMax; c++) headers.push({ col: c, header: rect.model.headers[c] });
  const seen = new Set();
  const combos = [];
  for (let r = rect.rowStart; r <= rect.rowEnd; r++) {
    const row = rect.model.rows[r];
    if (!row) continue;
    const values = headers.map(({ col }) => String(getDisplayValue(row, col) ?? "").replace(/\s+/g, " ").trim());
    const key = values.join("\u0000").toLowerCase();
    if (seen.has(key)) continue;
    seen.add(key);
    combos.push(values);
  }
  return { headers, combos };
}

// Zwraca { term, andChain } albo { error }. andChain = da się dokleić do bieżącego
// zapytania przez && (bez zagnieżdżania nawiasów).
function buildMultiCellFilterQuery(target, exclude) {
  const { headers, combos } = target;
  if (headers.some((h) => !String(h.header ?? "").trim())) return { error: t("cellMenuNoHeader") };
  if (combos.length > CELL_MENU_MULTI_MAX) {
    return { error: t("cellMenuMultiTooMany", { count: combos.length, max: CELL_MENU_MULTI_MAX }) };
  }
  const parts = [];
  for (const values of combos) {
    const terms = values.map((v, i) => buildCellFilterTerm(headers[i].header, v, false));
    if (terms.some((x) => !x)) return { error: t("cellMenuUnsupported") };
    parts.push(terms);
  }
  if (headers.length === 1) {
    if (exclude) {
      const term = combos.map((values) => buildCellFilterTerm(headers[0].header, values[0], true)).join(" && ");
      return { term, andChain: true };
    }
    const flat = parts.map((terms) => terms[0]);
    return { term: flat.length === 1 ? flat[0] : `{${flat.join(" || ")}}`, andChain: true };
  }
  if (exclude) return { term: parts.map((terms) => `!{${terms.join(" && ")}}`).join(" && "), andChain: true };
  if (parts.length === 1) return { term: parts[0].join(" && "), andChain: true };
  return { term: parts.map((terms) => `{${terms.join(" && ")}}`).join(" || "), andChain: false };
}

function applyCellFilter(exclude) {
  const target = cellMenuTarget;
  closeCellMenu({ restoreFocus: false });
  if (!target) return;
  let term;
  let andChain = true;
  if (target.multi) {
    const built = buildMultiCellFilterQuery(target, exclude);
    if (built.error) { toast(built.error, "warning"); return; }
    ({ term, andChain } = built);
  } else {
    if (!String(target.header ?? "").trim()) { toast(t("cellMenuNoHeader"), "warning"); return; }
    term = buildCellFilterTerm(target.header, target.value, exclude);
    if (!term) { toast(t("cellMenuUnsupported"), "warning"); return; }
  }
  // Suma kombinacji (…) || (…) nie da się dokleić przez && bez zagnieżdżania nawiasów,
  // których parser nie zna — wtedy bieżące szukanie jest zastępowane (z komunikatem).
  const hadQuery = !!String(searchQueryEl.value || "").trim()
    && !!(lastAppliedFilters && (lastAppliedFilters.filtering || lastAppliedFilters.marking));
  // Warunki, które bieżące szukanie JUŻ zawiera, nie są doklejane drugi raz
  // (np. zaznaczenie wiersza po wcześniejszym „Pokaż tylko takie" na jednej z jego komórek).
  if (target.multi && andChain && hadQuery && !/[{}]/.test(term)) {
    const prevLower = String(searchQueryEl.value || "").toLowerCase();
    const fresh = term.split(" && ").filter((part) => !prevLower.includes(part.toLowerCase()));
    if (fresh.length) term = fresh.join(" && ");
  }
  const { query, replaced } = andChain ? combineWithCurrentQuery(term) : { query: term, replaced: hadQuery };

  if (replaced && filterNegateEl.checked) filterNegateEl.checked = false;
  // Składnia „Kolumna:" działa tylko z operatorami — włączamy je wszędzie naraz.
  [filterOperatorsEl, quickSearchOperatorsEl, quickSearchPopupOperatorsEl].forEach((el) => { if (el) el.checked = true; });
  // „Pokaż tylko" ma filtrować; „Zaznacz"/„Podświetl" zostawiłyby wszystkie wiersze.
  const action = quickSearchActionEl ? quickSearchActionEl.value : "filter";
  const nextAction = action === "filter-cells" ? "filter-cells" : "filter";
  [quickSearchActionEl, quickSearchPopupActionEl].forEach((el) => { if (el) el.value = nextAction; });
  if (quickSearchEl) quickSearchEl.value = query;
  if (quickSearchPopupInput) quickSearchPopupInput.value = query;
  searchQueryEl.value = query;
  syncQuickSearchModeControls();
  applyQuickSearch();
  if (replaced) toast(t("cellMenuReplaced"), "info");
}

function cellMenuItems(target) {
  if (target.multi) {
    return [
      { id: "show", label: t("cellMenuMultiShow"), run: () => applyCellFilter(false) },
      { id: "hide", label: t("cellMenuMultiHide"), run: () => applyCellFilter(true) },
      { id: "copy", label: t("cellMenuMultiCopy"), run: () => { closeCellMenu(); copySelectionToClipboard(); } },
    ];
  }
  const empty = !String(target.value ?? "").trim();
  return [
    { id: "show", label: t(empty ? "cellMenuShowEmpty" : "cellMenuShowOnly"), run: () => applyCellFilter(false) },
    { id: "hide", label: t(empty ? "cellMenuHideEmpty" : "cellMenuHide"), run: () => applyCellFilter(true) },
    { id: "copy", label: t("cellMenuCopy"), run: () => { closeCellMenu(); copySelectionToClipboard(); } },
  ];
}

function openCellMenu(td, point, { viaKeyboard = false } = {}) {
  if (!cellMenuEl || !td || !currentDisplayModel) return false;
  const tr = td.parentElement;
  const rowKey = tr?.dataset.rowKey || "";
  const colIndex0 = parseInt(td.dataset.colIndex || "", 10);
  if (!rowKey || !Number.isFinite(colIndex0)) return false;
  const row = currentDisplayModel.rows.find((r) => getRowSelectionKey(r) === rowKey);
  if (!row) return false;
  hideCellTooltip();

  const head = document.createElement("div");
  head.className = "cell-menu-head";
  const h = document.createElement("span");
  h.className = "cell-menu-col";
  const v = document.createElement("span");
  v.className = "cell-menu-val";

  // Klik W OBRĘBIE zaznaczenia wielu komórek: zaznaczenie zostaje, menu działa na całość.
  const selRect = cellMenuSelectionRect(td);
  if (selRect) {
    const { headers, combos } = collectSelectionCombos(selRect);
    cellMenuTarget = { multi: true, rect: selRect, headers, combos, td };
    h.textContent = afShort(headers.map((x) => x.header || "—").join(", "), 34);
    v.textContent = t("cellMenuMultiHead", { cells: selRect.rowCount * selRect.colCount, distinct: combos.length });
  } else {
    // Komórka pod menu staje się aktywną (Kopiuj kopiuje właśnie ją, widać, o co chodzi).
    setSelectionKind("cell", { repaint: false });
    setFocusedCell(rowKey, colIndex0, { scroll: false });
    const header = currentDisplayModel.headers[colIndex0];
    const value = String(getDisplayValue(row, colIndex0) ?? "");
    cellMenuTarget = { header, value, td };
    h.textContent = afShort(header || "—", 28);
    v.textContent = value.trim() ? `„${afShort(value, 34)}”` : t("cellMenuEmptyValue");
  }
  cellMenuReturnFocus = viaKeyboard ? td : null;
  head.append(h, v);

  const frag = document.createDocumentFragment();
  frag.appendChild(head);
  cellMenuItems(cellMenuTarget).forEach((item, i) => {
    const btn = document.createElement("button");
    btn.type = "button";
    btn.className = `cell-menu-item cmi-${item.id}`;
    btn.setAttribute("role", "menuitem");
    btn.tabIndex = i === 0 ? 0 : -1;
    btn.textContent = item.label;
    btn.addEventListener("click", item.run);
    frag.appendChild(btn);
  });
  cellMenuEl.replaceChildren(frag);
  cellMenuEl.setAttribute("aria-label", t(cellMenuTarget.multi ? "cellMenuMultiAria" : "cellMenuAria"));
  cellMenuEl.classList.remove("hidden");

  // Pozycja: przy punkcie kliknięcia/palca, a z klawiatury pod komórką; zawsze w oknie.
  const rect = td.getBoundingClientRect();
  const vv = window.visualViewport;
  const vw = vv ? vv.width : window.innerWidth;
  const vh = vv ? vv.height : window.innerHeight;
  let x = point ? point.x : rect.left;
  let y = point ? point.y : rect.bottom + 4;
  const mw = cellMenuEl.offsetWidth;
  const mh = cellMenuEl.offsetHeight;
  if (x + mw > vw - 8) x = Math.max(8, vw - mw - 8);
  if (y + mh > vh - 8) y = Math.max(8, (point ? point.y : rect.top) - mh - 4);
  cellMenuEl.style.left = `${Math.round(x)}px`;
  cellMenuEl.style.top = `${Math.round(y)}px`;

  if (viaKeyboard) cellMenuEl.querySelector(".cell-menu-item")?.focus();
  return true;
}

function closeCellMenu({ restoreFocus = true } = {}) {
  if (!cellMenuEl || cellMenuEl.classList.contains("hidden")) return;
  cellMenuEl.classList.add("hidden");
  cellMenuEl.replaceChildren();
  const back = cellMenuReturnFocus;
  cellMenuTarget = null;
  cellMenuReturnFocus = null;
  if (restoreFocus && back && document.contains(back)) back.focus({ preventScroll: true });
}

function cellMenuTd(target) {
  const td = target?.closest?.("td[data-col-index]");
  if (!td || td.classList.contains("row-head")) return null;
  if (td.querySelector("input.cell-editor")) return null; // w edycji zostaje natywne menu pola
  return tbodyEl && tbodyEl.contains(td) ? td : null;
}

if (cellMenuEl && tbodyEl) {
  // Mysz: prawy klik. (Android przy przytrzymaniu też wysyła contextmenu.)
  tbodyEl.addEventListener("contextmenu", (e) => {
    const td = cellMenuTd(e.target);
    if (!td || !workbook) return;
    e.preventDefault();
    openCellMenu(td, { x: e.clientX, y: e.clientY });
    suppressCellClickUntil = Date.now() + 500;
  });

  // Palec: przytrzymanie ~0,5 s bez ruchu. Listenery passive — nigdy nie blokują
  // przewijania; ruch > 10 px = to przewijanie, nie przytrzymanie.
  let lpTimer = 0;
  let lpX = 0;
  let lpY = 0;
  const lpCancel = () => { if (lpTimer) { clearTimeout(lpTimer); lpTimer = 0; } };
  tbodyEl.addEventListener("touchstart", (e) => {
    lpCancel();
    if (e.touches.length !== 1 || !workbook) return;
    const td = cellMenuTd(e.target);
    if (!td) return;
    const touch = e.touches[0];
    lpX = touch.clientX;
    lpY = touch.clientY;
    lpTimer = setTimeout(() => {
      lpTimer = 0;
      if (openCellMenu(td, { x: lpX, y: lpY })) {
        suppressCellClickUntil = Date.now() + 800;
        if (navigator.vibrate) { try { navigator.vibrate(10); } catch { /* bez haptyki */ } }
      }
    }, 520);
  }, { passive: true });
  tbodyEl.addEventListener("touchmove", (e) => {
    if (!lpTimer) return;
    const touch = e.touches[0];
    if (!touch || Math.abs(touch.clientX - lpX) > 10 || Math.abs(touch.clientY - lpY) > 10) lpCancel();
  }, { passive: true });
  tbodyEl.addEventListener("touchend", lpCancel, { passive: true });
  tbodyEl.addEventListener("touchcancel", lpCancel, { passive: true });

  // Klik, który kończy przytrzymanie, nie może trafić do siatki (drugi tap = edytor).
  window.addEventListener("click", (e) => {
    if (Date.now() > suppressCellClickUntil) return;
    if (cellMenuEl.contains(e.target)) return;
    if (tbodyEl.contains(e.target)) { e.stopPropagation(); e.preventDefault(); }
  }, true);

  // Klawiatura: klawisz menu albo Shift+F10 na komórce siatki.
  tbodyEl.addEventListener("keydown", (e) => {
    if (!(e.key === "ContextMenu" || (e.key === "F10" && e.shiftKey))) return;
    const td = cellMenuTd(e.target) || cellMenuTd(document.activeElement)
      || (focusedCellState ? findCellElement(focusedCellState) : null);
    if (!td) return;
    e.preventDefault();
    e.stopPropagation();
    openCellMenu(td, null, { viaKeyboard: true });
  });

  cellMenuEl.addEventListener("keydown", (e) => {
    const items = Array.from(cellMenuEl.querySelectorAll(".cell-menu-item"));
    const i = items.indexOf(document.activeElement);
    if (e.key === "Escape") { e.preventDefault(); e.stopPropagation(); closeCellMenu(); return; }
    if (e.key === "Tab") { e.preventDefault(); closeCellMenu(); return; }
    let next = null;
    if (e.key === "ArrowDown") next = (i + 1) % items.length;
    if (e.key === "ArrowUp") next = (i - 1 + items.length) % items.length;
    if (e.key === "Home") next = 0;
    if (e.key === "End") next = items.length - 1;
    if (next === null) return;
    e.preventDefault();
    e.stopPropagation();
    items.forEach((b, k) => { b.tabIndex = k === next ? 0 : -1; });
    items[next].focus();
  });

  document.addEventListener("pointerdown", (e) => {
    if (!cellMenuEl.classList.contains("hidden") && !cellMenuEl.contains(e.target)) closeCellMenu({ restoreFocus: false });
  }, true);
  document.addEventListener("keydown", (e) => {
    if (e.key === "Escape" && !cellMenuEl.classList.contains("hidden")) { e.stopPropagation(); closeCellMenu(); }
  }, true);
  if (tableWrapEl) tableWrapEl.addEventListener("scroll", () => closeCellMenu({ restoreFocus: false }), { passive: true });
  window.addEventListener("resize", () => closeCellMenu({ restoreFocus: false }));
}
