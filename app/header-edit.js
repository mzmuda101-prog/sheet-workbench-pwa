// app/header-edit.js — ZMIANA NAZWY KOLUMNY (komórka w wierszu nagłówka), 2026-10-31.
//
// Dlaczego osobno, a nie „jak zwykła komórka”: nazwa kolumny to KLUCZ, po którym aplikacja
// odwołuje się do kolumny — sortowanie, wybrane kolumny filtrów, zapytania szukania
// („Status:="W trakcie"”), kolumny wyliczane („[Długość]”), zapisane widoki. Gdyby nagłówek
// zmieniał się jak zwykła wartość, wszystko to po cichu przestałoby trafiać w kolumnę.
//
// Zasady (porządnie i bezpiecznie):
//   1. Nazwy kolumn zawsze wynikają z ARKUSZA. Każda zmiana komórki w wierszu nagłówka
//      (zmiana nazwy, Cofnij, Ponów) kończy się swbSyncHeadersFromSheet(), które porównuje
//      nazwy z arkusza z currentHeaders i przenosi odwołania stanu ze starej nazwy na nową.
//      Dzięki temu Cofnij przywraca nie tylko komórkę, ale i sortowanie/filtry/szukanie.
//   2. Blokujemy to, czego nie da się zmienić bez ryzyka dla pliku: kolumnę TABELI EXCELA
//      (nazwa nagłówka = nazwa kolumny tabeli w xl/tables/*.xml i w odwołaniach
//      strukturalnych formuł w całym skoroszycie — Excel naprawiałby/odrzucał plik),
//      komórkę z formułą, komórkę przykrytą scaleniem, kolumnę wyliczaną (nie ma jej w pliku).
//   3. Nazwa musi być niepusta, unikalna i bez znaków składni szukania (" { } : && ||).
//   4. Zapisane widoki są wspólne dla wszystkich plików — NIE przepisujemy ich (ta sama
//      nazwa może znaczyć coś innego w innym pliku). Widok ze starą nazwą pokaże się jako
//      „brak kolumny”, tak jak dla każdego pliku bez tej kolumny.
//
// UI: prawy klik / przytrzymanie palcem na nagłówku kolumny → menu (zmiana nazwy + sortowanie).
// Zwykłe kliknięcie/stuknięcie nagłówka dalej sortuje — tego gestu nie ruszamy (dlatego też
// bez dwukliku: pierwsze kliknięcie dwukliku już by posortowało).

const SWB_HDR_BAD = /["{}:]|&&|\|\|/;

function swbHeaderRef(col) {
  return XLSX.utils.encode_cell({ r: currentHeaderRow - 1, c: currentStartCol + col });
}

// Ile kolumn pochodzi z pliku (kolumny wyliczane są dopisane na końcu currentHeaders).
function swbRealHeaderCount() {
  const derived = typeof dcAppliedCount === "number" ? dcAppliedCount : 0;
  return Math.max(0, currentHeaders.length - derived);
}

// Powód, dla którego nazwy tej kolumny nie wolno zmienić (null = wolno).
function swbHeaderRenameBlocker(col) {
  if (!workbook || !currentSheetName || !currentDisplayModel || currentDisplayModel.mode !== "wide") return t("editWideOnly");
  if (!Number.isInteger(col) || col < 0 || col >= swbRealHeaderCount()) return t("hdrRenameDerived");
  if (typeof cfmtIsRealCol === "function" && !cfmtIsRealCol(col)) return t("hdrRenameDerived");
  const r = currentHeaderRow - 1;
  const c = currentStartCol + col;
  for (const m of currentMerges || []) {
    if (!m || !m.s || !m.e) continue;
    if (r >= m.s.r && r <= m.e.r && c >= m.s.c && c <= m.e.c && (r !== m.s.r || c !== m.s.c)) return t("hdrRenameMerged");
  }
  const cell = workbook.Sheets[currentSheetName]?.[swbHeaderRef(col)];
  if (cell && cell.f) return t("hdrRenameFormula");
  const tables = typeof currentTables === "object" && currentTables ? currentTables : {};
  const seen = new Set();
  for (const [name, tb] of Object.entries(tables)) {
    if (!tb || seen.has(tb)) continue;
    seen.add(tb);
    if (!(tb.headerRows > 0)) continue;
    // arkusz tabeli nieznany → ostrożnie: traktuj jak ten arkusz
    if (tb.sheetName && tb.sheetName !== currentSheetName) continue;
    let rg;
    try { rg = XLSX.utils.decode_range(tb.ref); } catch { continue; }
    if (c >= rg.s.c && c <= rg.e.c && r >= rg.s.r && r < rg.s.r + tb.headerRows) return t("hdrRenameTable", { table: name });
  }
  return null;
}

// Nazwy kolumn tak, jak policzyłby je buildRows (table.js) z bieżącej zawartości arkusza.
function swbSheetHeaderNames() {
  const sheet = workbook?.Sheets?.[currentSheetName];
  if (!sheet) return null;
  const raw = [];
  for (let i = 0; i < swbRealHeaderCount(); i++) {
    const cell = sheet[swbHeaderRef(i)];
    const v = cell ? cell.v : null;
    raw.push(v ? String(v).trim() : XLSX.utils.encode_col(currentStartCol + i));
  }
  return makeHeadersUnique(raw);
}

const swbEscapeRe = (s) => String(s).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");

// „Status:="x" && Nr.:>3” → nazwy kolumn w tokenach zapytania (poza tekstem w cudzysłowie).
function swbRenameInQuery(query, renames) {
  const q = String(query || "");
  if (!q || !renames.length) return q;
  const parts = q.split(/("[^"]*")/);
  return parts.map((part, i) => {
    if (i % 2 === 1) return part; // wartość w cudzysłowie — nie ruszamy
    let out = part;
    renames.forEach(([from, to]) => {
      const re = new RegExp(`(^|[\\s{(!&|])${swbEscapeRe(from)}(\\s*:)`, "gi");
      out = out.replace(re, (m, pre, colon) => `${pre}${to}${colon}`);
    });
    return out;
  }).join("");
}

// Odwołania stanu ze starych nazw na nowe. renames = [[stara, nowa], …] (jedno przejście,
// więc zamiana nazw dwóch kolumn miejscami też jest poprawna).
function swbMigrateColumnRefs(renames) {
  const map = new Map(renames.map(([from, to]) => [from, to]));
  const ren = (name) => (map.has(name) ? map.get(name) : name);
  if (sortState && sortState.col) sortState.col = ren(sortState.col);
  if (Array.isArray(multiSortState)) multiSortState.forEach((rule) => { if (rule && rule.col) rule.col = ren(rule.col); });
  Object.values(columnSelections).forEach((set) => {
    const next = [...set].map(ren);
    set.clear();
    next.forEach((h) => set.add(h));
  });
  [searchQueryEl, searchQuery2El, typeof quickSearchEl !== "undefined" ? quickSearchEl : null,
    typeof quickSearchPopupInput !== "undefined" ? quickSearchPopupInput : null].forEach((el) => {
    if (el && typeof el.value === "string" && el.value) el.value = swbRenameInQuery(el.value, renames);
  });
  if (typeof derivedColumns !== "undefined" && Array.isArray(derivedColumns)) {
    derivedColumns.forEach((d) => {
      if (!d || typeof d.expr !== "string") return;
      renames.forEach(([from, to]) => {
        d.expr = d.expr.replace(new RegExp(`\\[${swbEscapeRe(from)}\\]`, "gi"), `[${to}]`);
      });
    });
  }
}

// Wiersz nagłówka w arkuszu się zmienił → nazwy kolumn i wszystko, co się do nich odwołuje.
// Zwraca liczbę zmienionych nazw.
function swbSyncHeadersFromSheet() {
  if (!workbook || !currentSheetName || !Array.isArray(currentHeaders) || !currentHeaders.length) return 0;
  const names = swbSheetHeaderNames();
  if (!names) return 0;
  const renames = [];
  names.forEach((name, i) => { if (name !== currentHeaders[i]) renames.push([currentHeaders[i], name, i]); });
  if (!renames.length) return 0;
  swbMigrateColumnRefs(renames.map(([from, to]) => [from, to]));
  renames.forEach(([, to, i]) => { currentHeaders[i] = to; });
  if (typeof dcClearCaches === "function") dcClearCaches();
  if (typeof applyDerivedColumns === "function" && typeof derivedColumns !== "undefined" && derivedColumns.length) applyDerivedColumns({ silent: true });
  if (typeof smartInvalidateModel === "function") smartInvalidateModel();
  if (typeof currentColumnProfiles !== "undefined") currentColumnProfiles = [];
  if (typeof bumpSheetDataStamp === "function") bumpSheetDataStamp();
  if (typeof populateSortColumnSelect === "function") populateSortColumnSelect();
  if (typeof populateEditColumnSelect === "function") populateEditColumnSelect();
  if (typeof populateValidationColumns === "function") populateValidationColumns();
  if (typeof updateColumnSummary === "function") updateColumnSummary();
  if (typeof updateSortControls === "function") updateSortControls();
  applyFilters();
  sortRows();
  scheduleViewRefresh({ table: true, analyses: true, filterBadge: true });
  log(`Nazwy kolumn: ${renames.map(([from, to]) => `${from} → ${to}`).join(", ")}`, "info");
  return renames.length;
}

// Wołane przez każdy zapis komórki w wierszu nagłówka (updateSheetCell, Cofnij/Ponów).
// Synchronizacja po zakończeniu bieżącego zadania — kilka zmian = jedno przeliczenie.
let _hdrSyncQueued = false;
function swbHeaderRowTouched(sheetName, rowIndex0) {
  if (sheetName !== currentSheetName || rowIndex0 !== currentHeaderRow - 1 || _hdrSyncQueued) return;
  _hdrSyncQueued = true;
  queueMicrotask(() => { _hdrSyncQueued = false; swbSyncHeadersFromSheet(); });
}

// Zmiana nazwy kolumny `col`. Zwraca true, gdy nazwa jest (już) taka, jak wpisano.
function swbRenameColumn(col, raw) {
  if (typeof swbEditGuard === "function" && !swbEditGuard()) return false;
  const reason = swbHeaderRenameBlocker(col);
  if (reason) { toast(reason, "info"); return false; }
  const name = String(raw ?? "").replace(/\s+/g, " ").trim();
  if (name === currentHeaders[col]) return true;
  if (!name) { toast(t("hdrRenameEmpty"), "warning"); return false; }
  if (SWB_HDR_BAD.test(name)) { toast(t("hdrRenameBadChars"), "warning"); return false; }
  const key = normalizeHeaderKey(name);
  if (currentHeaders.some((h, i) => i !== col && normalizeHeaderKey(h) === key)) {
    toast(t("hdrRenameDuplicate", { name }), "warning");
    return false;
  }
  // zawsze jako tekst: nagłówek „2026” ma zostać nazwą, nie liczbą
  updateSheetCell(currentHeaderRow - 1, col, { value: name, type: "string" });
  setDirtyState(true);
  swbSyncHeadersFromSheet();
  toast(t("hdrRenamed", { name }), "success");
  return true;
}

// ── Edytor w nagłówku ────────────────────────────────────────────────────────────
function swbHeaderTh(col) {
  return theadEl ? theadEl.querySelector(`tr.header-row th[data-col-index="${col}"]`) : null;
}

function swbOpenHeaderEditor(col) {
  const th = swbHeaderTh(col);
  if (!th || th.querySelector("input.hdr-editor")) return;
  if (typeof swbEditGuard === "function" && !swbEditGuard()) return;
  const reason = swbHeaderRenameBlocker(col);
  if (reason) { toast(reason, "info"); return; }
  if (typeof flushCellEditor === "function") flushCellEditor();
  const input = document.createElement("input");
  input.type = "text";
  input.className = "hdr-editor";
  input.value = currentHeaders[col];
  input.setAttribute("aria-label", t("hdrRenameAria", { name: currentHeaders[col] }));
  th.classList.add("hdr-editing");
  th.appendChild(input);
  let done = false;
  const close = () => {
    if (done) return;
    done = true;
    input.remove();
    th.classList.remove("hdr-editing");
  };
  const commit = () => {
    if (done) return;
    if (swbRenameColumn(col, input.value)) close();
    else { input.focus(); input.select(); }
  };
  input.addEventListener("keydown", (e) => {
    e.stopPropagation();
    if (e.key === "Enter") { e.preventDefault(); commit(); }
    else if (e.key === "Escape") { e.preventDefault(); close(); }
  });
  // wyjście z pola: zapisz; gdy się nie da (pusta/zdublowana nazwa) — komunikat i bez zmian
  input.addEventListener("blur", () => setTimeout(() => {
    if (done) return;
    swbRenameColumn(col, input.value);
    close();
  }, 0));
  input.addEventListener("click", (e) => e.stopPropagation()); // klik w pole ≠ sortowanie
  input.focus();
  input.select();
}

// ── Menu nagłówka (prawy klik / przytrzymanie) ───────────────────────────────────
function swbOpenHeaderMenu(th, point) {
  if (typeof cellMenuEl === "undefined" || !cellMenuEl || !th) return false;
  const col = parseInt(th.dataset.colIndex || "", 10);
  if (!Number.isInteger(col) || !currentHeaders[col]) return false;
  const name = currentHeaders[col];
  const blocker = swbIsEditMode() ? swbHeaderRenameBlocker(col) : t("viewModeBlocked");
  const head = document.createElement("div");
  head.className = "cell-menu-head";
  const h = document.createElement("span");
  h.className = "cell-menu-col";
  h.textContent = XLSX.utils.encode_col(currentStartCol + col) + currentHeaderRow;
  const v = document.createElement("span");
  v.className = "cell-menu-val";
  v.textContent = `„${name.length > 34 ? `${name.slice(0, 33)}…` : name}”`;
  head.append(h, v);
  const items = [
    { id: "hdr-rename", label: t("hdrRenameItem"), disabled: !!blocker, run: () => swbOpenHeaderEditor(col) },
    { id: "hdr-sort-asc", label: t("hdrSortAsc"), run: () => swbHeaderSort(name, "asc") },
    { id: "hdr-sort-desc", label: t("hdrSortDesc"), run: () => swbHeaderSort(name, "desc") },
  ];
  const frag = document.createDocumentFragment();
  frag.appendChild(head);
  items.forEach((item, i) => {
    const btn = document.createElement("button");
    btn.type = "button";
    btn.className = `cell-menu-item cmi-${item.id}`;
    btn.setAttribute("role", "menuitem");
    btn.tabIndex = i === 0 ? 0 : -1;
    btn.textContent = item.label;
    if (item.disabled) btn.setAttribute("aria-disabled", "true");
    btn.addEventListener("click", () => {
      closeCellMenu({ restoreFocus: false });
      if (item.disabled) { if (blocker) toast(blocker, "info"); return; }
      item.run();
    });
    frag.appendChild(btn);
    if (item.disabled && blocker) {
      const why = document.createElement("div");
      why.className = "cell-menu-note";
      why.textContent = blocker;
      frag.appendChild(why);
    }
  });
  cellMenuEl.replaceChildren(frag);
  cellMenuEl.setAttribute("aria-label", t("hdrMenuAria"));
  cellMenuEl.classList.remove("hidden");
  const rect = th.getBoundingClientRect();
  const vv = window.visualViewport;
  const vw = vv ? vv.width : window.innerWidth;
  const vh = vv ? vv.height : window.innerHeight;
  let x = point ? point.x : rect.left;
  let y = point ? point.y : rect.bottom + 4;
  const mw = cellMenuEl.offsetWidth;
  const mh = cellMenuEl.offsetHeight;
  if (x + mw > vw - 8) x = Math.max(8, vw - mw - 8);
  if (y + mh > vh - 8) y = Math.max(8, rect.top - mh - 4);
  cellMenuEl.style.left = `${Math.round(x + (vv ? vv.offsetLeft : 0))}px`;
  cellMenuEl.style.top = `${Math.round(y + (vv ? vv.offsetTop : 0))}px`;
  return true;
}

function swbHeaderSort(name, dir) {
  setPrimarySort(name, dir);
  if (tableViewMode !== "long") sortRows();
  updateSortControls();
  renderActiveTable();
}

if (theadEl) {
  const thFrom = (target) => {
    const th = target?.closest?.("tr.header-row th[data-sort-header]");
    return th && !target.closest("input") ? th : null;
  };
  theadEl.addEventListener("contextmenu", (e) => {
    const th = thFrom(e.target);
    if (!th || !workbook) return;
    e.preventDefault();
    swbOpenHeaderMenu(th, { x: e.clientX, y: e.clientY });
    suppressCellClickUntil = Date.now() + 500;
  });
  // palec: przytrzymanie ~0,5 s bez ruchu (jak menu komórki w filter-bar.js)
  let lpTimer = 0;
  let lpX = 0;
  let lpY = 0;
  const lpCancel = () => { if (lpTimer) { clearTimeout(lpTimer); lpTimer = 0; } };
  theadEl.addEventListener("touchstart", (e) => {
    lpCancel();
    if (e.touches.length !== 1 || !workbook) return;
    const th = thFrom(e.target);
    if (!th) return;
    lpX = e.touches[0].clientX;
    lpY = e.touches[0].clientY;
    lpTimer = setTimeout(() => {
      lpTimer = 0;
      if (swbOpenHeaderMenu(th, { x: lpX, y: lpY })) {
        suppressCellClickUntil = Date.now() + 800; // klik kończący przytrzymanie nie sortuje
        swbMenuOpenedByLongPress(); // …ani nie wybiera pozycji menu, która wypadła pod palcem
        if (navigator.vibrate) { try { navigator.vibrate(10); } catch { /* bez haptyki */ } }
      }
    }, 520);
  }, { passive: true });
  theadEl.addEventListener("touchmove", (e) => {
    if (!lpTimer) return;
    const tch = e.touches[0];
    if (!tch || Math.abs(tch.clientX - lpX) > 10 || Math.abs(tch.clientY - lpY) > 10) lpCancel();
  }, { passive: true });
  theadEl.addEventListener("touchend", lpCancel, { passive: true });
  theadEl.addEventListener("touchcancel", lpCancel, { passive: true });
}
