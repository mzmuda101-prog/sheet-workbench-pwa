// @ts-check
// changed-rows.js — filtr „Zmienione wiersze” (Mateusz 2026-11-20: po Wpisywaniu chcę zobaczyć, co zmieniłem).
//
// Dwa zakresy:
//   • „niezapisane” — wiersze z komórkami w pendingEdits bieżącego arkusza (po Zapisz znikają),
//   • „od otwarcia pliku” — każda komórka zmieniona w tej sesji z plikiem, także już zapisana
//     (ponowne wczytanie po operacji na arkuszu ich nie gubi; nowy plik — od zera).
// Wpięte jak „Tylko wiersze z notatkami”: warunek w swbBuildRowFilter (workbook.js), licznik filtrów,
// chip z ✕ (filter-bar.js). Po wyjściu z Wpisywania dymek ma „Pokaż zmienione” (entry.js).

/** @type {"" | "unsaved" | "all"} */
let swbChangedOnly = "";
/** @type {Map<string, Set<number>>} arkusz → numery wierszy (od 0) zmienione od otwarcia pliku */
const swbChangedSinceOpen = new Map();
const swbChangedOnlyEl = /** @type {HTMLSelectElement | null} */ (document.getElementById("changedOnly"));

/** Wiersz (od 0) z adresu komórki „C7” → 6. */
function swbRowFromRef(/** @type {string} */ cellRef) {
  const m = /([0-9]+)$/.exec(String(cellRef || ""));
  return m ? Number(m[1]) - 1 : -1;
}

/** Wołane z recordPendingEdit (workbook.js) przy każdej zmianie komórki. */
function swbNoteChangedCell(/** @type {string} */ sheetName, /** @type {string} */ cellRef) {
  const r = swbRowFromRef(cellRef);
  if (!sheetName || r < 0) return;
  let set = swbChangedSinceOpen.get(sheetName);
  if (!set) { set = new Set(); swbChangedSinceOpen.set(sheetName, set); }
  set.add(r);
}

/** Zbiór wierszy do pokazania dla bieżącego arkusza albo null, gdy filtr wyłączony. */
function swbChangedRowSet() {
  if (!swbChangedOnly) return null;
  // @ts-ignore — globalne z core.js
  const sheet = typeof currentSheetName === "string" ? currentSheetName : "";
  const out = new Set(swbChangedOnly === "all" ? (swbChangedSinceOpen.get(sheet) || []) : []);
  // @ts-ignore — globalne z core.js
  const pend = (typeof pendingEdits !== "undefined" && pendingEdits && pendingEdits[sheet]) || {};
  Object.keys(pend).forEach((ref) => { const r = swbRowFromRef(ref); if (r >= 0) out.add(r); });
  return out;
}

function swbChangedSyncUi() {
  if (!swbChangedOnlyEl) return;
  swbChangedOnlyEl.value = swbChangedOnly;
  const tt = (/** @type {string} */ k) => (typeof t === "function" ? t(k) : k);
  const label = document.getElementById("changedOnlyLabel");
  if (label) label.textContent = tt("changedOnlyLabel");
  const opts = { "": "changedOnlyOff", unsaved: "changedOnlyUnsaved", all: "changedOnlyAll" };
  Array.from(swbChangedOnlyEl.options).forEach((o) => { o.textContent = tt(opts[/** @type {"" | "unsaved" | "all"} */ (o.value)]); });
}

function swbSetChangedOnly(/** @type {string} */ mode, { rerun = true } = {}) {
  swbChangedOnly = mode === "unsaved" || mode === "all" ? mode : "";
  swbChangedSyncUi();
  // @ts-ignore — cell-notes.js
  if (rerun && typeof swbRerunFilters === "function") swbRerunFilters();
}

/** Nowy plik: zmiany i filtr od zera (ponowne wczytanie po operacji na arkuszu — bez zmian). */
function swbChangedReset() {
  swbChangedSinceOpen.clear();
  swbChangedOnly = "";
  swbChangedSyncUi();
}

if (swbChangedOnlyEl) swbChangedOnlyEl.addEventListener("change", () => swbSetChangedOnly(swbChangedOnlyEl.value));
swbChangedSyncUi();
