// cell-notes.js — notatki komórek jak w Excelu (Recenzja → Notatki).
//
// Po co (prośba Mateusza 2026-10-19): aplikacja tylko LICZYŁA komentarze z pliku
// („Komentarze: N” w panelu) — nie było ich widać ani nie dało się ich dopisać.
//
// Co robi:
//   • czerwony trójkącik w prawym górnym rogu komórki z notatką (idiom Excela),
//   • treść w podpowiedzi komórki: najechanie myszą / tapnięcie (workbook.js showCellTooltip),
//   • dodawanie, edycja i usuwanie: menu komórki (prawy klik / przytrzymanie palcem)
//     albo Shift+F2 na zaznaczonej komórce — jak w Excelu,
//   • zapis do pliku: pendingEdits[arkusz][A1].note = { text, author } | null
//     (xlsx-patch.js → commentsN.xml + kształt VML); Cofnij, szkice i licznik zmian
//     działają na tym samym wpisie bez osobnej hydrauliki.
// Komentarze WĄTKOWE (Excel 365, z odpowiedziami) są tylko do odczytu — ich zapis wymaga
// threadedComments + osób, a zła kopia psułaby wątek w Excelu.
//
// Dane w aplikacji: notatka żyje tam, gdzie kładzie ją SheetJS — cell.c = [{ a, t, T }].

const SWB_NOTE_AUTHOR_KEY = "swb-note-author-v1";

let swbNotesVersion = 0;
let swbNotesIndexCache = null; // { workbook, sheet, version, rows: Map<r, Set<absC>> }

function swbNoteEntries(cell) {
  return cell && Array.isArray(cell.c) && cell.c.length ? cell.c : null;
}

// SheetJS (0.20.3) czyta threadedComments bez dekodowania UTF-8: „Wątek” → „WÄtek”.
// Naprawiamy tylko tekst, który WYGLĄDA na UTF-8 przeczytane bajt po bajcie.
function swbFixUtf8Mojibake(s) {
  const str = String(s ?? "");
  if (!/[\u00c2-\u00f4][\u0080-\u00bf]/.test(str) || /[^\u0000-\u00ff]/.test(str)) return str;
  try { return decodeURIComponent(escape(str)); } catch { return str; }
}

// { text, author, threaded } albo null. Wątek = kolejne wpisy „autor: treść”.
function swbNoteOf(sheetName, ref) {
  const cell = workbook?.Sheets?.[sheetName]?.[ref];
  const list = swbNoteEntries(cell);
  if (!list) return null;
  const threaded = list.some((c) => c && c.T);
  if (threaded) {
    const text = list.map((c) => {
      const a = swbFixUtf8Mojibake(c.a);
      return `${a ? `${a}: ` : ""}${swbFixUtf8Mojibake(c.t).trim()}`;
    }).join("\n\n");
    return { text, author: swbFixUtf8Mojibake(list[0].a), threaded: true };
  }
  return { text: String(list[0].t ?? ""), author: list[0].a || "", threaded: false };
}

// Wpisanie notatki do obiektu komórki (też szkice: drafts.js). note = { text, author } | null.
function swbNotesApplyToCell(sheet, ref, note) {
  if (!sheet) return;
  const prev = sheet[ref];
  if (note) {
    const cell = prev ? { ...prev } : { t: "z" };
    cell.c = [{ a: note.author || "", t: String(note.text ?? ""), T: false }];
    sheet[ref] = cell;
  } else if (prev && prev.c) {
    const cell = { ...prev };
    delete cell.c;
    // sama zaślepka bez stylu i formatu — komórka znika jak w pliku
    if (cell.t === "z" && !cell.s && !cell.z && !cell.f) delete sheet[ref];
    else sheet[ref] = cell;
  }
  swbNotesVersion += 1;
}

function swbNotesInvalidate() { swbNotesVersion += 1; }

// Wiersz → zbiór kolumn (bezwzględnych) z notatką, dla bieżącego arkusza. Budowane leniwie
// raz na zmianę notatek — przy renderze to jedno Map.get na wiersz.
function swbNotesRowIndex() {
  if (!workbook || !currentSheetName) return null;
  const c = swbNotesIndexCache;
  if (c && c.workbook === workbook && c.sheet === currentSheetName && c.version === swbNotesVersion) return c.rows;
  const sheet = workbook.Sheets[currentSheetName];
  const rows = new Map();
  if (sheet) {
    for (const ref in sheet) {
      if (ref.charCodeAt(0) === 33) continue; // „!ref”, „!merges”…
      if (!swbNoteEntries(sheet[ref])) continue;
      const { r, c: col } = XLSX.utils.decode_cell(ref);
      if (!rows.has(r)) rows.set(r, new Set());
      rows.get(r).add(col);
    }
  }
  swbNotesIndexCache = { workbook, sheet: currentSheetName, version: swbNotesVersion, rows };
  return rows;
}

// Dla table.js: kolumny (bezwzględne) z notatką w danym wierszu arkusza albo null.
function swbNotesRowCols(rowIndex0) {
  if (!Number.isFinite(rowIndex0)) return null;
  const rows = swbNotesRowIndex();
  return rows && rows.size ? rows.get(rowIndex0) || null : null;
}

// <td> → adres komórki w arkuszu (tylko widok „wide” i wiersze nad nagłówkiem).
function swbNoteRefForTd(td) {
  if (!td || !td.dataset || td.classList.contains("row-head")) return null;
  const tr = td.closest("tr[data-row-index]");
  if (!tr) return null;
  const inBody = !!td.closest("tbody");
  if (inBody && currentDisplayModel?.mode !== "wide") return null;
  const r = parseInt(tr.dataset.rowIndex || "", 10);
  const col = parseInt(td.dataset.colIndex || "", 10);
  if (!Number.isFinite(r) || !Number.isFinite(col)) return null;
  if (typeof cfmtIsRealCol === "function" && !cfmtIsRealCol(col)) return null;
  return { ref: XLSX.utils.encode_cell({ r, c: currentStartCol + col }), r, col };
}

// Treść do podpowiedzi komórki (workbook.js).
function swbNoteTextForTd(td) {
  if (!td || !td.classList || !td.classList.contains("cell-has-note")) return "";
  const at = swbNoteRefForTd(td);
  const note = at ? swbNoteOf(currentSheetName, at.ref) : null;
  return note ? note.text.trim() : "";
}

// Znacznik na widocznych komórkach danego adresu (tabela + wiersze nad nagłówkiem).
function swbNotesRepaintRef(ref) {
  if (!ref || !currentSheetName) return;
  const { r, c } = XLSX.utils.decode_cell(ref);
  const col = c - currentStartCol;
  const has = !!swbNoteEntries(workbook?.Sheets?.[currentSheetName]?.[ref]);
  document.querySelectorAll(`#dataTable tr[data-row-index="${r}"] td[data-col-index="${col}"]`).forEach((td) => {
    if (td.closest("tbody") && currentDisplayModel?.mode !== "wide") return;
    td.classList.toggle("cell-has-note", has);
  });
}

function swbNotesRecord(sheetName, ref, note) {
  if (!pendingEdits[sheetName]) pendingEdits[sheetName] = {};
  const cur = pendingEdits[sheetName][ref];
  if (cur === null) pendingEdits[sheetName][ref] = { clear: true, note };
  else if (cur === undefined) pendingEdits[sheetName][ref] = { note };
  else pendingEdits[sheetName][ref] = { ...cur, note };
}

function swbNoteAuthor() {
  try { return localStorage.getItem(SWB_NOTE_AUTHOR_KEY) || ""; } catch { return ""; }
}

// Zapis notatki komórki (r = wiersz arkusza, col = kolumna widoku). text pusty = usuń.
function swbSetNote(r, col, text, author) {
  if (!workbook || !currentSheetName) return false;
  if (typeof swbEditGuard === "function" && !swbEditGuard()) return false;
  if (typeof cfmtIsRealCol === "function" && !cfmtIsRealCol(col)) return false;
  const sheet = workbook.Sheets[currentSheetName];
  if (!sheet) return false;
  const ref = XLSX.utils.encode_cell({ r, c: currentStartCol + col });
  const cur = swbNoteOf(currentSheetName, ref);
  if (cur && cur.threaded) return false;
  const clean = String(text ?? "").replace(/\r\n?/g, "\n").replace(/\s+$/, "");
  const note = clean.trim() ? { text: clean, author: cur ? cur.author : String(author || "").trim() } : null;
  if (!note && !cur) return false;
  if (note && cur && cur.text === note.text) return false;
  if (typeof undoCaptureCell === "function") undoCaptureCell(r, col); // Cofnij przywraca komórkę i wpis
  swbNotesApplyToCell(sheet, ref, note);
  swbNotesRecord(currentSheetName, ref, note);
  if (note && typeof swbEnsureSheetRef === "function") swbEnsureSheetRef(sheet, r, currentStartCol + col);
  swbNotesRepaintRef(ref);
  swbNotesSyncFilterUi();
  if (typeof setDirtyState === "function") setDirtyState(true);
  if (typeof scheduleViewRefresh === "function") scheduleViewRefresh({ filterBadge: true });
  return true;
}

// ── Szukanie i filtrowanie po notatkach (2026-10-20) ──────────────────────────
// Panel filtrów (pod „Tylko wiersze z danymi”, widoczne tylko gdy arkusz ma notatki):
//   • „Tylko wiersze z notatkami” — filtr wiersza jak „tylko niezgodne” z walidacji:
//     działa od razu (bez „Filtruj”), ma pigułkę w pasku „co teraz filtruje”,
//   • „Szukaj też w treści notatek” — szukanie (Filtr 1/2, szybkie szukanie) traktuje treść
//     notatki jak drugą zawartość jej komórki. Domyślnie wyłączone: wiersz pasujący „przez
//     notatkę” bez tego wyglądałby na błąd. Koszt: jedno Map.get na wiersz (notatek jest mało).
const SWB_NOTES_SEARCH_KEY = "swb-notes-search-v1";
let swbNotesOnly = false;
const swbNotesOnlyEl = document.getElementById("notesOnly");
const swbNotesSearchEl = document.getElementById("notesSearch");
const swbNotesFilterGroupEl = document.getElementById("notesFilterGroup");

function swbNotesSearchEnabled() {
  try { return localStorage.getItem(SWB_NOTES_SEARCH_KEY) === "1"; } catch { return false; }
}

// kolumny (widoku) z notatką w wierszu tabeli — tylko widok „wide”
function swbRowNoteCols(row) {
  if (!row || row.isLongViewRow || currentDisplayModel?.mode === "long") return null;
  const abs = swbNotesRowCols(row.rowIndex0);
  if (!abs) return null;
  const width = Array.isArray(currentHeaders) ? currentHeaders.length : 0;
  const out = [];
  abs.forEach((c) => { const col = c - currentStartCol; if (col >= 0 && col < width) out.push(col); });
  return out.length ? out.sort((a, b) => a - b) : null;
}

function swbRowHasNote(row) {
  return !!swbRowNoteCols(row);
}

// Dla workbook.js (scopedTermMatcher): czy treść notatki komórki pasuje do termu.
function swbNoteMatchesTerm(row, i, q, mode) {
  const abs = swbNotesRowCols(row.rowIndex0);
  if (!abs || !abs.has(currentStartCol + i)) return false;
  const note = swbNoteOf(currentSheetName, XLSX.utils.encode_cell({ r: row.rowIndex0, c: currentStartCol + i }));
  if (!note) return false;
  if (mode === "regex") {
    const re = typeof compileSearchRegex === "function" ? compileSearchRegex(q) : null;
    return re ? re.test(note.text) : false;
  }
  const text = note.text.toLowerCase();
  if (mode === "equals") return text.trim() === q;
  if (mode === "starts_with") return text.startsWith(q);
  return text.includes(q);
}

function swbSheetNoteCount() {
  const rows = swbNotesRowIndex();
  let n = 0;
  if (rows) rows.forEach((set) => { n += set.size; });
  return n;
}

// Pokazanie/ukrycie grupy w panelu + etykiety z licznikiem. Arkusz bez notatek: nic nie widać
// (a przełącznik „tylko z notatkami” gaśnie, żeby nie zostawić pustej tabeli bez wyjaśnienia).
let swbNotesUiKey = "";
function swbNotesSyncFilterUi() {
  if (!swbNotesFilterGroupEl) return;
  const n = workbook && currentSheetName ? swbSheetNoteCount() : 0;
  // wołane przy każdym renderze tabeli — DOM ruszamy tylko, gdy coś się zmieniło
  const key = `${n}|${swbNotesOnly}|${swbNotesSearchEnabled()}|${typeof currentLang !== "undefined" ? currentLang : ""}`;
  if (key === swbNotesUiKey) return;
  swbNotesUiKey = key;
  swbNotesFilterGroupEl.classList.toggle("hidden", n === 0 && !swbNotesOnly);
  if (swbNotesOnlyEl) {
    swbNotesOnlyEl.checked = swbNotesOnly;
    const label = swbNotesOnlyEl.closest("label")?.querySelector("span");
    if (label) label.textContent = t("notesOnlyLabel", { n });
  }
  if (swbNotesSearchEl) {
    swbNotesSearchEl.checked = swbNotesSearchEnabled();
    const label = swbNotesSearchEl.closest("label")?.querySelector("span");
    if (label) label.textContent = t("notesSearchLabel");
  }
}

function swbRerunFilters() {
  if (typeof applyFilters !== "function") return;
  applyFilters();
  if (typeof sortRows === "function") sortRows();
  if (typeof renderActiveTable === "function") renderActiveTable();
  if (typeof updateFilterBadge === "function") updateFilterBadge();
  if (typeof scheduleViewRefresh === "function") scheduleViewRefresh({ analyses: true });
}

function swbSetNotesOnly(on, { rerun = true } = {}) {
  swbNotesOnly = !!on;
  swbNotesSyncFilterUi();
  if (rerun) swbRerunFilters();
}

if (swbNotesOnlyEl) swbNotesOnlyEl.addEventListener("change", () => swbSetNotesOnly(swbNotesOnlyEl.checked));
if (swbNotesSearchEl) {
  swbNotesSearchEl.addEventListener("change", () => {
    try { localStorage.setItem(SWB_NOTES_SEARCH_KEY, swbNotesSearchEl.checked ? "1" : "0"); } catch { /* tryb prywatny */ }
    swbRerunFilters();
  });
}

// Token w polu szukania (jak @wtoku z Trybów auto): @notatka = wiersz ma notatkę,
// !@notatka = nie ma, @notatka:słowo / @notatka:"cała fraza" = treść notatki zawiera tekst.
// Synonimy: @notatki, @komentarz(e), @note(s), @comment(s). Wycinamy je z zapytania,
// zanim zobaczy je parser termów; reszta zapytania działa jak zwykle (AND z tokenem).
const SWB_NOTE_TOKEN_RE = /(^|\s)(!?)@(?:notatk[aiy]?|komentarz[ey]?|notes?|comments?)(?::("[^"]*"|\S+))?(?=\s|$)/giu;

function swbExtractNoteTokens(query) {
  const q = String(query || "");
  if (!q.includes("@")) return { query: q, rules: null };
  const rules = [];
  const rest = q.replace(SWB_NOTE_TOKEN_RE, (all, pre, neg, term) => {
    let text = term ? term.replace(/^"|"$/g, "").trim().toLowerCase() : "";
    rules.push({ negated: neg === "!", text });
    return pre;
  });
  if (!rules.length) return { query: q, rules: null };
  // spójniki zostawione przez wycięty token („@notatka && Kowalski” → „Kowalski”)
  const query2 = rest.replace(/^\s*(&&|\|\|)\s*/, "").replace(/\s*(&&|\|\|)\s*$/, "").replace(/\s+(&&)\s+(&&)\s+/g, " && ").trim();
  return { query: query2, rules };
}

// Wiersz spełnia wszystkie tokeny notatek (AND).
function swbRowPassesNoteRules(row, rules) {
  if (!rules || !rules.length) return true;
  const cols = swbRowNoteCols(row);
  return rules.every((rule) => {
    let hit;
    if (!rule.text) hit = !!cols;
    else hit = !!cols && cols.some((col) => {
      const note = swbNoteOf(currentSheetName, XLSX.utils.encode_cell({ r: row.rowIndex0, c: currentStartCol + col }));
      return !!note && note.text.toLowerCase().includes(rule.text);
    });
    return rule.negated ? !hit : hit;
  });
}

// Notatki widoczne w tabeli, w kolejności widoku (sortowanie/filtr), wiersz po wierszu.
function swbNotesInView() {
  const model = currentDisplayModel;
  if (!model || model.mode !== "wide" || !Array.isArray(model.rows)) return [];
  const rows = swbNotesRowIndex();
  if (!rows || !rows.size) return [];
  const out = [];
  model.rows.forEach((row) => {
    if (!row || !rows.has(row.rowIndex0)) return;
    const cols = swbRowNoteCols(row);
    if (cols) cols.forEach((col) => out.push({ r: row.rowIndex0, col, rowKey: getRowSelectionKey(row) }));
  });
  return out;
}

// Sąsiednia notatka względem pozycji { r, col } (albo aktywnej komórki), z zawijaniem.
function swbNoteNeighbor(from, dir) {
  const list = swbNotesInView();
  if (!list.length) return null;
  const model = currentDisplayModel;
  let pos = -1; // pozycja „między” notatkami: indeks wiersza w widoku * szerokość + kolumna
  const width = (model?.headers?.length || 0) + 1;
  const rowPos = new Map();
  model.rows.forEach((row, k) => { if (row && !rowPos.has(row.rowIndex0)) rowPos.set(row.rowIndex0, k); });
  if (from && rowPos.has(from.r)) pos = rowPos.get(from.r) * width + from.col;
  const key = (x) => rowPos.get(x.r) * width + x.col;
  let hit = null;
  if (dir > 0) hit = list.find((x) => key(x) > pos) || null;
  else for (let k = list.length - 1; k >= 0; k--) if (key(list[k]) < pos) { hit = list[k]; break; }
  if (!hit) {
    hit = dir > 0 ? list[0] : list[list.length - 1];
    if (list.length > 1 || !from || key(hit) !== pos) toast(t(dir > 0 ? "noteNavWrapStart" : "noteNavWrapEnd"), "info");
  }
  return hit;
}

function swbFocusNoteCell(n) {
  setFocusedCell(n.rowKey, n.col, { scroll: true });
  return document.querySelector(`#dataTable tbody tr[data-row-key="${CSS.escape(n.rowKey)}"] td[data-col-index="${n.col}"]`);
}

// Skrót w tabeli: przejdź do notatki i pokaż jej treść (podpowiedź komórki).
function swbGoToNote(dir) {
  if (!workbook || currentDisplayModel?.mode !== "wide") return false;
  const list = swbNotesInView();
  if (!list.length) {
    toast(t(swbSheetNoteCount() ? "noteNavHidden" : "noteNavNone"), "info");
    return false;
  }
  let from = null;
  if (focusedCellState) {
    const row = currentDisplayModel.rows.find((x) => getRowSelectionKey(x) === focusedCellState.rowKey);
    if (row) from = { r: row.rowIndex0, col: focusedCellState.colIndex0 };
  }
  const next = swbNoteNeighbor(from, dir);
  if (!next) return false;
  if (typeof hideCellTooltip === "function") hideCellTooltip(); // poprzednia treść nie może zostać pod nową komórką
  swbFocusNoteCell(next);
  // komórka mogła dopiero co wjechać w widok (przewijanie / wirtualizacja) — podpowiedź po klatce
  requestAnimationFrame(() => requestAnimationFrame(() => {
    const td = document.querySelector(`#dataTable tbody tr[data-row-key="${CSS.escape(next.rowKey)}"] td[data-col-index="${next.col}"]`);
    if (td && typeof showCellTooltip === "function") showCellTooltip(td, true);
  }));
  return true;
}

// Ctrl/⌘+Alt+N = następna notatka, z Shift = poprzednia; F8 / Shift+F8 jako alias (klawiatury
// tabletowe często nie mają rzędu F). e.key przy wciśniętym Ctrl/⌘ zostaje literą, a AltGr+N
// na polskim układzie daje „ń” — więc polskie znaki nie są przechwytywane.
document.addEventListener("keydown", (e) => {
  const meta = e.ctrlKey || e.metaKey;
  const combo = meta && e.altKey && (String(e.key).toLowerCase() === "n" || (e.metaKey && e.code === "KeyN"));
  const fkey = e.key === "F8" && !meta && !e.altKey;
  if (!combo && !fkey) return;
  if (!workbook) return;
  const dir = e.shiftKey ? -1 : 1;
  if (swbNoteModal.isOpen) { e.preventDefault(); e.stopPropagation(); swbNoteModal.go(dir); return; }
  if (document.querySelector(".modal:not(.hidden)")) return;
  if (typeof activeCellEditor !== "undefined" && activeCellEditor) return;
  const a = document.activeElement;
  const tag = a?.tagName;
  if ((tag === "INPUT" || tag === "TEXTAREA" || a?.isContentEditable) && !a.closest("#dataTable")) return;
  e.preventDefault();
  e.stopPropagation();
  swbGoToNote(dir);
}, true);

// ── Okno notatki ────────────────────────────────────────────────────────────────
const swbNoteModal = (() => {
  let el = null;
  let refs = null;
  let target = null; // { r, col, ref }
  let returnFocus = null;
  let viewOnly = false; // „Pokaż notatkę” z menu — samo czytanie, przejście do edycji przyciskiem

  function build() {
    el = document.createElement("div");
    el.className = "modal hidden note-modal";
    el.id = "noteModal";
    el.setAttribute("role", "dialog");
    el.setAttribute("aria-modal", "true");
    el.setAttribute("aria-labelledby", "noteModalTitle");
    el.innerHTML = `
      <div class="modal-content" tabindex="-1">
        <div class="modal-head">
          <div class="modal-title" id="noteModalTitle"></div>
          <div class="note-modal-nav" hidden>
            <button class="note-modal-navbtn" type="button" data-act="prev"><svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><polyline points="15 18 9 12 15 6"/></svg></button>
            <span class="note-modal-pos"></span>
            <button class="note-modal-navbtn" type="button" data-act="next"><svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><polyline points="9 18 15 12 9 6"/></svg></button>
          </div>
          <button class="modal-close" type="button" data-act="cancel">
            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><line x1="18" y1="6" x2="6" y2="18" /><line x1="6" y1="6" x2="18" y2="18" /></svg>
          </button>
        </div>
        <div class="note-modal-body">
          <p class="note-modal-info" hidden></p>
          <label class="field note-modal-field"><span class="note-modal-label"></span>
            <textarea class="note-modal-text" rows="5" spellcheck="true"></textarea>
          </label>
          <label class="field note-modal-author"><span class="note-modal-author-label"></span>
            <input class="note-modal-author-input" type="text" autocomplete="name" spellcheck="false" maxlength="80" />
          </label>
          <p class="note-modal-hint"></p>
        </div>
        <div class="modal-actions">
          <button class="btn ghost danger note-modal-delete" type="button" data-act="delete"></button>
          <span class="note-modal-spacer"></span>
          <button class="btn ghost note-modal-cancel" type="button" data-act="cancel"></button>
          <button class="btn note-modal-edit" type="button" data-act="edit"></button>
          <button class="btn note-modal-save" type="button" data-act="save"></button>
        </div>
      </div>`;
    document.body.appendChild(el);
    refs = {
      content: el.querySelector(".modal-content"),
      title: el.querySelector(".modal-title"),
      close: el.querySelector(".modal-close"),
      info: el.querySelector(".note-modal-info"),
      label: el.querySelector(".note-modal-label"),
      text: el.querySelector(".note-modal-text"),
      authorWrap: el.querySelector(".note-modal-author"),
      authorLabel: el.querySelector(".note-modal-author-label"),
      author: el.querySelector(".note-modal-author-input"),
      hint: el.querySelector(".note-modal-hint"),
      del: el.querySelector(".note-modal-delete"),
      cancel: el.querySelector(".note-modal-cancel"),
      save: el.querySelector(".note-modal-save"),
      edit: el.querySelector(".note-modal-edit"),
      nav: el.querySelector(".note-modal-nav"),
      pos: el.querySelector(".note-modal-pos"),
      prev: el.querySelector('[data-act="prev"]'),
      next: el.querySelector('[data-act="next"]'),
    };
    el.addEventListener("click", (e) => {
      if (e.target === el) { close(); return; }
      const act = e.target.closest("[data-act]")?.dataset.act;
      if (act === "cancel") close();
      else if (act === "save") save();
      else if (act === "delete") remove();
      else if (act === "edit") toEdit();
      else if (act === "prev" || act === "next") go(act === "next" ? 1 : -1);
    });
    // Klawisze okna łapiemy na dokumencie (faza przechwytywania): Esc wciśnięty tuż po
    // otwarciu, zanim fokus wejdzie do pola, trafiał do siatki i okno zostawało otwarte.
    document.addEventListener("keydown", (e) => { if (!el.classList.contains("hidden")) onKey(e); }, true);
  }

  function focusables() {
    return Array.from(el.querySelectorAll("button, textarea, input")).filter((x) => !x.disabled && !x.closest("[hidden]") && x.offsetParent !== null);
  }

  function onKey(e) {
    if (e.key === "Escape") { e.preventDefault(); e.stopPropagation(); close(); return; }
    if (e.key === "Enter" && (e.metaKey || e.ctrlKey)) { e.preventDefault(); e.stopPropagation(); save(); return; }
    if (!el.contains(e.target)) {
      // fokus jeszcze poza oknem — nic nie może trafić do siatki pod spodem
      if (e.key !== "Tab") return;
      e.preventDefault();
      refs.text.focus({ preventScroll: true });
      return;
    }
    if (e.key === "Tab") {
      const list = focusables();
      if (!list.length) return;
      const i = list.indexOf(document.activeElement);
      if (e.shiftKey && i <= 0) { e.preventDefault(); list[list.length - 1].focus(); }
      else if (!e.shiftKey && i === list.length - 1) { e.preventDefault(); list[0].focus(); }
    }
  }

  function open(r, col, { focusBack = null, view = false } = {}) {
    if (!workbook || !currentSheetName) return false;
    if (typeof cfmtIsRealCol === "function" && !cfmtIsRealCol(col)) return false;
    if (!el) build();
    const ref = XLSX.utils.encode_cell({ r, c: currentStartCol + col });
    const note = swbNoteOf(currentSheetName, ref);
    const locked = !!(note && note.threaded) || (typeof swbIsEditMode === "function" && !swbIsEditMode());
    viewOnly = !!view && !!note;
    const readOnly = locked || viewOnly;
    // Podgląd bez notatki: nie ma czego pokazać — bramka edycji mówi, jak to zmienić.
    if (!note && readOnly) { if (typeof swbEditGuard === "function") swbEditGuard(); return false; }
    target = { r, col, ref };
    returnFocus = focusBack;
    refs.title.textContent = t(note ? (note.threaded ? "noteTitleThread" : "noteTitleEdit") : "noteTitleNew", { ref });
    refs.close.setAttribute("aria-label", t("noteClose"));
    refs.label.textContent = t("noteTextLabel");
    refs.text.value = note ? note.text : "";
    refs.text.readOnly = readOnly;
    refs.text.placeholder = t("notePlaceholder");
    refs.info.hidden = !locked;
    refs.info.textContent = note && note.threaded ? t("noteThreadInfo") : t("noteViewModeInfo");
    // autor: przy nowej notatce do wpisania (zapamiętany), przy istniejącej tylko pokazany
    refs.authorWrap.hidden = readOnly && !note?.author;
    refs.authorLabel.textContent = t("noteAuthorLabel");
    refs.author.value = note ? note.author : swbNoteAuthor();
    refs.author.readOnly = !!note || readOnly;
    refs.author.placeholder = t("noteAuthorPlaceholder");
    refs.hint.textContent = readOnly ? "" : t("noteHint");
    refs.hint.hidden = readOnly;
    refs.del.textContent = t("noteDelete");
    refs.del.hidden = !note || readOnly;
    refs.cancel.textContent = t(readOnly ? "noteClose" : "noteCancel");
    refs.save.textContent = t("noteSave");
    refs.save.hidden = readOnly;
    // samo czytanie, ale notatkę da się zmienić — „Edytuj” w tym samym oknie
    refs.edit.textContent = t("noteEdit");
    refs.edit.hidden = !(viewOnly && !locked);
    // „Poprzednia / Następna notatka” jak w Excelu — po notatkach widocznych w tabeli
    const list = swbNotesInView();
    const at = list.findIndex((x) => x.r === r && x.col === col);
    const others = list.length - (at >= 0 ? 1 : 0);
    refs.nav.hidden = others < 1;
    refs.pos.textContent = at >= 0 ? t("noteNavPos", { i: at + 1, n: list.length }) : t("noteNavCount", { n: list.length });
    refs.prev.setAttribute("aria-label", t("noteNavPrev"));
    refs.next.setAttribute("aria-label", t("noteNavNext"));
    refs.prev.dataset.hint = "";
    refs.prev.dataset.hintPl = "Poprzednia notatka (Ctrl/⌘+Alt+Shift+N · Shift+F8)";
    refs.prev.dataset.hintEn = "Previous note (Ctrl/⌘+Alt+Shift+N · Shift+F8)";
    refs.next.dataset.hint = "";
    refs.next.dataset.hintPl = "Następna notatka (Ctrl/⌘+Alt+N · F8)";
    refs.next.dataset.hintEn = "Next note (Ctrl/⌘+Alt+N · F8)";
    if (typeof hideCellTooltip === "function") hideCellTooltip();
    el.classList.remove("hidden");
    // Dotyk + istniejąca notatka: najpierw do przeczytania — klawiatura wysunięta od razu
    // zasłoniłaby pół okna. Pisanie zaczyna się od tapnięcia w pole.
    const coarse = window.matchMedia && window.matchMedia("(pointer: coarse)").matches;
    if ((coarse || viewOnly) && note) {
      refs.content.focus({ preventScroll: true });
    } else {
      refs.text.focus({ preventScroll: true });
      const end = refs.text.value.length;
      try { refs.text.setSelectionRange(end, end); } catch { /* bez zaznaczenia */ }
    }
    return true;
  }

  function close() {
    if (!el || el.classList.contains("hidden")) return;
    el.classList.add("hidden");
    target = null;
    const back = returnFocus;
    returnFocus = null;
    if (back && document.contains(back)) back.focus({ preventScroll: true });
  }

  function save() {
    if (!target || refs.text.readOnly) { close(); return; }
    const author = refs.author.readOnly ? "" : refs.author.value.trim();
    if (!refs.author.readOnly) {
      try { localStorage.setItem(SWB_NOTE_AUTHOR_KEY, author); } catch { /* tryb prywatny */ }
    }
    const had = !!swbNoteOf(currentSheetName, target.ref);
    const text = refs.text.value;
    const { r, col } = target;
    close();
    if (swbSetNote(r, col, text, author)) {
      const removed = had && !text.trim();
      toast(t(removed ? "noteRemoved" : had ? "noteSaved" : "noteAdded", { ref: XLSX.utils.encode_cell({ r, c: currentStartCol + col }) }), "success");
    }
  }

  // Przejście do sąsiedniej notatki. Zmieniona treść bieżącej zapisuje się po drodze
  // (jak w Excelu — tam też nie ma „Anuluj” przy przechodzeniu między notatkami).
  function go(dir) {
    if (!target) return;
    const { r, col } = target;
    if (!refs.text.readOnly) {
      const cur = swbNoteOf(currentSheetName, target.ref);
      const text = refs.text.value;
      if ((cur ? cur.text : "") !== text.replace(/\r\n?/g, "\n").replace(/\s+$/, "")) {
        if (!refs.author.readOnly) { try { localStorage.setItem(SWB_NOTE_AUTHOR_KEY, refs.author.value.trim()); } catch { /* tryb prywatny */ } }
        swbSetNote(r, col, text, refs.author.readOnly ? "" : refs.author.value.trim());
      }
    }
    const next = swbNoteNeighbor({ r, col }, dir);
    if (!next) return;
    const td = swbFocusNoteCell(next);
    open(next.r, next.col, { focusBack: td || returnFocus, view: viewOnly });
  }

  function toEdit() {
    if (!target) return;
    const { r, col } = target;
    open(r, col, { focusBack: returnFocus });
    refs.text.focus({ preventScroll: true }); // świadomy wybór „Edytuj” — od razu do pisania, też na dotyku
    const end = refs.text.value.length;
    try { refs.text.setSelectionRange(end, end); } catch { /* bez zaznaczenia */ }
  }

  function remove() {
    if (!target) return;
    const { r, col, ref } = target;
    close();
    if (swbSetNote(r, col, "", "")) toast(t("noteRemoved", { ref }), "success");
  }

  return { open, close, go, get isOpen() { return !!el && !el.classList.contains("hidden"); } };
})();

// Pozycje menu komórki (filter-bar.js). target = { row, col, td }.
function swbNoteMenuItems(target) {
  if (!target || !target.row || !Number.isFinite(target.row.rowIndex0) || currentDisplayModel?.mode !== "wide") return [];
  if (typeof cfmtIsRealCol === "function" && !cfmtIsRealCol(target.col)) return [];
  const r = target.row.rowIndex0;
  const col = target.col;
  const note = swbNoteOf(currentSheetName, XLSX.utils.encode_cell({ r, c: currentStartCol + col }));
  const run = () => { closeCellMenu({ restoreFocus: false }); swbNoteModal.open(r, col, { focusBack: target.td }); };
  if (!note) return [{ id: "note", label: t("cellMenuNoteAdd"), run }];
  if (note.threaded) return [{ id: "note", label: t("cellMenuNoteThread"), run }];
  // „Pokaż” = przeczytać w spokoju (dymek po tapnięciu znika sam po kilku sekundach)
  const view = () => { closeCellMenu({ restoreFocus: false }); swbNoteModal.open(r, col, { focusBack: target.td, view: true }); };
  return [
    { id: "note-view", label: t("cellMenuNoteView"), run: view },
    { id: "note", label: t("cellMenuNoteEdit"), run },
    {
      id: "note-del",
      label: t("cellMenuNoteDelete"),
      run: () => {
        closeCellMenu();
        if (swbSetNote(r, col, "", "")) toast(t("noteRemoved", { ref: XLSX.utils.encode_cell({ r, c: currentStartCol + col }) }), "success");
      },
    },
  ];
}

// Shift+F2 — „Nowa notatka / Edytuj notatkę” jak w Excelu (na zaznaczonej komórce).
document.addEventListener("keydown", (e) => {
  if (e.key !== "F2" || !e.shiftKey || e.ctrlKey || e.metaKey || e.altKey) return;
  if (swbNoteModal.isOpen || !workbook || !focusedCellState || currentDisplayModel?.mode !== "wide") return;
  if (typeof activeCellEditor !== "undefined" && activeCellEditor) return;
  const tag = document.activeElement?.tagName;
  if ((tag === "INPUT" || tag === "TEXTAREA") && !document.activeElement.closest("#dataTable")) return;
  const row = currentDisplayModel.rows.find((x) => getRowSelectionKey(x) === focusedCellState.rowKey);
  if (!row || !Number.isFinite(row.rowIndex0)) return;
  e.preventDefault();
  e.stopPropagation();
  const td = document.querySelector(`#dataTable tbody tr[data-row-key="${CSS.escape(focusedCellState.rowKey)}"] td[data-col-index="${focusedCellState.colIndex0}"]`);
  swbNoteModal.open(row.rowIndex0, focusedCellState.colIndex0, { focusBack: td });
}, true);
