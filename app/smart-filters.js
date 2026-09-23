// Tryby auto (smart filtry) — filtrowanie po STANIE rekordu, nie po samym tekście.
//
// Arkusze operacyjne bardzo często są „szerokie": jeden wiersz = jeden obiekt
// (teren, sprawa, maszyna), a w kolumnach powtarzają się bloki — cykl 1, cykl 2, …
// Każdy blok to tak naprawdę osobny REKORD (kto / od / do / długość).
// Zwykłe szukanie po tekście widzi tylko komórki, więc nie potrafi odpowiedzieć na
// pytania typu „gdzie ktoś ZACZĄŁ, ale nie SKOŃCZYŁ", „ostatni wpis tej osoby",
// „co wisi dłużej niż 30 dni".
//
// Ten moduł dokłada nad istniejącym filtrem cienką warstwę:
//   1) model: z wykrytych bloków (getActiveRepeatingGroup) + ról kolumn (start/koniec/
//      osoba/długość) buduje dla wiersza listę rekordów,
//   2) stan: zestaw „trybów auto" (chipy w UI albo tokeny @wtoku / @ostatni w polu szukania),
//   3) korelacja: gdy tryb auto jest włączony, tekst filtra sprawdzamy W OBRĘBIE rekordu
//      (prefiks wiersza + kolumny bloku), więc „Kowalski + w toku" znaczy „Kowalski
//      w bloku, który jest w toku", a nie „Kowalski gdziekolwiek ORAZ cokolwiek w toku".
//
// Wykrywanie ról jest heurystyczne (nagłówki → dane), ale ZAWSZE można je nadpisać
// ręcznie w panelu — plik użytkownika nie musi mieć kolumn „od"/„do".

const SMART_COLS_STORE_KEY = "swb-smart-cycle-cols";
const SMART_STATE_RECORD_LEVEL = ["open", "closed", "orphan"];
const SMART_ROW_FLAGS = ["allClosed", "none"];

// Stan trybów auto. `states` łączy się przez OR (na poziomie rekordu),
// `rowFlags` przez AND (na poziomie wiersza), `only` i `minDays` zawężają rekordy.
let smartFilterState = {
  states: [],
  rowFlags: [],
  only: "",
  minDays: null,
};

// rowIndex0 -> Set(blockIndex) — rekordy, dzięki którym wiersz przeszedł ostatni filtr.
// Używane przez widok Wide-to-Long (pokazuje tylko trafione cykle) i podświetlanie komórek.
let smartMatchedBlocksByRow = new Map();

// Stan efektywny ostatniego przebiegu filtra (chipy + tokeny @ z pól szukania).
// Widok long i podświetlanie pytają o niego, żeby tokeny działały tak samo jak chipy.
let smartEffectiveState = null;

let _smartModelCache = null;
let _smartModelKey = "";
let _smartManualVersion = 0;

// ── Model: bloki + role kolumn ─────────────────────────────────────────────────

function smartManualOverrides() {
  try {
    const raw = localStorage.getItem(SMART_COLS_STORE_KEY);
    return raw ? (JSON.parse(raw) || {}) : {};
  } catch {
    return {};
  }
}

function smartOverrideKey() {
  return `${currentFileName || ""}${currentSheetName || ""}${currentHeaderRow || 0}`;
}

function smartReadOverride() {
  const all = smartManualOverrides();
  const hit = all[smartOverrideKey()];
  return hit && typeof hit === "object" ? hit : null;
}

function smartWriteOverride(patch) {
  const all = smartManualOverrides();
  const key = smartOverrideKey();
  const next = { ...(all[key] || {}), ...patch };
  Object.keys(next).forEach((k) => {
    if (next[k] === null || next[k] === undefined || next[k] === "") delete next[k];
  });
  if (Object.keys(next).length) all[key] = next;
  else delete all[key];
  try {
    localStorage.setItem(SMART_COLS_STORE_KEY, JSON.stringify(all));
  } catch {
    // brak miejsca / tryb prywatny — nadpisanie zadziała tylko w tej sesji
  }
  _smartManualVersion += 1;
  _smartModelKey = "";
}

function smartClearOverride() {
  const all = smartManualOverrides();
  delete all[smartOverrideKey()];
  try {
    localStorage.setItem(SMART_COLS_STORE_KEY, JSON.stringify(all));
  } catch {
    // jak wyżej
  }
  _smartManualVersion += 1;
  _smartModelKey = "";
}

function smartBlockHeaderInfos(group) {
  const first = group.blocks[0];
  const headers = Array.isArray(first.headers) ? first.headers : [];
  return headers.map((header, idx) => {
    const base = parseRepeatedHeader(header)?.base || cleanSectionLabel(header) || String(header || "");
    return { idx, header, base, norm: normalizeAnalysisKey(base) };
  });
}

// Czy komórka w ogóle coś ma — bez formatowania daty (getDisplayValue jest droższe).
function smartCellFilled(row, i) {
  if (!row || i < 0 || i >= row.values.length) return false;
  const v = row.values[i];
  if (v === null || v === undefined) return false;
  if (typeof v === "string") return v.trim() !== "";
  if (typeof v === "number") return true;
  if (typeof v === "boolean") return true;
  if (v instanceof Date) return true;
  return String(getDisplayValue(row, i)).trim() !== "";
}

// Fallback z danych: w obrębie bloku szukamy kolumn, które faktycznie wyglądają na daty.
// Skanujemy próbkę wierszy, żeby nie płacić za cały arkusz.
function smartInferRolesFromData(group) {
  const rows = Array.isArray(baseRows) ? baseRows : [];
  const span = group.blocks[0]?.span || 0;
  if (!span) return { startIdx: -1, endIdx: -1 };
  const dateHits = new Array(span).fill(0);
  const textHits = new Array(span).fill(0);
  const sample = rows.slice(0, 200);

  sample.forEach((row) => {
    group.blocks.forEach((block) => {
      for (let k = 0; k < span; k++) {
        const col = block.startIndex + k;
        if (!smartCellFilled(row, col)) continue;
        const raw = row.values[col];
        if (raw instanceof Date) { dateHits[k] += 1; continue; }
        const display = getDisplayValue(row, col);
        const ds = String(display);
        const looksDate = /\d[-/.]\d/.test(ds) || (/\d/.test(ds) && /[a-ząćęłńóśźż]/i.test(ds) && parseDateFlexible(raw ?? display) instanceof Date);
        if (looksDate && parseDateFlexible(raw ?? display) instanceof Date) dateHits[k] += 1;
        else if (/[a-ząćęłńóśźż]/i.test(ds)) textHits[k] += 1;
      }
    });
  });

  const dateCols = dateHits
    .map((count, idx) => ({ idx, count }))
    .filter((entry) => entry.count > 0)
    .sort((a, b) => b.count - a.count || a.idx - b.idx)
    .slice(0, 2)
    .sort((a, b) => a.idx - b.idx);

  const entityCandidate = textHits
    .map((count, idx) => ({ idx, count }))
    .filter((entry) => entry.count > 0)
    .sort((a, b) => b.count - a.count || a.idx - b.idx)[0];

  return {
    startIdx: dateCols[0] ? dateCols[0].idx : -1,
    endIdx: dateCols[1] ? dateCols[1].idx : -1,
    entityIdx: entityCandidate ? entityCandidate.idx : -1,
  };
}

function smartDetectRoles(group) {
  const infos = smartBlockHeaderInfos(group);
  const byHeader = {
    startIdx: findAnalysisColumnIndex(infos, [
      (norm) => norm === "od" || norm === "data od" || norm === "start",
      (norm) => /(^| )(od|start|from|poczatek|rozpoczecie|wydano|pobranie|wyjazd)( |$)/.test(norm),
    ]),
    endIdx: findAnalysisColumnIndex(infos, [
      (norm) => norm === "do" || norm === "data do" || norm === "koniec" || norm === "end",
      (norm) => /(^| )(do|koniec|zakonczenie|zakonczono|end|until|zwrot|oddano|powrot)( |$)/.test(norm),
    ]),
    entityIdx: findAnalysisColumnIndex(infos, [
      (norm) => /(^| )(imie|nazwisko|osoba|pracownik|opiekun|wlasciciel|owner|assignee|user|operator|kierowca)( |$)/.test(norm),
      (norm) => norm.includes("imie") || norm.includes("nazwisk"),
    ]),
    durationIdx: findAnalysisColumnIndex(infos, [
      (norm) => norm.includes("dlugosc") || norm.includes("czas trwania") || norm === "czas",
      (norm) => /(^| )(duration|days|dni)( |$)/.test(norm),
    ]),
  };

  const needsData = byHeader.startIdx < 0 || byHeader.endIdx < 0;
  const inferred = needsData ? smartInferRolesFromData(group) : {};

  const roles = {
    startIdx: byHeader.startIdx >= 0 ? byHeader.startIdx : (inferred.startIdx ?? -1),
    endIdx: byHeader.endIdx >= 0 ? byHeader.endIdx : (inferred.endIdx ?? -1),
    entityIdx: byHeader.entityIdx >= 0 ? byHeader.entityIdx : (inferred.entityIdx ?? -1),
    durationIdx: byHeader.durationIdx,
    source: (byHeader.startIdx >= 0 && byHeader.endIdx >= 0) ? "headers" : "data",
  };

  // „od" bez „do" (albo odwrotnie) nadal jest użyteczne: brak końca = w toku.
  if (roles.startIdx >= 0 && roles.startIdx === roles.endIdx) roles.endIdx = -1;
  return { roles, infos };
}

function smartModelKey() {
  const group = typeof getActiveRepeatingGroup === "function" ? getActiveRepeatingGroup() : null;
  if (!group || !Array.isArray(group.blocks) || group.blocks.length < 2) return "";
  const stamp = typeof sheetDataStamp === "number" ? sheetDataStamp : 0;
  return [
    currentSheetName || "",
    currentHeaderRow || 0,
    stamp,
    group.blocks.length,
    group.blocks[0]?.startIndex ?? -1,
    group.blocks[0]?.span ?? 0,
    _smartManualVersion,
  ].join("|");
}

// Model albo null, gdy arkusz nie ma powtarzalnych bloków (tryby auto są wtedy wyłączone).
function getSmartModel() {
  const key = smartModelKey();
  if (!key) {
    _smartModelKey = "";
    _smartModelCache = null;
    return null;
  }
  if (key === _smartModelKey && _smartModelCache) return _smartModelCache;

  const group = getActiveRepeatingGroup();
  const { roles, infos } = smartDetectRoles(group);
  const override = smartReadOverride() || {};
  const applyOverride = (name) => {
    const v = override[name];
    if (v === "none") return -1;
    const n = Number(v);
    return Number.isInteger(n) && n >= 0 && n < (group.blocks[0]?.span || 0) ? n : roles[name];
  };

  const model = {
    group,
    blocks: group.blocks,
    prefixCount: Math.max(0, Number(group.prefixCount) || 0),
    span: group.blocks[0]?.span || 0,
    columns: infos,
    startIdx: applyOverride("startIdx"),
    endIdx: applyOverride("endIdx"),
    entityIdx: applyOverride("entityIdx"),
    durationIdx: roles.durationIdx,
    manual: Object.keys(override).length > 0,
    source: Object.keys(override).length ? "manual" : roles.source,
  };
  model.hasStateColumns = model.startIdx >= 0 || model.endIdx >= 0;
  // Kolumny bloku i zakres widzenia tekstu zależą TYLKO od modelu — liczymy raz,
  // a nie dla każdego wiersza z osobna (przy 500+ wierszach × 8 blokach to widać).
  model.blockCols = model.blocks.map((block) => {
    const cols = [];
    for (let col = block.startIndex; col <= block.endIndex; col++) cols.push(col);
    return cols;
  });
  model.blockScope = model.blocks.map((block, i) => {
    const allowed = new Set();
    for (let k = 0; k < model.prefixCount; k++) allowed.add(k);
    model.blockCols[i].forEach((c) => allowed.add(c));
    return allowed;
  });

  _smartModelKey = key;
  _smartModelCache = model;
  return model;
}

function smartInvalidateModel() {
  _smartModelKey = "";
  _smartModelCache = null;
}

// ── Rekordy wiersza ────────────────────────────────────────────────────────────

function smartRecordDays(row, rec, model) {
  if (rec._days !== undefined) return rec._days;
  let days = null;
  const startCol = model.startIdx >= 0 ? rec.startCol : -1;
  const endCol = model.endIdx >= 0 ? rec.endCol : -1;
  const start = startCol >= 0 ? parseDateFlexible(row.values[startCol] ?? getDisplayValue(row, startCol)) : null;
  const end = endCol >= 0 ? parseDateFlexible(row.values[endCol] ?? getDisplayValue(row, endCol)) : null;
  if (start instanceof Date) days = diffDays(start, end instanceof Date ? end : new Date());
  if (days === null && model.durationIdx >= 0) {
    const col = rec.blockStart + model.durationIdx;
    days = parseDurationDaysFlexible(row.values[col] ?? getDisplayValue(row, col));
  }
  rec._days = days;
  return days;
}

// Lista rekordów wiersza — po jednym na blok. Puste bloki też są na liście
// (ze stanem "empty"), bo dzięki temu liczymy „pierwszy/ostatni wypełniony".
function buildSmartRecords(row, model = getSmartModel()) {
  if (!model || !row) return [];
  const records = [];
  let lastFilled = -1;
  let firstFilled = -1;

  model.blocks.forEach((block, blockIndex) => {
    const cols = model.blockCols[blockIndex];
    let filled = false;
    for (let i = 0; i < cols.length; i++) {
      if (smartCellFilled(row, cols[i])) { filled = true; break; }
    }
    const startCol = model.startIdx >= 0 ? block.startIndex + model.startIdx : -1;
    const endCol = model.endIdx >= 0 ? block.startIndex + model.endIdx : -1;
    const hasStart = startCol >= 0 && smartCellFilled(row, startCol);
    const hasEnd = endCol >= 0 && smartCellFilled(row, endCol);

    let state = "empty";
    if (filled) {
      if (!model.hasStateColumns) state = "filled";
      else if (hasStart && hasEnd) state = "closed";
      else if (hasStart && !hasEnd) state = "open";
      else if (!hasStart && hasEnd) state = "orphan";
      else state = "filled"; // blok ma dane, ale nie w kolumnach start/koniec
    }

    if (filled) {
      lastFilled = blockIndex;
      if (firstFilled < 0) firstFilled = blockIndex;
    }

    records.push({
      blockIndex,
      label: block.label,
      cols,
      blockStart: block.startIndex,
      blockEnd: block.endIndex,
      startCol,
      endCol,
      hasStart,
      hasEnd,
      filled,
      state,
      isFirst: false,
      isLast: false,
    });
  });

  records.forEach((rec) => {
    rec.isFirst = rec.blockIndex === firstFilled;
    rec.isLast = rec.blockIndex === lastFilled;
  });
  return records;
}

// ── Stan filtra ────────────────────────────────────────────────────────────────

function smartActiveState() {
  return smartEffectiveState || smartFilterState;
}

function smartNeedsRecordScope(state = smartActiveState()) {
  return !!(state.states.length || state.only || Number.isFinite(state.minDays));
}

function smartFilterIsActive() {
  if (!getSmartModel()) return false;
  const state = smartActiveState();
  return smartNeedsRecordScope(state) || state.rowFlags.length > 0;
}

// Licznik do plakietki „Filtry" — liczymy TYLKO chipy panelu (tokeny widać w polu szukania).
function smartActiveCount() {
  if (!getSmartModel()) return 0;
  return smartFilterState.states.length
    + smartFilterState.rowFlags.length
    + (smartFilterState.only ? 1 : 0)
    + (Number.isFinite(smartFilterState.minDays) ? 1 : 0);
}

function resetSmartFilterState({ silent = true } = {}) {
  smartFilterState = { states: [], rowFlags: [], only: "", minDays: null };
  smartEffectiveState = null;
  smartMatchedBlocksByRow = new Map();
  if (!silent && typeof renderSmartFilterUi === "function") renderSmartFilterUi();
}

function smartToggleValue(listName, value) {
  const list = smartFilterState[listName];
  const idx = list.indexOf(value);
  if (idx >= 0) list.splice(idx, 1);
  else list.push(value);
}

// ── Dopasowanie ────────────────────────────────────────────────────────────────

function smartRecordPasses(row, rec, model, state = smartActiveState()) {
  if (!rec.filled) return false;
  if (state.states.length) {
    // "closed" traktujemy szeroko: blok domknięty. "filled" (blok bez kolumn start/koniec)
    // pasuje tylko wtedy, gdy nie pytamy o konkretny stan.
    if (!state.states.includes(rec.state)) return false;
  }
  if (state.only === "last" && !rec.isLast) return false;
  if (state.only === "first" && !rec.isFirst) return false;
  if (Number.isFinite(state.minDays)) {
    const days = smartRecordDays(row, rec, model);
    if (!Number.isFinite(days) || days < state.minDays) return false;
  }
  return true;
}

function smartRowFlagsPass(records, state = smartActiveState()) {
  if (!state.rowFlags.length) return true;
  const anyFilled = records.some((rec) => rec.filled);
  for (const flag of state.rowFlags) {
    if (flag === "none" && anyFilled) return false;
    if (flag === "allClosed") {
      if (!anyFilled) return false;
      if (records.some((rec) => rec.state === "open" || rec.state === "orphan")) return false;
    }
  }
  return true;
}

// Rekordy wiersza, które przechodzą tryby auto (bez części tekstowej filtra).
function smartMatchingRecords(row, model = getSmartModel(), state = smartActiveState()) {
  if (!model) return [];
  const records = buildSmartRecords(row, model);
  if (!smartRowFlagsPass(records, state)) return [];
  if (!smartNeedsRecordScope(state)) return records.filter((rec) => rec.filled);
  return records.filter((rec) => smartRecordPasses(row, rec, model, state));
}

// Kolumny, w których szukamy tekstu dla danego rekordu: prefiks wiersza (Nr., Teren…)
// + kolumny tego bloku. Przecinamy z wyborem kolumn użytkownika; gdy wybór jest
// rozłączny z rekordem, zostawiamy oryginalny wybór (nie chcemy cicho gubić filtra).
function smartScopeIndexes(indexes, rec, model) {
  const allowed = model.blockScope[rec.blockIndex];
  const scoped = indexes.filter((i) => allowed.has(i));
  return scoped.length ? scoped : indexes;
}

// Wynik zależy od (kryteria, blok) — nie od wiersza — więc w jednym przebiegu
// filtra liczymy go raz na blok i podajemy dalej te same obiekty.
let _smartScopeCache = null;

function smartScopeCriteria(criteria, rec, model) {
  if (!_smartScopeCache) _smartScopeCache = new Map();
  const hit = _smartScopeCache.get(rec.blockIndex);
  if (hit) return hit;
  const scoped = criteria.map((criterion) => (
    criterion.query || (criterion.emptyMode && criterion.emptyMode !== "all")
      ? { ...criterion, indexes: smartScopeIndexes(criterion.indexes, rec, model), blockScope: model.blockScope[rec.blockIndex] }
      : criterion
  ));
  _smartScopeCache.set(rec.blockIndex, scoped);
  return scoped;
}

// Czy dany blok (rekord) ma się pokazać w widoku Wide-to-Long / w agregacji „long".
// Gdy wiersz przeszedł filtr, korzystamy z zapamiętanego dopasowania (uwzględnia tekst);
// w innym wypadku liczymy sam stan.
function smartRecordAccepted(row, blockIndex) {
  if (!smartFilterIsActive()) return true;
  const hit = smartMatchedBlocksByRow.get(row?.rowIndex0);
  if (hit) return hit.has(blockIndex);
  const model = getSmartModel();
  if (!model) return true;
  return smartMatchingRecords(row, model, smartActiveState()).some((rec) => rec.blockIndex === blockIndex);
}

// Start przebiegu filtrowania: zapamiętujemy stan efektywny (chipy + tokeny)
// i czyścimy mapę trafionych bloków.
function smartBeginFilterPass(state = null) {
  smartEffectiveState = state;
  smartMatchedBlocksByRow = new Map();
  _smartScopeCache = null;
}

function smartRememberMatch(row, recs) {
  if (!recs || !recs.length) return;
  smartMatchedBlocksByRow.set(row.rowIndex0, new Set(recs.map((rec) => rec.blockIndex)));
}

// Podświetlanie w trybie auto: zostawiamy tylko komórki z trafionych bloków
// (plus prefiks wiersza) i DOKŁADAMY komórki, które zdecydowały o stanie —
// datę „od" rekordu w toku widać wtedy od razu, nawet bez wpisanego tekstu.
function smartRestrictMatchedCols(row, cols) {
  const model = getSmartModel();
  const hit = smartMatchedBlocksByRow.get(row?.rowIndex0);
  if (!model || !hit || !hit.size) return cols;
  const allowed = new Set();
  for (let i = 0; i < model.prefixCount; i++) allowed.add(i);
  const evidence = new Set();
  model.blocks.forEach((block, blockIndex) => {
    if (!hit.has(blockIndex)) return;
    for (let col = block.startIndex; col <= block.endIndex; col++) allowed.add(col);
    if (model.startIdx >= 0) evidence.add(block.startIndex + model.startIdx);
    if (model.endIdx >= 0 && smartCellFilled(row, block.startIndex + model.endIdx)) evidence.add(block.startIndex + model.endIdx);
  });
  const out = new Set();
  cols.forEach((c) => { if (allowed.has(c)) out.add(c); });
  evidence.forEach((c) => { if (smartCellFilled(row, c)) out.add(c); });
  return out;
}

// ── Tokeny tekstowe (@wtoku, @ostatni, @dni>>30 …) ─────────────────────────────
// Pozwalają wpisać tryb auto wprost w pole szukania — także w szybkim szukaniu,
// gdzie nie ma miejsca na chipy. Token jest MODYFIKATOREM wiersza, nie zwykłym
// termem: wycinamy go z zapytania (razem z osieroconym &&/||) przed parserem.
const SMART_TOKENS = [
  { re: /^(wtoku|w-toku|otwarte|open|inprogress)$/, apply: (st) => { if (!st.states.includes("open")) st.states.push("open"); } },
  { re: /^(zakonczone|zakończone|zamkniete|zamknięte|closed|done)$/, apply: (st) => { if (!st.states.includes("closed")) st.states.push("closed"); } },
  { re: /^(bezstartu|brakstartu|orphan|nostart)$/, apply: (st) => { if (!st.states.includes("orphan")) st.states.push("orphan"); } },
  { re: /^(wszystkozamkniete|wszystkozamknięte|allclosed)$/, apply: (st) => { if (!st.rowFlags.includes("allClosed")) st.rowFlags.push("allClosed"); } },
  { re: /^(niezaczete|niezaczęte|puste|none|nothing)$/, apply: (st) => { if (!st.rowFlags.includes("none")) st.rowFlags.push("none"); } },
  { re: /^(ostatni|last)$/, apply: (st) => { st.only = "last"; } },
  { re: /^(pierwszy|first)$/, apply: (st) => { st.only = "first"; } },
];

function smartApplyToken(state, body) {
  const norm = String(body || "").trim().toLowerCase().replace(/\s+/g, "");
  if (!norm) return false;
  const days = norm.match(/^(dni|dluzejniz|dłużejniż|days|older)(?:>>=?|>=|>|:)?(\d+(?:[.,]\d+)?)$/);
  if (days) {
    const n = Number(days[2].replace(",", "."));
    if (Number.isFinite(n)) { state.minDays = n; return true; }
    return false;
  }
  for (const token of SMART_TOKENS) {
    if (token.re.test(norm)) { token.apply(state); return true; }
  }
  return false;
}

// Zwraca { query, state, found } — zapytanie bez tokenów i stan z nich zbudowany.
function extractSmartTokens(rawQuery) {
  const query = String(rawQuery || "");
  if (query.indexOf("@") < 0) return { query, state: null, found: false };
  const state = { states: [], rowFlags: [], only: "", minDays: null };
  let found = false;
  const cleaned = query.replace(/@[^\s&|(){}!]+/g, (match) => {
    if (smartApplyToken(state, match.slice(1))) { found = true; return " "; }
    return match;
  });
  if (!found) return { query, state: null, found: false };
  const tidy = cleaned
    .replace(/(^|\s)(&&|\|\|)(\s*(&&|\|\|))*(\s|$)/g, " ")
    .replace(/(&&|\|\|)\s*$/g, " ")
    .replace(/^\s*(&&|\|\|)/g, " ")
    .replace(/\s+/g, " ")
    .trim();
  return { query: tidy, state, found: true };
}

// Stan efektywny dla przebiegu filtrowania: chipy z panelu + tokeny z zapytań.
function resolveEffectiveSmartState(queries = []) {
  const merged = {
    states: smartFilterState.states.slice(),
    rowFlags: smartFilterState.rowFlags.slice(),
    only: smartFilterState.only,
    minDays: smartFilterState.minDays,
  };
  let tokenFound = false;
  const cleanedQueries = queries.map((q) => {
    const parsed = extractSmartTokens(q);
    if (!parsed.found) return q;
    tokenFound = true;
    parsed.state.states.forEach((s) => { if (!merged.states.includes(s)) merged.states.push(s); });
    parsed.state.rowFlags.forEach((f) => { if (!merged.rowFlags.includes(f)) merged.rowFlags.push(f); });
    if (parsed.state.only) merged.only = parsed.state.only;
    if (Number.isFinite(parsed.state.minDays)) merged.minDays = parsed.state.minDays;
    return parsed.query;
  });
  return { state: merged, queries: cleanedQueries, tokenFound };
}

// ── UI: chipy, status, konfiguracja kolumn ─────────────────────────────────────

const smartFilterPanelEl = document.getElementById("panel-smart-filters");
const smartFilterStatusEl = document.getElementById("smartFilterStatus");
const smartFilterChipsEl = document.getElementById("smartFilterChips");
const smartQuickChipsEl = document.getElementById("quickSearchSmartChips");
const smartFilterBadgeEl = document.getElementById("smartFilterBadge");
const smartDaysFieldEl = document.getElementById("smartDaysField");
const smartMinDaysEl = document.getElementById("smartMinDays");
const smartFilterConfigEl = document.getElementById("smartFilterConfig");
const smartStartColEl = document.getElementById("smartStartCol");
const smartEndColEl = document.getElementById("smartEndCol");
const smartEntityColEl = document.getElementById("smartEntityCol");
const smartResetColsBtn = document.getElementById("smartResetColsBtn");
const smartFilterHintEl = document.getElementById("smartFilterHint");

// Pełny zestaw chipów w panelu; `compact` trafia też do okna szybkiego szukania,
// gdzie liczy się każdy piksel (telefon).
const SMART_CHIPS = [
  { key: "open", group: "states", label: "smartChipOpen", hint: "smartChipOpenHint", compact: true, needsState: true },
  { key: "closed", group: "states", label: "smartChipClosed", hint: "smartChipClosedHint", compact: true, needsState: true },
  { key: "orphan", group: "states", label: "smartChipOrphan", hint: "smartChipOrphanHint", needsState: true },
  { key: "allClosed", group: "rowFlags", label: "smartChipAllClosed", hint: "smartChipAllClosedHint", needsState: true },
  { key: "none", group: "rowFlags", label: "smartChipNone", hint: "smartChipNoneHint" },
  { key: "last", group: "only", label: "smartChipLast", hint: "smartChipLastHint", compact: true },
  { key: "first", group: "only", label: "smartChipFirst", hint: "smartChipFirstHint" },
  { key: "days", group: "days", label: "smartChipDays", hint: "smartChipDaysHint", compact: true, needsState: true },
];

function smartChipIsActive(chip) {
  if (chip.group === "states") return smartFilterState.states.includes(chip.key);
  if (chip.group === "rowFlags") return smartFilterState.rowFlags.includes(chip.key);
  if (chip.group === "only") return smartFilterState.only === chip.key;
  if (chip.group === "days") return Number.isFinite(smartFilterState.minDays);
  return false;
}

function smartChipToggle(chip) {
  if (chip.group === "states") smartToggleValue("states", chip.key);
  else if (chip.group === "rowFlags") smartToggleValue("rowFlags", chip.key);
  else if (chip.group === "only") smartFilterState.only = smartFilterState.only === chip.key ? "" : chip.key;
  else if (chip.group === "days") {
    if (Number.isFinite(smartFilterState.minDays)) smartFilterState.minDays = null;
    else {
      const raw = parseFloat(String(smartMinDaysEl?.value || "30").replace(",", "."));
      smartFilterState.minDays = Number.isFinite(raw) && raw > 0 ? raw : 30;
    }
  }
}

function smartApplyChange() {
  if (!Array.isArray(currentHeaders) || !currentHeaders.length) {
    renderSmartFilterUi();
    return;
  }
  // Chip to świadoma decyzja „pokaż mi tylko te wiersze" — jak klik „Filtruj".
  if (smartActiveCount() > 0) {
    filtersCommitted = true;
    quickSearchHighlightMode = false;
    quickSearchCellsMode = false;
  }
  applyFilters();
  sortRows();
  renderSmartFilterUi();
  scheduleViewRefresh({ table: true, analyses: true, filterBadge: true });
}

function smartBuildChip(chip, model) {
  const btn = document.createElement("button");
  btn.type = "button";
  btn.className = "chip smart-chip";
  btn.tabIndex = 0; // jawnie: Safari pomija w nawigacji Tab przyciski bez tabindex
  btn.dataset.smartChip = chip.key;
  btn.textContent = t(chip.label);
  btn.setAttribute("aria-pressed", smartChipIsActive(chip) ? "true" : "false");
  btn.classList.toggle("active", smartChipIsActive(chip));
  const hint = t(chip.hint);
  if (hint) {
    btn.setAttribute("data-hint", hint);
    btn.setAttribute("data-hint-delay", "1.1");
    btn.setAttribute("data-hint-touch", "on");
  }
  const unsupported = !model || (chip.needsState && !model.hasStateColumns);
  btn.disabled = !!unsupported;
  if (unsupported) btn.title = t("smartChipUnavailable");
  btn.addEventListener("click", () => {
    smartChipToggle(chip);
    smartApplyChange();
  });
  return btn;
}

function smartFillColumnSelect(select, value, { allowNone = true } = {}) {
  if (!select) return;
  const model = getSmartModel();
  select.replaceChildren();
  if (allowNone) {
    const none = document.createElement("option");
    none.value = "none";
    none.textContent = t("smartColNone");
    select.appendChild(none);
  }
  (model?.columns || []).forEach((col) => {
    const option = document.createElement("option");
    option.value = String(col.idx);
    option.textContent = col.base || col.header || `#${col.idx + 1}`;
    select.appendChild(option);
  });
  select.value = value >= 0 ? String(value) : "none";
  select.disabled = !model;
}

function smartStatusText(model) {
  if (!model) return t("smartStatusNoBlocks");
  const startLabel = model.startIdx >= 0 ? (model.columns[model.startIdx]?.base || "?") : t("smartColNone");
  const endLabel = model.endIdx >= 0 ? (model.columns[model.endIdx]?.base || "?") : t("smartColNone");
  const base = t("smartStatusReady", { blocks: model.blocks.length, span: model.span });
  const roles = t("smartStatusRoles", { start: startLabel, end: endLabel });
  const src = model.source === "manual" ? t("smartSourceManual")
    : model.source === "data" ? t("smartSourceData")
    : t("smartSourceHeaders");
  return `${base} · ${roles} · ${src}`;
}

function renderSmartFilterUi() {
  const model = getSmartModel();
  const available = !!model;

  if (smartFilterPanelEl) {
    smartFilterPanelEl.classList.toggle("smart-unavailable", !available);
  }
  if (smartFilterStatusEl) {
    smartFilterStatusEl.textContent = smartStatusText(model);
    smartFilterStatusEl.classList.toggle("warn", !available);
  }
  if (smartFilterChipsEl) {
    smartFilterChipsEl.replaceChildren(...SMART_CHIPS.map((chip) => smartBuildChip(chip, model)));
    const clear = document.createElement("button");
    clear.type = "button";
    clear.className = "chip smart-chip-clear";
    clear.tabIndex = 0;
    clear.textContent = t("smartClear");
    clear.disabled = smartActiveCount() === 0;
    clear.addEventListener("click", () => {
      resetSmartFilterState();
      smartApplyChange();
    });
    smartFilterChipsEl.appendChild(clear);
  }
  // Most do widoku long: „pokaż mi tylko te cykle, a nie całe wiersze".
  // Widok long i tak pokazuje wyłącznie trafione bloki, więc to naturalny drugi krok.
  if (smartFilterChipsEl && available && typeof canUseLongView === "function" && canUseLongView()) {
    const longBtn = document.createElement("button");
    longBtn.type = "button";
    longBtn.className = "chip smart-chip-long";
    longBtn.tabIndex = 0;
    const isLong = typeof tableViewMode !== "undefined" && tableViewMode === "long";
    longBtn.textContent = isLong ? t("smartBackToRows") : t("smartShowCycles");
    longBtn.classList.toggle("active", isLong);
    longBtn.setAttribute("aria-pressed", isLong ? "true" : "false");
    longBtn.setAttribute("data-hint", t("smartShowCyclesHint"));
    longBtn.setAttribute("data-hint-delay", "1.1");
    longBtn.setAttribute("data-hint-touch", "on");
    longBtn.addEventListener("click", () => {
      tableViewMode = tableViewMode === "long" ? "wide" : "long";
      manualColumnWidths = {};
      renderSmartFilterUi();
      scheduleViewRefresh({ table: true, analyses: true, sync: true });
    });
    smartFilterChipsEl.appendChild(longBtn);
  }
  if (smartQuickChipsEl) {
    const showQuick = available && (typeof currentHeaders !== "undefined") && currentHeaders.length;
    smartQuickChipsEl.classList.toggle("hidden", !showQuick);
    if (showQuick) {
      smartQuickChipsEl.replaceChildren(...SMART_CHIPS.filter((chip) => chip.compact).map((chip) => smartBuildChip(chip, model)));
      smartApplyRovingTabindex();
    }
  }
  if (smartDaysFieldEl) {
    smartDaysFieldEl.classList.toggle("hidden", !Number.isFinite(smartFilterState.minDays));
    if (smartMinDaysEl && Number.isFinite(smartFilterState.minDays)) {
      smartMinDaysEl.value = String(smartFilterState.minDays);
    }
  }
  if (smartFilterBadgeEl) {
    const count = smartActiveCount();
    smartFilterBadgeEl.textContent = String(count);
    smartFilterBadgeEl.classList.toggle("hidden", count === 0);
  }
  if (smartFilterConfigEl) smartFilterConfigEl.classList.toggle("hidden", !available);
  smartFillColumnSelect(smartStartColEl, model ? model.startIdx : -1);
  smartFillColumnSelect(smartEndColEl, model ? model.endIdx : -1);
  smartFillColumnSelect(smartEntityColEl, model ? model.entityIdx : -1);
  if (smartResetColsBtn) smartResetColsBtn.disabled = !model || !model.manual;
  if (smartFilterHintEl) {
    smartFilterHintEl.textContent = available && !model.hasStateColumns
      ? t("smartHintNoStateCols")
      : t("smartHintTokens");
  }
}

// Pasek chipów w oknie szybkiego szukania to JEDEN przystanek Tab (roving tabindex),
// wybór strzałkami — dokładnie tak, jak grupy segmentów zastępujące tam <select>.
// Inaczej cztery chipy dokładałyby cztery przystanki do i tak długiego obiegu okna.
function smartApplyRovingTabindex(activeIndex = 0) {
  if (!smartQuickChipsEl) return;
  const chips = Array.from(smartQuickChipsEl.querySelectorAll("button"));
  if (!chips.length) return;
  const idx = Math.max(0, Math.min(activeIndex, chips.length - 1));
  chips.forEach((chip, i) => { chip.tabIndex = i === idx ? 0 : -1; });
  smartQuickChipsEl.setAttribute("role", "group");
}

if (smartQuickChipsEl) {
  smartQuickChipsEl.addEventListener("keydown", (e) => {
    if (e.key !== "ArrowLeft" && e.key !== "ArrowRight" && e.key !== "Home" && e.key !== "End") return;
    const chips = Array.from(smartQuickChipsEl.querySelectorAll("button"));
    if (!chips.length) return;
    const current = chips.indexOf(document.activeElement);
    if (current < 0) return;
    e.preventDefault();
    const next = e.key === "Home" ? 0
      : e.key === "End" ? chips.length - 1
      : e.key === "ArrowRight" ? (current + 1) % chips.length
      : (current - 1 + chips.length) % chips.length;
    smartApplyRovingTabindex(next);
    chips[next].focus();
  });
}

if (smartMinDaysEl) {
  smartMinDaysEl.addEventListener("change", () => {
    const raw = parseFloat(String(smartMinDaysEl.value || "").replace(",", "."));
    if (!Number.isFinite(raw) || raw <= 0) return;
    smartFilterState.minDays = raw;
    smartApplyChange();
  });
}

[[smartStartColEl, "startIdx"], [smartEndColEl, "endIdx"], [smartEntityColEl, "entityIdx"]].forEach(([select, role]) => {
  if (!select) return;
  select.addEventListener("change", () => {
    smartWriteOverride({ [role]: select.value });
    smartInvalidateModel();
    smartApplyChange();
  });
});

if (smartResetColsBtn) {
  smartResetColsBtn.addEventListener("click", () => {
    smartClearOverride();
    smartInvalidateModel();
    smartApplyChange();
  });
}
