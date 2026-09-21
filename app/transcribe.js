// Tryb spisywania — ekran do PRZEPISYWANIA danych ręcznie na papier.
//
// Po co osobny tryb: eksport (CSV / Druk-PDF) obsługuje wyjście MASZYNOWE.
// Ręczne spisywanie ma inne wąskie gardła: gubienie wiersza po podniesieniu wzroku
// znad kartki, brak śladu „gdzie skończyłem”, kolejność kolumn na ekranie inna niż
// kolejność rubryk na formularzu, gaszący się tablet i dłoń przewijająca tabelę.
//
// Model: JEDEN wiersz naraz jako duża karta pól, wielki przycisk „Spisane i dalej”,
// odhaczanie zapisywane lokalnie (wznowienie po przerwie), Wake Lock i blokada dotyku.
//
// Źródło danych: currentDisplayModel (ten sam, co eksport), więc tryb automatycznie
// respektuje filtry, sortowanie, Wide-to-Long i kolumny wyliczane. Wiersze są
// SNAPSHOTOWANE przy otwarciu — widok pod spodem nie przesunie się pod ręką.

const TR_STORE_KEY = "excel-workbench-transcribe";
const TR_MAX_SCOPES = 12;      // ile plików/arkuszy pamiętamy (reszta wypada, najstarsze pierwsze)
const TR_MAX_DONE = 20000;     // bezpiecznik na rozmiar localStorage
const TR_DEFAULT_FIELDS = 10;  // ile pól proponujemy przy pierwszym otwarciu arkusza
const TR_FONT_STEPS = [1, 2, 3];

const trOverlayEl = document.getElementById("transcribeOverlay");
const trBtn = document.getElementById("transcribeBtn");
const trCardEl = document.getElementById("trCard");
const trEmptyEl = document.getElementById("trEmpty");
const trCounterEl = document.getElementById("trCounter");
const trDoneCountEl = document.getElementById("trDoneCount");
const trProgressBarEl = document.getElementById("trProgressBar");
const trSourceEl = document.getElementById("trSource");
const trPrevBtn = document.getElementById("trPrevBtn");
const trNextBtn = document.getElementById("trNextBtn");
const trMarkBtn = document.getElementById("trMarkBtn");
const trMarkChipEl = document.getElementById("trMarkChip");
const trHideDoneEl = document.getElementById("trHideDone");
const trCloseBtn = document.getElementById("trCloseBtn");
const trFontBtn = document.getElementById("trFontBtn");
const trLockBtn = document.getElementById("trLockBtn");
const trFieldsBtn = document.getElementById("trFieldsBtn");
const trFieldsPanelEl = document.getElementById("trFieldsPanel");
const trFieldsListEl = document.getElementById("trFieldsList");
const trFieldsAllBtn = document.getElementById("trFieldsAllBtn");
const trFieldsNoneBtn = document.getElementById("trFieldsNoneBtn");
const trFieldsDoneBtn = document.getElementById("trFieldsDoneBtn");
const trResetBtn = document.getElementById("trResetBtn");
const trAutoFieldsEl = document.getElementById("trAutoFields");
const trAutoNoteEl = document.getElementById("trAutoNote");
const trInheritEl = document.getElementById("trInherit");
const trInheritNoteEl = document.getElementById("trInheritNote");
const trTouchShieldEl = document.getElementById("trTouchShield");
const trStageEl = document.getElementById("trStage");
const trStageWrapEl = document.getElementById("trStageWrap");
const trScrollRailEl = document.getElementById("trScrollRail");
const trScrollThumbEl = document.getElementById("trScrollThumb");
const trScrollMoreEl = document.getElementById("trScrollMore");
const trScrollMoreTextEl = document.getElementById("trScrollMoreText");
const trUndoBtn = document.getElementById("trUndoBtn");
const trProgressBtn = document.getElementById("trProgressBtn");
const trProgressPanelEl = document.getElementById("trProgressPanel");
const trProgressDoneBtn = document.getElementById("trProgressDoneBtn");
const trStatsEl = document.getElementById("trStats");
const trScopeNoteEl = document.getElementById("trScopeNote");
const trStoreListEl = document.getElementById("trStoreList");
const trStoreClearAllBtn = document.getElementById("trStoreClearAllBtn");
const trNoticeEl = document.getElementById("trNotice");
const trNoticeTextEl = document.getElementById("trNoticeText");
const trNoticeResetBtn = document.getElementById("trNoticeResetBtn");
const trNoticeKeepBtn = document.getElementById("trNoticeKeepBtn");
const trLiveEl = document.getElementById("trLive");
const trUnmatchedEl = document.getElementById("trUnmatched");
const trUnmatchedTitleEl = document.getElementById("trUnmatchedTitle");
const trUnmatchedHintEl = document.getElementById("trUnmatchedHint");
const trUnmatchedListEl = document.getElementById("trUnmatchedList");
const trUnmatchedToggleEl = document.getElementById("trUnmatchedToggle");

let trIsOpen = false;
let trRows = [];             // snapshot wierszy z modelu widoku
let trHeaders = [];
let trRowHeadFormatter = null;
let trFieldOrder = [];       // WSZYSTKIE indeksy kolumn w kolejności ustawionej przez użytkownika
let trSelected = new Set();  // które z nich trafiają na kartę
let trDone = new Set();      // klucze wierszy już spisanych
let trOrder = [];            // indeksy do trRows widoczne w bieżącym trybie (hideDone)
let trPos = 0;
let trHideDone = false;
let trAutoFields = false;
let trInheritOn = false;          // „dziedzicz z góry" — master switch
let trInheritCols = new Set();    // kolumny, które przenoszą wartość w dół
let trMergeCols = new Set();      // kolumny objęte PIONOWYM scaleniem (auto-podpowiedź + znacznik)
let trMergeRanges = new Map();    // kolumna -> zakresy scaleń = GRANICE przenoszenia
let trInheritMap = new Map();     // `col:rowIndex0` -> { text, from }
let trLongMode = false;           // Wide-to-Long: dziedziczenie nie ma tam sensu
let trFont = 2;
let trLocked = false;
let trScope = "";
let trWakeLock = null;
let trReturnFocusEl = null;
let trBulkMode = false;       // seria szybkiego odhaczania — wstrzymuje zapis do localStorage
let trVolatileCols = new Set(); // kolumny liczone „na dziś" (TODAY) — poza odciskiem wiersza
let trDoneSigs = new Map();     // klucz wiersza -> odcisk treści (przeżywa przesunięcie wierszy)
let trSigCache = new Map();     // rowIndex0 -> odcisk (liczony leniwie)
let trFingerprint = null;       // odcisk arkusza z TEJ sesji
let trChangeInfo = null;        // { level, savedRows, rows, moved, lost, savedAt } albo null
let trUnmatched = [];           // ✓ których nie dało się dopasować: [{ prev, best, score }]
let trScrollRaf = 0;
let trHoldTimer = 0;          // odliczanie do startu turbo (przytrzymanie)
let trHoldProgressTimer = 0;  // animacja paska „ładowania" przytrzymania
let trTurboTimer = 0;         // pętla szybkiego odhaczania
let trTurboCount = 0;
let trTurboSource = "";       // "key" albo "pointer" — kto trzyma, ten puszcza
let trBurstKeys = [];         // klucze odhaczone w ostatniej serii (do cofnięcia)
let trUndoTimer = 0;

// ── Trwałość ────────────────────────────────────────────────────────────────
// Jeden klucz, w środku mapa „zakresów” (plik + arkusz + tryb widoku). Dzięki temu
// wracasz do tego samego pliku i zastajesz swoje ✓ oraz swój układ pól.

function trLoadStore() {
  try {
    const parsed = JSON.parse(localStorage.getItem(TR_STORE_KEY) || "{}");
    if (!parsed || typeof parsed !== "object") return { scopes: {} };
    if (!parsed.scopes || typeof parsed.scopes !== "object") parsed.scopes = {};
    return parsed;
  } catch {
    return { scopes: {} };
  }
}

function trSaveStore(store) {
  try {
    const scopes = store.scopes || {};
    const keys = Object.keys(scopes);
    if (keys.length > TR_MAX_SCOPES) {
      keys
        .sort((a, b) => (scopes[a]?.ts || 0) - (scopes[b]?.ts || 0))
        .slice(0, keys.length - TR_MAX_SCOPES)
        .forEach((k) => { delete scopes[k]; });
    }
    localStorage.setItem(TR_STORE_KEY, JSON.stringify(store));
  } catch {
    // Brak miejsca. Odciski kolumnowe i podglądy to największa część zapisu, a zarazem
    // jedyna, bez której tryb DALEJ działa (zostaje dopasowanie po hashu całego wiersza).
    // Więc zanim stracimy ✓, zrzucamy balast — najpierw ze starszych zakresów, potem
    // ze wszystkich. Cichy brak zapisu byłby tu najgorszy: użytkownik traci postęp
    // godzinnego przepisywania i nie dowiaduje się o tym.
    const scopes = store.scopes || {};
    const order = Object.keys(scopes).sort((a, b) => (scopes[a]?.ts || 0) - (scopes[b]?.ts || 0));
    const strip = (k) => {
      const rec = scopes[k];
      if (!rec) return;
      rec.cols = [];
      rec.doneCells = [];
      rec.donePrev = [];
    };
    for (let i = 0; i <= order.length; i++) {
      if (i < order.length) strip(order[i]);
      try {
        localStorage.setItem(TR_STORE_KEY, JSON.stringify(store));
        return;
      } catch { /* dalej za duże — zrzucamy kolejny zakres */ }
    }
    /* prywatne okno / twardy brak miejsca — tryb działa dalej, tylko bez pamięci */
  }
}

function trPersist() {
  if (!trScope || trBulkMode) return; // w trakcie serii zapisujemy RAZ, na końcu
  const store = trLoadStore();
  store.font = trFont;
  store.hideDone = trHideDone;
  const doneList = Array.from(trDone).slice(0, TR_MAX_DONE);
  const prev = store.scopes?.[trScope] || null;
  store.scopes[trScope] = {
    order: trFieldOrder.slice(),
    sel: Array.from(trSelected),
    auto: trAutoFields,
    inherit: trInheritOn,
    inheritCols: Array.from(trInheritCols),
    done: doneList,
    // Odciski TREŚCI odhaczonych wierszy — równolegle do `done`. To one pozwalają odnaleźć
    // te same wiersze, gdy plik urośnie albo ktoś przestawi kolejność.
    doneSig: doneList.map((key) => trDoneSigs.get(key) || ""),
    // Odcisk KOLUMNOWY (v2) + czytelny podglad wiersza. Podglad jest jedyna rzecza,
    // po ktorej da sie pokazac uzytkownikowi, CZEGO nie udalo sie potem dopasowac.
    // Przy bardzo duzych zapisach odpuszczamy oba — localStorage ma ~5 MB na origin,
    // a tryb ma dzialac dalej (wtedy zostaje samo dopasowanie po hashu calego wiersza).
    cols: doneList.length <= TR_SIG_MAX_ROWS ? trSigCols.map((c) => c.name) : [],
    doneCells: doneList.length <= TR_SIG_MAX_ROWS ? doneList.map((key) => trDoneCells.get(key) || "") : [],
    donePrev: doneList.length <= TR_SIG_MAX_ROWS ? doneList.map((key) => trDonePrev.get(key) || "") : [],
    sig: trFingerprint,
    volCols: Array.from(trVolatileCols),
    rowsTotal: trRows.length,
    cursor: trCurrentKey(),
    started: prev?.started || Date.now(),
    ts: Date.now(),
  };
  trSaveStore(store);
}

function trScopeKey(model) {
  const file = currentFileName || "?";
  const sheet = (typeof sheetSelect !== "undefined" && sheetSelect?.value) || "?";
  return `${file}::${sheet}::${model.mode || "wide"}`;
}

// ── Odciski: „czy to nadal ten sam arkusz?" ─────────────────────────────────
//
// Klucz wiersza to `wide:<numer wiersza>` — czyli POZYCJA. Wystarczy, że ktoś wstawi
// w Excelu jeden wiersz na górze, a wszystkie ✓ przesuwają się o jeden i cicho lądują
// na cudzych wierszach. Przy przepisywaniu na papier to najgorszy możliwy błąd: wiersz,
// którego nikt nie przepisał, wygląda na zrobiony.
//
// Dlatego obok kluczy zapisujemy ODCISK TREŚCI każdego odhaczonego wiersza. Gdy plik się
// zmieni, ✓ przenosimy po treści, a nie po pozycji — i mówimy wprost, ile się udało.
//
// Z odcisku wiersza WYPADAJĄ kolumny liczone „na dziś" (formuły z TODAY() — patrz znacznik
// przeliczenia w tabeli). Inaczej ten sam wiersz miałby jutro inny odcisk i mechanizm
// psułby się sam z siebie, raz na dobę.

// Separator pol w odcisku wiersza: znak sterujacy, ktory nie wystapi w tresci komorki.
const TR_SEP = String.fromCharCode(1);

function trHash(text) {
  let h = 0x811c9dc5;
  const s = String(text);
  for (let i = 0; i < s.length; i++) {
    h ^= s.charCodeAt(i);
    h = Math.imul(h, 0x01000193) >>> 0;
  }
  return h.toString(36);
}

// Kolumny zmienne: te, w których przeliczenie na dziś dało inny wynik niż plik.
// Zapamiętujemy je w zakresie, bo przy ODZNACZONYM „Przeliczaj formuły z datą"
// nie da się ich wykryć — a odcisk musi wychodzić tak samo w obie strony.
function trCollectVolatileCols(savedCols) {
  const out = new Set(Array.isArray(savedCols) ? savedCols.filter((n) => Number.isInteger(n)) : []);
  const src = Array.isArray(baseRows) ? baseRows : [];
  for (const row of src) {
    if (!row?.recalcCells) continue;
    Object.keys(row.recalcCells).forEach((k) => {
      const n = parseInt(k, 10);
      if (Number.isInteger(n)) out.add(n);
    });
  }
  return out;
}

// Odcisk v1 (cały wiersz, kolumny po POZYCJI). Zostaje wyłącznie po to, żeby czytać
// zapisy zrobione przed wprowadzeniem odcisku kolumnowego — nowych już tak nie liczymy.
function trRowSig(row) {
  if (!row) return "";
  const cacheKey = row.rowIndex0;
  if (cacheKey !== undefined && trSigCache.has(cacheKey)) return trSigCache.get(cacheKey);
  const parts = [];
  const len = Array.isArray(row.values) ? row.values.length : 0;
  for (let i = 0; i < len; i++) {
    if (trVolatileCols.has(i)) continue;
    parts.push(String(getDisplayValue(row, i) ?? ""));
  }
  const sig = trHash(parts.join(TR_SEP));
  if (cacheKey !== undefined) trSigCache.set(cacheKey, sig);
  return sig;
}

// ── Odcisk v2: po NAZWACH kolumn, nie po pozycjach ───────────────────────────
//
// Czego nie umiał v1 (jeden hash całego wiersza, kolumny liczone po numerze):
//   • poprawka JEDNEJ komórki → inny hash → ✓ przepadało,
//   • dodana/przestawiona kolumna → inne hashe WSZYSTKICH wierszy → przepadało wszystko,
//   • inne formatowanie liczby/daty → inny hash, choć treść ta sama.
//
// v2 zapisuje osobny odcisk KAŻDEJ kolumny, podpisany NAZWĄ nagłówka. Przy dopasowaniu
// bierzemy część wspólną kolumn (te, które są w obu plikach) i widzimy, ile z nich się
// zgadza — czyli mamy stopień podobieństwa zamiast zero-jedynkowego hasha. Wartości
// normalizujemy (spacje, wielkość liter, separator dziesiętny, daty do ISO), więc
// kosmetyka formatowania przestaje cokolwiek psuć.

const TR_SIG_MAX_COLS = 16;   // ile kolumn wchodzi w odcisk (reszta to i tak szum)
const TR_SIG_MAX_ROWS = 4000; // powyżej tego nie zapisujemy odcisków kolumnowych (localStorage)
const TR_SIM_MIN = 0.7;       // próg tury „po podobieństwie"
const TR_SIM_COMMON = 8;      // wartość częstsza niż tyle razy nie identyfikuje wiersza

let trSigCols = [];            // [{ name, idx }] — kolumny wchodzące w odcisk, w kolejności arkusza
let trCellsCache = new Map();  // rowIndex0 -> odcisk kolumnowy (liczony leniwie)
let trDoneCells = new Map();   // klucz wiersza -> odcisk kolumnowy
let trDonePrev = new Map();    // klucz wiersza -> czytelny podgląd (lista nieodnalezionych)

// Nagłówek tak, jak widzi go użytkownik — do komunikatów. Do porównań służy postać
// znormalizowana (trNormHeaderName), ale pokazywanie jej w UI wyglądałoby jak literówka.
function trHeaderLabel(idx) {
  const src = Array.isArray(currentHeaders) && currentHeaders.length ? currentHeaders : trHeaders;
  const raw = String(src[idx] ?? "").trim();
  return raw || `#${idx + 1}`;
}

function trNormHeaderName(name) {
  return String(name ?? "").trim().toLowerCase().replace(/\s+/g, " ");
}

// Nazwy nagłówków jako podpisy kolumn. Duplikaty numerujemy, bo inaczej dwie kolumny
// „Uwagi" byłyby tą samą kolumną i dopasowanie mieszałoby je ze sobą.
function trSigHeaderNames() {
  const src = Array.isArray(currentHeaders) && currentHeaders.length ? currentHeaders : trHeaders;
  const seen = new Map();
  return src.map((h, i) => {
    const base = trNormHeaderName(h) || `#${i}`;
    const n = (seen.get(base) || 0) + 1;
    seen.set(base, n);
    return n > 1 ? `${base}#${n}` : base;
  });
}

// Wartość do odcisku: surowa (values), bo `display` zależy od ustawień wyświetlania.
function trRawValue(row, i) {
  if (row && Array.isArray(row.values) && i < row.values.length) return row.values[i];
  return getDisplayValue(row, i);
}

// Normalizacja decyduje, czy „12,50" i „12.5" to ta sama treść. Ta sama funkcja pracuje
// po obu stronach porównania, więc jest symetryczna z definicji.
function trNormValue(v) {
  if (v == null) return "";
  if (v instanceof Date) return Number.isNaN(v.getTime()) ? "" : v.toISOString().slice(0, 10);
  if (typeof v === "number") return Number.isFinite(v) ? String(Math.round(v * 1e6) / 1e6) : "";
  if (typeof v === "boolean") return v ? "1" : "0";
  const s = String(v).replace(/ /g, " ").trim().replace(/\s+/g, " ");
  if (!s) return "";
  // Liczba zapisana tekstem („1 234,50", „1,234.50") — do jednej postaci, żeby zmiana
  // formatu kolumny w Excelu nie wyglądała jak zmiana treści.
  if (/^-?[\d\s.,]+$/.test(s) && /\d/.test(s)) {
    const cleaned = s.replace(/\s/g, "").replace(/\.(?=\d{3}(\D|$))/g, "").replace(",", ".");
    const n = Number(cleaned);
    if (Number.isFinite(n)) return String(Math.round(n * 1e6) / 1e6);
  }
  return s.toLowerCase();
}

function trValueSig(row, idx) {
  const norm = trNormValue(trRawValue(row, idx));
  return norm ? trHash(norm) : "";
}

// Które kolumny biorą udział w odcisku. Liczy się informacja: kolumna wypełniona
// i zróżnicowana identyfikuje wiersz, kolumna z trzema powtarzającymi się wartościami
// prawie nic nie wnosi. Kolumny „na dziś" (TODAY) odpadają z definicji.
function trPickSigCols() {
  const names = trSigHeaderNames();
  const src = Array.isArray(baseRows) && baseRows.length ? baseRows : trRows;
  const step = Math.max(1, Math.floor(src.length / 400) || 1);
  const scored = [];
  for (let i = 0; i < names.length; i++) {
    if (trVolatileCols.has(i)) continue;
    let seen = 0;
    let filled = 0;
    const distinct = new Set();
    for (let r = 0; r < src.length; r += step) {
      seen += 1;
      const v = trNormValue(trRawValue(src[r], i));
      if (!v) continue;
      filled += 1;
      if (distinct.size <= 4000) distinct.add(v);
    }
    if (!seen || !filled) continue;
    const fill = filled / seen;
    const uniq = distinct.size / filled;
    scored.push({ name: names[i], idx: i, score: fill * (0.4 + 0.6 * uniq) });
  }
  scored.sort((a, b) => b.score - a.score);
  const keep = scored.slice(0, TR_SIG_MAX_COLS);
  keep.sort((a, b) => a.idx - b.idx);
  trSigCols = keep;
  trCellsCache.clear();
}

function trRowCells(row) {
  if (!row || !trSigCols.length) return "";
  const cacheKey = row.rowIndex0;
  if (cacheKey !== undefined && trCellsCache.has(cacheKey)) return trCellsCache.get(cacheKey);
  const out = trSigCols.map((c) => trValueSig(row, c.idx)).join("|");
  if (cacheKey !== undefined) trCellsCache.set(cacheKey, out);
  return out;
}

// Podgląd wiersza — jedyne, po czym człowiek pozna, CZEGO nie udało się dopasować.
// Bierzemy pola pokazywane na karcie, bo to one opisują wiersz w języku użytkownika.
function trRowPreview(row) {
  if (!row) return "";
  const order = Array.isArray(trFieldOrder) && trFieldOrder.length
    ? trFieldOrder
    : trHeaders.map((_, i) => i);
  const cols = trSelected && trSelected.size ? order.filter((i) => trSelected.has(i)) : order;
  const out = [];
  for (const i of cols) {
    const v = String(getDisplayValue(row, i) ?? "").trim();
    if (!v) continue;
    out.push(v.length > 28 ? `${v.slice(0, 27)}…` : v);
    if (out.length >= 3) break;
  }
  return out.join(" · ");
}

function trClearSigs() {
  trDoneSigs.clear();
  trDoneCells.clear();
  trDonePrev.clear();
}

// Jedno miejsce, w którym wiersz dostaje komplet odcisków. Wołane przy KAŻDYM ✓ —
// inaczej zapis miałby klucz bez odcisku i przy następnej zmianie pliku byłby nie do uratowania.
function trMarkSig(key, row) {
  if (!key || !row) return;
  trDoneSigs.set(key, trRowSig(row));
  const cells = trRowCells(row);
  if (cells) trDoneCells.set(key, cells);
  const prev = trRowPreview(row);
  if (prev) trDonePrev.set(key, prev);
}

// Odcisk ARKUSZA liczymy z `baseRows` (pełne, nieprzefiltrowane) — inaczej otwarcie trybu
// przy włączonym filtrze wyglądałoby jak „plik się zmienił”. Próbkujemy do 40 wierszy:
// tanio, a łapie realne edycje. Data zapisu pliku dokłada sygnał „to inna wersja pliku”.
function trSheetFingerprint() {
  const src = Array.isArray(baseRows) ? baseRows : [];
  const step = Math.max(1, Math.floor(src.length / 40) || 1);
  let sample = "";
  for (let i = 0; i < src.length; i += step) sample += trRowSig(src[i]) + ",";
  let mtime = "";
  try {
    const raw = workbook?.Props?.ModifiedDate;
    if (raw) mtime = String(raw instanceof Date ? raw.toISOString() : raw);
  } catch { /* brak właściwości dokumentu — zostaje reszta odcisku */ }
  return {
    rows: src.length,
    cols: trHeaders.length,
    h: trHash((Array.isArray(currentHeaders) ? currentHeaders : trHeaders).join(TR_SEP)),
    sample: trHash(sample),
    mtime,
  };
}

// Trzy poziomy, bo trzy różne wnioski dla użytkownika:
//   "same"  — nic się nie zmieniło,
//   "soft"  — plik był zapisywany ponownie, ale dane wyglądają tak samo (info, bez alarmu),
//   "hard"  — inna liczba wierszy / inne nagłówki / inna próbka treści (✓ mogą być nie na swoim).
function trCompareFingerprint(saved, now) {
  if (!saved || typeof saved !== "object") return "same"; // zapis sprzed tej wersji — nie strasz
  if (saved.rows !== now.rows || saved.h !== now.h || saved.sample !== now.sample) return "hard";
  if (saved.mtime && now.mtime && saved.mtime !== now.mtime) return "soft";
  return "same";
}

// Przeniesienie ✓ dla zapisów SPRZED odcisku kolumnowego: jedyne, co mamy, to hash całego
// wiersza — albo trafia w punkt, albo nie ma czego szukać. Duplikaty rozdajemy licznikowo.
function trRemapDoneLegacy(savedDone, savedSigs, savedPrev) {
  const keys = new Set();
  const wanted = new Map();
  const owner = new Map(); // odcisk -> indeksy w savedDone (do podglądu nieodnalezionych)
  savedDone.forEach((key, i) => {
    const sig = savedSigs[i];
    if (!sig) return;
    wanted.set(sig, (wanted.get(sig) || 0) + 1);
    if (!owner.has(sig)) owner.set(sig, []);
    owner.get(sig).push(i);
  });
  if (!wanted.size) {
    return {
      keys: new Set(savedDone),
      moved: 0, exact: 0, byKey: 0, similar: 0,
      lost: 0, unmatched: [], remapped: false, keyCol: "",
    };
  }
  const source = Array.isArray(baseRows) && baseRows.length ? baseRows : trRows;
  let moved = 0;
  for (const row of source) {
    const sig = trRowSig(row);
    const left = wanted.get(sig);
    if (!left) continue;
    const key = trKeyOf(row);
    keys.add(key);
    trMarkSig(key, row);
    wanted.set(sig, left - 1);
    const queue = owner.get(sig);
    if (queue && queue.length) queue.shift();
    moved += 1;
  }
  const unmatched = [];
  wanted.forEach((left, sig) => {
    const queue = owner.get(sig) || [];
    for (let n = 0; n < left; n++) {
      const i = queue[n];
      unmatched.push({ prev: (savedPrev && savedPrev[i]) || "", best: null, score: 0 });
    }
  });
  return {
    keys,
    moved, exact: moved, byKey: 0, similar: 0,
    lost: unmatched.length, unmatched, remapped: true, keyCol: "",
  };
}

// ── Kaskada dopasowania ✓ ───────────────────────────────────────────────────
//
// Trzy tury, od najpewniejszej do najluźniejszej. Każdy wiersz docelowy może zostać zajęty
// tylko RAZ, a tury idą po kolei — więc pewne dopasowanie zawsze wygrywa z domysłem,
// niezależnie od kolejności wierszy w pliku.
//
//   1. DOKŁADNIE  — zgadzają się wszystkie wspólne kolumny.
//   2. PO KLUCZU  — zgadza się kolumna, która w tym arkuszu jednoznacznie identyfikuje wiersz
//                   (prawie same unikaty, prawie zawsze wypełniona — np. „Nr"). Wymagamy
//                   jednoznaczności po OBU stronach: jeden zapis, jeden wiersz.
//   3. PO PODOBIEŃSTWIE — co najmniej 70% wspólnych kolumn i JEDEN wyraźny zwycięzca (bez
//                   remisu). Jeśli arkusz ma kolumnę-klucz, jej wartość MUSI się zgadzać —
//                   inaczej „podobny" wiersz to po prostu cudzy wiersz.
//
// Czego kaskada NIE robi: nie zgaduje przy remisie i nie odhacza niczego „na oko" — to,
// czego nie dopasuje, ląduje na liście nieodnalezionych razem z podglądem treści.
function trMatchDone(rec) {
  const savedDone = Array.isArray(rec?.done) ? rec.done : [];
  const savedSigs = Array.isArray(rec?.doneSig) ? rec.doneSig : [];
  const savedPrev = Array.isArray(rec?.donePrev) ? rec.donePrev : [];
  const savedCols = Array.isArray(rec?.cols) ? rec.cols : [];
  const savedCells = Array.isArray(rec?.doneCells) ? rec.doneCells : [];
  if (!savedDone.length) {
    return { keys: new Set(), moved: 0, exact: 0, byKey: 0, similar: 0, lost: 0, unmatched: [], remapped: false, keyCol: "" };
  }
  if (!savedCols.length || !savedCells.some(Boolean)) {
    return trRemapDoneLegacy(savedDone, savedSigs, savedPrev);
  }

  // Część wspólna kolumn: po NAZWIE, z pominięciem kolumn liczonych „na dziś" po TEJ stronie.
  // Tu leżał błąd przy „Przenieś": odciski źródła liczone były bez kolumny TODAY, a odciski
  // celu razem z nią — porównywaliśmy dwie różne rzeczy i ✓ nie miały prawa trafić.
  const names = trSigHeaderNames();
  const byName = new Map();
  names.forEach((n, i) => { if (!byName.has(n)) byName.set(n, i); });
  const pairs = [];
  savedCols.forEach((name, pos) => {
    const idx = byName.get(name);
    if (idx === undefined || trVolatileCols.has(idx)) return;
    pairs.push({ pos, idx, name });
  });
  if (!pairs.length) return trRemapDoneLegacy(savedDone, savedSigs, savedPrev);

  const rows = Array.isArray(baseRows) && baseRows.length ? baseRows : trRows;
  // Indeksy odwrotne: hash wartości -> wiersze. Dają naraz turę 1 (pełny klucz), wykrycie
  // kolumny-klucza (ile unikatów) i głosowanie tury 3 — bez skanowania każdy z każdym.
  const colIndex = pairs.map(() => new Map());
  const fullIndex = new Map();
  rows.forEach((row, r) => {
    const cells = pairs.map((p, pi) => {
      const h = trValueSig(row, p.idx);
      if (h) {
        const list = colIndex[pi].get(h);
        if (list) list.push(r); else colIndex[pi].set(h, [r]);
      }
      return h;
    });
    const full = cells.join("|");
    const bucket = fullIndex.get(full);
    if (bucket) bucket.push(r); else fullIndex.set(full, [r]);
  });

  const taken = new Set();
  const keys = new Set();
  let exact = 0;
  let byKey = 0;
  let similar = 0;
  const claim = (r) => {
    taken.add(r);
    const key = trKeyOf(rows[r]);
    keys.add(key);
    trMarkSig(key, rows[r]);
  };
  const freeIn = (list) => {
    if (!list) return -1;
    for (const r of list) if (!taken.has(r)) return r;
    return -1;
  };

  const pending = [];
  const legacyPending = [];
  savedDone.forEach((key, i) => {
    const raw = savedCells[i];
    if (!raw) {
      legacyPending.push({ i, sig: savedSigs[i] || "", prev: savedPrev[i] || "" });
      return;
    }
    const all = String(raw).split("|");
    pending.push({ i, cells: pairs.map((p) => all[p.pos] || ""), prev: savedPrev[i] || "" });
  });

  // ── Tura 1: dokładnie ──
  const minShared = Math.min(2, pairs.length);
  let rest = [];
  pending.forEach((p) => {
    if (p.cells.filter(Boolean).length < minShared) { rest.push(p); return; }
    const hit = freeIn(fullIndex.get(p.cells.join("|")));
    if (hit < 0) { rest.push(p); return; }
    claim(hit);
    exact += 1;
  });

  // Wiersze zapisane starszą wersją (mają tylko hash całego wiersza) — dokładamy je do
  // tury 1, żeby mieszany zapis nie tracił ✓ tylko dlatego, że część jest starsza.
  legacyPending.forEach((p) => {
    let hit = -1;
    if (p.sig) {
      for (let r = 0; r < rows.length; r++) {
        if (taken.has(r)) continue;
        if (trRowSig(rows[r]) === p.sig) { hit = r; break; }
      }
    }
    if (hit < 0) { rest.push({ i: p.i, cells: [], prev: p.prev }); return; }
    claim(hit);
    exact += 1;
  });

  // ── Kolumna-klucz: ta, która w TYM arkuszu identyfikuje wiersz jednoznacznie ──
  let keyPi = -1;
  let keyScore = 0;
  pairs.forEach((_, pi) => {
    const m = colIndex[pi];
    let filled = 0;
    m.forEach((list) => { filled += list.length; });
    if (!filled || !rows.length) return;
    const uniq = m.size / filled;
    const fill = filled / rows.length;
    if (uniq < 0.9 || fill < 0.8) return;
    const score = uniq * fill;
    if (score > keyScore) { keyScore = score; keyPi = pi; }
  });

  // ── Tura 2: po kluczu ──
  if (keyPi >= 0 && rest.length) {
    const demand = new Map(); // ilu zapisanych wierszy chce tej samej wartości klucza
    rest.forEach((p) => {
      const h = p.cells[keyPi];
      if (h) demand.set(h, (demand.get(h) || 0) + 1);
    });
    const next = [];
    rest.forEach((p) => {
      const h = p.cells[keyPi];
      // Jednoznacznie po obu stronach — inaczej nie wiadomo, który wiersz jest czyj.
      if (!h || demand.get(h) !== 1) { next.push(p); return; }
      const list = colIndex[keyPi].get(h);
      if (!list || list.length !== 1 || taken.has(list[0])) { next.push(p); return; }
      claim(list[0]);
      byKey += 1;
    });
    rest = next;
  }

  // ── Tura 3: po podobieństwie ──
  const unmatched = [];
  rest.forEach((p) => {
    const denom = p.cells.filter(Boolean).length;
    if (!denom) { unmatched.push({ prev: p.prev, best: null, score: 0 }); return; }
    const votes = new Map();
    pairs.forEach((_, pi) => {
      const h = p.cells[pi];
      if (!h) return;
      const list = colIndex[pi].get(h);
      // Wartość pospolita („Teren" w 900 wierszach) niczego nie identyfikuje — pomijamy,
      // inaczej głosowanie wygrywałby przypadkowy wiersz o tym samym statusie.
      if (!list || list.length > TR_SIM_COMMON) return;
      list.forEach((r) => { if (!taken.has(r)) votes.set(r, (votes.get(r) || 0) + 1); });
    });
    let best = -1;
    let bestV = 0;
    let secondV = 0;
    votes.forEach((v, r) => {
      if (v > bestV) { secondV = bestV; bestV = v; best = r; }
      else if (v > secondV) secondV = v;
    });
    const score = denom ? bestV / denom : 0;
    const keyOk = keyPi < 0 || !p.cells[keyPi]
      || (best >= 0 && trValueSig(rows[best], pairs[keyPi].idx) === p.cells[keyPi]);
    if (best >= 0 && score >= TR_SIM_MIN && bestV > secondV && keyOk) {
      claim(best);
      similar += 1;
      return;
    }
    unmatched.push({
      prev: p.prev,
      best: best >= 0 ? trKeyOf(rows[best]) : null,
      score: best >= 0 ? score : 0,
    });
  });

  return {
    keys,
    moved: exact + byKey + similar,
    exact,
    byKey,
    similar,
    lost: unmatched.length,
    unmatched,
    remapped: true,
    keyCol: keyPi >= 0 ? trHeaderLabel(pairs[keyPi].idx) : "",
  };
}

// ── Model / wiersze ─────────────────────────────────────────────────────────

function trKeyOf(row) {
  return typeof getRowSelectionKey === "function" ? getRowSelectionKey(row) : String(row?.rowIndex0 ?? "");
}

function trCurrentRow() {
  const idx = trOrder[trPos];
  return Number.isInteger(idx) ? trRows[idx] : null;
}

function trCurrentKey() {
  const row = trCurrentRow();
  return row ? trKeyOf(row) : "";
}

function trFieldLabel(idx) {
  return typeof exportColLabel === "function"
    ? exportColLabel(trHeaders[idx], idx)
    : String(trHeaders[idx] ?? idx + 1);
}

// Dwa tryby doboru pól:
//   ręczny  — dokładnie to, co zaznaczone (stałe rubryki formularza),
//   auto    — pola z wartością W TYM wierszu, brane ze WSZYSTKICH kolumn.
// Auto jest odpowiedzią na arkusze z powtarzanymi blokami (Kw1_*, Kw2_*, Kw3_*)
// i na „przesunięte” wiersze: raz dane siedzą jedną kolumnę w prawo, raz dwie.
// Zaznaczenia w trybie auto nie mają wpływu, ale KOLEJNOŚĆ owszem.
// ── Dziedziczenie z góry (scalone komórki / „wartość tylko w pierwszym wierszu") ──
//
// Bardzo częsty układ: nazwisko scalone przez 5 wierszy, pod spodem pozycje. W danych
// wiersze-kontynuacje są PUSTE (build-rows-core czyta `!merges` tylko po to, żeby
// wyznaczyć zakres arkusza — wartości nie rozlewa), więc bez tego mechanizmu tryb
// „dobieraj pola z wiersza" ukrywałby pole tożsamości akurat tam, gdzie jest najbardziej
// potrzebne.
//
// DWIE DECYZJE PROJEKTOWE, które trzymają to w ryzach:
//  1. Liczymy po `baseRows` — pełnym, NIEPRZEFILTROWANYM zestawie w kolejności arkusza.
//     Liczenie po widoku dawałoby inny wynik po filtrze albo sortowaniu, a „wiersz wyżej"
//     to własność PLIKU, nie bieżącego widoku.
//  2. Kolumny wskazuje użytkownik (scalenia tylko je podpowiadają). Zgadywanie „ta kolumna
//     chyba się przenosi" mogłoby wpisać cudzą wartość do rubryki na papierze, a tego się
//     nie cofa gumką.
// Pionowe scalenia arkusza, pogrupowane po kolumnie MODELU. Trzymamy pełne zakresy,
// nie same numery kolumn, bo zakres jest jednocześnie GRANICĄ przenoszenia.
function trDetectMergeRanges() {
  const byCol = new Map();
  if (trLongMode) return byCol;
  try {
    const sheet = workbook?.Sheets?.[currentSheetName];
    const merges = Array.isArray(sheet?.["!merges"]) ? sheet["!merges"] : [];
    const startCol = Number.isFinite(currentStartCol) ? currentStartCol : 0;
    merges.forEach((m) => {
      if (!m?.s || !m?.e || m.e.r <= m.s.r) return; // interesują nas tylko scalenia PIONOWE
      for (let c = m.s.c; c <= m.e.c; c++) {
        const col = c - startCol;
        if (col < 0 || col >= trHeaders.length) continue;
        if (!byCol.has(col)) byCol.set(col, []);
        byCol.get(col).push({ start: m.s.r, end: m.e.r });
      }
    });
  } catch {
    /* brak dostępu do arkusza — zostaje ręczny wybór kolumn */
  }
  return byCol;
}

// DWIE REGUŁY, świadomie różne — bo różna jest pewność, skąd bierze się pustka:
//
//  • kolumna ZE SCALENIAMI → przenosimy DOKŁADNIE w granicach scalenia. Plik sam mówi,
//    dokąd sięga rekord, więc nie ma miejsca na wyciek wartości do następnego rekordu.
//  • kolumna wskazana RĘCZNIE (bez scaleń) → najbliższa wartość powyżej. Tu granicy
//    rekordu nikt nie zapisał, więc to świadomy wybór użytkownika; wartość zawsze
//    dostaje na karcie numer wiersza źródłowego, żeby dało się ją sprawdzić.
function trBuildInheritance() {
  trInheritMap.clear();
  if (!trInheritOn || trLongMode || !trInheritCols.size) return;
  const source = Array.isArray(baseRows) ? baseRows : [];
  if (!source.length) return;

  const byRowIndex = new Map();
  source.forEach((row) => { byRowIndex.set(row.rowIndex0, row); });

  const mergeCols = [];
  const carryCols = [];
  trInheritCols.forEach((col) => {
    if (trMergeRanges.has(col)) mergeCols.push(col);
    else carryCols.push(col);
  });

  // 1) Kolumny ze scaleniami — zakres po zakresie, kotwicą jest pierwszy wiersz scalenia.
  mergeCols.forEach((col) => {
    (trMergeRanges.get(col) || []).forEach((range) => {
      const anchorRow = byRowIndex.get(range.start);
      if (!anchorRow) return; // kotwica nad wierszem nagłówka albo poza zakresem danych
      const text = String(getDisplayValue(anchorRow, col) ?? "").trim();
      if (!text) return;
      const from = (range.start ?? 0) + 1;
      for (let r = range.start + 1; r <= range.end; r++) {
        const row = byRowIndex.get(r);
        if (!row) continue;
        if (String(getDisplayValue(row, col) ?? "").trim()) continue; // własna wartość wygrywa
        trInheritMap.set(`${col}:${r}`, { text, from });
      }
    });
  });

  // 2) Kolumny wskazane ręcznie — jeden przebieg po wierszach dla wszystkich naraz.
  if (!carryCols.length) return;
  const carry = new Array(carryCols.length).fill(null);
  for (const row of source) {
    // Podnagłówek = granica sekcji, więc zrywa przenoszenie. UWAGA: markSubheaderRows
    // sprawdza tylko pierwsze wiersze arkusza, więc to zabezpieczenie łapie nagłówki
    // sekcji u góry pliku, a nie jest pełnym wykrywaniem rekordów.
    if (row.isSubheader) {
      carry.fill(null);
      continue;
    }
    const rowIdx = row.rowIndex0;
    for (let k = 0; k < carryCols.length; k++) {
      const col = carryCols[k];
      const txt = String(getDisplayValue(row, col) ?? "").trim();
      if (txt) carry[k] = { text: txt, from: (rowIdx ?? 0) + 1 };
      else if (carry[k]) trInheritMap.set(`${col}:${rowIdx}`, carry[k]);
    }
  }
}

// Jedno miejsce, które odpowiada „co ma stanąć w tej rubryce”: wartość własna wiersza,
// a jak jej nie ma — odziedziczona (z numerem wiersza źródłowego, żeby dało się sprawdzić).
function trResolveField(row, idx) {
  const raw = String(getDisplayValue(row, idx) ?? "").trim();
  if (raw) return { text: raw, from: 0 };
  if (!trInheritOn || trLongMode || !trInheritCols.has(idx)) return { text: "", from: 0 };
  const hit = trInheritMap.get(`${idx}:${row?.rowIndex0}`);
  return hit ? { text: hit.text, from: hit.from } : { text: "", from: 0 };
}

function trHasValue(row, idx) {
  return trResolveField(row, idx).text !== "";
}

function trVisibleCols(row) {
  if (trAutoFields) {
    if (!row) return [];
    return trFieldOrder.filter((idx) => trHasValue(row, idx));
  }
  return trFieldOrder.filter((idx) => trSelected.has(idx));
}

// Ile pól tryb auto przemilczał — bez tego licznika znikające rubryki wyglądają
// jak zgubione dane, a nie jak świadome pominięcie pustych.
function trSkippedCount(row) {
  if (!trAutoFields || !row) return 0;
  return trFieldOrder.length - trVisibleCols(row).length;
}

// Domyślny zestaw pól przy pierwszym otwarciu arkusza: kolumny, które faktycznie
// coś zawierają (próbka wierszy), przycięte do TR_DEFAULT_FIELDS. Lepszy start niż
// 40 pustych rubryk — resztę użytkownik dokłada w panelu „Pola”.
function trDefaultSelection() {
  const sample = trRows.slice(0, 300);
  const filled = [];
  trHeaders.forEach((_, idx) => {
    const has = sample.some((row) => String(getDisplayValue(row, idx) ?? "").trim() !== "");
    if (has) filled.push(idx);
  });
  const base = filled.length ? filled : trHeaders.map((_, i) => i);
  return new Set(base.slice(0, TR_DEFAULT_FIELDS));
}

function trRebuildOrder(preserveKey) {
  trOrder = [];
  trRows.forEach((row, i) => {
    if (!trHideDone || !trDone.has(trKeyOf(row))) trOrder.push(i);
  });
  if (preserveKey) {
    const at = trOrder.findIndex((i) => trKeyOf(trRows[i]) === preserveKey);
    if (at >= 0) trPos = at;
  }
  trPos = Math.max(0, Math.min(trPos, Math.max(0, trOrder.length - 1)));
}

// ── Belka przewijania / „jest tego więcej" ──────────────────────────────────
//
// Problem z tabletu: karta z kilkunastoma polami nie mieści się na ekranie, a jedyną
// informacją o tym jest natywny pasek, który na iPadOS pojawia się DOPIERO w trakcie
// przewijania. Przy przepisywaniu na papier to realna pomyłka — pole 12 zostaje
// nieprzepisane, bo nikt nie wiedział, że istnieje.
//
// Stąd trzy sygnały naraz: własna belka z boku (jest tu w ogóle co przewijać?),
// cienie-krawędzie (treść jest ucięta w tę stronę) i pigułka z LICZBĄ pól poniżej
// (ile dokładnie zostało) — pigułka jest jednocześnie przyciskiem „przewiń o ekran".

function trFieldsBelowFold() {
  if (!trStageEl) return 0;
  const foldY = trStageEl.getBoundingClientRect().bottom;
  let n = 0;
  trCardEl?.querySelectorAll(".tr-field").forEach((el) => {
    // liczymy pole jako „poniżej”, gdy jego etykieta i wartość nie są w całości widoczne
    if (el.getBoundingClientRect().bottom > foldY + 2) n += 1;
  });
  return n;
}

function trUpdateScrollUi() {
  if (!trStageEl || !trStageWrapEl) return;
  const max = trStageEl.scrollHeight - trStageEl.clientHeight;
  const overflow = max > 4;
  const top = trStageEl.scrollTop;
  trStageWrapEl.classList.toggle("has-overflow", overflow);
  trStageWrapEl.classList.toggle("at-top", top <= 2);
  trStageWrapEl.classList.toggle("at-bottom", !overflow || top >= max - 2);

  if (trScrollThumbEl && trScrollRailEl) {
    const railH = trScrollRailEl.clientHeight;
    const ratio = trStageEl.scrollHeight ? trStageEl.clientHeight / trStageEl.scrollHeight : 1;
    const thumbH = Math.max(22, Math.round(railH * Math.min(1, ratio)));
    const travel = Math.max(0, railH - thumbH);
    const progress = max > 0 ? Math.min(1, Math.max(0, top / max)) : 0;
    trScrollThumbEl.style.height = `${thumbH}px`;
    trScrollThumbEl.style.transform = `translateY(${Math.round(travel * progress)}px)`;
  }

  if (trScrollMoreTextEl) {
    const below = overflow ? trFieldsBelowFold() : 0;
    // Gdy pole jest jedno, ale bardzo wysokie, licznik pokazałby „0” — wtedy sam napis.
    trScrollMoreTextEl.textContent = below > 0 ? t("trScrollMore", { n: below }) : t("trScrollMoreMore");
  }
}

function trScheduleScrollUi() {
  if (trScrollRaf) return;
  trScrollRaf = requestAnimationFrame(() => {
    trScrollRaf = 0;
    trUpdateScrollUi();
  });
}

// Nowy wiersz = nowa kartka: zawsze zaczynamy od GÓRY. Bez tego po przewinięciu
// długiego wiersza następny otwierał się w połowie i pierwsze pola uciekały nad ekran.
function trResetScroll() {
  if (!trStageEl) return;
  trStageEl.scrollTop = 0;
  trScheduleScrollUi();
}

if (trStageEl) trStageEl.addEventListener("scroll", trScheduleScrollUi, { passive: true });
if (trScrollMoreEl) {
  trScrollMoreEl.addEventListener("click", () => {
    if (!trStageEl) return;
    const step = Math.max(120, Math.round(trStageEl.clientHeight * 0.82));
    trStageEl.scrollBy({ top: step, behavior: "smooth" });
  });
}
window.addEventListener("resize", () => { if (trIsOpen) trScheduleScrollUi(); });

// ── Render ──────────────────────────────────────────────────────────────────

function trRenderCard() {
  if (!trCardEl) return;
  const row = trCurrentRow();
  const cols = trVisibleCols(row);
  const total = trOrder.length;

  trCardEl.replaceChildren();
  const noRows = !row;
  const noFields = !cols.length;
  trCardEl.classList.toggle("hidden", noRows || noFields);
  if (trEmptyEl) {
    trEmptyEl.classList.toggle("hidden", !(noRows || noFields));
    if (noRows || noFields) {
      trEmptyEl.replaceChildren();
      const title = document.createElement("div");
      title.className = "tr-empty-title";
      const sub = document.createElement("div");
      sub.className = "tr-empty-sub";
      if (noFields && trAutoFields) {
        // W trybie auto „brak pól” znaczy: cały wiersz jest pusty. To inny komunikat
        // niż „nic nie zaznaczyłeś” — inaczej wygląda na awarię.
        title.textContent = t("trRowEmpty");
        sub.textContent = t("trRowEmptySub");
      } else if (noFields) {
        title.textContent = t("trNoFields");
        sub.textContent = t("trNoFieldsSub");
      } else if (trDone.size >= trRows.length && trRows.length) {
        title.textContent = t("trAllDone");
        sub.textContent = t("trAllDoneSub");
      } else {
        title.textContent = t("trNothingToShow");
        sub.textContent = t("trAllDoneSub");
      }
      trEmptyEl.append(title, sub);
    }
  }

  if (!noRows && !noFields) {
    const head = document.createElement("div");
    head.className = "tr-card-head";
    const srcRow = document.createElement("span");
    srcRow.className = "tr-card-rownum";
    const rowLabel = trRowHeadFormatter ? trRowHeadFormatter(row) : String((row.rowIndex0 ?? 0) + 1);
    srcRow.textContent = t("trSheetRow", { n: rowLabel });
    const skipped = trSkippedCount(row);
    if (skipped > 0) {
      const skippedEl = document.createElement("span");
      skippedEl.className = "tr-card-skipped";
      skippedEl.textContent = t("trSkipped", { n: skipped });
      head.appendChild(skippedEl);
    }
    head.appendChild(srcRow);
    trCardEl.appendChild(head);

    cols.forEach((ci) => {
      const field = document.createElement("div");
      field.className = "tr-field";
      const label = document.createElement("div");
      label.className = "tr-field-label";
      const labelText = document.createElement("span");
      labelText.textContent = trFieldLabel(ci);
      label.appendChild(labelText);
      const resolved = trResolveField(row, ci);
      if (resolved.from) {
        // Wartość odziedziczona MUSI być rozpoznawalna — inaczej przepisze się ją
        // jak własną i nie da się już wychwycić pomyłki.
        field.classList.add("is-inherited");
        const badge = document.createElement("span");
        badge.className = "tr-inherited-from";
        badge.textContent = t("trInheritedFrom", { n: resolved.from });
        label.appendChild(badge);
      }
      const value = document.createElement("div");
      value.className = "tr-field-value";
      if (resolved.text) {
        value.textContent = resolved.text;
      } else {
        value.textContent = "—";
        value.classList.add("is-empty");
      }
      field.append(label, value);
      trCardEl.appendChild(field);
    });
  }

  // Licznik, pasek postępu, stan ✓ bieżącego wiersza
  const pos = total ? trPos + 1 : 0;
  if (trCounterEl) trCounterEl.textContent = t("trCounter", { pos, total });
  if (trDoneCountEl) trDoneCountEl.textContent = t("trDoneCount", { done: trDone.size, all: trRows.length });
  if (trProgressBarEl) {
    const pct = trRows.length ? Math.round((trDone.size / trRows.length) * 100) : 0;
    trProgressBarEl.style.width = `${pct}%`;
  }
  const isDone = !!row && trDone.has(trKeyOf(row));
  if (trMarkChipEl) {
    trMarkChipEl.classList.toggle("is-done", isDone);
    trMarkChipEl.textContent = isDone ? t("trChipDone") : t("trChipPending");
    trMarkChipEl.setAttribute("aria-pressed", isDone ? "true" : "false");
    trMarkChipEl.disabled = !row;
  }
  if (trPrevBtn) trPrevBtn.disabled = !total || trPos <= 0;
  if (trNextBtn) trNextBtn.disabled = !total || trPos >= total - 1;
  if (trMarkBtn) trMarkBtn.disabled = !row;

  if (trLiveEl && row) trLiveEl.textContent = t("trLiveRow", { pos, total });
  // Pomiar po wstawieniu pól do DOM — inaczej scrollHeight jest jeszcze sprzed renderu.
  trScheduleScrollUi();
  if (trStatsEl && trProgressPanelEl && !trProgressPanelEl.classList.contains("hidden")) trRenderProgressPanel();
  trPersist();
}

// ── Nawigacja ───────────────────────────────────────────────────────────────

function trGo(delta) {
  if (!trOrder.length) return;
  const next = trPos + delta;
  if (next < 0 || next >= trOrder.length) return;
  trPos = next;
  trResetScroll();
  trRenderCard();
}

function trGoEdge(which) {
  if (!trOrder.length) return;
  trPos = which < 0 ? 0 : trOrder.length - 1;
  trResetScroll();
  trRenderCard();
}

function trToggleDone() {
  const row = trCurrentRow();
  if (!row) return;
  const key = trKeyOf(row);
  if (trDone.has(key)) trDone.delete(key);
  else { trDone.add(key); trMarkSig(key, row); }
  if (trHideDone) {
    const keep = trPos;
    trRebuildOrder(null);
    trPos = Math.max(0, Math.min(keep, Math.max(0, trOrder.length - 1)));
  }
  trRenderCard();
}

// Główna akcja: odhacz i przejdź dalej. Przy „ukryj spisane” bieżący wiersz znika,
// więc pozycja ZOSTAJE na miejscu i sama pokazuje następny — bez przeskoku o dwa.
// Zwraca klucz odhaczonego wiersza (albo "" gdy nie było czego odhaczyć) — potrzebne
// szybkiemu odhaczaniu, żeby dało się całą serię cofnąć jednym ruchem.
function trMarkAndNext(options = {}) {
  const row = trCurrentRow();
  if (!row) return "";
  const key = trKeyOf(row);
  const wasDone = trDone.has(key);
  trDone.add(key);
  trMarkSig(key, row);
  const before = trPos;
  if (trHideDone) {
    const keep = trPos;
    trRebuildOrder(null);
    trPos = Math.max(0, Math.min(keep, Math.max(0, trOrder.length - 1)));
  } else if (trPos < trOrder.length - 1) {
    trPos += 1;
  }
  const moved = trHideDone ? true : trPos !== before;
  trResetScroll();
  trRenderCard();
  if (!options.quiet && trDone.size >= trRows.length && trRows.length) toast(t("trAllDone"), "success");
  return { key, wasDone, moved };
}

// ── Szybkie odhaczanie (przytrzymanie) ──────────────────────────────────────
//
// Po co: apka w tle na tablecie potrafi zostać ubita, a po powrocie trzeba dojść do
// miejsca sprzed przerwy. Klikanie „Spisane i dalej” 200 razy to nie jest plan.
// PRZYTRZYMANIE (palcem na przycisku albo spacją, gdy przycisk ma fokus) rozpędza
// odhaczanie: po ~0,55 s startuje pętla, która przyspiesza z 200 ms do 60 ms na wiersz.
//
// Trzy bezpieczniki, bo to operacja masowa:
//  1. próg czasu — zwykły tap NIGDY nie wejdzie w tryb szybki,
//  2. pasek na przycisku pokazuje, ile zostało do startu (nic nie dzieje się „nagle”),
//  3. cała seria cofa się jednym przyciskiem w pasku meta (i wraca na wiersz startowy).
const TR_HOLD_MS = 550;        // ile trzymać, zanim ruszy tryb szybki
const TR_TURBO_START_MS = 200; // pierwszy krok
const TR_TURBO_MIN_MS = 45;    // najszybszy krok
const TR_TURBO_ACCEL = 0.86;   // mnożnik między krokami

let trBurstStartKey = "";
let trSuppressNextClick = false;

// Licznik „+N" na przycisku żyje w custom property, bo tekst przycisku podmienia i18n.
function trSetTurboLabel(text) {
  if (!trMarkBtn) return;
  if (text) trMarkBtn.style.setProperty("--tr-turbo-label", JSON.stringify(text));
  else trMarkBtn.style.removeProperty("--tr-turbo-label");
}

function trSetHoldProgress(pct) {
  if (trMarkBtn) trMarkBtn.style.setProperty("--tr-hold", `${Math.round(pct)}%`);
}

function trHoldStart(source) {
  if (!trIsOpen || trTurboTimer || trHoldTimer) return;
  if (!trCurrentRow()) return;
  trTurboSource = source;
  const began = Date.now();
  trHoldProgressTimer = setInterval(() => {
    trSetHoldProgress(Math.min(100, ((Date.now() - began) / TR_HOLD_MS) * 100));
  }, 60);
  trHoldTimer = setTimeout(() => {
    trHoldTimer = 0;
    trTurboStart();
  }, TR_HOLD_MS);
}

// Zatrzymanie odliczania. `fired` = czy zdążył wystartować tryb szybki — od tego zależy,
// czy puszczenie klawisza/palca ma jeszcze odhaczyć pojedynczy wiersz.
function trHoldCancel() {
  const fired = !!trTurboTimer;
  if (trHoldTimer) { clearTimeout(trHoldTimer); trHoldTimer = 0; }
  if (trHoldProgressTimer) { clearInterval(trHoldProgressTimer); trHoldProgressTimer = 0; }
  trSetHoldProgress(0);
  if (fired) trTurboStop();
  trTurboSource = "";
  return fired;
}

function trTurboStart() {
  if (trHoldProgressTimer) { clearInterval(trHoldProgressTimer); trHoldProgressTimer = 0; }
  trSetHoldProgress(100);
  trTurboCount = 0;
  trBurstKeys = [];
  trBurstStartKey = trCurrentKey();
  // Zapis stanu to serializacja całego zbioru ✓ (do 20 tys. kluczy). Przy 20 wierszach
  // na sekundę robiłoby to z tabletu podkładkę pod kawę — zapisujemy raz, po serii.
  trBulkMode = true;
  if (trMarkBtn) trMarkBtn.classList.add("is-turbo");
  if (typeof navigator !== "undefined" && navigator.vibrate) { try { navigator.vibrate(12); } catch { /* brak wsparcia */ } }
  toast(t("trTurboStarted"), "info");

  let delay = TR_TURBO_START_MS;
  const step = () => {
    const res = trMarkAndNext({ quiet: true });
    if (!res || !res.key) { trTurboStop(); return; }
    if (!res.wasDone) trBurstKeys.push(res.key);
    trTurboCount += 1;
    trSetTurboLabel(`+${trTurboCount}`);
    if (!res.moved) { // koniec listy — dalej nie ma dokąd
      toast(t("trTurboEnd"), "info");
      trTurboStop();
      return;
    }
    delay = Math.max(TR_TURBO_MIN_MS, Math.round(delay * TR_TURBO_ACCEL));
    trTurboTimer = setTimeout(step, delay);
  };
  trTurboTimer = setTimeout(step, 0);
}

function trTurboStop() {
  if (trTurboTimer) { clearTimeout(trTurboTimer); trTurboTimer = 0; }
  const wasBulk = trBulkMode;
  trBulkMode = false;
  if (wasBulk) trPersist();
  if (trMarkBtn) trMarkBtn.classList.remove("is-turbo");
  trSetTurboLabel("");
  trSetHoldProgress(0);
  if (trTurboCount > 0) {
    toast(t("trTurboDone", { n: trTurboCount }), "success");
    trShowUndo(trBurstKeys.length);
  }
  trTurboCount = 0;
}

// Cofnięcie serii: zdejmujemy ✓ tylko z wierszy odhaczonych W TEJ serii (te, które
// były odhaczone wcześniej, zostają) i wracamy kursorem na wiersz startowy.
function trShowUndo(n) {
  if (!trUndoBtn) return;
  if (!n) { trHideUndo(); return; }
  trUndoBtn.textContent = t("trUndoBurst", { n });
  trUndoBtn.classList.remove("hidden");
  if (trUndoTimer) clearTimeout(trUndoTimer);
  trUndoTimer = setTimeout(trHideUndo, 12000);
}

function trHideUndo() {
  if (trUndoTimer) { clearTimeout(trUndoTimer); trUndoTimer = 0; }
  if (trUndoBtn) trUndoBtn.classList.add("hidden");
  trBurstKeys = [];
  trBurstStartKey = "";
}

function trUndoBurst() {
  if (!trBurstKeys.length) { trHideUndo(); return; }
  const n = trBurstKeys.length;
  const back = trBurstStartKey;
  trBurstKeys.forEach((key) => trDone.delete(key));
  trHideUndo();
  trRebuildOrder(back || null);
  trResetScroll();
  trRenderCard();
  toast(t("trTurboUndone", { n }), "success");
}

function trSetHideDone(on) {
  trHideDone = !!on;
  if (trHideDoneEl) trHideDoneEl.checked = trHideDone;
  trRebuildOrder(trHideDone ? null : trCurrentKey());
  trRenderCard();
}

function trResetProgress() {
  trDone.clear();
  trClearSigs();
  trUnmatched = [];
  trChangeInfo = null;   // czyścimy też powód ostrzeżenia — nie ma już czego ratować
  trHideChangeNotice();
  trRebuildOrder(null);
  trPos = 0;
  trRenderCard();
  if (trProgressPanelEl && !trProgressPanelEl.classList.contains("hidden")) trRenderProgressPanel();
  toast(t("trResetDone"), "success");
}

// Kasowanie ✓ to operacja nieodwracalna — dwustopniowy przycisk zamiast confirm(),
// żeby nie wyrywać użytkownika z pełnoekranowego trybu systemowym oknem.
// Uogólnione na dowolny przycisk, bo kasowań jest teraz kilka (✓ tego arkusza,
// pojedynczy zapis, wszystkie zapisy) i każde ma być tak samo trudne do przypadkowego kliknięcia.
const trArmedButtons = new Map(); // przycisk -> { timer, label }

function trArmDanger(btn, label, action) {
  if (!btn) return;
  const armed = trArmedButtons.get(btn);
  if (armed) {
    clearTimeout(armed.timer);
    trArmedButtons.delete(btn);
    btn.classList.remove("is-armed");
    btn.textContent = armed.label;
    action();
    return;
  }
  const timer = setTimeout(() => {
    const cur = trArmedButtons.get(btn);
    if (!cur) return;
    trArmedButtons.delete(btn);
    btn.classList.remove("is-armed");
    btn.textContent = cur.label;
  }, 3500);
  trArmedButtons.set(btn, { timer, label });
  btn.classList.add("is-armed");
  btn.textContent = t("trResetConfirm");
}

function trDisarmAll() {
  trArmedButtons.forEach((state, btn) => {
    clearTimeout(state.timer);
    btn.classList.remove("is-armed");
    btn.textContent = state.label;
  });
  trArmedButtons.clear();
}

function trArmReset() {
  trArmDanger(trResetBtn, t("trReset"), trResetProgress);
}

function trUpdateInheritNote() {
  if (trInheritNoteEl) {
    let note;
    if (trLongMode) note = t("trInheritLong");
    else if (trMergeCols.size) note = t("trInheritMergeInfo", { cols: trMergeCols.size });
    else note = t("trInheritNoMerges");
    trInheritNoteEl.textContent = note;
  }
  if (trInheritEl) trInheritEl.disabled = trLongMode;
}

function trSetInherit(on, options = {}) {
  trInheritOn = !!on && !trLongMode;
  // Pierwsze włączenie na arkuszu ze scaleniami samo proponuje kolumny — użytkownik
  // i tak może je dowolnie zmienić, ale nie musi zaczynać od pustej listy.
  if (trInheritOn && !trInheritCols.size && trMergeCols.size && !options.keepCols) {
    trInheritCols = new Set(trMergeCols);
  }
  if (trInheritEl) trInheritEl.checked = trInheritOn;
  trUpdateInheritNote();
  trBuildInheritance();
  if (trFieldsListEl && !trFieldsPanelEl?.classList.contains("hidden")) trRenderFields();
  if (options.silent) return;
  trRenderCard();
}

function trToggleInheritCol(colIdx) {
  if (!trInheritOn || trLongMode) return;
  if (trInheritCols.has(colIdx)) trInheritCols.delete(colIdx);
  else trInheritCols.add(colIdx);
  trBuildInheritance();
  trRenderFields();
  trRenderCard();
}

function trSetAutoFields(on) {
  trAutoFields = !!on;
  if (trAutoFieldsEl) trAutoFieldsEl.checked = trAutoFields;
  if (trAutoNoteEl) trAutoNoteEl.classList.toggle("hidden", !trAutoFields);
  if (trFieldsBtn) trFieldsBtn.textContent = trAutoFields ? t("trFieldsAuto") : t("trFields");
  if (trFieldsListEl) trFieldsListEl.classList.toggle("is-auto", trAutoFields);
  trRenderCard();
}

// ── Rozmiar tekstu / blokada dotyku / Wake Lock ─────────────────────────────

function trApplyFont() {
  if (trOverlayEl) trOverlayEl.dataset.font = String(trFont);
  if (trFontBtn) trFontBtn.textContent = `A${"+".repeat(Math.max(0, trFont - 1))}`;
}

function trCycleFont() {
  const at = TR_FONT_STEPS.indexOf(trFont);
  trFont = TR_FONT_STEPS[(at + 1) % TR_FONT_STEPS.length];
  trApplyFont();
  trScheduleScrollUi(); // większa czcionka = karta może przestać się mieścić
  trPersist();
}

// Blokada dotyku: tarcza przykrywa kartę (dłoń oparta o tablet nie przewinie ani nie
// zaznaczy), natomiast dolny pasek i sam przycisk blokady zostają nad nią klikalne.
function trSetLocked(on) {
  if (on && trTurboSource) trHoldCancel(); // blokada dotyku w trakcie serii = stop
  trLocked = !!on;
  if (trOverlayEl) trOverlayEl.classList.toggle("is-locked", trLocked);
  if (trTouchShieldEl) trTouchShieldEl.classList.toggle("hidden", !trLocked);
  if (trLockBtn) {
    trLockBtn.setAttribute("aria-pressed", trLocked ? "true" : "false");
    trLockBtn.textContent = trLocked ? t("trUnlock") : t("trLock");
  }
  if (trLocked && trFieldsPanelEl) trCloseFields();
  if (trLocked && trProgressPanelEl && !trProgressPanelEl.classList.contains("hidden")) trCloseProgress();
}

async function trRequestWakeLock() {
  try {
    if (!("wakeLock" in navigator)) return;
    trWakeLock = await navigator.wakeLock.request("screen");
    trWakeLock.addEventListener("release", () => { trWakeLock = null; });
    if (trOverlayEl) trOverlayEl.classList.add("has-wakelock");
  } catch {
    /* brak zgody / nieobsługiwane — tryb działa, ekran po prostu może zgasnąć */
  }
}

function trReleaseWakeLock() {
  try { trWakeLock?.release?.(); } catch { /* ignore */ }
  trWakeLock = null;
  if (trOverlayEl) trOverlayEl.classList.remove("has-wakelock");
}

document.addEventListener("visibilitychange", () => {
  if (trIsOpen && document.visibilityState === "visible" && !trWakeLock) trRequestWakeLock();
});

// ── Baner „ten arkusz wygląda inaczej” ──────────────────────────────────────
// Pokazujemy TYLKO przy twardej zmianie i TYLKO z konkretnymi liczbami. „Coś się zmieniło”
// bez liczb nie pomaga podjąć decyzji, a decyzja należy do użytkownika — nic nie kasujemy sami.

function trFormatWhen(ts) {
  if (!ts) return "";
  try {
    const d = new Date(ts);
    if (Number.isNaN(d.getTime())) return "";
    const today = new Date();
    const sameDay = d.toDateString() === today.toDateString();
    const time = d.toLocaleTimeString(I18N[currentLang]?.locale || "pl-PL", { hour: "2-digit", minute: "2-digit" });
    if (sameDay) return t("trWhenToday", { time });
    const date = typeof formatLocalizedDateDisplay === "function"
      ? formatLocalizedDateDisplay(d, { month: "short", year: "numeric" })
      : d.toLocaleDateString();
    return `${date}, ${time}`;
  } catch {
    return "";
  }
}

function trShowChangeNotice() {
  if (!trNoticeEl) return;
  const info = trChangeInfo;
  if (!info || info.level !== "hard") {
    trNoticeEl.classList.add("hidden");
    return;
  }
  if (trNoticeTextEl) {
    const when = trFormatWhen(info.savedAt);
    trNoticeTextEl.textContent = info.remapped
      ? t("trNoticeRemapped", {
        savedRows: info.savedRows ?? "?",
        rows: info.rows,
        moved: info.moved,
        all: info.savedDone,
        lost: info.lost,
        when,
      })
      : t("trNoticeUnknown", { all: info.savedDone, when });
    // Baner mówi „nieodnalezione: 60" — bez tego zdania nie widać, że da się je obejrzeć.
    if (info.remapped && info.lost > 0) {
      trNoticeTextEl.textContent += ` ${t("trNoticeSeeProgress")}`;
    }
  }
  trNoticeEl.classList.remove("hidden");
}

function trHideChangeNotice() {
  if (trNoticeEl) trNoticeEl.classList.add("hidden");
}

// ── Panel „Postęp” ──────────────────────────────────────────────────────────
// Odpowiada na trzy pytania naraz: ile zrobione, skąd apka to wie i jak to skasować.

function trStatTile(label, value, tone = "") {
  const box = document.createElement("div");
  box.className = `tr-stat${tone ? ` ${tone}` : ""}`;
  const v = document.createElement("div");
  v.className = "tr-stat-value";
  v.textContent = String(value);
  const l = document.createElement("div");
  l.className = "tr-stat-label";
  l.textContent = label;
  box.append(v, l);
  return box;
}

function trRenderProgressPanel() {
  if (!trStatsEl) return;
  const all = trRows.length;
  const done = Math.min(trDone.size, all);
  const left = Math.max(0, all - done);
  const pct = all ? Math.round((done / all) * 100) : 0;
  trStatsEl.replaceChildren(
    trStatTile(t("trStatDone"), done, "is-done"),
    trStatTile(t("trStatLeft"), left),
    trStatTile(t("trStatPercent"), `${pct}%`),
  );

  if (trScopeNoteEl) {
    const lines = [];
    lines.push(t("trScopeRows", { rows: all, cols: trHeaders.length }));
    const store = trLoadStore();
    const saved = store.scopes?.[trScope];
    const when = trFormatWhen(saved?.ts);
    if (when) lines.push(t("trScopeLastSession", { when }));
    if (trChangeInfo?.level === "soft") lines.push(t("trScopeFileResaved"));
    if (trChangeInfo?.level === "hard") {
      lines.push(trChangeInfo.remapped
        ? t("trScopeChangedRemapped", { moved: trChangeInfo.moved, all: trChangeInfo.savedDone, lost: trChangeInfo.lost })
        : t("trScopeChangedUnknown"));
    }
    if (trChangeInfo?.level === "import") {
      lines.push(t("trScopeImported", { moved: trChangeInfo.moved, all: trChangeInfo.savedDone, lost: trChangeInfo.lost }));
    }
    // Rozbicie na tury mówi, CZEMU można temu wynikowi ufać: „dokładnie" to zero domysłu,
    // „po kluczu" to zgodny numer/ID, „po podobieństwie" to jedyny wyraźny kandydat.
    if (trChangeInfo?.remapped && trChangeInfo.moved) {
      const parts = [t("trMatchExact", { n: trChangeInfo.exact || 0 })];
      if (trChangeInfo.byKey) {
        parts.push(trChangeInfo.keyCol
          ? t("trMatchByKeyCol", { n: trChangeInfo.byKey, col: trChangeInfo.keyCol })
          : t("trMatchByKey", { n: trChangeInfo.byKey }));
      }
      if (trChangeInfo.similar) parts.push(t("trMatchSimilar", { n: trChangeInfo.similar }));
      lines.push(`${t("trMatchHow")} ${parts.join(" · ")}`);
    }
    trScopeNoteEl.replaceChildren();
    lines.forEach((text) => {
      const div = document.createElement("div");
      div.textContent = text;
      trScopeNoteEl.appendChild(div);
    });
  }

  trRenderUnmatched();
  trRenderStoreList();
}

// ── Lista nieodnalezionych ✓ ────────────────────────────────────────────────
// Samo „60 nie znaleziono" to ślepy zaułek: nie wiadomo, czego dotyczy, więc nie da się
// nic z tym zrobić. Pokazujemy więc, KTÓRE to były wiersze (podgląd zapisany razem z ✓),
// a przy każdym — najbardziej podobny wiersz w tym pliku, jeśli w ogóle jakiś jest.
// „Pokaż" tylko przeskakuje na ten wiersz. Odhaczenie zostaje decyzją użytkownika:
// odhaczenie cudzego wiersza jest przy przepisywaniu na papier najgorszym możliwym błędem.
function trRenderUnmatched() {
  if (!trUnmatchedEl) return;
  const items = Array.isArray(trUnmatched) ? trUnmatched : [];
  if (!items.length) {
    trUnmatchedEl.classList.add("hidden");
    if (trUnmatchedListEl) trUnmatchedListEl.replaceChildren();
    return;
  }
  trUnmatchedEl.classList.remove("hidden");
  if (trUnmatchedTitleEl) trUnmatchedTitleEl.textContent = t("trUnmatchedTitle", { count: items.length });
  if (trUnmatchedHintEl) trUnmatchedHintEl.textContent = t("trUnmatchedHint");
  if (!trUnmatchedListEl) return;

  const open = trUnmatchedToggleEl?.getAttribute("aria-expanded") === "true";
  if (trUnmatchedToggleEl) trUnmatchedToggleEl.textContent = open ? t("trUnmatchedHide") : t("trUnmatchedShowAll");
  trUnmatchedListEl.classList.toggle("hidden", !open);
  if (!open) return;

  trUnmatchedListEl.replaceChildren();
  // Sufit na długość listy: przy setkach pozycji panel przestaje być czytelny, a sens
  // listy jest taki, żeby dało się ją przejrzeć okiem.
  const shown = items.slice(0, 200);
  shown.forEach((it) => {
    const row = document.createElement("div");
    row.className = "tr-unmatched-item";
    const text = document.createElement("div");
    text.className = "tr-unmatched-text";
    text.textContent = it.prev || t("trUnmatchedNoPreview");
    row.appendChild(text);
    if (it.best) {
      const near = document.createElement("button");
      near.type = "button";
      near.className = "btn btn-xs ghost";
      near.textContent = t("trUnmatchedNear", { pct: Math.round((it.score || 0) * 100) });
      near.addEventListener("click", () => trJumpToKey(it.best));
      row.appendChild(near);
    } else {
      const none = document.createElement("span");
      none.className = "tr-unmatched-none";
      none.textContent = t("trUnmatchedNoCandidate");
      row.appendChild(none);
    }
    trUnmatchedListEl.appendChild(row);
  });
  if (items.length > shown.length) {
    const more = document.createElement("div");
    more.className = "tr-unmatched-none";
    more.textContent = t("trUnmatchedMore", { n: items.length - shown.length });
    trUnmatchedListEl.appendChild(more);
  }
}

// Skok na wskazany wiersz BEZ odhaczania — użytkownik porównuje kartę z papierem i decyduje.
function trJumpToKey(rowKey) {
  if (!rowKey) return;
  const at = trOrder.findIndex((i) => trKeyOf(trRows[i]) === rowKey);
  if (at < 0) {
    toast(t("trUnmatchedGone"), "warning");
    return;
  }
  trPos = at;
  trCloseProgress();
  trRenderCard();
  toast(t("trUnmatchedJumped"), "info");
}

// Lista WSZYSTKICH zapamiętanych spisywań — także z innych plików. Bez niej „wyczyść"
// dotyczyłoby tylko tego, co akurat otwarte, a pamięć rosłaby w tle niewidzialnie.
function trRenderStoreList() {
  if (!trStoreListEl) return;
  const store = trLoadStore();
  const scopes = store.scopes || {};
  const entries = Object.entries(scopes)
    .map(([key, rec]) => ({ key, rec }))
    .sort((a, b) => (b.rec?.ts || 0) - (a.rec?.ts || 0));
  trStoreListEl.replaceChildren();
  if (!entries.length) {
    const empty = document.createElement("div");
    empty.className = "tr-store-empty";
    empty.textContent = t("trStoreEmpty");
    trStoreListEl.appendChild(empty);
    return;
  }
  entries.forEach(({ key, rec }) => {
    const item = document.createElement("div");
    item.className = "tr-store-item";
    if (key === trScope) item.classList.add("is-current");

    const main = document.createElement("div");
    main.className = "tr-store-main";
    const nameEl = document.createElement("div");
    nameEl.className = "tr-store-name";
    // klucz = plik::arkusz::tryb — rozbijamy, żeby dało się to przeczytać
    const parts = key.split("::");
    nameEl.textContent = parts[0] || key;
    const metaEl = document.createElement("div");
    metaEl.className = "tr-store-meta";
    const doneN = Array.isArray(rec?.done) ? rec.done.length : 0;
    const total = Number(rec?.rowsTotal) || 0;
    const bits = [parts[1] || "", t("trStoreDone", { done: doneN, all: total || "?" })];
    const when = trFormatWhen(rec?.ts);
    if (when) bits.push(when);
    metaEl.textContent = bits.filter(Boolean).join(" · ");
    main.append(nameEl, metaEl);

    const del = document.createElement("button");
    del.type = "button";
    del.className = "btn btn-xs ghost tr-store-del";
    del.textContent = "✕";
    del.setAttribute("aria-label", `${t("trStoreDelete")}: ${parts[0] || key}`);
    del.addEventListener("click", () => trDeleteScope(key));

    item.append(main);
    // Import ma sens tylko dla INNEGO pliku (nie bieżącego) i tylko gdy ten zapis ma
    // odciski treści — bez nich nie da się bezpiecznie dopasować wierszy.
    const hasSigs = Array.isArray(rec?.doneSig) && rec.doneSig.some(Boolean);
    if (key !== trScope && hasSigs) {
      const imp = document.createElement("button");
      imp.type = "button";
      imp.className = "btn btn-xs ghost tr-store-import";
      imp.textContent = t("trStoreImportBtn");
      imp.setAttribute("aria-label", `${t("trStoreImportAria")}: ${parts[0] || key}`);
      imp.addEventListener("click", () => trImportFromScope(key));
      item.append(imp);
    }
    item.append(del);
    trStoreListEl.appendChild(item);
  });
}

// Kasowanie + odświeżenie karty MUSI iść bez zapisu: trRenderCard woła trPersist,
// więc „wyczyść wszystko" natychmiast odtwarzałoby przed chwilą skasowany wpis.
// Po wyczyszczeniu pamięć ma zostać pusta aż do pierwszej realnej akcji użytkownika.
function trWithoutPersist(fn) {
  const before = trBulkMode;
  trBulkMode = true;
  try { fn(); } finally { trBulkMode = before; }
}

// Import zaznaczeń Z INNEGO zapamiętanego pliku do BIEŻĄCEGO — po treści wiersza,
// nigdy po pozycji. To dokładnie ten sam silnik, co przy „ten sam plik, ale zmieniony"
// (trRemapDone), tylko odpalony ręcznie i między RÓŻNYMI nazwami plików: np. gdy wczoraj
// spisałeś 150 wierszy w „Obieg.xlsx", a dziś wczytałeś nowszą wersję zapisaną pod inną
// nazwą — bez importu apka widziałaby to jako zupełnie nowy, pusty plik.
// Addytywne i nieniszczące: tylko DOKŁADA dopasowane ✓ do już zaznaczonych, źródłowy
// zapis zostaje nietknięty (można go potem osobno usunąć przyciskiem ✕).
function trImportFromScope(key) {
  const store = trLoadStore();
  const rec = store.scopes?.[key];
  if (!rec) return;
  const savedDone = Array.isArray(rec.done) ? rec.done : [];
  const sourceName = key.split("::")[0] || key;
  // Zapisy v2 niosą nazwy kolumn, więc kolumny „na dziś" odpadają po obu stronach same
  // z siebie. Zapisy sprzed v2 (same hashe pozycyjne) zostają dopasowaniem awaryjnym —
  // nie da się ich naprawić wstecz, ale przy pierwszym ✓ dostają już komplet odcisków.
  const remap = trMatchDone(rec);
  if (!remap.moved) {
    toast(t("trImportNone"), "warning");
    trUnmatched = remap.unmatched;
    trRenderProgressPanel();
    return;
  }
  remap.keys.forEach((k) => trDone.add(k));
  trUnmatched = remap.unmatched;
  trChangeInfo = {
    level: "import",
    rows: trFingerprint?.rows ?? trRows.length,
    savedDone: savedDone.length,
    moved: remap.moved,
    exact: remap.exact,
    byKey: remap.byKey,
    similar: remap.similar,
    keyCol: remap.keyCol,
    lost: remap.lost,
    savedAt: rec?.ts || 0,
    remapped: true,
  };
  trRebuildOrder(trCurrentKey());
  trRenderCard();
  trRenderProgressPanel();
  trPersist();
  toast(t("trImportDone", { moved: remap.moved, all: savedDone.length, lost: remap.lost, name: sourceName }), "success");
}

function trDeleteScope(key) {
  const store = trLoadStore();
  if (store.scopes) delete store.scopes[key];
  trSaveStore(store);
  if (key === trScope) {
    // Skasowaliśmy zapis otwartego arkusza — stan w pamięci musi za tym pójść,
    // inaczej najbliższy zapis wskrzesiłby go z powrotem.
    trWithoutPersist(() => {
      trDone.clear();
      trClearSigs();
      trUnmatched = [];
      trChangeInfo = null;
      trHideChangeNotice();
      trRebuildOrder(null);
      trPos = 0;
      trRenderCard();
    });
  }
  trRenderStoreList();
  if (trStatsEl) trRenderProgressPanel();
  toast(t("trStoreDeleted"), "success");
}

function trClearAllScopes() {
  try {
    localStorage.removeItem(TR_STORE_KEY);
  } catch { /* prywatne okno — i tak nie było czego kasować */ }
  trWithoutPersist(() => {
    trDone.clear();
    trClearSigs();
    trUnmatched = [];
    trChangeInfo = null;
    trHideChangeNotice();
    trRebuildOrder(null);
    trPos = 0;
    trRenderCard();
  });
  trRenderProgressPanel();
  toast(t("trStoreClearedAll"), "success");
}

function trOpenProgress() {
  if (!trProgressPanelEl) return;
  if (trFieldsPanelEl && !trFieldsPanelEl.classList.contains("hidden")) trCloseFields();
  trRenderProgressPanel();
  trProgressPanelEl.classList.remove("hidden");
  if (trProgressBtn) trProgressBtn.setAttribute("aria-expanded", "true");
  const first = trProgressPanelEl.querySelector("button");
  if (first) first.focus();
}

function trCloseProgress() {
  if (!trProgressPanelEl) return;
  trDisarmAll();
  trProgressPanelEl.classList.add("hidden");
  if (trProgressBtn) {
    trProgressBtn.setAttribute("aria-expanded", "false");
    trProgressBtn.focus();
  }
}

// ── Panel „Pola” ────────────────────────────────────────────────────────────

function trRenderFields() {
  if (!trFieldsListEl) return;
  trFieldsListEl.replaceChildren();
  trFieldOrder.forEach((colIdx, pos) => {
    const item = document.createElement("div");
    item.className = "tr-field-row";

    const cb = document.createElement("input");
    cb.type = "checkbox";
    cb.id = `trfield-${colIdx}`;
    cb.checked = trSelected.has(colIdx);
    cb.addEventListener("change", () => {
      if (cb.checked) trSelected.add(colIdx);
      else trSelected.delete(colIdx);
      trRenderCard();
    });

    const label = document.createElement("label");
    label.htmlFor = cb.id;
    label.className = "tr-field-name";
    label.textContent = trFieldLabel(colIdx);

    const actions = document.createElement("div");
    actions.className = "tr-field-actions";

    const inh = document.createElement("button");
    inh.type = "button";
    inh.className = "btn btn-xs ghost tr-inherit-btn";
    inh.textContent = "⤓";
    inh.disabled = !trInheritOn || trLongMode;
    inh.classList.toggle("is-on", trInheritCols.has(colIdx));
    inh.classList.toggle("is-merge", trMergeCols.has(colIdx));
    inh.setAttribute("aria-pressed", trInheritCols.has(colIdx) ? "true" : "false");
    inh.setAttribute("aria-label", `${t("trInheritColAria")}: ${trFieldLabel(colIdx)}`);
    inh.addEventListener("click", () => trToggleInheritCol(colIdx));
    actions.appendChild(inh);

    const up = document.createElement("button");
    up.type = "button";
    up.className = "btn btn-xs ghost tr-move-up";
    up.textContent = "▲";
    up.setAttribute("aria-label", `${t("moveUp")}: ${trFieldLabel(colIdx)}`);
    up.disabled = pos === 0;
    up.addEventListener("click", () => trMoveField(pos, -1));
    const down = document.createElement("button");
    down.type = "button";
    down.className = "btn btn-xs ghost tr-move-down";
    down.textContent = "▼";
    down.setAttribute("aria-label", `${t("moveDown")}: ${trFieldLabel(colIdx)}`);
    down.disabled = pos === trFieldOrder.length - 1;
    down.addEventListener("click", () => trMoveField(pos, 1));
    actions.append(up, down);

    item.append(cb, label, actions);
    trFieldsListEl.appendChild(item);
  });
  if (typeof ensureKeyboardReachable === "function") ensureKeyboardReachable(trFieldsListEl);
}

function trMoveField(pos, delta) {
  const next = pos + delta;
  if (next < 0 || next >= trFieldOrder.length) return;
  const [moved] = trFieldOrder.splice(pos, 1);
  trFieldOrder.splice(next, 0, moved);
  trRenderFields();
  trRenderCard();
  const list = trFieldsListEl.querySelectorAll(".tr-field-row");
  const btn = list[next]?.querySelector(`.tr-field-actions .tr-move-${delta < 0 ? "up" : "down"}`);
  if (btn && !btn.disabled) btn.focus();
}

function trOpenFields() {
  if (!trFieldsPanelEl) return;
  if (trProgressPanelEl && !trProgressPanelEl.classList.contains("hidden")) trCloseProgress();
  trRenderFields();
  trFieldsPanelEl.classList.remove("hidden");
  if (trFieldsBtn) trFieldsBtn.setAttribute("aria-expanded", "true");
  const first = trFieldsPanelEl.querySelector("input, button");
  if (first) first.focus();
}

function trCloseFields() {
  if (!trFieldsPanelEl) return;
  trFieldsPanelEl.classList.add("hidden");
  if (trFieldsBtn) {
    trFieldsBtn.setAttribute("aria-expanded", "false");
    trFieldsBtn.focus();
  }
  trRenderCard();
}

// ── Otwarcie / zamknięcie ───────────────────────────────────────────────────

function trBackgroundInert(on) {
  document.querySelectorAll(".app, .hero-overlay").forEach((el) => {
    if (on) el.setAttribute("inert", "");
    else el.removeAttribute("inert");
  });
}

function openTranscribe() {
  if (!trOverlayEl) return;
  const model = (typeof currentDisplayModel !== "undefined" && currentDisplayModel) || getDisplayModel();
  if (!model?.headers?.length || !model?.rows?.length) {
    toast(t("noDataForExport"), "warning");
    return;
  }

  trRows = model.rows.slice();
  trHeaders = model.headers.slice();
  trRowHeadFormatter = typeof model.rowHeadFormatter === "function" ? model.rowHeadFormatter : null;
  trScope = trScopeKey(model);

  const store = trLoadStore();
  const saved = store.scopes?.[trScope] || null;
  trFont = TR_FONT_STEPS.includes(store.font) ? store.font : 2;
  trHideDone = !!store.hideDone;

  const validCol = (i) => Number.isInteger(i) && i >= 0 && i < trHeaders.length;
  if (saved && Array.isArray(saved.order) && saved.order.length) {
    // Układ z poprzedniej sesji, ale arkusz mógł zmienić liczbę kolumn — dokładamy brakujące.
    trFieldOrder = saved.order.filter(validCol);
    trHeaders.forEach((_, i) => { if (!trFieldOrder.includes(i)) trFieldOrder.push(i); });
    trSelected = new Set((saved.sel || []).filter(validCol));
    if (!trSelected.size) trSelected = trDefaultSelection();
  } else {
    trFieldOrder = trHeaders.map((_, i) => i);
    trSelected = trDefaultSelection();
  }
  // ── Czy to nadal ten sam arkusz? ───────────────────────────────────────────
  // Kolejność jest istotna: najpierw kolumny zmienne (bo wchodzą w odcisk), potem odcisk
  // arkusza, dopiero na końcu decyzja, co zrobić z zapamiętanymi ✓.
  trSigCache.clear();
  trCellsCache.clear();
  trClearSigs();
  trVolatileCols = trCollectVolatileCols(saved?.volCols);
  trPickSigCols();          // kolumny odcisku — muszą być gotowe PRZED liczeniem czegokolwiek
  trFingerprint = trSheetFingerprint();

  const savedDone = Array.isArray(saved?.done) ? saved.done : [];
  const savedSigs = Array.isArray(saved?.doneSig) ? saved.doneSig : [];
  const level = savedDone.length ? trCompareFingerprint(saved?.sig, trFingerprint) : "same";
  trChangeInfo = null;
  trUnmatched = [];   // lista dotyczy KONKRETNEGO wczytania — nie może przeżyć zmiany arkusza

  const hasAnySig = savedSigs.some(Boolean)
    || (Array.isArray(saved?.doneCells) && saved.doneCells.some(Boolean));
  if (level === "hard" && hasAnySig) {
    // Plik inny, ale mamy odciski treści → przenosimy ✓ tam, gdzie ich miejsce.
    const remap = trMatchDone(saved);
    trDone = remap.keys;
    trChangeInfo = {
      level: "hard",
      savedRows: saved?.sig?.rows ?? saved?.rowsTotal ?? null,
      rows: trFingerprint.rows,
      savedDone: savedDone.length,
      moved: remap.moved,
      exact: remap.exact,
      byKey: remap.byKey,
      similar: remap.similar,
      keyCol: remap.keyCol,
      lost: remap.lost,
      savedAt: saved?.ts || 0,
      remapped: true,
    };
    trUnmatched = remap.unmatched;
  } else if (level === "hard") {
    // Zapis sprzed wersji z odciskami — kluczy nie ma jak zweryfikować. Zostawiamy je
    // (nic nie kasujemy bez pytania), ale mówimy wprost, że mogą być nie na swoim miejscu.
    trDone = new Set(savedDone);
    trChangeInfo = {
      level: "hard",
      savedRows: saved?.sig?.rows ?? saved?.rowsTotal ?? null,
      rows: trFingerprint.rows,
      savedDone: savedDone.length,
      moved: 0,
      lost: 0,
      savedAt: saved?.ts || 0,
      remapped: false,
    };
  } else {
    trDone = new Set(savedDone);
    const savedCells = Array.isArray(saved?.doneCells) ? saved.doneCells : [];
    const savedPrev = Array.isArray(saved?.donePrev) ? saved.donePrev : [];
    savedDone.forEach((key, i) => {
      if (savedSigs[i]) trDoneSigs.set(key, savedSigs[i]);
      if (savedCells[i]) trDoneCells.set(key, savedCells[i]);
      if (savedPrev[i]) trDonePrev.set(key, savedPrev[i]);
    });
    // Zapis mógł powstać w wersji sprzed odcisków treści (albo pojedynczy klucz je zgubił).
    // Skoro fingerprint arkusza się zgadza (same/soft), pozycje są wciąż te same — możemy
    // bezpiecznie DOPISAĆ brakujący odcisk od razu, zamiast czekać aż ktoś ręcznie tknie
    // wiersz. Inaczej te konkretne ✓ zostałyby bez odcisku w nieskończoność i przy KOLEJNEJ
    // realnej zmianie pliku znów wpadłyby w gałąź „nie da się zweryfikować".
    if (trDone.size) {
      const missing = new Set();
      // Brakuje CZEGOKOLWIEK z kompletu (hash wiersza / odcisk kolumnowy / podgląd) —
      // pozycje są wciąż te same, więc dopisujemy od razu, zamiast czekać, aż ktoś ruszy wiersz.
      trDone.forEach((key) => {
        if (!trDoneSigs.has(key) || !trDoneCells.has(key) || !trDonePrev.has(key)) missing.add(key);
      });
      if (missing.size) {
        for (const row of trRows) {
          const key = trKeyOf(row);
          if (!missing.has(key)) continue;
          trMarkSig(key, row);
          missing.delete(key);
          if (!missing.size) break;
        }
      }
    }
    if (level === "soft") {
      trChangeInfo = { level: "soft", rows: trFingerprint.rows, savedDone: savedDone.length, savedAt: saved?.ts || 0 };
    }
  }
  trAutoFields = !!saved?.auto;
  trLongMode = model.mode === "long";
  trMergeRanges = trDetectMergeRanges();
  trMergeCols = new Set(trMergeRanges.keys());
  const validCol2 = (i) => Number.isInteger(i) && i >= 0 && i < trHeaders.length;
  trInheritCols = new Set(Array.isArray(saved?.inheritCols) ? saved.inheritCols.filter(validCol2) : []);
  trInheritOn = !!saved?.inherit && !trLongMode;

  trRebuildOrder(null);
  // Wznowienie: wracamy na zapamiętany wiersz, a jak go nie ma — na pierwszy nieodhaczony.
  const resumeKey = saved?.cursor;
  let at = resumeKey ? trOrder.findIndex((i) => trKeyOf(trRows[i]) === resumeKey) : -1;
  if (at < 0) at = trOrder.findIndex((i) => !trDone.has(trKeyOf(trRows[i])));
  trPos = at >= 0 ? at : 0;

  if (trSourceEl) {
    const sheet = (typeof sheetSelect !== "undefined" && sheetSelect?.value) || "";
    trSourceEl.textContent = [currentFileName || "", sheet].filter(Boolean).join(" · ");
  }
  if (trHideDoneEl) trHideDoneEl.checked = trHideDone;
  trSetAutoFields(trAutoFields);
  trSetInherit(trInheritOn, { keepCols: true, silent: true });
  trApplyFont();
  trSetLocked(false);
  if (trFieldsPanelEl) trFieldsPanelEl.classList.add("hidden");
  if (trProgressPanelEl) trProgressPanelEl.classList.add("hidden");

  trReturnFocusEl = document.activeElement;
  trOverlayEl.classList.remove("hidden");
  document.body.classList.add("tr-active");
  trBackgroundInert(true);
  trIsOpen = true;
  trHideUndo();
  trResetScroll();
  trShowChangeNotice();
  trRenderCard();
  trRequestWakeLock();
  if (trMarkBtn) trMarkBtn.focus();
  // Przy twardej zmianie pliku mówi baner — i to on ma przyciski decyzji. Toast tylko
  // zasłaniałby te przyciski przez pierwsze sekundy, czyli dokładnie wtedy, gdy są potrzebne.
  if (trChangeInfo?.level !== "hard" && trDone.size) {
    toast(t("trResumed", { done: trDone.size }), "info");
  }
}

function closeTranscribe() {
  if (!trOverlayEl || !trIsOpen) return;
  trHoldCancel();
  trBulkMode = false;
  trDisarmAll();
  trHideUndo();
  trPersist();
  trIsOpen = false;
  trSetLocked(false);
  trOverlayEl.classList.add("hidden");
  document.body.classList.remove("tr-active");
  trBackgroundInert(false);
  trReleaseWakeLock();
  const back = trReturnFocusEl;
  trReturnFocusEl = null;
  if (back && document.contains(back) && !back.closest("[inert]")) back.focus();
  else if (trBtn) trBtn.focus();
}

// ── Klawiatura ──────────────────────────────────────────────────────────────
// Capture na document: dopóki nakładka jest otwarta, klawisze NIE docierają do
// globalnego handlera w bootstrap.js (inaczej strzałki przesuwałyby zaznaczenie
// w tabeli pod spodem, a Cmd+Shift+F otwierałby okno szukania).

function trFocusables() {
  if (!trOverlayEl) return [];
  const sel = 'button:not([disabled]), input:not([disabled]), [tabindex]:not([tabindex="-1"])';
  return Array.from(trOverlayEl.querySelectorAll(sel)).filter((el) => {
    if (el.closest(".hidden")) return false;
    if (el.getAttribute("tabindex") === "-1") return false; // np. pigułka „więcej ↓" — klikalna, ale poza Tabem
    return el.offsetParent !== null || el === document.activeElement;
  });
}

function trTrapTab(e) {
  const list = trFocusables();
  if (!list.length) return;
  const first = list[0];
  const last = list[list.length - 1];
  if (e.shiftKey && document.activeElement === first) {
    e.preventDefault();
    last.focus();
  } else if (!e.shiftKey && document.activeElement === last) {
    e.preventDefault();
    first.focus();
  }
}

document.addEventListener("keydown", (e) => {
  if (!trIsOpen) return;
  e.stopPropagation();

  if (e.key === "Escape") {
    e.preventDefault();
    if (trFieldsPanelEl && !trFieldsPanelEl.classList.contains("hidden")) trCloseFields();
    else if (trProgressPanelEl && !trProgressPanelEl.classList.contains("hidden")) trCloseProgress();
    else if (trLocked) trSetLocked(false);
    else closeTranscribe();
    return;
  }
  if (e.key === "Tab") {
    trTrapTab(e);
    return;
  }
  if (trFieldsPanelEl && trFieldsPanelEl.contains(e.target)) return;
  if (trProgressPanelEl && trProgressPanelEl.contains(e.target)) return;

  const tag = String(e.target?.tagName || "").toLowerCase();
  if (tag === "input" || tag === "select" || tag === "textarea") return;

  switch (e.key) {
    case "ArrowRight":
    case "PageDown":
      e.preventDefault();
      trGo(1);
      break;
    case "ArrowLeft":
    case "PageUp":
      e.preventDefault();
      trGo(-1);
      break;
    case "Home":
      e.preventDefault();
      trGoEdge(-1);
      break;
    case "End":
      e.preventDefault();
      trGoEdge(1);
      break;
    case " ":
    case "Enter": {
      // Na „Spisane i dalej" spacja jest PRZYTRZYMYWALNA (szybkie odhaczanie), więc
      // blokujemy natywną aktywację przycisku i sami decydujemy przy puszczeniu klawisza.
      const onMark = e.target === trMarkBtn;
      if (tag === "button" && !onMark) return; // inny przycisk niech zadziała sam
      e.preventDefault();
      if (!e.repeat) trHoldStart("key");
      break;
    }
    default:
      break;
  }
}, true);

// Puszczenie klawisza: albo kończy tryb szybki, albo — gdy nie zdążył wystartować —
// wykonuje zwykłe „Spisane i dalej". Capture, bo keydown wyżej też jest w capture.
document.addEventListener("keyup", (e) => {
  if (!trIsOpen || trTurboSource !== "key") return;
  if (e.key !== " " && e.key !== "Enter") return;
  e.stopPropagation();
  e.preventDefault();
  const fired = trHoldCancel();
  if (!fired) trMarkAndNext();
}, true);

// Utrata fokusu / przejście w tło w trakcie trzymania — nie zostawiamy pętli w biegu.
window.addEventListener("blur", () => { if (trTurboSource) trHoldCancel(); });
document.addEventListener("visibilitychange", () => {
  if (document.visibilityState !== "visible" && trTurboSource) trHoldCancel();
});

// ── Podpięcie ───────────────────────────────────────────────────────────────

if (trBtn) trBtn.addEventListener("click", openTranscribe);
if (trCloseBtn) trCloseBtn.addEventListener("click", closeTranscribe);
if (trMarkBtn) {
  trMarkBtn.addEventListener("click", () => {
    // Klik doleci też po zakończeniu przytrzymania — wtedy go połykamy, żeby seria
    // nie dostała jednego wiersza w bonusie.
    if (trSuppressNextClick) { trSuppressNextClick = false; return; }
    trMarkAndNext();
  });
  trMarkBtn.addEventListener("pointerdown", (e) => {
    if (e.button != null && e.button > 0) return; // tylko lewy / dotyk / pióro
    trHoldStart("pointer");
  });
  const endPointerHold = () => {
    if (trTurboSource !== "pointer") return;
    if (trHoldCancel()) trSuppressNextClick = true;
  };
  ["pointerup", "pointercancel", "pointerleave"].forEach((evt) => trMarkBtn.addEventListener(evt, endPointerHold));
  // Palec puszczony poza przyciskiem (albo mysz zwolniona gdzie indziej) też kończy serię.
  window.addEventListener("pointerup", endPointerHold);
}
if (trUndoBtn) trUndoBtn.addEventListener("click", trUndoBurst);
if (trMarkChipEl) trMarkChipEl.addEventListener("click", trToggleDone);
if (trPrevBtn) trPrevBtn.addEventListener("click", () => trGo(-1));
if (trNextBtn) trNextBtn.addEventListener("click", () => trGo(1));
if (trHideDoneEl) trHideDoneEl.addEventListener("change", () => trSetHideDone(trHideDoneEl.checked));
if (trFontBtn) trFontBtn.addEventListener("click", trCycleFont);
if (trLockBtn) trLockBtn.addEventListener("click", () => trSetLocked(!trLocked));
if (trResetBtn) trResetBtn.addEventListener("click", trArmReset);
if (trAutoFieldsEl) trAutoFieldsEl.addEventListener("change", () => trSetAutoFields(trAutoFieldsEl.checked));
if (trInheritEl) trInheritEl.addEventListener("change", () => trSetInherit(trInheritEl.checked));
if (trFieldsBtn) {
  trFieldsBtn.addEventListener("click", () => {
    if (trFieldsPanelEl && trFieldsPanelEl.classList.contains("hidden")) trOpenFields();
    else trCloseFields();
  });
}
if (trFieldsDoneBtn) trFieldsDoneBtn.addEventListener("click", trCloseFields);
if (trProgressBtn) {
  trProgressBtn.addEventListener("click", () => {
    if (trProgressPanelEl && trProgressPanelEl.classList.contains("hidden")) trOpenProgress();
    else trCloseProgress();
  });
}
if (trProgressDoneBtn) trProgressDoneBtn.addEventListener("click", trCloseProgress);
if (trUnmatchedToggleEl) {
  trUnmatchedToggleEl.addEventListener("click", () => {
    const open = trUnmatchedToggleEl.getAttribute("aria-expanded") === "true";
    trUnmatchedToggleEl.setAttribute("aria-expanded", open ? "false" : "true");
    trRenderUnmatched();
  });
}
if (trStoreClearAllBtn) trStoreClearAllBtn.addEventListener("click", () => trArmDanger(trStoreClearAllBtn, t("trStoreClearAll"), trClearAllScopes));
if (trNoticeKeepBtn) trNoticeKeepBtn.addEventListener("click", trHideChangeNotice);
if (trNoticeResetBtn) {
  trNoticeResetBtn.addEventListener("click", () => {
    trResetProgress();
    trHideChangeNotice();
  });
}
if (trFieldsAllBtn) {
  trFieldsAllBtn.addEventListener("click", () => {
    trFieldOrder.forEach((i) => trSelected.add(i));
    trRenderFields();
    trRenderCard();
  });
}
if (trFieldsNoneBtn) {
  trFieldsNoneBtn.addEventListener("click", () => {
    trSelected.clear();
    trRenderFields();
    trRenderCard();
  });
}
if (trTouchShieldEl) {
  // Tarcza połyka dotyk i klik — ale nie „na ślepo”: podwójny tap odblokowuje,
  // żeby nie dało się zamknąć w trybie bez wyjścia, gdy przycisk zniknie z pola widzenia.
  ["touchstart", "touchmove", "pointerdown", "click", "wheel"].forEach((evt) => {
    trTouchShieldEl.addEventListener(evt, (e) => { e.preventDefault(); e.stopPropagation(); }, { passive: false });
  });
  trTouchShieldEl.addEventListener("dblclick", () => trSetLocked(false));
}

// Hook testowy — Playwright steruje trybem bez klikania po pikselach.
window.__transcribe = {
  open: openTranscribe,
  close: closeTranscribe,
  state: () => ({
    open: trIsOpen,
    pos: trPos,
    total: trOrder.length,
    rows: trRows.length,
    done: trDone.size,
    cols: trVisibleCols(trCurrentRow()),
    auto: trAutoFields,
    skipped: trSkippedCount(trCurrentRow()),
    inherit: trInheritOn,
    inheritCols: Array.from(trInheritCols),
    mergeCols: Array.from(trMergeCols),
    longMode: trLongMode,
    fields: (() => {
      const row = trCurrentRow();
      if (!row) return [];
      return trVisibleCols(row).map((ci) => {
        const r = trResolveField(row, ci);
        return { col: ci, label: trFieldLabel(ci), text: r.text, from: r.from };
      });
    })(),
    locked: trLocked,
    font: trFont,
    hideDone: trHideDone,
    scrollTop: trStageEl ? trStageEl.scrollTop : 0,
    canScroll: !!(trStageEl && trStageEl.scrollHeight - trStageEl.clientHeight > 4),
    overflowUi: !!trStageWrapEl?.classList.contains("has-overflow"),
    atBottom: !!trStageWrapEl?.classList.contains("at-bottom"),
    turbo: !!trTurboTimer,
    volatileCols: Array.from(trVolatileCols),
    changed: trChangeInfo ? { ...trChangeInfo } : null,
    unmatched: trUnmatched.map((u) => ({ prev: u.prev, best: u.best, score: u.score })),
    sigCols: trSigCols.map((c) => c.name),
    undoVisible: !!(trUndoBtn && !trUndoBtn.classList.contains("hidden")),
    burst: trBurstKeys.length,
    values: (() => {
      const row = trCurrentRow();
      return row ? trVisibleCols(row).map((ci) => String(getDisplayValue(row, ci) ?? "")) : [];
    })(),
  }),
  mark: trMarkAndNext,
  isDone: (key) => trDone.has(key),
  openProgress: trOpenProgress,
  closeProgress: trCloseProgress,
  clearAllScopes: trClearAllScopes,
  holdStart: (source = "key") => trHoldStart(source),
  holdCancel: () => trHoldCancel(),
  undoBurst: trUndoBurst,
  scrollBy: (px) => { if (trStageEl) { trStageEl.scrollTop += px; trUpdateScrollUi(); } },
  go: trGo,
  setHideDone: trSetHideDone,
  setAutoFields: trSetAutoFields,
  setInherit: (on) => trSetInherit(on),
  toggleInheritCol: trToggleInheritCol,
  reset: trResetProgress,
  setFields: (cols) => {
    trSelected = new Set(cols);
    trRenderCard();
  },
  moveField: trMoveField,
};
