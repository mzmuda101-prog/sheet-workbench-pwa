// render-budget.js — limit wyświetlanych wierszy dobierany POMIAREM na tym urządzeniu.
//
// Dlaczego nie heurystyka: `hardwareConcurrency`/`deviceMemory` kłamią na iOS (każdy
// iPhone zgłasza ~4 rdzenie, pamięci nie ma wcale), więc nowy telefon dostawał ten sam
// sztywny limit 200 co stary. Zamiast zgadywać — mierzymy, ile NAPRAWDĘ kosztuje tu
// przebudowa tabeli (JS + styl + layout aż do następnej klatki), w przeliczeniu na komórkę.
//
// Dlaczego koszt na komórkę jest dobrym miernikiem: `content-visibility` NIE działa na
// wierszach tabeli (zawieranie nie dotyczy <tr> — zmierzone 2026-09-25: z nim i bez niego
// ten sam koszt), więc każdy wiersz w DOM jest przeliczany przy każdym sorcie/filtrze
// i koszt rośnie liniowo z liczbą komórek. Stały budżet czasu → liczba wierszy na miarę
// urządzenia: słabe zostaje przy 200 (jak dotąd, nigdy mniej), mocne dostaje więcej.
//
// Zasady:
// - limit wpisany ręcznie przez użytkownika (zapisany w MAX_ROWS_KEY) ZAWSZE wygrywa;
//   wyczyszczenie pola wraca do automatu;
// - automat zmienia limit tylko przy wczytaniu arkusza (tabela nie „rośnie" pod palcem);
// - średnia krocząca w localStorage, więc tryb oszczędzania baterii / starzejący się
//   telefon same obniżą limit po kilku renderach;
// - sufit komórek chroni pamięć (iOS potrafi zabić kartę PWA przy ogromnym DOM).
//
// `?rowbudget=off` wyłącza automat (A/B), `?rowbudget=reset` czyści pomiar.

const RENDER_BUDGET_KEY = "swb-render-cost-v1";
const RENDER_BUDGET_MS = 180;         // ≈ tyle, ile dziś kosztuje sort 200 wierszy na słabym sprzęcie
const RENDER_BUDGET_FLOOR_ROWS = 200; // dawny sztywny limit — automat nigdy nie schodzi niżej
const RENDER_BUDGET_MAX_ROWS = 2000;
const RENDER_BUDGET_MAX_CELLS = 24000; // 17 500 komórek było płynnych na iPhonie 15 Pro (2026-09-18)
const RENDER_BUDGET_MIN_CELLS = 1500;  // mniejsze rendery to głównie stały narzut, nie próbka

const renderBudget = (() => {
  const params = (() => { try { return new URLSearchParams(location.search); } catch (_) { return new URLSearchParams(); } })();
  const mode = params.get("rowbudget");
  if (mode === "reset") { try { localStorage.removeItem(RENDER_BUDGET_KEY); } catch (_) {} }
  const disabled = mode === "off";

  let state = null; // { c: ms na komórkę (EMA), n: liczba próbek }
  try {
    const raw = JSON.parse(localStorage.getItem(RENDER_BUDGET_KEY) || "null");
    if (raw && Number.isFinite(raw.c) && raw.c > 0) state = { c: raw.c, n: raw.n | 0 };
  } catch (_) { state = null; }
  let lastSample = null;

  function record(ms, cells) {
    if (!(cells >= RENDER_BUDGET_MIN_CELLS) || !(ms > 0)) return;
    if (document.visibilityState === "hidden") return; // karta w tle = rAF zamrożony, pomiar bez sensu
    const perCell = ms / cells;
    lastSample = { ms: Math.round(ms), cells };
    if (!state) {
      state = { c: perCell, n: 1 };
    } else {
      // Pojedyncze skoki (GC, powiadomienie, przełączenie apki) nie mogą zbić limitu
      // jednym strzałem — ale trwałe spowolnienie (oszczędzanie baterii) ma przejść.
      const alpha = perCell > state.c * 3 ? 0.05 : 0.25;
      state = { c: state.c + (perCell - state.c) * alpha, n: state.n + 1 };
    }
    try { localStorage.setItem(RENDER_BUDGET_KEY, JSON.stringify({ c: state.c, n: state.n })); } catch (_) {}
  }

  // Mierzy od t0 (start renderu) do chwili PO najbliższej klatce (rAF → wiadomość = po stylu/layoucie/malowaniu).
  function measureRender(cells, t0) {
    if (!(cells >= RENDER_BUDGET_MIN_CELLS)) return;
    requestAnimationFrame(() => {
      const ch = new MessageChannel();
      ch.port1.onmessage = () => record(performance.now() - t0, cells);
      ch.port2.postMessage(0);
    });
  }

  function recommendedRows(colCount) {
    const cols = Math.max(1, colCount | 0) + 1; // + kolumna numerów wierszy
    const cellCap = Math.floor(RENDER_BUDGET_MAX_CELLS / cols);
    if (!state) return RENDER_BUDGET_FLOOR_ROWS; // pierwsze uruchomienie: jak dawniej
    const byTime = RENDER_BUDGET_MS / (state.c * cols);
    const rows = Math.min(RENDER_BUDGET_MAX_ROWS, cellCap, byTime);
    // w dół do pełnej setki — żeby liczba w polu była „ludzka" i nie skakała o kilka wierszy
    return Math.max(RENDER_BUDGET_FLOOR_ROWS, Math.floor(rows / 100) * 100);
  }

  function isManual() {
    try { return !!localStorage.getItem(MAX_ROWS_KEY); } catch (_) { return false; }
  }

  // Wołane przy wczytaniu arkusza, zanim powstanie tabela.
  function applyToSheet(colCount) {
    if (disabled || isManual() || !maxRowsEl) return;
    maxRowsEl.value = String(recommendedRows(colCount));
  }

  function describe() {
    const cols = (typeof currentHeaders !== "undefined" && currentHeaders.length) || 20;
    if (!state) return "budżet wierszy: brak pomiaru (200)";
    const per100 = (state.c * 100).toFixed(2).replace(".", ",");
    return "budżet wierszy " + (disabled ? "WYŁ" : isManual() ? "ręczny" : "auto") + ": " + recommendedRows(cols)
      + " · " + per100 + " ms/100 kom. · próbek " + state.n
      + (lastSample ? "\nostatni render " + lastSample.ms + " ms / " + lastSample.cells + " kom." : "");
  }

  return { measureRender, recommendedRows, applyToSheet, isManual, describe, get state() { return state; } };
})();
