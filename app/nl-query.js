// nl-query.js — „✦ Zapytaj zwykłymi słowami” w oknie szybkiego szukania.
//
// „otwarte sprawy Kowalskiej dłużej niż 30 dni, ile na osobę” → szukaj „Kowalsk” + @wtoku
// + @dni>>30 i grupowanie wg kolumny osoby. Wynik to ZWYKŁE zapytanie aplikacji, widoczne i
// edytowalne w polu szukania — liczy zawsze silnik arkusza, nigdy model.
//
// 1. Reguły (bez AI, natychmiast): tryby auto, dni/tygodnie/miesiące, pierwszy/ostatni cykl,
//    „według / wg / na / per <kolumna>”. Słowa, które zostały, dopasowujemy do DANYCH
//    arkusza: najdłuższy początek wyrazu, który występuje w komórkach („Kowalskiej” → „Kowalsk”).
// 2. Gdy coś zostało niezrozumiane — przycisk „Zapytaj AI” (local-ai.js). Odpowiedź modelu
//    przechodzi przez to samo sito: tylko znane tryby, istniejące kolumny i słowa obecne w danych.

const NLQ_STOP = new Set(("pokaż pokaz wyświetl wyswietl znajdź znajdz szukaj wypisz daj chcę chce lista listę liste " +
  "wszystkie wszystkich wszystko te ten ta tych które ktore który ktory która ktora co czy jakie jaki są sa jest był była było " +
  "i oraz a albo lub ze z ze w we na do od po dla przez przy o u nad pod bez niż niz jak tylko nadal jeszcze już juz " +
  "sprawy sprawa spraw sprawę rekord rekordy rekordów wiersze wierszy pozycje pozycji osoby osób osoba wpisy wpisów " +
  "trwają trwa trwające trwajace ile policz liczba sztuk razem łącznie lacznie dni dzień dzien tygodni tydzień miesięcy miesiąca " +
  "show find list all the which that are is with in of for by and or me please").split(/\s+/));

const NLQ_RULES = [
  // kolejność ma znaczenie: zaprzeczenia przed pozytywami („niezakończone” zawiera „zakończone”)
  { re: /\b(niezakończon\w*|niezakonczon\w*|niezamknięt\w*|niezamkniet\w*|niezałatwion\w*|niezalatwion\w*|zaległ\w*|zalegl\w*|otwart\w*|w\s+toku|w\s+trakcie|w\s+realizacji|trwając\w*|trwajac\w*|aktywn\w*|open|in\s+progress|ongoing|overdue|pending)\b/giu, token: "@wtoku", label: "nlqOpen", needsState: true },
  { re: /\b(wszystk\w*\s+zamknięt\w*|wszystk\w*\s+zamkniet\w*|all\s+closed)\b/giu, token: "@wszystkozamkniete", label: "nlqAllClosed", needsState: true },
  { re: /\b(zamknięt\w*|zamkniet\w*|zakończon\w*|zakonczon\w*|skończon\w*|skonczon\w*|załatwion\w*|zalatwion\w*|rozliczon\w*|sfinalizowan\w*|closed|finished|done|completed)\b/giu, token: "@zamkniete", label: "nlqClosed", needsState: true },
  { re: /\b(bez\s+(startu|początku|poczatku|daty\s+(rozpoczęcia|rozpoczecia|startu))|no\s+start)\b/giu, token: "@bezstartu", label: "nlqNoStart", needsState: true },
  { re: /\b(niezaczęt\w*|niezaczet\w*|nierozpoczęt\w*|nierozpoczet\w*|nothing\s+started)\b/giu, token: "@niezaczete", label: "nlqNone" },
  { re: /\b(ostatni\w*(\s+(cykl\w*|etap\w*|blok\w*|wpis\w*))?|last(\s+cycle)?)\b/giu, token: "@ostatni", label: "nlqLast" },
  { re: /\b(pierwsz\w*(\s+(cykl\w*|etap\w*|blok\w*|wpis\w*))?|first(\s+cycle)?)\b/giu, token: "@pierwszy", label: "nlqFirst" },
];
const NLQ_DAYS_RE = /\b(dłużej|dluzej|ponad|powyżej|powyzej|więcej|wiecej|over|more|longer|starsz\w*|older)\s*(niż|niz|od|than)?\s*(\d+(?:[.,]\d+)?)\s*(dni|dnia|dzień|dzien|days?|tyg\w*|tydzie\w*|weeks?|miesi\w*|mies\.?|months?|lat|lata|rok\w*|years?)\b/iu;
// Liczebniki słowami → cyfry („dwóch miesięcy”, „pół roku”); jednostka bez liczby = 1 („dłużej niż miesiąc”).
const NLQ_NUMBER_WORDS = [
  [/(?<!\p{L})(jeden|jednego|jednej|jedna)(?!\p{L})/giu, "1"], [/(?<!\p{L})(dwa|dwóch|dwoch|dwu|dwie)(?!\p{L})/giu, "2"], [/(?<!\p{L})(trzy|trzech)(?!\p{L})/giu, "3"],
  [/(?<!\p{L})(cztery|czterech)(?!\p{L})/giu, "4"], [/(?<!\p{L})(pięć|piec|pięciu|pieciu)(?!\p{L})/giu, "5"], [/(?<!\p{L})(sześć|szesc|sześciu|szesciu)(?!\p{L})/giu, "6"],
  [/(?<!\p{L})(siedem|siedmiu)(?!\p{L})/giu, "7"], [/(?<!\p{L})(osiem|ośmiu|osmiu)(?!\p{L})/giu, "8"], [/(?<!\p{L})(dziewięć|dziewiec|dziewięciu|dziewieciu)(?!\p{L})/giu, "9"],
  [/(?<!\p{L})(dziesięć|dziesiec|dziesięciu|dziesieciu)(?!\p{L})/giu, "10"], [/(?<!\p{L})(czternaście|czternastu)(?!\p{L})/giu, "14"], [/(?<!\p{L})(piętnaście|pietnascie|piętnastu|pietnastu)(?!\p{L})/giu, "15"],
  [/(?<!\p{L})(dwadzieścia|dwadziescia|dwudziestu)(?!\p{L})/giu, "20"], [/(?<!\p{L})(trzydzieści|trzydziesci|trzydziestu)(?!\p{L})/giu, "30"], [/(?<!\p{L})(pół|pol)(?!\p{L})/giu, "0.5"],
];
function nlqNormalizeNumbers(text) {
  let s = String(text);
  // tylko przed jednostką czasu — „jedna sprawa” zostaje słowem
  const unit = /^\s+(dni|dnia|dzień|dzien|tyg\S*|tydz\S*|miesi\S*|mies\.|lat|lata|rok\S*)/iu;
  for (const [re, n] of NLQ_NUMBER_WORDS) s = s.replace(re, (m, _w, off, all) => (unit.test(all.slice(off + m.length)) ? n : m));
  return s.replace(/(?<!\p{L})(dłużej|dluzej|ponad|powyżej|powyzej|więcej|wiecej)\s*(niż|niz|od)?\s*(tydzień|tydzien|tygodnia|miesiąc|miesiac|miesiąca|miesiaca|rok|roku)(?!\p{L})/giu, (m, a, b, unit) => `${a} ${b || ""} 1 ${unit}`);
}
const NLQ_GROUP_RE = /\b(według|wedlug|wg\.?|na\s+(?:każd\w*\s+)?|per|po|dla\s+każd\w*|for\s+each|by|grupuj\w*\s+(?:po|wg|według)?)\s+([\p{L}][\p{L}\-]{2,})/giu;
const NLQ_GROUP_HINTS = { osob: "entity", pracownik: "entity", kto: "entity", person: "entity", employee: "entity", owner: "entity", nazwisk: "entity" };

function nlqFold(s) {
  return String(s ?? "").normalize("NFD").replace(/\p{M}/gu, "").replace(/ł/g, "l").replace(/Ł/g, "L").toLocaleLowerCase("pl");
}

// Wszystkie komórki (cały arkusz, nie tylko widok) w postaci do szukania — raz na zapytanie.
function nlqCorpus() {
  const out = [];
  const rows = typeof baseRows !== "undefined" ? baseRows : [];
  for (const row of rows) {
    for (let c = 0; c < currentHeaders.length; c++) {
      const v = typeof getDisplayValue === "function" ? getDisplayValue(row, c) : row.values?.[c];
      if (v != null && v !== "") out.push(nlqFold(v));
    }
  }
  return out.join("\n");
}

// Najdłuższy początek wyrazu obecny w danych (min. 4 znaki lub całe krótsze słowo) — odmiana
// polska bez słownika: „Kowalskiej” → „Kowalsk”, „Warszawie” → „Warszaw”.
function nlqMatchInData(word, corpus) {
  const w = String(word);
  const f = nlqFold(w);
  const min = Math.min(f.length, Math.max(4, f.length - 4));
  for (let len = f.length; len >= min; len--) {
    if (corpus.includes(f.slice(0, len))) return w.slice(0, len);
  }
  return null;
}

function nlqFindHeader(word) {
  const f = nlqFold(word);
  const stem = f.slice(0, Math.max(4, f.length - 3));
  const hint = Object.keys(NLQ_GROUP_HINTS).find((k) => f.startsWith(k));
  const headers = currentHeaders.map((h) => ({ h, f: nlqFold(h) }));
  const exact = headers.find((x) => x.f === f) || headers.find((x) => x.f.startsWith(stem)) || headers.find((x) => x.f.includes(stem));
  if (exact) return exact.h;
  if (hint && typeof getSmartModel === "function") {
    const model = getSmartModel();
    if (model && model.entityIdx >= 0 && model.blocks?.[0]) return currentHeaders[model.blocks[0].startIndex + model.entityIdx] || null;
  }
  return null;
}

// → { terms: [...], tokens: [{ token, label }], days, groupBy, unknown: [...] }
function parseNlQuery(text, { corpus = null } = {}) {
  let rest = ` ${nlqNormalizeNumbers(text || "")} `;
  const tokens = [];
  const addToken = (token, label) => { if (!tokens.some((x) => x.token === token)) tokens.push({ token, label }); };
  let days = null;
  const dm = rest.match(NLQ_DAYS_RE);
  if (dm) {
    const n = Number(dm[3].replace(",", "."));
    const unit = nlqFold(dm[4]);
    days = Math.round(n * (/^(tyg|tydz|week)/.test(unit) ? 7 : /^(mies|month)/.test(unit) ? 30 : /^(lat|rok|year)/.test(unit) ? 365 : 1));
    rest = rest.replace(dm[0], " ");
  }
  for (const rule of NLQ_RULES) {
    rest = rest.replace(rule.re, () => { addToken(rule.token, rule.label); return " "; });
  }
  if (days != null) addToken(`@dni>>${days}`, "nlqDays");
  let groupBy = null;
  rest = rest.replace(NLQ_GROUP_RE, (m, _p, word) => {
    const h = nlqFindHeader(word);
    if (h && !groupBy) { groupBy = h; return " "; }
    return m;
  });
  if (!groupBy && /\b(ile|policz|how\s+many|count)\b/iu.test(text) && /\b(osob\w*|pracownik\w*|kto|person|employee)\b/iu.test(rest)) {
    groupBy = nlqFindHeader("osoba");
  }
  const words = (rest.match(/[\p{L}][\p{L}\p{N}.\-']*|\d[\d.,\-/]*/gu) || [])
    .map((w) => w.replace(/[.,]+$/, ""))
    .filter((w) => w && !NLQ_STOP.has(nlqFold(w)) && !NLQ_STOP.has(w.toLocaleLowerCase("pl")) && !/^(osob|pracownik)/i.test(nlqFold(w)));
  const data = corpus ?? nlqCorpus();
  const terms = [];
  const unknown = [];
  for (const w of words) {
    const hit = nlqMatchInData(w, data);
    if (hit) { if (!terms.includes(hit)) terms.push(hit); } else unknown.push(w);
  }
  return { terms, tokens, days, groupBy, unknown };
}

// Zapytanie dla pola szukania (składnia aplikacji). Kilka słów → operatory „&&”.
function nlqToQuery(parsed) {
  const parts = [parsed.terms.length > 1 ? parsed.terms.join(" && ") : parsed.terms[0] || ""].filter(Boolean);
  parsed.tokens.forEach((x) => parts.push(x.token));
  return parts.join(" ").trim();
}

// ── AI (opcjonalnie) ─────────────────────────────────────────────────────────
const NLQ_TOKEN_LABEL = {
  "@wtoku": "nlqOpen", "@zamkniete": "nlqClosed", "@bezstartu": "nlqNoStart", "@niezaczete": "nlqNone",
  "@wszystkozamkniete": "nlqAllClosed", "@ostatni": "nlqLast", "@pierwszy": "nlqFirst",
};
const NLQ_ALLOWED = Object.keys(NLQ_TOKEN_LABEL);

function nlqAiSystem() {
  return "Zamieniasz pytanie o dane z arkusza na filtr aplikacji. Kolumny arkusza: " +
    currentHeaders.slice(0, 60).map((h) => `„${h}”`).join(", ") + ". " +
    "Dozwolone tryby: @wtoku (sprawa w trakcie), @zamkniete (zakończona), @bezstartu, @niezaczete, @wszystkozamkniete, " +
    "@ostatni (ostatni cykl), @pierwszy (pierwszy cykl). Odpowiedz WYŁĄCZNIE obiektem JSON: " +
    '{"szukaj": ["słowa do znalezienia w komórkach"], "tryby": ["..."], "dni_ponad": liczba albo null, "grupuj": "nazwa kolumny albo null"}';
}

async function nlqAskAi(text, corpus) {
  const who = currentHeaders.find((h) => /imi|nazwisk|osob|pracownik/i.test(h)) || currentHeaders[0] || "";
  const { value, problem } = await LocalAI.json({
    system: nlqAiSystem(),
    shots: [
      ["otwarte sprawy Wiśniewskiej dłużej niż tydzień", '{"szukaj": ["Wiśniewsk"], "tryby": ["@wtoku"], "dni_ponad": 7, "grupuj": null}'],
      ["ile zamkniętych na osobę", `{"szukaj": [], "tryby": ["@zamkniete"], "dni_ponad": null, "grupuj": ${JSON.stringify(who)}}`],
    ],
    prefix: '{"szukaj": ["',
    user: text,
    maxTokens: 100,
    validate: (j) => (Array.isArray(j.tryby) || j.tryby == null ? "" : "tryby-not-array"),
  });
  if (!value) return { error: problem || "invalid" };
  // sito: tylko znane tryby, istniejąca kolumna, słowa obecne w danych
  const tokens = (value.tryby || []).map((x) => String(x).trim().toLowerCase()).filter((x) => NLQ_ALLOWED.includes(x)).map((token) => ({ token, label: NLQ_TOKEN_LABEL[token] }));
  const days = Number.isFinite(Number(value.dni_ponad)) && Number(value.dni_ponad) > 0 ? Math.round(Number(value.dni_ponad)) : null;
  if (days != null) tokens.push({ token: `@dni>>${days}`, label: "nlqDays" });
  const terms = [];
  const unknown = [];
  (Array.isArray(value.szukaj) ? value.szukaj : value.szukaj ? [value.szukaj] : []).forEach((w) => {
    String(w).split(/\s+/).filter(Boolean).forEach((part) => {
      const hit = nlqMatchInData(part, corpus);
      if (hit) { if (!terms.includes(hit)) terms.push(hit); } else unknown.push(part);
    });
  });
  const groupBy = value.grupuj && currentHeaders.includes(String(value.grupuj)) ? String(value.grupuj) : null;
  return { terms, tokens, days, groupBy, unknown, ai: true };
}

// ── UI w oknie szybkiego szukania ────────────────────────────────────────────
const nlqInputEl = document.getElementById("nlqInput");
const nlqBtnEl = document.getElementById("nlqBtn");
const nlqResultEl = document.getElementById("nlqResult");
let nlqLast = null;

function nlqChip(text, cls = "") {
  return Object.assign(document.createElement("span"), { className: `nlq-chip ${cls}`.trim(), textContent: text });
}

function renderNlq(parsed, sourceText) {
  nlqLast = { parsed, text: sourceText };
  nlqResultEl.replaceChildren();
  nlqResultEl.hidden = false;
  const hasState = typeof getSmartModel === "function" && !!getSmartModel()?.hasStateColumns;
  const row = document.createElement("div");
  row.className = "nlq-chips";
  row.append(Object.assign(document.createElement("span"), { className: "nlq-lead", textContent: t("nlqUnderstood") }));
  if (parsed.ai) row.append(Object.assign(document.createElement("span"), { className: "ai-badge", textContent: t("aiBadge") }));
  parsed.terms.forEach((w) => row.append(nlqChip(t("nlqSearch", { text: w }))));
  parsed.tokens.forEach((x) => row.append(nlqChip(x.label === "nlqDays" ? t("nlqDays", { n: parsed.days ?? x.token.replace(/\D+/g, "") }) : t(x.label), "is-mode")));
  if (parsed.groupBy) row.append(nlqChip(t("nlqGroup", { col: parsed.groupBy }), "is-group"));
  if (!parsed.terms.length && !parsed.tokens.length && !parsed.groupBy) row.append(nlqChip(t("nlqNothing"), "is-empty"));
  nlqResultEl.append(row);
  const notes = [];
  if (parsed.unknown.length) notes.push(t("nlqUnknown", { words: parsed.unknown.join(", ") }));
  if (parsed.tokens.some((x) => x.token !== "@ostatni" && x.token !== "@pierwszy" && x.token !== "@niezaczete") && !hasState) notes.push(t("nlqNoState"));
  notes.forEach((n) => nlqResultEl.append(Object.assign(document.createElement("p"), { className: "nlq-note", textContent: n })));
  const actions = document.createElement("div");
  actions.className = "nlq-actions";
  const apply = Object.assign(document.createElement("button"), { type: "button", className: "btn btn-sm", textContent: t("nlqApply") });
  apply.disabled = !parsed.terms.length && !parsed.tokens.length && !parsed.groupBy;
  apply.addEventListener("click", () => applyNlq(parsed));
  actions.append(apply);
  // „Zapytaj AI” tylko przy dokładniejszym modelu: lekki 0.8B nie umiał ani złożyć filtra, ani
  // przypisać słowa do trybu (pomiar docs/llm-bench/prompts.html, 2026-10-17: 0/3 i 2/8).
  const st = typeof LocalAI !== "undefined" ? LocalAI.status() : null;
  const weakModel = st && st.status === "ready" && st.backend === "webgpu" && st.loadedChoice === "qwen3.5-0.8b";
  if (parsed.unknown.length && !parsed.ai && typeof LocalAI !== "undefined" && !weakModel) {
    const ai = Object.assign(document.createElement("button"), { type: "button", className: "btn btn-sm ghost", textContent: t("nlqAskAi") });
    ai.addEventListener("click", () => runNlqAi(sourceText));
    actions.append(ai);
  }
  nlqResultEl.append(actions);
}

function applyNlq(parsed) {
  const query = nlqToQuery(parsed);
  const input = typeof quickSearchPopupInput !== "undefined" && quickSearchPopupInput ? quickSearchPopupInput : quickSearchEl;
  const ops = typeof quickSearchPopupOperatorsEl !== "undefined" ? quickSearchPopupOperatorsEl : null;
  // bez zdarzeń „input/change” — podgląd na żywo (z opóźnieniem) cofnąłby zatwierdzony filtr
  if (parsed.terms.length > 1 && ops) ops.checked = true;
  if (input) input.value = query;
  commitQuickSearch();
  if (parsed.groupBy) {
    Object.assign(aggregationWorkbenchState, { groupBy: parsed.groupBy, groupBy2: "", groupBy3: "", measures: ["count_rows"], aggregation: "count" });
    const panel = document.getElementById("panel-aggregation-workbench");
    if (panel) panel.open = true;
    if (typeof ensureAnalysisHeavy === "function") ensureAnalysisHeavy();
    if (typeof renderAggregationWorkbench === "function") renderAggregationWorkbench();
    toast(t("nlqGroupDone", { col: parsed.groupBy }), "info");
  }
}

async function runNlqAi(text) {
  if (!(await LocalAI.ensure({ reason: t("nlqAiReason") }))) return;
  nlqBtnEl.disabled = true;
  try {
    const got = await nlqAskAi(text, nlqCorpus());
    if (got.error) { nlqResultEl.append(Object.assign(document.createElement("p"), { className: "nlq-note", textContent: t("nlqAiFailed") })); return; }
    renderNlq(got, text);
  } catch (err) {
    nlqResultEl.append(Object.assign(document.createElement("p"), { className: "nlq-note", textContent: t("aiFailed", { msg: err.message || err }) }));
  } finally {
    nlqBtnEl.disabled = false;
  }
}

function runNlq() {
  const text = String(nlqInputEl?.value || "").trim();
  if (!text) { nlqInputEl?.focus(); return; }
  if (!currentHeaders.length) { toast(t("nlqNoSheet"), "info"); return; }
  renderNlq(parseNlQuery(text), text);
}

nlqBtnEl?.addEventListener("click", runNlq);
nlqInputEl?.addEventListener("keydown", (e) => {
  if (e.key === "Enter" && !e.isComposing) { e.preventDefault(); e.stopPropagation(); runNlq(); }
});

document.getElementById("nlqOpenBtn")?.addEventListener("click", () => {
  if (typeof openQuickSearchPopup === "function" && openQuickSearchPopup()) nlqInputEl?.focus();
});
