// entry-core.js — silnik Wpisywania (kartka papieru → arkusz), bez DOM.
//
// Tu tylko decyzje, które da się sprawdzić bez przeglądarki: który wiersz jest celem kartki
// (szukanie po kolumnie-kluczu z tolerancją literówek), która kolumna jest kluczem, jakiego
// typu jest kolumna i czy wpisana wartość wygląda na literówkę znanej. Zapis i UI: entry.js.
//
// Ocena kandydata (score):
//   3   — ten sam tekst (bez wielkości liter i polskich znaków),
//   2   — te same słowa w innej kolejności („Kowalski Jan”),
//   1   — każde słowo zapytania = słowo wiersza albo literówka (1 znak, słowa ≥ 5 zn.)
//         — tu też „Kowalski” w „Jan Kowalski”,
//   0.5 — każde słowo zapytania to początek słowa wiersza („Kow”).
// Ryzyko duplikatu (pytanie przed nowym wierszem) od score 1 — literówka w nazwisku to
// dokładnie ten przypadek, w którym powstałyby dwa wiersze tej samej osoby.

const SwbEntryCore = (() => {
  const TYPE_SHARE = 0.8;      // przeważający typ kolumny
  const DICT_SHARE = 0.3;      // kolumna „słownikowa”: mniej niż 30% unikalnych
  const DICT_MIN = 5;          // za mało wartości, żeby mówić o słowniku
  const TYPO_MIN_LEN = 5;      // literówki tylko w dłuższych słowach
  const DISTINCT_SHARE = 0.5;  // kolumna rozróżniająca kandydatów

  function fold(s) {
    if (s == null) return "";
    return String(s).normalize("NFD").replace(/\p{M}/gu, "").replace(/ł/g, "l").replace(/Ł/g, "L")
      .toLocaleLowerCase("pl").replace(/\s+/g, " ").trim();
  }
  const words = (s) => fold(s).split(/[\s,;/|()\-]+/).filter(Boolean);
  const isEmpty = (v) => v === null || v === undefined || (typeof v === "string" && v.trim() === "");
  const asText = (v) => (v instanceof Date ? v.toISOString().slice(0, 10) : String(v));

  // Odległość Levenshteina ≤ 1 (bez pełnej macierzy: wystarczy jedno przejście).
  function within1(a, b) {
    if (a === b) return true;
    const la = a.length, lb = b.length;
    if (Math.abs(la - lb) > 1) return false;
    let i = 0, j = 0, diff = 0;
    while (i < la && j < lb) {
      if (a[i] === b[j]) { i++; j++; continue; }
      if (++diff > 1) return false;
      if (la > lb) i++; else if (lb > la) j++; else { i++; j++; }
    }
    return diff + (la - i) + (lb - j) <= 1;
  }
  const typoMatch = (q, w) => q.length >= TYPO_MIN_LEN && w.length >= TYPO_MIN_LEN && within1(q, w);

  function kindOf(v) {
    if (v instanceof Date) return "date";
    if (typeof v === "number") return "number";
    if (typeof v === "boolean") return "bool";
    const s = String(v).trim();
    return /^-?\d+([.,]\d+)?$/.test(s) ? "number" : "text";
  }

  function profileColumn(values) {
    const ne = (values || []).filter((v) => !isEmpty(v));
    const counts = {};
    ne.forEach((v) => { const k = kindOf(v); counts[k] = (counts[k] || 0) + 1; });
    let type = "mixed";
    for (const k of ["date", "number", "text"]) if (ne.length && (counts[k] || 0) / ne.length >= TYPE_SHARE) type = k;
    const seen = new Map();
    ne.forEach((v) => { const f = fold(asText(v)); if (!seen.has(f)) seen.set(f, asText(v)); });
    const dictionary = type === "text" && ne.length >= DICT_MIN && seen.size / ne.length < DICT_SHARE;
    return { type, dictionary, known: dictionary ? [...seen.values()] : [] };
  }

  const uniqueShare = (values) => {
    const ne = (values || []).filter((v) => !isEmpty(v));
    if (!ne.length) return 0;
    return new Set(ne.map((v) => fold(asText(v)))).size / ne.length;
  };

  function pickKeyColumn(headers, columns) {
    let best = -1, bestShare = -1;
    (columns || []).forEach((col, i) => {
      if (profileColumn(col).type !== "text") return;
      const share = uniqueShare(col);
      if (share > bestShare) { best = i; bestShare = share; }
    });
    return best < 0 ? 0 : best;
  }

  function distinguishingCols(headers, columns, keyCol, n = 2) {
    const out = [];
    for (let i = 0; i < (columns || []).length && out.length < n; i++) {
      if (i === keyCol) continue;
      if (uniqueShare(columns[i]) >= DISTINCT_SHARE) out.push(i);
    }
    return out;
  }

  // foldCache (opcjonalnie, Map tekst → { f, w }): normalizacja każdej wartości klucza raz,
  // a nie przy każdym naciśnięciu klawisza — przy dziesiątkach tysięcy wierszy to różnica w płynności
  function folded(value, cache) {
    const text = asText(value);
    if (!cache) return { f: fold(text), w: words(text) };
    let hit = cache.get(text);
    if (!hit) { hit = { f: fold(text), w: words(text) }; cache.set(text, hit); }
    return hit;
  }

  function scoreOf(qWords, qFold, value, fuzzy, cache) {
    if (isEmpty(value)) return 0;
    const { f: vFold, w: vWords } = folded(value, cache);
    if (vFold === qFold) return 3;
    if (vWords.length === qWords.length && [...vWords].sort().join(" ") === [...qWords].sort().join(" ")) return 2;
    // każde słowo zapytania do innego słowa wiersza
    const match = (test) => {
      const free = vWords.slice();
      return qWords.every((q) => {
        const k = free.findIndex((w) => test(q, w));
        if (k < 0) return false;
        free.splice(k, 1);
        return true;
      });
    };
    if (match((q, w) => q === w || (fuzzy && typoMatch(q, w)))) return 1;
    if (match((q, w) => q.length >= 2 && w.startsWith(q))) return 0.5;
    return 0;
  }

  function findTargets(rows, keyCol, query, opts = {}) {
    const limit = opts.limit || 6;
    const fuzzy = opts.fuzzy !== false;
    const qFold = fold(query);
    if (!qFold) return [];
    const qWords = words(query);
    const out = [];
    (rows || []).forEach((r) => {
      const v = r && Array.isArray(r.values) ? r.values[keyCol] : null;
      const score = scoreOf(qWords, qFold, v, fuzzy, opts.foldCache);
      if (score > 0) out.push({ rowIndex0: r.rowIndex0, score, exact: score === 3 });
    });
    out.sort((a, b) => b.score - a.score || a.rowIndex0 - b.rowIndex0);
    return out.slice(0, limit);
  }

  function typoSuggest(value, known) {
    const f = fold(value);
    if (f.length < TYPO_MIN_LEN) return null;
    const list = known || [];
    if (list.some((k) => fold(k) === f)) return null;
    const hits = list.filter((k) => { const kf = fold(k); return kf.length >= TYPO_MIN_LEN && within1(f, kf); });
    return hits.length === 1 ? hits[0] : null;
  }

  const isDuplicateRisk = (targets) => (targets || []).some((t) => t.score >= 1);

  return { fold, pickKeyColumn, findTargets, distinguishingCols, profileColumn, typoSuggest, isDuplicateRisk, within1 };
})();

if (typeof module !== "undefined" && module.exports) module.exports = SwbEntryCore;
