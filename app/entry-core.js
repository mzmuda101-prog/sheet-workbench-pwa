// @ts-check
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
  const FILL_SHARE = 0.5;     // kolumna rozróżniająca: wypełniona w co najmniej połowie wierszy

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

  // Kolumny opisujące kandydata obok klucza: WYPEŁNIONE (≥ połowa wierszy) i najbardziej
  // zróżnicowane. Sama unikalność wybierała kolumny prawie puste (29 z 563 wierszy w pliku
  // Mateusza) — przy każdym kandydacie „—”. Zwracane w kolejności kolumn arkusza.
  function distinguishingCols(headers, columns, keyCol, n = 2) {
    const scored = [];
    (columns || []).forEach((col, i) => {
      if (i === keyCol || !Array.isArray(col) || !col.length) return;
      const ne = col.filter((v) => !isEmpty(v));
      if (ne.length / col.length < FILL_SHARE) return;
      const distinct = new Set(ne.map((v) => fold(asText(v)))).size;
      if (distinct < 2) return;
      scored.push({ i, distinct });
    });
    return scored.sort((a, b) => b.distinct - a.distinct || a.i - b.i).slice(0, n).map((x) => x.i).sort((a, b) => a - b);
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
    // numer z kartki bez względu na zera wiodące: „032” = „32” = liczba 32 (kolumna „Nr.” w formacie 000)
    if (/^\d+$/.test(qFold) && /^\d+$/.test(vFold) && Number(qFold) === Number(vFold)) return 3;
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

  // Pole celu jako zapytanie w składni szybkiego szukania (&&, ||, !słowo, "fraza", >> <<,
  // Kolumna:…) — wtedy karta szuka jak tabela, a nie po samym kluczu z tolerancją literówek.
  // Zwykły tekst (nazwisko, „Lis-Nowak”, „8:30”) zostaje zwykłym szukaniem.
  // Wiszący operator na końcu („032||”, „Ewa Lis && ”) to pisanie w trakcie albo literówka —
  // pusta połówka nie może przełączać karty w zapytanie ani zmieniać wyników.
  const trimQuery = (q) => String(q ?? "").replace(/(\s*(&&|\|\|)\s*)+$/, "").trim();

  function looksLikeQuery(q, headers) {
    const s = trimQuery(q);
    if (!s) return false;
    if (/&&|\|\||>>|<<|"/.test(s)) return true;
    if (/^!\S/.test(s)) return true;
    const m = /^([^:]+):/.exec(s);
    if (!m) return false;
    const name = fold(m[1]);
    return (headers || []).some((h) => fold(h) === name);
  }

  // ── Kolejka: lista celów w jednym wpisie („031, 032, 045-048”, „Jan Kowalski; Anna Nowak”) ──
  // Separatory: przecinek, średnik, nowa linia; same liczby rozdzielone spacjami też; zakres liczb
  // całkowitych a-b (rosnąco, max 500, zera wiodące jak w pierwszej). Jedno nazwisko ze spacją,
  // nazwisko z myślnikiem i zapytanie z operatorami to NIE lista. null = zwykłe szukanie.
  const MAX_RANGE = 500;
  function expandToken(tok) {
    const m = /^(\d+)\s*-\s*(\d+)$/.exec(tok);
    if (!m) return [tok];
    const a = Number(m[1]), b = Number(m[2]);
    if (!(b >= a) || b - a + 1 > MAX_RANGE) return null;
    const width = m[1].length;
    const out = [];
    for (let n = a; n <= b; n++) out.push(String(n).padStart(width, "0"));
    return out;
  }
  function parseTargetList(text, headers, opts = {}) {
    let raw = String(text ?? "").trim();
    // „032||023||019” — same numery połączone „lub” to w karcie LISTA (zgłoszenie 2026-11-16),
    // bo człowiek wpisuje tak kilka kart; „Kowalski || Nowak” albo warunek kolumny zostają zapytaniem
    if (opts.orList !== false && /\|\|/.test(raw) && !/&&/.test(raw)) {
      const parts = raw.split("||").map((x) => x.trim()).filter(Boolean);
      if (parts.length >= 2 && parts.every((x) => /^\d+(\s*-\s*\d+)?$/.test(x))) raw = parts.join(", ");
    }
    if (!raw || looksLikeQuery(raw, headers)) return null;
    const glued = raw.replace(/(\d)\s*-\s*(\d)/g, "$1-$2"); // „045 - 048” → „045-048”
    let parts;
    if (/[,;\n]/.test(glued)) parts = glued.split(/[,;\n]+/).map((x) => x.trim()).filter(Boolean);
    else if (/^\d+(-\d+)?(\s+\d+(-\d+)?)*$/.test(glued)) parts = glued.split(/\s+/);
    else return null;
    // „031 032, 045” — liczby po spacji wewnątrz pozycji też rozdzielamy
    parts = parts.flatMap((p) => (/^\d+(-\d+)?(\s+\d+(-\d+)?)+$/.test(p) ? p.split(/\s+/) : [p]));
    const items = [];
    for (const p of parts) {
      const ex = expandToken(p);
      if (!ex) return null;
      items.push(...ex);
    }
    if (items.length < 2) return null;
    const seen = new Set();
    const uniq = [];
    const dupes = [];
    items.forEach((it) => { const k = fold(it); if (seen.has(k)) dupes.push(it); else { seen.add(k); uniq.push(it); } });
    return { items: uniq, dupes };
  }

  // Pozycje kolejki → wiersze. Tylko DOKŁADNE trafienia (lista ma być precyzyjna — literówka
  // w kolejce to „brak”, nie zgadywanie): ok (jeden), ambiguous (kilka), missing (żaden),
  // same-row (wiersz już zajęty przez wcześniejszą pozycję).
  function resolveList(items, rows, keyCol, opts = {}) {
    const used = new Set();
    return (items || []).map((text) => {
      const exact = findTargets(rows, keyCol, text, { fuzzy: false, limit: 50, foldCache: opts.foldCache }).filter((x) => x.score === 3);
      if (!exact.length) return { text, status: "missing" };
      if (exact.length > 1) return { text, status: "ambiguous", options: exact.map((x) => x.rowIndex0) };
      const r = exact[0].rowIndex0;
      if (used.has(r)) return { text, status: "same-row", rowIndex0: r };
      used.add(r);
      return { text, status: "ok", rowIndex0: r };
    });
  }

  const isDuplicateRisk = (targets) => (targets || []).some((t) => t.score >= 1);

  // ── Tryb smart: cykle od/do (projekt: docs/superpowers/specs/2026-11-16-wpisywanie-smart-cykle-design.md) ──
  // „Źródło cykli” zamienia wiersz na listę cykli { n, start, end, data, skip }; stan i akcje znają tylko
  // ten kształt — źródło „w pionie” (kolejne wiersze) da się dodać bez ruszania akcji i kolejki.
  const ROLE_START = /^(od|from|start|begin|wyd|pobr|wzi|out)/;
  const ROLE_END = /^(do|to|end|until|zwr|odd|zda|back|ret)/;
  const ROLES = ["data", "start", "end", "skip"];
  function splitHeader(h) {
    const s = String(h ?? "").trim();
    const m = /^(.*?)[\s_.-]*(\d+)$/.exec(s);
    return m && m[1] ? { base: m[1].trim(), n: Number(m[2]) } : { base: s, n: null };
  }
  function guessRoles(positions) { // positions[p] = { label, types: [...], formula }
    const roles = positions.map((p) => (p.formula ? "skip" : null));
    const dates = positions.map((p, i) => (roles[i] ? false : p.types.includes("date")));
    const dateIdx = dates.map((d, i) => (d ? i : -1)).filter((i) => i >= 0);
    if (dateIdx.length >= 2) { roles[dateIdx[0]] = "start"; roles[dateIdx[1]] = "end"; }
    else positions.forEach((p, i) => {
      if (roles[i]) return;
      const f = fold(p.label);
      if (!roles.includes("start") && ROLE_START.test(f)) roles[i] = "start";
      else if (!roles.includes("end") && ROLE_END.test(f)) roles[i] = "end";
    });
    return roles.map((r) => r || "data");
  }
  const validRoles = (r, span) => Array.isArray(r) && r.length === span && r.every((x) => ROLES.includes(x))
    && r.filter((x) => x === "start").length === 1 && r.filter((x) => x === "end").length === 1;

  // headers + colInfo[i] = { type (profileColumn), formula } → { kind:"columns", first, span, count, roles, labels } | null
  function detectCycleColumns(headers, colInfo, opts = {}) {
    const hs = (headers || []).map(splitHeader);
    let best = null;
    for (let first = 0; first < hs.length; first++) {
      for (let span = 2; first + span * 2 <= hs.length; span++) {
        const tpl = hs.slice(first, first + span);
        if (!tpl.every((x) => x.base) || !tpl.every((x) => x.n === null || x.n === 1)) continue;
        if (new Set(tpl.map((x) => fold(x.base))).size !== span) continue;
        let count = 1;
        while (first + (count + 1) * span <= hs.length
          && hs.slice(first + count * span, first + (count + 1) * span).every((x, p) => x.n === count + 1 && fold(x.base) === fold(tpl[p].base))) count++;
        if (count < 2) continue;
        if (!best || count * span > best.count * best.span) best = { first, span, count };
      }
    }
    if (!best) return null;
    const positions = Array.from({ length: best.span }, (_, p) => {
      const cols = Array.from({ length: best.count }, (_, k) => best.first + k * best.span + p);
      const info = cols.map((c) => (colInfo && colInfo[c]) || {});
      return { label: hs[best.first + p].base, types: info.map((x) => x.type), formula: info.some((x) => x.formula) };
    });
    const roles = validRoles(opts.roles, best.span) ? opts.roles.slice() : guessRoles(positions);
    if (!validRoles(roles, best.span)) return null;
    return { kind: "columns", ...best, roles, labels: positions.map((p) => p.label) };
  }

  function cyclesOfRow(values, src) {
    const out = [];
    if (!src) return out;
    for (let k = 0; k < src.count; k++) {
      const c = { n: k + 1, start: null, end: null, data: [], skip: [] };
      src.roles.forEach((role, p) => {
        const col = src.first + k * src.span + p;
        const cell = { col, v: values ? values[col] ?? null : null };
        if (role === "start") c.start = cell; else if (role === "end") c.end = cell;
        else if (role === "skip") c.skip.push(col); else c.data.push(cell);
      });
      out.push(c);
    }
    return out;
  }

  function cycleState(cycles) {
    const list = (cycles || []).map((c) => {
      const s = !isEmpty(c.start && c.start.v), e = !isEmpty(c.end && c.end.v), d = c.data.some((x) => !isEmpty(x.v));
      const status = s ? (e ? "closed" : "open") : (e || d ? "odd" : "empty");
      return { n: c.n, status };
    });
    const open = list.map((c, i) => (c.status === "open" ? i : -1)).filter((i) => i >= 0);
    let last = -1;
    list.forEach((c, i) => { if (c.status !== "empty") last = i; });
    const free = last + 1 < list.length ? last + 1 : -1;
    const issues = [];
    if (list.some((c) => c.status === "odd")) issues.push("odd");
    if (open.length > 1) issues.push("many-open");
    return { cycles: list, open, free, issues };
  }

  // action: "close" (Zdaj) | "issue" (Wydaj); opts: { date, data: {indeks pola danych: tekst}, resolve }
  // writes = { kolumna: tekst } — wyłącznie puste komórki wybranego cyklu.
  function planCycleAction(cycles, action, opts = {}) {
    const st = cycleState(cycles);
    const date = String(opts.date ?? "").trim();
    if (st.issues.length) return { status: "check", issues: st.issues };
    if (action === "close") {
      if (!st.open.length) return { status: "not-out" };
      const c = cycles[st.open[0]];
      if (!date) return { status: "need-date", cycle: c.n };
      return { status: "ok", cycle: c.n, writes: { [c.end.col]: date } };
    }
    const openC = st.open.length ? cycles[st.open[0]] : null;
    const target = st.free >= 0 ? cycles[st.free] : null;
    if (openC && !opts.resolve) return { status: "open-prev", openCycle: openC.n, cycle: target ? target.n : null };
    if (openC && opts.resolve === "note") return { status: "note", cycle: openC.n };
    if (!target) return { status: "full" };
    if (!date) return { status: "need-date", cycle: target.n };
    const writes = {};
    if (openC && opts.resolve === "close-issue") writes[openC.end.col] = date;
    target.data.forEach((cell, i) => { const v = opts.data && opts.data[i]; if (!isEmpty(v)) writes[cell.col] = String(v).trim(); });
    writes[target.start.col] = date;
    const out = { status: "ok", cycle: target.n };
    if (openC && opts.resolve === "close-issue") out.closeCycle = openC.n;
    out.writes = writes;
    return out;
  }

  const todayText = (d = new Date()) => `${String(d.getDate()).padStart(2, "0")}.${String(d.getMonth() + 1).padStart(2, "0")}.${d.getFullYear()}`;

  return { fold, pickKeyColumn, findTargets, distinguishingCols, profileColumn, typoSuggest, isDuplicateRisk, looksLikeQuery, trimQuery, parseTargetList, resolveList, within1,
    detectCycleColumns, cyclesOfRow, cycleState, planCycleAction, todayText };
})();

if (typeof module !== "undefined" && module.exports) module.exports = SwbEntryCore;
