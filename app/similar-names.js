// similar-names.js — „Podobne nazwiska”: scalanie wariantów tej samej osoby/wartości w grupowaniu.
//
// ROADMAP „Shared Words Grouping”: „Julian Olsz”, „Gr1 Julian Olsz”, „Olsz Julian”, „J. Olsz”
// to jedna osoba; „Julian Olasz” — osobna (może literówka, może inna osoba → tylko PROPOZYCJA).
// Bez AI, deterministycznie — ten sam plik daje zawsze te same grupy.
//
// Kroki: (1) słowa bez znaczników grup („Gr 1”, „gr.4”, „Z2”), tytułów (mgr, dr, inż., pan/pani)
// i cyfr; wielkość liter i polskie znaki bez znaczenia („Łukasz” = „Lukasz”), kolejność słów też.
// (2) ten sam zestaw słów = ta sama grupa. (3) inicjał / samo nazwisko dołącza do JEDYNEJ grupy,
// do której pasuje. (4) różnica jednej litery w dłuższym słowie → propozycja do zatwierdzenia.
// Nazwa grupy = najczęstszy wariant (remis: najdłuższy bez znaczników).

(function (root) {
  "use strict";

  const TITLES = new Set(["mgr", "dr", "inz", "inż", "prof", "hab", "lek", "pan", "pani", "p", "sp", "ks"]);
  const GROUP_MARKERS = /^(gr|grupa|zesp|zespol|zespół|z|team|dz|dzial|dział|sekcja|sek)$/i;

  function fold(s) {
    return String(s).normalize("NFD").replace(/\p{M}/gu, "").replace(/ł/g, "l").replace(/Ł/g, "L").toLocaleLowerCase("pl");
  }

  // { words: [pełne słowa], initials: [litery] } albo null (pusta / sama liczba — bez scalania)
  function nameParts(label) {
    const raw = String(label ?? "").trim();
    if (!raw) return null;
    const tokens = raw.split(/[\s,;/|()\[\]]+|(?<=\.)(?=\p{L})/u).filter(Boolean);
    const words = [];
    const initials = [];
    for (let i = 0; i < tokens.length; i++) {
      const tok = tokens[i];
      const next = tokens[i + 1] || "";
      const bare = fold(tok).replace(/[.\-_'’:]+$/g, "").replace(/^[.\-_'’]+/g, "");
      if (!bare) continue;
      if (/\d/.test(bare)) continue; // Gr1, 4, Z2
      // „Gr 1”, „gr. 4” — krótki znacznik przed liczbą
      if (GROUP_MARKERS.test(bare) && /^\d/.test(next)) { i++; continue; }
      if (TITLES.has(bare)) continue;
      if (/^\p{L}$/u.test(bare)) { initials.push(bare); continue; }
      if (bare.length < 2) continue;
      words.push(bare);
    }
    if (!words.length && !initials.length) return null;
    return { words, initials };
  }

  function keyOf(parts) {
    return parts ? [...new Set(parts.words)].sort().join(" ") : "";
  }

  // jedna litera różnicy (zamiana/brak/wstawka) — tylko dla słów ≥ 4 znaków
  function oneEditApart(a, b) {
    if (a === b) return false;
    if (Math.min(a.length, b.length) < 4 || Math.abs(a.length - b.length) > 1) return false;
    let i = 0;
    let j = 0;
    let edits = 0;
    while (i < a.length && j < b.length) {
      if (a[i] === b[j]) { i++; j++; continue; }
      if (++edits > 1) return false;
      if (a.length > b.length) i++;
      else if (b.length > a.length) j++;
      else { i++; j++; }
    }
    return edits + (a.length - i) + (b.length - j) === 1;
  }

  // items: [{ label, count }]; opts.extraPairs: [[a, b]] zatwierdzone ręcznie (literówki)
  function buildSimilarNameGroups(items, opts = {}) {
    const byKey = new Map(); // klucz → { members: Map(label→count), words:Set }
    const loose = []; // inicjały / niepełne — do dopasowania w kroku 3
    const passthrough = new Map();
    for (const { label, count } of items) {
      const trimmed = String(label ?? "").trim();
      const parts = nameParts(trimmed);
      if (!parts) { passthrough.set(trimmed, trimmed); continue; }
      if (parts.initials.length && parts.words.length) {
        loose.push({ label: trimmed, count, parts });
        continue;
      }
      const key = parts.words.length ? keyOf(parts) : `#${parts.initials.join("")}`;
      if (!byKey.has(key)) byKey.set(key, { key, members: new Map(), words: new Set(parts.words) });
      const g = byKey.get(key);
      g.members.set(trimmed, (g.members.get(trimmed) || 0) + count);
    }
    // (3a) „J. Olsz” → grupa z „olsz” i słowem na „j” — tylko gdy dokładnie jedna taka
    for (const it of loose) {
      const candidates = [...byKey.values()].filter((g) => g.words.size > it.parts.words.length
        && it.parts.words.every((w) => g.words.has(w))
        && it.parts.initials.every((ch) => [...g.words].some((w) => !it.parts.words.includes(w) && w.startsWith(ch))));
      const key = candidates.length === 1 ? candidates[0].key : `${keyOf(it.parts)}|${it.parts.initials.join("")}`;
      if (!byKey.has(key)) byKey.set(key, { key, members: new Map(), words: new Set(it.parts.words) });
      const g = byKey.get(key);
      g.members.set(it.label, (g.members.get(it.label) || 0) + it.count);
    }
    // (3b) samo nazwisko „Olsz” → jedyna grupa z tym słowem i co najmniej dwoma słowami
    for (const g of [...byKey.values()]) {
      if (g.words.size !== 1 || !byKey.has(g.key)) continue;
      const [w] = g.words;
      const targets = [...byKey.values()].filter((x) => x !== g && x.words.size >= 2 && x.words.has(w));
      if (targets.length !== 1) continue;
      g.members.forEach((c, l) => targets[0].members.set(l, (targets[0].members.get(l) || 0) + c));
      byKey.delete(g.key);
    }
    // ręcznie zatwierdzone pary (literówki) — łączymy całe grupy
    const groupOfLabel = () => {
      const m = new Map();
      byKey.forEach((g) => g.members.forEach((_, l) => m.set(l, g)));
      return m;
    };
    for (const [a, b] of opts.extraPairs || []) {
      const idx = groupOfLabel();
      const ga = idx.get(String(a).trim());
      const gb = idx.get(String(b).trim());
      if (!ga || !gb || ga === gb) continue;
      gb.members.forEach((c, l) => ga.members.set(l, (ga.members.get(l) || 0) + c));
      gb.words.forEach((w) => ga.words.add(w));
      byKey.delete(gb.key);
    }
    // (4) propozycje: grupy o tej samej liczbie słów różniące się jedną literą w jednym słowie
    const groups = [...byKey.values()];
    const suggestions = [];
    for (let i = 0; i < groups.length; i++) {
      for (let j = i + 1; j < groups.length; j++) {
        const A = [...groups[i].words].sort();
        const B = [...groups[j].words].sort();
        if (A.length !== B.length || !A.length) continue;
        const diff = A.filter((w) => !groups[j].words.has(w));
        const diffB = B.filter((w) => !groups[i].words.has(w));
        if (diff.length === 1 && diffB.length === 1 && oneEditApart(diff[0], diffB[0])) {
          suggestions.push({ a: canonical(groups[i]), b: canonical(groups[j]) });
        }
      }
    }
    function canonical(g) {
      return [...g.members.entries()].sort((x, y) => y[1] - x[1]
        || (/\d/.test(x[0]) - /\d/.test(y[0]))
        || y[0].length - x[0].length
        || x[0].localeCompare(y[0], "pl"))[0][0];
    }
    const map = new Map(passthrough);
    const clusters = [];
    for (const g of byKey.values()) {
      const name = canonical(g);
      g.members.forEach((_, l) => map.set(l, name));
      if (g.members.size > 1) {
        clusters.push({ canonical: name, members: [...g.members.entries()].map(([label, count]) => ({ label, count })).sort((a, b) => b.count - a.count) });
      }
    }
    clusters.sort((a, b) => b.members.length - a.members.length || a.canonical.localeCompare(b.canonical, "pl"));
    return { map, clusters, suggestions };
  }

  const api = { buildSimilarNameGroups, nameParts, oneEditApart, fold };
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  root.SimilarNames = api;
})(typeof window !== "undefined" ? window : globalThis);
