#!/usr/bin/env node
// Testy silnika Wpisywania (app/entry-core.js) — bez przeglądarki.
const C = require("../app/entry-core.js");

let fails = 0, passes = 0;
function ok(cond, msg) { if (cond) passes++; else { fails++; console.log("  ✗", msg); } }
function eq(a, b, msg) { ok(JSON.stringify(a) === JSON.stringify(b), `${msg} — jest ${JSON.stringify(a)}, oczekiwano ${JSON.stringify(b)}`); }

// ── fold ─────────────────────────────────────────────
eq(C.fold("Łukasz Żółć"), "lukasz zolc", "fold: polskie znaki i wielkość liter");
eq(C.fold(null), "", "fold: null → pusty");

// ── pickKeyColumn ────────────────────────────────────
eq(C.pickKeyColumn(["Lp", "Osoba", "Dział"], [[1, 2, 3], ["Jan", "Ewa", "Olaf"], ["A", "A", "B"]]), 1, "klucz: kolumna tekstowa o największej unikalności, liczby odpadają");
eq(C.pickKeyColumn(["A", "B"], [["x", "y"], ["p", "q"]]), 0, "klucz: remis → bardziej z lewej");
eq(C.pickKeyColumn(["A"], [[1, 2]]), 0, "klucz: brak kolumny tekstowej → 0");

// ── findTargets ──────────────────────────────────────
const rows = [
  { rowIndex0: 5, values: ["Jan Kowalski"] },
  { rowIndex0: 9, values: ["Kowalski Jan"] },
  { rowIndex0: 12, values: ["Jan Kowalsky"] },
  { rowIndex0: 14, values: ["Anna Nowak"] },
  { rowIndex0: 3, values: ["Ewa Lis"] },
  { rowIndex0: 7, values: ["Ewa Lis"] },
  { rowIndex0: 20, values: [null] },
];
const r1 = C.findTargets(rows, 0, "jan kowalski");
eq(r1.map((x) => [x.rowIndex0, x.score, x.exact]), [[5, 3, true], [9, 2, false], [12, 1, false]], "cel: dokładny, ten sam zbiór słów, literówka");
const r2 = C.findTargets(rows, 0, "Kowalski").map((x) => x.rowIndex0);
ok(r2.includes(5) && r2.includes(9), "cel: samo nazwisko znajduje oba warianty — " + JSON.stringify(r2));
eq(C.findTargets(rows, 0, ""), [], "cel: puste zapytanie → brak");
eq(C.findTargets(rows, 0, "   "), [], "cel: same spacje → brak");
eq(C.findTargets(rows, 0, "jan kowalski", { fuzzy: false }).map((x) => x.rowIndex0), [5, 9], "cel: fuzzy wyłączone — bez literówek");
eq(C.findTargets(rows, 0, "ewa lis").map((x) => x.rowIndex0), [3, 7], "cel: dwa identyczne klucze — oba, po numerze wiersza");
eq(C.findTargets(rows, 0, "jan", { limit: 2 }).length, 2, "cel: limit");
eq(C.findTargets(rows, 0, "zzzz"), [], "cel: nic nie pasuje");
eq(C.findTargets(rows, 0, "an").map((x) => x.score), [0.5], "cel: prefiks słowa (Anna) → score 0.5");
ok(C.findTargets(rows, 0, "Łukasz").length === 0, "cel: nie wybucha na polskich znakach");

// ── foldCache: te same wyniki, normalizacja raz na wartość ──
const cache = new Map();
eq(C.findTargets(rows, 0, "jan kowalski", { foldCache: cache }).map((x) => x.rowIndex0), [5, 9, 12], "foldCache: te same wyniki");
ok(cache.get("Jan Kowalski") && cache.get("Jan Kowalski").f === "jan kowalski", "foldCache: zapamiętana postać — " + JSON.stringify([...cache]));
const size = cache.size;
C.findTargets(rows, 0, "anna", { foldCache: cache });
eq(cache.size, size, "foldCache: drugie szukanie nie dokłada wartości");

// ── distinguishingCols ───────────────────────────────
eq(C.distinguishingCols(["Osoba", "Dział", "Tel", "Miasto"], [["a", "b", "c", "d"], ["X", "X", "X", "X"], ["1", "2", "3", "4"], ["K", "L", "M", "K"]], 0), [2, 3], "rozróżniające: unikalność ≥ 50%, bez klucza");

// ── profileColumn ────────────────────────────────────
const d = new Date(2026, 0, 1);
eq(C.profileColumn([d, d, d, d, d, d, d, d, d, "x"]).type, "date", "profil: 90% dat → date");
eq(C.profileColumn([1, 2, 3, "a", "b"]).type, "mixed", "profil: mieszanka → mixed");
eq(C.profileColumn([1, 2, 3, 4, null, ""]).type, "number", "profil: puste pomijane");
const cities = [];
for (let i = 0; i < 21; i++) cities.push(["Kraków", "Gdańsk", "Lublin"][i % 3]);
const pc = C.profileColumn(cities);
ok(pc.dictionary === true && pc.known.length === 3, "profil: słownikowa kolumna — " + JSON.stringify(pc));
ok(C.profileColumn(["a", "b", "c", "d"]).dictionary === false, "profil: same unikalne → nie słownik");

// ── typoSuggest ──────────────────────────────────────
eq(C.typoSuggest("Krakuw", ["Kraków", "Gdańsk"]), "Kraków", "literówka: 1 litera od znanej");
eq(C.typoSuggest("kraków", ["Kraków"]), null, "literówka: ta sama po fold → brak");
eq(C.typoSuggest("Kot", ["Kat"]), null, "literówka: krótkie słowa pomijane");
eq(C.typoSuggest("Lubiin", ["Lublin", "Lubiim"]), null, "literówka: dwie kandydatki → brak (niejednoznaczne)");
eq(C.typoSuggest("", ["Kraków"]), null, "literówka: puste → brak");

// ── isDuplicateRisk ──────────────────────────────────
ok(C.isDuplicateRisk([{ score: 2 }]), "duplikat: score 2 → ryzyko");
ok(C.isDuplicateRisk([{ score: 1 }]), "duplikat: literówka → też ryzyko (spec: anty-cel dwóch wierszy tej samej osoby)");
ok(!C.isDuplicateRisk([{ score: 0.5 }]), "duplikat: sam prefiks → bez ryzyka");
ok(!C.isDuplicateRisk([]), "duplikat: pusto → bez ryzyka");

console.log(`entry-core: ${passes} ✓, ${fails} ✗`);
process.exit(fails ? 1 : 0);
