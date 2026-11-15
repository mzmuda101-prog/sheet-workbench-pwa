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

// ── numery z zerami wiodącymi (kolumna „Nr.” = liczba 32 w formacie 000 → widać „032”) ──
const nr = [{ rowIndex0: 33, values: ["031"] }, { rowIndex0: 34, values: ["032"] }, { rowIndex0: 35, values: ["033"] }, { rowIndex0: 40, values: [32] }];
eq(C.findTargets(nr, 0, "032").map((x) => [x.rowIndex0, x.score]), [[34, 3], [40, 3]], "numer: „032” = tekst „032” i liczba 32");
eq(C.findTargets(nr, 0, "32").map((x) => x.rowIndex0), [34, 40], "numer: „32” znajduje „032” (bez zer wiodących)");
eq(C.findTargets(nr, 0, "03").map((x) => x.rowIndex0).slice(0, 3), [33, 34, 35], "numer: początek „03” — wszystkie 03x");
ok(C.isDuplicateRisk(C.findTargets(nr, 0, "032")), "numer: istniejący numer = ryzyko duplikatu");
eq(C.findTargets([{ rowIndex0: 1, values: ["3,20"] }], 0, "3.2").length, 0, "numer: „3.2” nie łapie tekstu „3,20” jako tego samego numeru (tylko liczby całkowite)");

// ── foldCache: te same wyniki, normalizacja raz na wartość ──
const cache = new Map();
eq(C.findTargets(rows, 0, "jan kowalski", { foldCache: cache }).map((x) => x.rowIndex0), [5, 9, 12], "foldCache: te same wyniki");
ok(cache.get("Jan Kowalski") && cache.get("Jan Kowalski").f === "jan kowalski", "foldCache: zapamiętana postać — " + JSON.stringify([...cache]));
const size = cache.size;
C.findTargets(rows, 0, "anna", { foldCache: cache });
eq(cache.size, size, "foldCache: drugie szukanie nie dokłada wartości");

// ── looksLikeQuery: kiedy pole celu to zapytanie z operatorami (jak szybkie szukanie) ──
const H = ["Osoba", "Data", "Status", "Ilość"];
ok(C.looksLikeQuery("Ewa Lis && Status:=wzięta", H), "zapytanie: &&");
ok(C.looksLikeQuery("Lis || Nowak", H), "zapytanie: ||");
ok(C.looksLikeQuery("!Ewa", H), "zapytanie: wykluczenie na początku");
ok(C.looksLikeQuery('"Ewa Lis"', H), "zapytanie: cudzysłów");
ok(C.looksLikeQuery("Ilość:>>2", H), "zapytanie: porównanie");
ok(C.looksLikeQuery("status: zdana", H), "zapytanie: nazwa kolumny z dwukropkiem (bez wielkości liter)");
ok(C.looksLikeQuery("ilosc:2", H), "zapytanie: nazwa kolumny bez polskich znaków");
ok(!C.looksLikeQuery("Jan Kowalski", H), "nie zapytanie: zwykłe nazwisko");
ok(!C.looksLikeQuery("Ewa Lis-Nowak", H), "nie zapytanie: myślnik w nazwisku");
ok(!C.looksLikeQuery("Godz 8:30", H), "nie zapytanie: dwukropek nie po nazwie kolumny");
ok(!C.looksLikeQuery("Wykrzyknik!", H), "nie zapytanie: ! na końcu");
ok(!C.looksLikeQuery("", H), "nie zapytanie: puste");
ok(!C.looksLikeQuery("032||", H), "nie zapytanie: wiszące || (pisanie w trakcie albo literówka)");
ok(!C.looksLikeQuery("Ewa Lis &&  ", H), "nie zapytanie: wiszące &&");
ok(C.looksLikeQuery("032 || 033", H), "zapytanie: || z obu stron");
eq(C.trimQuery("032||"), "032", "trimQuery: zdejmuje wiszący operator");
eq(C.trimQuery(" Ewa Lis && Status:=wzięta && "), "Ewa Lis && Status:=wzięta", "trimQuery: tylko końcówka");

// ── distinguishingCols ───────────────────────────────
eq(C.distinguishingCols(["Osoba", "Dział", "Tel", "Miasto"], [["a", "b", "c", "d"], ["X", "X", "X", "X"], ["1", "2", "3", "4"], ["K", "L", "M", "K"]], 0), [2, 3], "rozróżniające: unikalność ≥ 50%, bez klucza");

// rozróżniające: kolumna prawie pusta (29 z 563 w pliku Mateusza) nie opisuje wierszy, choć ma unikalne wartości
const many = (f) => Array.from({ length: 20 }, (_, i) => f(i));
eq(C.distinguishingCols(["Nr", "Imię", "od", "Imię2", "Długość"], [
  many((i) => String(i)), many((i) => `Osoba ${i % 7}`), many((i) => `0${i % 4}.03`), many((i) => (i < 2 ? `X${i}` : "")), many((i) => `${i % 5}m`),
], 0), [1, 4], "rozróżniające: wypełnione i najbardziej zróżnicowane, w kolejności kolumn; prawie pusta odpada");

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

// ── parseTargetList: kolejka numerów / nazwisk ──
const L = (t) => { const r = C.parseTargetList(t, H); return r ? r.items : null; };
eq(L("031, 032, 045-048"), ["031", "032", "045", "046", "047", "048"], "lista: przecinki + zakres z zerami wiodącymi");
eq(L("031 032 033"), ["031", "032", "033"], "lista: same liczby po spacji");
eq(L("031\n032\n\n033"), ["031", "032", "033"], "lista: nowe linie, puste pomijane");
eq(L("Jan Kowalski, Anna Nowak"), ["Jan Kowalski", "Anna Nowak"], "lista: nazwiska po przecinku");
eq(L("Jan Kowalski; Anna Nowak"), ["Jan Kowalski", "Anna Nowak"], "lista: średnik");
eq(L("8-10"), ["8", "9", "10"], "lista: sam zakres bez zer");
eq(L("Jan Kowalski"), null, "nie lista: jedno nazwisko ze spacją");
eq(L("032"), null, "nie lista: jeden numer");
eq(L("Lis-Nowak"), null, "nie lista: nazwisko z myślnikiem");
eq(L("Ewa Lis && Status:=wzięta"), null, "nie lista: zapytanie z operatorami");
eq(L("032 || 033"), null, "nie lista: zapytanie z ||");
eq(L("10-8"), null, "nie lista: zakres malejący");
eq(L("1-9999"), null, "nie lista: zakres ponad 500 pozycji");
const dl = C.parseTargetList("031, 032, 031", H);
eq([dl.items, dl.dupes], [["031", "032"], ["031"]], "lista: duplikaty usunięte i zgłoszone");
eq(L("031,"), null, "nie lista: jedna pozycja z przecinkiem (pisanie w trakcie)");

// ── resolveList: pozycje kolejki → wiersze (tylko dokładne trafienia) ──
const rl = [{ rowIndex0: 1, values: ["031"] }, { rowIndex0: 2, values: ["032"] }, { rowIndex0: 3, values: ["Ewa Lis"] }, { rowIndex0: 4, values: ["Ewa Lis"] }, { rowIndex0: 5, values: [32] }];
const res = C.resolveList(["031", "Ewa Lis", "099", "Ewa Lyss", "32"], rl, 0);
eq(res.map((x) => [x.text, x.status, x.rowIndex0 ?? null, x.options || null]), [
  ["031", "ok", 1, null],
  ["Ewa Lis", "ambiguous", null, [3, 4]],
  ["099", "missing", null, null],
  ["Ewa Lyss", "missing", null, null],
  ["32", "ambiguous", null, [2, 5]],
], "kolejka: ok / kilka wierszy / brak; bez literówek; numer = numer");
const res2 = C.resolveList(["031", "31"], rl, 0);
eq(res2.map((x) => x.status), ["ok", "same-row"], "kolejka: dwie pozycje na ten sam wiersz — druga oznaczona");

console.log(`entry-core: ${passes} ✓, ${fails} ✗`);
process.exit(fails ? 1 : 0);
