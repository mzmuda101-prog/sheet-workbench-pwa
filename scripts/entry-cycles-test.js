#!/usr/bin/env node
// Testy trybu smart Wpisywania: wykrywanie cykli od/do i decyzje Wydaj/Zdaj (app/entry-core.js), bez przeglądarki.
// Wzorce z pomiaru pliku „RODO Obieg terenów 2026 V5”, arkusz „2025-2026” (8 cykli × 4 kolumny).
const C = require("../app/entry-core.js");

let fails = 0, passes = 0;
function ok(cond, msg) { if (cond) passes++; else { fails++; console.log("  ✗", msg); } }
function eq(a, b, msg) { ok(JSON.stringify(a) === JSON.stringify(b), `${msg} — jest ${JSON.stringify(a)}, oczekiwano ${JSON.stringify(b)}`); }

// ── nagłówki i kolumny jak w pliku ───────────────────
const H = ["Nr."];
for (let k = 1; k <= 8; k++) { const s = k === 1 ? "" : String(k); H.push(`Imię i Nazwisko${s}`, `od${s}`, `do${s}`, `Długość${s}`); }
// typy z danych: w cyklu 1–2 są daty i tekst, cykle 3–8 puste (typ „mixed” = brak danych); Długość = formuła
const info = H.map((h, i) => {
  if (i === 0) return { type: "number", formula: false };
  const pos = (i - 1) % 4, k = Math.floor((i - 1) / 4) + 1;
  if (pos === 3) return { type: "mixed", formula: true };
  if (k > 2) return { type: "mixed", formula: false };
  return { type: pos === 0 ? "text" : "date", formula: false };
});

const src = C.detectCycleColumns(H, info);
ok(!!src, "wykrycie: plik RODO ma cykle");
eq(src && src.kind, "columns", "wykrycie: źródło w poziomie");
eq(src && src.count, 8, "wykrycie: 8 cykli");
eq(src && src.span, 4, "wykrycie: 4 kolumny w cyklu");
eq(src && src.first, 1, "wykrycie: cykl zaczyna się po Nr.");
eq(src && src.roles, ["data", "start", "end", "skip"], "wykrycie: role osoba/od/do/formuła");
eq(src && src.labels, ["Imię i Nazwisko", "od", "do", "Długość"], "wykrycie: nazwy bazowe pozycji");

// cykle bez dat w danych (pusty arkusz) → role z nazw
const infoEmpty = H.map((h, i) => ({ type: "mixed", formula: i > 0 && (i - 1) % 4 === 3 }));
eq(C.detectCycleColumns(H, infoEmpty)?.roles, ["data", "start", "end", "skip"], "wykrycie: bez danych role z nazw od/do");

// angielskie nazwy, numeracja od 1
const HE = ["Key", "Name1", "From1", "To1", "Name2", "From2", "To2", "Name3", "From3", "To3"];
const eSrc = C.detectCycleColumns(HE, HE.map(() => ({ type: "mixed", formula: false })));
eq(eSrc && [eSrc.first, eSrc.span, eSrc.count, ...eSrc.roles], [1, 3, 3, "data", "start", "end"], "wykrycie: From/To, numery od 1");

// arkusz bez cykli
eq(C.detectCycleColumns(["Osoba", "Data", "Status", "Ilość", "Razem"], [{ type: "text" }, { type: "date" }, { type: "text" }, { type: "number" }, { type: "number" }]), null, "wykrycie: zwykła tabela → null");
eq(C.detectCycleColumns(["Nr", "Kwota1", "Kwota2", "Kwota3"], [{}, { type: "number" }, { type: "number" }, { type: "number" }]), null, "wykrycie: Kwota1..3 bez dat → null");
eq(C.detectCycleColumns(["A", "Data", "Data2"], [{}, { type: "date" }, { type: "date" }]), null, "wykrycie: jedna data na cykl (bez końca) → null");

// role z Ustawień nadpisują zgadywanie
const forced = C.detectCycleColumns(H, info, { roles: ["data", "end", "start", "skip"] });
eq(forced && forced.roles, ["data", "end", "start", "skip"], "wykrycie: role z Ustawień");
eq(C.detectCycleColumns(H, info, { roles: ["data", "data", "end", "skip"] })?.roles, ["data", "start", "end", "skip"], "wykrycie: złe role z Ustawień (bez startu) → zgadywanie");

// ── cykle wiersza i stan ─────────────────────────────
const D = (d) => new Date(Date.UTC(2026, 0, d));
function row(cycles) { // cycles: [[osoba, od, do], …] → wartości wiersza
  const v = [32];
  for (let k = 0; k < 8; k++) { const c = cycles[k] || [null, null, null]; v.push(c[0], c[1], c[2], c[1] ? "1m" : ""); }
  return v;
}
const cyc = (values) => C.cyclesOfRow(values, src);
eq(cyc(row([["Ewa", D(4), null]]))[0], { n: 1, start: { col: 2, v: D(4) }, end: { col: 3, v: null }, data: [{ col: 1, v: "Ewa" }], skip: [4] }, "cykle wiersza: kształt cyklu 1");
eq(cyc(row([]))[1].end.col, 7, "cykle wiersza: do2 = kolumna 7");

const st = (cycles) => C.cycleState(cyc(row(cycles)));
eq(st([["Ewa", D(4), null]]).cycles.map((c) => c.status), ["open", "empty", "empty", "empty", "empty", "empty", "empty", "empty"], "stan: wzięta, nieoddana");
eq(st([["Ewa", D(4), D(9)]]).free, 1, "stan: oddana w 1 → wolny cykl 2 (indeks 1)");
eq(st([]).free, 0, "stan: bez cyklu → wolny cykl 1");
eq(st([["Ewa", D(4), D(9)], ["Jan", D(10), null]]).open, [1], "stan: otwarty cykl 2");
eq(st([["Ewa", D(4), null], ["Jan", D(10), null]]).issues, ["many-open"], "stan: dwa otwarte → do sprawdzenia");
eq(st([["Ewa", null, D(9)]]).issues, ["odd"], "stan: do bez od → do sprawdzenia");
eq(st([["Ewa", null, null]]).issues, ["odd"], "stan: osoba bez od → do sprawdzenia");
const all8 = Array.from({ length: 8 }, (_, k) => ["X", D(k + 1), D(k + 2)]);
eq(st(all8).free, -1, "stan: 8 zajętych → brak wolnego");
eq(st([null, ["Jan", D(10), D(11)]]).free, 2, "stan: luka w cyklu 1 — wolny PO ostatnim zajętym");

// ── akcje ────────────────────────────────────────────
const act = (cycles, action, opts) => C.planCycleAction(cyc(row(cycles)), action, { date: "16.11.2026", ...opts });
eq(act([["Ewa", D(4), null]], "close"), { status: "ok", cycle: 1, writes: { 3: "16.11.2026" } }, "Zdaj: otwarty cykl 1 → do");
eq(act([["Ewa", D(4), D(9)], ["Jan", D(10), null]], "close"), { status: "ok", cycle: 2, writes: { 7: "16.11.2026" } }, "Zdaj: otwarty cykl 2 → do2");
eq(act([["Ewa", D(4), D(9)]], "close").status, "not-out", "Zdaj: nic otwartego → nie jest wydana");
eq(act([], "close").status, "not-out", "Zdaj: pusty wiersz → nie jest wydana");
eq(act([["Ewa", D(4), null], ["Jan", D(10), null]], "close").status, "check", "Zdaj: dwa otwarte → do sprawdzenia");
eq(act([["Ewa", null, D(9)]], "close").status, "check", "Zdaj: do bez od → do sprawdzenia");

eq(act([], "issue", { data: { 0: "Anna Lis" } }), { status: "ok", cycle: 1, writes: { 1: "Anna Lis", 2: "16.11.2026" } }, "Wydaj: pusty → cykl 1 osoba + od");
eq(act([["Ewa", D(4), D(9)]], "issue", { data: { 0: "Anna Lis" } }), { status: "ok", cycle: 2, writes: { 5: "Anna Lis", 6: "16.11.2026" } }, "Wydaj: po oddanej → cykl 2");
eq(act([["Ewa", D(4), D(9)]], "issue", { data: {} }).writes, { 6: "16.11.2026" }, "Wydaj: bez osoby — sama data");
eq(act(all8, "issue", { data: { 0: "A" } }).status, "full", "Wydaj: brak wolnego cyklu → tylko ostrzeżenie");
eq(act([["Ewa", null, D(9)]], "issue").status, "check", "Wydaj: anomalia → do sprawdzenia");

const prev = act([["Ewa", D(4), null]], "issue", { data: { 0: "Anna" } });
eq([prev.status, prev.openCycle, prev.cycle], ["open-prev", 1, 2], "Wydaj przy otwartym: pytanie, nie automat");
ok(!prev.writes, "Wydaj przy otwartym: bez rozstrzygnięcia nic nie pisze");
eq(act([["Ewa", D(4), null]], "issue", { data: { 0: "Anna" }, resolve: "close-issue" }),
  { status: "ok", cycle: 2, closeCycle: 1, writes: { 3: "16.11.2026", 5: "Anna", 6: "16.11.2026" } }, "Wydaj: zamknij poprzedni i wydaj");
eq(act([["Ewa", D(4), null]], "issue", { data: { 0: "Anna" }, resolve: "issue-anyway" }),
  { status: "ok", cycle: 2, writes: { 5: "Anna", 6: "16.11.2026" } }, "Wydaj mimo to — poprzedni zostaje otwarty");
eq(act([["Ewa", D(4), null]], "issue", { resolve: "note" }), { status: "note", cycle: 1 }, "Tylko notatka — nic nie wydaje");
const prevFull = act(Array.from({ length: 8 }, (_, k) => ["X", D(k + 1), k === 7 ? null : D(k + 2)]), "issue", { data: { 0: "A" } });
eq([prevFull.status, prevFull.openCycle, prevFull.cycle], ["open-prev", 8, null], "Wydaj: otwarty ostatni, brak wolnego → pytanie bez cyklu docelowego");
eq(act(Array.from({ length: 8 }, (_, k) => ["X", D(k + 1), k === 7 ? null : D(k + 2)]), "issue", { resolve: "close-issue" }).status, "full", "zamknij i wydaj bez wolnego → full");
eq(act([["Ewa", D(4), null]], "close", { date: "" }).status, "need-date", "Zdaj bez daty → potrzebna data");

// nigdy nie pisze do zajętej komórki — także gdy w wolnym cyklu ktoś zostawił datę „do” (odd i tak blokuje)
const anyWrite = [act([], "issue", { data: { 0: "A" } }), act([["E", D(1), null]], "close"), act([["E", D(1), null]], "issue", { data: { 0: "A" }, resolve: "close-issue" })];
const filled = [row([]), row([["E", D(1), null]]), row([["E", D(1), null]])];
anyWrite.forEach((a, i) => ok(Object.keys(a.writes).every((c) => filled[i][c] == null || filled[i][c] === ""), `nigdy nie nadpisuje (przypadek ${i + 1})`));

// ── data dziś ────────────────────────────────────────
eq(C.todayText(new Date(2026, 10, 6)), "06.11.2026", "data dziś: dd.mm.rrrr");

console.log(`entry-cycles: ${passes} ok, ${fails} błędów`);
process.exit(fails ? 1 : 0);
