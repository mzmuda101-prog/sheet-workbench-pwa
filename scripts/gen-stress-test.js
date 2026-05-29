// Generator arkusza testowego dla Excel Workbench PWA — stress-test agregacji & Wide-to-Long
//
// Użycie:
//   npm install xlsx        (jednorazowo, w folderze projektu)
//   node scripts/gen-stress-test.js
//
// Generuje: scripts/stress-test-workbench.xlsx
const XLSX = require("xlsx");

// ── dane bazowe ──────────────────────────────────────────────────────────────

const OSOBY = [
  "Jan Kowalski",
  "jan kowalski",        // duplikat po normalizacji (case)
  "Jan  Kowalski",       // duplikat po normalizacji (podwójna spacja)
  "Anna Nowak",
  "Anna Nowak-Wiśniewska",
  "Łukasz Ąkowski",      // polskie znaki
  "Żaneta Źródłowska",   // polskie znaki ekstremalne
  "Piotr Śliwiński",
  "Katarzyna Błaszczyk",
  "Michał Ćwikła",
  "Nikodem Wrześniewski",
  "",                    // pusta osoba → "(puste)"
  null,                  // null → "(puste)"
  "A",                   // bardzo krótka wartość
  "Osoba z bardzo długą nazwą która przekracza normalną szerokość kolumny w tabeli excela",
];

const PROJEKTY = ["Alpha", "Beta", "Gamma", "Delta", "Epsilon", "Alpha", "Beta"]; // duplikaty celowe
const STATUSY  = ["Aktywny", "Zakończony", "Wstrzymany", "Planowany", "aktywny", "AKTYWNY"]; // case test
const MIESIACE = ["Styczeń","Luty","Marzec","Kwiecień","Maj","Czerwiec",
                  "Lipiec","Sierpień","Wrzesień","Październik","Listopad","Grudzień"];
const LATA     = [2022, 2023, 2024, 2025];

function rand(arr) { return arr[Math.floor(Math.random() * arr.length)]; }
function randInt(min, max) { return Math.floor(Math.random() * (max - min + 1)) + min; }
function randFloat(min, max) { return parseFloat((Math.random() * (max - min) + min).toFixed(2)); }
function excelDate(year, month, day) {
  // Excel serial date
  const d = new Date(year, month - 1, day);
  return Math.floor((d - new Date(1899, 11, 30)) / 86400000);
}

// ── arkusz 1: główna tabela (3 000 wierszy, bloki kwartalne Wide-to-Long) ───

const HDR1 = [
  "Osoba", "Projekt", "Status", "Miesiąc", "Rok", "Kategoria",
  // bloki kwartalne — detektor suffix-cycle: Kw1_X, Kw2_X, Kw3_X, Kw4_X
  "Kw1_Kwota", "Kw1_Sztuki", "Kw1_Uwagi",
  "Kw2_Kwota", "Kw2_Sztuki", "Kw2_Uwagi",
  "Kw3_Kwota", "Kw3_Sztuki", "Kw3_Uwagi",
  "Kw4_Kwota", "Kw4_Sztuki", "Kw4_Uwagi",
  "Suma_kontrolna",
];

const rows1 = [HDR1];
for (let i = 0; i < 3000; i++) {
  const osoba    = rand(OSOBY);
  const projekt  = rand(PROJEKTY);
  const status   = rand(STATUSY);
  const miesiac  = rand(MIESIACE);
  const rok      = rand(LATA);
  const kategoria = rand(["Sprzedaż", "Koszty", "Marża", "", "Inwestycje", "HR"]);

  // kwartalne bloki — niektóre celowo puste
  const makeBlok = () => {
    if (Math.random() < 0.12) return [null, null, null]; // całkowicie pusty blok
    const kwota  = Math.random() < 0.05 ? null : randFloat(-5000, 150000);
    const sztuki = Math.random() < 0.08 ? null : randInt(0, 9999);
    const uwagi  = Math.random() < 0.7  ? null : rand(["OK", "Do weryfikacji", "Błąd", "brak", ""]);
    return [kwota, sztuki, uwagi];
  };

  const b1 = makeBlok(), b2 = makeBlok(), b3 = makeBlok(), b4 = makeBlok();

  // suma kontrolna: suma kwot niezerowych — testuje liczenie przez agregację
  const kwoty  = [b1[0], b2[0], b3[0], b4[0]].filter(v => typeof v === "number");
  const suma   = kwoty.length ? parseFloat(kwoty.reduce((a, b) => a + b, 0).toFixed(2)) : null;

  rows1.push([osoba, projekt, status, miesiac, rok, kategoria,
    ...b1, ...b2, ...b3, ...b4, suma]);
}

// ── arkusz 2: tabela z merged headers (detektor merge-based) ───────────────

// Wiersz 1: nagłówki grup (merge ponad podkolumnami) — symulowany płasko
// Bo xlsx-js XLSX.utils.aoa_to_sheet nie obsługuje merge ręcznie,
// robimy to przez worksheet.merges

const HDR2_GRP = ["", "", "Okres A", "", "", "Okres B", "", "", "Okres C", ""];
const HDR2_COL = [
  "ID", "Typ",
  "OkrA_Start", "OkrA_Koniec", "OkrA_Wartość",
  "OkrB_Start", "OkrB_Koniec", "OkrB_Wartość",
  "OkrC_Start", "OkrC_Koniec", "OkrC_Wartość",
  "Suma",
];

const rows2 = [HDR2_GRP, HDR2_COL];
for (let i = 1; i <= 800; i++) {
  const typ = rand(["Typ A", "Typ B", "Typ C", "Typ A", "Typ B"]); // duplikaty
  const row = [i, typ];
  for (let blok = 0; blok < 3; blok++) {
    const rok = 2023 + blok;
    const od  = Math.random() < 0.1 ? null : excelDate(rok, randInt(1,6), 1);
    const do_ = Math.random() < 0.1 ? null : excelDate(rok, randInt(7,12), 28);
    const war = Math.random() < 0.15 ? null : randFloat(0, 50000);
    row.push(od, do_, war);
  }
  row.push(row.slice(4).filter(v => typeof v === "number").reduce((a,b) => a+b, 0) || null);
  rows2.push(row);
}

// ── arkusz 3: edge-cases agregacji ─────────────────────────────────────────

const HDR3 = ["Klucz_grupowania", "Wartość_liczbowa", "Wartość_tekstowa", "Data", "Flaga"];
const KLUCZE = [
  "Normalny",
  "  Spacja na początku",
  "Spacja na końcu  ",
  "Dwie  Spacje",
  "normalny",            // duplikat po normalize
  "NORMALNY",            // duplikat po normalize
  "Normalny ",           // trailing space
  "Znak / ukośnik",
  "Znak \\ backslash",
  'Cudzysłów "podwójny"',
  "Apostrof 'pojedynczy'",
  "Liczba 123",
  "Liczba 0",
  "0",                   // może być traktowane jako liczba
  "FALSE",
  "TRUE",
  "NULL",
  "#REF!",
  "=FORMUŁA",
  "Łańcuch z\nnową linią",
  "Emoji 🎉",
  "Emoji 🔥🔥🔥",
  "Bardzo długi klucz grupowania który nie zmieści się w normalnej szerokości pola i wymaga przewijania poziomego lub obcięcia",
  null,
  "",
  0,                     // liczba 0 jako klucz
  false,
  true,
];

const rows3 = [HDR3];
for (let i = 0; i < 500; i++) {
  const klucz  = rand(KLUCZE);
  const liczba = rand([
    randFloat(-1e9, 1e9),  // bardzo duże/małe
    randFloat(-0.001, 0.001), // bliskie zeru
    0,
    -0,
    Infinity,              // uwaga: xlsx serializuje jako błąd lub null
    null,
    randInt(1, 100),
  ]);
  const liczbaSafe = Number.isFinite(liczba) ? liczba : null;
  const tekst  = rand(["OK", "Błąd", "N/A", null, "", "0", "false", "123abc", "abc123"]);
  const data   = Math.random() < 0.3 ? null : excelDate(randInt(2020,2025), randInt(1,12), randInt(1,28));
  const flaga  = rand([true, false, null, 1, 0, "tak", "nie", "yes", "no"]);
  rows3.push([klucz == null ? null : String(klucz), liczbaSafe, tekst, data, flaga]);
}

// ── arkusz 4b: suffix trailing-digit cycles — Osoba, Osoba1, Osoba2, Osoba3 ─
// Testuje detectSuffixedCycleBlocks (klasyczny pattern: base+cyfra na końcu)

const HDR4B_OSOBY = ["Osoba", "Osoba1", "Osoba2", "Osoba3"];  // 3 cykle osoby
const HDR4B_KWOTY = ["Kwota", "Kwota1", "Kwota2", "Kwota3"];  // 3 cykle kwoty
const HDR4B_DATY  = ["Data", "Data1", "Data2", "Data3"];       // 3 cykle daty

const HDR4B = [
  "ID", "Projekt", "Status",
  ...HDR4B_OSOBY, ...HDR4B_KWOTY, ...HDR4B_DATY,
  "Uwagi",
];

const OSOBY_SIMPLE = [
  "Anna Nowak", "Piotr Wiśniewski", "Katarzyna Kowalska",
  "Michał Zając", "Agnieszka Dąbrowska", "Tomasz Lewandowski",
  null, "",
];

const rows4b = [HDR4B];
for (let i = 1; i <= 1200; i++) {
  const projekt = rand(PROJEKTY);
  const status  = rand(STATUSY);

  // 4 osoby — z okazjonalnymi "brudnymi" wartościami
  const dirtyOsoba = () => {
    const r = Math.random();
    if (r < 0.12) return null;
    if (r < 0.14) return 12345;          // liczba zamiast imienia
    if (r < 0.16) return "?";
    if (r < 0.17) return "N/A";
    if (r < 0.18) return "-";
    if (r < 0.19) return "Nieznany";
    return rand(OSOBY_SIMPLE);
  };
  const osoby = HDR4B_OSOBY.map(dirtyOsoba);

  // 4 kwoty — z okazjonalnymi "brudnymi" wartościami
  const dirtyKwota = () => {
    const r = Math.random();
    if (r < 0.08) return null;
    if (r < 0.10) return "brak";         // tekst zamiast liczby
    if (r < 0.11) return "?";
    if (r < 0.12) return "-";
    if (r < 0.13) return "1 234,56";    // liczba z separatorem jako tekst
    if (r < 0.14) return "100 zł";      // liczba z jednostką jako tekst
    return randFloat(100, 99999);
  };
  const kwoty = HDR4B_KWOTY.map(dirtyKwota);

  // 4 daty — z okazjonalnymi "brudnymi" wartościami zamiast daty
  const dirtyDate = () => {
    const r = Math.random();
    if (r < 0.20) return null;
    if (r < 0.23) return "?";
    if (r < 0.26) return "brak";
    if (r < 0.28) return "nd.";
    if (r < 0.30) return "TBD";
    if (r < 0.32) return randFloat(0, 9999);  // liczba zamiast daty
    if (r < 0.33) return "01.13.2024";        // zły format (miesiąc 13)
    if (r < 0.34) return "2024-99-01";        // dzień 99
    return excelDate(randInt(2022, 2025), randInt(1, 12), randInt(1, 28));
  };
  const daty = HDR4B_DATY.map(dirtyDate);

  const uwagi = rand(["OK", "Do weryfikacji", null, "", "Pilne"]);
  rows4b.push([i, projekt, status, ...osoby, ...kwoty, ...daty, uwagi]);
}

// ── arkusz 4: duże n — 8 000 wierszy, 3 kolumny ────────────────────────────

const HDR4 = ["Region", "Dział", "Kwota"];
const REGIONY = ["Północ","Południe","Wschód","Zachód","Centrum"];
const DZIALY  = ["Sprzedaż","Marketing","IT","HR","Finanse","Produkcja","Logistyka"];
const rows4 = [HDR4];
for (let i = 0; i < 8000; i++) {
  rows4.push([rand(REGIONY), rand(DZIALY), randFloat(-10000, 500000)]);
}

// ── budowa workbooka ─────────────────────────────────────────────────────────

const wb = XLSX.utils.book_new();

// Arkusz 1 — Wide-to-Long
const ws1 = XLSX.utils.aoa_to_sheet(rows1);
// Ustaw format dat (kolumny nie mają dat w ark1, ok)
XLSX.utils.book_append_sheet(wb, ws1, "WideToLong_Kwartaly");

// Arkusz 2 — Merged headers (bloki okresów)
const ws2 = XLSX.utils.aoa_to_sheet(rows2);
// Dodaj merges nad HDR2_GRP (Okres A: C1:E1, Okres B: F1:H1, Okres C: I1:K1 — 0-indexed)
ws2["!merges"] = [
  { s: { r: 0, c: 2 }, e: { r: 0, c: 4 } }, // Okres A
  { s: { r: 0, c: 5 }, e: { r: 0, c: 7 } }, // Okres B
  { s: { r: 0, c: 8 }, e: { r: 0, c: 10 } }, // Okres C
];
XLSX.utils.book_append_sheet(wb, ws2, "MergedHeaders_Okresy");

// Arkusz 3 — Edge cases
const ws3 = XLSX.utils.aoa_to_sheet(rows3);
XLSX.utils.book_append_sheet(wb, ws3, "EdgeCases_Klucze");

// Arkusz 4b — Suffix trailing-digit (Osoba/Osoba1/Osoba2/Osoba3)
const ws4b = XLSX.utils.aoa_to_sheet(rows4b);
XLSX.utils.book_append_sheet(wb, ws4b, "SuffixCycle_Osoba123");

// Arkusz 4 — Duże N
const ws4 = XLSX.utils.aoa_to_sheet(rows4);
XLSX.utils.book_append_sheet(wb, ws4, "DuzeN_8000wierszy");

// ── zapis ─────────────────────────────────────────────────────────────────────

const outPath = require("path").resolve(__dirname, "stress-test-workbench.xlsx");
XLSX.writeFile(wb, outPath);
console.log(`✓ Wygenerowano: ${outPath}`);
console.log(`  Arkusz 1 (WideToLong_Kwartaly):    ${rows1.length - 1} wierszy, ${HDR1.length} kolumn — mid-digit: Kw1_Kwota`);
console.log(`  Arkusz 2 (MergedHeaders_Okresy):   ${rows2.length - 2} wierszy, ${HDR2_COL.length} kolumn — merged headers`);
console.log(`  Arkusz 3 (EdgeCases_Klucze):       ${rows3.length - 1} wierszy, ${HDR3.length} kolumn — edge cases`);
console.log(`  Arkusz 4 (SuffixCycle_Osoba123):   ${rows4b.length - 1} wierszy, ${HDR4B.length} kolumn — trailing digit: Osoba/Osoba1/Osoba2`);
console.log(`  Arkusz 5 (DuzeN_8000wierszy):      ${rows4.length - 1} wierszy, ${HDR4.length} kolumn — wydajność`);
