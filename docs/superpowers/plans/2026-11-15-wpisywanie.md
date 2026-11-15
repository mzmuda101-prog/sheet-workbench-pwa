# Wpisywanie — plan wdrożenia

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Tryb „Wpisywanie”: jedna fizyczna kartka = jedna karta (cel → pola → podgląd), wpis do właściwego wiersza istniejącego albo nowego, bez psucia pliku.

**Architecture:** Czysty silnik bez DOM (`app/entry-core.js`, testowany w Node) + moduł aplikacji (`app/entry.js`: zapis przez istniejącą bramkę `writeCellFromInput`, nakładka UI, pamięć). Zapis całej kartki w jednym zadaniu JS → jeden krok Cofnij. Wejście programowe `swbEntry.fill()` pod przyszłe OCR.

**Tech Stack:** vanilla JS (skrypty `defer`, globalne funkcje jak w reszcie `app/`), SheetJS, Playwright (chromium + webkit), openpyxl do fixture, `scripts/ooxml-validate`.

**Spec:** `docs/superpowers/specs/2026-11-15-wpisywanie-design.md`

## Global Constraints

- Każdy zapis komórki wyłącznie przez `writeCellFromInput(currentDisplayModel, row, col, raw)` (`app/table.js`); nowy wiersz wyłącznie przez `swbRowGrow.addDraft()`.
- Tylko tryb Edycja (`swbIsEditMode()`), tylko widok `currentDisplayModel.mode === "wide"`.
- Nadpisanie niepustej komórki tylko z jawną zgodą przy tym polu (domyślnie odznaczona).
- Puste pole na karcie **nigdy nie czyści** komórki (moja decyzja — anty-cel „utrata danych”; czyszczenie robi się w tabeli).
- Kolumny z formułą / obliczeniowe / wyliczane (`!cfmtIsRealCol`) — tylko do odczytu.
- Walidacja CAŁEJ kartki przed pierwszym zapisem; błąd = zero zapisów.
- Bez emoji (ikony lucide inline SVG, `currentColor`); teksty PL + EN w `app/language.js`; `npm run check:i18n` zielone.
- Klucze localStorage: `excel-workbench-entry` (ustawienia i niedokończona kartka per zakres `plik\tarkusz`), max 30 zakresów jak `TR_MAX_SCOPES`.
- Progi heurystyk: typ kolumny ≥ 80% niepustych; kolumna „słownikowa” < 30% unikalnych; literówka = odległość 1 w słowie ≥ 5 znaków; do 6 kandydatów celu.
- Style z tokenów `styles/app.css`, cele dotyku ≥ 44 px, `prefers-reduced-motion`, widoczny fokus.
- **Opcje silnika (prośba Mateusza 2026-11-15):** każda „sprytna” funkcja włączana/wyłączana w panelu „Opcje” karty, zapamiętana globalnie (nie per arkusz) w `excel-workbench-entry.options`, domyślnie wszystkie włączone: `fuzzyTarget` (tolerancja literówek w szukaniu celu; wyłączone = tylko dokładne i ten sam zbiór słów), `typoHints` (podpowiedź „Miałeś na myśli…”), `typeHints` (ostrzeżenie o typie i „Zapisze się jako”), `suggestions` (lista podpowiedzi przy polu; listy wyboru z pliku zostają zawsze), `duplicateGuard` (pytanie przed nowym wierszem przy podobnym kluczu), `confirmStep` (krok 3 podglądu; wyłączone = zatwierdzenie od razu, ALE nadpisanie niepustej komórki i tak pokazuje podgląd). Silnik dostaje opcje parametrem, nie czyta UI.
- Nowe testy dopisane do `test:serial` w `package.json` (chromium i `ENGINE=webkit`).

## Review Focus

1. **Dwa wiersze z identycznym kluczem** („Jan Kowalski” ×2) — oba kandydaci widoczni z kolumnami rozróżniającymi, nic nie wybiera się samo. → test w Task 1 (`findTargets`) i Task 3.
2. **Tabela posortowana / przewinięta** — cel identyfikowany po `rowIndex0` arkusza, nigdy po pozycji w widoku. → test w Task 2 (sortowanie przed `fill`).
3. **Arkusz zmieniony między otwarciem karty a zatwierdzeniem** (Cofnij w tabeli, inna edycja) — commit sprawdza wartości „było”, przy różnicy odmawia z `reason: "stale"`. → test w Task 2.
4. **Wiersze nad nagłówkiem, podnagłówki, wiersze widoku long** — nie są kandydatami celu. → test w Task 1 (wiersze z `isLongViewRow` / `rowIndex0 < currentHeaderRow` odfiltrowane przez wywołującego, `findTargets` dostaje gotową listę) i Task 2.
5. **Częściowa data / liczba z przecinkiem** („12.03”, „1 234,5”) — podgląd pokazuje to samo, co trafi do komórki (ten sam `parseCellInput`). → test w Task 3.

---

### Task 1: Silnik bez DOM — `app/entry-core.js`

**Files:**
- Create: `app/entry-core.js`
- Test: `scripts/entry-core-test.js`

**Interfaces:**
- Produces (obiekt `SwbEntryCore`, `module.exports` jak w `anonymize-core.js`):
  - `fold(s: string) → string` — jak `fold` w `similar-names.js` (NFD, bez znaków diakrytycznych, ł→l, lower pl).
  - `pickKeyColumn(headers: string[], columns: any[][]) → number` — kolumna tekstowa (≥ 80% niepustych to tekst) o największym udziale unikalnych niepustych; remis → mniejszy indeks; brak → 0.
  - `findTargets(rows: {rowIndex0:number, values:any[]}[], keyCol: number, query: string, opts?: {limit?: number, fuzzy?: boolean /* domyślnie true; false = bez score 1 */}) → {rowIndex0:number, score:number, exact:boolean}[]` — score: 3 dokładne (po fold), 2 ten sam zbiór słów, 1 każde słowo zapytania = prefiks albo odległość 1 (słowa ≥ 5 zn.); malejąco, potem `rowIndex0`; limit domyślnie 6; pusty query → `[]`.
  - `distinguishingCols(headers, columns, keyCol, n=2) → number[]` — pierwsze kolumny ≠ keyCol z unikalnością ≥ 50%.
  - `profileColumn(values: any[]) → {type: "date"|"number"|"text"|"mixed", dictionary: boolean, known: string[]}` — `known` = unikalne niepuste jako tekst (gdy `dictionary`).
  - `typoSuggest(value: string, known: string[]) → string|null` — jedyna znana wartość w odległości 1 (słowo ≥ 5), inna niż value po fold; dokładne trafienie → null.
  - `isDuplicateRisk(targets) → boolean` — jest kandydat ze `score >= 2`.

- [ ] **Step 1: Write the failing test** `scripts/entry-core-test.js` (styl `anonymize-core-test.js`, `ok`/`eq`):
  - `eq(C.fold("Łukasz Żółć"), "lukasz zolc")`
  - `pickKeyColumn(["Lp","Osoba","Dział"], [[1,2,3],["Jan","Ewa","Olaf"],["A","A","B"]]) === 1` (liczby to nie klucz tekstowy)
  - `findTargets` na wierszach `Jan Kowalski`(5), `Kowalski Jan`(9), `Jan Kowalsky`(12), `Anna Nowak`(14): query `"jan kowalski"` → `[5 exact score3, 9 score2, 12 score1]`; query `"Kowalski"` → zawiera 5 i 9; query `""` → `[]`; `"jan kowalski"` z `{fuzzy:false}` → `[5, 9]` (bez 12); dwa wiersze `"Ewa Lis"` (3, 7) → oba, kolejność 3, 7.
  - `profileColumn(["2026-01-01" jako Date ×9, "x"]).type === "date"`; `profileColumn(["Kraków","Kraków","Gdańsk",…20 wierszy 3 miast]).dictionary === true`.
  - `typoSuggest("Krakuw", ["Kraków","Gdańsk"]) === "Kraków"`; `typoSuggest("Kraków", …) === null`; `typoSuggest("Kot", ["Kat"]) === null` (krótkie).
- [ ] **Step 2: Run** `node scripts/entry-core-test.js` — Expected: FAIL (`Cannot find module`).
- [ ] **Step 3: Implement** interfejsy z bloku wyżej w `app/entry-core.js` (IIFE `const SwbEntryCore = (() => {…})()`, odległość Levenshteina ograniczona do 1).
- [ ] **Step 4: Run** `node scripts/entry-core-test.js` — Expected: wszystkie ✓, kod wyjścia 0.
- [ ] **Step 5: Commit** `feat(sheet): Wpisywanie — silnik szukania celu i profilu kolumn`

### Task 2: Zapis kartki — `swbEntry.plan` / `swbEntry.commit` / `swbEntry.fill`

**Files:**
- Create: `app/entry.js` (część „zapis”; UI dochodzi w Task 3)
- Modify: `index.html` (skrypty `app/entry-core.js`, `app/entry.js` po `app/transcribe.js`, z `?v=` jak sąsiedzi)
- Test: `scripts/entry-playwright.js` (fixture openpyxl jak `calc-column-playwright.js`)

**Interfaces:**
- Consumes: `SwbEntryCore.*` (Task 1), `writeCellFromInput`, `swbRowGrow.addDraft/canAdd/filtered`, `parseCellInput`, `cfmtIsRealCol`, `getCellDataValidation`, `swbFormulaText`, `undoLast`, `renderActiveTable`, `setDirtyState`, `rememberColumnValue`.
- Produces (globalny `swbEntry`, test przez `window.__entry = swbEntry`):
  - `swbEntry.candidates() → {rowIndex0, values}[]` — wiersze danych `baseRows` (bez `isLongViewRow`, `rowIndex0 >= currentHeaderRow`).
  - `swbEntry.plan({target: {rowIndex0:number}|{newRow:true}, fields: {[col:number]: string}, overwrite?: number[]}) → {ok:boolean, changes: {col, before:string, after:string, parsedType:string, overwrite:boolean}[], errors: {col, reason:"formula"|"dv-stop"|"overwrite-not-allowed"|"not-real-col"}[], warnings: {col, reason:"dv-warning"|"type"|"typo", suggest?:string}[]}` — bez zapisu; puste wartości pomijane; wartość równa obecnej (po `parseCellInput`) pomijana.
  - `swbEntry.commit(planInput) → {ok:true, rowIndex0:number, undoSteps:1|2} | {ok:false, reason:"guard"|"filtered"|"invalid"|"stale", plan?}` — ponowny `plan()` tuż przed zapisem (wykrywa „stale”: `before` z planu ≠ obecne); wszystkie zapisy synchronicznie w jednym zadaniu.
  - `swbEntry.fill({target, fields})` — ustawia stan karty (Task 3 rysuje); w Task 2 tylko przechowuje i zwraca `plan()`.
  - `swbEntry.undoCard() → boolean` — `undoLast()` `undoSteps` razy dla ostatniej kartki.

Fixture `entry.xlsx` (openpyxl): arkusz „Prace”, tabela Excela `Prace` A1:E6: `Osoba` (tekst, 5 osób, w tym dwa razy „Ewa Lis”), `Data` (daty, format `dd.mm.yyyy`), `Status` (lista wyboru `"wzięta,zdana"` errorStyle stop), `Ilość` (liczby), `Razem` (kolumna obliczeniowa `=D2*2`).

- [ ] **Step 1: Write the failing test** — scenariusze (każdy `ok(nazwa, warunek, dostałem)`):
  1. uzupełnienie: `commit({target:{rowIndex0:2}, fields:{1:"12.03.2026", 2:"zdana"}})` → `ok`, `B3` data 2026-03-12, `C3 = "zdana"`, `undoSteps === 1`.
  2. nowy wiersz: `commit({target:{newRow:true}, fields:{0:"Olga Wrona",3:"4"}})` → `rowIndex0 === 6`, `E7` ma formułę `D7*2` i wynik 8, styl `A7` = styl `A6`; po zapisie (`buildOutputBytes`) tabela `Prace` ma ref `A1:E7`; walidator OOXML bez nowych błędów.
  3. nadpisanie: pole w niepustej komórce bez `overwrite` → `ok:false, reason:"invalid"`, `errors[0].reason === "overwrite-not-allowed"`, arkusz bez zmian; z `overwrite:[col]` → zapis.
  4. lista `stop`: `2:"oddana"` → `invalid`, `dv-stop`, **żadna** komórka nie zmieniona (także pole poprawne w tej samej kartce).
  5. kolumna obliczeniowa: `4:"99"` → `errors[0].reason === "formula"`.
  6. puste pole w niepustej komórce → brak zmiany w `changes`.
  7. Cofnij kartki: po scenariuszu 2 `undoCard()` → wiersz 7 pusty, tabela z powrotem `A1:E6` po zapisie.
  8. Review Focus 2: posortuj tabelę po `Osoba` malejąco, `commit` na `rowIndex0:2` → zmienia `A3`-wiersz arkusza, nie trzeci widoczny.
  9. Review Focus 3: `plan()`, potem zmiana `B3` przez `writeCellFromInput`, potem `commit` z tym samym wejściem i `overwrite` → `reason:"stale"`.
  10. filtr aktywny + `newRow` → `reason:"filtered"`; tryb Podgląd → `reason:"guard"`.
  11. brak `pageerror`.
- [ ] **Step 2: Run** `node scripts/entry-playwright.js` (serwer `npm run serve:test`) — Expected: FAIL (`swbEntry is not defined`).
- [ ] **Step 3: Implement** `app/entry.js` (część zapisu) zgodnie z blokiem Interfaces; DV jak w commicie edytora (`ui-controls.js` ~5408): porównanie bez wielkości liter, `allowBlank`.
- [ ] **Step 4: Run** `node scripts/entry-playwright.js && ENGINE=webkit node scripts/entry-playwright.js` — Expected: `OK` w obu.
- [ ] **Step 5: Commit** `feat(sheet): Wpisywanie — zapis kartki do istniejącego albo nowego wiersza`

### Task 3: Nakładka „Wpisywanie” — trzy kroki karty

**Files:**
- Modify: `index.html` (przycisk `#entryBtn` obok `#transcribeBtn`; nakładka `#entryOverlay role="dialog" aria-modal="true"`, wzór `#transcribeOverlay`)
- Modify: `app/entry.js` (UI), `styles/app.css` (sekcja `/* Wpisywanie */`, klasy `en-*`, tokeny jak `tr-*`), `app/language.js` (klucze `en*` PL+EN, `setText`/`setAttr` jak dla `trBtn`)
- Test: `scripts/entry-ui-playwright.js`

**Interfaces:**
- Consumes: `swbEntry.plan/commit/fill/candidates/undoCard` (Task 2), `SwbEntryCore.*`, `createCellSuggestions(input, values, onPick, opts)`, `buildColumnSuggestions(col)`.
- Produces: `swbEntry.open() → boolean`, `swbEntry.close()`, `swbEntry.state() → {open, step:1|2|3, target, fields, pinned:number[], keyCol, sessionLog:{key, rowIndex0, at}[]}`.

Zachowanie (copy dokładne):
- Otwarcie odmawia toastem: Podgląd → „Wpisywanie działa w trybie Edycja. Przełącz na Edycja i spróbuj ponownie.”; widok long → „Wpisywanie działa w zwykłym widoku tabeli. Wyłącz Wide-to-Long.”. Filtr → okno `appDialog`: „Wpisywanie widzi cały arkusz, ale nowe wiersze wymagają wyczyszczenia filtra.” [Wyczyść filtr i otwórz] [Anuluj].
- Krok 1: pole „Kogo/czego dotyczy?” (`#enTargetInput`), lista kandydatów (`role="listbox"`, strzałki + Enter), wiersz kandydata: klucz, 2 kolumny rozróżniające, „wiersz N”; ostatnia pozycja „Nowy wiersz: «tekst»”. Strażnik: `isDuplicateRisk` przy wyborze „Nowy” → okno „W arkuszu jest już «X» (wiersz N). Uzupełnić go zamiast tworzyć nowy?” [Uzupełnij wiersz N] [Utwórz nowy].
- Krok 2: `<label>` + `<input>` na pole; obecna wartość jako placeholder (pusta komórka) albo wartość + znacznik „nadpiszesz” po zmianie; pod polem `aria-live`: „Zapisze się jako: …” (z `parseCellInput` + `toDisplay`), ostrzeżenia typu/literówki z przyciskiem „Popraw na «X»”; pole formuły `readonly` z „Liczy się samo”; pinezka (`aria-pressed`). Enter = następne pole, Shift+Enter = poprzednie, Ctrl/⌘+Enter = krok 3; `enterkeyhint="next"`.
- Krok 3: lista „było → będzie”, nadpisania z checkboxem „Tak, nadpisz” (odznaczony); [Wpisz] nieaktywny przy błędach z komunikatem przy polu; po zapisie toast „Wpisano do wiersza N” [Cofnij] [Pokaż w tabeli]; karta czyści się, przypięte zostają, fokus w polu celu.
- Esc zamyka (z niezatwierdzonymi polami — kartka zostaje zapamiętana, Task 4); klawisze nie przeciekają do tabeli (capture, jak transcribe).

- [ ] **Step 1: Write the failing test** `scripts/entry-ui-playwright.js` (fixture z Task 2, wspólna funkcja `makeEntryFixture` wydzielona do `scripts/entry-fixture.js`):
  - klik `#entryBtn` → nakładka widoczna, fokus w `#enTargetInput`;
  - wpis „ewa lis” → 2 kandydaci z różnymi wartościami rozróżniającymi, żaden nie wybrany (Review Focus 1);
  - wpis „Jan Kowalsky” → pierwszy kandydat „Jan Kowalski”; wybór „Nowy” → okno strażnika;
  - krok 2: wpis „12.03” w Data → podgląd „Zapisze się jako:” równy `toDisplay(parseCellInput("12.03",…).value)`; po zapisie komórka pokazuje ten sam tekst (Review Focus 5);
  - „Status” = „zdana” z podpowiedzi; `Razem` ma `readonly`;
  - krok 3: nadpisanie bez zaznaczenia → [Wpisz] `disabled`; po zaznaczeniu → zapis, toast z „Cofnij”, klik Cofnij → wartość wraca;
  - pinezka na Data → po zapisie następna karta ma tę samą datę;
  - Esc zamyka, `Ctrl+Z` po zamknięciu nie zostało przechwycone w trakcie (klawisze nie przeciekły: tabela bez zmian podczas pisania w karcie);
  - Podgląd → toast z dokładnym tekstem; brak `pageerror`.
- [ ] **Step 2: Run** `node scripts/entry-ui-playwright.js` — Expected: FAIL (`#entryBtn` nie istnieje).
- [ ] **Step 3: Implement** HTML, CSS, i18n i UI w `app/entry.js` wg „Zachowanie”.
- [ ] **Step 4: Run** `node scripts/entry-ui-playwright.js && ENGINE=webkit node scripts/entry-ui-playwright.js && ENGINE=webkit TOUCH=1 node scripts/entry-ui-playwright.js && npm run check:i18n && node scripts/contrast-playwright.js` — Expected: wszystkie zielone.
- [ ] **Step 5: Commit** `feat(sheet): Wpisywanie — karta: cel, pola z podpowiedziami, podgląd zmian`

### Task 4: Pamięć — ustawienia zakresu, niedokończona kartka, dziennik sesji

**Files:**
- Modify: `app/entry.js`, `app/language.js`, `index.html` (panel „Ustawienia karty”: kolumna-klucz, kolejność i wybór pól — wzór `#trFieldsPanel`)
- Test: `scripts/entry-ui-playwright.js` (dopisane scenariusze)

**Interfaces:**
- Consumes: Task 3.
- Produces: localStorage `excel-workbench-entry` = `{scopes: {[plik\tarkusz]: {keyCol, fieldOrder:number[], selected:number[], pinned:{[col]:string}, draft: {target, fields, savedAt}|null, at}}}`; `swbEntry.state().sessionLog`.

- [ ] **Step 1: Write the failing test** (dopisane):
  - zmiana kolumny-klucza i kolejności pól → przeładowanie strony → te same ustawienia;
  - wpis 2 pól bez zatwierdzenia → `page.reload()` + wczytanie pliku → otwarcie karty pokazuje „Masz niedokończoną kartkę («Jan Kowalski», 2 pola)” [Wróć do niej] [Odrzuć]; „Wróć” przywraca cel i pola;
  - `visibilitychange` → `hidden` zapisuje kartkę od razu (bez czekania na debounce 300 ms);
  - 31 zakresów → najstarszy usunięty;
  - panel „Opcje”: wyłączenie `fuzzyTarget` → „Jan Kowalsky” nie znajduje „Jan Kowalski”; wyłączenie `typoHints` / `typeHints` → brak odpowiednich komunikatów; wyłączenie `duplicateGuard` → „Nowy” bez okna; wyłączenie `confirmStep` → Ctrl+Enter zapisuje od razu, ale kartka z nadpisaniem i tak pokazuje krok 3; opcje przeżywają przeładowanie i inny plik;
  - po 2 kartkach panel „Wpisane w tej sesji: 2” z „Pokaż w tabeli” ustawiającym fokus na `wide:<rowIndex0>`;
  - zamknięcie nakładki przy niezapisanym pliku → toast „Zmiany są w aplikacji, ale nie w pliku. Zapisz, kiedy skończysz.” — raz na sesję.
- [ ] **Step 2: Run** `node scripts/entry-ui-playwright.js` — Expected: FAIL na nowych scenariuszach.
- [ ] **Step 3: Implement** zapis/odczyt w `try/catch` (brak localStorage = działa bez pamięci).
- [ ] **Step 4: Run** `node scripts/entry-ui-playwright.js && ENGINE=webkit node scripts/entry-ui-playwright.js` — Expected: zielone.
- [ ] **Step 5: Commit** `feat(sheet): Wpisywanie — ustawienia per arkusz, niedokończona kartka, dziennik sesji`

### Task 5: Wpięcie w zestaw testów i sprawdzenie w prawdziwym Excelu

**Files:**
- Modify: `package.json` (`test:serial`: `node scripts/entry-core-test.js && node scripts/entry-playwright.js && ENGINE=webkit node scripts/entry-playwright.js && node scripts/entry-ui-playwright.js && ENGINE=webkit node scripts/entry-ui-playwright.js`; `test:excel`: `EXCEL_DIR=… node scripts/entry-playwright.js`)
- Modify: `scripts/entry-playwright.js` (przy `EXCEL_DIR` zapisuje `entry-nowy-wiersz.xlsx` i `entry-uzupelnienie.xlsx`), `scripts/excel-real-check.py` (oba pliki na liście)
- Modify: `sw.js` (nowe pliki w liście cache, jeśli lista jest jawna)

- [ ] **Step 1:** Dopisz kroki; `node scripts/pwa-offline-playwright.js` — Expected: zielone (nowe skrypty w cache offline).
- [ ] **Step 2:** `npm test -- entry` — Expected: wszystkie kroki `entry` zielone.
- [ ] **Step 3:** `npm run test:excel` wg umowy testów z Excelem (`/private/tmp/swb-excel-check`, `open -a`) — Expected: oba pliki otwierają się bez naprawy, tabela `Prace` zawiera nowy wiersz z formułą `Razem`.
- [ ] **Step 4:** `npm run test:all` — Expected: zielone (bez regresji Spisywania i row-grow).
- [ ] **Step 5: Commit** `test(sheet): Wpisywanie w zestawie testów i sprawdzeniu w Excelu`

Wydanie (bump wersji, release-PR) — osobno, na prośbę Mateusza.
