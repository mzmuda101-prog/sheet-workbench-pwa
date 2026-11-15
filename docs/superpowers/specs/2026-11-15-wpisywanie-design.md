# Wpisywanie — projekt modułu (2026-11-15)

## Po co

Odwrotność Spisywania (`app/transcribe.js`, ekran → papier). Mateusz przepisuje dane
z **fizycznych kartek i kart pracy** do arkusza. Dwa wymagania, oba twarde:

1. **Nic się nie wysypie** — plik po zapisie otwiera się w Excelu bez naprawy, formuły,
   tabele Excela, CF i listy wyboru działają dalej.
2. **Wpis trafia tam, gdzie trzeba** — kartka dotycząca istniejącego rekordu uzupełnia
   ten rekord; kartka z nowym rekordem dopisuje nowy wiersz. Nigdy po cichu w zły wiersz.

## Cele / nie-cele / anty-cele

- **Cele:** karta-formularz jednej kartki; wybór celu (istniejący wiersz albo nowy);
  podpowiedzi i sprawdzanie wartości; podgląd zmian przed zatwierdzeniem; Cofnij całej kartki.
- **Nie-cele:** OCR / zdjęcie kartki, dyktowanie, tworzenie nowych kolumn z karty,
  wpisywanie do widoku Wide-to-Long, wiele arkuszy na jednej karcie.
- **Anty-cele (porażka, nawet jeśli „działa”):** nadpisanie niepustej komórki bez wyraźnej
  zgody; wpis w kolumnę z formułą / kolumnę wyliczaną; nowy wiersz pod ukrytymi filtrem
  danymi; zgubiona niezatwierdzona kartka po zgaszeniu tabletu; dwa wiersze tej samej osoby,
  bo literówka w nazwisku nie znalazła istniejącego.

## Podejścia (rozważone)

- **A. Karta z wyborem celu — WYBRANE (Mateusz, 2026-11-15).** Najpierw „kogo/czego
  dotyczy”, potem pola, potem podgląd zmian. Obsługuje uzupełnianie i dopisywanie.
- B. Karta tylko do nowych wierszy — prostsza, ale nie uzupełnia istniejących rekordów.
- C. Ulepszone wpisywanie w tabeli — przy stosie kartek łatwo zgubić wiersz.

## Ograniczenia (z kodu)

- Jedyna droga zapisu komórki: `writeCellFromInput(model, row, col, raw)` (`app/table.js`)
  → `parseCellInput` → `updateSheetCell` / `swbFormulaCommit`. Ta sama co wklejanie —
  daje parsowanie PL (daty, liczby, %), zachowanie stylu, `pendingEdits`, Cofnij, row-grow.
- Nowy wiersz: `swbRowGrow.addDraft()` — wymaga trybu `wide`, trybu Edycja i **braku filtra**
  (`swbRowGrow.filtered()`); styl i formuły kolumn z wiersza wyżej dopisuje sam przy 1. zapisie.
- Wszystkie zapisy w jednym zadaniu JS = **jeden krok Cofnij** (`undo.js`); część automatyczna
  row-grow jest osobnym krokiem (1. Cofnij zdejmuje automat, 2. wpis) — tak zostaje.
- Bramka trybu: `swbEditGuard()` / `swbIsEditMode()`.
- Kolumny wyliczane: `cfmtIsRealCol(col) === false` — nie do zapisu.
- Listy wyboru: `getCellDataValidation(sheet, r, absCol)`, egzekwowanie jak w edytorze
  komórki (stop / warning / info).
- Podpowiedzi: `buildColumnSuggestions(col)`, `createCellSuggestions(input, values, onPick)`,
  `rememberColumnValue`.
- Utrwalenie zatwierdzonych zmian po awarii robi już `drafts.js` (IndexedDB) — moduł
  utrwala tylko **niezatwierdzoną kartkę**.

## Jak działa (decyzje)

### Wejście
Przycisk „Wpisywanie” obok „Spisywanie” (to samo menu). Pełnoekranowa nakładka
(`#entryOverlay`), ten sam układ i style co Spisywanie (karta, duża czcionka, A/A/A,
Wake Lock). Otwarcie odmawia z wyjaśnieniem, gdy: tryb Podgląd, widok Long, brak pliku.
Gdy aktywny filtr/szukanie: okno „Wpisywanie widzi cały arkusz; nowe wiersze wymagają
wyczyszczenia filtra. [Wyczyść filtr i otwórz] [Anuluj]”. *(moja decyzja)*

### Krok 1 — Cel kartki
- Pole **„Kogo/czego dotyczy?”** szuka w **kolumnie-kluczu**. Klucz wybierany
  automatycznie: kolumna tekstowa z największym udziałem unikalnych niepustych wartości
  (remis → bardziej z lewej); użytkownik zmienia w „Ustawieniach karty”; zapamiętany per
  plik+arkusz w localStorage (wzór: zakresy w transcribe.js). *(moja decyzja)*
- Szukanie po **wszystkich** wierszach danych (`baseRows`), bez wielkości liter i polskich
  znaków (fold jak w `similar-names.js`), kolejność słów bez znaczenia, tolerancja 1 litery
  w słowie ≥ 5 znaków. Do 6 kandydatów; każdy pokazuje klucz + 2 kolumny rozróżniające
  (pierwsze niepuste kolumny o wysokiej unikalności) + numer wiersza arkusza.
- Wynik: wybierasz wiersz **albo** „＋ Nowy wiersz: «wpisany tekst»”.
- **Strażnik duplikatu:** wybór „Nowy” przy kandydacie z wynikiem ≥ dokładne/1 literówka
  pyta: „W arkuszu jest już «Jan Kowalski» (wiersz 42). Uzupełnić go zamiast tworzyć nowy?”.

### Krok 2 — Pola
- Pola = kolumny w kolejności tabeli; użytkownik wybiera i układa jak w Spisywaniu
  („Pola” — kolejność rubryk na kartce może być inna niż na ekranie). Zapamiętane per zakres.
- **Kolumny z formułą** (komórka celu ma formułę albo kolumna jest kolumną obliczeniową /
  wyliczaną): pole tylko do odczytu z etykietą „liczy się samo”; po zatwierdzeniu pokazuje
  wynik.
- Istniejący wiersz: każde pole pokazuje obecną wartość. Puste pole — zwykłe wpisywanie.
  Niepuste — wartość jako tekst; edycja zaznacza pole jako **nadpisanie**.
- **Sprytne wpisywanie (per pole):**
  - podpowiedzi: lista wyboru z pliku (pierwszeństwo) albo `buildColumnSuggestions`;
  - **typ kolumny** = przeważający typ niepustych wartości (data / liczba / tekst, próg 80%);
    pod polem na żywo „zapisze się jako: 12 marca 2026” (z `parseCellInput`); wpis innego typu
    w kolumnie dat/liczb = ostrzeżenie żółte, nie blokada;
  - lista wyboru w trybie `stop` blokuje zatwierdzenie (czerwony komunikat przy polu),
    `warning` ostrzega;
  - **literówka wobec słownika kolumny** (wartość spoza znanych, a 1 litera od znanej, dla
    kolumn „słownikowych” — < 30% unikalnych): „Miałeś na myśli «Kraków»?” [Popraw];
  - **przypinanie** wartości (ikona pinezki lucide): przypięte pole przechodzi na
    następną kartkę (np. data obchodu, osoba wpisująca). Przypięcia per zakres.
- Klawiatura: Enter = następne pole, w ostatnim polu → krok 3; Shift+Enter wstecz;
  Ctrl/⌘+Enter = od razu krok 3. Na dotyku `enterkeyhint="next"`.

### Krok 3 — Podgląd i zatwierdzenie
- Lista zmian: „Wiersz 42 · Jan Kowalski” albo „Nowy wiersz 118”; każda pozycja:
  kolumna, było → będzie. Nadpisania na czerwono z checkboxem „Tak, nadpisz” (domyślnie
  **odznaczony**). Pola bez zmian pomijane.
- **Zatwierdź** (wiele zapisów w jednym zadaniu JS → jeden krok Cofnij):
  1. nowy wiersz: `swbRowGrow.addDraft()` i znalezienie obiektu wiersza;
  2. dla każdego pola `writeCellFromInput(currentDisplayModel, row, col, raw)`;
     wynik ≠ "ok" → przerwanie przed jakimkolwiek zapisem (walidacja całości PRZED zapisem);
  3. `renderActiveTable()`, `setDirtyState(true)`, `rememberColumnValue`;
  4. dymek „Wpisano do wiersza 42 [Cofnij] [Pokaż w tabeli]” i czysta karta (przypięte zostają).
- **Prawda w komunikacie:** dymek dopiero po sprawdzeniu, że wartości są w arkuszu
  (odczyt z `workbook.Sheets`).
- **Nowa kartka po zmianie arkusza pod spodem** (inna karta przeglądarki, Cofnij w tabeli):
  przed zapisem sprawdzamy, że wartości „było” nadal są w komórkach; inaczej podgląd
  odświeża się z komunikatem.

### Sesja, Cofnij, utrwalenie
- „Wpisane w tej sesji: N” — lista kartek (klucz, wiersz, godzina) z „Pokaż w tabeli”.
- „Cofnij ostatnią kartkę” = `undoLast()` (×2, gdy kartka utworzyła wiersz z automatem —
  liczymy kroki po `undo` stosie, nie na ślepo).
- Niezatwierdzona kartka: localStorage per zakres, zapis przy każdej zmianie pola (debounce
  300 ms) i na `visibilitychange`; po otwarciu „Masz niedokończoną kartkę (Jan Kowalski,
  3 pola) [Wróć do niej] [Odrzuć]”.
- Plik nie zapisuje się sam — zwykłe „Zapisz”; przy zamknięciu nakładki z niezapisanym
  plikiem przypomnienie w dymku (raz na sesję).

### Dostępność i wygląd
Tokeny i style Spisywania; bez emoji (ikony lucide); fokus widoczny; `aria-live` dla
komunikatów pól; każde pole z `<label>`; cele dotyku ≥ 44 px; `prefers-reduced-motion`.
PL + EN w `language.js` (`check:i18n`).

## Testy (definicja „gotowe”)
- `scripts/entry-playwright.js` (chromium + webkit): uzupełnienie istniejącego wiersza;
  nowy wiersz (styl + formuły kolumn z wiersza wyżej, tabela Excela rośnie); literówka
  w kluczu znajduje wiersz; strażnik duplikatu; nadpisanie wymaga zgody; lista `stop`
  blokuje; kolumna z formułą tylko do odczytu; Cofnij kartki (z nowym wierszem i bez);
  przywrócenie niedokończonej kartki po przeładowaniu; filtr → pytanie o wyczyszczenie;
  Podgląd → odmowa.
- Round-trip: zapis → ponowne otwarcie → wartości w tych samych komórkach; walidator
  OOXML; prawdziwy Excel (umowa testów z Excelem) — plik otwiera się bez naprawy, tabela
  zawiera nowy wiersz.
- Dopisanie do `test:serial`.

## Na przyszłość: skan karty (OCR) — NIE w tym etapie
Mateusz chce kiedyś skanować kartę pracy aparatem: OCR wypełnia kartę, on tylko wybiera
status („wzięta” / „zdana”) i zatwierdza. Dlatego karta ma wejście programowe
(`swbEntry.fill({ target, fields })` — cel + wartości pól ze zwykłego obiektu), które
przechodzi przez te same kroki: cel, pola ze sprawdzaniem, podgląd, zatwierdzenie.
Ręczne wpisywanie korzysta z tej samej funkcji, więc OCR nie będzie osobną drogą zapisu.

## Zostawione budowniczemu
Nazwy plików (`app/entry.js`, `styles/…`), dokładny wygląd kroku 1, progi heurystyk
(80% / 30% / 1 litera) — do strojenia na plikach Mateusza (`scripts/*.xlsx`).
