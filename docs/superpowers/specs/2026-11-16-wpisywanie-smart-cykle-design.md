# Wpisywanie — tryb smart: cykle od/do (PROJEKT, 2026-11-16)

Podstawa: szkic z decyzjami Mateusza (`2026-11-16-wpisywanie-smart-cykle-draft.md`). Tu: jak to zbudować
w istniejącym module (`app/entry-core.js`, `app/entry.js`), bez nowej drogi zapisu.

## 1. Silnik (entry-core.js, bez DOM — testy w Node)

**Źródło cykli** (decyzja 4: pion później, bez przepisywania akcji). Źródło zamienia wiersz na listę
cykli w jednym kształcie; stan i akcje znają tylko ten kształt:

```
cycle = { n, start: {col, v}, end: {col, v}, data: [{col, v}], skip: [col] }
```

- `detectCycleColumns(headers, colInfo)` — źródło „w poziomie”. `colInfo[i] = { type, formula }`
  (typ z `profileColumn`, formuła z komórki). Nazwa bazowa = nagłówek bez cyfry na końcu („od2” → „od”).
  Szuka bloku `span ≥ 2` powtórzonego `≥ 2` razy, gdzie blok k ma numer k+1 (blok 1 bez numeru albo „1”)
  — tak Excel numeruje powtórzone nagłówki tabeli. Najdłuższy blok wygrywa.
  Role pozycji w bloku (z danych WSZYSTKICH cykli na tej pozycji, bo cykle 3–8 bywają puste):
  formuła → `skip`; daty → pierwsza `start`, druga `end`; brak danych → podpowiedź z nazwy
  (od/from/start/wyd…, do/to/end/zwrot/odd…); reszta → `data` (osoba). Bez jednego `start` i jednego
  `end` → `null` (arkusz bez cykli, karta jak dziś). `roles` z Ustawień nadpisują zgadywanie.
- `cyclesOfRow(values, src)` → lista cykli wiersza.
- `cycleState(cycles)` → każdy cykl: `empty | open | closed | odd` (odd = `do` bez `od`, dane bez `od`).
  Wiersz: `open` (indeksy), `free` (pierwszy pusty cykl PO ostatnim zajętym), `issues`
  (`odd`, `many-open`).
- `planCycleAction(cycles, action, { date, data, resolve })` → `{ status, writes, cycle, closeCycle }`:
  - **close** (Zdaj): issues → `check`; brak otwartego → `not-out`; jeden otwarty → `ok`, `writes = {end: date}`.
  - **issue** (Wydaj): issues → `check`; otwarty poprzedni → `open-prev` (bez `resolve`) albo wg `resolve`:
    `close-issue` (zamknij datą + wydaj) · `issue-anyway` · `note`; brak wolnego → `full`; inaczej `ok`,
    `writes = {start: date, data…}`.
  - `writes` tylko do PUSTYCH komórek wybranego cyklu (smart nie nadpisuje). Zapis i tak przechodzi
    przez `plan()`, który blokuje nadpisanie bez zgody — druga linia obrony.

## 2. Zapis (entry.js, bez nowej drogi)

Akcja smart = wyliczone `fields` (+ ewentualnie notatka) → istniejące `plan` / `commit` (writeCellFromInput,
jedno Cofnij, „stale”). Kolejka: `planMany` / `commitMany` przyjmują pozycje z WŁASNYMI polami
(`items: [{ rowIndex0, fields, notes }]`) — każdy wiersz dostaje datę w swojej kolumnie, jeden zapis,
jedno Cofnij. Przed zatwierdzeniem decyzja smart liczona jest jeszcze raz na świeżych danych; inna niż
w podglądzie → „stale”, nic nie piszemy.

## 3. Interfejs

- **Karta jednego wiersza** (krok 2): przy aktywnym trybie zamiast pól — pasek stanu
  („Cykl 1 otwarty od 1 sty 26 · Adam Kamiński” / „Wolny: cykl 2”), przełącznik akcji
  **[Zdaj] [Wydaj]** (nazwy z Ustawień), pole daty (domyślnie dziś) i pola danych wydania (osoba,
  z podpowiedziami jak dziś). Link „Wszystkie pola” wraca do zwykłej karty. „Dalej” → istniejący
  podgląd „było → będzie” z nagłówkiem cyklu („→ do2 · cykl 2”).
- **Wydaj przy otwartym poprzednim** (decyzja 3): okno ze stanem i wyborem *Zamknij poprzedni (data)
  i wydaj · Wydaj mimo to · Tylko notatka · Anuluj*. „Tylko notatka” zapisuje notatkę w komórce klucza.
- **Zdaj bez otwartego cyklu**: komunikat, nic się nie zapisuje. **Brak wolnego cyklu** (decyzja 2): tylko
  ostrzeżenie. **Anomalie**: „do sprawdzenia”, bez zgadywania.
- **Kolejka**: na ekranie kolejki wybór akcji (*Pola · Zdaj · Wydaj*). Przy akcji każda pozycja pokazuje
  swoją kolumnę („→ do2”) albo status: „nie wydana” (pomiń), „do sprawdzenia” / „poprzedni otwarty” —
  wybór per pozycja (te same opcje co w oknie + Pomiń). „Wszystkim naraz” = jeden podgląd, jedno Cofnij.

## 4. Ustawienia → Tryb smart (per arkusz, ta sama pamięć co klucz i pola)

Włączony (domyślnie tak, gdy wykryto cykle) · nazwy akcji (domyślnie „Wydaj” / „Zdaj”) · słowo na
pozycję (domyślnie „Pozycja”; używane tylko w mianowniku: „Karta 032 · cykl 1 otwarty”) · data domyślna
(dziś / pusta) · anomalie w kolejce (pokaż do decyzji / pomiń) · wykryte role z ręczną poprawką.

## 5. Poza zakresem

Cykle w pionie (tylko interfejs źródła), dopisywanie kolumn cyklu 9, OCR.
