# Od zera — czy da się zrobić taki arkusz w aplikacji?

Pomysł Mateusza (2026-11-14): nie tylko otwierać i poprawiać pliki z bazy testów, ale umieć
każdy z nich **zrobić od zera** — docelowo na poziomie Excela. Wynik przeglądu bazy i lista braków.

**Jak powstał:** `node tools/od-zera/inventory-xlsx.js <pliki|foldery>` wypisuje, czego z Excela
używa każdy plik (formuły i funkcje, scalenia, formaty, formatowanie warunkowe, poprawność
danych, tabele, wykresy, wydruk…). Baza: `arkusze/` (kosztorys, obieg, zadania, porównanie),
`scripts/fixtures/anon-*` (zanonimizowane prawdziwe), plik z `test-fixtures` i dwa wzorcowe
(`openpyxl-features`, `excel-parity-styles`). Pliki `accuracy-suite` to testy dokładności formuł —
pominięte (sprawdzają liczenie, nie budowanie).

## Liczenie: zrobione

Wszystkie funkcje z prawdziwych plików silnik już liczy (m.in. SUMIFS, COUNTIFS, INDEX/MATCH,
DATEDIF, IFERROR, HLOOKUP, SUMPRODUCT, TEXT po polsku, UNIQUE, VSTACK, TOCOL, SUBTOTAL),
też formuły rozlewane, odwołania między arkuszami i kolumny obliczeniowe tabel.

## Co da się zrobić od zera dziś

Wpisywanie wartości i formuł, nowe wiersze (formuły i wygląd z wierszy wyżej), format komórki
(pogrubienie, kursywa, podkreślenie, przekreślenie, rozmiar, kolor, wypełnienie, wyrównanie,
zawijanie, format liczb, obramowanie), notatki komórek, nazwa kolumny, zablokowanie kolumn /
wiersza nagłówka (w widoku).

## Braki — w kolejności (ile z 11 plików tego używa)

| # | Brak | Plików | Uwagi |
|---|---|---|---|
| 0 | ~~**Nowy skoroszyt**~~ — zrobione 2026-11-14 | — | Pusty skoroszyt, dodaj / usuń / zmień nazwę arkusza; zapis sprawdzony prawdziwym Excelem (szczegóły niżej). |
| 1 | ~~**Szerokość kolumn i wysokość wierszy w pliku**~~ — zrobione 2026-11-15 | 9 | Przeciągnięcie w Edycji trafia do pliku; szczegóły niżej. |
| 2 | ~~**Formatuj jako tabelę**~~ — zrobione 2026-11-15 | 8 | Przycisk na wstążce formatu; szczegóły niżej. |
| 3 | **Scalanie komórek** | 7 | Scal / Scal i wyśrodkuj / Rozdziel. Zapis dziś zostawia `mergeCells` bez zmian. |
| 4 | **Poprawność danych: lista** | 7 | Dziś tylko odczyt z pliku (rozwijana lista działa). Brak „Dane → Poprawność danych → Lista” od zera. |
| 5 | **Nazwy zakresów** | 7 | Menedżer nazw: utwórz / zmień. |
| 6 | **Formatowanie warunkowe** | 5 | Dziś tylko odczyt i liczenie reguł. Od zera: „większe niż”, „duplikaty”, skala kolorów, pasek danych, reguła formułą. |
| 7 | **Zablokuj okienka w pliku** | 5 | Blokada działa w widoku, nie trafia do pliku. |
| 8 | **Wydruk** | 3 | Obszar wydruku, poziomo / dopasuj do strony, tytuły wydruku, nagłówek i stopka. |
| 9 | **Ukryj wiersze / kolumny, ochrona arkusza** | 3 / 2 | |
| 10 | **Wykresy** | 3 | Duży etap; pliki mają je w obiegu i kosztorysie. |
| 11 | **Tabele przestawne** | 2 | Aplikacja ma własne zestawienia (agregacja) — do decyzji, czy zapisywać je jako tabelę przestawną Excela. |

## Dom poleceń (gdzie co jest)

Każde polecenie ma jedno miejsce; telefon (do 480 px) — te same miejsca.

| Polecenie | Dom | Skróty / inne wejścia |
|---|---|---|
| Nowy skoroszyt | panel **Plik i arkusz**, przycisk pod wyborem pliku | ekran startowy: „Nowy skoroszyt” obok „Wybierz plik z dysku” |
| Nowy arkusz, Zmień nazwę arkusza…, Usuń arkusz… | lista pod nazwą arkusza w nagłówku (▾), na dole listy; działa na bieżącym arkuszu | — |
| Szerokość kolumny / wysokość wiersza | przeciągnięcie krawędzi litery kolumny / numeru wiersza (Edycja) | Cofnij / Ponów |
| Formatuj jako tabelę | wstążka formatu (Edycja), przycisk z siatką za Autosumowaniem | — (Ctrl+T / Ctrl+L zajmuje przeglądarka) |
| Nazwa kolumny w pustej siatce | klik w pusty nagłówek pustej kolumny (Edycja) | prawy klik / przytrzymanie nagłówka → „Zmień nazwę kolumny” |

### 0. Nowy skoroszyt — jak działa (2026-11-14)

- **Nowy skoroszyt**: `Zeszyt1.xlsx` z jednym arkuszem `Arkusz1` (jak polski Excel; części pliku jak z Excela
  Mac — Aptos Narrow 12, motyw Office 2023, bez autora). Otwiera się w Edycji. Pusty arkusz to siatka A–H
  z wierszem 1 na nazwy kolumn i pustymi wierszami pod spodem; pierwszy wpis trafia do wiersza 2.
  W Podglądzie pusty arkusz pokazuje komunikat. Nie trafia do „Ostatnio otwierane”; pierwsze „Zapisz” pyta o miejsce.
- **Nowy arkusz**: `ArkuszN` z najniższym wolnym N, wstawiony za bieżącym (Excel po ponownym otwarciu liczy tak samo;
  w jednej sesji Excel liczy dalej — po usunięciu Arkusz2 daje Arkusz3; tego nie naśladujemy).
- **Zmiana nazwy**: zasady Excela (1–31 znaków, bez `: \ / ? * [ ]`, bez apostrofu na brzegach, bez powtórki,
  nie „Historia”). Odwołania w formułach, nazwach zdefiniowanych, formatowaniu warunkowym, listach wyboru,
  wykresach i źródle tabel przestawnych idą za nową nazwą; apostrofy jak w Excelu (`'Moje dane'!A1`, `'2025'!A1`).
- **Usunięcie**: pytanie, gdy arkusz nie jest pusty; odwołania → `#ADR!` (w pliku `#REF!`, jak zapisuje Excel);
  znikają części tylko tego arkusza (tabele, notatki, rysunki), nazwy lokalne arkusza, łańcuch obliczeń.
  Ostatniego widocznego arkusza nie da się usunąć.
- **Jak**: poprawki trafiają do pliku w pamięci, operacja zmienia plik (`app/sheet-ops.js`), skoroszyt wczytuje się
  od nowa (`app/sheet-ops-ui.js`). Jak w Excelu, operacji na arkuszach nie da się cofnąć (historia Cofnij od nowa).
- **Testy**: `node scripts/sheet-ops-test.js` (plik, bez przeglądarki) i `scripts/sheet-ops-playwright.js`
  (aplikacja, też `ENGINE=webkit TOUCH=1`); `EXCEL_DIR=/private/tmp/swb-excel-check` + `scripts/excel-real-check.py`
  — prawdziwy Excel otwiera wyniki bez naprawiania.
- **Na później**: odwołania 3D (`Arkusz1:Arkusz3!A1`) — zmiana nazwy pierwszego z pary i usuwanie skrajnego arkusza; przenoszenie / kopiowanie
  arkusza i kolejność zakładek; zmiana nazwy podwójnym kliknięciem; filtry i sortowanie widoku po operacji na
  arkuszu wracają do domyślnych (plik wczytuje się od nowa).

### 1. Szerokości kolumn i wysokości wierszy — jak działa (2026-11-15)

- Było: przeciąganie zmieniało tylko widok. Jest: w **Edycji** przeciągnięta szerokość kolumny i wysokość wiersza
  idą do pliku (`app/layout-edit.js` → `app/xlsx-layout.js`), liczą się do „Zapisz (N)”, mają Cofnij / Ponów,
  zostają po przełączeniu arkusza i w szkicu. W **Podglądzie** przeciąganie dalej zmienia tylko widok.
- Przeliczenie jak Excel na Macu (sprawdzone): na ekranie znaki × 6 + 5 pkt, w pliku `width = px / 6`;
  wysokość w punktach = px w aplikacji. Excel otwiera wynik z tą samą szerokością co aplikacja (138 px → 138 pkt).
  Zakres `<col min max>` jest dzielony, styl kolumny i ukrycie zostają.
- Uwaga: wiersze bez wysokości w pliku aplikacja rysuje w swojej gęstości (~29 px), Excel w 15–16 pkt — wiersz
  przeciągnięty z 29 na 54 px będzie w Excelu miał 54 pkt.
- Test: `scripts/layout-edit-playwright.js` (Chromium + Safari), Excel: `EXCEL_DIR=…` + `excel-real-check.py`.
- **Na później**: dwuklik na krawędzi = dopasuj do zawartości (Excel), wpisanie szerokości liczbą
  („Szerokość kolumny…”), zmiana kilku zaznaczonych kolumn naraz.

### 2. Formatuj jako tabelę — jak działa (2026-11-15)

- Było: aplikacja czytała tabele Excela (paski, kolor nagłówka, kolumny obliczeniowe, odwołania Tabela[Kolumna],
  rośnięcie przy nowych wierszach), ale nie umiała ich tworzyć.
- Jest: przycisk na wstążce formatu (działa na aktywnej komórce, jak w Excelu). Zakres = wiersz nagłówka + wszystkie
  wiersze danych; kolumny zaznaczone (gdy zaznaczono kilka), inaczej wszystkie z treścią (puste z prawej odpadają).
- Jak Excel (sprawdzone): nazwa `Tabela1`, `Tabela2`…; nagłówek pusty → „Kolumna1”, powtórzony → „Nazwa2”,
  liczba / formuła → tekst; styl `TableStyleMedium2`, filtr na całej tabeli, filtr arkusza przechodzi do tabeli.
  Odmowa z komunikatem: zakres nachodzi na inną tabelę albo na scalone komórki.
- Plik: `swbSheetOps.addTable` (część tabeli, relacja, `<tableParts>`, typ zawartości); nazwy kolumn muszą być
  równe komórkom nagłówka — aplikacja wpisuje je przed utworzeniem (inaczej Excel „naprawia” plik).
- Jak operacje na arkuszach: plik wczytuje się od nowa. **Cofnij / Ponów (2026-11-20):** utworzenie tabeli to
  krok Cofnij „cały plik” — Ctrl+Z, przycisk Cofnij i „Cofnij” w komunikacie wracają do pliku sprzed tabeli (także
  sprzed wpisania nazw kolumn do nagłówka), Ctrl+Y ją przywraca. Ograniczenie: zmiany komórek sprzed utworzenia
  tabeli nie dają się cofnąć po niej (historia zaczyna się od nowa, jak przy innych operacjach na pliku).
- Komunikat (decyzja Mateusza 2026-11-20: zakres zostaje): „Utworzono tabelę „Tabela2” (H5:I25): nagłówek z wiersza 5
  i wszystkie wiersze danych w kolumnach H–I”. Gdy zaznaczenie (ciągłe wiersze danych, co najmniej dwa) jest inne
  niż ten zakres, drugi przycisk „Tylko zaznaczenie (H8:I13)” cofa tabelę i tworzy ją dokładnie na zaznaczeniu,
  z pierwszym wierszem jako nagłówkiem — jak Ctrl+T w Excelu. Test: `scripts/table-undo-playwright.js`.
- Na później: wybór stylu z galerii, wiersz sumy, „Moja tabela nie ma nagłówków”, konwersja tabeli z powrotem
  na zakres.
- Przy okazji: po usunięciu arkusza formuły z `#REF!` przeliczają się od razu (wcześniej pokazywały starą wartość).

## Jak to sprawdzać

Każda pozycja kończy się testem „zrobione w aplikacji = jak w Excelu”: arkusz zbudowany od zera,
zapisany i otwarty prawdziwym Excelem (wyrocznia jak przy TEKST()). Miara postępu: dla każdego
pliku z bazy lista „czego nie da się zrobić od zera” ma maleć do zera.
