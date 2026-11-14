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
| 0 | **Nowy skoroszyt** | — | Dziś nie ma „Nowy plik”: wszystko zaczyna się od otwarcia istniejącego. Podstawa „od zera”: pusty skoroszyt, dodaj / usuń / zmień nazwę arkusza. |
| 1 | **Szerokość kolumn i wysokość wierszy w pliku** | 9 | Przeciąganie zmienia widok, ale zapis nie dopisuje `<cols>` / wysokości. |
| 2 | **Formatuj jako tabelę** | 8 | Pliki mają tabele Excela (z nich biorą listy i kolumny obliczeniowe). Aplikacja je obsługuje, ale nie tworzy. |
| 3 | **Scalanie komórek** | 7 | Scal / Scal i wyśrodkuj / Rozdziel. Zapis dziś zostawia `mergeCells` bez zmian. |
| 4 | **Poprawność danych: lista** | 7 | Dziś tylko odczyt z pliku (rozwijana lista działa). Brak „Dane → Poprawność danych → Lista” od zera. |
| 5 | **Nazwy zakresów** | 7 | Menedżer nazw: utwórz / zmień. |
| 6 | **Formatowanie warunkowe** | 5 | Dziś tylko odczyt i liczenie reguł. Od zera: „większe niż”, „duplikaty”, skala kolorów, pasek danych, reguła formułą. |
| 7 | **Zablokuj okienka w pliku** | 5 | Blokada działa w widoku, nie trafia do pliku. |
| 8 | **Wydruk** | 3 | Obszar wydruku, poziomo / dopasuj do strony, tytuły wydruku, nagłówek i stopka. |
| 9 | **Ukryj wiersze / kolumny, ochrona arkusza** | 3 / 2 | |
| 10 | **Wykresy** | 3 | Duży etap; pliki mają je w obiegu i kosztorysie. |
| 11 | **Tabele przestawne** | 2 | Aplikacja ma własne zestawienia (agregacja) — do decyzji, czy zapisywać je jako tabelę przestawną Excela. |

## Jak to sprawdzać

Każda pozycja kończy się testem „zrobione w aplikacji = jak w Excelu”: arkusz zbudowany od zera,
zapisany i otwarty prawdziwym Excelem (wyrocznia jak przy TEKST()). Miara postępu: dla każdego
pliku z bazy lista „czego nie da się zrobić od zera” ma maleć do zera.
