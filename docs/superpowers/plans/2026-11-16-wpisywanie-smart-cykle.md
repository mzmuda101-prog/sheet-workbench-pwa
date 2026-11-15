# Plan wdrożenia — Wpisywanie, tryb smart (cykle od/do)

Projekt: `docs/superpowers/specs/2026-11-16-wpisywanie-smart-cykle-design.md`. Każdy krok: najpierw test
(czerwony), potem kod (zielony), potem commit.

1. **Silnik wykrywania** — `detectCycleColumns`, `cyclesOfRow` w `entry-core.js`.
   Test `scripts/entry-cycles-test.js`: nagłówki z pliku RODO (8 × 4, „Długość” = formuła → skip),
   cykle 3–8 bez danych (role z cyklu 1 / z nazw), arkusz bez cykli → null, „Kwota1, Kwota2” bez dat → null,
   role z Ustawień nadpisują.
2. **Stan i akcje** — `cycleState`, `planCycleAction`. Testy na wzorcach z pomiaru: wzięta-nieoddana,
   oddana w 1, pusty, w cyklu 2, cykl 1 otwarty + 2 zajęty (many-open → check), do bez od (odd → check),
   wszystkie 8 zajęte (full), Wydaj przy otwartym (open-prev + trzy rozwiązania), nigdy nie pisze do zajętej.
3. **Zapis per wiersz w kolejce** — `planMany`/`commitMany` z `items`. Test w `entry-playwright.js`
   (różne kolumny w różnych wierszach, jedno Cofnij przywraca wszystko).
4. **Karta smart jednego wiersza** — panel akcji, okno „poprzedni otwarty”, podgląd z cyklem.
   Test `scripts/entry-smart-playwright.js` na kopii pliku RODO (arkusz 2025-2026, klucz Nr.).
5. **Kolejka z akcją** — wybór akcji, statusy per pozycja, „naraz” z jednym Cofnij. Ten sam test.
6. **Ustawienia → Tryb smart** — per arkusz, w pamięci. Ten sam test + przeładowanie strony.
7. **Sprawdzenie**: cały `npm test`, plik RODO zapisany z aplikacji → prawdziwy Excel
   (`/private/tmp/swb-excel-check`, `open -a`, round-trip), telefon 382 px (WebKit, dotyk, bez
   poziomego przewijania), i18n.
8. **Wydanie** (dopiero gdy wszystko zielone): wersja, PR, scalenie, sprawdzenie produkcji.
