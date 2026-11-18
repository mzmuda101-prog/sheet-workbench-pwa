# Wstążka — etap 2: ekran Plik (plan wdrożenia)

**Cel:** „Plik” w rzędzie 1 otwiera ekran za kulisami (makieta v2 `.bs`), zamiast tymczasowego menu z etapu 1.
**Projekt:** `docs/WSTAZKA.md` (etap 2) + `docs/makiety/wstazka-v2.html` (bsHtml). Zasada: bez dubli.

## Zawartość (jedna funkcja = jedno miejsce)

| Sekcja | Co | Skąd (bez kopii logiki) |
|---|---|---|
| Wróć | zamyka ekran | — |
| Nowy | kafelki: Otwórz plik z urządzenia (wiodący), Pusty skoroszyt, Przykładowy plik; lista „Ostatnio otwierane” | `#emptyOpenBtn`, `#emptyNewBtn`, `#loadSampleBtn` (klik); `swbRecent.render(host)` |
| Otwórz | wybór pliku | `#emptyOpenBtn` (FSA albo `#fileInput`) |
| Informacje | nazwa pliku, arkusze, bieżący arkusz, wiersze/kolumny; „Wczytywanie: arkusz i nagłówek…” | globalne `workbook`, `currentFileName`, `currentSheetName`; `swbRibbon.openPanel("panel-file-sheet")` |
| Zapisz jako | — | `#saveAsBtn` (klik) |
| Eksportuj | przycisk „Eksport…” | `#exportCsvBtn` przeniesiony (gniazdo) |
| Opcje | Język, Motyw, Kopia ustawień | wiersze z menu ⋯ przeniesione (gniazda) |
| Pomoc i skróty | Skróty klawiszowe, Dziennik zdarzeń | `#shortcutsMenuItem` przeniesiony; `openPanel("panel-log")` |

Pominięte: Anonimizuj kopię (jest na karcie Dane), Zamknij i Zgłoś błąd (takich funkcji w Sheet nie ma).
Menu ⋯ po przeniesieniu: Asystent AI… (już na karcie Asystent), Odśwież, inne aplikacje, stan, wersja.

## Zachowanie
- Wjeżdża pod rzędem 1: przenikanie + 12 px w dół, 280 ms `--ease-drawer`; przy „ogranicz ruch” samo przenikanie.
- Otwarcie: fokus na wybranej sekcji, reszta aplikacji `inert`. Esc / Wróć / Plik: zamyka, fokus na „Plik”.
- Telefon bez zmian (gniazda oddają kontrolki do menu ⋯).

## Kroki
1. Test (czerwony) w `scripts/ribbon-playwright.js`: otwarcie, inert, sekcje, Informacje z nazwą pliku,
   Opcje mają `#themeSeg` i `#langSwitch` (a ⋯ już nie), Eksport z ekranu otwiera okno i zamyka ekran,
   Esc oddaje fokus, 768 px: motyw wraca do ⋯.
2. Kod: znaczniki `#rbBackstage` w `#ribbon`, `app/backstage.js`, style w `styles/ribbon.css`, teksty PL/EN,
   `swbRecent.render(host)`. Usunąć tymczasowe `#rbFileMenu`.
3. Stare testy, które szukają tych kontrolek w ⋯ / w menu Plik — gałąź „wstążka”.
4. Pełny `--all`, strażnik buildu, zrzuty, wydanie v1.05.00, PR.
