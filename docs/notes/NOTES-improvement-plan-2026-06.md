# Plan ulepszeń — czerwiec 2026

Notatka robocza ustalona wspólnie (Mateusz + Claude) 1 czerwca 2026.
Kolejność i zakres uzgodnione; każdy większy krok = osobny commit (bezpieczny punkt powrotu).

## Zasada pracy

- Zmiany robimy **małymi krokami**, po każdym większym kroku osobny commit.
- Claude **nie commituje sam** — podaje gotowy tekst commita do skopiowania, Mateusz zatwierdza.
- Każdy commit = bezpieczny punkt, do którego łatwo wrócić, jeśli coś nie podejdzie wizualnie/UX.
- Trzymamy się filozofii produktu: local-first, offline, bezpieczne dla plików źródłowych, wygodne na tablecie, „workbench wokół Excela, nie klon".

## Kolejność wdrażania

### 1. Szybkie wygrane (niskie ryzyko)
- [x] Martwy kod klawiatury w `bootstrap.js` — był już usunięty wcześniej; sprzątnięto tylko
      zbędny podwójny pusty wiersz w handlerze `keydown`.
- [x] Artefakt `e.stopPropagation()` w `ui-controls.js` — usunięte dwa martwe wywołania
      (apply/close pickera); globalny handler kliknięć i tak robi early-return wewnątrz pickera.
- [x] Bug `count_rows` — w pickerze miar `Liczebność` jest teraz wzajemnie wykluczająca
      z miarami kolumnowymi (silnik i tak nadpisuje pozostałe miary przy obecności count_rows),
      więc agregacja nie zwija się po cichu do Liczebności (ani przez „Zaznacz wszystko", ani ręcznie).
- [x] Tłumaczenia EN — **analiza czasu** w pełni EN; **inspektor / sekcje / wykrywacz bloków /
      profiler kolumn / KPI** w pełni EN (~80 nowych kluczy i18n). Naprawiono też szerszy problem:
      `applyLanguage` nie re-renderował dynamicznych paneli → teraz przełączają język na żywo.
      Przy okazji znaleziono i naprawiono 2 realne bugi raw-key: `t("choose")` i `t("hintDefault")`
      pokazywały surowe klucze (klucze istniały tylko w `STATIC_TRANSLATIONS`, nie w `I18N`).
      Toasty i puste stany były już wcześniej czyste. Pozostaje świadomy edge: `entityLabel`
      (Osoba/Właściciel) wyprowadzany z danych — wymagałby przerobienia `pluralizeEntityLabel`.
- [ ] (z ToDo) Detekcja dat dla nagłówków `od2`/`do2`/`data2` bez wymuszania Wide-to-Long —
      realny bug poprawności, nie tylko wygoda. **NASTĘPNE.**

> Uwaga procesowa: niektóre punkty ToDo bywają nieaktualne (np. martwy kod był już usunięty,
> toasty/puste stany już przetłumaczone). Zawsze najpierw weryfikujemy w kodzie, czy punkt
> faktycznie wymaga pracy, zanim się go robi.

### 2. A — Przykładowy plik w pustym stanie ⭐
- Przycisk „Wypróbuj na przykładowym pliku" w empty state.
- **Plik ma być lekki — cel ~10–40 KB, NIE 2 MB** (szybki start, brak spowolnień).
  - `stress-test-workbench.xlsx` = 3,4 MB → za duży na demo.
  - Najlepiej wygenerować dedykowany mały plik nowym skryptem `scripts/gen-sample.js`
    (analogicznym do `gen-stress-test.js`).
- Demo ma pokazać od razu: filtry, daty, **powtarzalne bloki → Wide-to-Long**, agregacje.
- Cel: onboarding bez własnego pliku, bez ujawniania prywatnych danych (offline).
  Rozwiązuje też problem z ToDo: nowi nie wiedzą, jak działa Wide-to-Long.

### 3. B — Pasek statystyk na dole (Excel-style status bar) ⭐
- Pojawia się **tylko** przy zaznaczeniu komórek/wierszy z liczbami; w spoczynku znika
  (nie zabiera miejsca tabeli).
- Pokazuje: liczbę zaznaczonych komórek, Σ suma, średnia, min, max.
- Reuse istniejącej selekcji komórek + logiki agregacji.

### 4. D — Pogrupowanie sidebara w sekcje
- Z 14 paneli → **5 grup** (4–6 wg ustaleń, 3 to za mało):
  | Sekcja | Co wchodzi |
  |---|---|
  | Dane | Plik i arkusz · Widok |
  | Filtry | Filtr tekstowy (1+2 scalone) · Filtr dat · Akcje · Sortowanie i presety |
  | Inspekcja | Analiza workbench · Układ arkusza · KPI |
  | Analiza | Agregacje · Formula Workbench |
  | Pomoc | Log · Skróty i info |
- **Scalić dwa filtry tekstowe** w jeden panel z „+ dodaj filtr" (jak sort-builder).

### 5. C — Powiadomienia o aktualizacji (PWA)
- Problem: użytkownicy z PWA na ekranie głównym siedzą na starej wersji.
- Service Worker (`sw.js`) + `SKIP_WAITING` + istniejący `#appUpdateBtn` + toast
  „Nowa wersja — odśwież". Podpiąć `updatefound` / `controllerchange`.

### 6. Handle sidebara na mobilkach
- Już częściowo zrobione — przejrzeć i poprawić/zmienić, jeśli znajdzie się lepsze
  miejsce/pomysł. Nie przebudowa od zera.

## Pomysły na później (zaakceptowane kierunkowo, nie teraz)

- Sparkline / mini-histogramy w profilerze kolumn (liczbowe: histogram, tekstowe: top-N paski).
- Zapisane pełne sesje analizy (filtry + sort + widoczne kolumny + zoom) per typ pliku —
  rozszerzenie obecnych presetów sortowania; realizuje cel „macro-substitute".
- Command Palette (uwaga: Ctrl/⌘+K jest dziś zajęty na „Kolumny").
- Freeze pierwszej kolumny (nie tylko nagłówka) — jak panele w Excelu/Sheets.
- Data bars / heatmapa w komórkach liczbowych jako przełącznik widoku.
- Automatyczne podbijanie `?v=...` w `index.html` przez `scripts/bump-version.js`
  (dziś ręczne przy każdym `<script>` — łatwo o pominięcie).

## Powiązane notatki
- [ROADMAP.md](../../ROADMAP.md)
- [NOTES-priority-plan.md](./NOTES-priority-plan.md)
</content>
</invoke>
