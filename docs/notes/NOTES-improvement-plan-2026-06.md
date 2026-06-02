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
- [x] Detekcja dat `od2`/`do2`/`data2` w normalnym trybie — **zweryfikowana jako już działająca**
      (naprawiona wcześniej przez `parseRepeatedHeader`): klasyfikacja nagłówka, miary agregacji,
      pary od/do, analiza czasu, parsowanie wartości (też seriale Excela) i filtr dat wszystkie
      zdejmują sufiks. Jedyny realny gap znaleziony i naprawiony: flaga „arkusz procesu / SLA"
      (`collectSheetInsights`) — `normalizedHeaders` używa teraz bazowego nagłówka, więc arkusze
      z cyklicznymi `od2/do2` dostają flagę bez Wide-to-Long. To był kolejny „przeterminowany" punkt ToDo.

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

### 3. B — Pasek statystyk na dole (Excel-style status bar) — ✅ ZROBIONE (2 czerwca 2026)
- Pojawia się **tylko** przy zaznaczeniu zakresu (≥2 komórek); pojedyncza komórka = cisza
  (jak w Excelu). W spoczynku `hidden`, nie zabiera miejsca tabeli.
- Pokazuje: `Zakres RxC`, `Liczba` (niepuste komórki) + gdy są liczby: `Suma`, `Średnia`,
  `Min`, `Maks`. Formatowanie przez `toLocaleString(I18N[currentLang].locale)` (PL przecinek).
- **Model zakresu = prostokąt kotwica↔koniec**: `focusedCellState` (klik) = kotwica,
  `selectedCellState` (Shift+strzałki / Shift+klik) = ruchomy koniec. Idealnie mapuje istniejący
  dual-state, zero nowego stanu globalnego.
- Implementacja w `table.js`: `getSelectionRectangle()`, `parseCellNumber()` (tolerancyjny:
  spacje/nbsp = tysiące, `,`/`.` = dziesiętny, `%`), `computeSelectionStats()`, `formatStatNumber()`,
  `syncRangeHighlightInDom()` (klasa `.cell-in-range`), `updateCellStats()`. Wpięte w
  `setFocusedCell`/`setSelectedCell` (też gałąź czyszczenia) + koniec `renderTable` + `applyLanguage`.
- **Shift+klik** dodany w `ui-controls.js` (handler `tbodyEl` click) → zakres myszą/dotykiem
  bez klawiatury (tablet). i18n: klucze w `I18N` (PL+EN), świadomie NIE w STATIC_TRANSLATIONS
  (omija dual-dict trap). DOM: `<div id="cellStatsBar">` po `#tableScrollbar` w `.table-panel`.
- Zweryfikowane Playwrightem (przykładowy plik → zakres Shift+strzałki i Shift+klik):
  pasek startuje hidden → pokazuje Σ=54/śr=13,5/min=9/maks=18 dla 4×1, podświetlenie zakresu,
  live-switch PL→EN etykiet (13,5→13.5), znika po Shift+Esc, zero błędów konsoli. Smoke-test OK.
- Build bump: `20260602-01` → `20260602-02` (`scripts/bump-version.js`).

### 4. D — Pogrupowanie sidebara w sekcje — ✅ ZROBIONE (2 czerwca 2026)
- 14 paneli owinięte w **5 sekcji** `<section class="sidebar-group">` z nagłówkami (PL/EN,
  przez STATIC_TRANSLATIONS + setText, te same co tytuły paneli). Bez przenoszenia bloków
  (grupy = ciągłe zakresy DOM), więc zero ryzyka regresji; `:has(#panel-X[open])` i
  `querySelectorAll("details.panel")` działają dalej bo panele zachowały ID/klasy.
  | Sekcja (id) | Co wchodzi |
  |---|---|
  | Dane (`group-data`) | Plik i arkusz |
  | Filtry i widok roboczy (`group-work`) | Filtr 1, Filtr 2, Filtr dat, Akcje, Sortowanie, Widok |
  | Inspekcja arkusza (`group-inspect`) | Analiza workbench, KPI, Układ arkusza |
  | Analiza (`group-analyze`) | Agregacje, Formula Workbench |
  | Pomoc (`group-help`) | Log, Skróty i info |
  - Decyzja: „Widok" (zoom/freeze) trafił do grupy roboczej z filtrami (a nie do „Dane"),
    bo to kontrola widoku roboczego i pozwoliło uniknąć przenoszenia bloku HTML. Zweryfikowane
    wizualnie (Playlwright screenshot, PL i EN, brak błędów konsoli).
- [x] ZROBIONE (2 czerwca): **scalenie dwóch filtrów tekstowych** w jeden panel
  `#panel-text-filters` „Filtry tekstowe". Filtr 1 zawsze widoczny; Filtr 2 ukryty za
  przyciskiem „+ Dodaj drugi filtr" (z opcją „Usuń" czyszczącą jego stan). WSZYSTKIE ID inputów
  zachowane (searchQuery/searchQuery2, filterMode/2, filter1/2Columns+Pick, filterEmptyMode/2,
  filterNegate/2, filterOperators/2, onlyNonEmpty) → silnik filtrów (workbook.js applyFilters)
  NIETKNIĘTY, zero refaktoru. i18n przez STATIC+setText (textFilters/filterBlock1/2/
  addSecondFilter/removeFilter). resetFilterInputs() chowa Filtr 2. Zweryfikowane screenshotem PL+EN.
- [x] Kolizja nazw: grupa „ANALIZA" → przemianowana na **„Agregacje i formuły"** (PL) /
  „Aggregation & formulas" (EN), żeby nie zgrzytała z panelem „Analiza workbench" w Inspekcji.

### 4b. „Morph pill" — wspólny język animacji (2 czerwca 2026) ✅
Mateuszowi bardzo spodobała się animacja handle sidebara na mobilkach (morfująca pigułka:
płynny morph kształtu/pozycji/koloru przez CSS `transition` + glassmorphism + kapsuła
`border-radius:999px` + badge `::before` ze zmiennym glifem). Poprosił o przeniesienie tego
„feelu" w inne miejsca. Wybrane 3 (NIE pasek statystyk):
- **Przełączniki trybu (toolbar)** — `.toggle-pill` na `#wideLongToggle`/`#excelLayoutToggle`/
  `#readingToggle`. W stanie `[aria-pressed="true"]`/`.active` morfują z zaokrąglonego
  prostokąta (`--r-sm`) w pełną kapsułę w kolorze akcentu, z „wjeżdżającą" białą kropką stanu
  (`::before scale(0)→scale(1)`). Dodano `aria-pressed` do `readingToggle` w `setReadingMode`
  (pozostałe dwa już je miały). Specyficzność: `.btn.ghost.toggle-pill[aria-pressed="true"]`.
- **Przycisk Aktualizuj (PWA)** — `.app-update-btn:not(.hidden)` dostaje `updateMorphIn`
  (wjazd scale+radius) + `updatePulse` (pulsująca poświata, nieskończona). Styl GOTOWY pod
  Feature C (JS usuwa `.hidden`, gdy jest nowa wersja).
- **FAB „do góry"** — nowy `#scrollTopFab` w `.table-panel` (`position:relative` dodane).
  Jak handle sidebara: kółko 44px rozwija się w pigułkę z etykietą na hover/focus
  (`width 44→128px`, label `max-width 0→90`). Widoczny gdy `tableWrapEl.scrollTop > 120`
  (próg 240 był za wysoki — przykładowy plik przewija się max ~229px). Klik → `scrollTo top`
  (smooth; `auto` przy reduced-motion). i18n w STATIC (`scrollTop`/`scrollTopAria`, PL+EN).
- **PUŁAPKA do zapamiętania:** `--t-slow` = `"250ms ease"` (zawiera już timing-function).
  Użycie `var(--t-slow) cubic-bezier(...)` w `animation`/`transition` daje DWA timing-functions
  → cała deklaracja odrzucana po cichu. Dla własnego easingu podawać jawny czas
  (`260ms cubic-bezier(...)`), NIE `var(--t-slow)`.
- **Korekta (Mateusz, 2 czerwca):** przełączniki i przycisk Aktualizuj NIE mają morfować
  zaokrąglenia rogów — zostają na stałym `--r-sm`/pigułce (zmienia się tylko kolor + kropka +
  scale/puls). Morph kształtu zarezerwowany wyłącznie dla FAB (jego sens). Strzałka `↑` w FAB
  była nie wycentrowana → `left:50% + translate(-50%,-50%)` w stanie collapsed, zjeżdża do
  `left:20px` przy rozwinięciu (miejsce na etykietę). Build → `20260602-04`.
- **„Plumknięcie" / pop pojawienia (Mateusz, 2 czerwca, build `20260602-05`):** spodobała mu się
  animacja handle Schowaj/Wysuń na mobilkach — poprosił o lekki pop (skala „plop") dodany do
  handle GENERALNIE (desktop+mobile, nie tylko mobile) oraz lekko do przełączników toolbara.
  Zrobione keyframes `popPlop` (handle, mocniej) / `popPlopLight` (przełączniki, lekko) skalujące
  **niezależną właściwością `scale:`** — NIE `transform` — bo handle pozycjonuje się przez
  `transform: translateY/translateX(-50%)` i scale przez transform by to zepsuł. Wyzwalane JS:
  `replayPop(el, cls)` (remove→reflow→add, re-triggerowalne); handle „plumka" w `setSidebarOpen`
  TYLKO przy realnej zmianie stanu (`prevSidebarOpenState`, więc bez popu na init/`setSidebarOpen(true)`
  przy starcie); przełączniki — pop na każdy klik (listener na wideLong/excelLayout/reading).
  Reduced-motion zeruje. Zweryfikowane Playwrightem (brak popu na load, handle-pop+popPlop po zmianie,
  btn-pop+popPlopLight po kliknięciu, zero błędów).
- Reduced-motion: `@media (prefers-reduced-motion: reduce)` zeruje te animacje/transitiony.
- Zweryfikowane Playwrightem: toggle radius 8px→999px + kropka scale(1), reading aria-pressed,
  FAB hidden→visible→hover 44→128→scroll-to-0→hidden + i18n PL/EN, update btn animationName
  `updateMorphIn, updatePulse`. Zero błędów. Build `20260602-02 → 20260602-03`.

### 5. C — Powiadomienia o aktualizacji (PWA) — ✅ ZROBIONE (2 czerwca 2026)
- Problem: użytkownicy z PWA na ekranie głównym siedzą na starej wersji.
- **Okazało się przeterminowane — logika BYŁA już w `bootstrap.js`** (SW `SKIP_WAITING`,
  rejestracja, `updatefound`→`statechange`, `registration.waiting` na starcie,
  `controllerchange`→reload, `showAppUpdate` = odkrycie `#appUpdateBtn` + `toast(updateAvailable)`,
  klik→postMessage SKIP_WAITING). Klucze i18n (updateNow/updateReady/updateAvailable/refreshingApp)
  PL+EN obecne. Styl morph przycisku dodaliśmy wcześniej (build -03/-04).
- **Realny gap dodany (-06):** `registration.update()` leciał TYLKO raz na starcie → PWA wznowiona
  z ikony na ekranie głównym nie wykrywała nowego buildu. Dodano `checkForUpdate()` na
  `visibilitychange` (powrót na pierwszy plan) + `setInterval` co 30 min. To dokładnie adresuje
  motywację punktu.
- Zweryfikowane E2E (Playwright): kopia aplikacji serwowana, strona kontrolowana przez SW →
  podmiana `sw.js` (nowy CACHE_VERSION) → `visibilitychange` → przycisk „Aktualizuj" + toast
  się pojawiają, zero błędów. Build → `20260602-06`.

### Osobne ToDo (poza planem czerwcowym) — ✅ ZROBIONE (2 czerwca, build `20260602-07`)
- **Zwijany pasek narzędzi nad tabelą (mobile)** — `#toolbarToggle` (chevron) tuż za statusem
  zsuwa/wysuwa `.table-actions` + quick-search (animacja `max-height`), zostawiając tylko status →
  więcej miejsca na tabelę na telefonie. **Funkcja mobile-only** (`@media max-width:768px`): na
  desktopie toggle ukryty (`.btn.toolbar-toggle{display:none}` — wyższa specyficzność niż `.btn`,
  bo inaczej `.btn{display:inline-flex}` wygrywał) i klasa `toolbar-collapsed` jest bezczynna, więc
  układ desktopu (status + przyciski w jednej linii) nietknięty. Na mobile `.table-actions`/
  `.quick-search` dostają `flex-basis:100%` → przyciski w osobnej linii pod [status + strzałka].
  Stan zapamiętany w `localStorage` (`TOOLBAR_COLLAPSED_KEY`), etykieta i18n w `I18N`
  (toolbarCollapse/toolbarExpand, odświeżana w `applyLanguage`), „plumknięcie" przy kliknięciu
  (uogólniony `.btn-pop`). Logika: `setToolbarCollapsed`/`updateToolbarToggleLabel` w ui-controls.js.
  Zweryfikowane Playwrightem (collapse/expand, persist po reloadzie, PL/EN, desktop ukryty) + screenshoty.
- **Glassmorphism paneli Apple-style** — `.panel` dostał: półprzezroczyste tło + górny `linear-gradient`
  „sheen", mocniejszy `backdrop-filter: blur(22px) saturate(1.8)`, warstwowy miękki cień i wewnętrzną
  krawędź światła u góry (`inset 0 1px 0 ...`). Osobny wariant `[data-theme="dark"] .panel`.
  Zweryfikowane screenshotami (jasny + ciemny) — frosted, czyste, Apple-like.
  - **Refinement (Mateusz, build `20260602-10`):** podciągnięte w stronę „liquid glass" —
    `blur(24px) saturate(1.9) brightness(1.02)` (dark: 26/1.7/1.04), mocniejszy specular na górnej
    krawędzi (`inset 0 1px 0 rgba(255,255,255,.72)`), delikatny pełnokrawędziowy pierścień światła
    (`inset 0 0 0 1px ...`) i subtelna głębia u dołu (`inset 0 -12px 26px -20px ...`). Lekko, ale widać.

### 6. Handle sidebara na mobilkach — ✅ ZROBIONE (2 czerwca 2026)
- Domknięte w tej sesji bez osobnej przebudowy: morph pasek→pigułka „Schowaj panel" na dole,
  naprawione przewijanie palcem (touch axis-lock + overscroll-behavior), „plumknięcie" przy
  otwarciu/zamknięciu, pełne tłumaczenia PL/EN (CSS custom property na etykietę). Mateusz uznał
  za wystarczające — bez dalszych zmian.

---
## ✅ PLAN CZERWCOWY ZAMKNIĘTY (2 czerwca 2026)
Wszystkie etapy zrobione: szybkie wygrane, A (przykładowy plik), B (pasek statystyk), C (update PWA),
D (grupy sidebara + scalenie filtrów), morph + plumknięcie, osobne ToDo (zwijany toolbar mobile +
glassmorphism), dług i18n (fallback + walidator), pkt 6 (handle). Build na koniec: `20260602-09`.
Dalsze kierunki — patrz „Pomysły na później" poniżej.

## Dług techniczny i18n — ✅ UJEDNOLICONE (2 czerwca 2026, build `20260602-09`)

System miał **dwa równoległe słowniki** = powtarzalne źródło bugów (złapane 3×: `choose`,
`hintDefault` → surowy klucz; `sampleBtn` → „undefined"):
- `I18N` — przez `t(key)` (dynamiczne stringi w JS).
- `STATIC_TRANSLATIONS` — przez `applyStaticTranslations` (`copy.key`, statyczny DOM).

**Rozwiązanie (wdrożone):**
1. `t()` — łańcuch fallbacków: `I18N[lang] → STATIC[lang] → I18N.pl → STATIC.pl → key`.
   Tekst w „drugim" słowniku nie zwraca już surowego klucza.
2. `applyStaticTranslations` — `copy` to teraz **scalony słownik**
   `{...I18N.pl, ...STATIC.pl, ...I18N[lang], ...STATIC[lang]}` (STATIC bieżącego języka wygrywa,
   reszta to fallback) → `copy.X` nigdy nie jest „undefined".
3. **Walidator dev-time** `scripts/check-i18n.js` (+ `npm run check:i18n`, wpięty w `npm test`
   PRZED smoke-testem): wyłuskuje literały obiektów I18N/STATIC (dopasowanie nawiasów + `Function`),
   skanuje `t("…")` i `copy.X` w `app/*.js`, sprawdza pokrycie i parytet PL/EN. Pomija dynamiczne
   prefiksy (`t("profType_" + x)` — peek na `+`). Exit 1 przy braku pokrycia.
   Aktualny stan: I18N 368/368 (pl/en), STATIC 174/174, 312 literalnych `t()`, 109 `copy.X` — wszystko OK.

**Konwencja na przyszłość:** nowy tekst dodawaj do JEDNEGO słownika (statyczny DOM → STATIC;
budowany w JS → I18N) w OBU językach; fallback i walidator pilnują reszty. `npm test` złapie braki.

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
