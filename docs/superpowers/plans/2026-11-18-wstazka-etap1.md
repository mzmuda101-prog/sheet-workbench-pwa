# Wstążka — etap 1: rama z dwoma rzędami (plan wdrożenia)

> **Dla agentów:** wykonuj zadanie po zadaniu (superpowers:subagent-driven-development albo
> superpowers:executing-plans). Kroki z `- [ ]`.

**Cel:** na komputerze (≥ 769 px) nagłówek „hero” i paski nad tabelą zamieniają się w dwa rzędy
wstążki z makiety v2; telefon bez zmian.

**Architektura:** kopia mechaniki z Documents Workbench (`../documents-workbench-pwa/app/ribbon.js`).
Znaczniki `#ribbon` w `index.html` mają gniazda `[data-rb="<selektor>"]`; `app/ribbon.js` przy
≥ 769 px przenosi do nich ISTNIEJĄCE elementy (te same id, ta sama obsługa), a poniżej wstawia je z
powrotem na miejsce znacznika-komentarza. Nowe przyciski kart tylko klikają istniejące kontrolki
albo wołają istniejące funkcje.

**Technologia:** vanilla JS (globalne `const` w module IIFE, jak `dwbRibbon`), CSS w `styles/app.css`,
testy Playwright (Chromium + WebKit) uruchamiane przez `npm test`.

**Projekt:** `docs/WSTAZKA.md` (etapy i decyzje) + makieta `docs/makiety/wstazka-v2.html`
(rząd 1, rząd 2 Uproszczona, `.fbar` z chipami, menu trybu — teksty 1:1).

## Ograniczenia globalne

- Próg komputera: `matchMedia("(min-width: 769px)")`; klasa `body.rb-on`, gdy wstążka działa.
- Wysokości (Mysz): rząd 1 = 38 px, rząd 2 = 36 px, przyciski 28 px; Plik 26 px.
- Kolory i krój wyłącznie z tokenów `:root` / `[data-theme="dark"]` w `styles/app.css`
  (`--accent`, `--accent-h`, `--accent-light`, `--accent-border`, `--on-accent`, `--ink*`, `--muted`,
  `--border*`, `--surface`); paski wstążki płaskie (bez `--glass`), decyzja 1 z `WSTAZKA.md`.
- Każdy tekst w `I18N.pl` i `I18N.en` (`app/language.js`), `npm run check:i18n` zielony.
- Bez emoji; ikony jako inline SVG w stylu lucide, `aria-hidden="true"`; przycisk z samą ikoną ma
  `aria-label` i podpowiedź `data-hint` ze skrótem.
- Ruch: przewijanie rzędu `scroll-behavior:smooth` i jednorazowe drgnięcie strzałki „+N” — oba
  wyłączone przy `prefers-reduced-motion`. Karty strzałkami: bez animacji. Wciśnięcie
  `scale(.97)` 160 ms `--ease-out`; hover tylko w `@media (hover:hover) and (pointer:fine)`.
- Nowy plik JS: dopisać do `index.html` (z `?v=`), do listy w `sw.js` i do każdej listy plików, w
  której dziś jest `app/frame-compact.js` (`grep -rln frame-compact.js --exclude-dir=node_modules`).
- Commity bez stopki AI (pamięć: no-coauthor-trailer).

## Na co uważać przy przeglądzie

1. **Przejście 768 ↔ 769 px w trakcie pracy** (obrót iPada, zmiana okna) — każda przeniesiona
   kontrolka wraca na swoje miejsce, fokus i otwarte menu nie giną. Test: zadanie 1, krok „resize”.
2. **Pusty stan (brak pliku)** — wstążka nie pokazuje martwych poleceń: karty poza „Narzędzia
   główne” są, ale przyciski wymagające pliku są `disabled` tak samo jak dziś ich źródła. Test: zad. 2.
3. **Tryb Podgląd** — rząd 2 znika, w rzędzie 1 nazwa pliku na środku i „Edytuj” zamiast „Zapisz”;
   powrót do Edycji przywraca wybraną kartę. Test: zad. 2.
4. **Wąskie okno 769–900 px** — nic nie wychodzi poza ekran (`scrollWidth === innerWidth`), rzędy
   przewijają się ze strzałką „+N”, Szukaj zwija się do ikony. Test: zad. 1.
5. **Chipy filtrów w pasku formuły** — przy wielu filtrach chipy schodzą pod formułę zamiast
   ściskać pole formuły poniżej 160 px. Test: zad. 3.

---

### Zadanie 1: szkielet wstążki i przenoszenie kontrolek

**Pliki:**
- Utwórz: `app/ribbon.js`, `scripts/ribbon-playwright.js`
- Zmień: `index.html` (znaczniki `#ribbon` przed `<header class="hero">`, skrypt po `app-frame.js`),
  `styles/app.css` (nowa sekcja „Wstążka” na końcu), `sw.js`, `package.json` (`test:serial`:
  `node scripts/ribbon-playwright.js && ENGINE=webkit node scripts/ribbon-playwright.js`)

**Interfejsy:**
- Tworzy: globalne `const swbRibbon = { isOn(): boolean, select(tabId: string, opts?: {focus?: boolean}): void, current(): string, openPanel(detailsId: string): void }`
  oraz zdarzenie `document` `"swb:ribbon"` z `detail: { on: boolean }` przy każdej zmianie trybu.
- Karty (`data-rtab`): `home`, `ins`, `data`, `fx`, `an`, `view`, `ai`.
- Rząd 1 — gniazda: `#quickSearchWrap` (Szukaj), `#heroSaveBtn` (Zapisz z licznikiem),
  `#appMenuBtn` + `#appMenu` (⋯), tymczasowo `#panelToggle` (do etapu 3). Nowy przycisk trybu `#rbModeBtn`.
- Rząd 2 — gniazda stałe: `#undoBtn`, `#redoBtn`; przycisk „Opcje wstążki” jeszcze niewidoczny (etap 6).

- [ ] **Krok 1: test (czerwony).** `scripts/ribbon-playwright.js` (wzór uruchamiania i wczytania
  pliku: `scripts/app-frame-playwright.js`; plik `arkusze/obieg.xlsx`), viewport 1440×900:
  - `body.rb-on` jest; `#ribbon` widoczny; `.hero` ma `display:none`;
  - `[role=tablist]` ma 7 kart w kolejności z makiety, wybrana „Narzędzia główne”;
  - `#quickSearch`, `#heroSaveBtn`, `#appMenuBtn`, `#undoBtn` mają przodka `#ribbon`;
  - górna krawędź `#formulaBar` ≤ 82 px od góry okna (38 + 36 + margines);
  - `ArrowRight` na wybranej karcie wybiera „Wstaw” i przenosi fokus, bez klasy animacji;
  - resize do 768×900: `body.rb-on` znika, `#quickSearch` NIE ma przodka `#ribbon`,
    `.hero` widoczny; resize z powrotem do 1440: wraca;
  - 820×900: `document.documentElement.scrollWidth === innerWidth`; prawa strzałka „+N” rzędu
    kart widoczna i ma liczbę > 0;
  - 390×844 (dotyk): `#ribbon` `hidden`, układ jak dziś (`.hero` widoczny).
- [ ] **Krok 2:** `node scripts/ribbon-playwright.js` → FAIL (`body.rb-on` nie istnieje).
- [ ] **Krok 3: kod.** Znaczniki rzędów jak w makiecie v2 (`.r1`, `.r2`, gniazda `data-rb`), moduł
  `app/ribbon.js` z: `place(desk)` (przenoszenie przez komentarz-znacznik, jak w Docs), karty z
  klawiaturą (Strzałki, Home, End), `makeStrip(el)` z wygaszeniem krawędzi, strzałkami „+N”,
  kółkiem myszy w bok i jednym drgnięciem na sesję (`sessionStorage` `swb-rb-hinted-v1`, w try/catch),
  `ResizeObserver` przeliczający strzałki. Na `rb-on` chowaj CSS-em `.hero`, `#heroGrip`,
  `#toolbarToggle`, `#sheetModeSwitch`, rząd `.table-toolbar` (pusty po przeniesieniu) i
  zwijanie nagłówka (`frame-compact.js` ma nie zwijać, gdy `swbRibbon.isOn()`).
  Szukaj w rzędzie 1: widoczne pole + `#nlqOpenBtn` (✦); reszta `#quickSearchWrap`
  (`#qsFlags`, `#quickSearchAction`, `#quickSearchColumnsBtn`…) tylko przy `:focus-within`,
  szerokość 230 → 380 px w 180 ms `--ease-out`; poniżej 900 px szerokości okna sama ikona.
- [ ] **Krok 4:** test → PASS w Chromium i `ENGINE=webkit`.
- [ ] **Krok 5:** commit `feat(ribbon): etap 1 — szkielet wstążki, przenoszenie kontrolek, przewijane rzędy`.

### Zadanie 2: polecenia kart i menu trybu

**Pliki:** zmień `index.html` (zawartość paneli `.rb-panel[data-rtab]`), `app/ribbon.js`,
`app/language.js`, `styles/app.css`; test w `scripts/ribbon-playwright.js`.

**Interfejsy:** używa `swbRibbon.select`, `swbRibbon.openPanel` z zadania 1.
`openPanel(id)` = `setSidebarOpen(true)` + `document.getElementById(id).open = true` +
`scrollIntoView({block:"nearest"})` (panel boczny zostaje do etapu 3).

Zawartość kart (kolejność i nazwy z makiety v2; źródło = co klika / co woła):
- **Narzędzia główne:** cały `#formatRibbon` przeniesiony do gniazda (B, I, U, S, rozmiar, kolor,
  wypełnienie, obramowanie, wyrównanie, zawijanie, format liczb, dziesiętne, malarz, wyczyść, Σ,
  tabela). Nowe przed nim: Wklej / Wytnij / Kopiuj → to samo co Ctrl/⌘+V/X/C
  (`cutSelectionToClipboard`, `copySelectionToClipboard`; Wklej przez `navigator.clipboard.read()`
  → to samo zdarzenie „paste”, odmowa = komunikat „użyj Ctrl/⌘+V”).
- **Wstaw:** Wiersz, Kolumna (te same akcje co menu komórki), Kolumna wyliczana →
  `openPanel("panel-derived-columns")`, Lista rozwijana → `openPanel("panel-validation")`,
  Notatka → akcja `note` z `app/cell-notes.js` (Shift+F2), Arkusz → „Nowy arkusz” z `sheet-ops-ui.js`.
- **Dane:** Filtruj → `openPanel("panel-text-filters")`, Tryby auto → `panel-smart-filters`,
  Filtr dat → `panel-date-filter`, Wyczyść filtry → „Wyczyść” z `#activeFilters`, Sortuj / Presety →
  `panel-sort-workbench`, Domyślne sortowanie → klik `#resetSortBtn`, Z szerokiej na długą →
  `#wideLongToggle`, Sprawdzanie wartości → `panel-validation`, Podobne nazwy → tryb z
  `similar-names.js`, Anonimizuj → `panel-anonymize`.
- **Formuły:** Autosuma → `autoSum`, Szablony formuł / Wstaw funkcję → okno fx z `formula-ui.js`,
  Wyjaśnij formułę → `formula-explain.js` dla bieżącej komórki, Formuły w arkuszu →
  `panel-formula-workbench`.
- **Analiza:** Przegląd arkusza `panel-workbench-analysis`, Układ arkusza `panel-sheet-inspector`,
  Kluczowe liczby `panel-kpi-extractor`, Agregacje `panel-aggregation-workbench`, Podsumowanie
  miesięczne `panel-monthly-summary`, Raport → klik `#reportBtn`.
- **Widok:** powiększenie (`#zoomGroup` przeniesiony do gniazda — do etapu 4), `#viewMenuBtn`
  (Układ jak w Excelu, wiersze nad nagłówkiem, pasek cyfr), Gęstość → `setDensity`,
  Zablokowany nagłówek / szerokości kolumn → `openPanel("panel-view")`.
- **Asystent:** Zapytaj słowami → klik `#nlqOpenBtn`, Model AI na urządzeniu… → klik `#aiMenuItem`.
- **Menu trybu `#rbModeBtn`** (role `menu`, pozycje `menuitemradio`, opisy i skróty z makiety):
  Edycja / Podgląd → klik odpowiedniego `#sheetModeSwitch .mode-btn[data-mode]`; Spisywanie →
  klik `#transcribeBtn`; Wpisywanie → klik `#entryBtn`. Esc zamyka i oddaje fokus.
  `#moreMenuBtn` znika na `rb-on` (wszystkie jego pozycje mają miejsce). W Podglądzie: rząd 2 `hidden`,
  w rzędzie 1 nazwa pliku i arkusza (`#heroTitle`, nazwa arkusza) na środku, „Edytuj” zamiast „Zapisz”.

- [ ] **Krok 1: test (czerwony):** każda karta ma ≥ 1 przycisk i żaden nie ma pustej nazwy
  dostępnej; bez pliku przycisk „Wiersz” jest `disabled`; po wczytaniu klik „Agregacje” otwiera
  `#panel-aggregation-workbench` (`open === true`); menu trybu → „Podgląd”: `.r2` niewidoczny,
  tekst „obieg.xlsx” w rzędzie 1, przycisk „Edytuj”; „Edytuj” → wraca ostatnio wybrana karta;
  „Spisywanie” otwiera ten sam ekran co `#transcribeBtn` (to samo, co sprawdza
  `scripts/transcribe-playwright.js`); Esc w menu oddaje fokus `#rbModeBtn`.
- [ ] **Krok 2:** test → FAIL.
- [ ] **Krok 3:** zawartość kart, menu trybu, teksty PL/EN.
- [ ] **Krok 4:** test → PASS (oba silniki), `npm run check:i18n` → zielony.
- [ ] **Krok 5:** commit `feat(ribbon): karty Narzędzia główne…Asystent i menu trybu`.

### Zadanie 3: pasek formuły z filtrami, arkusze pod tabelą, „Ostatnie”

**Pliki:** `app/ribbon.js`, `app/app-frame.js` (`mountSheetTabs`), `styles/app.css`; test w
`scripts/ribbon-playwright.js`.

- Na `rb-on` `#activeFilters` przenosi się do gniazda na końcu `#formulaBar` (flex z zawijaniem:
  formuła `flex: 1 1 160px`, chipy `margin-inline-start:auto`, kreska oddzielająca). Liczba
  „Widać X z Y” zostaje w chipach do etapu 4 (potem przechodzi na dolny pasek).
- `mountSheetTabs()` nie przenosi `#sheetTabs` do `#heroTabsSlot`, gdy `swbRibbon.isOn()`;
  zakładki zostają na swoim miejscu pod tabelą (i wracają do nagłówka na telefonie / bez wstążki).
- `#recentStrip` („Ostatnie”) na `rb-on`: lista pod polem Szukaj, widoczna przy `:focus-within`.

- [ ] **Krok 1: test (czerwony):** filtr „W toku” → `#activeFilters` ma przodka `#formulaBar` i
  ten sam `top` co `#nameBox` (± 2 px); 6 filtrów → `.fx` ≥ 160 px szerokości, chipy w drugiej
  linii; `#sheetTabs` pod tabelą (`top` > `top` siatki); szukanie „Osoba” → po fokusie w Szukaj
  widać chip „Osoba” z `#recentStrip`.
- [ ] **Krok 2:** FAIL. **Krok 3:** kod. **Krok 4:** PASS (oba silniki).
- [ ] **Krok 5:** commit `feat(ribbon): filtry w pasku formuły, arkusze pod tabelą, Ostatnie pod Szukaj`.

### Zadanie 4: stare testy, pełny przebieg, wydanie v1.04.00

- [ ] **Krok 1:** `npm run test:all`. Testy, które na komputerze szukają nagłówka (dziś:
  `app-frame`, `cell-format`, `hero-collapse-guard`, `hero-pull-expand`, `keyboard-nav`, `key-bar`,
  `quick-search-flags`, `pre-row-select`, `sheet-ops`, `ux-contract`, `undo-history`), poprawić tak,
  żeby sprawdzały TO SAMO zachowanie w nowym miejscu (np. nazwa pliku: w Plik/Podgląd albo w
  `document.title`). Testu nie wolno osłabić — każda zmiana asercji ma komentarz „etap 1 wstążki: …”.
  Testy hero na komputerze (`hero-collapse-guard`, `hero-pull-expand`) dostają viewport telefonu albo
  sprawdzają, że na `rb-on` nagłówek się nie zwija.
- [ ] **Krok 2:** `npm run test:all` → zielony (niestabilne wypisane osobno — zgłosić Mateuszowi).
- [ ] **Krok 3:** zrzuty 1440 / 1024 / 820 / 390 px, jasny i ciemny, porównanie z makietą v2; lista
  odstępstw dopisana do `docs/WSTAZKA.md` („Etap 1 — co weszło”, „Tymczasowe odstępstwa”).
- [ ] **Krok 4:** `npm run release` (v1.04.00), commit `chore(release): v1.04.00 — wstążka, etap 1`,
  PR do `main` — scalenie dopiero po zgodzie Mateusza.
