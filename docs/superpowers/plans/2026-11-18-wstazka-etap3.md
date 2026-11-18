# Wstążka — etap 3: szyna i jeden panel naraz (plan wdrożenia)

**Decyzje Mateusza (2026-11-18):** szyna 4 pozycje (Filtry z licznikiem, Sortowanie, Formuły, Sprawdzanie
wartości) — bez Notatek i Asystenta (nie ma takich paneli); każdy przycisk otwiera JEDEN panel z tytułem,
pinezką i ×; znikają 🔧 i „Znajdź ustawienie…”; bez dubli na kartach; szyna od 769 px; karta „Tabela”
osobno na koniec (pamięć: ribbon-table-tab-later).

**Architektura:** nowy `app/rail.js` (`swbRail`) korzysta z istniejącego panelu bocznego i paska ikon
(`frame-compact.js`: wysuwanie nad tabelę, pinezka ≥ 1024 px, Esc, klik obok). Przy wstążce:
- `#sidebarRail` dostaje 4 przyciski `.rb-rail-btn[data-rb-rail]` (stare przyciski grup chowane CSS-em);
  widoczny od 769 px (CSS dla 769–1023 jak „docked” ≥ 1024);
- `swbRail.show(panelIds, titleKey)`: klasa `:root.rb-pane`, widać tylko wskazane `details.panel`
  (`.rb-pane-on`), tytuł `#rbPaneTitle` w `.sidebar-top`, sekcje otwarte; ≥ 1024 przez
  `frameCompact.openGroup(grupa)`, węziej `setSidebarOpen(true)`; zamknięcie panelu zdejmuje tryb
  (po 320 ms, jak `releaseFocusLater`);
- `swbRibbon.openPanel(id)` → `swbRail.show([id], tytuł panelu)` — karty i ekran Plik używają tego samego.

**Bez dubli / każdy panel ma dom:**
- szyna: Filtry = text-filters + date-filter + smart-filters; Sortowanie = sort-workbench;
  Formuły = formula-workbench; Sprawdzanie = validation;
- z kart znikają: Filtry tekstowe, Tryby auto, Filtr dat, Sortowanie, Sprawdzanie wartości (Dane),
  Formuły w arkuszu (Formuły);
- Narzędzia główne + „Narzędzia edycji” (panel-edit-tools — dziś tylko przez 🔧);
- Plik › Pomoc + „Skróty wyszukiwania” (panel-shortcuts);
- reszta bez zmian (Analiza, Wstaw, Dane: Anonimizuj, Widok, Plik › Informacje, Plik › Pomoc › Dziennik).

**Liczniki:** tylko Filtry (`#panelToggleCount`). Formuły / Sprawdzanie: brak gotowego licznika,
liczenie przy każdej zmianie to koszt na dużych arkuszach — pominięte (do decyzji później).

## Testy (`scripts/ribbon-playwright.js`, sekcja E3; chromium + webkit)
1. 1440: szyna widoczna, 4 przyciski w kolejności, stare przyciski grup niewidoczne, 🔧 i szukajka panelu niewidoczne.
2. Filtry → otwarty panel z tytułem „Filtry” i dokładnie 3 sekcjami; drugi klik zamyka.
3. Analiza → Agregacje → tylko panel Agregacje, tytuł „Agregacje”.
4. Pinezka (≥ 1024): panel obok tabeli, dalej jeden panel; odpięcie.
5. Esc zamyka i oddaje fokus przyciskowi szyny.
6. 900 px: szyna widoczna, tabela jej nie zasłania, Sortowanie otwiera jeden panel.
7. Każdy `details.panel` ma dokładnie jeden dom (szyna albo przycisk wstążki / ekranu Plik); żaden
   przycisk karty nie otwiera panelu z szyny.
8. 768 px: bez szyny wstążki, panel jak dotąd (grupy, szukajka).
