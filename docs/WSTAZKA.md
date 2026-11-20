# Wstążka — wdrożenie makiety v2

Wzór: `docs/makiety/wstazka-v2.html` (zaakceptowana ogólnie 2026-11-18; v1 zostaje dla porównania).
Zasada jak w Documents Workbench (`../documents-workbench-pwa/docs/WSTAZKA.md`): **układ,
rozmieszczenie, teksty i zachowania 1:1 z makietą**, animacje mogą się różnić, każde konieczne
odstępstwo najpierw do decyzji Mateusza. Krój i kolory zostają z `styles/app.css`.

Mechanika jak w Docs: kontrolki to **te same elementy (te same id)**. Na komputerze (≥ 769 px)
`app/ribbon.js` przenosi je do gniazd `[data-rb]` wstążki, na telefonie wracają na swoje miejsca.
Moduły i testy ich nie gubią, a nowe przyciski klikają istniejące przyciski albo wołają istniejące
funkcje — bez kopii logiki. Jedna funkcja = jedno miejsce na wstążce.

## Etapy (każdy to osobne wydanie)

| # | Etap | Co wchodzi | Stan |
|---|------|------------|------|
| 1 | Rama: dwa rzędy u góry (v1.04.00) | Rząd 1: logo, Plik, karty (Narzędzia główne, Wstaw, Dane, Formuły, Analiza, Widok, Asystent), Szukaj, Tryb (menu: Edycja, Podgląd, Spisywanie, Wpisywanie), Zapisz z licznikiem, ⋯. Rząd 2: Cofnij/Ponów + polecenia karty w jednej linii (Uproszczona). Przewijane rzędy z podpowiedzią „+N”. Chipy filtrów w rzędzie paska formuły. Nagłówek „hero” schowany na komputerze; karty arkuszy wracają pod tabelę. | v1.04.00 (PR #13–15) |
| 2 | Plik za kulisami | Ekran Plik: Nowy, Otwórz, Informacje, Zapisz jako, Eksportuj, Anonimizuj kopię, Zamknij · Opcje, Pomoc i skróty, Zgłoś błąd. Menu ⋯ odchudzone. | v1.05.00 (PR #16) |
| 3 | Szyna i panele + karta Tabela | Szyna 40 px: Filtry, Sortowanie, Notatki, Formuły, Sprawdzanie, Asystent z licznikami; panel wysuwany nad tabelę albo przypięty. Zastępuje długi panel z sekcjami i „Znajdź ustawienie…”. Kontekstowa karta Tabela (kursor w tabeli Excela), mignie dwa razy, nie przełącza się sama. | v1.06.00 (PR #17); karta Tabela przeniesiona na koniec |
| 4 | Dolny pasek | Karty arkuszy (z kolorem karty z pliku, przewijane), Widoczne X/Y, Średnia / Liczba / Suma zaznaczenia, powiększenie. | v1.07.00 (PR #19) |
| 5 | Mini pasek | Nad zaznaczeniem myszą (pod nim, gdy brak miejsca): B, wypełnienie, format liczb, Σ, Filtruj wg wartości, Notatka, Wyjaśnij. Klawiatura i dotyk bez paska. | v1.08.00 |
| 6 | Opcje wstążki (6a pokazywanie i gęstość, 6b Klasyczna) | Klasyczna / Uproszczona, Zawsze / Chowaj przy przewijaniu / Tylko karty (Ctrl+F1), gęstość Mysz / Dotyk, gęstość wierszy z Automatyczną. | 6a: v1.09.00 · 6b: v1.10.00 |
| 7 | Telefon | Pasek u góry (wstecz, cofnij, nazwa + arkusz, tryb, ⋯), Szukaj z ✦ i filtrami, pasek formuły tylko przy edycji, rząd poleceń nad paskiem cyfr, arkusz kart w miejscu klawiatury. Trzy części: 7a góra, 7b dół, 7c panele na cały ekran. | 7a: v1.11.00 · 7b: v1.12.00 · 7c: v1.13.00 |

## Do decyzji Mateusza (domyślnie — jak w makiecie)

1. Szkło na górnych paskach: **płaskie** (szkło zostaje w oknach i powiadomieniach).
2. „Analiza”: **osobna karta**.
3. Spisywanie i Wpisywanie: **w menu trybu**.
4. Anonimizuj: **na karcie Dane** (i w Plik → Anonimizuj kopię od etapu 2).
5. Wiersze tabeli 24 px z makiety v2: **nie w etapie 1** — dziś wysokość ustawia „Gęstość wierszy”;
   zmiana domyślnej gęstości to osobna decyzja (etap 6).
   Rozstrzygnięte w etapie 6a: nowa domyślna gęstość wierszy „Automatyczna” (mysz = Standardowa 23 px,
   dotyk = Wygodna 29 px).

## Etap 1 — co weszło

`app/ribbon.js` + znaczniki `#ribbon` w `index.html` + `styles/ribbon.css` (osobny plik).
Od 769 px: rząd 1 (logo, Plik, 7 kart, Szukaj, Tryb, Narzędzia, Zapisz, ⋯), rząd 2 (Cofnij/Ponów +
polecenia karty), przewijane rzędy ze strzałką „+N”. Nagłówek „hero” i pasek nad tabelą schowane;
przycisk arkusza (lista + Nowy / Zmień nazwę / Usuń), pasek formuły i chipy filtrów w jednym rzędzie; „Ostatnie” zapytania w
rozwiniętym polu Szukaj. Podgląd: bez kart i poleceń, nazwa pliku na środku, „Edytuj” zamiast „Zapisz”.
Telefon bez zmian (wszystko wraca na swoje miejsca). Test: `scripts/ribbon-playwright.js`.

Rozmieszczenie (jedna funkcja = jedno miejsce; przyciski z menu „Więcej ▾” i „Widok ▾” są
PRZENIESIONE, nie skopiowane):
- **Narzędzia główne:** Wklej / Wytnij / Kopiuj + cały dotychczasowy pasek formatowania (w nim Σ i
  „Formatuj jako tabelę”).
- **Wstaw:** Kolumna wyliczana, Notatka (Shift+F2).
- **Dane:** Filtry tekstowe, Tryby auto, Filtr dat · Sortowanie, Domyślne sortowanie · Wide-to-Long,
  Sprawdzanie wartości · Anonimizuj.
- **Formuły:** Formuły w arkuszu, Wyjaśnij formułę.
- **Analiza:** Przegląd arkusza, Układ arkusza, Kluczowe liczby · Agregacje, Podsumowanie miesięczne · Raport.
- **Widok:** powiększenie i gęstość, menu Widok (układ jak w Excelu, wiersze nad nagłówkiem, pasek
  cyfr), Ustawienia widoku.
- **Asystent:** Asystent AI… (przeniesiony z ⋯).
- **Tryb ▾:** Edycja, Podgląd, Spisywanie, Wpisywanie.
- **Plik ▾ (tymczasowo, do etapu 2):** Plik i arkusz…, Eksport….

Świadomie bez dubli (do decyzji Mateusza, czy któryś przywrócić):
- Wiersz / Kolumna — są pod tabelą („+ Nowy wiersz”, „+ Nowa kolumna”), więc nie na karcie Wstaw;
- Wstaw funkcję — przycisk fx w pasku formuły; Autosuma — Σ na Narzędziach głównych;
- Zapytaj słowami — ✦ w polu Szukaj (karta Asystent ma tylko ustawienia modelu);
- Podobne nazwy — tryb grupowania w Agregacjach;
- Wyczyść filtry — „Wyczyść” przy chipach filtrów;
- Arkusz — „Nowy arkusz” jest na liście przycisku arkusza.

Tymczasowe odstępstwa (znikają w kolejnych etapach):
- 🔧 Narzędzia (cały panel boczny) stoi w rzędzie 1 — do etapu 3 (szyna);
- przyciski kart otwierają sekcję w panelu bocznym — do etapu 3 (wysuwany panel);
- „Plik” to małe menu — do etapu 2 (ekran Plik);
- powiększenie stoi na karcie Widok — do etapu 4 (dolny pasek);
- arkusze: przycisk z nazwą w rzędzie paska formuły, bez zakładek (jak przed wstążką, test
  app-frame S1) — zakładki u dołu wchodzą w etapie 4.

Poza tym: F6 traktuje wstążkę jako region „pasek”; po zamknięciu Spisywania / Wpisywania fokus
wraca na „Tryb ▾”. Stare testy, które klikały kontrolki w dawnych miejscach, mają gałąź „wstążka”
(to samo zachowanie, nowe miejsce; komentarz „etap 1 wstążki”).

## Etap 2 — co weszło (ekran Plik)

`app/backstage.js` + `#rbBackstage` w `#ribbon` + sekcja w `styles/ribbon.css`; plan:
`docs/superpowers/plans/2026-11-18-wstazka-etap2.md`. „Plik” otwiera ekran pod rzędem kart (przenikanie
+ 12 px, 280 ms; wyjście 200 ms; „ogranicz ruch” = samo przenikanie), reszta aplikacji `inert`, Esc /
Wróć / Plik zamyka i oddaje fokus „Plik”. Tymczasowe menu Plik z etapu 1 zniknęło.

- **Nowy:** Otwórz plik z urządzenia (wiodący), Pusty skoroszyt, Przykładowy plik + „Ostatnio otwierane”
  (ta sama lista co na pustym ekranie, tu także przy otwartym pliku).
- **Otwórz**, **Zapisz jako** — te same przyciski co dotąd.
- **Informacje:** plik, bieżący arkusz, liczba arkuszy, wiersze w widoku, kolumny; „Wczytywanie: arkusz
  i nagłówek…” (sekcja Plik i arkusz w panelu).
- **Eksportuj:** Eksport… (przeniesiony z menu Plik).
- **Opcje:** Język, Motyw, Kopia ustawień — przeniesione z ⋯.
- **Pomoc i skróty:** Skróty klawiszowe (z ⋯), Dziennik zdarzeń.
- Menu ⋯ przy wstążce: Asystent AI…, Odśwież, inne aplikacje, stan, wersja.

Bez dubli: Anonimizuj kopię pominięte (Anonimizuj jest na karcie Dane); Raport zostaje na Analizie.
Pominięte, bo takich funkcji w Sheet nie ma: Zamknij, Zgłoś błąd (makieta v2 ich nie wymagała).

## Etap 3 — co weszło (szyna i jeden panel naraz)

`app/rail.js` (`swbRail`) + przyciski `.rb-rail-btn` w `#sidebarRail` + sekcja w `styles/ribbon.css`;
plan: `docs/superpowers/plans/2026-11-18-wstazka-etap3.md`. Decyzje Mateusza z 2026-11-18.

- Szyna przy wstążce (od 769 px): **Filtry** (z licznikiem; szukanie + filtr dat + tryby auto),
  **Sortowanie**, **Formuły w arkuszu**, **Sprawdzanie wartości**. Bez Notatek i Asystenta — nie ma
  takich paneli.
- Każdy przycisk (szyny, karty, ekranu Plik) pokazuje JEDEN panel z tytułem; pinezka (≥ 1024 px) stawia
  go obok tabeli. Wysuwanie, Esc i klik obok — jak dotąd (`frame-compact.js`). Znikają 🔧, uchwyt
  panelu, „Znajdź ustawienie…” i zakładki grup (na telefonie zostają).
- Bez dubli: z kart zniknęły przyciski paneli z szyny (Dane: filtry, Sortowanie, Sprawdzanie; Formuły:
  Formuły w arkuszu). Nowe: Narzędzia główne → **Narzędzia edycji** (dotąd tylko przez 🔧), Plik › Pomoc →
  **Skróty wyszukiwania**. Test pilnuje, że każdy panel ma dokładnie jeden dom.
- Liczniki błędów przy Formułach / Sprawdzaniu: pominięte (brak gotowego źródła, liczenie przy każdej
  zmianie to koszt) — do decyzji później.
- Karta kontekstowa „Tabela”: osobny etap na koniec, prawie od zera (decyzja Mateusza).

## Etap 4 — co weszło (dolny pasek)

Dolny pasek (`#tableFoot`) przy wstążce: od lewej przycisk listy arkuszy (przełączanie + Nowy /
Zmień nazwę / Usuń; przy kilku arkuszach sama ikona, nazwa w aria-label) i karty arkuszy
(przewijane z wygaszeniem krawędzi), dalej „+ Nowy wiersz / Nowa kolumna”, statystyki (sumy widoku,
a przy zaznaczeniu zakres / liczba / suma / średnia — jak dotąd), powiększenie po prawej (PR #18).
Lista arkuszy otwiera się w górę, gdy pod przyciskiem brak miejsca.

Bez dubli: bez osobnego „+” przy kartach („Nowy arkusz” jest na liście) i bez „Widoczne X z Y”
(liczba wierszy jest w chipach filtrów). Pominięte: kolory kart z pliku (Sheet ich dziś nie czyta).

## Etap 5 — co weszło (mini pasek)

`app/mini-bar.js` + style na końcu `styles/ribbon.css`. Po zaznaczeniu myszą co najmniej dwóch
komórek (przeciągnięcie albo Shift+klik) nad zakresem staje mały pasek: pogrubienie, wypełnienie
(ostatni kolor), format liczb · Σ, Filtruj wg wartości · Notatka, Wyjaśnij. Gdy nad zakresem brak
miejsca (nagłówek kolumn) — pod nim. Przy przewijaniu gaśnie i wraca po zatrzymaniu; znika po
kliknięciu obok, po klawiszu (też Esc) i po Σ / Filtruj / Notatce / Wyjaśnij (po B, wypełnieniu i
formacie zostaje). Jedna komórka, klawiatura, dotyk, Podgląd i telefon — bez paska.

Powtórki ze wstążki są celowe (decyzja Mateusza 2026-11-19): to skrót przy zaznaczeniu. Przyciski
klikają te same kontrolki co wstążka (`#formatRibbon`, `#fmtFillBtn`, `#fmtNumFormat`, `#fmtSumBtn`,
Notatka i Wyjaśnij z kart Wstaw / Formuły), a „Filtruj wg wartości” to „Pokaż tylko …” z menu
komórki (`applyCellFilter`). Do decyzji później: pasek także przy kliknięciu jednej komórki.

## Etap 6a — co weszło (pokazywanie wstążki i gęstość; v1.09.00)

- Przycisk na końcu rzędu poleceń (ikona zależna od wyboru) → menu: **Zawsze** / **Chowaj przy
  przewijaniu** / **Tylko karty** (Ctrl/⌘+F1 przełącza Tylko karty ↔ Zawsze). To samo w Plik → Opcje.
- Chowaj przy przewijaniu: tabela w dół — rząd poleceń zwija się (160 ms), tabela zajmuje miejsce;
  w górę, klik karty albo Tab do poleceń — wraca.
- Tylko karty: klik karty wysuwa polecenia nad tabelą (tabela się nie przesuwa); ta sama karta drugi
  raz, klik poza wstążką albo Esc — chowa.
- Gęstość wstążki **Mysz / Dotyk** (Plik → Opcje): Dotyk = przyciski 36 px, rząd 50 px; mini pasek
  (etap 5) wtedy się nie pokazuje.
- Domyślne bez zapisanego wyboru (decyzja Mateusza 2026-11-19): sam dotyk (iPad bez myszki) = Dotyk,
  a iPad pionowo także Chowaj przy przewijaniu; po obróceniu do poziomu — Zawsze. Mysz = Zawsze + Mysz.
  Wybór zapamiętany na urządzeniu (`swb-rb-prefs-v1`).
- **Gęstość wierszy** przebudowana (`app/frame-compact.js`): nowa opcja **Automatyczna** (domyślna) —
  mysz = Standardowa 23 px, dotyk albo wstążka Dotyk = Wygodna 29 px; nadąża za podłączeniem myszki do
  iPada. Wygodna / Standardowa / Kompaktowa ręcznie (wygrywają z automatem). Wybór w menu powiększenia
  na dolnym pasku i w Plik → Opcje; lista w panelu Widok schowana przy wstążce (bez dubla).
- Dolny pasek przy 1024 px w jednej linii: statystyki zaznaczenia przewijają się w bok, prawa krawędź
  wygasa, gdy coś jest dalej (wcześniej pasek zawijał się na 86 px).
- Naprawa silnika przewijania (`app/virt-rows.js`), ujawniona przez nową domyślną gęstość: po skoku
  w środek tabeli okno wierszy liczyło się ze starą poprawką (drift), zapas nad widokiem znikał i
  pierwszy krok w górę podskakiwał (virt-guard 11). Dotyczyło też ręcznie wybranej Standardowej.
- Bez „Klasycznej” w menu — wejdzie z etapem 6b.

## Etap 6b — co weszło (wstążka Klasyczna; v1.10.00)

- Wybór: menu opcji przy rzędzie poleceń (Układ: Uproszczona / Klasyczna) i Plik → Opcje (Układ wstążki).
  Domyślnie Uproszczona; wybór zapamiętany na urządzeniu (`swb-rb-prefs-v1`).
- `app/ribbon-classic.js` buduje panele Klasycznej (rząd 92 px): duże przyciski (ikona nad podpisem),
  kolumny małych z podpisami, nazwy grup pod spodem. Przyciski to klony Uproszczonej (te same
  `data-rb-act` / `data-rb-panel`), kontrolki z id (Raport, Widok ▾, Asystent AI…) przenoszą się do gniazd.
- Blok formatu (`#formatRibbon`) zostaje w całości — jego obsługa w `sheet-edit-ui.js` słucha na całym
  bloku — a w grupy Czcionka / Wyrównanie / Liczba / Edycja układa go siatka CSS (podpisy grup i pod
  dużymi Σ i „Formatuj jako tabelę” widać tylko w Klasycznej).
- Grupy (tylko polecenia, które aplikacja ma — decyzja Mateusza 2026-11-19): Narzędzia główne: Schowek ·
  Czcionka · Wyrównanie · Liczba · Edycja · Narzędzia; Wstaw: Dane · Komórka; Dane: Sortowanie · Narzędzia
  danych · Ochrona; Formuły: Inspekcja; Analiza: Arkusz · Zestawienia · Wynik; Widok: Widok; Asystent.
  Grupa bez widocznych poleceń (np. Sortowanie bez sortowania) znika.
- Różnice od makiety: Malarz formatów w grupie Edycja (nie w Schowku — blok formatu jest jedną kontrolką);
  gęstość Dotyk dotyczy tylko Uproszczonej.
- Przy okazji (nagranie Mateusza 2026-11-19): przy ponownym wczytaniu pliku (np. „Formatuj jako tabelę”)
  pasek formuły nie miga już napisem o Wide-to-Long.

### Polecenia z makiety, których Sheet nie ma — do decyzji później

Wstaw / Usuń wiersz i kolumnę (Narzędzia główne → Komórki; Wstaw → Wiersze i kolumny), Lista rozwijana,
Arkusz (Shift+F11), Tryby auto, Filtr dat, Puste / niepuste, Presety sortowania, Sprawdzanie wartości,
Podobne nazwy (Dane), Wstaw funkcję (Shift+F3), Szablony formuł, Przelicz (F9), Czas trwania, Zablokuj
okienka i zablokowany nagłówek, Gęstość wierszy i Przywróć szerokości kolumn (Widok → Siatka), Pasek cyfr,
Skupienie, Pełny ekran, Opisz formułę słowami, Wygładź podsumowanie, Procent i separator tysięcy (Liczba).
Część z nich istnieje w innym miejscu (szyna: filtry i sortowanie; menu Widok; panele) — przy dokładaniu
pilnować zasady bez dubli.

## Etap 7a — co weszło (góra ekranu na telefonie; v1.11.00)

Plan zaakceptowany przez Mateusza 2026-11-20 (podział 7a → 7b → 7c). Do 768 px, z otwartym plikiem:
- **Pasek u góry** (`#phTop`, `app/ribbon.js` placePhone): ← · Cofnij · nazwa pliku i arkusz · Tryb (ikona,
  to samo menu co na komputerze) · Zapisz · ⋯. Te same elementy (te same id) co we wstążce i w nagłówku.
  Duża karta „hero” z uchwytem zwijania jest wtedy schowana (bez pliku zostaje jak dotąd).
- **Zapisz** z licznikiem tylko przy niezapisanych zmianach. **Cofnij** stoi zawsze, wyszarzony bez historii.
- **←** = to samo pytanie co przy „Aktualizuj” (Zapisz / Nie zapisuj / Anuluj), potem ekran startowy z
  ostatnimi plikami (przeładowanie zamiast ręcznego czyszczenia stanu modułów).
- **Szukaj**: 16 px tekstu (iPhone nie powiększa strony), bez lupy (szuka klawisz „Szukaj” na klawiaturze).
  „Aa”, klucz i ▾ zostają w tym rzędzie do etapu 7b (rząd poleceń na dole) — żeby nic nie zniknęło w międzyczasie.
- **Pasek formuły** tylko podczas edycji komórki, w miejscu rzędu Szukaj (klasa `ph-editing`).
- **Karty arkuszy** pod tabelą jak na komputerze (przycisk listy jako ikona), bez pływającego przycisku arkusza.
- Pasek pod wycięciem ekranu (safe area), cele 44 px, reakcja przy dotknięciu.

Testy: `scripts/phone-top-playwright.js` (Chromium + WebKit). Testy zwijanej karty (`hero-pull-expand`,
`hero-collapse-guard`) usunięte — tej karty na telefonie z plikiem już nie ma; ich zasady „strona się nie
przewija” i „dolny pasek na ekranie” sprawdza nowy test. Do sprawdzenia na prawdziwym iPhonie: wycięcie
ekranu, klawiatura przy edycji, ← z niezapisanymi zmianami.

## Etap 7b — co weszło (dół ekranu na telefonie; v1.12.00)

Decyzje Mateusza 2026-11-20: polecenia te same co we wstążce na komputerze (wariant A, bez osobnego
zestawu), rząd zawsze w Edycji, pasek „Kopiuj / Wklej / Wypełnij” znika (menu kontekstowe może powtarzać).
- **Rząd poleceń** nad dolnym paskiem (`#phBottom`): przycisk z nazwą karty („Główne ▾”) + `#rbCmds`
  przeniesiony ze wstążki (te same przyciski i obsługa), 40 px, przewija się w bok. W Podglądzie go nie ma.
- **Arkusz kart** (`#phSheet`) po stuknięciu w nazwę karty: klawiatura się chowa, karty jako pigułki,
  polecenia w siatce 4 kolumn z podpisami (przyciski z samą ikoną dostają podpis z aria-label).
- **Klawiatura**: rząd poleceń wchodzi do paska cyfr (key-bar.js) i jedzie z nim nad klawiaturą.
- **Przeniesione na telefonie** (bez dubli): Spisywanie / Wpisywanie → menu Trybu; powiększenie → karta
  Widok; Eksport → ⋯ (ekranu Plik na telefonie nie ma); Ponów → pasek u góry, gdy jest co ponawiać;
  Wypełnij w dół / w prawo → karta Główne (także na komputerze).
- **Rząd Szukaj**: bez „Aa” i ▾; „Narzędzia” → **Filtry** (ikona lejka, licznik działających filtrów).
- Dymki na telefonie stoją pod paskiem u góry (wcześniej przykrywały Tryb i ⋯).

Test: `scripts/phone-bottom-playwright.js` (Chromium + WebKit). Na prawdziwym iPhonie: rząd poleceń nad
klawiaturą przy edycji komórki, arkusz kart, czy rząd nie jest za nisko przy kresce systemu.

## Etap 7c — co weszło (panele na cały ekran na telefonie; v1.13.00)

- Telefon z plikiem: panele otwierają się po jednym na cały ekran (ten sam tryb jednego panelu co szyna
  na komputerze, `app/rail.js`), z ← u góry (przycisk „Schowaj panel”) i tytułem. Wjazd z prawej.
- Znikają: uchwyt „Wysuń”, „Znajdź ustawienie…”, zakładki grup i pinezka (bez pliku zostają jak dotąd).
- Filtry (rząd Szukaj) otwiera panel filtrów; licznik działających filtrów zostaje.
- Szyna na kartach: Sortowanie i Sprawdzanie na Danych, Formuły na Formułach (na komputerze dalej szyna).
- Linki z ekranu Plik: Wczytywanie (arkusz i nagłówek), Skróty wyszukiwania, Dziennik zdarzeń — w ⋯.

Test: `scripts/phone-panels-playwright.js`. `drawer-tabs-playwright.js` sprawdza zakładki grup bez pliku.

## Etap 8 — co weszło (ekran startowy jak w Documents Workbench; v1.14.00)

Dodany przez Mateusza 2026-11-20 jako końcowy (przed kartą Tabela). Decyzje: karty „Zobacz na przykładzie”
zostają (mniejsze), bez „Przewodnika”, tytuł „Otwórz arkusz”.
- **Bez pliku** (`body.swb-nofile`, ribbon.js): na komputerze logo · Plik · „Sheet Workbench” · ⋯ — bez kart,
  poleceń, szukania, trybu, szyny i dolnego paska; na telefonie karta z nazwą i ⋯, bez klucza i „Wysuń”.
- **Środek**: kafelek z ikoną, „Otwórz arkusz” (28 px), jedno zdanie o prywatności, „Otwórz plik…” (główny),
  „Nowy skoroszyt”, „Przykładowy plik” (na telefonie główny na całą szerokość, dwa pod nim), ostatnie pliki,
  „albo przeciągnij plik w dowolne miejsce okna” (bez tego na dotyku).
- **Przeciąganie** pliku w dowolne miejsce okna (bootstrap.js; jak app-frame.js w Documents).
- **Ustaw jako arkusz startowy** (prawy klik / przytrzymanie na karcie arkusza i w liście arkuszy): activeTab
  + jedno tabSelected w pliku — Excel otwiera plik na tym arkuszu (sprawdzone w Excelu 2026-11-20).

Testy: `scripts/start-screen-playwright.js`, `scripts/start-sheet-playwright.js`.

