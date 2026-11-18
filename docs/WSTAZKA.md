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
| 1 | Rama: dwa rzędy u góry (v1.04.00) | Rząd 1: logo, Plik, karty (Narzędzia główne, Wstaw, Dane, Formuły, Analiza, Widok, Asystent), Szukaj, Tryb (menu: Edycja, Podgląd, Spisywanie, Wpisywanie), Zapisz z licznikiem, ⋯. Rząd 2: Cofnij/Ponów + polecenia karty w jednej linii (Uproszczona). Przewijane rzędy z podpowiedzią „+N”. Chipy filtrów w rzędzie paska formuły. Nagłówek „hero” schowany na komputerze; karty arkuszy wracają pod tabelę. | plan: `docs/superpowers/plans/2026-11-18-wstazka-etap1.md` |
| 2 | Plik za kulisami | Ekran Plik: Nowy, Otwórz, Informacje, Zapisz jako, Eksportuj, Anonimizuj kopię, Zamknij · Opcje, Pomoc i skróty, Zgłoś błąd. Menu ⋯ odchudzone. | — |
| 3 | Szyna i panele + karta Tabela | Szyna 40 px: Filtry, Sortowanie, Notatki, Formuły, Sprawdzanie, Asystent z licznikami; panel wysuwany nad tabelę albo przypięty. Zastępuje długi panel z sekcjami i „Znajdź ustawienie…”. Kontekstowa karta Tabela (kursor w tabeli Excela), mignie dwa razy, nie przełącza się sama. | — |
| 4 | Dolny pasek | Karty arkuszy (z kolorem karty z pliku, przewijane), Widoczne X/Y, Średnia / Liczba / Suma zaznaczenia, powiększenie. | — |
| 5 | Mini pasek | Nad zaznaczeniem myszą (pod nim, gdy brak miejsca): B, wypełnienie, format liczb, Σ, Filtruj wg wartości, Notatka, Wyjaśnij. Klawiatura i dotyk bez paska. | — |
| 6 | Opcje wstążki | Klasyczna / Uproszczona, Zawsze / Tylko karty (Ctrl+F1), gęstość Mysz / Dotyk. | — |
| 7 | Telefon | Pasek u góry (wstecz, cofnij, nazwa + arkusz, tryb, ⋯), Szukaj z ✦ i filtrami, pasek formuły tylko przy edycji, rząd poleceń nad paskiem cyfr, arkusz kart w miejscu klawiatury. | — |

## Do decyzji Mateusza (domyślnie — jak w makiecie)

1. Szkło na górnych paskach: **płaskie** (szkło zostaje w oknach i powiadomieniach).
2. „Analiza”: **osobna karta**.
3. Spisywanie i Wpisywanie: **w menu trybu**.
4. Anonimizuj: **na karcie Dane** (i w Plik → Anonimizuj kopię od etapu 2).
5. Wiersze tabeli 24 px z makiety v2: **nie w etapie 1** — dziś wysokość ustawia „Gęstość wierszy”;
   zmiana domyślnej gęstości to osobna decyzja (etap 6).

## Etap 1 — co weszło

`app/ribbon.js` + znaczniki `#ribbon` w `index.html` + `styles/ribbon.css` (osobny plik).
Od 769 px: rząd 1 (logo, Plik, 7 kart, Szukaj, Tryb, Narzędzia, Zapisz, ⋯), rząd 2 (Cofnij/Ponów +
polecenia karty), przewijane rzędy ze strzałką „+N”. Nagłówek „hero” i pasek nad tabelą schowane;
pasek formuły i chipy filtrów w jednym rzędzie; karty arkuszy pod tabelą; „Ostatnie” zapytania w
rozwiniętym polu Szukaj. Podgląd: bez kart i poleceń, nazwa pliku na środku, „Edytuj” zamiast „Zapisz”.
Telefon bez zmian (wszystko wraca na swoje miejsca). Test: `scripts/ribbon-playwright.js`.

Rozmieszczenie (jedna funkcja = jedno miejsce; przyciski z menu „Więcej ▾” i „Widok ▾” są
PRZENIESIONE, nie skopiowane):
- **Narzędzia główne:** Wklej / Wytnij / Kopiuj + cały dotychczasowy pasek formatowania (w nim Σ i
  „Formatuj jako tabelę”).
- **Wstaw:** Kolumna wyliczana, Notatka (Shift+F2), Arkusz.
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
- Wyczyść filtry — „Wyczyść” przy chipach filtrów.

Tymczasowe odstępstwa (znikają w kolejnych etapach):
- 🔧 Narzędzia (cały panel boczny) stoi w rzędzie 1 — do etapu 3 (szyna);
- przyciski kart otwierają sekcję w panelu bocznym — do etapu 3 (wysuwany panel);
- „Plik” to małe menu — do etapu 2 (ekran Plik);
- powiększenie stoi na karcie Widok — do etapu 4 (dolny pasek).
