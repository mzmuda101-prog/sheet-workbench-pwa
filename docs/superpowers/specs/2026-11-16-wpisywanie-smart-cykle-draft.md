# Wpisywanie — tryb smart: cykle „od lewej do prawej” (SZKIC, 2026-11-16)

Status: **plan z decyzjami Mateusza, nic nie zbudowane.** Prośba Mateusza 2026-11-16: „smart ogarnia
powtórzenia … od do, od2 do2 … wypełnia to co trzeba, tam gdzie trzeba, tyle co trzeba”, w opcjach,
na różnych arkuszach.

## Co jest w arkuszu (pomiar: „RODO Obieg terenów 2026 V5”, arkusz „2025-2026”)
- 33 kolumny: `Nr.` + **8 cykli** × [`Imię i Nazwisko`, `od`, `do`, `Długość` (formuła)].
  Nagłówki 2..8 mają cyfrę na końcu (`od2`, `do8`) — `parseRepeatedHeader` (analysis.js) już to rozpoznaje.
- Wzorce wypełnienia (563 wiersze): 259 „wzięta, nieoddana” (imię+od, bez do) · 215 „oddana” w cyklu 1 ·
  57 bez cyklu · 23 w cyklu 2 · pojedyncze anomalie: cykl 1 otwarty ORAZ cykl 2 zajęty (4),
  imię+do bez od (3), wszystkie 8 cykli zajęte — teoretycznie możliwe.

## Idea
Zamiast 33 pól karta pokazuje **akcje na wierszu**, a smart sam wybiera kolumny:
- **Wydaj** (nowy cykl): `Imię i Nazwisko` + `od` (domyślnie dziś) → pierwszy WOLNY cykl po ostatnim zajętym.
- **Zdaj** (zamknij cykl): `do` (domyślnie dziś) → ostatni OTWARTY cykl (ma `od`, nie ma `do`).
- Podgląd mówi wprost: „Wiersz 34 · 032 → do2: 16.11.2026 (cykl 2)”.
- Działa z kolejką: „Zdaj 031, 032, 045-048, data dziś” — **każdy wiersz dostaje datę w SWOJEJ kolumnie**
  (031 → do, 032 → do2…), jeden podgląd, jedno Cofnij. To jest główny zysk: stos zwróconych kart w kilka sekund.

## Wykrywanie (różne arkusze)
- Cykl = blok kolejnych kolumn o tych samych nazwach bazowych z rosnącym numerem (1, 2, 3…);
  kolumny przed blokiem (`Nr.`) to klucz/dane stałe. Wykorzystać `detectRepeatingBlocks` /
  `detectSignatureRepeatingBlocks` (analysis.js), także bloki ze scaleniami nad nagłówkiem.
- Role w cyklu zgadywane: pierwsza kolumna dat = **start**, druga kolumna dat = **koniec**, formuły = pomijane,
  tekst = **dane wydania** (osoba). Zgadywanie do poprawienia w Ustawieniach (np. „cykl zamyka: do”).
- Arkusz bez cykli → karta działa jak dziś (tryb smart się nie pokazuje).

## Bezpieczniki (anty-cele: wpis w zły cykl, nadpisanie historii)
- Smart **nigdy nie nadpisuje** zajętego cyklu; wpisuje tylko w puste komórki wybranego cyklu.
- Anomalie (dwa otwarte cykle, `do` bez `od`, wszystkie cykle zajęte) → wiersz oznaczony „do sprawdzenia”,
  bez zgadywania; w kolejce osobny status (jak „pasuje kilka wierszy”).
- „Zdaj” na wierszu bez otwartego cyklu → komunikat „ta karta nie jest wydana”, nic się nie zapisuje.
- „Wydaj” na wierszu z otwartym cyklem → pytanie „karta nie jest oddana — zamknąć poprzedni cykl dziś?”.
- Zapis tą samą drogą co dziś (writeCellFromInput, jedno Cofnij, kontrola „stale”).

## Opcje (Ustawienia → Tryb smart)
- Włącz tryb smart dla arkuszy z cyklami (domyślnie wł., gdy wykryto cykle).
- Data domyślna przy Wydaj/Zdaj: dziś / pusta.
- Role kolumn w cyklu (podgląd wykrycia + ręczna poprawka), zapamiętane per arkusz.
- Co przy braku wolnego cyklu: tylko ostrzeż / zaproponuj dodanie kolumn cyklu 9 (istniejące „＋ Nowa kolumna” — row-grow/xlsx-grow).
- Anomalie: pomiń / pokaż do decyzji.

## Decyzje Mateusza (2026-11-16)
1. **Nic na sztywno.** Nazwy akcji (domyślnie np. „Wydaj / Zdaj”) i słowo na rzecz w obiegu (karta, zakres,
   sprzęt, klucz…) ustawiane **per arkusz** — z tego korzystają różne osoby z różnymi arkuszami.
   W interfejsie neutralnie: „cykl”, „pozycja”, a słowa z Ustawień tam, gdzie mówią o akcjach.
2. **Brak wolnego cyklu → tylko ostrzeżenie.** (Dopisywanie kolumn cyklu 9 — może kiedyś.)
3. **Wydanie, gdy poprzedni cykl jest otwarty → ostrzeżenie, nie automat.** To nie jest zero-jedynkowe —
   to sygnał do wyjaśnienia (telefon do osoby, notatka). Okno pokazuje stan („cykl 1 otwarty od 04.01: Osoba 001”)
   i daje wybór: **Zamknij poprzedni (data) i wydaj** · **Wydaj mimo to** (poprzedni zostaje otwarty) ·
   **Tylko notatka** (zapisz notatkę na wierszu, nic nie wydawaj) · **Anuluj**. W kolejce taki wiersz dostaje
   status „do sprawdzenia” i ten sam wybór per pozycja. Te same zasady dla innych nietypowych wierszy.
4. **Cykle w pionie (kolejne wiersze)** — na razie nie, ale budowa ma to dopuszczać: wykrywanie cykli za
   jednym interfejsem („źródło cykli”), żeby pion dało się dodać bez przepisywania akcji i kolejki.

## Kolejność prac (po akceptacji)
silnik wykrywania cykli i stanu wiersza (Node, testy na wzorcach z pomiaru) → akcje Wydaj/Zdaj dla jednej
kartki → kolejka z akcją (różne kolumny per wiersz) → Ustawienia → prawdziwy plik + Excel + telefon → wydanie.
