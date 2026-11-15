# Wpisywanie — kolejka numerów (2026-11-16)

## Po co
Mateusz przepisuje stosy kart (zakresy, karty pracy): zwrot kilku kart z tą samą datą, wydanie kilku
numerów jednej osobie, albo karty z różnymi danymi. Wpisywanie numer po numerze od nowa jest wolne.
Zatwierdzony kształt (rozmowa 2026-11-16): **jedna kolejka**, a „po kolei” vs „wszystkim naraz”
wynika z pól — pole **wspólne** (wypełnione na ekranie kolejki = przypięte) idzie do wszystkich,
reszta jest **osobna** dla każdego numeru.

## Zachowanie
1. **Lista w polu celu**: przecinek / średnik / nowa linia rozdziela pozycje; same liczby rozdzielone
   spacjami też; zakres `045-048` (liczby całkowite, rosnąco, max 500, zera wiodące jak w pierwszej).
   Pojedyncze „Jan Kowalski” i zapytanie z operatorami (`&&`, `||`…) to NIE lista. Duplikaty z listy
   usuwane (z informacją).
2. **Ekran kolejki** przed zapisem: każda pozycja ze statusem —
   ✓ jeden wiersz (dokładne trafienie, bez literówek — lista ma być precyzyjna),
   ? kilka wierszy (wybór z listy), ✗ brak w arkuszu (pomiń / nowy wiersz — domyślnie z Ustawień),
   ⊘ ten sam wiersz co wcześniejsza pozycja (pominięta).
   Pod listą pola karty: **wypełnione tutaj = wspólne dla wszystkich**.
3. **„Wpisz wszystkim naraz”** (tylko pola wspólne, tylko istniejące wiersze): jeden podgląd z listą
   wierszy, nadpisania jednym zaznaczeniem „nadpisz w N wierszach”, walidacja całości przed zapisem,
   jeden krok Cofnij na całą serię.
4. **„Po kolei”**: karta za kartą z polami wspólnymi już wpisanymi; „Wpisz i dalej →” od razu otwiera
   następną; pasek pozycji („3 z 6”, klikalne, „Pomiń”); nowe wiersze tylko tutaj.
5. Podsumowanie „Wpisano N · pominięto M”; niedokończona kolejka przeżywa zgaszenie tabletu.

## Ustawienia (Ustawienia → Kolejka)
`queueLists` (rozpoznawaj listy, wł.), `queueMissing` (`skip` | `new`, domyślnie `skip`),
`queueAutoNext` (po wpisaniu od razu następny, wł.).

## Granice
Bez notatek wspólnych (notatka jest per karta). Bez fuzzy w liście. „Naraz” nie tworzy wierszy.
Pod OCR: skan stosu kart wypełni tę samą kolejkę (`swbEntry.fillQueue`).
