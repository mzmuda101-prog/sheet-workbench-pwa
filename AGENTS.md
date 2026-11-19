# Sheet Workbench — instrukcje dla agentów AI

PWA do pracy z plikami Excela (.xlsx) w przeglądarce, także offline i na telefonie/iPadzie.
Czysty JavaScript + HTML + CSS, bez frameworka. esbuild tylko skleja pliki w `dist/`.

Wspólne zasady wszystkich projektów (przeczytaj na starcie): `../ZASADY-AGENTOW.md`
@../ZASADY-AGENTOW.md

## Zanim zaczniesz
- W tym folderze może równolegle pracować inna sesja lub agent. Najpierw `git status` i `git log -3`;
  nie nadpisuj cudzych niezacommitowanych zmian.
- Zachowanie wzorowane na Excelu sprawdzaj w prawdziwym Excelu (AppleScript `formula local`,
  `string value`, `number format`), nie zgaduj. Wynik zapisz w komentarzu/teście jako „wyrocznia Excela”.
  Zamykaj tylko skoroszyty otwarte przez siebie, po nazwie — nigdy `close every workbook`.

## Mapa
- `index.html` — wejście; `app/*.js` — moduły apki (`core.js` = rdzeń, wersja w `APP_BUILD_VERSION`).
- `styles/` — CSS; `sw.js` — service worker (offline); `lib/` — biblioteki obce (nie edytuj).
- `scripts/*-playwright.js` — testy; `scripts/run-tests.mjs` — runner; `docs/samples/` — pliki testowe.

## Uruchomienie
- `npm run dev` → http://127.0.0.1:3000 · serwer testów: port 7821 (WebKit nie wpuszcza na niektóre
  porty, np. 4190 — pusta strona = sprawdź konsolę).

## Testy
- `npm test` — sam pomija kroki, których zmiana nie dotyczy (`scripts/test-impact.mjs`).
  `npm test -- słowo` = tylko pasujące; `-- --dry` = podgląd; `-- --why` = powody.
- Mała zmiana → testy powiązane. `npm run test:all` tylko przed dużym wydaniem / przy zmianach przekrojowych.
  Nie puszczaj nic dłuższego niż kilka minut bez uprzedzenia.
- Nowa funkcja = nowy test w `scripts/` (Chromium + WebKit, telefon 390 px, gdy dotyczy UI).
- Testy starego silnika tabeli wymuszają `virt=0`; domyślnie działa nowy (wirtualizacja).

## Wydanie
- Wersja `vMAJOR.MM.PP`: `npm run release` (+1), `-- skok` (do okrągłej …0/…5), `-- minor`, `-- major`.
- Przed podaniem tekstu commita sprawdź `git log -3` — jeśli ten numer już poszedł, podbij wersję.
- Commit: `typ(sheet): krótki tytuł po polsku (vX.YY.ZZ)`, potem punkty „- …”. Bez nazw prywatnych plików użytkownika.

## Zasady UI (skrót; pełne: `../.claude/rules/`)
- Każde `:hover` w `@media (hover: hover)`; ramki fokusu przycisków przez `:focus-visible` / `.kbd-focus`
  (przyciski mają `tabindex="0"`, więc w Safari stuknięcie daje fokus).
- Podpowiedzi przez `data-hint` (cursor-hint), nie `title`.
- Litery nie jako skróty klawiszowe (klawiatury tabletowe); F1–F12 tylko jako alias.
- To, co może się nie zmieścić, przewija się — nic nie ucina się po cichu.
- Bez emoji w UI, kodzie i commitach.

## Sprawdzanie kodu
- `npm run check` — TypeScript sprawdza pliki z `// @ts-check` na górze (bez zmiany budowania).
  Nowe pliki zaczynaj od `// @ts-check`.
- `npm run lint` — Biome, tylko reguły na prawdziwe błędy.
