// anonymize-ui.js — okno „Anonimizacja danych” (etapy 1–2, 2026-10-31; 3–4 2026-11-06: profile,
// polecenie tekstem, raport ryzyka k-anonimowości, klucz odwracania — logika w anonymize-tools.js).
//
// Przebieg: zakres (zaznaczenie / kolumny / arkusz / cały plik) → Skanuj → raport typów
// i kolumn z metodą do wyboru → ręczne korekty (odznaczanie wartości, wyjątki, własne słowa
// i wzorce) → podgląd „przed → po” → wynik:
//   • „Nowa kopia” — oryginał nietknięty; plik budowany z pendingEdits + zmian anonimizacji
//     (buildPatchedXlsx), czyszczony z śladów (anonymize-zip.js) i zapisywany pod nową nazwą;
//   • „Zmień otwarty plik” — zmiany w arkuszach jak inne narzędzia edycji (jeden krok Cofnij
//     dla wszystkich arkuszy), a czyszczenie śladów dzieje się przy KAŻDYM zapisie
//     (swbAnonPendingCleanup, wpięte w buildOutputBytes).
// Plan zmian liczy jedna funkcja (anonBuildPlan) — podgląd i wykonanie nie mogą się rozjechać.
// Wynik jest deterministyczny w obrębie okna (ziarno); „klucz spójności” (hasło) daje te same
// zamiany także w innym pliku i innym dniu.

let swbAnonPendingCleanup = null; // { session, opts, workbook } — czyszczenie śladów przy zapisie (tryb „w miejscu”)

const swbAnon = (() => {
  const PREF_KEY = "swb-anon-prefs-v1";
  const TXT = {
    pl: {
      panelTitle: "Anonimizacja",
      panelHint: "Ukryj dane osobowe w zaznaczeniu, kolumnach, arkuszu albo całym pliku: imiona i nazwiska, PESEL, NIP, konta, telefony, e-maile, adresy… Wszystko dzieje się na tym urządzeniu.",
      open: "Anonimizuj dane…",
      title: "Anonimizacja danych",
      scope: "Zakres",
      scopeSelection: "Zaznaczenie", scopeColumns: "Kolumny", scopeSheet: "Arkusz", scopeFile: "Cały plik",
      filteredOnly: "Tylko przefiltrowane wiersze",
      everything: "Zamień każdą niepustą komórkę w zakresie (nie tylko wykryte dane)",
      pickColumns: "Kolumny do przeszukania",
      key: "Klucz spójności (opcjonalnie)",
      keyHint: "To samo hasło = te same pseudonimy (z hasłem czterocyfrowe, np. „Osoba 4821”) i fałszywe dane w każdym pliku i przy każdym uruchomieniu. Puste = pseudonimy po kolei od 001, reszta za każdym razem losowo.",
      scan: "Skanuj", rescan: "Skanuj ponownie", scanning: "Przeszukiwanie… {done} / {total} komórek",
      found: "Przeszukano {cells} komórek ({sheets}). Do zmiany: {changes} komórek, {values} różnych wartości.",
      nothing: "Nie znaleziono danych osobowych w tym zakresie. Możesz dopisać własne słowa i wzorce albo wybrać „Zamień każdą niepustą komórkę”.",
      types: "Rodzaje danych i metoda",
      columns: "Kolumny",
      colAuto: "Wykryte w tekście", colSkip: "Pomiń kolumnę", wholeCell: "Cała komórka: ",
      colsMore: "Pokaż pozostałe kolumny ({n})", colsLess: "Pokaż tylko kolumny ze zmianami",
      values: "Wartości (odznacz, czego nie ruszać)",
      search: "Szukaj wartości",
      all: "Wszystkie", none: "Żadne",
      more: "…i jeszcze {n}. Zawęź wyszukiwaniem.",
      rules: "Własne reguły",
      terms: "Zawsze ukryj (słowo lub fraza w linii)",
      exceptions: "Nigdy nie ruszaj (wyjątki, w linii)",
      patterns: "Własne wzorce (wyrażenia regularne, w linii)",
      rulesNote: "Po zmianie reguł kliknij „Skanuj ponownie”.",
      preview: "Podgląd zmian",
      output: "Wynik",
      outCopy: "Nowa kopia pliku — oryginał zostaje bez zmian",
      outInPlace: "Zmień otwarty plik (działa Cofnij; zapisz potem plik)",
      traces: "Ślady poza komórkami (etap 2)",
      trMeta: "Właściwości pliku: autor, firma, tytuł, ścieżka do pliku",
      trNotes: "Notatki i komentarze (z autorami)",
      trNotesAnon: "anonimizuj", trNotesRemove: "usuń", trNotesKeep: "zostaw",
      trHF: "Nagłówki i stopki wydruku",
      trHFClear: "wyczyść",
      trFormula: "Zapamiętane wyniki formuł (Excel przeliczy przy otwarciu)",
      trCharts: "Wykresy i teksty w rysunkach",
      trPivots: "Pamięć tabel przestawnych (odświeżą się przy otwarciu)",
      trLinks: "Hiperłącza (mailto:), listy wyboru, stałe w formułach i nazwach",
      trStrings: "Stare teksty, które zostają w pliku po podmianie komórek",
      cancel: "Zamknij", run: "Anonimizuj", runCopy: "Utwórz zanonimizowaną kopię", runInPlace: "Anonimizuj w otwartym pliku",
      working: "Tworzenie pliku…",
      needScan: "Najpierw kliknij „Skanuj”.",
      noFile: "Najpierw otwórz plik.",
      noSelection: "Zaznacz komórki w tabeli (klik + Shift albo przeciągnięcie).",
      noColumns: "Wybierz co najmniej jedną kolumnę.",
      noBytes: "Kopii nie da się zbudować z tego pliku — użyj „Zmień otwarty plik”.",
      reportTitle: "Raport anonimizacji",
      reportCells: "Zmienione komórki: {n}.",
      reportSaved: "Zapisano: {name}",
      reportInPlace: "Zmiany są w otwartym pliku. Zapisz plik (Zapisz / Zapisz jako) — przy zapisie apka wyczyści też ślady poza komórkami. Cofnij: Ctrl/⌘+Z.",
      leakOk: "Test szczelności: w żadnej części pliku nie ma już oryginalnych danych ({parts} sprawdzonych części).",
      leakFound: "Test szczelności: oryginalne dane zostały jeszcze w {n} miejscach (lista niżej). Często to komórki poza wybranym zakresem — wybierz „Cały plik”, żeby je też objąć.",
      leakFoundFile: "Test szczelności: zanonimizowane wartości zostały jeszcze w {n} miejscach z danymi (lista niżej, z adresami komórek).",
      structTitle: "Do sprawdzenia: {n} nazw i nagłówków w strukturze pliku (nie wliczone do testu)",
      structNote: "Nazwy zakresów, arkuszy, tabel i ich kolumn oraz nagłówki apka zostawia — zależą od nich listy, tabele i formuły. Jeśli któraś z nich to imię i nazwisko, zmień ją ręcznie w Excelu.",
      lkCells: "Arkusz „{s}” — komórki:", lkSheetOther: "Arkusz „{s}” — poza komórkami (lista wyboru, nagłówek wydruku, hiperłącze…):",
      lkSheetName: "Nazwy arkuszy:", lkDefinedName: "Nazwy zakresów (Formuły → Menedżer nazw):", lkDefinedValue: "Treść nazw zakresów:",
      lkTableName: "Nazwy tabel:", lkTableColumn: "Tabela „{t}” — nazwy kolumn:", lkAppTitles: "Właściwości pliku — lista nazw arkuszy i zakresów:",
      lkPart: "{p}:", lkHeader: "nagłówek", lkFormula: "wynik formuły",
      lkHintNames: "To nazwa w strukturze pliku, a nie dane w komórce. Jeśli to nie są dane osobowe (np. „mikrofony”), to fałszywy alarm.",
      lkHintHeaders: "To nagłówki — apka ich nie zmienia (od nich zależą tabele, listy i formuły). Jeśli to nie dane osobowe, to fałszywy alarm.",
      plCore: "Właściwości pliku", plChart: "Wykres", plDrawing: "Rysunek / pole tekstowe", plPivot: "Tabela przestawna", plNotes: "Notatki i komentarze",
      linksOk: "Powiązania: {rules} list wyboru, {cells} komórek sprawdzonych — każda wartość nadal jest na swojej liście.",
      linksBroken: "Powiązania: {n} komórek ma wartość spoza swojej listy wyboru po anonimizacji:",
      lowHint: "„?” ({n}) — są w kolumnie z osobami, ale nie wyglądają na imię i nazwisko (np. etykieta „Porządkowi”). Zostają bez zmian; zaznacz, jeśli to jednak osoba.",
      personFormulaHint: "Plik ma formuły. Jeśli któraś składa imię i nazwisko w innej formie (np. „Nazwisko Imię”, „I. Nazwisko”), wybierz dla osób „Fałszywe dane” — z pseudonimu „Osoba 001” formuła zrobi „001 Osoba”.",
      linksFix: "Zwykle lista źródłowa (np. arkusz z bazą) była poza zakresem anonimizacji. Wybierz „Cały plik” — wtedy lista i komórki zmienią się tą samą zamianą.",
      linksNote: "Te same wartości dostają wszędzie tę samą zamianę, więc listy z „bazy”, WYSZUKAJ/X.WYSZUKAJ po nazwisku, nazwy zakresów i tabele dalej się zgadzają. Formuły przeliczą się z nowych danych.",
      leakPart: "{where}: {terms}",
      warnImages: "Plik ma obrazy — ich treści apka nie sprawdza.",
      warnRevisions: "Plik ma historię zmian współdzielenia — może zawierać stare wartości.",
      warnExternalLinks: "Plik ma łącza do innych plików — ścieżki mogą zawierać nazwę użytkownika.",
      warnEmbedded: "Plik ma makra albo osadzone obiekty — ich treści apka nie sprawdza.",
      cleaned: "Wyczyszczone ślady: {list}.",
      cl_metadata: "właściwości pliku", cl_authors: "autorzy", cl_formulaResults: "wyniki formuł", cl_formulaConstants: "stałe w formułach",
      cl_links: "łącza i listy", cl_headersFooters: "nagłówki/stopki", cl_notes: "notatki", cl_notesRemoved: "usunięte notatki",
      cl_unusedStrings: "stare teksty", cl_charts: "wykresy", cl_pivots: "tabele przestawne", cl_tablesSynced: "nazwy kolumn tabel",
      sharedStrings: "teksty komórek pliku", inSheet: "arkusz „{s}”",
      toastInPlace: "Zanonimizowano {n} komórek. Zapisz plik, żeby zmiany trafiły na dysk.",
      toastLeakOnSave: "Zapisano, ale oryginalne dane zostały jeszcze w {n} miejscach pliku (poza zakresem anonimizacji).",
      toastCleanFailed: "Zapisano BEZ czyszczenia śladów anonimizacji (notatki, wykresy, tabele przestawne) — szczegóły w Logu. Zapisz jeszcze raz albo użyj „Zapisz kopię”.",
      fileSuffix: " — zanonimizowany",
      conf_high: "pewne", conf_medium: "prawdopodobne",
      sheetsN: "{n} ark.",
      formulaSkipped: "Komórki z formułami ({n}) nie są zmieniane — ich wyniki przeliczą się z zanonimizowanych danych.",
      m_keep: "Zostaw", m_pseudo: "Pseudonim (Osoba 001)", m_fake: "Fałszywe dane (realistyczne)", m_mask: "Maskuj (J***, ***123)",
      m_generalize: "Uogólnij (rok, 62-xxx…)", m_redact: "Znacznik [TYP]", m_remove: "Usuń (pusta komórka)", m_hash: "Skrót z kluczem", m_shift: "Przesuń daty o stałą liczbę dni",
      t_custom: "Własne", t_person: "Osoba", t_pesel: "PESEL", t_nip: "NIP", t_regon: "REGON", t_iban: "Nr konta", t_card: "Karta płatnicza",
      t_idcard: "Dowód osobisty", t_passport: "Paszport", t_kw: "Księga wieczysta", t_email: "E-mail", t_phone: "Telefon", t_address: "Adres",
      t_postcode: "Kod pocztowy", t_place: "Miejscowość", t_dob: "Data urodzenia", t_plate: "Rejestracja", t_vin: "VIN", t_ip: "Adres IP", t_other: "Inne",
      label_person: "Osoba", label_address: "Adres", label_place: "Miejscowość", label_custom: "UKRYTE", label_other: "UKRYTE",
      label_iban: "KONTO", label_card: "KARTA", label_idcard: "DOWÓD", label_passport: "PASZPORT", label_phone: "TELEFON", label_email: "E-MAIL",
      label_postcode: "KOD", label_dob: "DATA", label_plate: "REJESTRACJA", label_kw: "KW",
    },
    en: {
      panelTitle: "Anonymization",
      panelHint: "Hide personal data in a selection, columns, a sheet or the whole file: names, national IDs, tax IDs, accounts, phones, e-mails, addresses… Everything happens on this device.",
      open: "Anonymize data…",
      title: "Data anonymization",
      scope: "Scope",
      scopeSelection: "Selection", scopeColumns: "Columns", scopeSheet: "Sheet", scopeFile: "Whole file",
      filteredOnly: "Filtered rows only",
      everything: "Replace every non-empty cell in scope (not only detected data)",
      pickColumns: "Columns to scan",
      key: "Consistency key (optional)",
      keyHint: "Same password = same pseudonyms (four digits with a password, e.g. “Person 4821”) and fake data in every file and every run. Empty = pseudonyms numbered from 001, everything else random each time.",
      scan: "Scan", rescan: "Scan again", scanning: "Scanning… {done} / {total} cells",
      found: "Scanned {cells} cells ({sheets}). To change: {changes} cells, {values} distinct values.",
      nothing: "No personal data found in this scope. Add your own words and patterns or choose “Replace every non-empty cell”.",
      types: "Data types and method",
      columns: "Columns",
      colAuto: "Detected in text", colSkip: "Skip column", wholeCell: "Whole cell: ",
      colsMore: "Show other columns ({n})", colsLess: "Show only columns with changes",
      values: "Values (uncheck what to keep)",
      search: "Search values",
      all: "All", none: "None",
      more: "…and {n} more. Narrow it down with search.",
      rules: "Custom rules",
      terms: "Always hide (word or phrase per line)",
      exceptions: "Never touch (exceptions, per line)",
      patterns: "Custom patterns (regular expressions, per line)",
      rulesNote: "After changing rules click “Scan again”.",
      preview: "Preview",
      output: "Result",
      outCopy: "New copy of the file — the original stays unchanged",
      outInPlace: "Change the open file (Undo works; save the file afterwards)",
      traces: "Traces outside cells (stage 2)",
      trMeta: "File properties: author, company, title, file path",
      trNotes: "Notes and comments (with authors)",
      trNotesAnon: "anonymize", trNotesRemove: "remove", trNotesKeep: "keep",
      trHF: "Print headers and footers",
      trHFClear: "clear",
      trFormula: "Cached formula results (Excel recalculates on open)",
      trCharts: "Charts and text in drawings",
      trPivots: "Pivot table caches (refresh on open)",
      trLinks: "Hyperlinks (mailto:), dropdown lists, constants in formulas and names",
      trStrings: "Old texts left in the file after cells are replaced",
      cancel: "Close", run: "Anonymize", runCopy: "Create anonymized copy", runInPlace: "Anonymize the open file",
      working: "Building the file…",
      needScan: "Click “Scan” first.",
      noFile: "Open a file first.",
      noSelection: "Select cells in the table (click + Shift or drag).",
      noColumns: "Pick at least one column.",
      noBytes: "A copy can't be built from this file — use “Change the open file”.",
      reportTitle: "Anonymization report",
      reportCells: "Changed cells: {n}.",
      reportSaved: "Saved: {name}",
      reportInPlace: "The changes are in the open file. Save it (Save / Save as) — saving also cleans traces outside cells. Undo: Ctrl/⌘+Z.",
      leakOk: "Leak test: no original data left in any part of the file ({parts} parts checked).",
      leakFound: "Leak test: original data is still in {n} places (listed below). Often these are cells outside the chosen scope — choose “Whole file” to cover them.",
      leakFoundFile: "Leak test: anonymized values are still in {n} places with data (listed below, with cell addresses).",
      structTitle: "To review: {n} names and headers in the file structure (not counted in the test)",
      structNote: "Range, sheet, table and column names and headers are kept — lists, tables and formulas depend on them. If one of them is a person's name, change it by hand in Excel.",
      lkCells: "Sheet “{s}” — cells:", lkSheetOther: "Sheet “{s}” — outside cells (dropdown list, print header, hyperlink…):",
      lkSheetName: "Sheet names:", lkDefinedName: "Range names (Formulas → Name Manager):", lkDefinedValue: "Range name contents:",
      lkTableName: "Table names:", lkTableColumn: "Table “{t}” — column names:", lkAppTitles: "File properties — list of sheet and range names:",
      lkPart: "{p}:", lkHeader: "header", lkFormula: "formula result",
      lkHintNames: "This is a name in the file structure, not data in a cell. If it isn't personal data (e.g. “microphones”), it's a false alarm.",
      lkHintHeaders: "These are headers — the app doesn't change them (tables, lists and formulas depend on them). If it isn't personal data, it's a false alarm.",
      plCore: "File properties", plChart: "Chart", plDrawing: "Drawing / text box", plPivot: "Pivot table", plNotes: "Notes and comments",
      linksOk: "Links: {rules} dropdown lists, {cells} cells checked — every value is still on its list.",
      linksBroken: "Links: {n} cells have a value outside their dropdown list after anonymization:",
      lowHint: "“?” ({n}) — in a column of people, but they don't look like a name (e.g. a label “Ushers”). They stay unchanged; tick one if it is a person after all.",
      personFormulaHint: "The file has formulas. If one builds the name in another form (e.g. “Last First”, “F. Last”), choose “Fake data” for people — from the pseudonym “Person 001” a formula makes “001 Person”.",
      linksFix: "Usually the source list (e.g. a database sheet) was outside the anonymized scope. Choose “Whole file” — then the list and the cells change with the same replacement.",
      linksNote: "The same values get the same replacement everywhere, so lists from a “database” sheet, VLOOKUP/XLOOKUP by name, range names and tables still match. Formulas recalculate from the new data.",
      leakPart: "{where}: {terms}",
      warnImages: "The file has images — the app doesn't check their content.",
      warnRevisions: "The file has shared-workbook revision history — it may hold old values.",
      warnExternalLinks: "The file links to other files — paths may contain a user name.",
      warnEmbedded: "The file has macros or embedded objects — the app doesn't check their content.",
      cleaned: "Cleaned traces: {list}.",
      cl_metadata: "file properties", cl_authors: "authors", cl_formulaResults: "formula results", cl_formulaConstants: "formula constants",
      cl_links: "links and lists", cl_headersFooters: "headers/footers", cl_notes: "notes", cl_notesRemoved: "removed notes",
      cl_unusedStrings: "old texts", cl_charts: "charts", cl_pivots: "pivot tables", cl_tablesSynced: "table column names",
      sharedStrings: "file cell texts", inSheet: "sheet “{s}”",
      toastInPlace: "Anonymized {n} cells. Save the file to write the changes to disk.",
      toastLeakOnSave: "Saved, but original data is still in {n} places of the file (outside the anonymized scope).",
      toastCleanFailed: "Saved WITHOUT cleaning anonymization traces (notes, charts, pivot tables) — see the Log. Save again or use “Save copy”.",
      fileSuffix: " — anonymized",
      conf_high: "certain", conf_medium: "likely",
      sheetsN: "{n} sheets",
      formulaSkipped: "Formula cells ({n}) are not changed — their results recalculate from the anonymized data.",
      m_keep: "Keep", m_pseudo: "Pseudonym (Person 001)", m_fake: "Fake data (realistic)", m_mask: "Mask (J***, ***123)",
      m_generalize: "Generalize (year, 62-xxx…)", m_redact: "Tag [TYPE]", m_remove: "Remove (empty cell)", m_hash: "Keyed hash", m_shift: "Shift dates by a fixed number of days",
      t_custom: "Custom", t_person: "Person", t_pesel: "PESEL", t_nip: "Tax ID (NIP)", t_regon: "REGON", t_iban: "Account no.", t_card: "Payment card",
      t_idcard: "ID card", t_passport: "Passport", t_kw: "Land register", t_email: "E-mail", t_phone: "Phone", t_address: "Address",
      t_postcode: "Postal code", t_place: "Town", t_dob: "Date of birth", t_plate: "Plate", t_vin: "VIN", t_ip: "IP address", t_other: "Other",
      label_person: "Person", label_address: "Address", label_place: "Town", label_custom: "HIDDEN", label_other: "HIDDEN",
      label_iban: "ACCOUNT", label_card: "CARD", label_idcard: "ID", label_passport: "PASSPORT", label_phone: "PHONE", label_email: "E-MAIL",
      label_postcode: "ZIP", label_dob: "DATE", label_plate: "PLATE", label_kw: "LR",
    },
  };
  // Etapy 3–4 (2026-11-06): profile, polecenie tekstem, raport ryzyka, klucz odwracania.
  Object.assign(TXT.pl, {
    profileSec: "Profil i polecenie",
    profile: "Profil", profileCurrent: "— bieżące ustawienia —", profileBuiltin: "Wbudowane", profileOwn: "Moje profile",
    profileSave: "Zapisz jako profil…", profileDelete: "Usuń profil", profileExport: "Eksportuj profile", profileImport: "Importuj profile…",
    profileNamePrompt: "Nazwa profilu (np. „RODO — wysyłka do kontrahenta”):",
    profileSaved: "Zapisano profil „{name}”.", profileDeleted: "Usunięto profil „{name}”.", profileApplied: "Profil „{name}”: ustawienia wczytane.",
    profileImported: "Zaimportowano profile: {n}.", profileImportBad: "To nie jest plik profili anonimizacji.", profileExportNone: "Nie masz jeszcze własnych profili.",
    profileHint: "Profil pamięta metody, ślady, reguły, wynik i typy kolumn po nazwie nagłówka — ten sam plik co miesiąc w kilka sekund. Hasła profil nie zapisuje.",
    pr_rodo_share: "RODO — wysyłka do kontrahenta", prd_rodo_share: "Osoby → pseudonimy, numery zamaskowane, daty urodzenia do roku, adresy do miejscowości, notatki usunięte.",
    pr_test_data: "Dane do testów (realistyczne)", prd_test_data: "Wszystko podmienione na fałszywe dane z poprawnymi sumami, daty przesunięte — plik wygląda jak prawdziwy.",
    pr_max: "Maksymalnie bezpiecznie", prd_max: "Znaczniki [TYP] albo usunięcie, daty do roku, notatki i stopki usunięte.",
    command: "Polecenie tekstem", commandPh: "np. PESEL zamaskuj, nazwiska na pseudonimy, daty do miesiąca, notatki usuń",
    commandApply: "Zastosuj", commandEmpty: "Wpisz polecenie.",
    cmdUnderstood: "Zrozumiałem:", cmdUnknown: "Nie zrozumiałem: {list}. Ustaw to ręcznie niżej albo napisz inaczej (np. „telefony maskuj”).",
    cmdNothing: "Nic nie zrozumiałem z tego polecenia. Przykład: „PESEL zamaskuj, nazwiska na pseudonimy, daty do miesiąca”.",
    cmdRest: "pozostałe ({n})", cmdDateMonth: "daty do miesiąca", cmdDateYear: "daty do roku",
    cmdColumn: "kolumna „{h}” → {m}", cmdScope: "zakres: {v}", cmdOutput: "wynik: {v}",
    tr_notes: "notatki", tr_metadata: "właściwości pliku", tr_headersFooters: "nagłówki i stopki",
    trv_remove: "usuń", trv_keep: "zostaw", trv_anonymize: "anonimizuj", trv_clear: "wyczyść", trv_true: "wyczyść", trv_false: "zostaw",
    risk: "Ryzyko rozpoznania osoby (po anonimizacji)",
    riskNone: "Brak cech pośrednich (data urodzenia, kod pocztowy, miejscowość, płeć, wiek, stanowisko…) w zakresie — raport nie ma czego liczyć.",
    riskOk: "{where}: każda kombinacja cech ({cols}) występuje co najmniej {k}× w {n} wierszach — nikt nie wyróżnia się sam.",
    riskBad: "{where}: {u} z {n} wierszy ({p}%) ma UNIKALNĄ kombinację cech: {cols}. Ktoś, kto zna te cechy, rozpozna tę osobę mimo pseudonimu.",
    riskSmall: "{where}: nikt nie jest unikalny, ale {s} wierszy ma kombinację wspólną z najwyżej {k} osobami ({cols}).",
    riskDrop: "Najwięcej daje „{c}” — bez tej cechy unikalnych zostaje {u}. Uogólnij ją (np. data do roku, kod do 62-xxx) albo usuń kolumnę.",
    riskExamples: "Przykłady unikalnych kombinacji: {list}",
    rkOpt: "Zapisz klucz odwracania (plik zaszyfrowany hasłem)",
    rkHint: "Klucz pozwala później przywrócić oryginalne wartości komórek (pseudonimizacja w rozumieniu RODO). Zawiera oryginalne dane — trzymaj go osobno od zanonimizowanego pliku. Bez hasła nie da się go otworzyć.",
    rkPass: "Hasło klucza (min. 8 znaków)", rkPass2: "Powtórz hasło",
    rkWeak: "Hasło klucza musi mieć co najmniej 8 znaków.", rkMismatch: "Hasła klucza się różnią.",
    rkSaved: "Zapisano klucz odwracania: {name} ({n} komórek).", rkFailed: "Nie udało się zapisać klucza odwracania ({msg}).",
    rkSuffix: " — klucz odwracania",
    restoreOpen: "Przywróć z klucza…",
    restoreTitle: "Przywróć oryginalne dane z klucza",
    restoreIntro: "Otwórz zanonimizowany plik, wybierz plik klucza i wpisz jego hasło. Wrócą oryginalne wartości komórek; komórki zmienione po anonimizacji zostaną bez zmian. Notatek, wykresów i właściwości pliku klucz nie przywraca.",
    restoreFile: "Plik klucza", restorePass: "Hasło", restoreRun: "Przywróć",
    restoreBadPassword: "Złe hasło albo uszkodzony plik klucza.", restoreBadFile: "To nie jest plik klucza anonimizacji.", restoreNoCrypto: "Ta przeglądarka nie ma szyfrowania WebCrypto.",
    restoreNoFile: "Wybierz plik klucza.",
    restoreDone: "Przywrócono {n} komórek.", restoreSkipped: "Pominięto {n} — ich wartość zmieniła się po anonimizacji albo nie ma arkusza: {list}",
    restoreNothing: "Żadna komórka nie pasuje do klucza — to chyba inny plik.",
    restoreOtherFile: "Klucz utworzono dla pliku „{f}” ({d}).",
    aiSec: "✦ Wolny tekst (AI na urządzeniu)",
    aiIntro: "Kolumny ze zdaniami ({cols}) mogą zawierać nazwiska i adresy wpisane w tekst, których reguły nie łapią. Asystent AI przeczyta {n} tekstów na tym urządzeniu i zaproponuje fragmenty — niczego nie zmienia sam; propozycje są niżej na liście wartości ze znakiem ✦, do zaznaczenia.",
    aiRun: "✦ Szukaj osób i adresów w zdaniach", aiRunAgain: "✦ Szukaj ponownie", aiStop: "Zatrzymaj",
    aiReason: "Wyszukiwanie osób i adresów w wolnym tekście potrzebuje asystenta AI na urządzeniu.",
    aiProgress: "Czytanie tekstów… {done} / {total}",
    aiDone: "AI zaproponowało {n} fragmentów ({m} nowych). Zaznacz na liście wartości (✦) te, które mają zostać ukryte.",
    aiNone: "AI nie znalazło dodatkowych danych osobowych w tych tekstach.",
    aiFailed: "Asystent AI nie odpowiedział poprawnie ({msg}) — reguły działają dalej bez niego.",
    aiHint: "✦ ({n}) — propozycje AI z wolnego tekstu. Zaznacz te, które mają zostać ukryte; reszta zostaje bez zmian.",
  });
  Object.assign(TXT.en, {
    profileSec: "Profile and command",
    profile: "Profile", profileCurrent: "— current settings —", profileBuiltin: "Built-in", profileOwn: "My profiles",
    profileSave: "Save as profile…", profileDelete: "Delete profile", profileExport: "Export profiles", profileImport: "Import profiles…",
    profileNamePrompt: "Profile name (e.g. “GDPR — sending to a contractor”):",
    profileSaved: "Saved profile “{name}”.", profileDeleted: "Deleted profile “{name}”.", profileApplied: "Profile “{name}”: settings loaded.",
    profileImported: "Imported profiles: {n}.", profileImportBad: "This is not an anonymization profiles file.", profileExportNone: "You have no profiles of your own yet.",
    profileHint: "A profile remembers methods, traces, rules, the result and column types by header name — the same monthly file in seconds. It never stores the password.",
    pr_rodo_share: "GDPR — sending to a contractor", prd_rodo_share: "People → pseudonyms, numbers masked, birth dates to year, addresses to town, notes removed.",
    pr_test_data: "Test data (realistic)", prd_test_data: "Everything replaced with fake data with valid checksums, dates shifted — the file looks real.",
    pr_max: "Maximum safety", prd_max: "[TYPE] tags or removal, dates to year, notes and footers removed.",
    command: "Command in words", commandPh: "e.g. mask PESEL, names to pseudonyms, dates to month, remove notes",
    commandApply: "Apply", commandEmpty: "Type a command.",
    cmdUnderstood: "Understood:", cmdUnknown: "Not understood: {list}. Set it by hand below or rephrase (e.g. “mask phones”).",
    cmdNothing: "Nothing understood in this command. Example: “mask PESEL, names to pseudonyms, dates to month”.",
    cmdRest: "others ({n})", cmdDateMonth: "dates to month", cmdDateYear: "dates to year",
    cmdColumn: "column “{h}” → {m}", cmdScope: "scope: {v}", cmdOutput: "result: {v}",
    tr_notes: "notes", tr_metadata: "file properties", tr_headersFooters: "headers and footers",
    trv_remove: "remove", trv_keep: "keep", trv_anonymize: "anonymize", trv_clear: "clear", trv_true: "clear", trv_false: "keep",
    risk: "Re-identification risk (after anonymization)",
    riskNone: "No indirect traits (birth date, postal code, town, sex, age, job…) in scope — nothing to compute.",
    riskOk: "{where}: every combination of traits ({cols}) appears at least {k}× in {n} rows — nobody stands out alone.",
    riskBad: "{where}: {u} of {n} rows ({p}%) have a UNIQUE combination of traits: {cols}. Someone who knows these traits will recognize the person despite the pseudonym.",
    riskSmall: "{where}: nobody is unique, but {s} rows share their combination with at most {k} people ({cols}).",
    riskDrop: "“{c}” contributes most — without it {u} stay unique. Generalize it (e.g. date to year, code to 62-xxx) or remove the column.",
    riskExamples: "Examples of unique combinations: {list}",
    rkOpt: "Save a reversal key (password-encrypted file)",
    rkHint: "The key lets you restore the original cell values later (pseudonymization under GDPR). It holds the original data — keep it apart from the anonymized file. It can't be opened without the password.",
    rkPass: "Key password (min. 8 characters)", rkPass2: "Repeat password",
    rkWeak: "The key password needs at least 8 characters.", rkMismatch: "The key passwords differ.",
    rkSaved: "Saved the reversal key: {name} ({n} cells).", rkFailed: "Couldn't save the reversal key ({msg}).",
    rkSuffix: " — reversal key",
    restoreOpen: "Restore from key…",
    restoreTitle: "Restore original data from a key",
    restoreIntro: "Open the anonymized file, choose the key file and type its password. The original cell values come back; cells changed after anonymization stay as they are. The key doesn't restore notes, charts or file properties.",
    restoreFile: "Key file", restorePass: "Password", restoreRun: "Restore",
    restoreBadPassword: "Wrong password or damaged key file.", restoreBadFile: "This is not an anonymization key file.", restoreNoCrypto: "This browser has no WebCrypto encryption.",
    restoreNoFile: "Choose a key file.",
    restoreDone: "Restored {n} cells.", restoreSkipped: "Skipped {n} — their value changed after anonymization or the sheet is missing: {list}",
    restoreNothing: "No cell matches the key — it's probably a different file.",
    restoreOtherFile: "The key was made for “{f}” ({d}).",
    aiSec: "✦ Free text (on-device AI)",
    aiIntro: "Columns with sentences ({cols}) may hold names and addresses typed into text that rules don't catch. The AI assistant reads {n} texts on this device and suggests fragments — it changes nothing by itself; suggestions appear in the value list below marked ✦, to tick.",
    aiRun: "✦ Find people and addresses in sentences", aiRunAgain: "✦ Search again", aiStop: "Stop",
    aiReason: "Finding people and addresses in free text needs the on-device AI assistant.",
    aiProgress: "Reading texts… {done} / {total}",
    aiDone: "AI suggested {n} fragments ({m} new). Tick the ones to hide in the value list (✦).",
    aiNone: "AI found no extra personal data in these texts.",
    aiFailed: "The AI assistant didn't answer correctly ({msg}) — rules keep working without it.",
    aiHint: "✦ ({n}) — AI suggestions from free text. Tick the ones to hide; the rest stays unchanged.",
  });
  const L = () => (typeof currentLang !== "undefined" && currentLang === "en" ? "en" : "pl");
  function tx(key, vars = {}) {
    const s = TXT[L()][key] ?? TXT.pl[key] ?? key;
    return s.replace(/\{(\w+)\}/g, (m, k) => (vars[k] ?? m));
  }
  const fmtN = (n) => Number(n).toLocaleString(L() === "en" ? "en-US" : "pl-PL");
  const el = (tag, cls, text) => { const e = document.createElement(tag); if (cls) e.className = cls; if (text != null) e.textContent = text; return e; };

  // ── stan okna ──
  const st = {
    scope: "sheet", filteredOnly: false, everything: false, columns: new Set(), key: "",
    methods: { ...SwbAnon.DEFAULT_METHODS }, colModes: {}, excluded: new Set(), included: new Set(),
    rules: { terms: "", exceptions: "", patterns: "" },
    output: "copy", traces: { ...SWB_ANON_ZIP_DEFAULTS }, search: "",
    scan: null, seed: null, plan: null, showAllCols: false,
    // etapy 3–4
    dateGeneralize: null, headerModes: {}, command: "", cmdResult: null, profileId: "",
    keyOn: false, keyPass: "", keyPass2: "",
    // etap 5: propozycje AI (typ|wartość → {type, value}); aiStamp unieważnia pamięć trafień
    aiTerms: new Map(), aiStamp: 0, aiBusy: false, aiMsg: "",
  };
  const PROFILES_KEY = "swb-anon-profiles-v1";
  function loadPrefs() {
    try {
      const p = JSON.parse(localStorage.getItem(PREF_KEY) || "null");
      if (!p) return;
      if (p.methods) Object.assign(st.methods, p.methods);
      if (p.rules) Object.assign(st.rules, p.rules);
      if (p.output === "copy" || p.output === "inplace") st.output = p.output;
      if (p.traces) Object.assign(st.traces, p.traces);
      if (p.scope) st.scope = p.scope;
      if (p.dateGeneralize === "month" || p.dateGeneralize === "year") st.dateGeneralize = p.dateGeneralize;
    } catch { /* brak pamięci — domyślne */ }
  }
  function savePrefs() {
    try { localStorage.setItem(PREF_KEY, JSON.stringify({ methods: st.methods, rules: st.rules, output: st.output, traces: st.traces, scope: st.scope, dateGeneralize: st.dateGeneralize })); } catch { /* bez pamięci */ }
  }
  const lines = (s) => String(s || "").split(/\r?\n/).map((x) => x.trim()).filter(Boolean);
  const labels = () => {
    const out = {};
    SwbAnon.TYPES.forEach((type) => { const k = `label_${type}`; out[type] = TXT[L()][k] || type.toUpperCase(); });
    return out;
  };
  function seedFromKey() {
    if (st.key.trim()) return parseInt(SwbAnon.sha256Hex(`swb-anon|${st.key.trim()}`).slice(0, 8), 16);
    if (st.seed == null) st.seed = (Math.random() * 2 ** 32) >>> 0;
    return st.seed;
  }

  // ── zakres → komórki ─────────────────────────────────────────────────
  function headerRow0Of(sheetName) {
    if (sheetName === currentSheetName && typeof currentHeaderRow === "number") return currentHeaderRow - 1;
    const ws = workbook.Sheets[sheetName];
    try { return detectHeaderRowSimple(ws) - 1; } catch { return 0; }
  }
  // { sheets:[{name, header0, ranges:{rows:Set|null, cols:Set|null, r0,r1,c0,c1}}] } albo { err }
  function resolveScope() {
    if (!workbook) return { err: "noFile" };
    const out = [];
    if (st.scope === "file") {
      workbook.SheetNames.forEach((name) => { if (workbook.Sheets[name] && workbook.Sheets[name]["!ref"]) out.push({ name, rows: null, cols: null }); });
      return { sheets: out };
    }
    const name = currentSheetName;
    if (!name || !workbook.Sheets[name]) return { err: "noFile" };
    if (st.scope === "selection") {
      const rect = typeof getSelectionRectangle === "function" ? getSelectionRectangle() : null;
      if (!rect) return { err: "noSelection" };
      const rows = new Set();
      for (let i = rect.rowStart; i <= rect.rowEnd; i++) {
        const row = rect.model.rows[i];
        if (row && Number.isInteger(row.rowIndex0)) rows.add(row.rowIndex0);
      }
      const cols = new Set();
      for (let c = rect.colMin; c <= rect.colMax; c++) cols.add(currentStartCol + c);
      return { sheets: [{ name, rows, cols, selection: true }] };
    }
    let rows = null;
    if (st.filteredOnly && Array.isArray(viewRows)) {
      rows = new Set();
      viewRows.forEach((row) => { if (row && !row.isLongViewRow && Number.isInteger(row.rowIndex0)) rows.add(row.rowIndex0); });
    }
    let cols = null;
    if (st.scope === "columns") {
      if (!st.columns.size) return { err: "noColumns" };
      cols = new Set([...st.columns].map((c) => currentStartCol + c));
    }
    return { sheets: [{ name, rows, cols }] };
  }

  const isDate = (v) => v instanceof Date && !Number.isNaN(v.getTime());
  // Skan: dla każdej kolumny w zakresie — nagłówek, profil, komórki (bez formuł).
  async function scan(onProgress) {
    const sc = resolveScope();
    if (sc.err) return { err: sc.err };
    SwbAnon.resetLearned();
    const compiled = SwbAnon.compileRules({ terms: lines(st.rules.terms), patterns: lines(st.rules.patterns), exceptions: lines(st.rules.exceptions) });
    const result = { cols: [], cells: 0, formulas: 0, sheets: sc.sheets.length, compiled, protectedCells: new Set(), headerCells: new Set() };
    let total = 0;
    const plans = sc.sheets.map((s) => {
      const ws = workbook.Sheets[s.name];
      const range = XLSX.utils.decode_range(ws["!ref"]);
      const header0 = s.selection ? -1 : headerRow0Of(s.name);
      total += (range.e.r - range.s.r + 1) * (range.e.c - range.s.c + 1);
      return { ...s, ws, range, header0 };
    });
    let done = 0;
    let tick = performance.now();
    for (const p of plans) {
      const { ws, range, header0 } = p;
      const tables = tablesOfSheet(p.name);
      for (let c = range.s.c; c <= range.e.c; c++) {
        if (p.cols && !p.cols.has(c)) { done += range.e.r - range.s.r + 1; continue; }
        const headCell = header0 >= 0 ? ws[XLSX.utils.encode_cell({ r: header0, c })] : null;
        const sheetHeader = headCell && headCell.v != null ? String(headCell.w ?? headCell.v) : "";
        // Kolumna dzieli się na odcinki: każda tabela Excela ma własny nagłówek (i profil),
        // reszta arkusza — nagłówek arkusza. Nagłówki i wiersze sum tabel są nietykalne:
        // Excel wymaga, żeby nazwa kolumny tabeli była identyczna z komórką nagłówka.
        const segs = new Map(); // klucz tabeli ("" = poza tabelą) → komórki
        for (let r = range.s.r; r <= range.e.r; r++) {
          done++;
          const tb = tables.find((t) => r >= t.r0 && r <= t.r1 && c >= t.c0 && c <= t.c1) || null;
          if (tb && (r < tb.r0 + tb.headerRows || r > tb.r1 - tb.totalsRows)) { result.protectedCells.add(`${p.name}!${XLSX.utils.encode_cell({ r, c })}`); continue; }
          if (!tb && r === header0) continue;
          if (p.rows && !p.rows.has(r)) continue;
          const cell = ws[XLSX.utils.encode_cell({ r, c })];
          if (!cell || cell.v == null || cell.v === "" || cell.t === "z" || cell.t === "e") continue;
          if (cell.f) { result.formulas++; continue; }
          result.cells++;
          const key = tb ? tb.name : "";
          if (!segs.has(key)) segs.set(key, { tb, cells: [] });
          segs.get(key).cells.push({ r, v: cell.v, above: !tb && header0 >= 0 && r < header0 });
          if (performance.now() - tick > 40) {
            onProgress && onProgress(done, total);
            await new Promise((res) => setTimeout(res, 0));
            tick = performance.now();
          }
        }
        for (const [key, seg] of segs) {
          const header = seg.tb ? (seg.tb.colNames[c - seg.tb.c0] ?? "") : sheetHeader;
          const cells = seg.cells;
          const dataCells = cells.filter((x) => !x.above);
          const profile = SwbAnon.profileColumn(header, dataCells.slice(0, 2000), compiled);
          // liczba: PESEL 2000–2009 zapisany liczbą ma 10 cyfr (zero z przodu zgubione) — szukamy w 11-cyfrowej formie
          const numText = (n) => { const t = String(n); return t.length === 10 && SwbAnon.validators.peselValid(`0${t}`) ? `0${t}` : t; };
          cells.forEach((x) => { x.hits = typeof x.v === "string" ? SwbAnon.findInText(x.v, compiled, header) : typeof x.v === "number" && Number.isInteger(x.v) && x.v > 999999 ? SwbAnon.findInText(numText(x.v), compiled, header) : []; });
          result.cols.push({ sheet: p.name, c, table: key, header, profile, cells, letter: XLSX.utils.encode_col(c) });
        }
      }
      if (header0 >= 0) for (let c = range.s.c; c <= range.e.c; c++) result.headerCells.add(`${p.name}!${XLSX.utils.encode_cell({ r: header0, c })}`);
    }
    onProgress && onProgress(total, total);
    return result;
  }

  const colKey = (col) => `${col.sheet}|${col.c}|${col.table || ""}`;

  // Tabele Excela arkusza (z currentTables, czytanych przy otwarciu pliku).
  function tablesOfSheet(sheetName) {
    const out = [];
    const seen = new Set();
    const all = typeof currentTables === "object" && currentTables ? currentTables : {};
    for (const tb of Object.values(all)) {
      if (!tb || seen.has(tb) || tb.sheetName !== sheetName) continue;
      seen.add(tb);
      let rg;
      try { rg = XLSX.utils.decode_range(tb.ref); } catch { continue; }
      out.push({ name: tb.displayName || tb.name || tb.ref, r0: rg.s.r, r1: rg.e.r, c0: rg.s.c, c1: rg.e.c, headerRows: Number.isFinite(tb.headerRows) ? tb.headerRows : 1, totalsRows: Number.isFinite(tb.totalsRows) ? tb.totalsRows : 0, colNames: tb.colNames || [] });
    }
    return out;
  }
  // typ kolumny z profilu/polecenia po nazwie nagłówka — tylko WŁASNE klucze (nagłówek „constructor”
  // czy „__proto__” nie może zwrócić czegoś z Object.prototype)
  function headerModeOf(header) {
    if (!header) return null;
    const k = SwbAnonTools.fold(header);
    return Object.prototype.hasOwnProperty.call(st.headerModes, k) ? st.headerModes[k] : null;
  }
  function colMode(col) {
    const m = st.colModes[colKey(col)];
    if (m) return m;
    const hm = headerModeOf(col.header);
    if (hm && hm !== "auto") return hm;
    if (st.everything) return col.profile.type || "custom";
    return col.profile.type || "auto";
  }
  const valueKey = (type, v) => `${type}|${SwbAnon.norm(v instanceof Date ? v.toISOString() : v)}`;

  // Numery pseudonimów już w pliku (po wcześniejszym przebiegu) — nowa osoba nie dostanie tego samego.
  // Pamięć na skoroszyt + stan danych: buildPlan leci przy każdej literze klucza.
  let reservedCache = null;
  function reservedPseudo() {
    const stamp = typeof sheetDataStamp !== "undefined" ? sheetDataStamp : 0;
    const lang = L();
    if (reservedCache && reservedCache.wb === workbook && reservedCache.stamp === stamp && reservedCache.lang === lang) return reservedCache.out;
    const values = [];
    (workbook?.SheetNames || []).forEach((n) => {
      const ws = workbook.Sheets[n];
      if (ws) for (const ref in ws) { const c = ref.charCodeAt(0) !== 33 && ws[ref]; if (c && typeof c.v === "string") values.push(c.v); }
    });
    const out = SwbAnon.collectReservedPseudo(values, labels());
    reservedCache = { wb: workbook, stamp, lang, out };
    return out;
  }

  // Plan: lista zmian {sheet, r, c, before, after:{value,type}|null(=wyczyść), type, original}
  function buildPlan() {
    const scanR = st.scan;
    if (!scanR) return null;
    const session = SwbAnon.createSession({
      seed: seedFromKey(), salt: st.key.trim() ? `k:${st.key.trim()}` : undefined,
      keyed: !!st.key.trim(), reserved: reservedPseudo(), dateGeneralize: st.dateGeneralize || undefined,
      methods: st.methods, labels: labels(),
      rules: { terms: lines(st.rules.terms), patterns: lines(st.rules.patterns), exceptions: lines(st.rules.exceptions) },
    });
    // 1. osoby: najpierw poznaj pełne imiona (żeby „A. Kamiński” trafił do „Adam Kamiński”).
    // Tylko z komórek, które wyglądają na osobę — etykieta „Porządkowi” w kolumnie nazwisk
    // nie może zostać „nazwiskiem” (szukanym potem w całym pliku).
    scanR.cols.forEach((col) => {
      const mode = colMode(col);
      col.cells.forEach((x) => {
        if (mode === "person" && !x.above && typeof x.v === "string" && SwbAnon.personCellScore(x.v) > 0) session.learnPerson(x.v);
        x.hits.forEach((h) => { if (h.type === "person") session.learnPerson(h.value); });
      });
    });
    // imiona/nazwiska spoza słownika poznane z kolumn osób → wykryj je też w wolnym tekście
    const learnedAny = session.learned.first.size || session.learned.sur.size;
    const aiList = [...st.aiTerms.values()];
    const hitsOf = (col, x) => {
      let base = x.hits;
      if (learnedAny && typeof x.v === "string") {
        if (!x.hits2) x.hits2 = session.findInText(x.v, col.header);
        base = x.hits2;
      }
      if (!aiList.length || typeof x.v !== "string") return base;
      if (x.hitsAiStamp !== st.aiStamp || x.hitsAiBase !== base) { x.hitsAi = SwbAnonTools.aiHits(x.v, aiList, base); x.hitsAiStamp = st.aiStamp; x.hitsAiBase = base; }
      return x.hitsAi;
    };
    scanR.cols.forEach((col) => col.cells.forEach((x) => { delete x.hits2; delete x.hitsAiBase; }));
    const changes = [];
    const values = new Map(); // valueKey → {type, value, count, conf}
    const typeCounts = {};
    // conf "low" = w kolumnie osób, ale nie wygląda na osobę (etykieta) — domyślnie odznaczone
    const note = (type, value, conf) => {
      const k = valueKey(type, value);
      const e = values.get(k);
      // ta sama wartość wykryta regułą i przez AI: pewniejsze wygrywa
      if (e) { e.count++; if (e.conf === "ai" && conf !== "ai") e.conf = conf; } else values.set(k, { key: k, type, value, count: 1, conf });
      if ((conf === "low" || conf === "ai") && !st.included.has(k)) return false;
      typeCounts[type] = (typeCounts[type] || 0) + 1;
      return !st.excluded.has(k);
    };
    for (const col of scanR.cols) {
      const mode = colMode(col);
      col.planned = 0;
      if (mode === "skip") continue;
      for (const x of col.cells) {
        let wholeType = !x.above && mode !== "auto" ? mode : null;
        // kolumna osób rozpoznana automatycznie: komórka, która nie wygląda na imię/nazwisko
        // („Porządkowi”, „STARSI”, „Data”) czeka na decyzję użytkownika; „Gr3 P. Król” ma osobę
        // w środku — zamieniamy tylko ją (kod grupy zostaje)
        const auto = !st.colModes[colKey(col)] && !headerModeOf(col.header) && !st.everything;
        let low = wholeType === "person" && auto && typeof x.v === "string" && SwbAnon.personCellScore(x.v) === 0;
        if (low && hitsOf(col, x).some((h) => h.type === "person")) { wholeType = null; low = false; }
        if (wholeType) {
          if (typeof x.v === "string" && !x.v.trim()) continue;
          if (!note(wholeType, x.v, low ? "low" : "high")) continue;
          const method = st.methods[wholeType] || "redact";
          if (method === "keep") continue;
          const after = session.replaceWholeCell(wholeType, x.v, method);
          if (after === null) continue;
          if (after.value === x.v) continue;
          if (typeof x.v === "string" && typeof after.value === "string") session.remember(x.v, after.value);
          changes.push({ sheet: col.sheet, r: x.r, c: col.c, before: x.v, after: after.type === null ? null : after, type: wholeType });
          col.planned++;
          continue;
        }
        const hits = hitsOf(col, x);
        if (!hits.length) continue;
        const kept = hits.filter((h) => note(h.type, h.value, h.conf) && (st.methods[h.type] || "redact") !== "keep");
        if (!kept.length) continue;
        if (typeof x.v === "number") {
          const h = kept[0];
          const after = session.replaceWholeCell(h.type, x.v, st.methods[h.type]);
          if (!after) continue;
          changes.push({ sheet: col.sheet, r: x.r, c: col.c, before: x.v, after: after.type === null ? null : after, type: h.type });
          col.planned++;
          continue;
        }
        const out = session.replaceInText(x.v, kept, (type) => st.methods[type]);
        kept.forEach((h) => { const rep = session.replaceText(h.type, h.value, st.methods[h.type]); session.remember(h.value, rep); });
        if (out === x.v) continue;
        changes.push({ sheet: col.sheet, r: x.r, c: col.c, before: x.v, after: out === "" ? null : { value: out, type: "string" }, type: kept[0].type });
        col.planned++;
      }
    }
    session.freezeKnown();
    return { session, changes, values, typeCounts };
  }

  // ── wykonanie ───────────────────────────────────────────────────────
  function payloadOf(after) {
    if (!after || after.value == null) return null;
    if (after.type === "date") return { v: after.value, t: "d" };
    if (after.type === "number") return { v: after.value, t: "n" };
    return { v: String(after.value), t: "s" };
  }
  function cellObjOf(prev, after) {
    const cell = {};
    if (prev && prev.s) cell.s = prev.s;
    if (prev && prev.z && after.type !== "string") cell.z = prev.z;
    if (prev && prev.c) cell.c = prev.c;
    if (after.type === "date") { cell.t = "d"; cell.v = after.value; } else if (after.type === "number") { cell.t = "n"; cell.v = after.value; } else { cell.t = "s"; cell.v = String(after.value); }
    return cell;
  }

  // opts.restore — przywracanie z klucza: bez czyszczenia śladów przy zapisie (to są znowu dane oryginalne).
  function applyInPlace(plan, opts = {}) {
    if (typeof swbEditGuard === "function" && !swbEditGuard()) return false;
    if (typeof undoCloseStep === "function") undoCloseStep();
    const bySheet = new Map();
    plan.changes.forEach((ch) => { if (!bySheet.has(ch.sheet)) bySheet.set(ch.sheet, []); bySheet.get(ch.sheet).push(ch); });
    const group = [];
    const rowMap = typeof undoRowMap === "function" ? undoRowMap() : new Map();
    for (const [sheetName, list] of bySheet) {
      const ws = workbook.Sheets[sheetName];
      const isCur = sheetName === currentSheetName;
      const step = { workbook, sheet: sheetName, changes: [] };
      for (const ch of list) {
        const ref = XLSX.utils.encode_cell({ r: ch.r, c: ch.c });
        const col = ch.c - (isCur ? currentStartCol : 0);
        if (typeof undoSnapshot === "function") step.changes.push({ cellRef: ref, rowIndex0: ch.r, col, absCol: ch.c, snap: undoSnapshot(sheetName, ref, ch.r, col, isCur ? rowMap : null) });
        const prev = ws[ref];
        if (!ch.after) {
          if (prev && (prev.s || prev.z || prev.c)) { const stub = { t: "z" }; if (prev.s) stub.s = prev.s; if (prev.z) stub.z = prev.z; if (prev.c) stub.c = prev.c; ws[ref] = stub; } else delete ws[ref];
        } else {
          ws[ref] = cellObjOf(prev, ch.after);
        }
        recordPendingEdit(sheetName, ref, payloadOf(ch.after));
        if (isCur) {
          const row = rowMap.get(ch.r);
          if (row && col >= 0) {
            const v = ch.after ? ch.after.value : null;
            if (Array.isArray(row.values)) row.values[col] = v;
            if (Array.isArray(row.rawValues)) row.rawValues[col] = v;
            if (Array.isArray(row.display)) row.display[col] = v == null ? "" : toDisplay(v);
            if (typeof cfmtRefreshDisplay === "function") cfmtRefreshDisplay(row, col);
          }
        }
        if (typeof swbFormulaTouched === "function") swbFormulaTouched(sheetName, ch.r, ch.c);
      }
      if (step.changes.length) group.push(step);
    }
    bumpSheetDataStamp();
    if (typeof swbNotesInvalidate === "function") swbNotesInvalidate();
    // Drugi przebieg „w miejscu” przed zapisem: czyszczenie pierwszego zostaje w łańcuchu (prev) —
    // dawniej nowe nadpisywało stare i ślady 1. przebiegu (notatki, wykresy, pivot) zostawały w pliku.
    const prevCleanup = swbAnonPendingCleanup && swbAnonPendingCleanup.workbook === workbook ? swbAnonPendingCleanup : null;
    const cleanup = opts.restore ? null : { session: plan.session, opts: { ...st.traces }, workbook, headerCells: st.scan && st.scan.headerCells, protectedCells: st.scan && st.scan.protectedCells, prev: prevCleanup };
    if (group.length && typeof undoStack !== "undefined") {
      const cur = group.find((s) => s.sheet === currentSheetName) || group[0];
      const all = [cur, ...group.filter((s) => s !== cur)];
      undoStack.push({ workbook, sheet: cur.sheet, group: all, changes: all.flatMap((s) => s.changes), anonCleanup: swbAnonPendingCleanup });
      if (undoStack.length > UNDO_LIMIT) undoStack.shift();
      redoStack = [];
      updateUndoButtons();
    }
    swbAnonPendingCleanup = cleanup;
    setDirtyState(true);
    renderActiveTable();
    scheduleViewRefresh({ analyses: true, filterBadge: true });
    log(`Anonimizacja: ${opts.restore ? "przywrócono z klucza" : "zmieniono"} ${plan.changes.length} komorek w ${bySheet.size} arkuszach`, "success");
    return true;
  }

  // traces/noClean — warianty do testów (szukanie, który krok czyszczenia psuje plik)
  async function buildCopyBytes(plan, traces = st.traces, noClean = false) {
    const merged = {};
    Object.entries(pendingEdits || {}).forEach(([sn, cells]) => { merged[sn] = { ...cells }; });
    plan.changes.forEach((ch) => {
      const ref = XLSX.utils.encode_cell({ r: ch.r, c: ch.c });
      if (!merged[ch.sheet]) merged[ch.sheet] = {};
      const prev = merged[ch.sheet][ref];
      let payload = payloadOf(ch.after);
      if (prev && (prev.fmt || "note" in prev)) {
        const keep = {};
        if (prev.fmt) keep.fmt = prev.fmt;
        if ("note" in prev) keep.note = prev.note;
        payload = payload === null ? { clear: true, ...keep } : { ...payload, ...keep };
      }
      merged[ch.sheet][ref] = payload;
    });
    const patched = await buildPatchedXlsx(originalFileBytes, merged, buildSheetFormulaMaps());
    if (noClean) return { bytes: patched, report: null };
    return swbAnonCleanZip(patched, plan.session, traces, { personMethod: st.methods.person, headerCells: st.scan.headerCells, protectedCells: st.scan.protectedCells });
  }

  async function saveCopy(bytes) {
    const base = typeof workbookBaseName === "function" ? workbookBaseName() : "plik";
    const ext = currentFileName && /\.xlsm$/i.test(currentFileName) ? "xlsm" : "xlsx";
    return saveFile(bytes, `${base}${tx("fileSuffix")}.${ext}`, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fsaSaveTypes(ext));
  }
  async function saveFile(bytes, name, mime, types) {
    if (canFSA) {
      try {
        const handle = await window.showSaveFilePicker({ suggestedName: name, types });
        const w = await handle.createWritable();
        await w.write(new Blob([bytes]));
        await w.close();
        return handle.name || name;
      } catch (err) {
        if (err && err.name === "AbortError") return null;
      }
    }
    const blob = new Blob([bytes], { type: mime });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url; a.download = name;
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 30000);
    return name;
  }

  // Wpięte w buildOutputBytes (ui-controls.js): czyszczenie śladów przy zapisie po anonimizacji „w miejscu”.
  async function cleanOnSave(bytes) {
    const c = swbAnonPendingCleanup;
    if (!c || c.workbook !== workbook || typeof swbAnonCleanZip !== "function") return bytes;
    try {
      const chain = [];
      for (let x = c; x && x.workbook === workbook && chain.length < 20; x = x.prev) chain.unshift(x);
      let out = null;
      for (const x of chain) {
        out = await swbAnonCleanZip(out ? out.bytes : bytes, x.session, x.opts, { personMethod: x.session.methods.person, headerCells: x.headerCells, protectedCells: x.protectedCells });
      }
      if (out.report.leaks.length) setTimeout(() => toast(tx("toastLeakOnSave", { n: out.report.leaks.length }), "warning"), 400);
      swbAnon.lastReport = out.report;
      return out.bytes;
    } catch (err) {
      log(`Anonimizacja: czyszczenie śladów przy zapisie nie powiodło się (${err && err.message})`, "warning");
      setTimeout(() => toast(tx("toastCleanFailed"), "error"), 400);
      return bytes;
    }
  }

  // ── UI ──────────────────────────────────────────────────────────────
  let modal = null;
  const R = {};
  function methodOptions(type) {
    const list = ["keep", "pseudo", "fake", "mask", "generalize", "redact", "remove", "hash"];
    if (type === "dob" || type === "custom" || type === "other") list.push("shift");
    return list;
  }
  function select(options, value, onChange) {
    const s = el("select", "anon-select");
    options.forEach(([v, label]) => { const o = el("option", null, label); o.value = v; s.appendChild(o); });
    s.value = value;
    s.addEventListener("change", () => onChange(s.value));
    return s;
  }
  function section(title, cls) {
    const sec = el("section", `anon-sec${cls ? " " + cls : ""}`);
    const h = el("h3", "anon-sec-title", title);
    sec.appendChild(h);
    return sec;
  }
  function check(label, checked, onChange, hint) {
    const lab = el("label", "anon-check");
    const inp = el("input");
    inp.type = "checkbox"; inp.checked = !!checked;
    inp.addEventListener("change", () => onChange(inp.checked));
    lab.append(inp, el("span", null, label));
    if (hint) { lab.setAttribute("data-hint", ""); lab.setAttribute("data-hint-pl", hint); lab.setAttribute("data-hint-en", hint); }
    return lab;
  }

  function buildModal() {
    modal = el("div", "modal hidden anon-modal");
    modal.id = "anonModal";
    modal.setAttribute("role", "dialog");
    modal.setAttribute("aria-modal", "true");
    modal.setAttribute("aria-labelledby", "anonTitle");
    const content = el("div", "modal-content anon-content");
    const head = el("div", "modal-head");
    R.title = el("div", "modal-title"); R.title.id = "anonTitle";
    const close = el("button", "modal-close");
    close.type = "button"; close.id = "anonClose";
    close.innerHTML = '<svg viewBox="0 0 24 24" width="16" height="16" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><path d="M18 6 6 18M6 6l12 12"/></svg>';
    close.addEventListener("click", closeModal);
    head.append(R.title, close);
    R.body = el("div", "anon-body");
    const actions = el("div", "modal-actions anon-actions");
    R.status = el("div", "anon-status");
    R.status.setAttribute("aria-live", "polite");
    R.cancel = el("button", "btn ghost"); R.cancel.type = "button"; R.cancel.addEventListener("click", closeModal);
    R.run = el("button", "btn"); R.run.type = "button"; R.run.id = "anonRunBtn"; R.run.addEventListener("click", run);
    actions.append(R.status, R.cancel, R.run);
    content.append(head, R.body, actions);
    modal.appendChild(content);
    modal.addEventListener("click", (e) => { if (e.target === modal) closeModal(); });
    modal.addEventListener("keydown", (e) => { if (e.key === "Escape") { e.preventDefault(); closeModal(); } });
    document.body.appendChild(modal);
  }

  // ── Etap 3: profile ────────────────────────────────────────────────
  const T = () => SwbAnonTools;
  function loadProfiles() {
    try { const a = JSON.parse(localStorage.getItem(PROFILES_KEY) || "[]"); return Array.isArray(a) ? a.map((p) => T().sanitizeProfile(p)).filter((p) => p && p.name) : []; } catch { return []; }
  }
  function saveProfiles(list) {
    try { localStorage.setItem(PROFILES_KEY, JSON.stringify(list)); } catch { /* bez pamięci */ }
  }
  const presetName = (p) => tx(`pr_${p.id.replace(/-/g, "_")}`);
  // Profil z bieżących ustawień okna (bez hasła). Typy kolumn zapisywane po nazwie nagłówka.
  function currentAsProfile(name) {
    const columns = { ...st.headerModes };
    if (st.scan) st.scan.cols.forEach((col) => { const m = st.colModes[colKey(col)]; if (m && col.header) columns[T().fold(col.header)] = m; });
    return T().sanitizeProfile({ name, methods: st.methods, traces: st.traces, rules: st.rules, output: st.output, scope: st.scope, dateGeneralize: st.dateGeneralize, everything: st.everything, columns, command: st.command });
  }
  function applyProfile(p) {
    st.methods = { ...SwbAnon.DEFAULT_METHODS, ...p.methods };
    st.traces = { ...SWB_ANON_ZIP_DEFAULTS, ...p.traces };
    if (p.rules) st.rules = { terms: "", exceptions: "", patterns: "", ...p.rules };
    if (p.output) st.output = p.output;
    if (p.scope && p.scope !== "selection") st.scope = p.scope;
    st.dateGeneralize = p.dateGeneralize || null;
    st.everything = !!p.everything;
    st.headerModes = { ...(p.columns || {}) };
    st.colModes = {};
    savePrefs();
  }
  function profileById(id) {
    if (!id) return null;
    if (id.startsWith("b:")) { const b = T().PRESETS.find((x) => x.id === id.slice(2)); return b ? { ...b, name: presetName(b) } : null; }
    const own = loadProfiles();
    const i = Number(id.slice(2));
    return own[i] || null;
  }
  function onProfilePick(id) {
    st.profileId = id;
    const p = profileById(id);
    if (!p) { render(); return; }
    applyProfile(p);
    if (p.command) st.command = p.command;
    st.cmdResult = null;
    if (st.scan) st.plan = buildPlan();
    render();
    toast(tx("profileApplied", { name: p.name }), "success");
  }
  function saveAsProfile() {
    const name = String(window.prompt(tx("profileNamePrompt"), "") || "").trim().slice(0, 80);
    if (!name) return;
    const list = loadProfiles().filter((p) => p.name !== name);
    list.push(currentAsProfile(name));
    saveProfiles(list);
    st.profileId = `u:${list.length - 1}`;
    render();
    toast(tx("profileSaved", { name }), "success");
  }
  function deleteProfile() {
    if (!st.profileId.startsWith("u:")) return;
    const list = loadProfiles();
    const i = Number(st.profileId.slice(2));
    const p = list[i];
    if (!p) return;
    list.splice(i, 1);
    saveProfiles(list);
    st.profileId = "";
    render();
    toast(tx("profileDeleted", { name: p.name }), "success");
  }
  async function exportProfiles() {
    const list = loadProfiles();
    if (!list.length) { toast(tx("profileExportNone"), "warning"); return; }
    const text = T().exportProfiles(list);
    await saveFile(new TextEncoder().encode(text), "profile-anonimizacji.json", "application/json", [{ description: "JSON", accept: { "application/json": [".json"] } }]);
  }
  function importProfiles() {
    const inp = document.createElement("input");
    inp.type = "file"; inp.accept = ".json,application/json";
    inp.addEventListener("change", async () => {
      const f = inp.files && inp.files[0];
      if (!f) return;
      try {
        const { profiles } = T().importProfiles(await f.text());
        const list = loadProfiles();
        profiles.forEach((p) => { const k = list.findIndex((x) => x.name === p.name); if (k >= 0) list[k] = p; else list.push(p); });
        saveProfiles(list);
        render();
        toast(tx("profileImported", { n: fmtN(profiles.length) }), "success");
      } catch { toast(tx("profileImportBad"), "warning"); }
    });
    inp.click();
  }

  // ── Etap 3: polecenie tekstem ─────────────────────────────────────
  function applyCommand() {
    const text = st.command.trim();
    if (!text) { toast(tx("commandEmpty"), "warning"); return; }
    const headers = st.scan ? st.scan.cols.filter((c) => c.header).map((c) => ({ key: colKey(c), header: c.header })) : (currentHeaders || []).map((h, i) => ({ key: `h:${i}`, header: h })).filter((h) => h.header);
    const r = T().parseCommand(text, { headers });
    Object.assign(st.methods, r.methods);
    Object.assign(st.traces, r.traces);
    if (r.output) st.output = r.output;
    if (r.scope && r.scope !== st.scope) { st.scope = r.scope; st.scan = null; st.plan = null; }
    if (r.dateGeneralize) st.dateGeneralize = r.dateGeneralize;
    r.columns.forEach((c) => {
      // z komendy przed skanem (klucz „h:i”) — po nazwie nagłówka, jak profil
      if (st.scan && !c.key.startsWith("h:")) st.colModes[c.key] = c.mode; else st.headerModes[T().fold(c.header)] = c.mode;
      if (c.method && c.mode === "custom") st.methods.custom = c.method;
    });
    st.cmdResult = r;
    st.profileId = "";
    savePrefs();
    if (st.scan) st.plan = buildPlan();
    render();
  }
  // Zrozumiane polecenie → krótka lista („PESEL → Maskuj; Osoba → Pseudonim; daty do miesiąca”).
  function describeCommand(r) {
    const byMethod = new Map();
    const parts = [];
    r.understood.forEach((u) => {
      if (u.kind === "method") { if (!byMethod.has(u.method)) byMethod.set(u.method, []); if (!byMethod.get(u.method).includes(u.type)) byMethod.get(u.method).push(u.type); }
      else if (u.kind === "trace") parts.push(`${tx(`tr_${u.trace}`)} → ${tx(`trv_${u.value}`)}`);
      else if (u.kind === "column") parts.push(tx("cmdColumn", { h: u.header, m: u.mode === "skip" ? tx("colSkip") : u.mode === "custom" ? tx(`m_${u.method}`) : tx(`t_${u.mode}`) }));
      else if (u.kind === "scope") parts.push(tx("cmdScope", { v: tx(`scope${u.scope[0].toUpperCase()}${u.scope.slice(1)}`) }));
      else if (u.kind === "output") parts.push(tx("cmdOutput", { v: tx(u.output === "copy" ? "outCopy" : "outInPlace").split(/[—(]/)[0].trim() }));
    });
    const methodParts = [...byMethod].map(([m, types]) => `${types.length > 6 ? tx("cmdRest", { n: types.length }) : types.map((t) => tx(`t_${t}`)).join(", ")} → ${tx(`m_${m}`)}`);
    if (r.dateGeneralize) methodParts.push(tx(r.dateGeneralize === "month" ? "cmdDateMonth" : "cmdDateYear"));
    return [...methodParts, ...parts];
  }
  function renderProfileSection() {
    const sec = section(tx("profileSec"), "anon-profile-sec");
    const row = el("div", "anon-profile-row");
    const sel = el("select", "anon-select anon-profile-select");
    sel.id = "anonProfile";
    sel.setAttribute("aria-label", tx("profile"));
    const o0 = el("option", null, tx("profileCurrent")); o0.value = ""; sel.appendChild(o0);
    const g1 = el("optgroup"); g1.label = tx("profileBuiltin");
    T().PRESETS.forEach((p) => { const o = el("option", null, presetName(p)); o.value = `b:${p.id}`; o.title = tx(`prd_${p.id.replace(/-/g, "_")}`); g1.appendChild(o); });
    sel.appendChild(g1);
    const own = loadProfiles();
    if (own.length) {
      const g2 = el("optgroup"); g2.label = tx("profileOwn");
      own.forEach((p, i) => { const o = el("option", null, p.name); o.value = `u:${i}`; g2.appendChild(o); });
      sel.appendChild(g2);
    }
    sel.value = st.profileId;
    sel.addEventListener("change", () => onProfilePick(sel.value));
    const mk = (label, fn, id) => { const b = el("button", "btn btn-sm ghost", label); b.type = "button"; if (id) b.id = id; b.addEventListener("click", fn); return b; };
    const btns = el("div", "anon-profile-btns");
    btns.append(mk(tx("profileSave"), saveAsProfile, "anonProfileSave"));
    if (st.profileId.startsWith("u:")) btns.append(mk(tx("profileDelete"), deleteProfile, "anonProfileDelete"));
    btns.append(mk(tx("profileExport"), exportProfiles, "anonProfileExport"), mk(tx("profileImport"), importProfiles, "anonProfileImport"));
    row.append(sel);
    sec.append(row, btns);
    const pid = st.profileId.startsWith("b:") ? st.profileId.slice(2).replace(/-/g, "_") : null;
    sec.appendChild(el("p", "field-note", pid ? tx(`prd_${pid}`) : tx("profileHint")));
    // polecenie
    const cmd = el("div", "anon-cmd");
    const lab = el("label", "anon-sub", tx("command")); lab.htmlFor = "anonCommand";
    const inp = el("input", "anon-cmd-input"); inp.type = "text"; inp.id = "anonCommand"; inp.placeholder = tx("commandPh"); inp.value = st.command; inp.autocomplete = "off"; inp.enterKeyHint = "done";
    inp.addEventListener("input", () => { st.command = inp.value; });
    inp.addEventListener("keydown", (e) => { if (e.key === "Enter") { e.preventDefault(); applyCommand(); } });
    const go = el("button", "btn btn-sm", tx("commandApply")); go.type = "button"; go.id = "anonCommandApply"; go.addEventListener("click", applyCommand);
    const line = el("div", "anon-cmd-line"); line.append(inp, go);
    cmd.append(lab, line);
    if (st.cmdResult) {
      const r = st.cmdResult;
      const box = el("div", "anon-cmd-result"); box.id = "anonCommandResult"; box.setAttribute("aria-live", "polite");
      const parts = describeCommand(r);
      if (parts.length) {
        box.appendChild(el("span", "anon-cmd-head", tx("cmdUnderstood")));
        const ul = el("ul", "anon-cmd-list");
        parts.forEach((t) => ul.appendChild(el("li", null, t)));
        box.appendChild(ul);
      }
      if (r.unknown.length) box.appendChild(el("p", parts.length ? "anon-warn" : "anon-warn", parts.length ? tx("cmdUnknown", { list: r.unknown.map((x) => `„${x}”`).join(", ") }) : tx("cmdNothing")));
      cmd.appendChild(box);
    }
    sec.appendChild(cmd);
    return sec;
  }

  // ── Etap 4: raport ryzyka ─────────────────────────────────────────
  // Dla każdego arkusza/tabeli: kolumny z cechami pośrednimi → k-anonimowość PO zamianach.
  function computeRisk(plan) {
    if (!st.scan || !plan) return [];
    const after = new Map(plan.changes.map((ch) => [`${ch.sheet}|${ch.r}|${ch.c}`, ch.after ? ch.after.value : null]));
    const groups = new Map();
    st.scan.cols.forEach((col) => {
      const mode = colMode(col);
      if (mode === "skip") return;
      const type = mode !== "auto" ? mode : col.profile.type;
      if (!T().isQuasiColumn(col.header, type)) return;
      const k = `${col.sheet}|${col.table || ""}`;
      if (!groups.has(k)) groups.set(k, { sheet: col.sheet, table: col.table, cols: [] });
      groups.get(k).cols.push(col);
    });
    const out = [];
    for (const g of groups.values()) {
      const rowsSet = new Set();
      const vals = g.cols.map((col) => { const m = new Map(); col.cells.forEach((x) => { if (!x.above) { m.set(x.r, x.v); rowsSet.add(x.r); } }); return m; });
      if (rowsSet.size < 5) continue;
      const records = [...rowsSet].map((r) => g.cols.map((col, i) => { const k = `${col.sheet}|${r}|${col.c}`; return after.has(k) ? after.get(k) : vals[i].get(r) ?? null; }));
      const names = g.cols.map((c) => c.header || c.letter);
      out.push({ where: `${st.scan.sheets > 1 ? `${g.sheet}${g.table ? " › " : ""}` : ""}${g.table || (st.scan.sheets > 1 ? "" : g.sheet)}`, ...T().riskReport(records, names) });
    }
    return out;
  }
  function renderRisk(plan) {
    const sec = section(tx("risk"), "anon-risk");
    sec.id = "anonRisk";
    const rs = computeRisk(plan);
    if (!rs.length) { sec.appendChild(el("p", "field-note", tx("riskNone"))); return sec; }
    const show = (v) => (v instanceof Date ? toDisplay(v) : v === "" ? "∅" : String(v));
    rs.forEach((r) => {
      const cols = r.columns.join(" + ");
      if (r.unique) {
        sec.appendChild(el("p", "anon-warn", tx("riskBad", { where: r.where, u: fmtN(r.unique), n: fmtN(r.total), p: fmtN(Math.round((r.unique / r.total) * 1000) / 10), cols })));
        const best = r.drop[0];
        if (best && best.unique < r.unique) sec.appendChild(el("p", "field-note", tx("riskDrop", { c: best.name, u: fmtN(best.unique) })));
        if (r.examples.length) sec.appendChild(el("p", "field-note anon-risk-ex", tx("riskExamples", { list: r.examples.slice(0, 3).map((e) => `(${e.map(show).join(" · ")})`).join(", ") })));
      } else if (r.small) {
        sec.appendChild(el("p", "anon-warn", tx("riskSmall", { where: r.where, s: fmtN(r.small), k: r.smallK, cols })));
      } else {
        sec.appendChild(el("p", "anon-ok", tx("riskOk", { where: r.where, cols, k: r.minK, n: fmtN(r.total) })));
      }
    });
    return sec;
  }

  // ── Etap 4: klucz odwracania ──────────────────────────────────────
  function keyProblem() {
    if (!st.keyOn) return null;
    if (st.keyPass.length < 8) return "rkWeak";
    if (st.keyPass !== st.keyPass2) return "rkMismatch";
    return null;
  }
  async function saveReverseKey(plan) {
    try {
      const base = typeof workbookBaseName === "function" ? workbookBaseName() : "plik";
      const text = await T().encryptReverseKey(plan.changes.map((ch) => ({ sheet: ch.sheet, r: ch.r, c: ch.c, before: ch.before, after: ch.after ? ch.after.value : null })), st.keyPass, { file: currentFileName || base });
      const name = await saveFile(new TextEncoder().encode(text), `${base}${tx("rkSuffix")}.json`, "application/json", [{ description: "JSON", accept: { "application/json": [".json"] } }]);
      if (name) toast(tx("rkSaved", { name, n: fmtN(plan.changes.length) }), "success");
      return name;
    } catch (err) {
      toast(tx("rkFailed", { msg: (err && err.message) || "?" }), "error");
      return null;
    }
  }
  function renderKeyOption(out) {
    const box = el("div", "anon-key-opt");
    box.appendChild(check(tx("rkOpt"), st.keyOn, (v) => { st.keyOn = v; renderResults(); }, tx("rkHint")));
    if (st.keyOn) {
      const mkPass = (label, k, id) => {
        const f = el("label", "field");
        f.appendChild(el("span", null, tx(label)));
        const i = el("input"); i.type = "password"; i.autocomplete = "new-password"; i.id = id; i.value = st[k];
        i.addEventListener("input", () => { st[k] = i.value; });
        f.appendChild(i);
        return f;
      };
      const grid = el("div", "anon-key-pass");
      grid.append(mkPass("rkPass", "keyPass", "anonKeyPass"), mkPass("rkPass2", "keyPass2", "anonKeyPass2"));
      box.append(grid, el("p", "field-note", tx("rkHint")));
    }
    out.appendChild(box);
  }

  // ── Etap 4: przywracanie z klucza ─────────────────────────────────
  let restoreModal = null;
  function openRestore() {
    if (!workbook) { toast(tx("noFile"), "warning"); return; }
    if (!restoreModal) {
      restoreModal = el("div", "modal hidden anon-modal anon-restore-modal");
      restoreModal.id = "anonRestoreModal";
      restoreModal.setAttribute("role", "dialog"); restoreModal.setAttribute("aria-modal", "true"); restoreModal.setAttribute("aria-labelledby", "anonRestoreTitle");
      restoreModal.addEventListener("click", (e) => { if (e.target === restoreModal) restoreModal.classList.add("hidden"); });
      restoreModal.addEventListener("keydown", (e) => { if (e.key === "Escape") { e.preventDefault(); restoreModal.classList.add("hidden"); } });
      document.body.appendChild(restoreModal);
    }
    const content = el("div", "modal-content anon-content");
    const head = el("div", "modal-head");
    const title = el("div", "modal-title", tx("restoreTitle")); title.id = "anonRestoreTitle";
    const close = el("button", "modal-close"); close.type = "button"; close.setAttribute("aria-label", tx("cancel"));
    close.innerHTML = '<svg viewBox="0 0 24 24" width="16" height="16" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><path d="M18 6 6 18M6 6l12 12"/></svg>';
    close.addEventListener("click", () => restoreModal.classList.add("hidden"));
    head.append(title, close);
    const body = el("div", "anon-body");
    body.appendChild(el("p", "field-note", tx("restoreIntro")));
    const fField = el("label", "field"); fField.appendChild(el("span", null, tx("restoreFile")));
    const fInp = el("input"); fInp.type = "file"; fInp.accept = ".json,application/json"; fInp.id = "anonRestoreFile"; fField.appendChild(fInp);
    const pField = el("label", "field"); pField.appendChild(el("span", null, tx("restorePass")));
    const pInp = el("input"); pInp.type = "password"; pInp.autocomplete = "off"; pInp.id = "anonRestorePass"; pField.appendChild(pInp);
    const result = el("div", "anon-restore-result"); result.id = "anonRestoreResult"; result.setAttribute("aria-live", "polite");
    body.append(fField, pField, result);
    const actions = el("div", "modal-actions anon-actions");
    const cancel = el("button", "btn ghost", tx("cancel")); cancel.type = "button"; cancel.addEventListener("click", () => restoreModal.classList.add("hidden"));
    const go = el("button", "btn", tx("restoreRun")); go.type = "button"; go.id = "anonRestoreRun";
    go.addEventListener("click", async () => {
      result.replaceChildren();
      const f = fInp.files && fInp.files[0];
      if (!f) { result.appendChild(el("p", "anon-warn", tx("restoreNoFile"))); return; }
      go.disabled = true;
      try {
        const data = await T().decryptReverseKey(await f.text(), pInp.value);
        pInp.value = "";
        result.append(...restoreFromKey(data));
      } catch (err) {
        const code = err && err.message;
        result.appendChild(el("p", "anon-warn", tx(code === "badPassword" ? "restoreBadPassword" : code === "noCrypto" ? "restoreNoCrypto" : "restoreBadFile")));
      } finally { go.disabled = false; }
    });
    pInp.addEventListener("keydown", (e) => { if (e.key === "Enter") { e.preventDefault(); go.click(); } });
    actions.append(cancel, go);
    content.append(head, body, actions);
    restoreModal.replaceChildren(content);
    restoreModal.classList.remove("hidden");
  }
  // Zwraca elementy raportu. Przywraca tylko komórki, które wciąż mają wartość z anonimizacji.
  function restoreFromKey(data) {
    const nodes = [];
    const changes = [];
    const skipped = [];
    data.changes.forEach((ch) => {
      const ref = XLSX.utils.encode_cell({ r: ch.r, c: ch.c });
      // nazwa arkusza z pliku klucza (niezaufana): tylko arkusze tego skoroszytu — „__proto__” czy
      // „constructor” dałyby Object.prototype i zapis do niego (jak znalezisko audytu z 2026-11-06)
      const own = Object.prototype.hasOwnProperty.call(workbook.Sheets, ch.sheet) && (workbook.SheetNames || []).includes(ch.sheet);
      const ws = own ? workbook.Sheets[ch.sheet] : null;
      if (!ws || ch.r > 1048575 || ch.c > 16383) { skipped.push(`${ch.sheet}!${ref}`); return; }
      const cell = ws[ref];
      const cur = cell && cell.t !== "z" ? cell.v : null;
      if (!T().sameValue(cur, ch.after)) { skipped.push(`${ch.sheet}!${ref}`); return; }
      if (T().sameValue(cur, ch.before)) return;
      const v = ch.before;
      changes.push({ sheet: ch.sheet, r: ch.r, c: ch.c, before: cur, after: v == null ? null : { value: v, type: v instanceof Date ? "date" : typeof v === "number" ? "number" : "string" } });
    });
    if (data.file && currentFileName && T().fold(data.file) !== T().fold(currentFileName)) nodes.push(el("p", "field-note", tx("restoreOtherFile", { f: data.file, d: (data.created || "").slice(0, 10) })));
    if (!changes.length) { nodes.push(el("p", "anon-warn", tx("restoreNothing"))); return nodes; }
    if (!applyInPlace({ changes, session: null }, { restore: true })) return nodes;
    nodes.push(el("p", "anon-ok", tx("restoreDone", { n: fmtN(changes.length) })));
    if (skipped.length) nodes.push(el("p", "field-note", tx("restoreSkipped", { n: fmtN(skipped.length), list: skipped.slice(0, 12).join(", ") + (skipped.length > 12 ? ` +${skipped.length - 12}` : "") })));
    toast(tx("restoreDone", { n: fmtN(changes.length) }), "success");
    return nodes;
  }

  // ── Etap 5: AI ✦ w wolnym tekście ─────────────────────────────────
  // Teksty z kolumn ze zdaniami (bez kolumn pominiętych i „całej komórki” — tam i tak wszystko znika).
  function freeTexts() {
    if (!st.scan) return { cols: [], texts: [] };
    const cols = st.scan.cols.filter((col) => {
      const mode = colMode(col);
      if (mode === "skip" || (mode !== "auto" && mode !== "custom")) return false;
      return SwbAnonTools.isFreeTextColumn(col.header, col.cells.slice(0, 300).map((x) => x.v));
    });
    const seen = new Set();
    const texts = [];
    cols.forEach((col) => col.cells.forEach((x) => {
      if (typeof x.v !== "string" || x.v.trim().length < 12 || !/\p{L}{3}/u.test(x.v)) return;
      const k = SwbAnon.norm(x.v);
      if (seen.has(k)) return;
      seen.add(k);
      texts.push(x.v);
    }));
    // najpierw teksty z wyrazem z wielkiej litery w środku zdania (podejrzane o nazwisko)
    const score = (t) => (/[^.!?]\s+\p{Lu}\p{Ll}{2,}/u.test(t) ? 1 : 0);
    texts.sort((a, b) => score(b) - score(a));
    return { cols, texts: texts.slice(0, 120) };
  }
  let aiAbort = null;
  async function runAi() {
    if (typeof LocalAI === "undefined") return;
    if (st.aiBusy) { if (aiAbort) aiAbort.abort(); return; }
    const { texts } = freeTexts();
    if (!texts.length) return;
    if (!(await LocalAI.ensure({ reason: tx("aiReason") }))) return;
    st.aiBusy = true; st.aiMsg = "";
    aiAbort = typeof AbortController !== "undefined" ? new AbortController() : null;
    renderResults();
    const before = st.aiTerms.size;
    let found = 0;
    const BATCH = 6;
    try {
      for (let i = 0; i < texts.length; i += BATCH) {
        if (aiAbort && aiAbort.signal.aborted) break;
        R.status.textContent = tx("aiProgress", { done: fmtN(i), total: fmtN(texts.length) });
        const chunk = texts.slice(i, i + BATCH);
        const { value, problem } = await LocalAI.json({ ...SwbAnonTools.aiPrompt(chunk), maxTokens: 300, signal: aiAbort && aiAbort.signal, validate: (j) => (j && (Array.isArray(j.znalezione) || Array.isArray(j.found)) ? "" : "no-list") });
        if (!value) { if (i === 0) throw new Error(problem || "invalid"); continue; }
        const exc = new Set(lines(st.rules.exceptions).map(SwbAnon.norm));
        SwbAnonTools.parseAiAnswer(chunk, value).forEach((h) => {
          if (exc.has(SwbAnon.norm(h.value))) return;
          found++;
          const k = `${h.type}|${SwbAnon.norm(h.value)}`;
          if (!st.aiTerms.has(k)) st.aiTerms.set(k, { type: h.type, value: h.value });
        });
      }
      const fresh = st.aiTerms.size - before;
      st.aiMsg = found ? tx("aiDone", { n: fmtN(found), m: fmtN(fresh) }) : tx("aiNone");
    } catch (err) {
      if (!(err && err.name === "AbortError")) st.aiMsg = tx("aiFailed", { msg: (err && err.message) || "?" });
    } finally {
      st.aiBusy = false; aiAbort = null;
      R.status.textContent = "";
      st.aiStamp++;
      st.plan = buildPlan();
      renderResults();
    }
  }
  function renderAiSection() {
    if (typeof LocalAI === "undefined" || !st.scan) return null;
    const { cols, texts } = freeTexts();
    if (!texts.length && !st.aiTerms.size) return null;
    const sec = section(tx("aiSec"), "anon-ai");
    sec.id = "anonAi";
    sec.appendChild(el("p", "field-note", tx("aiIntro", { cols: cols.map((c) => `„${c.header || c.letter}”`).slice(0, 4).join(", ") + (cols.length > 4 ? "…" : ""), n: fmtN(texts.length) })));
    const b = el("button", "btn btn-sm", st.aiBusy ? tx("aiStop") : st.aiTerms.size ? tx("aiRunAgain") : tx("aiRun"));
    b.type = "button"; b.id = "anonAiRun";
    b.addEventListener("click", runAi);
    sec.appendChild(b);
    if (st.aiMsg) { const m = el("p", /nie odpowiedział|didn't answer/.test(st.aiMsg) ? "anon-warn" : "field-note", st.aiMsg); m.id = "anonAiMsg"; sec.appendChild(m); }
    return sec;
  }

  function render() {
    if (!modal) return;
    R.title.textContent = tx("title");
    R.cancel.textContent = tx("cancel");
    R.body.replaceChildren();
    R.body.appendChild(renderProfileSection());

    // Zakres
    const sScope = section(tx("scope"));
    const seg = el("div", "anon-seg");
    seg.setAttribute("role", "radiogroup");
    [["selection", "scopeSelection"], ["columns", "scopeColumns"], ["sheet", "scopeSheet"], ["file", "scopeFile"]].forEach(([v, k]) => {
      const b = el("button", "anon-seg-btn", tx(k));
      b.type = "button"; b.dataset.scope = v;
      b.setAttribute("role", "radio");
      b.setAttribute("aria-checked", String(st.scope === v));
      b.addEventListener("click", () => { st.scope = v; st.scan = null; savePrefs(); render(); });
      seg.appendChild(b);
    });
    sScope.appendChild(seg);
    if (st.scope === "columns") {
      const box = el("div", "anon-cols-pick");
      box.appendChild(el("div", "anon-sub", tx("pickColumns")));
      const list = el("div", "anon-chips");
      (currentHeaders || []).forEach((h, i) => {
        if (typeof cfmtIsRealCol === "function" && !cfmtIsRealCol(i)) return;
        const lab = el("label", "anon-chip");
        const inp = el("input"); inp.type = "checkbox"; inp.checked = st.columns.has(i);
        inp.addEventListener("change", () => { if (inp.checked) st.columns.add(i); else st.columns.delete(i); st.scan = null; renderResults(); });
        lab.append(inp, el("span", null, h || `(${i + 1})`));
        list.appendChild(lab);
      });
      box.appendChild(list);
      sScope.appendChild(box);
    }
    if (st.scope === "sheet" || st.scope === "columns") sScope.appendChild(check(tx("filteredOnly"), st.filteredOnly, (v) => { st.filteredOnly = v; st.scan = null; renderResults(); }));
    sScope.appendChild(check(tx("everything"), st.everything, (v) => { st.everything = v; st.colModes = {}; if (st.scan) refreshPlan(); else renderResults(); }));
    const keyField = el("label", "field anon-key");
    keyField.appendChild(el("span", null, tx("key")));
    const keyInp = el("input"); keyInp.type = "password"; keyInp.autocomplete = "off"; keyInp.value = st.key; keyInp.id = "anonKey";
    keyInp.addEventListener("input", () => { st.key = keyInp.value; if (st.scan) refreshPlan(); });
    keyField.appendChild(keyInp);
    keyField.appendChild(el("span", "field-note", tx("keyHint")));
    sScope.appendChild(keyField);
    const scanBtn = el("button", "btn full", st.scan ? tx("rescan") : tx("scan"));
    scanBtn.type = "button"; scanBtn.id = "anonScanBtn";
    scanBtn.addEventListener("click", doScan);
    sScope.appendChild(scanBtn);
    R.body.appendChild(sScope);

    R.results = el("div", "anon-results");
    R.body.appendChild(R.results);
    renderResults();
  }

  function refreshPlan() {
    st.plan = buildPlan();
    renderResults();
  }

  async function doScan() {
    if (!workbook) { toast(tx("noFile"), "warning"); return; }
    R.status.textContent = tx("scanning", { done: 0, total: "…" });
    const res = await scan((done, total) => { R.status.textContent = tx("scanning", { done: fmtN(done), total: fmtN(total) }); });
    R.status.textContent = "";
    if (res.err) { toast(tx(res.err), "warning"); return; }
    st.scan = res;
    st.colModes = {};
    st.excluded = new Set();
    st.included = new Set();
    st.plan = buildPlan();
    render();
    const first = R.results.querySelector(".anon-sec");
    if (first && first.scrollIntoView) first.scrollIntoView({ block: "start", behavior: "smooth" });
  }

  function renderResults() {
    if (!R.results) return;
    R.results.replaceChildren();
    const plan = st.plan;
    R.run.textContent = st.output === "copy" ? tx("runCopy") : tx("runInPlace");
    R.run.disabled = !plan || !plan.changes.length;
    R.run.classList.toggle("hidden", !plan);
    if (!st.scan || !plan) return;

    const sum = el("p", "anon-summary", tx("found", {
      cells: fmtN(st.scan.cells), sheets: tx("sheetsN", { n: st.scan.sheets }),
      changes: fmtN(plan.changes.length), values: fmtN(plan.values.size),
    }));
    R.results.appendChild(sum);
    if (st.scan.formulas) R.results.appendChild(el("p", "field-note", tx("formulaSkipped", { n: fmtN(st.scan.formulas) })));
    if (!plan.values.size) R.results.appendChild(el("p", "anon-empty", tx("nothing")));

    // typy + metoda
    const types = Object.keys(plan.typeCounts).sort((a, b) => SwbAnon.TYPES.indexOf(a) - SwbAnon.TYPES.indexOf(b));
    if (types.length) {
      const sec = section(tx("types"));
      const grid = el("div", "anon-types");
      types.forEach((type) => {
        const row = el("div", "anon-type-row");
        row.dataset.type = type;
        const name = el("span", "anon-type-name");
        name.append(el("span", `anon-tag t-${type}`, tx(`t_${type}`)));
        // przykład pseudonimu wg rodzaju: „Adres 001”, „PESEL 001” (nie zawsze „Osoba 001”)
        const mLabel = (m) => (m === "pseudo" && type !== "email" ? tx("m_pseudo").replace(/\(.*\)/, `(${labels()[type]} 001)`) : tx(`m_${m}`));
        const sel = select(methodOptions(type).map((m) => [m, mLabel(m)]),st.methods[type] || "redact", (v) => { st.methods[type] = v; savePrefs(); refreshPlan(); });
        sel.setAttribute("aria-label", tx(`t_${type}`));
        const cnt = el("span", "anon-count", `${fmtN(plan.typeCounts[type])}×`);
        cnt.style.textAlign = "right";
        row.append(name, sel, cnt);
        grid.appendChild(row);
      });
      sec.appendChild(grid);
      // formuły budujące formy nazwiska („Nazwisko Imię”, „I. Nazwisko”) z pseudonimu „Osoba 001” robią „001 Osoba”
      if (plan.typeCounts.person && st.scan.formulas && st.methods.person !== "fake") sec.appendChild(el("p", "field-note", tx("personFormulaHint")));
      R.results.appendChild(sec);
    }

    // kolumny
    const candidates = st.scan.cols.filter((c) => c.profile.type || c.cells.some((x) => x.hits.length) || st.colModes[colKey(c)] || st.everything);
    // kolumny bez żadnej zmiany (np. same formuły) chowamy za przyciskiem — inaczej lista puchnie od zer
    const cols = st.showAllCols ? st.scan.cols : candidates.filter((c) => c.planned > 0 || st.colModes[colKey(c)]);
    const hiddenCols = st.scan.cols.length - cols.length;
    if (cols.length || hiddenCols) {
      const sec = section(tx("columns"));
      const list = el("div", "anon-col-list");
      const multiSheet = st.scan.sheets > 1;
      cols.forEach((col) => {
        const row = el("div", "anon-col-row");
        const name = el("div", "anon-col-name");
        const where = `${multiSheet ? `${col.sheet} › ` : ""}${col.letter}${col.table ? ` · ${col.table}` : ""}`;
        name.appendChild(el("span", "anon-col-letter", where));
        name.appendChild(el("span", "anon-col-head", col.header || "—"));
        const opts = [["auto", tx("colAuto")], ...SwbAnon.TYPES.filter((t) => t !== "other").map((t) => [t, `${tx("wholeCell")}${tx(`t_${t}`)}`]), ["skip", tx("colSkip")]];
        const sel = select(opts, colMode(col), (v) => { st.colModes[colKey(col)] = v; refreshPlan(); });
        sel.setAttribute("aria-label", col.header || col.letter);
        const n = el("span", "anon-count", fmtN(col.planned || 0));
        row.append(name, sel, n);
        list.appendChild(row);
      });
      sec.appendChild(list);
      if (hiddenCols > 0 || st.showAllCols) {
        const more = el("button", "btn btn-sm ghost anon-show-cols", st.showAllCols ? tx("colsLess") : tx("colsMore", { n: fmtN(hiddenCols) }));
        more.type = "button";
        more.addEventListener("click", () => { st.showAllCols = !st.showAllCols; renderResults(); });
        sec.appendChild(more);
      }
      R.results.appendChild(sec);
    }

    // wartości (ręczne korekty)
    if (plan.values.size) {
      const sec = section(tx("values"), "anon-values-sec");
      const bar = el("div", "anon-values-bar");
      const search = el("input", "anon-search");
      search.type = "search"; search.placeholder = tx("search"); search.value = st.search;
      const allB = el("button", "btn btn-sm ghost", tx("all")); allB.type = "button";
      const noneB = el("button", "btn btn-sm ghost", tx("none")); noneB.type = "button";
      bar.append(search, allB, noneB);
      const list = el("div", "anon-values");
      const drawValues = () => {
        list.replaceChildren();
        const q = SwbAnon.norm(st.search);
        const items = [...plan.values.values()].filter((v) => !q || SwbAnon.norm(v.value instanceof Date ? v.value.toISOString() : v.value).includes(q) || SwbAnon.norm(tx(`t_${v.type}`)).includes(q))
          .sort((a, b) => (a.conf === "low") - (b.conf === "low") || b.count - a.count);
        const lows = items.filter((v) => v.conf === "low").length;
        const ais = items.filter((v) => v.conf === "ai").length;
        if (lows) list.appendChild(el("div", "anon-low-hint", tx("lowHint", { n: fmtN(lows) })));
        if (ais) list.appendChild(el("div", "anon-low-hint anon-ai-hint", tx("aiHint", { n: fmtN(ais) })));
        // najpierw do decyzji („?”, „✦”), potem reszta
        const pend = (v) => v.conf === "low" || v.conf === "ai";
        items.sort((a, b) => pend(b) - pend(a) || b.count - a.count);
        items.slice(0, 200).forEach((v) => {
          const lab = el("label", "anon-value");
          const isOn = (x) => !st.excluded.has(x.key) && (!pend(x) || st.included.has(x.key));
          const inp = el("input"); inp.type = "checkbox"; inp.checked = isOn(v);
          inp.addEventListener("change", () => {
            if (inp.checked) { st.excluded.delete(v.key); if (pend(v)) st.included.add(v.key); } else { st.excluded.add(v.key); st.included.delete(v.key); }
            refreshPlanKeepValues();
          });
          if (v.conf === "low") lab.classList.add("is-low");
          if (v.conf === "ai") lab.classList.add("is-ai");
          const txt = el("span", "anon-value-text", v.value instanceof Date ? toDisplay(v.value) : String(v.value));
          lab.append(inp, el("span", `anon-tag t-${v.type}`, tx(`t_${v.type}`)), txt, el("span", "anon-count", `${fmtN(v.count)}×`));
          if (v.conf === "medium") lab.classList.add("is-medium");
          list.appendChild(lab);
        });
        if (items.length > 200) list.appendChild(el("div", "anon-more", tx("more", { n: fmtN(items.length - 200) })));
        list._items = items;
      };
      search.addEventListener("input", () => { st.search = search.value; drawValues(); });
      allB.addEventListener("click", () => { (list._items || []).forEach((v) => { st.excluded.delete(v.key); if (v.conf === "low" || v.conf === "ai") st.included.add(v.key); }); refreshPlanKeepValues(); drawValues(); });
      noneB.addEventListener("click", () => { (list._items || []).forEach((v) => { st.excluded.add(v.key); st.included.delete(v.key); }); refreshPlanKeepValues(); drawValues(); });
      drawValues();
      sec.append(bar, list);
      R.results.appendChild(sec);
    }

    const aiSec = renderAiSection();
    if (aiSec) R.results.appendChild(aiSec);

    // własne reguły
    const rules = el("details", "anon-sec anon-rules");
    rules.appendChild(el("summary", "anon-sec-title", tx("rules")));
    [["terms", "terms"], ["exceptions", "exceptions"], ["patterns", "patterns"]].forEach(([k, label]) => {
      const f = el("label", "field");
      f.appendChild(el("span", null, tx(label)));
      const ta = el("textarea"); ta.rows = 2; ta.value = st.rules[k]; ta.dataset.rule = k;
      ta.addEventListener("input", () => { st.rules[k] = ta.value; savePrefs(); });
      f.appendChild(ta);
      rules.appendChild(f);
    });
    rules.appendChild(el("p", "field-note", tx("rulesNote")));
    if (lines(st.rules.terms).length || lines(st.rules.exceptions).length || lines(st.rules.patterns).length) rules.open = true;
    R.results.appendChild(rules);

    // podgląd
    if (plan.changes.length) {
      const sec = section(tx("preview"));
      sec.appendChild(renderPreview(plan));
      R.results.appendChild(sec);
    }
    if (plan.changes.length) R.results.appendChild(renderRisk(plan));

    // wynik + ślady
    const out = section(tx("output"));
    const radios = el("div", "anon-radios");
    [["copy", "outCopy"], ["inplace", "outInPlace"]].forEach(([v, k]) => {
      const lab = el("label", "anon-check");
      const inp = el("input"); inp.type = "radio"; inp.name = "anonOut"; inp.value = v; inp.checked = st.output === v;
      inp.addEventListener("change", () => { st.output = v; savePrefs(); R.run.textContent = v === "copy" ? tx("runCopy") : tx("runInPlace"); });
      lab.append(inp, el("span", null, tx(k)));
      radios.appendChild(lab);
    });
    out.appendChild(radios);
    renderKeyOption(out);
    out.appendChild(el("div", "anon-sub", tx("traces")));
    const tr = el("div", "anon-traces");
    tr.appendChild(check(tx("trMeta"), st.traces.metadata, (v) => { st.traces.metadata = v; savePrefs(); }));
    const notesRow = el("div", "anon-trace-row");
    notesRow.append(el("span", null, tx("trNotes")), select([["anonymize", tx("trNotesAnon")], ["remove", tx("trNotesRemove")], ["keep", tx("trNotesKeep")]], st.traces.notes, (v) => { st.traces.notes = v; savePrefs(); }));
    const hfRow = el("div", "anon-trace-row");
    hfRow.append(el("span", null, tx("trHF")), select([["anonymize", tx("trNotesAnon")], ["clear", tx("trHFClear")], ["keep", tx("trNotesKeep")]], st.traces.headersFooters, (v) => { st.traces.headersFooters = v; savePrefs(); }));
    tr.append(notesRow, hfRow);
    [["formulaCache", "trFormula"], ["charts", "trCharts"], ["pivots", "trPivots"], ["links", "trLinks"], ["sharedStrings", "trStrings"]].forEach(([k, label]) => {
      tr.appendChild(check(tx(label), st.traces[k], (v) => { st.traces[k] = v; savePrefs(); }));
    });
    out.appendChild(tr);
    R.results.appendChild(out);
    if (R.report) R.results.appendChild(R.report);
  }

  // Zmiana zaznaczenia wartości: przelicz plan, ale nie przerysowuj listy (fokus, przewinięcie).
  function refreshPlanKeepValues() {
    st.plan = buildPlan();
    const prev = R.results.querySelector(".anon-preview");
    if (prev) prev.replaceWith(renderPreview(st.plan));
    const risk = R.results.querySelector("#anonRisk");
    if (risk && st.plan.changes.length) risk.replaceWith(renderRisk(st.plan));
    const sum = R.results.querySelector(".anon-summary");
    if (sum) sum.textContent = tx("found", { cells: fmtN(st.scan.cells), sheets: tx("sheetsN", { n: st.scan.sheets }), changes: fmtN(st.plan.changes.length), values: fmtN(st.plan.values.size) });
    R.run.disabled = !st.plan.changes.length;
  }

  function showVal(v) {
    if (v == null || v === "") return "∅";
    return v instanceof Date ? toDisplay(v) : String(v);
  }
  function renderPreview(plan) {
    const box = el("div", "anon-preview edit-preview");
    const pairs = new Map();
    plan.changes.forEach((c) => {
      const from = showVal(c.before);
      const to = showVal(c.after ? c.after.value : null);
      const k = `${from}\u0000${to}`;
      const p = pairs.get(k);
      if (p) p.n++; else pairs.set(k, { from, to, n: 1 });
    });
    [...pairs.values()].slice(0, 14).forEach((p) => {
      const row = el("div", "edit-preview-row");
      row.append(el("span", "from", p.from), el("span", null, "→"), el("span", "to", p.to), el("span", "n", `${p.n}×`));
      box.appendChild(row);
    });
    if (pairs.size > 14) box.appendChild(el("div", "unify-more", tx("more", { n: fmtN(pairs.size - 14) })));
    return box;
  }

  // Wycieki: pełne wartości i dokładne miejsca — żeby dało się ocenić, czy to dane osobowe,
  // czy np. nazwa zakresu albo nagłówek (fałszywy alarm).
  function renderLeaks(leaks, structural) {
    const box = el("div", "anon-leaks-box");
    const places = leaks.reduce((n, e) => n + e.items.length, 0);
    if (!structural) box.appendChild(el("p", "anon-warn", tx(st.lastScope === "file" ? "leakFoundFile" : "leakFound", { n: fmtN(places) })));
    const q = (v) => `„${v}”`;
    leaks.forEach((e) => {
      const block = el("div", "anon-leak");
      let head;
      if (e.kind === "cells") head = tx("lkCells", { s: e.sheet });
      else if (e.kind === "sheetOther") head = tx("lkSheetOther", { s: e.sheet });
      else if (e.kind === "sheetName") head = tx("lkSheetName");
      else if (e.kind === "definedName") head = tx("lkDefinedName");
      else if (e.kind === "tableName") head = tx("lkTableName");
      else if (e.kind === "tableColumn") head = tx("lkTableColumn", { t: e.label });
      else if (e.kind === "appTitles") head = tx("lkAppTitles");
      else if (e.label === "definedNameValue") head = tx("lkDefinedValue");
      else head = tx("lkPart", { p: partLabel(e.label) });
      block.appendChild(el("div", "anon-leak-head", head));
      const ul = el("ul", "anon-leaks");
      e.items.slice(0, 40).forEach((it) => {
        const li = el("li");
        li.appendChild(el("span", "anon-leak-term", q(it.term)));
        if (it.refs && it.refs.length) {
          const refs = it.refs.slice(0, 12).map((r) => `${r.ref}${r.header ? ` (${tx("lkHeader")})` : r.formula ? ` (${tx("lkFormula")})` : ""}`).join(", ");
          li.appendChild(el("span", "anon-leak-refs", ` — ${refs}${it.refs.length > 12 ? ` +${it.refs.length - 12}` : ""}`));
        }
        ul.appendChild(li);
      });
      if (e.items.length > 40) ul.appendChild(el("li", null, tx("more", { n: fmtN(e.items.length - 40) })));
      block.appendChild(ul);
      const hint = structural ? null : { sheetName: "lkHintNames", definedName: "lkHintNames", tableName: "lkHintNames", tableColumn: "lkHintNames", appTitles: "lkHintNames" }[e.kind]
        || (e.kind === "cells" && e.items.every((it) => it.refs.every((r) => r.header)) ? "lkHintHeaders" : null);
      if (hint) block.appendChild(el("p", "field-note", tx(hint)));
      box.appendChild(block);
    });
    return box;
  }
  function partLabel(p) {
    if (/^docProps\/core/i.test(p)) return tx("plCore");
    if (/^xl\/charts\//i.test(p)) return tx("plChart");
    if (/^xl\/drawings\//i.test(p)) return tx("plDrawing");
    if (/^xl\/pivot/i.test(p)) return tx("plPivot");
    if (/^xl\/(comments|threadedComments)/i.test(p)) return tx("plNotes");
    if (/sharedStrings/i.test(p)) return tx("sharedStrings");
    return p;
  }

  // ── kontrola powiązań: listy wyboru (Sprawdzanie poprawności → Lista) ──────
  // Komórka, która przed anonimizacją miała wartość z listy, musi ją mieć także po — lista
  // źródłowa (zakres, nazwa, kolumna tabeli albo lista wpisana) zmienia się tą samą zamianą.
  function checkLinks(plan) {
    if (typeof dvRulesBySheet === "undefined" || !dvRulesBySheet || !plan) return null;
    const changed = new Map(plan.changes.map((ch) => [`${ch.sheet}!${ch.r},${ch.c}`, ch.after ? ch.after.value : null]));
    const orig = (sheet, r, c) => { const cell = workbook.Sheets[sheet]?.[XLSX.utils.encode_cell({ r, c })]; return cell ? (cell.w ?? cell.v) : null; };
    const after = (sheet, r, c) => { const k = `${sheet}!${r},${c}`; return changed.has(k) ? changed.get(k) : orig(sheet, r, c); };
    const norm = (v) => (v == null ? "" : String(v instanceof Date ? toDisplay(v) : v).trim().toLowerCase());
    const rangeOf = (f) => {
      let ref = String(f || "").trim().replace(/^=/, "");
      if (typeof dvNameRefs !== "undefined" && dvNameRefs && dvNameRefs.has(ref)) ref = dvNameRefs.get(ref);
      if (typeof dvResolveStructuredToA1 === "function") ref = dvResolveStructuredToA1(ref) || ref;
      return ref;
    };
    const readSet = (ref, sheet, valueAt) => {
      let sn = sheet, part = ref;
      const bang = ref.lastIndexOf("!");
      if (bang >= 0) { sn = ref.slice(0, bang).replace(/^'(.*)'$/, "$1").replace(/''/g, "'"); part = ref.slice(bang + 1); }
      let rg;
      try { rg = XLSX.utils.decode_range(part.replace(/\$/g, "")); } catch { return null; }
      if (!workbook.Sheets[sn]) return null;
      const set = new Set();
      for (let r = rg.s.r; r <= rg.e.r && set.size < 5000; r++) for (let c = rg.s.c; c <= rg.e.c; c++) { const v = norm(valueAt(sn, r, c)); if (v) set.add(v); }
      return set;
    };
    let rules = 0, cells = 0;
    const broken = [];
    for (const [sheet, list] of dvRulesBySheet) {
      for (const rule of list) {
        if (rule.type !== "list") continue;
        const f = String(rule.formula1Raw || "").trim();
        let before, now;
        if (/^".*"$/.test(f)) {
          const items = f.slice(1, -1).split(",").map((x) => x.trim()).filter(Boolean);
          before = new Set(items.map(norm));
          now = new Set(items.map((x) => norm(plan.session.rewriteString(x))));
        } else {
          const ref = rangeOf(f);
          before = readSet(ref, sheet, orig);
          now = readSet(ref, sheet, after);
        }
        if (!before || !now) continue;
        rules++;
        for (const rg of rule.ranges) {
          for (let r = rg.s.r; r <= Math.min(rg.e.r, rg.s.r + 5000); r++) for (let c = rg.s.c; c <= rg.e.c; c++) {
            const v0 = norm(orig(sheet, r, c));
            if (!v0 || !before.has(v0)) continue; // pusta albo spoza listy już przed — nie nasza sprawa
            cells++;
            const v1 = norm(after(sheet, r, c));
            if (!now.has(v1)) broken.push(`${sheet}!${XLSX.utils.encode_cell({ r, c })}`);
          }
        }
      }
    }
    return rules ? { rules, cells, broken } : null;
  }
  function renderLinks(lk) {
    const box = el("div", "anon-links");
    if (!lk.broken.length) box.appendChild(el("p", "anon-ok", tx("linksOk", { rules: fmtN(lk.rules), cells: fmtN(lk.cells) })));
    else {
      box.appendChild(el("p", "anon-warn", tx("linksBroken", { n: fmtN(lk.broken.length) })));
      box.appendChild(el("p", "anon-leak-refs", lk.broken.slice(0, 20).join(", ") + (lk.broken.length > 20 ? ` +${lk.broken.length - 20}` : "")));
    }
    box.appendChild(el("p", "field-note", tx(lk.broken.length ? "linksFix" : "linksNote")));
    return box;
  }

  function renderReport({ changed, savedName, report, inPlace, links }) {
    const sec = section(tx("reportTitle"), "anon-report");
    sec.id = "anonReport";
    sec.appendChild(el("p", null, tx("reportCells", { n: fmtN(changed) })));
    if (savedName) sec.appendChild(el("p", null, tx("reportSaved", { name: savedName })));
    if (inPlace) sec.appendChild(el("p", null, tx("reportInPlace")));
    if (report) {
      const cl = Object.entries(report.cleaned).map(([k, n]) => `${tx(`cl_${k}`)} (${fmtN(n)})`);
      if (cl.length) sec.appendChild(el("p", "field-note", tx("cleaned", { list: cl.join(", ") })));
      if (!report.leaks.length) sec.appendChild(el("p", "anon-ok", tx("leakOk", { parts: fmtN(report.partsChecked || 0) })));
      else sec.appendChild(renderLeaks(report.leaks));
      if (report.structural && report.structural.length) {
        // nazwy i nagłówki: nie wliczane do testu, zwinięte — do przejrzenia
        const det = el("details", "anon-struct");
        const n = report.structural.reduce((k, e) => k + e.items.length, 0);
        det.appendChild(el("summary", null, tx("structTitle", { n: fmtN(n) })));
        det.appendChild(el("p", "field-note", tx("structNote")));
        det.appendChild(renderLeaks(report.structural, true));
        sec.appendChild(det);
      }
      report.warnings.forEach((w) => sec.appendChild(el("p", "field-note", tx(`warn${w[0].toUpperCase()}${w.slice(1)}`))));
    }
    if (links) sec.appendChild(renderLinks(links));
    R.report = sec;
    if (R.results) R.results.appendChild(sec);
    sec.scrollIntoView && sec.scrollIntoView({ block: "nearest", behavior: "smooth" });
  }

  async function run() {
    if (!st.plan) { toast(tx("needScan"), "warning"); return; }
    if (!st.plan.changes.length) return;
    const kp = keyProblem();
    if (kp) { toast(tx(kp), "warning"); const f = modal.querySelector("#anonKeyPass"); if (f) f.focus(); return; }
    // świeży plan tuż przed wykonaniem — ten sam, który widać w podglądzie
    const plan = buildPlan();
    st.plan = plan;
    st.lastScope = st.scope;
    R.report = null;
    const links = checkLinks(plan);
    if (st.output === "inplace") {
      if (!applyInPlace(plan)) return;
      toast(tx("toastInPlace", { n: fmtN(plan.changes.length) }), "success");
      if (st.keyOn) await saveReverseKey(plan);
      st.scan = null; st.plan = null;
      render();
      renderReport({ changed: plan.changes.length, inPlace: true, links });
      return;
    }
    if (!originalFileBytes || typeof buildPatchedXlsx !== "function") { toast(tx("noBytes"), "warning"); return; }
    R.run.disabled = true;
    R.status.textContent = tx("working");
    try {
      if (typeof ensureXlsxLibs === "function") await ensureXlsxLibs(false);
      const { bytes, report } = await buildCopyBytes(plan);
      report.partsChecked = report.partsChecked || 0;
      const savedName = await saveCopy(bytes);
      swbAnon.lastBytes = bytes;
      swbAnon.lastReport = report;
      R.status.textContent = "";
      if (savedName === null) { R.run.disabled = false; return; }
      log(`Anonimizacja: zapisano kopie ${savedName} (${plan.changes.length} komorek)`, "success");
      if (st.keyOn) await saveReverseKey(plan);
      renderReport({ changed: plan.changes.length, savedName, report, links });
    } catch (err) {
      R.status.textContent = "";
      toast(`${tx("noBytes")} (${err && err.message})`, "error");
    } finally {
      R.run.disabled = !(st.plan && st.plan.changes.length);
    }
  }

  function open(opts = {}) {
    if (!workbook) { toast(tx("noFile"), "warning"); return; }
    if (!modal) buildModal();
    if (opts.scope) st.scope = opts.scope;
    else if (typeof getSelectionRectangle === "function" && getSelectionRectangle() && getSelectionRectangle().colCount * getSelectionRectangle().rowCount > 1) st.scope = "selection";
    st.scan = null; st.plan = null; st.seed = null; R.report = null;
    if (st.aiWorkbook !== workbook) { st.aiTerms = new Map(); st.aiMsg = ""; st.aiWorkbook = workbook; }
    render();
    modal.classList.remove("hidden");
    // fokus tylko przy klawiaturze — na dotyku zakres wyglądał na „wciśnięty” (obwódka fokusu)
    if (!(window.matchMedia && matchMedia("(pointer: coarse)").matches)) {
      requestAnimationFrame(() => { const b = modal.querySelector(".anon-seg-btn[aria-checked='true']") || modal.querySelector("button"); if (b) b.focus({ focusVisible: false }); });
    }
  }
  function closeModal() {
    if (modal) modal.classList.add("hidden");
  }

  function syncPanel() {
    const t = document.querySelector("#panel-anonymize .panel-title");
    if (t) t.textContent = tx("panelTitle");
    const h = document.getElementById("anonPanelHint");
    if (h) h.textContent = tx("panelHint");
    const b = document.getElementById("anonOpenBtn");
    if (b) b.textContent = tx("open");
    const rb = document.getElementById("anonRestoreBtn");
    if (rb) rb.textContent = tx("restoreOpen");
    if (modal && !modal.classList.contains("hidden")) render();
  }

  loadPrefs();
  const openBtn = document.getElementById("anonOpenBtn");
  if (openBtn) openBtn.addEventListener("click", () => open());
  const restoreBtn = document.getElementById("anonRestoreBtn");
  if (restoreBtn) restoreBtn.addEventListener("click", () => openRestore());
  syncPanel();

  return { open, close: closeModal, scan, buildPlan, buildCopyBytes, state: st, cleanOnSave, syncPanel, openRestore, restoreFromKey, computeRisk, applyCommand, lastReport: null, lastBytes: null };
})();

function swbAnonSyncLanguage() { swbAnon.syncPanel(); }
