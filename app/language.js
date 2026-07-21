// Language dictionaries, translation helpers, and locale-specific UI updates.

const LANG_KEY = "excel-workbench-lang";
let currentLang = localStorage.getItem(LANG_KEY) === "en" ? "en" : "pl";
let BASE_TITLE = document.title || "Sheet Workbench";

const I18N = {
  pl: {
    locale: "pl-PL",
    title: "Sheet Workbench",
    description: "Lokalne przeglądanie, filtrowanie i analiza arkuszy Excel",
    values: "Wartości",
    formulas: "Formuły",
    contains: "Zawiera",
    startsWith: "Zaczyna się",
    equals: "Równa się",
    dateBetween: "Między",
    dateBefore: "DO",
    dateAfter: "OD",
    dateLastN: "Ostatnie N dni",
    allLabel: "Wszystkie",
    any: "dowolnie",
    anyNonEmpty: "komórka niepusta (przynajmniej jedna)",
    allNonEmpty: "komórki niepuste (wszystkie)",
    anyEmpty: "komórka pusta (przynajmniej jedna)",
    allEmpty: "komórki puste (wszystkie)",
    dateAnyNonEmpty: "komórka niepusta (przynajmniej jedna)",
    dateAllNonEmpty: "komórki niepuste (wszystkie)",
    dateAnyEmpty: "komórka pusta (przynajmniej jedna)",
    dateAllEmpty: "komórki puste (wszystkie)",
    sortAsc: "Rosnąco",
    sortDesc: "Malejąco",
    allFunctions: "Wszystkie funkcje",
    noResult: "Bez wyniku",
    withError: "Z błędem",
    online: "Online",
    offline: "Offline",
    panelOpen: "Zamknij filtry",
    panelClosed: "Filtry",
    readingStandard: "Tryb standardowy",
    readingQuick: "Tryb szybkie szukanie",
    quickSearchColumns: "Kolumny",
    quickSearchActionFilter: "Filtruj",
    quickSearchActionHighlight: "Zaznacz",
    quickSearchActionCells: "Podświetl pasujące",
    quickSearchActionFilterCells: "Filtruj + podświetl",
    allColumns: "Wszystkie kolumny",
    selectedColumnsCount: "{count} kolumn",
    choosePreset: "Wybierz preset",
    noSavedPresets: "Brak zapisanych presetów",
    moveUp: "Góra",
    moveDown: "Dół",
    changeDirection: "Zmień kierunek",
    remove: "Usuń",
    defaultSort: "Domyślne sortowanie",
    excelView: "Wymiary z Excela",
    excelViewOn: "Wymiary z Excela: ON",
    classicView: "Widok klasyczny",
    loadingFile: "Wczytywanie pliku...",
    loadingSheet: "Budowanie tabeli...",
    loadingGeneric: "Wczytywanie...",
    skeletonAria: "Ładowanie tabeli",
    statusFileLoaded: "Plik wczytany",
    fileLoaded: "Plik wczytany",
    guideChooseSheet: "Wybierz arkusz i kliknij „Wczytaj arkusz”",
    sheetLoaded: "Arkusz wczytany",
    filtersApplied: "Zastosowano filtry",
    filtersReset: "Reset filtrów",
    firstLoadSheet: "Najpierw wczytaj arkusz",
    loadSheetToPickColumns: "Wczytaj arkusz, żeby wybrac kolumny",
    quickSearchColumnsTitleShort: "Kolumny szybkiego szukania",
    textFilter2ColumnsTitle: "Kolumny filtru tekstowego 2",
    dateFilterColumnsTitle: "Kolumny filtru dat",
    selectMeasuresTitle: "Wybierz miary",
    groupByPickerTitle: "Wybierz kolumny grupowania",
    groupByLimitReached: "Możesz wybrać maksymalnie 3 kolumny grupowania",
    addedSortRule: "Dodano sortowanie do kolejki",
    noSortsToSave: "Brak sortowan do zapisania",
    presetNamePrompt: "Nazwa presetu sortowania:",
    sortPresetSaved: "Zapisano preset sortowania",
    choosePresetToast: "Wybierz preset",
    presetNotFound: "Nie znaleziono presetu",
    sortPresetLoaded: "Wczytano preset sortowania",
    choosePresetToDelete: "Wybierz preset do usuniecia",
    sortPresetDeleted: "Usunieto preset sortowania",
    defaultSortRestored: "Przywrocono domyslne sortowanie",
    webSaveInfo: "Wersja webowa nie nadpisuje pliku. Uzyj Zapisz jako...",
    widthsRestored: "Przywrócono automatyczne szerokości",
    cacheRefresh: "Czyszcze cache i odswiezam aplikacje...",
    refreshingApp: "Odswiezam aplikacje...",
    updateAvailable: "Jest dostepna nowsza wersja aplikacji.",
    updateReady: "Nowa wersja gotowa. Kliknij Aktualizuj.",
    updateNow: "Aktualizuj",
    fileLoadFailed: "Nie udalo sie wczytac pliku",
    noDataForExport: "Brak danych do eksportu",
    csvExported: "Wyeksportowano CSV",
    exportNoColumns: "Zaznacz przynajmniej jedną kolumnę",
    exportColumnFallback: "(kolumna {n})",
    exportReportTitle: "Raport",
    exportRowsMeta: "{count} wierszy",
    exportModalTitle: "Eksport / Raport",
    exportModalSub: "Wybierz kolumny do eksportu (bieżący, przefiltrowany widok)",
    exportSelectAll: "Zaznacz wszystko",
    exportClearAll: "Wyczyść",
    exportActionCsv: "CSV",
    exportActionPrint: "Drukuj / PDF",
    validationNeedValues: "Podaj listę dozwolonych wartości (lub wybierz kolumnę-słownik)",
    validationAllValid: "Wszystko zgodne — 0 wartości spoza listy ({total} wierszy)",
    validationSummaryText: "{bad} z {total} wierszy poza listą · {values} różnych wartości spoza listy",
    validationBadValuesTitle: "Wartości spoza listy:",
    validationPanelTitle: "Walidacja listą",
    validationHintText: "Wskaż kolumnę i listę dozwolonych wartości (wpisaną albo wziętą z innej kolumny). Pokażę, które wiersze mają wartości SPOZA listy — to szybka kontrola jakości danych.",
    validationColumnLabel: "Kolumna do sprawdzenia",
    validationSourceLabel: "Źródło dozwolonych wartości",
    validationSourceListOpt: "Wpisana lista",
    validationSourceColumnOpt: "Inna kolumna (słownik)",
    validationAllowedLabel: "Dozwolone wartości (po jednej w wierszu lub po przecinku)",
    validationDictColumnLabel: "Kolumna-słownik (jej unikalne wartości = dozwolone)",
    validationIgnoreEmptyLabel: "Pomijaj puste komórki",
    validationCaseLabel: "Ignoruj wielkość liter i spacje brzegowe",
    validationCheckLabel: "Sprawdź",
    validationClearLabel: "Wyczyść",
    validationShowOnlyLabel: "Pokaż w tabeli tylko niezgodne wiersze",
    derivedPanelTitle: "Kolumny wyliczane",
    derivedHintText: "Twórz wirtualne kolumny z wyrażeń odwołujących się do innych kolumn przez [Nazwa] (np. [Kwota]*1,23). Funkcją WYSZUKAJ dociągniesz wartość z innego arkusza po kluczu. Kolumny są tylko w widoku — plik zostaje nietknięty.",
    derivedFormulaHead: "Kolumna z wyrażenia",
    derivedNameLabel: "Nazwa kolumny",
    derivedExprLabel: "Wyrażenie",
    derivedAdd: "Dodaj kolumnę",
    derivedUpdate: "Zaktualizuj kolumnę",
    derivedJoinHead: "Dołącz kolumnę z arkusza (WYSZUKAJ)",
    derivedJoinNameLabel: "Nazwa nowej kolumny (opcjonalnie)",
    derivedJoinKeyLabel: "Klucz w tym arkuszu",
    derivedJoinSheetLabel: "Arkusz źródłowy",
    derivedJoinSrcKeyLabel: "Kolumna-klucz w nim",
    derivedJoinSrcRetLabel: "Kolumna do dociągnięcia",
    derivedJoinAdd: "Dołącz kolumnę",
    derivedRecalc: "Przelicz",
    derivedClearAll: "Usuń wszystkie",
    derivedEmpty: "Brak kolumn wyliczanych. Dodaj wyrażenie albo dołącz kolumnę z innego arkusza.",
    derivedDefaultName: "Kolumna",
    derivedEdit: "Edytuj",
    derivedRemove: "Usuń",
    derivedNa: "#N/D",
    derivedErr: "#BŁĄD",
    derivedNeedExpr: "Wpisz wyrażenie kolumny.",
    derivedBadSyntax: "Błąd składni wyrażenia.",
    derivedSyntaxOk: "Składnia OK",
    derivedSyntaxBad: "Błąd składni",
    derivedBadgeSyntax: "błąd składni",
    derivedBadgeErrRows: "{n} bez wyniku",
    derivedJoinNeedAll: "Uzupełnij klucz, arkusz i obie kolumny.",
    noFileToSave: "Brak pliku do zapisu",
    xlsmConfirm: "Plik .xlsm moze utracic makra. Kontynuowac zapis?",
    fileSaved: "Zapisano plik",
    saveFailed: "Nie udało się zapisać pliku",
    saveFellBackToDownload: "Nie udało się zapisać w miejscu — pobrano kopię pliku.",
    savePermissionDenied: "Brak zgody na zapis do pliku",
    saveInPlaceWarn: "Zapis NADPISZE oryginalny plik w miejscu. Funkcja jest nowa — na czas testów zrób wcześniej kopię pliku albo użyj „Zapisz jako…”. Kontynuować nadpisywanie?",
    saveAsPrompt: "Podaj nazwe pliku (xlsx lub xlsm):",
    chooseFileFirst: "Najpierw wybierz plik",
    noSheet: "Brak arkusza",
    xlsxMissingStatus: "Brak biblioteki XLSX",
    xlsxMissingToast: "Brak biblioteki XLSX. Odśwież stronę lub sprawdź połączenie.",
    xlsxMissingEmpty: "Aplikacja nie załadowała silnika arkuszy. Odśwież stronę i sprawdź połączenie z internetem.",
    tableNoData: "Brak danych",
    tableNoResults: "Brak wyników",
    tableNoResultsHint: "Zmień filtry albo wybierz inny arkusz.",
    statusTableRows: "Wierszy: {total} (pokazano: {shown}){mode}{focused}",
    statusTableRowsEmpty: "Wierszy: 0",
    statusFocusedRow: " • rekord {pos}",
    statusLongMode: " • tryb long",
    formulaEditBlocked: "Edycja formuł jest zablokowana",
    dvRejected: "Wartość spoza listy dozwolonych (reguła z pliku) — wpisz jedną z podpowiedzi",
    dvWarning: "Uwaga: wartość spoza listy dozwolonych — zapisuję mimo to",
    editWideOnly: "Edycja komórek działa tylko w trybie szerokim (wide).",
    editBlockedRow: "Tej komórki nie można edytować (wiersz pochodny lub podnagłówek).",
    editCellAria: "Edycja komórki",
    editToolsPanel: "Narzędzia edycji",
    editScopeLabel: "Zakres",
    editColumnLabel: "Kolumna",
    editOpLabel: "Operacja",
    editPatternModeLabel: "Tryb",
    editPatternLabel: "Wzorzec",
    editFindLabel: "Znajdź",
    editReplaceLabel: "Zamień na",
    editReplaceNote: "Wskazówka: datę wpisz tak, jak ją widzisz w tabeli; liczbę — jako wartość surową (kropka dziesiętna, bez spacji i symboli jak „zł” czy „%”).",
    editRegexLabel: "Wyrażenie regularne",
    editCaseLabel: "Wielkość liter",
    editApply: "Zastosuj",
    editScopeColumn: "Kolumna",
    editScopeSelection: "Zaznaczenie",
    editOpPattern: "Zdejmij śmieć wzorcem",
    editOpReplace: "Znajdź i zamień",
    editOpCase: "Zmiana wielkości",
    editOpTrim: "Przytnij / spacje",
    editOpAffix: "Prefiks / sufiks",
    editOpPad: "Wyrównaj długość",
    editOpConvert: "Konwersja typu",
    editPModePattern: "Wzorzec",
    editPModeFuzzy: "Rozmyte",
    editCaseUpper: "WIELKIE",
    editCaseLower: "małe",
    editCaseTitle: "Jak Nazwa Własna",
    editTrimModeLabel: "Tryb",
    editTrimEnds: "Przytnij końce",
    editTrimCollapse: "Zwiń wielokrotne spacje",
    editTrimHard: "Zamień twarde spacje",
    editPrefixLabel: "Prefiks (z przodu)",
    editSuffixLabel: "Sufiks (z tyłu)",
    editPadLenLabel: "Długość",
    editPadCharLabel: "Znak wypełnienia",
    editPadSideLabel: "Strona",
    editPadStart: "Z przodu (z lewej)",
    editPadEnd: "Z tyłu (z prawej)",
    editConvertToLabel: "Na typ",
    editConvertNumber: "Liczba",
    editConvertDate: "Data",
    editConvertText: "Tekst",
    editFilteredOnlyLabel: "Tylko przefiltrowane wiersze",
    editErrNoPattern: "Podaj wzorzec (np. Gr*= albo =#)",
    editErrNoFind: "Podaj tekst do znalezienia",
    editErrBadRegex: "Niepoprawne wyrażenie regularne",
    editErrNoAffix: "Podaj prefiks lub sufiks",
    editErrBadPadLen: "Podaj poprawną długość (liczbę ≥ 1)",
    editErrNoColumn: "Wybierz kolumnę",
    editErrNoSelection: "Najpierw zaznacz zakres komórek",
    editToolApplied: "Zmieniono {count} komórek",
    editToolNoChange: "Brak zmian do zastosowania",
    sortRulesEmpty: "Brak aktywnych sortowań. Kliknij nagłówek tabeli albo dodaj regułę tutaj.",
    sectionHeaderSet: "Ustawiono wiersz nagłówka {row}",
    sectionOutsideLimit: "Ta sekcja nie mieści się w aktualnym limicie wierszy",
    blockOutsideView: "Tego bloku nie widać jeszcze w aktualnym widoku arkusza",
    kpiAboveTable: "Źródło KPI jest nad aktualną tabelą: wiersz {row}",
    columnOutsideView: "Ta kolumna nie mieści się jeszcze w aktualnym widoku tabeli",
    formulaOutsideView: "Ta formuła nie mieści się w aktualnym widoku tabeli",
    wideLongOn: "Włączono Wide-to-Long",
    wideLongOff: "Wrócono do klasycznego widoku",
    longColBlockNum: "Nr bloku",
    longColBlock: "Blok",
    switchToLongTitle: "Przełącz wykryte bloki kolumn na długi widok analityczny",
    backToClassicTitle: "Wróć do klasycznego układu arkusza",
    positiveHeaderRow: "Podaj dodatni numer wiersza nagłówka.",
    invalidHeaderRow: "Wiersz {row} nie wygląda na poprawny nagłówek dla tego arkusza.",
    filteredFor: "Przefiltrowano widok dla: {value}",
    monthlyNoData: "Wczytaj arkusz, aby zobaczyć rozkład miesięczny.",
    monthlyNoDate: "Brak kolumny z datą w aktualnym widoku — ten panel grupuje dane po miesiącach.",
    monthlyDateColumn: "Kolumna z datą",
    monthlyDateColumns: "Kolumny z datą (można kilka)",
    monthlyMetric: "Miara",
    monthlyMetricQ: "Co policzyć?",
    monthlyMetricCount: "Liczba wierszy",
    monthlyMetricOccurrences: "Liczba wystąpień (cykli)",
    monthlyMetricRows: "Liczba wierszy (unikalne)",
    monthlyMetricSum: "Suma",
    monthlyMetricAvg: "Średnia",
    monthlyMetricMin: "Minimum",
    monthlyMetricMax: "Maksimum",
    monthlySplit: "Rozbij słupki na kolumny",
    monthlyMeasureColumn: "Kolumna miary (liczba/data)",
    monthlyMeasureOf: "…z kolumn (można kilka cykli)",
    monthlyMeasureHint: "Wybierz kolumnę liczbową lub duracji (np. „Długość”). Wskaż 2+ kolumny z datą (📅) → policzy ODSTĘP między nimi (np. od→do = czas trzymania).",
    monthlyDateGroup: "Grupuj po miesiącu — wg daty (można kilka)",
    monthlyReadOccurrences: "Liczba wystąpień",
    monthlyReadRows: "Liczba unikalnych wierszy",
    monthlyReadOf: "z „{col}”",
    monthlyReadGap: "z odstępu między: {col}",
    monthlyGapToggle: "Policz odstęp między datami",
    monthlyReadback: "{what} na miesiąc — wg daty: {dates}",
    monthlyWindow: "Zakres",
    monthlyMonthsOption: "{n} mies.",
    monthlyAnchor: "Koniec okna",
    monthlyAnchorData: "Do ostatniej daty",
    monthlyAnchorToday: "Do dziś",
    monthlyTotalRows: "Trafień z datą: {count}",
    freezeHeadersOn: "Zablokowano wiersze nagłówków",
    freezeHeadersOff: "Odblokowano wiersze nagłówków",
    headerDetected: "Wykryto wiersz nagłówka",
    loadSheetToSearch: "Wczytaj arkusz, żeby szukać",
    duplicatedHeaders: "Zdublowane naglowki rozrozniono ({count})",
    themeToggleTitle: "Zmień motyw (jasny / ciemny)",
    themeToggleAria: "Zmień motyw",
    heroGripAria: "Zwiń lub rozwiń nagłówek",
    brandRefreshTitle: "Odśwież aplikację",
    brandRefreshAria: "Odśwież aplikację",
    networkSafety: "Pliki Excel są wczytywane i przetwarzane lokalnie na Twoim urządzeniu.",
    networkOnlineTitle: "Połączenie aktywne. {note}",
    networkOfflineTitle: "Brak połączenia sieciowego. {note}",
    sidebarCloseAria: "Zamknij panel filtrów",
    sidebarOpenAria: "Otwórz panel filtrów",
    sidebarHideTitle: "Schowaj filtry",
    sidebarShowTitle: "Pokaż filtry",
    sidebarHandleLabel: "Wysuń",
    emptyTitle: "Wczytaj plik Excel, aby zacząć",
    emptySub: "wybierz plik poniżej albo przeciągnij go do strefy w panelu bocznym",
    openFileBtn: "Wybierz plik z dysku",
    sampleBtn: "Wypróbuj na przykładowym pliku",
    measureOccurrences: "Liczba rekordów / wystąpień",
    aggregationGroupBy: "Grupuj po",
    aggregationGroupBy2: "Potem po",
    aggregationGroupBy3: "I jeszcze po",
    aggregationNone: "Brak",
    aggregationMeasure: "Mierz",
    aggregationMethod: "Agregacja",
    aggregationMatchText: "Kliknij wynik → filtruj tabelę:",
    aggregationMeasureFilter: "Zawiera w mierze",
    aggregationShowResults: "Pokaż wyników",
    aggregationGroupFilter: "Filtrowanie grup",
    aggregationSearch: "Szukaj:",
    aggregationSearchPlaceholder: "np. Julian...",
    aggregationSearchTable: "Szukaj w tabeli",
    aggregationSource: "Źródło",
    aggregationScope: "Zakres",
    aggregationHeader: "Nagłówek agregacji",
    aggregationAuto: "Auto",
    aggregationClassic: "Widok klasyczny",
    aggregationCurrentView: "Aktualny widok",
    aggregationWholeSheet: "Cały arkusz",
    aggregationCustomNumber: "Własny numer",
    aggregationAll: "Wszystkie",
    aggregationContains: "Zawiera",
    aggregationExact: "Dokładnie",
    aggGroupMode: "Scalanie grup",
    aggGroupExact: "Dokładnie (1:1)",
    aggGroupFuzzy: "Rozmyte (auto)",
    aggGroupPattern: "Wzorzec…",
    aggGroupPatternPlaceholder: "np. Gr*=  ·  =#",
    aggregationCount: "Liczebność",
    aggregationAvg: "Średnia",
    aggregationMedian: "Mediana",
    aggregationMin: "Minimum",
    aggregationMax: "Maksimum",
    aggregationSum: "Suma",
    aggregationDistinct: "Unikalnych",
    aggregationEarliest: "Najwcześniej",
    aggregationLatest: "Najpóźniej",
    aggregationGroups: "Grupy",
    aggregationSourceRows: "Wiersze źródła",
    aggregationMeasuredRows: "Zmierzonych rekordów",
    aggregationMissing: "brak",
    aggregationEmpty: "puste",
    aggregationRecordCount: "Liczba {count}",
    aggregationMeta: "Liczba {count} • średnia {avg} • mediana {median} • zakres {min} -> {max}",
    aggregationDateMeta: "Liczba {count} • najwcześniej {min} • najpóźniej {max}",
    aggregationTextMeta: "Liczba {count} • unikalnych {distinct}",
    aggregationNote: "Aktualne źródło: {source} • zakres: {scope} • nagłówek: {header}{helper} • poziomy grupowania: {depth} • dopasowanie: {match}{having}.",
    aggregationMeasureNote: "Miary: {measures}. Niezgodne dodatkowe miary są automatycznie pomijane dla wybranej agregacji.",
    aggregationHeaderAuto: "auto -> wiersz {row}",
    aggregationHeaderRow: "wiersz {row}",
    aggregationHelper: " (pomocniczy)",
    aggregationHavingValue: " • filtr: > {value}",
    aggregationHavingPercent: " • filtr: > {value}%",
    aggregationHavingAboveValue: "Wartość >",
    aggregationHavingAboveTotal: "% sumy >",
    aggregationHavingAboveMax: "% max >",
    aggregationMeasureSearchPlaceholder: "szukaj...",
    aggregationSearchCount: "{visible} z {total}",
    aggregationResult: "Wynik",
    aggregationTooltipCount: "Ile rekordów lub wystąpień jest w każdej grupie",
    aggregationTooltipAvg: "Średnia wartość, np. średni czas lub średnia kwota",
    aggregationTooltipMedian: "Środkowa wartość: połowa wartości jest mniejsza, połowa większa",
    aggregationTooltipMin: "Najmniejsza wartość w grupie",
    aggregationTooltipMax: "Największa wartość w grupie",
    aggregationTooltipSum: "Łączna suma wszystkich wartości w grupie",
    aggregationTooltipDistinct: "Ile różnych, niepowtarzalnych wartości jest w grupie",
    aggregationTooltipEarliest: "Najstarsza / najwcześniejsza data w grupie",
    aggregationTooltipLatest: "Najnowsza / najpóźniejsza data w grupie",
    aggregationNoData: "Wczytaj arkusz, aby uruchomić agregacje.",
    aggregationNoOptions: "Brak sensownych opcji grupowania lub mierzenia dla aktualnego źródła danych.",
    aggregationNoResults: "Aktualna kombinacja grupowania i mierzenia nie zwróciła żadnych wyników.",
    noInsightFlags: "Brak istotnych flag dla aktualnego pliku.",
    workbookInsightsEmpty: "Wczytaj plik, aby zobaczyć metadane skoroszytu.",
    sheetInsightsEmpty: "Wczytaj arkusz, aby zobaczyć sygnały jakości danych i struktury.",
    sectionsEmpty: "Wczytaj arkusz, aby wykryć sekcje i bloki layoutu.",
    sheetSummaryEmpty: "Wczytaj arkusz, aby zobaczyć szybkie podsumowanie struktury i najważniejszych sygnałów.",
    analysisComputing: "Liczę analizę…",
    durationNoGroup: "Wykryj najpierw powtarzalne bloki kolumn. Ten panel najlepiej działa na arkuszach z cyklami albo seriami podobnych pól.",
    durationNoConfig: "Wykryto bloki, ale nie udało się znaleźć pary typu osoba + od/do albo osoba + długość. Jeśli nagłówek jest nietypowy, moduł próbuje też zgadywać po danych, ale tu to wciąż za mało.",
    durationNoRecords: "Bloki zostały rozpoznane, ale w aktualnym widoku nie ma rekordów z pełnymi danymi czasu dla tej samej wartości.",
    durationRecordsCount: "Rekordy czasu",
    durationAvgTime: "Średni czas",
    durationMedian: "Mediana",
    durationMin: "Min",
    durationMax: "Max",
    durationInProgress: "W toku",
    durationClosed: "Zamknięte",
    durationNoteFiltered: 'Analiza dotyczy aktualnie przefiltrowanego widoku ({visible} z {source} wierszy). Otwarte rekordy bez daty "do" są liczone do dzisiaj.',
    durationNoteFull: 'Analiza dotyczy całego aktualnego widoku arkusza. Otwarte rekordy bez daty "do" są liczone do dzisiaj.',
    durationNoteInferred: " Układ kolumn został częściowo odgadnięty na podstawie danych, bo nagłówek nie był idealny.",
    durationNoteHelper: " Do tej analizy użyto pomocniczo wiersza nagłówka {row}, bo lepiej pasował niż aktualnie wybrany {current}.",
    durationStatusLabel: "Status",
    durationStatusAll: "Wszystkie",
    durationStatusClosed: "Tylko zamknięte",
    durationStatusOpen: "Tylko otwarte",
    durationSortLabel: "Sortuj po",
    durationSortAvg: "Średniej",
    durationSortMedian: "Medianie",
    durationSortCount: "Liczbie rekordów",
    durationSortMax: "Maksimum",
    durationSortMin: "Minimum",
    durationShowLabel: "Pokaż rekordów",
    durationShowAll: "Wszystkie",
    durationViewClassic: "Widok klasyczny",
    durationShowFull: "Pokaż całość",
    durationShownPartial: "Pokazano {shown} z {total} wyników.",
    durationShownAll: "Pokazano wszystkie wyniki: {total}.",
    durationDaysUnit: "{n} dni",
    durationMonthsShort: "mies.",
    durationNone: "brak",
    durationMetaLine: "Średnio {avg} • mediana {median} • rekordy {count} • w toku {open} • zakres {min} -> {max}",
    durationShowInTable: "Pokaż w tabeli",
    secKindTable: "Tabela",
    secKindGroup: "Blok",
    secKindCandidate: "Nagłówek",
    secKindSubheader: "Sekcja",
    secKindLayout: "Układ",
    secTableData: "Tabela danych",
    secFallback: "Sekcja arkusza",
    secMetaHeader: "Nagłówek: wiersz {row} • kolumny {cols}",
    secCurrentHeaderRow: "Aktualny wiersz nagłówka: {row}",
    secPossibleHeaderRow: "Możliwy wiersz nagłówka: {row}",
    secMetaRowCols: "Wiersz {row} • kolumny {cols}",
    secMetaDataRow: "Wiersz danych {row}",
    secSetHeader: "Ustaw nagłówek",
    secJump: "Skocz",
    blkFallbackLabel: "Blok {n}",
    blkSignatureLabel: "Powtarzalny układ: {n} bloków",
    blkSignatureLabelSingle: "Powtarzalny układ kolumn",
    blkSignatureMeta: "{n} bloków po {span} kolumny",
    blkCycleLabel: "Powtarzalny układ: {n} cykli",
    blkCycleMeta: "{n} cykli • {fields} pól w schemacie",
    blkMergedMeta: "{n} bloków po {span} kolumny • {first} -> {last}",
    blkPrefixNote: " • stałe kolumny przed blokami: {cols}",
    blkSpanBadge: "{n} kol.",
    blkColumnsMeta: "Kolumny {range}",
    blkJumpTo: "Skocz do bloku",
    inspSheets: "Arkusze",
    inspHiddenSheets: "Ukryte arkusze",
    inspVeryHidden: "Very hidden",
    inspDefinedNames: "Nazwane zakresy",
    inspMacroFile: "Plik makr .xlsm",
    inspVisibleAllRows: "Widoczne / wszystkie wiersze",
    inspColumns: "Kolumny",
    inspFormulas: "Formuły",
    inspMerges: "Scalenia (zakresy / komórki)",
    inspHiddenColsRows: "Ukryte kolumny / wiersze",
    inspNumericDateCols: "Kolumny liczbowe / datowe",
    inspSparseCols: "Rzadkie kolumny",
    inspLongText: "Długie teksty",
    flagProcessSheet: "Wygląda jak arkusz procesu / SLA",
    flagDuplicateHeaders: "Zdublowane nagłówki: {n}",
    flagDuplicateRows: "Duplikaty wierszy: {n}",
    flagFormulasNoResult: "Formuły bez wyniku: {n}",
    flagComments: "Komentarze: {n}",
    flagLinks: "Linki: {n}",
    flagVeryHiddenSheets: "Są arkusze very hidden",
    kpiCandidates: "Kandydaci KPI",
    kpiSource: "Źródło",
    kpiRowsAboveHeader: "Wiersze nad nagłówkiem {row}",
    kpiRowsAboveDetected: "Wiersze nad wykrytym nagłówkiem {row}",
    kpiTitleAliases: "{label}: {value}\nRównież jako: {aliases}",
    kpiMetaPlain: "{address} • etykieta {labelAddress}",
    kpiMetaAliases: "{address} • etykieta {labelAddress} • również jako: {aliases}{more}",
    kpiShowSource: "Pokaż źródło",
    aggHavingTitle: "Podaj wartość progową (np. 10 oznacza >10)",
    inspChipSections: "Sekcje",
    inspChipBlocks: "Bloki",
    inspChipFlagged: "Kolumny z flagami",
    inspTopSignal: "Top sygnał",
    inspSetHeaderRow: "Ustaw nagłówek: {row}",
    inspBackToClassic: "Wróć do widoku klasycznego",
    inspSwitchToLong: "Przełącz na Wide-to-Long",
    inspJumpToCol: "Skocz do kolumny: {header}",
    colFallback: "Kolumna {n}",
    profMeta: "Kolumna {col} • puste {empty}% • unikalne {unique}",
    profJumpToCol: "Skocz do kolumny",
    profType_pusta: "pusta",
    profType_formuly: "formuły",
    profType_daty: "daty",
    profType_liczby: "liczby",
    profType_tekst: "tekst",
    profType_mixed: "mieszane",
    profFlag_sparse: "rzadka",
    profFlag_mixed: "mieszane",
    profFlag_longText: "długie teksty",
    profFlag_mostlyUnique: "prawie same unikalne",
    profFlag_formulaColumn: "kolumna formuł",
    profFlag_empty: "pusta",
    aggGroupKind_person: "osoba/właściciel",
    aggGroupKind_category: "status/kategoria",
    aggGroupKind_entity: "klient/firma",
    aggGroupKind_place: "miejsce/dział",
    aggGroupKind_item: "produkt/projekt",
    aggGroupKind_time: "czas/data",
    aggGroupKind_id: "identyfikator",
    aggGroupKind_other: "kolumna",
    aggGroupProfileMeta: "{kind} • {unique} unikalnych • {nonEmpty} niepustych",
    choose: "Wybierz",
    hintDefault: "Kliknij",
    loadingSample: "Generuję przykładowy plik...",
    sampleFileName: "Przyklad-obieg-terenow.xlsx",
    sampleLoadFailed: "Nie udało się wygenerować przykładu",
    repeatBlocksEmpty: "Brak wyraźnych powtarzalnych bloków dla aktualnego arkusza. Najlepiej działa na szerokich tabelach z cyklami, etapami albo seriami podobnych kolumn.",
    kpiSummaryEmpty: "Brak wykrytych KPI lub podsumowań dla aktualnego arkusza.",
    kpiNoSummary: "Brak podsumowania KPI.",
    kpiListEmpty: "Nie wykryto wiarygodnych KPI nad aktualną tabelą danych.",
    columnProfilesEmpty: "Wczytaj arkusz, aby zobaczyć profil kolumn i szybkie sygnały problemowości.",
    formulaSummaryFormulas: "Formuły",
    formulaSummaryMissing: "Bez wyniku",
    formulaSummaryErrors: "Z błędem",
    formulaSummaryTop: "Top funkcja",
    formulaSummaryVisible: "Widoczne po filtrze",
    formulaSummaryGroups: "Grupy",
    formulaNoSummary: "Brak podsumowania formuł.",
    formulaNoFormulas: "Aktualny arkusz nie ma wykrytych formuł albo nie został jeszcze wczytany.",
    formulaNoList: "Brak formuł do pokazania dla aktualnego arkusza.",
    formulaNoFilterMatch: "Brak formuł pasujących do bieżącego filtru.",
    formulaSameCount: "{count} takich samych",
    formulaJumpFirst: "Skocz do pierwszej komórki",
    formulaJumpCell: "Skocz do komórki",
    formulaResultLabel: "wynik",
    formulaAddresses: "Adresy",
    formulaOutsideTable: "poza tabelą",
    formulaEmptyResult: "pusty wynik",
    qsAllSheetsLabel: "Wszystkie arkusze",
    globalSearchEmpty: "Brak trafień",
    globalSearchMore: "+{count} więcej w tym arkuszu",
    cellStatsRange: "Zakres",
    cellStatsCount: "Liczba",
    cellStatsSum: "Suma",
    cellStatsAvg: "Średnia",
    cellStatsMin: "Min",
    cellStatsMax: "Maks",
    toolbarCollapse: "Zwiń pasek narzędzi",
    toolbarExpand: "Rozwiń pasek narzędzi",
  },
  en: {
    locale: "en-US",
    title: "Sheet Workbench",
    description: "Local browsing, filtering, and analysis of Excel sheets",
    values: "Values",
    formulas: "Formulas",
    contains: "Contains",
    startsWith: "Starts with",
    equals: "Equals",
    dateBetween: "Between",
    dateBefore: "Before",
    dateAfter: "After",
    dateLastN: "Last N days",
    allLabel: "All",
    any: "any",
    anyNonEmpty: "non-empty cell (at least one)",
    allNonEmpty: "non-empty cells (all)",
    anyEmpty: "empty cell (at least one)",
    allEmpty: "empty cells (all)",
    dateAnyNonEmpty: "non-empty cell (at least one)",
    dateAllNonEmpty: "non-empty cells (all)",
    dateAnyEmpty: "empty cell (at least one)",
    dateAllEmpty: "empty cells (all)",
    allFunctions: "All functions",
    noResult: "No result",
    withError: "With error",
    online: "Online",
    offline: "Offline",
    panelOpen: "Close filters",
    panelClosed: "Filters",
    readingStandard: "Standard mode",
    readingQuick: "Quick search mode",
    quickSearchColumns: "Columns",
    quickSearchActionFilter: "Filter",
    quickSearchActionHighlight: "Highlight",
    quickSearchActionCells: "Highlight matches",
    quickSearchActionFilterCells: "Filter + highlight",
    allColumns: "All columns",
    selectedColumnsCount: "{count} columns",
    choosePreset: "Choose preset",
    noSavedPresets: "No saved presets",
    sortAsc: "Ascending",
    sortDesc: "Descending",
    moveUp: "Up",
    moveDown: "Down",
    changeDirection: "Change direction",
    remove: "Remove",
    defaultSort: "Default sort",
    excelView: "Excel dimensions",
    excelViewOn: "Excel dimensions: ON",
    classicView: "Classic view",
    loadingFile: "Loading file...",
    loadingSheet: "Building table...",
    loadingGeneric: "Loading...",
    skeletonAria: "Loading table",
    statusFileLoaded: "File loaded",
    fileLoaded: "File loaded",
    guideChooseSheet: "Pick a sheet and click “Load sheet”",
    sheetLoaded: "Sheet loaded",
    filtersApplied: "Filters applied",
    filtersReset: "Filters reset",
    firstLoadSheet: "Load a sheet first",
    loadSheetToPickColumns: "Load a sheet to choose columns",
    quickSearchColumnsTitleShort: "Quick search columns",
    textFilter2ColumnsTitle: "Text filter 2 columns",
    dateFilterColumnsTitle: "Date filter columns",
    selectMeasuresTitle: "Select measures",
    groupByPickerTitle: "Select grouping columns",
    groupByLimitReached: "You can select up to 3 grouping columns",
    addedSortRule: "Sort rule added to the queue",
    noSortsToSave: "No sort rules to save",
    presetNamePrompt: "Sort preset name:",
    sortPresetSaved: "Sort preset saved",
    choosePresetToast: "Choose a preset",
    presetNotFound: "Preset not found",
    sortPresetLoaded: "Sort preset loaded",
    choosePresetToDelete: "Choose a preset to delete",
    sortPresetDeleted: "Sort preset deleted",
    defaultSortRestored: "Default sorting restored",
    webSaveInfo: "The web version does not overwrite the original file. Use “Save as...”",
    widthsRestored: "Automatic widths restored",
    cacheRefresh: "Clearing cache and refreshing the app...",
    refreshingApp: "Refreshing the app...",
    updateAvailable: "A newer app version is available.",
    updateReady: "New version ready. Click Update.",
    updateNow: "Update",
    fileLoadFailed: "Failed to load the file",
    noDataForExport: "No data to export",
    csvExported: "CSV exported",
    exportNoColumns: "Select at least one column",
    exportColumnFallback: "(column {n})",
    exportReportTitle: "Report",
    exportRowsMeta: "{count} rows",
    exportModalTitle: "Export / Report",
    exportModalSub: "Choose columns to export (current, filtered view)",
    exportSelectAll: "Select all",
    exportClearAll: "Clear",
    exportActionCsv: "CSV",
    exportActionPrint: "Print / PDF",
    validationNeedValues: "Provide a list of allowed values (or pick a dictionary column)",
    validationAllValid: "All valid — 0 values outside the list ({total} rows)",
    validationSummaryText: "{bad} of {total} rows outside the list · {values} distinct off-list values",
    validationBadValuesTitle: "Values outside the list:",
    validationPanelTitle: "List validation",
    validationHintText: "Pick a column and a list of allowed values (typed or taken from another column). I'll show which rows hold values OUTSIDE the list — a quick data-quality check.",
    validationColumnLabel: "Column to check",
    validationSourceLabel: "Source of allowed values",
    validationSourceListOpt: "Typed list",
    validationSourceColumnOpt: "Another column (dictionary)",
    validationAllowedLabel: "Allowed values (one per line or comma-separated)",
    validationDictColumnLabel: "Dictionary column (its unique values = allowed)",
    validationIgnoreEmptyLabel: "Skip empty cells",
    validationCaseLabel: "Ignore case and surrounding spaces",
    validationCheckLabel: "Check",
    validationClearLabel: "Clear",
    validationShowOnlyLabel: "Show only non-matching rows in the table",
    derivedPanelTitle: "Calculated columns",
    derivedHintText: "Build virtual columns from expressions that reference other columns via [Name] (e.g. [Amount]*1.23). Use LOOKUP to pull a value from another sheet by key. Columns live in the view only — the file stays untouched.",
    derivedFormulaHead: "Column from an expression",
    derivedNameLabel: "Column name",
    derivedExprLabel: "Expression",
    derivedAdd: "Add column",
    derivedUpdate: "Update column",
    derivedJoinHead: "Pull a column from a sheet (LOOKUP)",
    derivedJoinNameLabel: "New column name (optional)",
    derivedJoinKeyLabel: "Key in this sheet",
    derivedJoinSheetLabel: "Source sheet",
    derivedJoinSrcKeyLabel: "Its key column",
    derivedJoinSrcRetLabel: "Column to pull",
    derivedJoinAdd: "Pull column",
    derivedRecalc: "Recalculate",
    derivedClearAll: "Remove all",
    derivedEmpty: "No calculated columns yet. Add an expression or pull a column from another sheet.",
    derivedDefaultName: "Column",
    derivedEdit: "Edit",
    derivedRemove: "Remove",
    derivedNa: "#N/A",
    derivedErr: "#ERR",
    derivedNeedExpr: "Enter a column expression.",
    derivedBadSyntax: "Expression syntax error.",
    derivedSyntaxOk: "Syntax OK",
    derivedSyntaxBad: "Syntax error",
    derivedBadgeSyntax: "syntax error",
    derivedBadgeErrRows: "{n} unmatched",
    derivedJoinNeedAll: "Fill in the key, sheet and both columns.",
    noFileToSave: "No file to save",
    xlsmConfirm: ".xlsm files may lose macros. Continue saving?",
    fileSaved: "File saved",
    saveFailed: "Could not save the file",
    saveFellBackToDownload: "Couldn't save in place — downloaded a copy instead.",
    savePermissionDenied: "Write permission denied",
    saveInPlaceWarn: "This will OVERWRITE the original file in place. The feature is new — while testing, keep a backup copy first or use “Save as…”. Continue overwriting?",
    saveAsPrompt: "Enter a file name (xlsx or xlsm):",
    chooseFileFirst: "Choose a file first",
    noSheet: "No sheet selected",
    xlsxMissingStatus: "XLSX library missing",
    xlsxMissingToast: "XLSX library is missing. Refresh the page or check your connection.",
    xlsxMissingEmpty: "The app did not load the spreadsheet engine. Refresh the page and check your internet connection.",
    tableNoData: "No data",
    tableNoResults: "No results",
    tableNoResultsHint: "Change filters or choose another sheet.",
    statusTableRows: "Rows: {total} (shown: {shown}){mode}{focused}",
    statusTableRowsEmpty: "Rows: 0",
    statusFocusedRow: " • record {pos}",
    statusLongMode: " • long mode",
    formulaEditBlocked: "Formula editing is blocked",
    dvRejected: "Value not in the allowed list (rule from the file) — pick one of the suggestions",
    dvWarning: "Heads up: value not in the allowed list — saving anyway",
    editWideOnly: "Cell editing works only in wide mode.",
    editBlockedRow: "This cell can't be edited (derived row or subheader).",
    editCellAria: "Edit cell",
    editToolsPanel: "Edit tools",
    editScopeLabel: "Scope",
    editColumnLabel: "Column",
    editOpLabel: "Operation",
    editPatternModeLabel: "Mode",
    editPatternLabel: "Pattern",
    editFindLabel: "Find",
    editReplaceLabel: "Replace with",
    editReplaceNote: "Tip: type a date as you see it in the table; type a number as its raw value (dot decimal, no separators or symbols like “zł” or “%”).",
    editRegexLabel: "Regular expression",
    editCaseLabel: "Letter case",
    editApply: "Apply",
    editScopeColumn: "Column",
    editScopeSelection: "Selection",
    editOpPattern: "Strip noise by pattern",
    editOpReplace: "Find & replace",
    editOpCase: "Change case",
    editOpTrim: "Trim / spaces",
    editOpAffix: "Prefix / suffix",
    editOpPad: "Pad to length",
    editOpConvert: "Convert type",
    editPModePattern: "Pattern",
    editPModeFuzzy: "Fuzzy",
    editCaseUpper: "UPPERCASE",
    editCaseLower: "lowercase",
    editCaseTitle: "Title Case",
    editTrimModeLabel: "Mode",
    editTrimEnds: "Trim ends",
    editTrimCollapse: "Collapse repeated spaces",
    editTrimHard: "Normalize hard spaces",
    editPrefixLabel: "Prefix (front)",
    editSuffixLabel: "Suffix (back)",
    editPadLenLabel: "Length",
    editPadCharLabel: "Fill char",
    editPadSideLabel: "Side",
    editPadStart: "Start (left)",
    editPadEnd: "End (right)",
    editConvertToLabel: "To type",
    editConvertNumber: "Number",
    editConvertDate: "Date",
    editConvertText: "Text",
    editFilteredOnlyLabel: "Filtered rows only",
    editErrNoPattern: "Enter a pattern (e.g. Gr*= or =#)",
    editErrNoFind: "Enter text to find",
    editErrBadRegex: "Invalid regular expression",
    editErrNoAffix: "Enter a prefix or suffix",
    editErrBadPadLen: "Enter a valid length (number ≥ 1)",
    editErrNoColumn: "Choose a column",
    editErrNoSelection: "Select a cell range first",
    editToolApplied: "Changed {count} cells",
    editToolNoChange: "No changes to apply",
    sortRulesEmpty: "No active sort rules. Click a table header or add a rule here.",
    sectionHeaderSet: "Header row set to {row}",
    sectionOutsideLimit: "This section is outside the current row limit",
    blockOutsideView: "This block is not visible in the current sheet view yet",
    kpiAboveTable: "The KPI source is above the current table: row {row}",
    columnOutsideView: "This column is not available in the current table view yet",
    formulaOutsideView: "This formula is not available in the current table view",
    wideLongOn: "Wide-to-Long enabled",
    wideLongOff: "Back to classic view",
    longColBlockNum: "Block #",
    longColBlock: "Block",
    switchToLongTitle: "Switch detected column blocks to long analytical view",
    backToClassicTitle: "Return to classic spreadsheet layout",
    positiveHeaderRow: "Enter a positive header row number.",
    invalidHeaderRow: "Row {row} does not look like a valid header for this sheet.",
    filteredFor: "Filtered view for: {value}",
    monthlyNoData: "Load a sheet to see the monthly breakdown.",
    monthlyNoDate: "No date column in the current view — this panel groups data by month.",
    monthlyDateColumn: "Date column",
    monthlyDateColumns: "Date columns (multiple)",
    monthlyMetric: "Metric",
    monthlyMetricQ: "What to show?",
    monthlyMetricCount: "Row count",
    monthlyMetricOccurrences: "Occurrences (cycles)",
    monthlyMetricRows: "Rows (distinct)",
    monthlyMetricSum: "Sum",
    monthlyMetricAvg: "Average",
    monthlyMetricMin: "Minimum",
    monthlyMetricMax: "Maximum",
    monthlySplit: "Split bars by column",
    monthlyMeasureColumn: "Measure column (number/date)",
    monthlyMeasureOf: "…of columns (one per cycle)",
    monthlyMeasureHint: "Pick a number or duration column (e.g. “Length”). Select 2+ date columns (📅) → it computes the GAP between them (e.g. start→end = holding time).",
    monthlyDateGroup: "Group by month — by date (one or more)",
    monthlyReadOccurrences: "Occurrences",
    monthlyReadRows: "Distinct rows",
    monthlyReadOf: "of “{col}”",
    monthlyReadGap: "of the gap between: {col}",
    monthlyGapToggle: "Compute gap between dates",
    monthlyReadback: "{what} per month — by date: {dates}",
    monthlyWindow: "Window",
    monthlyMonthsOption: "{n} mo",
    monthlyAnchor: "Window end",
    monthlyAnchorData: "To latest date",
    monthlyAnchorToday: "To today",
    monthlyTotalRows: "Date hits: {count}",
    freezeHeadersOn: "Header rows locked",
    freezeHeadersOff: "Header rows unlocked",
    headerDetected: "Header row detected",
    loadSheetToSearch: "Load a sheet to search",
    duplicatedHeaders: "Duplicate headers were disambiguated ({count})",
    themeToggleTitle: "Change theme (light / dark)",
    themeToggleAria: "Change theme",
    heroGripAria: "Collapse or expand the header",
    brandRefreshTitle: "Refresh app",
    brandRefreshAria: "Refresh app",
    networkSafety: "Excel files are loaded and processed locally on your device.",
    networkOnlineTitle: "Connection active. {note}",
    networkOfflineTitle: "No network connection. {note}",
    sidebarCloseAria: "Close filters panel",
    sidebarOpenAria: "Open filters panel",
    sidebarHideTitle: "Hide filters",
    sidebarShowTitle: "Show filters",
    sidebarHandleLabel: "Open",
    emptyTitle: "Load an Excel file to get started",
    emptySub: "choose a file below or drag it into the drop zone in the sidebar",
    openFileBtn: "Choose a file from disk",
    sampleBtn: "Try a sample file",
    measureOccurrences: "Record / occurrence count",
    aggregationGroupBy: "Group by",
    aggregationGroupBy2: "Then by",
    aggregationGroupBy3: "And then by",
    aggregationNone: "None",
    aggregationMeasure: "Measure",
    aggregationMethod: "Aggregation",
    aggregationMatchText: "Click result → filter table:",
    aggregationMeasureFilter: "Contains in measure",
    aggregationShowResults: "Show results",
    aggregationGroupFilter: "Group filtering",
    aggregationSearch: "Search:",
    aggregationSearchPlaceholder: "e.g. Julian...",
    aggregationSearchTable: "Search in table",
    aggregationSource: "Source",
    aggregationScope: "Scope",
    aggregationHeader: "Aggregation header",
    aggregationAuto: "Auto",
    aggregationClassic: "Classic view",
    aggregationCurrentView: "Current view",
    aggregationWholeSheet: "Whole sheet",
    aggregationCustomNumber: "Custom number",
    aggregationAll: "All",
    aggregationContains: "Contains",
    aggregationExact: "Exact",
    aggGroupMode: "Group merging",
    aggGroupExact: "Exact (1:1)",
    aggGroupFuzzy: "Fuzzy (auto)",
    aggGroupPattern: "Pattern…",
    aggGroupPatternPlaceholder: "e.g. Gr*=  ·  =#",
    aggregationCount: "Count",
    aggregationAvg: "Average",
    aggregationMedian: "Median",
    aggregationMin: "Minimum",
    aggregationMax: "Maximum",
    aggregationSum: "Sum",
    aggregationDistinct: "Distinct",
    aggregationEarliest: "Earliest",
    aggregationLatest: "Latest",
    aggregationGroups: "Groups",
    aggregationSourceRows: "Source rows",
    aggregationMeasuredRows: "Measured records",
    aggregationMissing: "missing",
    aggregationEmpty: "empty",
    aggregationRecordCount: "Count {count}",
    aggregationMeta: "Count {count} • average {avg} • median {median} • range {min} -> {max}",
    aggregationDateMeta: "Count {count} • earliest {min} • latest {max}",
    aggregationTextMeta: "Count {count} • distinct {distinct}",
    aggregationNote: "Current source: {source} • scope: {scope} • header: {header}{helper} • grouping levels: {depth} • match: {match}{having}.",
    aggregationMeasureNote: "Measures: {measures}. Incompatible extra measures are skipped automatically for the selected aggregation.",
    aggregationHeaderAuto: "auto -> row {row}",
    aggregationHeaderRow: "row {row}",
    aggregationHelper: " (helper)",
    aggregationHavingValue: " • filter: > {value}",
    aggregationHavingPercent: " • filter: > {value}%",
    aggregationHavingAboveValue: "Value >",
    aggregationHavingAboveTotal: "% total >",
    aggregationHavingAboveMax: "% max >",
    aggregationMeasureSearchPlaceholder: "search...",
    aggregationSearchCount: "{visible} of {total}",
    aggregationResult: "Result",
    aggregationTooltipCount: "How many records or occurrences are in each group",
    aggregationTooltipAvg: "Average value, for example average duration or amount",
    aggregationTooltipMedian: "Middle value: half the values are lower and half are higher",
    aggregationTooltipMin: "Smallest value in the group",
    aggregationTooltipMax: "Largest value in the group",
    aggregationTooltipSum: "Total sum of all values in the group",
    aggregationTooltipDistinct: "How many different non-repeated values are in the group",
    aggregationTooltipEarliest: "Oldest / earliest date in the group",
    aggregationTooltipLatest: "Newest / latest date in the group",
    aggregationNoData: "Load a sheet to run aggregations.",
    aggregationNoOptions: "No useful grouping or measure options for the current data source.",
    aggregationNoResults: "The current grouping and measure combination returned no results.",
    noInsightFlags: "No important flags for the current file.",
    workbookInsightsEmpty: "Load a file to see workbook metadata.",
    sheetInsightsEmpty: "Load a sheet to see data quality and structure signals.",
    sectionsEmpty: "Load a sheet to detect sections and layout blocks.",
    sheetSummaryEmpty: "Load a sheet to see a quick structure summary and key signals.",
    analysisComputing: "Computing…",
    durationNoGroup: "Detect repeating column blocks first. This panel works best on sheets with cycles or series of similar fields.",
    durationNoConfig: "Blocks were detected, but no person + from/to or person + duration pair could be found. If the header is unusual, the module also tries to infer from data, but there is still not enough here.",
    durationNoRecords: "Blocks were recognized, but the current view has no records with complete time data for the same value.",
    durationRecordsCount: "Time records",
    durationAvgTime: "Average time",
    durationMedian: "Median",
    durationMin: "Min",
    durationMax: "Max",
    durationInProgress: "In progress",
    durationClosed: "Closed",
    durationNoteFiltered: 'Analysis covers the currently filtered view ({visible} of {source} rows). Open records without an "end" date are counted up to today.',
    durationNoteFull: 'Analysis covers the entire current sheet view. Open records without an "end" date are counted up to today.',
    durationNoteInferred: " The column layout was partly inferred from the data because the header was not ideal.",
    durationNoteHelper: " This analysis used header row {row} as a helper because it fit better than the currently selected {current}.",
    durationStatusLabel: "Status",
    durationStatusAll: "All",
    durationStatusClosed: "Closed only",
    durationStatusOpen: "Open only",
    durationSortLabel: "Sort by",
    durationSortAvg: "Average",
    durationSortMedian: "Median",
    durationSortCount: "Record count",
    durationSortMax: "Maximum",
    durationSortMin: "Minimum",
    durationShowLabel: "Show records",
    durationShowAll: "All",
    durationViewClassic: "Classic view",
    durationShowFull: "Show all",
    durationShownPartial: "Showing {shown} of {total} results.",
    durationShownAll: "Showing all results: {total}.",
    durationDaysUnit: "{n} days",
    durationMonthsShort: "mo",
    durationNone: "none",
    durationMetaLine: "Average {avg} • median {median} • records {count} • in progress {open} • range {min} -> {max}",
    durationShowInTable: "Show in table",
    secKindTable: "Table",
    secKindGroup: "Block",
    secKindCandidate: "Header",
    secKindSubheader: "Section",
    secKindLayout: "Layout",
    secTableData: "Data table",
    secFallback: "Sheet section",
    secMetaHeader: "Header: row {row} • columns {cols}",
    secCurrentHeaderRow: "Current header row: {row}",
    secPossibleHeaderRow: "Possible header row: {row}",
    secMetaRowCols: "Row {row} • columns {cols}",
    secMetaDataRow: "Data row {row}",
    secSetHeader: "Set header",
    secJump: "Jump",
    blkFallbackLabel: "Block {n}",
    blkSignatureLabel: "Repeating layout: {n} blocks",
    blkSignatureLabelSingle: "Repeating column layout",
    blkSignatureMeta: "{n} blocks of {span} columns",
    blkCycleLabel: "Repeating layout: {n} cycles",
    blkCycleMeta: "{n} cycles • {fields} fields in schema",
    blkMergedMeta: "{n} blocks of {span} columns • {first} -> {last}",
    blkPrefixNote: " • fixed columns before blocks: {cols}",
    blkSpanBadge: "{n} col.",
    blkColumnsMeta: "Columns {range}",
    blkJumpTo: "Jump to block",
    inspSheets: "Sheets",
    inspHiddenSheets: "Hidden sheets",
    inspVeryHidden: "Very hidden",
    inspDefinedNames: "Named ranges",
    inspMacroFile: ".xlsm macro file",
    inspVisibleAllRows: "Visible / total rows",
    inspColumns: "Columns",
    inspFormulas: "Formulas",
    inspMerges: "Merges (ranges / cells)",
    inspHiddenColsRows: "Hidden columns / rows",
    inspNumericDateCols: "Numeric / date columns",
    inspSparseCols: "Sparse columns",
    inspLongText: "Long texts",
    flagProcessSheet: "Looks like a process / SLA sheet",
    flagDuplicateHeaders: "Duplicate headers: {n}",
    flagDuplicateRows: "Duplicate rows: {n}",
    flagFormulasNoResult: "Formulas without a result: {n}",
    flagComments: "Comments: {n}",
    flagLinks: "Links: {n}",
    flagVeryHiddenSheets: "Has very hidden sheets",
    kpiCandidates: "KPI candidates",
    kpiSource: "Source",
    kpiRowsAboveHeader: "Rows above header {row}",
    kpiRowsAboveDetected: "Rows above detected header {row}",
    kpiTitleAliases: "{label}: {value}\nAlso as: {aliases}",
    kpiMetaPlain: "{address} • label {labelAddress}",
    kpiMetaAliases: "{address} • label {labelAddress} • also as: {aliases}{more}",
    kpiShowSource: "Show source",
    aggHavingTitle: "Enter a threshold value (e.g. 10 means >10)",
    inspChipSections: "Sections",
    inspChipBlocks: "Blocks",
    inspChipFlagged: "Flagged columns",
    inspTopSignal: "Top signal",
    inspSetHeaderRow: "Set header: {row}",
    inspBackToClassic: "Back to classic view",
    inspSwitchToLong: "Switch to Wide-to-Long",
    inspJumpToCol: "Jump to column: {header}",
    colFallback: "Column {n}",
    profMeta: "Column {col} • empty {empty}% • unique {unique}",
    profJumpToCol: "Jump to column",
    profType_pusta: "empty",
    profType_formuly: "formulas",
    profType_daty: "dates",
    profType_liczby: "numbers",
    profType_tekst: "text",
    profType_mixed: "mixed",
    profFlag_sparse: "sparse",
    profFlag_mixed: "mixed",
    profFlag_longText: "long texts",
    profFlag_mostlyUnique: "mostly unique",
    profFlag_formulaColumn: "formula column",
    profFlag_empty: "empty",
    aggGroupKind_person: "person/owner",
    aggGroupKind_category: "status/category",
    aggGroupKind_entity: "client/company",
    aggGroupKind_place: "place/department",
    aggGroupKind_item: "product/project",
    aggGroupKind_time: "time/date",
    aggGroupKind_id: "identifier",
    aggGroupKind_other: "column",
    aggGroupProfileMeta: "{kind} • {unique} unique • {nonEmpty} non-empty",
    choose: "Choose",
    hintDefault: "Click",
    loadingSample: "Generating a sample file...",
    sampleFileName: "Sample-territory-process.xlsx",
    sampleLoadFailed: "Could not generate the sample",
    repeatBlocksEmpty: "No clear repeating blocks for the current sheet. This works best on wide tables with cycles, stages, or similar column series.",
    kpiSummaryEmpty: "No KPI or summary values detected for the current sheet.",
    kpiNoSummary: "No KPI summary.",
    kpiListEmpty: "No reliable KPIs detected above the current data table.",
    columnProfilesEmpty: "Load a sheet to see column profiles and quick issue signals.",
    formulaSummaryFormulas: "Formulas",
    formulaSummaryMissing: "No result",
    formulaSummaryErrors: "With error",
    formulaSummaryTop: "Top function",
    formulaSummaryVisible: "Visible after filter",
    formulaSummaryGroups: "Groups",
    formulaNoSummary: "No formula summary.",
    formulaNoFormulas: "The current sheet has no detected formulas or has not been loaded yet.",
    formulaNoList: "No formulas to show for the current sheet.",
    formulaNoFilterMatch: "No formulas match the current filter.",
    formulaSameCount: "{count} identical",
    formulaJumpFirst: "Jump to first cell",
    formulaJumpCell: "Jump to cell",
    formulaResultLabel: "result",
    formulaAddresses: "Addresses",
    formulaOutsideTable: "outside table",
    formulaEmptyResult: "empty result",
    qsAllSheetsLabel: "All sheets",
    globalSearchEmpty: "No matches",
    globalSearchMore: "+{count} more in this sheet",
    cellStatsRange: "Range",
    cellStatsCount: "Count",
    cellStatsSum: "Sum",
    cellStatsAvg: "Average",
    cellStatsMin: "Min",
    cellStatsMax: "Max",
    toolbarCollapse: "Collapse toolbar",
    toolbarExpand: "Expand toolbar",
  },
};

function t(key, vars = {}) {
  const lang = currentLang;
  // Jedno źródło prawdy z fallbackiem: I18N bieżącego języka → STATIC bieżącego →
  // I18N.pl → STATIC.pl → sam klucz. Dzięki temu tekst zdefiniowany w „drugim"
  // słowniku nie zwraca już surowego klucza. Walidacja: scripts/check-i18n.js.
  let value =
    I18N[lang]?.[key] ??
    STATIC_TRANSLATIONS[lang]?.[key] ??
    I18N.pl[key] ??
    STATIC_TRANSLATIONS.pl[key] ??
    key;
  Object.entries(vars).forEach(([name, replacement]) => {
    value = value.replaceAll(`{${name}}`, String(replacement));
  });
  return value;
}

const SELECT_VALUE_ALIASES_PL = {
  "dowolnie": "all",
  "nie puste (przynajmniej jedna)": "any_non_empty",
  "nie puste (przynajmniej jedno)": "any_non_empty",
  "nie puste (wszystkie)": "all_non_empty",
  "puste (przynajmniej jedna)": "any_empty",
  "puste (przynajmniej jedno)": "any_empty",
  "puste (wszystkie)": "all_empty",
  "z datą (przynajmniej jedna)": "any_non_empty",
  "z data (przynajmniej jedna)": "any_non_empty",
  "z datą (wszystkie)": "all_non_empty",
  "z data (wszystkie)": "all_non_empty",
  "bez daty (przynajmniej jedna)": "any_empty",
  "bez daty (wszystkie)": "all_empty",
  "zawiera": "contains",
  "zaczyna się": "starts_with",
  "równa się": "equals",
  "między": "between",
  "do": "before",
  "od": "after",
  "ostatnie n dni": "last_n_days",
  "dokładnie": "exact",
  "dokladnie": "exact",
  "rosnąco": "asc",
  "rosnaco": "asc",
  "malejąco": "desc",
  "malejaco": "desc",
  "wszystkie funkcje": "all",
  "bez wyniku": "missing",
  "z błędem": "error",
  "z błędami": "error",
  "wartości": "values",
  "formuły": "formulas",
};

const STATIC_TRANSLATIONS = {
  pl: {
    title: "Sheet Workbench",
    description: "Lokalne przeglądanie, filtrowanie i analiza arkuszy Excel",
    heroSub: "Lokalne przeglądanie, filtrowanie i analiza arkuszy",
    introAria: "Intro Mateusza",
    panelToggleAria: "Zwin/rozwin panel",
    link1Title: "Przejdź do strony Mateusz App | formularz i eksport Excel",
    link2Title: "Przejdź do strony Mateusz App | Portal Ogloszeniowy",
    groupData: "Dane",
    groupWork: "Filtry i widok roboczy",
    groupInspect: "Inspekcja arkusza",
    groupAnalyze: "Agregacje i formuły",
    groupHelp: "Pomoc",
    fileAndSheet: "Plik i arkusz",
    dropText: "Przeciągnij plik <strong>.xlsx</strong>",
    dropOr: "lub",
    chooseFile: "Wybierz plik",
    fileInputAria: "Wybierz plik Excel",
    noFile: "Brak pliku",
    sheet: "Arkusz",
    headerRow: "Wiersz nagłówka",
    autoDetectHeader: "Auto wykryj wiersz nagłówka",
    displayMode: "Tryb wyświetlania",
    rowsLimit: "Limit wierszy",
    loadSheet: "Wczytaj arkusz",
    textFilter1: "Filtr tekstowy 1",
    textFilter2: "Filtr tekstowy 2",
    textFilters: "Filtry tekstowe",
    filterBlock1: "Filtr 1",
    filterBlock2: "Filtr 2",
    addSecondFilter: "+ Dodaj drugi filtr",
    removeFilter: "Usuń",
    search: "Szukaj",
    searchInvoice: "np. faktura",
    searchClient: "np. klient",
    mode: "Tryb",
    columns: "Kolumny",
    choose: "Wybierz",
    emptyOrNot: "Puste / niepuste komórki",
    invert: "Odwróć",
    onlyRowsWithData: "Tylko wiersze z danymi",
    highlightMatchCells: "Podświetl pasujące komórki",
    freezeHeadersLabel: "Zablokuj wiersze nagłówków",
    cellStyleOptionsTitle: "Style komórek z pliku",
    showFontColors: "Pokaż kolory czcionek",
    showCellFills: "Pokaż wypełnienia komórek",
    showCellFonts: "Pokaż formatowanie tekstu",
    showCellBorders: "Pokaż obramowania",
    showConditionalFormatting: "Pokaż formatowanie warunkowe",
    showSubheaders: "Wyróżniaj podnagłówki",
    recalcDates: "Przeliczaj formuły z datą (na dziś)",
    smartColWidths: "Inteligentne dopasowanie szerokości",
    wrapCells: "Zawijaj tekst w komórkach",
    dateFilter: "Filtr dat",
    lastDays: "Ostatnie dni",
    from: "Od",
    to: "Do",
    dateColumns: "Kolumny dat",
    actions: "Akcje",
    filter: "Filtruj",
    resetFilters: "Reset filtrów",
    resetWidths: "Reset szerokości",
    save: "Zapisz",
    saveAs: "Zapisz jako...",
    saveInPlace: "Zapisz (w miejscu lub jako)",
    sortingPresets: "Sortowanie i presety",
    sortColumn: "Kolumna sortowania",
    direction: "Kierunek",
    addSort: "Dodaj sortowanie",
    sortPreset: "Preset sortowania",
    savePreset: "Zapisz preset",
    loadPreset: "Wczytaj preset",
    deletePreset: "Usuń preset",
    view: "Widok",
    zoom: "Powiększenie",
    rowHeightLabel: "Wysokość wierszy (px)",
    colWidthLabel: "Szerokość kolumn (px)",
    freezeFirstColLabel: "Zablokuj pierwszą kolumnę",
    workbenchAnalysis: "Analiza workbench",
    file: "Plik",
    sheetSection: "Arkusz",
    flags: "Flagi",
    kpiSummary: "KPI / Podsumowanie",
    kpiHint: "Pomocnicze wyciąganie najważniejszych liczb i wskaźników z górnej części arkusza.",
    sheetLayout: "Układ arkusza",
    sheetLayoutHint: "Jedno miejsce na orientację w arkuszu: sekcje, kolumny, powtarzalne bloki i szybkie sygnały o układzie danych.",
    mapAndColumns: "Mapa arkusza i kolumn",
    mapHint: "Tu masz jednocześnie szybkie skoki po sekcjach oraz profil kolumn, żeby łatwiej rozumieć układ arkusza bez przeskakiwania między dwoma podobnymi blokami.",
    sectionsJumps: "Sekcje i skoki",
    columnsSignals: "Kolumny i sygnały",
    blockDetector: "Wykrywacz bloków",
    blockHint: "Pomocniczy widok dla szerokich arkuszy z cyklami, turami albo powtarzalnymi sekcjami kolumn.",
    durationAnalysis: "Analiza czasu / osób",
    durationHint: "Lokalna analiza dla wykrytych bloków: łączenie tych samych osób lub innych wartości w jednej kolumnie i liczenie średnich czasów na podstawie dat lub długości.",
    aggregations: "Agregacje",
    monthlySummaryTitle: "Podsumowanie miesięczne",
    monthlySummaryHint: "Rozkład wierszy na 12 ostatnich miesięcy wg kolumny z datą. Wybierz datę i miarę (liczba / suma / średnia). Liczone na aktualnym widoku.",
    aggregationsHint: "Lekki kreator typu grupuj / mierz / agreguj. Działa na aktualnym widoku albo całym arkuszu, z opcją pracy na Wide-to-Long tam, gdzie ma to sens.",
    formulaHint: "Pogrupowany przegląd formuł z aktualnego arkusza: wyszukiwanie, szybkie flagi, skrócone podglądy i skok do komórki.",
    searchFormula: "Szukaj formuły",
    formulaPlaceholder: "np. XLOOKUP, SUMIFS, A1, kwota",
    formulaFilter: "Filtr",
    function: "Funkcja",
    log: "Log",
    logAria: "Log zdarzeń",
    shortcuts: "Skróty i info",
    quickSearchActionFilter: "Filtruj",
    quickSearchActionHighlight: "Zaznacz",
    quickSearchActionCells: "Podświetl pasujące",
    quickSearchActionFilterCells: "Filtruj + podświetl",
    quickSearchOperatorsLabel: "&&·||",
    quickSearchOperatorsTitle: "Włącz operatory szukania: || (lub) · && (oraz) · ! (bez słowa) · {…} (grupowanie) · >> (większe) · << (mniejsze; z „=” → ≥/≤). /| Przykłady: „Kowalski && Faktura”, „>>1000”, „=<<50”, „>>2026-01-01”.",
    searchOperatorsToggle: "Operatory wyszukiwania",
    searchOperatorsTitle: "Operatory wyszukiwania",
    arrows: "Strzałki",
    clickKey: "Klik",
    clickCell: "zaznacz jedną komórkę (aktywną); kolejny klik resetuje zaznaczenie",
    arrowFocus: "przesuwasz aktywną komórkę (resetuje zaznaczenie)",
    or: "LUB",
    clearRowFocus: "Zwiń zaznaczenie do komórki, a potem odznacz fokus",
    afterCellClick: "(po wczesniejszym kliknieciu na komorke) =",
    showCellSelection: "= zaznacz prostokąt od aktywnej komórki do klikniętej",
    moveCellSelection: "rozszerz zaznaczenie od aktywnej komórki",
    editingTitle: "Edycja komórki",
    editOpen: "Dwuklik, Enter lub zacznij pisać — otwórz edytor komórki",
    editCommitVertical: "zatwierdź i w dół / w górę",
    editCommitHorizontal: "zatwierdź i w prawo / w lewo",
    editMoveArrows: "zatwierdź i przejdź do sąsiedniej komórki w kierunku strzałki",
    editCaret: "ruch kursora w tekście (jak w edytorze)",
    editCancel: "anuluj edycję (przywróć poprzednią wartość)",
    clearCellSelection: "Odznacz wszystko (zaznaczenie i fokus)",
    quickSearch: "Szybkie szukanie",
    quickSearchSlash: "Szybkie szukanie (bez aktywnej komórki / poza polem)",
    quickSearchEnter: "w oknie szukania = Szukaj (zawsze, gdy popup otwarty)",
    quickSearchArrows: "w szukaniu = nawigacja po podglądzie wyników",
    columns: "Kolumny",
    theme: "Motyw",
    closeModalPanel: "Zamknij modal / panel",
    orStrong: "LUB",
    andStrong: "ORAZ",
    searchOperatorOrRest: "— wiersz zawiera przynajmniej jeden z wyrazów",
    searchOperatorOr: "LUB — wiersz zawiera przynajmniej jeden z wyrazów",
    searchOperatorOrExample: "np. 'Kowalski || Nowak'",
    searchOperatorAndRest: "— wiersz zawiera wszystkie wyrazy jednocześnie",
    searchOperatorAnd: "ORAZ — wiersz zawiera wszystkie wyrazy jednocześnie",
    searchOperatorAndExample: "np. 'Kowalski && Faktura' · ma pierwszeństwo nad ||",
    searchOperatorNot: "wykluczenie pojedynczego warunku, np. 'Kowalski !J.'",
    searchOperatorCombined: "możesz łączyć operatory, np. 'Kowalski && !Anulowana || Nowak'",
    searchCmpStrong: "Porównania liczb i dat",
    searchCmpRest: "— '>>' większe, '<<' mniejsze; dodaj '=' dla ≥/≤ (>>=, =<<)",
    searchCmpExample: "np. '>>1000', '=<<50', '>>2026-01-01' · zakres: '>>5 && <<10'",
    aggGroupLegendTitle: "Scalanie grup (agregacje)",
    aggGroupLegendCore: "rdzeń — to zostaje i jest zliczane jako grupa",
    aggGroupLegendStar: "jeden krótki „token” śmiecia (do separatora)",
    aggGroupLegendStar2: "dłuższy śmieć (kilka tokenów)",
    aggGroupLegendHash: "tylko cyfry",
    aggGroupLegendAt: "tylko litery",
    aggGroupLegendQ: "jeden znak",
    aggGroupLegendLiteral: "tekst dosłowny (ignoruje wielkość liter)",
    aggGroupLegendExample: "scala „Gr 1 J. Kowalski” + „J. Kowalski” → jedna grupa. Tryb „Rozmyte” robi to automatycznie.",
    searchOperatorBracket: "nawiasy {} wymuszają kolejność, np. '{Kowalski || Nowak} && Faktura'",
    hintDefault: "Kliknij",
    hintLangPl: "Przełącz język na polski",
    hintLangEn: "Switch language to English",
    hintLoad: "Wczytaj wybrany arkusz",
    hintApplyFilters: "Zastosuj aktywne filtry",
    hintResetFilters: "Wyczyść wszystkie filtry",
    hintResetWidths: "Przywróć szerokości kolumn",
    hintSaveAs: "Pobierz kopię pliku",
    hintQuickSearch: "Szukaj szybko w tabeli",
    hintQuickColumns: "Wybierz kolumny szukania",
    hintQuickApply: "Uruchom szybkie szukanie",
    hintExportCsv: "Wyeksportuj widok do CSV",
    statusNoData: "Brak danych",
    quickSearchPlaceholder: "Szybkie szukanie...",
    quickSearchAria: "Tryb szybkiego szukania",
    quickSearchColumnsTitle: "Wybierz kolumny dla szybkiego szukania",
    resetSort: "Domyślne sortowanie",
    exportCsv: "Eksport",
    sidebarScrimAria: "Zamknij panel filtrów",
    chooseColumns: "Wybierz kolumny",
    close: "Zamknij",
    searchColumns: "Szukaj kolumn",
    selectAll: "Zaznacz wszystko",
    clear: "Wyczyść",
    apply: "Zastosuj",
    quickSearchDialogAria: "Szybkie szukanie",
    searchInTable: "Szukaj w tabeli",
    quickSearchPopupPlaceholder: "np. faktura...",
    quickSearchPopupModeAria: "Tryb szybkiego szukania w oknie",
    quickSearchHint: "Enter – zastosuj · ↓↑ – wyniki · Esc – zamknij · / – otwórz",
    dateBetween: "Między",
    dateBefore: "DO",
    dateAfter: "OD",
    dateLastN: "Ostatnie N dni",
    any: "dowolnie",
    anyNonEmpty: "komórka niepusta (przynajmniej jedna)",
    allNonEmpty: "komórki niepuste (wszystkie)",
    anyEmpty: "komórka pusta (przynajmniej jedna)",
    allEmpty: "komórki puste (wszystkie)",
    dateAnyNonEmpty: "komórka niepusta (przynajmniej jedna)",
    dateAllNonEmpty: "komórki niepuste (wszystkie)",
    dateAnyEmpty: "komórka pusta (przynajmniej jedna)",
    dateAllEmpty: "komórki puste (wszystkie)",
    allFunctions: "Wszystkie funkcje",
    noResult: "Bez wyniku",
    withError: "Z błędem",
    scrollTop: "Do góry",
    scrollTopAria: "Przewiń tabelę do góry",
  },
  en: {
    title: "Sheet Workbench",
    description: "Local browsing, filtering, and analysis of Excel sheets",
    heroSub: "Local browsing, filtering, and sheet analysis",
    introAria: "Mateusz intro",
    panelToggleAria: "Collapse/expand panel",
    link1Title: "Go to Mateusz App | form and Excel export",
    link2Title: "Go to Mateusz App | Listings Portal",
    groupData: "Data",
    groupWork: "Filters & working view",
    groupInspect: "Sheet inspection",
    groupAnalyze: "Aggregation & formulas",
    groupHelp: "Help",
    fileAndSheet: "File and sheet",
    dropText: "Drag a <strong>.xlsx</strong> file",
    dropOr: "or",
    chooseFile: "Choose file",
    fileInputAria: "Choose an Excel file",
    noFile: "No file",
    sheet: "Sheet",
    headerRow: "Header row",
    autoDetectHeader: "Auto-detect header row",
    displayMode: "Display mode",
    rowsLimit: "Row limit",
    loadSheet: "Load sheet",
    textFilter1: "Text filter 1",
    textFilter2: "Text filter 2",
    textFilters: "Text filters",
    filterBlock1: "Filter 1",
    filterBlock2: "Filter 2",
    addSecondFilter: "+ Add a second filter",
    removeFilter: "Remove",
    search: "Search",
    searchInvoice: "e.g. invoice",
    searchClient: "e.g. client",
    mode: "Mode",
    columns: "Columns",
    choose: "Choose",
    emptyOrNot: "Empty / non-empty cells",
    invert: "Invert",
    onlyRowsWithData: "Only rows with data",
    highlightMatchCells: "Highlight matching cells",
    freezeHeadersLabel: "Freeze header rows",
    cellStyleOptionsTitle: "Cell styles from file",
    showFontColors: "Show font colors",
    showCellFills: "Show cell fills",
    showCellFonts: "Show text formatting",
    showCellBorders: "Show borders",
    showConditionalFormatting: "Show conditional formatting",
    showSubheaders: "Highlight subheaders",
    recalcDates: "Recalculate date formulas (today)",
    smartColWidths: "Smart column widths",
    wrapCells: "Wrap cell text",
    dateFilter: "Date filter",
    lastDays: "Last days",
    from: "From",
    to: "To",
    dateColumns: "Date columns",
    actions: "Actions",
    filter: "Filter",
    resetFilters: "Reset filters",
    resetWidths: "Reset widths",
    save: "Save",
    saveAs: "Save as...",
    saveInPlace: "Save (in place or save as)",
    sortingPresets: "Sorting and presets",
    sortColumn: "Sort column",
    direction: "Direction",
    addSort: "Add sort rule",
    sortPreset: "Sort preset",
    savePreset: "Save preset",
    loadPreset: "Load preset",
    deletePreset: "Delete preset",
    view: "View",
    zoom: "Zoom",
    rowHeightLabel: "Row height (px)",
    colWidthLabel: "Column width (px)",
    freezeFirstColLabel: "Freeze first column",
    workbenchAnalysis: "Workbench analysis",
    file: "File",
    sheetSection: "Sheet",
    flags: "Flags",
    kpiSummary: "KPI / Summary",
    kpiHint: "A helper extraction of the most important numbers and indicators from the upper part of the sheet.",
    sheetLayout: "Sheet layout",
    sheetLayoutHint: "One place for orienting yourself in the sheet: sections, columns, repeating blocks, and quick layout signals.",
    mapAndColumns: "Sheet map and columns",
    mapHint: "Here you get quick jumps between sections and a column profile at the same time, so it is easier to understand the sheet layout without bouncing between two similar blocks.",
    sectionsJumps: "Sections and jumps",
    columnsSignals: "Columns and signals",
    blockDetector: "Block detector",
    blockHint: "A helper view for wide sheets with cycles, rounds, or repeating column sections.",
    durationAnalysis: "Time / person analysis",
    durationHint: "Local analysis for detected blocks: linking the same people or other values in one column and calculating average times from dates or durations.",
    aggregations: "Aggregations",
    monthlySummaryTitle: "Monthly summary",
    monthlySummaryHint: "Rows distributed over the last 12 months by a date column. Pick the date and a metric (count / sum / average). Computed on the current view.",
    aggregationsHint: "A lightweight group / measure / aggregate builder. It works on the current view or the full sheet, with an option to use Wide-to-Long where it makes sense.",
    formulaHint: "A grouped review of formulas from the current sheet: search, quick flags, shortened previews, and a jump to the cell.",
    searchFormula: "Search formulas",
    formulaPlaceholder: "e.g. XLOOKUP, SUMIFS, A1, amount",
    formulaFilter: "Filter",
    function: "Function",
    log: "Log",
    logAria: "Event log",
    shortcuts: "Shortcuts & info",
    quickSearchActionFilter: "Filter",
    quickSearchActionHighlight: "Highlight",
    quickSearchActionCells: "Highlight matches",
    quickSearchActionFilterCells: "Filter + highlight",
    quickSearchOperatorsLabel: "&&·||",
    quickSearchOperatorsTitle: "Enable search operators: || (or) · && (and) · ! (without word) · {…} (grouping) · >> (greater) · << (less; add „=” → ≥/≤). /| Examples: „Kowalski && Invoice”, „>>1000”, „=<<50”, „>>2026-01-01”.",
    searchOperatorsToggle: "Search operators",
    searchOperatorsTitle: "Search operators",
    arrows: "Arrow keys",
    clickKey: "Click",
    clickCell: "select one cell (active); another click resets the selection",
    arrowFocus: "move the active cell (resets the selection)",
    or: "OR",
    clearRowFocus: "Collapse selection to the cell, then clear focus",
    afterCellClick: "(after clicking a cell) =",
    showCellSelection: "= select a rectangle from the active cell to the clicked one",
    moveCellSelection: "extend the selection from the active cell",
    editingTitle: "Cell editing",
    editOpen: "Double-click, Enter or just start typing — open the cell editor",
    editCommitVertical: "commit and move down / up",
    editCommitHorizontal: "commit and move right / left",
    editMoveArrows: "commit and move to the adjacent cell in the arrow direction",
    editCaret: "move the text caret (like a text editor)",
    editCancel: "cancel editing (restore the previous value)",
    clearCellSelection: "Clear selected cell",
    quickSearch: "Quick search",
    quickSearchSlash: "Quick search (no active cell / outside a field)",
    quickSearchEnter: "in search dialog = Search (whenever the popup is open)",
    quickSearchArrows: "in search = navigate live result preview",
    columns: "Columns",
    theme: "Theme",
    closeModalPanel: "Close modal / panel",
    orStrong: "OR",
    andStrong: "AND",
    searchOperatorOrRest: "— row contains at least one of the terms",
    searchOperatorOr: "OR — row contains at least one of the terms",
    searchOperatorOrExample: "e.g. \'Kowalski || Nowak\'",
    searchOperatorAndRest: "— row contains all terms simultaneously",
    searchOperatorAnd: "AND — row contains all terms simultaneously",
    searchOperatorAndExample: "e.g. \'Kowalski && Invoice\' · takes precedence over ||",
    searchOperatorNot: "exclude a single condition, e.g. 'Kowalski !J.'",
    searchOperatorBracket: "brackets {} force grouping, e.g. '{Kowalski || Nowak} && Invoice'",
    searchOperatorCombined: "operators can be combined, e.g. 'Kowalski && !Canceled || Nowak'",
    searchCmpStrong: "Number & date comparisons",
    searchCmpRest: "— '>>' greater, '<<' less; add '=' for ≥/≤ (>>=, =<<)",
    searchCmpExample: "e.g. '>>1000', '=<<50', '>>2026-01-01' · range: '>>5 && <<10'",
    aggGroupLegendTitle: "Group merging (aggregations)",
    aggGroupLegendCore: "core — kept and counted as the group",
    aggGroupLegendStar: "one short junk „token” (up to a separator)",
    aggGroupLegendStar2: "longer junk (several tokens)",
    aggGroupLegendHash: "digits only",
    aggGroupLegendAt: "letters only",
    aggGroupLegendQ: "one character",
    aggGroupLegendLiteral: "literal text (case-insensitive)",
    aggGroupLegendExample: "merges „Gr 1 J. Kowalski” + „J. Kowalski” → one group. „Fuzzy” mode does it automatically.",
    hintDefault: "Click",
    hintLangPl: "Switch language to Polish",
    hintLangEn: "Switch language to English",
    hintLoad: "Load the selected sheet",
    hintApplyFilters: "Apply active filters",
    hintResetFilters: "Clear all filters",
    hintResetWidths: "Restore column widths",
    hintSaveAs: "Download a copy of the file",
    hintQuickSearch: "Quickly search the table",
    hintQuickColumns: "Choose search columns",
    hintQuickApply: "Run quick search",
    hintExportCsv: "Export the view to CSV",
    statusNoData: "No data",
    quickSearchPlaceholder: "Quick search...",
    quickSearchAria: "Quick search mode",
    quickSearchColumnsTitle: "Choose columns for quick search",
    resetSort: "Default sort",
    exportCsv: "Export",
    sidebarScrimAria: "Close filters panel",
    chooseColumns: "Choose columns",
    close: "Close",
    searchColumns: "Search columns",
    selectAll: "Select all",
    clear: "Clear",
    apply: "Apply",
    quickSearchDialogAria: "Quick search",
    searchInTable: "Search in table",
    quickSearchPopupPlaceholder: "e.g. invoice...",
    quickSearchPopupModeAria: "Quick search mode in dialog",
    quickSearchHint: "Enter – apply · ↓↑ – results · Esc – close · / – open",
    dateBetween: "Between",
    dateBefore: "Before",
    dateAfter: "After",
    dateLastN: "Last N days",
    any: "any",
    anyNonEmpty: "non-empty cell (at least one)",
    allNonEmpty: "non-empty cells (all)",
    anyEmpty: "empty cell (at least one)",
    allEmpty: "empty cells (all)",
    dateAnyNonEmpty: "non-empty cell (at least one)",
    dateAllNonEmpty: "non-empty cells (all)",
    dateAnyEmpty: "empty cell (at least one)",
    dateAllEmpty: "empty cells (all)",
    allFunctions: "All functions",
    noResult: "No result",
    withError: "With error",
    scrollTop: "Back to top",
    scrollTopAria: "Scroll table to top",
  },
};

function setText(selector, value) {
  const el = document.querySelector(selector);
  if (el) el.textContent = value;
}

function setHtml(selector, value) {
  const el = document.querySelector(selector);
  if (!el) return;
  
  // [EN] Safe implementation without innerHTML - only allow specific trusted markup
  // Currently only used for drop text with <strong> tag
  el.textContent = '';
  
  // Handle .xlsx strong formatting
  if (value.includes('<strong>')) {
    const parts = value.split(/<\/?strong>/);
    parts.forEach((part, index) => {
      if (index % 2 === 1) {
        const strong = document.createElement('strong');
        strong.textContent = part;
        el.appendChild(strong);
      } else {
        el.appendChild(document.createTextNode(part));
      }
    });
  } else {
    el.textContent = value;
  }
}

function setAttr(selector, attr, value) {
  const el = document.querySelector(selector);
  if (el) el.setAttribute(attr, value);
}

function setHint(selector, value) {
  setAttr(selector, "data-hint", value);
}

function setFieldLabel(controlId, text) {
  const control = document.getElementById(controlId);
  const label = control?.closest("label.field");
  if (!label) return;
  const textNode = Array.from(label.childNodes).find((node) => node.nodeType === Node.TEXT_NODE && node.textContent.trim());
  if (textNode) {
    textNode.textContent = `${text}\n`;
  }
}

function setCheckboxText(controlId, text) {
  const control = document.getElementById(controlId);
  const label = control?.closest("label.checkbox");
  if (!label) return;
  const span = label.querySelector("span");
  if (span) { span.textContent = text; return; }
  const textNode = Array.from(label.childNodes).find((n) => n.nodeType === Node.TEXT_NODE && n.textContent.trim());
  if (textNode) textNode.textContent = ` ${text}`;
}

function setOperatorsToggleText(controlId, text, title) {
  const control = document.getElementById(controlId);
  const label = control?.closest(".qs-operators-toggle");
  if (!label) return;
  const span = label.querySelector(".qs-operators-toggle-text");
  if (span) span.textContent = text;
  if (title) {
    label.setAttribute("data-hint", title);
    control.removeAttribute("title");
    label.removeAttribute("title");
  }
}

function setShortcutTexts(copy) {
  document.querySelectorAll("[data-shortcut-text]").forEach((el) => {
    const key = el.dataset.shortcutText;
    if (copy[key]) el.textContent = copy[key];
  });
}

function setButtonLabel(selector, text) {
  const button = document.querySelector(selector);
  if (!button) return;
  const badge = button.querySelector(".filter-badge");
  let textNode = Array.from(button.childNodes).find((node) => node.nodeType === Node.TEXT_NODE && node.textContent.trim());
  if (!textNode) {
    textNode = document.createTextNode("");
    if (badge) {
      button.insertBefore(textNode, badge);
    } else {
      button.appendChild(textNode);
    }
  }
  textNode.textContent = ` ${text} `;
}

const SELECT_VALUE_ALIASES = {
  filterMode: {
    contains: "contains",
    "zawiera": "contains",
    "starts_with": "starts_with",
    "zaczyna się": "starts_with",
    "zaczyna sie": "starts_with",
    "starts with": "starts_with",
    equals: "equals",
    "równa się": "equals",
    "rowna sie": "equals",
  },
  filterMode2: {
    contains: "contains",
    "zawiera": "contains",
    "starts_with": "starts_with",
    "zaczyna się": "starts_with",
    "zaczyna sie": "starts_with",
    "starts with": "starts_with",
    equals: "equals",
    "równa się": "equals",
    "rowna sie": "equals",
  },
  dateMode: {
    between: "between",
    "między": "between",
    "miedzy": "between",
    before: "before",
    do: "before",
    after: "after",
    od: "after",
    last_n_days: "last_n_days",
    "ostatnie n dni": "last_n_days",
    "last n days": "last_n_days",
  },
  filterEmptyMode: {
    all: "all",
    "dowolnie": "all",
    any: "all",
    any_non_empty: "any_non_empty",
    "nie puste (przynajmniej jedna)": "any_non_empty",
    "nie puste (przynajmniej jedno)": "any_non_empty",
    "not empty (at least one)": "any_non_empty",
    all_non_empty: "all_non_empty",
    "nie puste (wszystkie)": "all_non_empty",
    "not empty (all)": "all_non_empty",
    any_empty: "any_empty",
    "puste (przynajmniej jedna)": "any_empty",
    "puste (przynajmniej jedno)": "any_empty",
    "empty (at least one)": "any_empty",
    all_empty: "all_empty",
    "puste (wszystkie)": "all_empty",
    "empty (all)": "all_empty",
  },
  filterEmptyMode2: {
    all: "all",
    "dowolnie": "all",
    any: "all",
    any_non_empty: "any_non_empty",
    "nie puste (przynajmniej jedna)": "any_non_empty",
    "nie puste (przynajmniej jedno)": "any_non_empty",
    "not empty (at least one)": "any_non_empty",
    all_non_empty: "all_non_empty",
    "nie puste (wszystkie)": "all_non_empty",
    "not empty (all)": "all_non_empty",
    any_empty: "any_empty",
    "puste (przynajmniej jedna)": "any_empty",
    "puste (przynajmniej jedno)": "any_empty",
    "empty (at least one)": "any_empty",
    all_empty: "all_empty",
    "puste (wszystkie)": "all_empty",
    "empty (all)": "all_empty",
  },
  dateEmptyMode: {
    all: "all",
    "dowolnie": "all",
    any: "all",
    any_non_empty: "any_non_empty",
    "z datą (przynajmniej jedna)": "any_non_empty",
    "z data (przynajmniej jedna)": "any_non_empty",
    "with date (at least one)": "any_non_empty",
    all_non_empty: "all_non_empty",
    "z datą (wszystkie)": "all_non_empty",
    "z data (wszystkie)": "all_non_empty",
    "with date (all)": "all_non_empty",
    any_empty: "any_empty",
    "bez daty (przynajmniej jedna)": "any_empty",
    "without date (at least one)": "any_empty",
    all_empty: "all_empty",
    "bez daty (wszystkie)": "all_empty",
    "without date (all)": "all_empty",
  },
  quickSearchMode: {
    contains: "contains",
    "zawiera": "contains",
    exact: "exact",
    "dokladnie": "exact",
    "dokładnie": "exact",
    equals: "exact",
  },
  quickSearchPopupMode: {
    contains: "contains",
    "zawiera": "contains",
    exact: "exact",
    "dokladnie": "exact",
    "dokładnie": "exact",
    equals: "exact",
  },
  displayMode: {
    values: "values",
    "wartości": "values",
    formulas: "formulas",
    "formuły": "formulas",
  },
  sortDirectionSelect: {
    asc: "asc",
    "rosnąco": "asc",
    "rosnaco": "asc",
    "ascending": "asc",
    desc: "desc",
    "malejąco": "desc",
    "malejaco": "desc",
    "descending": "desc",
  },
  formulaFilter: {
    all: "all",
    "wszystkie": "all",
    "all functions": "all",
    missing: "missing",
    "bez wyniku": "missing",
    "no result": "missing",
    error: "error",
    "z błędem": "error",
    "with error": "error",
  },
};

function normalizeSelectAliasKey(value) {
  return String(value || "").trim().toLowerCase().replace(/\s+/g, " ");
}

function normalizeSelectValue(id, value) {
  const aliases = SELECT_VALUE_ALIASES[id];
  if (!aliases) return value;
  const normalized = aliases[normalizeSelectAliasKey(value)];
  return normalized || value;
}

function normalizeSelectElement(select) {
  if (!select?.id) return;
  Array.from(select.options).forEach((option) => {
    const canonical = normalizeSelectValue(select.id, option.value) || normalizeSelectValue(select.id, option.textContent);
    if (canonical && canonical !== option.value) option.value = canonical;
  });
  const normalizedValue = normalizeSelectValue(select.id, select.value);
  if (normalizedValue && normalizedValue !== select.value) select.value = normalizedValue;
}

function getNormalizedSelectValue(select) {
  if (!select) return "";
  normalizeSelectElement(select);
  return normalizeSelectValue(select.id, select.value);
}

function applyStaticTranslations() {
  const lang = currentLang;
  // Scalony słownik: STATIC bieżącego języka wygrywa, ale brakujące klucze spadają
  // na I18N i na pl — żeby odczyt pola nigdy nie był „undefined" (patrz check-i18n.js).
  const copy = {
    ...I18N.pl,
    ...STATIC_TRANSLATIONS.pl,
    ...(I18N[lang] || {}),
    ...(STATIC_TRANSLATIONS[lang] || {}),
  };
  document.documentElement.lang = currentLang;
  document.title = copy.title;
  BASE_TITLE = copy.title;
  const meta = document.getElementById("pageDescription");
  if (meta) meta.setAttribute("content", copy.description);

  setAttr("#introVideo", "aria-label", copy.introAria);
  setText(".hero-sub", copy.heroSub);
  // #panelToggle bez aria-label — nazwa dostępna to widoczny tekst (panelOpen/panelClosed)
  /* Podpowiedzi przez data-hint (silnik cursor-hint), nie natywne title */
  setAttr("#link1", "data-hint", copy.link1Title);
  setAttr("#link1", "aria-label", copy.link1Title);
  setAttr("#link2", "data-hint", copy.link2Title);
  setAttr("#link2", "aria-label", copy.link2Title);
  setAttr("#brandRefresh", "aria-label", t("brandRefreshAria"));
  setText("#appUpdateBtn", t("updateNow"));
  setAttr("#themeToggle", "aria-label", t("themeToggleAria"));
  setAttr("#heroGrip", "aria-label", t("heroGripAria"));

  setText("#group-data-title", copy.groupData);
  setText("#group-work-title", copy.groupWork);
  setText("#group-inspect-title", copy.groupInspect);
  setText("#group-analyze-title", copy.groupAnalyze);
  setText("#group-help-title", copy.groupHelp);
  setText("#panel-file-sheet .panel-title", copy.fileAndSheet);
  setHtml(".drop-text", copy.dropText);
  setText(".drop-or", copy.dropOr);
  setText(".drop-btn", copy.chooseFile);
  setAttr("#fileInput", "aria-label", copy.fileInputAria);
  setText("#fileNameText", currentFileName || copy.noFile);
  setFieldLabel("sheetSelect", copy.sheet);
  setFieldLabel("headerRow", copy.headerRow);
  setCheckboxText("autoHeaderRow", copy.autoDetectHeader);
  setFieldLabel("displayMode", copy.displayMode);
  setFieldLabel("maxRows", copy.rowsLimit);
  setButtonLabel("#loadBtn", copy.loadSheet);
  setButtonLabel("#loadSampleBtn", t("sampleBtn"));
  setButtonLabel("#emptyOpenBtn", t("openFileBtn"));
  setText("#qsAllSheets .qs-scope-text", t("qsAllSheetsLabel"));
  setAttr("#qsAllSheets", "aria-label", t("qsAllSheetsLabel"));
  setText("#qsAllSheetsPopup .qs-scope-text", t("qsAllSheetsLabel"));
  setAttr("#qsAllSheetsPopup", "aria-label", t("qsAllSheetsLabel"));
  setText("#panel-text-filters .panel-title", copy.textFilters);
  setText("#filter1BlockTitle", copy.filterBlock1);
  setFieldLabel("searchQuery", copy.search);
  setAttr("#searchQuery", "placeholder", copy.searchInvoice);
  setFieldLabel("filterMode", copy.mode);
  setFieldLabel("filter1Columns", copy.columns);
  setText("#filter1Pick", copy.choose);
  setFieldLabel("filterEmptyMode", copy.emptyOrNot);
  setCheckboxText("filterNegate", copy.invert);
  setOperatorsToggleText("filterOperators", copy.searchOperatorsToggle, copy.quickSearchOperatorsTitle);
  setCheckboxText("onlyNonEmpty", copy.onlyRowsWithData);
  setCheckboxText("highlightMatchCells", copy.highlightMatchCells);
  setCheckboxText("highlightMatchCellsDate", copy.highlightMatchCells);
  setCheckboxText("freezeHeaders", copy.freezeHeadersLabel);
  setText("#cellStyleOptionsTitle", copy.cellStyleOptionsTitle);
  setCheckboxText("showFontColors", copy.showFontColors);
  setCheckboxText("showCellFills", copy.showCellFills);
  setCheckboxText("showCellFonts", copy.showCellFonts);
  setCheckboxText("showCellBorders", copy.showCellBorders);
  setCheckboxText("showConditionalFormatting", copy.showConditionalFormatting);
  setCheckboxText("showSubheaders", copy.showSubheaders);
  setCheckboxText("recalcDates", copy.recalcDates);
  setCheckboxText("smartColWidths", copy.smartColWidths);
  setCheckboxText("wrapCells", copy.wrapCells);
  setText("#filter2BlockTitle", copy.filterBlock2);
  setButtonLabel("#addFilter2Btn", copy.addSecondFilter);
  setButtonLabel("#removeFilter2Btn", copy.removeFilter);
  setFieldLabel("searchQuery2", copy.search);
  setAttr("#searchQuery2", "placeholder", copy.searchClient);
  setFieldLabel("filterMode2", copy.mode);
  setFieldLabel("filter2Columns", copy.columns);
  setText("#filter2Pick", copy.choose);
  setFieldLabel("filterEmptyMode2", copy.emptyOrNot);
  setCheckboxText("filterNegate2", copy.invert);
  setOperatorsToggleText("filterOperators2", copy.searchOperatorsToggle, copy.quickSearchOperatorsTitle);
  setText("#panel-date-filter .panel-title", copy.dateFilter);
  setFieldLabel("dateMode", copy.mode);
  setFieldLabel("lastDays", copy.lastDays);
  setFieldLabel("dateFrom", copy.from);
  setFieldLabel("dateTo", copy.to);
  setFieldLabel("dateColumns", copy.dateColumns);
  setText("#datePick", copy.choose);
  setFieldLabel("dateEmptyMode", copy.emptyOrNot);
  setCheckboxText("dateNegate", copy.invert);
  setText("#panel-actions .panel-title", copy.actions);
  setButtonLabel("#applyFilterBtn", copy.filter);
  setButtonLabel("#resetFiltersBtn", copy.resetFilters);
  setButtonLabel("#resetWidthsBtn", copy.resetWidths);
  setText("#saveBtn", copy.save);
  // Podpowiedź (#saveBtn) jest statyczna w index.html przez silnik cursor-hint
  // (data-hint-pl/en) — language.js jej nie ustawia.
  setButtonLabel("#saveAsBtn", copy.saveAs);
  // Panel "Narzędzia edycji"
  setText("#panel-edit-tools .panel-title", copy.editToolsPanel);
  setFieldLabel("editScope", copy.editScopeLabel);
  setFieldLabel("editColumnSelect", copy.editColumnLabel);
  setFieldLabel("editOp", copy.editOpLabel);
  setFieldLabel("editPatternMode", copy.editPatternModeLabel);
  setFieldLabel("editPatternInput", copy.editPatternLabel);
  setFieldLabel("editFind", copy.editFindLabel);
  setFieldLabel("editReplace", copy.editReplaceLabel);
  setText("#editReplaceNote", copy.editReplaceNote);
  setCheckboxText("editRegex", copy.editRegexLabel);
  setFieldLabel("editCaseMode", copy.editCaseLabel);
  setCheckboxText("editFilteredOnly", copy.editFilteredOnlyLabel);
  setFieldLabel("editTrimMode", copy.editTrimModeLabel);
  setFieldLabel("editPrefix", copy.editPrefixLabel);
  setFieldLabel("editSuffix", copy.editSuffixLabel);
  setFieldLabel("editPadLen", copy.editPadLenLabel);
  setFieldLabel("editPadChar", copy.editPadCharLabel);
  setFieldLabel("editPadSide", copy.editPadSideLabel);
  setFieldLabel("editConvertTo", copy.editConvertToLabel);
  setText("#applyEditToolBtn", copy.editApply);
  setText("#panel-sort-workbench .panel-title", copy.sortingPresets);
  setFieldLabel("sortColumnSelect", copy.sortColumn);
  setFieldLabel("sortDirectionSelect", copy.direction);
  setText("#addSortRuleBtn", copy.addSort);
  setFieldLabel("sortPresetSelect", copy.sortPreset);
  setText("#saveSortPresetBtn", copy.savePreset);
  setText("#applySortPresetBtn", copy.loadPreset);
  setText("#deleteSortPresetBtn", copy.deletePreset);
  setText("#panel-view .panel-title", copy.view);
  setFieldLabel("zoomLevel", copy.zoom);
  setFieldLabel("rowHeightAll", copy.rowHeightLabel);
  setFieldLabel("colWidthAll", copy.colWidthLabel);
  setCheckboxText("freezeFirstCol", copy.freezeFirstColLabel);
  setText("#panel-workbench-analysis .panel-title", copy.workbenchAnalysis);
  setText('#subtitle-workbook', copy.file);
  setText('#subtitle-sheet', copy.sheetSection);
  setText('#subtitle-flags', copy.flags);
  setText("#panel-kpi-extractor .panel-title", copy.kpiSummary);
  setText("#panel-kpi-extractor .panel-hint", copy.kpiHint);
  setText("#panel-sheet-inspector .panel-title", copy.sheetLayout);
  setText("#panel-sheet-inspector .panel-hint", copy.sheetLayoutHint);
  setText('#subtitle-map', copy.mapAndColumns);
  setText('#mini-title-sections', copy.sectionsJumps);
  setText('#mini-title-columns', copy.columnsSignals);
  setText('#subtitle-blocks', copy.blockDetector);
  setText('#subtitle-duration', copy.durationAnalysis);
  const inspectorHints = document.querySelectorAll("#panel-sheet-inspector .panel-hint");
  if (inspectorHints[0]) inspectorHints[0].textContent = copy.sheetLayoutHint;
  if (inspectorHints[1]) inspectorHints[1].textContent = copy.mapHint;
  if (inspectorHints[2]) inspectorHints[2].textContent = copy.blockHint;
  if (inspectorHints[3]) inspectorHints[3].textContent = copy.durationHint;
  setText("#panel-aggregation-workbench .panel-title", copy.aggregations);
  setText("#panel-aggregation-workbench .panel-hint", copy.aggregationsHint);
  setText("#panel-monthly-summary .panel-title", copy.monthlySummaryTitle);
  setText("#panel-monthly-summary .panel-hint", copy.monthlySummaryHint);
  setText("#panel-formula-workbench .panel-hint", copy.formulaHint);
  setFieldLabel("formulaSearch", copy.searchFormula);
  setAttr("#formulaSearch", "placeholder", copy.formulaPlaceholder);
  setFieldLabel("formulaFilter", copy.formulaFilter);
  setFieldLabel("formulaFunctionFilter", copy.function);
  setText("#panel-log .panel-title", copy.log);
  setAttr("#log", "aria-label", copy.logAria);
  setText("#panel-shortcuts .panel-title", copy.shortcuts);
  setOperatorsToggleText("quickSearchOperators", copy.searchOperatorsToggle, copy.quickSearchOperatorsTitle);
  setOperatorsToggleText("quickSearchPopupOperators", copy.searchOperatorsToggle, copy.quickSearchOperatorsTitle);
  setShortcutTexts(copy);


  const heroStatus = statusEl
    && !statusEl.classList.contains("unsaved")
    && (statusEl.textContent === STATIC_TRANSLATIONS.pl.statusNoData || statusEl.textContent === STATIC_TRANSLATIONS.en.statusNoData);
  if (heroStatus) setStatus(copy.statusNoData);

  const emptyStateVisible = emptyStateEl && !emptyStateEl.classList.contains("hidden");
  if (emptyStateVisible) setEmptyState(t("emptyTitle"), t("emptySub"));

  setAttr("#sidebarScrim", "aria-label", copy.sidebarScrimAria);
  setAttr("#quickSearchPopup", "aria-label", copy.quickSearchDialogAria);
  setAttr("#quickSearchMode", "aria-label", copy.quickSearchAria);
  setAttr("#quickSearchPopupMode", "aria-label", copy.quickSearchPopupModeAria);
  setAttr("#quickSearchAction", "aria-label", copy.quickSearchAria);
  setAttr("#quickSearchPopupAction", "aria-label", copy.quickSearchPopupModeAria);
  setAttr("#quickSearchColumnsBtn", "data-hint", copy.quickSearchColumnsTitle);
  setAttr("#closePicker", "aria-label", copy.close);

  setText("#columnPickerTitle", copy.chooseColumns);
  setAttr("#columnSearch", "placeholder", copy.searchColumns);
  setText("#selectAllBtn", copy.selectAll);
  setText("#clearAllBtn", copy.clear);
  setText("#applyPickBtn", copy.apply);
  setText(".quick-search-popup-label", copy.searchInTable);
  setAttr("#quickSearchPopupInput", "placeholder", copy.quickSearchPopupPlaceholder);
  setText(".quick-search-popup-hint", copy.quickSearchHint);
  setAttr("#tableSkeleton", "aria-label", t("skeletonAria"));
  setAttr("#quickSearch", "placeholder", copy.quickSearchPlaceholder);
  setText("#exportCsvBtn", copy.exportCsv);
  setText("#exportModalTitle", t("exportModalTitle"));
  setText("#exportModalSub", t("exportModalSub"));
  setText("#exportSelectAll", t("exportSelectAll"));
  setText("#exportClearAll", t("exportClearAll"));
  setText("#exportCsvAction", t("exportActionCsv"));
  setText("#exportPrintAction", t("exportActionPrint"));
  setText("#validationPanelTitle", t("validationPanelTitle"));
  setText("#validationHint", t("validationHintText"));
  setText("#validationColumnLabel", t("validationColumnLabel"));
  setText("#validationSourceLabel", t("validationSourceLabel"));
  setText("#validationSourceList", t("validationSourceListOpt"));
  setText("#validationSourceColumn", t("validationSourceColumnOpt"));
  setText("#validationAllowedLabel", t("validationAllowedLabel"));
  setText("#validationDictColumnLabel", t("validationDictColumnLabel"));
  setText("#validationIgnoreEmptyLabel", t("validationIgnoreEmptyLabel"));
  setText("#validationCaseLabel", t("validationCaseLabel"));
  setText("#validationCheckBtn", t("validationCheckLabel"));
  setText("#validationClearBtn", t("validationClearLabel"));
  setText("#validationShowOnlyLabel", t("validationShowOnlyLabel"));
  setText("#derivedPanelTitle", t("derivedPanelTitle"));
  setText("#derivedHint", t("derivedHintText"));
  setText("#derivedFormulaHead", t("derivedFormulaHead"));
  setText("#derivedNameLabel", t("derivedNameLabel"));
  setText("#derivedExprLabel", t("derivedExprLabel"));
  setText("#dcAddBtn", t("derivedAdd"));
  setText("#derivedJoinHead", t("derivedJoinHead"));
  setText("#derivedJoinNameLabel", t("derivedJoinNameLabel"));
  setText("#derivedJoinKeyLabel", t("derivedJoinKeyLabel"));
  setText("#derivedJoinSheetLabel", t("derivedJoinSheetLabel"));
  setText("#derivedJoinSrcKeyLabel", t("derivedJoinSrcKeyLabel"));
  setText("#derivedJoinSrcRetLabel", t("derivedJoinSrcRetLabel"));
  setText("#dcJoinBtn", t("derivedJoinAdd"));
  setText("#dcRecalcBtn", t("derivedRecalc"));
  setText("#dcClearAllBtn", t("derivedClearAll"));
  setText("#resetSortBtn", copy.resetSort);
  setText("#loadingText", t("loadingGeneric"));
  setText("#scrollTopFab .fab-label", copy.scrollTop);
  setAttr("#scrollTopFab", "aria-label", copy.scrollTopAria);

  if (langButtons.length) {
    langButtons.forEach((button) => {
      const isActive = button.dataset.lang === currentLang;
      button.classList.toggle("is-active", isActive);
      button.setAttribute("aria-pressed", String(isActive));
    });
  }

  applySelectTranslations();
}

function applySelectTranslations() {
  const selectValueToI18nKey = {
    displayMode: {
      values: "values",
      formulas: "formulas",
    },
    editScope: { column: "editScopeColumn", selection: "editScopeSelection" },
    editOp: {
      pattern: "editOpPattern", replace: "editOpReplace", case: "editOpCase",
      trim: "editOpTrim", affix: "editOpAffix", pad: "editOpPad", convert: "editOpConvert",
    },
    editPatternMode: { pattern: "editPModePattern", fuzzy: "editPModeFuzzy" },
    editCaseMode: { upper: "editCaseUpper", lower: "editCaseLower", title: "editCaseTitle" },
    editTrimMode: { ends: "editTrimEnds", collapse: "editTrimCollapse", hard: "editTrimHard" },
    editPadSide: { start: "editPadStart", end: "editPadEnd" },
    editConvertTo: { number: "editConvertNumber", date: "editConvertDate", text: "editConvertText" },
    filterMode: {
      contains: "contains",
      starts_with: "startsWith",
      equals: "equals",
    },
    filterMode2: {
      contains: "contains",
      starts_with: "startsWith",
      equals: "equals",
    },
    dateMode: {
      between: "dateBetween",
      before: "dateBefore",
      after: "dateAfter",
      last_n_days: "dateLastN",
    },
    filterEmptyMode: {
      all: "any",
      any_non_empty: "anyNonEmpty",
      all_non_empty: "allNonEmpty",
      any_empty: "anyEmpty",
      all_empty: "allEmpty",
    },
    filterEmptyMode2: {
      all: "any",
      any_non_empty: "anyNonEmpty",
      all_non_empty: "allNonEmpty",
      any_empty: "anyEmpty",
      all_empty: "allEmpty",
    },
    dateEmptyMode: {
      all: "any",
      any_non_empty: "dateAnyNonEmpty",
      all_non_empty: "dateAllNonEmpty",
      any_empty: "dateAnyEmpty",
      all_empty: "dateAllEmpty",
    },
    quickSearchMode: {
      contains: "contains",
      exact: "equals",
    },
    quickSearchPopupMode: {
      contains: "contains",
      exact: "equals",
    },
    quickSearchAction: {
      filter: "quickSearchActionFilter",
      highlight: "quickSearchActionHighlight",
      cells: "quickSearchActionCells",
      "filter-cells": "quickSearchActionFilterCells",
    },
    quickSearchPopupAction: {
      filter: "quickSearchActionFilter",
      highlight: "quickSearchActionHighlight",
      cells: "quickSearchActionCells",
      "filter-cells": "quickSearchActionFilterCells",
    },
    sortDirectionSelect: {
      asc: "sortAsc",
      desc: "sortDesc",
    },
    formulaFilter: {
      all: "allLabel",
      missing: "noResult",
      error: "withError",
    },
  };

  const plToEn = {
    "wartości": "values",
    "formuły": "formulas",
    "zawiera": "contains",
    "zaczyna się": "starts_with",
    "równa się": "equals",
    "między": "between",
    "do": "before",
    "od": "after",
    "ostatnie n dni": "last_n_days",
    "dowolnie": "all",
    "nie puste (przynajmniej jedna)": "any_non_empty",
    "nie puste (przynajmniej jedno)": "any_non_empty",
    "nie puste (wszystkie)": "all_non_empty",
    "puste (przynajmniej jedna)": "any_empty",
    "puste (przynajmniej jedno)": "any_empty",
    "puste (wszystkie)": "all_empty",
    "z datą (przynajmniej jedna)": "any_non_empty",
    "z datą (wszystkie)": "all_non_empty",
    "bez daty (przynajmniej jedna)": "any_empty",
    "bez daty (wszystkie)": "all_empty",
    "dokładnie": "exact",
    "dokladnie": "exact",
    "filtruj": "filter",
    "zaznacz": "highlight",
    "rosnąco": "asc",
    "rosnaco": "asc",
    "malejąco": "desc",
    "malejaco": "desc",
    "wszystkie funkcje": "all_functions",
    "bez wyniku": "missing",
    "z błędem": "error",
    "z błędami": "error",
  };

  const ids = [
    "displayMode", "filterMode", "filterMode2", "dateMode",
    "filterEmptyMode", "filterEmptyMode2", "dateEmptyMode",
    "quickSearchMode", "quickSearchPopupMode",
    "quickSearchAction", "quickSearchPopupAction",
    "sortDirectionSelect", "formulaFilter",
    "editScope", "editOp", "editPatternMode", "editCaseMode",
    "editTrimMode", "editPadSide", "editConvertTo"
  ];

  ids.forEach((id) => {
    const select = document.getElementById(id);
    if (!select) return;
    const optionMap = selectValueToI18nKey[id] || {};
    Array.from(select.options).forEach((option) => {
      const value = plToEn[option.value] || option.value;
      const i18nKey = optionMap[value];
      if (i18nKey) option.textContent = t(i18nKey);
    });
  });

  const formulaFunctionPlaceholder = document.querySelector("#formulaFunctionFilter option[value=\"\"]");
  if (formulaFunctionPlaceholder) formulaFunctionPlaceholder.textContent = t("allFunctions");
}

function updateLangSwitchIndicator() {
  if (!langSwitchEl || !langSwitchIndicatorEl) return;

  const activeButton = langButtons.find((button) => button.classList.contains("is-active"));
  if (!activeButton) return;

  langSwitchIndicatorEl.style.setProperty("--lang-pill-x", `${activeButton.offsetLeft}px`);
  langSwitchIndicatorEl.style.setProperty("--lang-pill-width", `${activeButton.offsetWidth}px`);

  if (!langSwitchEl.classList.contains("is-ready")) {
    window.requestAnimationFrame(() => {
      langSwitchEl.classList.add("is-ready");
    });
  }
}

function applyLanguage(lang) {
  currentLang = lang === "en" ? "en" : "pl";
  localStorage.setItem(LANG_KEY, currentLang);
  applyStaticTranslations();
  updateColumnSummary();
  updateNetworkBadge();
  updateQuickSearchColumnButtons();
  updateExcelLayoutButtonLabel();
  syncSidebarHandle();
  setReadingMode(rootEl.classList.contains("reading"));
  renderSortRules();
  renderSortPresets();
  // Dynamiczne panele budują DOM w JS, więc muszą się przerenderować,
  // żeby ich etykiety przełączyły język na żywo (nie tylko statyczny tekst).
  scheduleViewRefresh({ analyses: true, formula: true, sync: true });
  updateCellStats();
  updateToolbarToggleLabel();
  updateLangSwitchIndicator();
}