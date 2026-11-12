// row-grow.js — nowe wiersze pod danymi (Etap 1a „dodawanie do pliku”, 2026-11-10).
//
// Tabela pokazuje tylko wiersze z danymi (build-rows-core pomija puste), więc pod danymi nie ma
// „pustej siatki” jak w Excelu. Tu:
//   • „＋ Nowy wiersz” pod tabelą (tryb Edycja) i Enter w ostatnim wierszu dodają WIERSZ ROBOCZY:
//     zwykły obiekt wiersza w baseRows/viewRows z numerem za ostatnim wierszem danych arkusza.
//     Bez przebudowy danych — wpisywanie idzie tą samą drogą co w każdym innym wierszu.
//   • Pierwszy zapis w nowym wierszu (beforeWrite z updateSheetCell / updateSheetFormulaCell):
//     – wygląd jak w wierszu nad nim: indeks stylu (s=) z komórki powyżej dla KAŻDEJ kolumny
//       (ramki, tło, format dat/liczb), także pustych — zapis: payload.xf (xlsx-patch.js);
//     – formuły (tylko wiersz z przycisku/Entera): kolumna, w której 3 ostatnie wiersze mają tę
//       samą formułę względną, dostaje ją przesuniętą o wiersz (jak Excel przy dopisywaniu).
//     Wszystko w tym samym zadaniu JS co wpis → jeden krok Cofnij.
//   • Zapis: growInfo() mówi xlsx-grow.js, o ile urosły dane każdego arkusza (tabele Excela,
//     autofiltr, formatowanie warunkowe, listy wyboru, wymiar arkusza rosną razem z nimi).
// Puste nowe wiersze na KOŃCU danych (np. po Cofnij albo Enter bez wpisu) znikają z widoku przy
// rysowaniu tabeli — ich miejsce zajmują „duchy”; zostaje tylko ten, w którym stoi kursor
// (dawniej zostawały i pod danymi rósł pas pustych wierszy łudząco podobnych do duchów, film 2026-11-10).
//
// Etap 1b — nowe kolumny z prawej („＋ Nowa kolumna”):
//   • nazwa we własnym oknie (zasady jak przy zmianie nazwy: niepusta, unikalna, bez " { } : && ||);
//   • nagłówek + wygląd z kolumny po lewej (nagłówek i każdy wiersz danych), jeden krok Cofnij;
//   • tabela Excela kończąca się na ostatniej kolumnie dostaje kolumnę (w pamięci od razu, w pliku
//     <tableColumn> z tą samą nazwą — xlsx-grow.js); CF, listy wyboru i autofiltr rosną w bok;
//   • kolumny wynikają z arkusza, więc układ wczytuje się na nowo „po cichu” (reloadStructure):
//     filtry, szukanie i sortowanie (po nazwach), szerokości i przewinięcie zostają. Cofnij/Ponów,
//     które zmienia liczbę kolumn, robi to samo.
//
// „Duchy” (po Etapie 1b): 3 blade puste wiersze pod danymi w trybie Edycja, jak pusta siatka
// Excela. Klik / stuknięcie w komórkę ducha = nowy wiersz (addDraft) i pisanie w TEJ kolumnie.
// Osobny <tbody class="ghost-body"> ZA tabelą danych: silnik przewijania, zaznaczanie, filtry,
// liczniki i eksport pracują na tbodyEl i go nie widzą; komórki bez data-col-index (selektory
// komórek danych ich nie łapią), szerokości z <colgroup> tabeli.
// 2 „duchy-kolumny” z prawej: blade litery i „＋” w nagłówku, dwie dodatkowe <col>; wiersze danych
// po prostu kończą się wcześniej (silnik przewijania i rysowanie wierszy bez zmian), a klik w puste
// pole po prawej / w nagłówek ducha = okno nazwy nowej kolumny i pisanie w klikniętym wierszu.
// Bez duchów-kolumn, gdy są kolumny wyliczane (nowa kolumna stanęłaby PRZED nimi, nie w duchu).

const swbRowGrow = (() => {
  const origLast = new Map(); // arkusz → ostatni wiersz z treścią w otwartym pliku (0-based, -1 = pusty)
  const liveLast = new Map(); // arkusz → ostatni wiersz, za którym zaczynają się „nowe” (rośnie)
  const drafts = new Set();   // `${arkusz}\t${r}` — wiersze z przycisku/Entera (dostają formuły)
  const newXf = new Map();    // arkusz → Map(ref → indeks stylu) dla komórek, których nie było w pliku
  const origLastCol = new Map(); // arkusz → ostatnia kolumna z treścią w otwartym pliku (0-based)
  const origTables = new Map();  // tabela (lower) → { sheet, ref } z otwartego pliku

  const keyOf = (sheet, r) => `${sheet}\t${r}`;
  const enc = (r, c) => XLSX.utils.encode_cell({ r, c });

  function lastRowOf(ws) {
    let last = -1;
    if (!ws) return last;
    for (const k in ws) {
      if (k[0] === "!") continue;
      if (r0(k) <= last) continue;
      if (typeof isMeaningfulSheetCell === "function" ? isMeaningfulSheetCell(ws[k]) : ws[k] && ws[k].v != null) last = r0(k);
    }
    return last;
  }
  function r0(ref) { const m = /\d+$/.exec(ref); return m ? Number(m[0]) - 1 : -1; }
  function lastColOf(ws) {
    let last = -1;
    if (!ws) return last;
    for (const k in ws) {
      if (k[0] === "!") continue;
      const c = XLSX.utils.decode_cell(k).c;
      if (c <= last) continue;
      if (typeof isMeaningfulSheetCell === "function" ? isMeaningfulSheetCell(ws[k]) : ws[k] && ws[k].v != null) last = c;
    }
    return last;
  }

  // Nowy plik: zapamiętaj, gdzie kończyły się dane każdego arkusza (zanim cokolwiek dopiszemy).
  function snapshot() {
    origLast.clear(); liveLast.clear(); drafts.clear(); newXf.clear(); origLastCol.clear(); origTables.clear();
    rowToastShown = false;
    if (typeof workbook === "undefined" || !workbook) return;
    (workbook.SheetNames || []).forEach((name) => {
      const last = lastRowOf(workbook.Sheets[name]);
      origLast.set(name, last);
      liveLast.set(name, last);
      origLastCol.set(name, lastColOf(workbook.Sheets[name]));
    });
    if (typeof currentTables !== "undefined" && currentTables) {
      Object.entries(currentTables).forEach(([k, m]) => { if (m && m.ref && m.sheetName) origTables.set(k, { sheet: m.sheetName, ref: m.ref, display: m.displayName || k }); });
    }
  }

  const styleMapOf = (sheet) => {
    if (typeof currentStyleIndexMap === "undefined" || !currentStyleIndexMap) return null;
    if (!currentStyleIndexMap.get(sheet)) currentStyleIndexMap.set(sheet, new Map());
    return currentStyleIndexMap.get(sheet);
  };
  const xfMapOf = (sheet) => { if (!newXf.has(sheet)) newXf.set(sheet, new Map()); return newXf.get(sheet); };

  function canAdd() {
    return typeof workbook !== "undefined" && !!workbook && !!currentSheetName
      && typeof currentDisplayModel !== "undefined" && !!currentDisplayModel && currentDisplayModel.mode === "wide"
      && Array.isArray(currentDisplayModel.headers) && currentDisplayModel.headers.length > 0
      && (typeof swbIsEditMode !== "function" || swbIsEditMode());
  }

  // Filtr / szukanie ukrywa część wierszy: ostatni WIDOCZNY wiersz nie jest końcem danych, a nowy
  // wiersz trafiłby pod ukryte (albo od razu zniknął, bo nie pasuje do filtra). Wtedy bez duchów,
  // bez Entera-dopisywania i z wyłączonym przyciskiem — najpierw trzeba wyczyścić filtr.
  // (Wiersze robocze są i w baseRows, i w viewRows, więc bez filtra długości są równe.)
  const filtered = () => typeof viewRows !== "undefined" && Array.isArray(viewRows) && Array.isArray(baseRows) && viewRows.length < baseRows.length;

  const isEmptyRow = (row) => !row || !Array.isArray(row.values) || row.values.every((v) => v === null || v === undefined || v === "");

  // Wiersz roboczy za ostatnim wierszem danych arkusza. Zwraca klucz wiersza (do fokusu) albo "".
  // gap = ile pustych wierszy zostawić nad nowym (klik w 2./3. ducha). Jak w Excelu: wiersz za
  // odstępem nie jest ciągiem danych — bez formuł i wyglądu z góry, tabele i zakresy nie rosną.
  function addDraft({ col = 0, gap = 0 } = {}) {
    if (!canAdd()) return "";
    if (filtered()) { toast(t("addRowFiltered"), "info"); return ""; }
    const name = currentSheetName;
    const tail = baseRows[baseRows.length - 1];
    // pusty wiersz roboczy już czeka na końcu — nie mnożymy pustych
    if (!gap && tail && tail.isDraftRow && isEmptyRow(tail)) {
      if (!viewRows.includes(tail)) {
        viewRows.push(tail);
        justAdded = tail;
        try { renderActiveTable(); } finally { justAdded = null; }
      }
      const k = getRowSelectionKey(tail);
      setFocusedCell(k, Math.max(0, Math.min(col, currentDisplayModel.headers.length - 1)), { scroll: true });
      return k;
    }
    // ostatni wiersz z TREŚCIĄ teraz, nie „dokąd kiedyś doszliśmy” (liveLast): po Cofnij nowego wiersza
    // kolejny stawał o jeden niżej, nad nim zostawał pusty wiersz i formuły już się nie kopiowały
    let lastData = lastRowOf(workbook.Sheets[name]);
    baseRows.forEach((r) => { if (Number.isFinite(r.rowIndex0) && r.rowIndex0 > lastData) lastData = r.rowIndex0; });
    // Tabela Excela z wierszem sumy na końcu danych: nowy wiersz pod sumą wypadłby z tabeli.
    // Excel wstawia wtedy NAD sumą — to wstawianie w środku (Etap 4); na razie mówimy wprost.
    if (totalsTableEndingAt(name, lastData)) { toast(t("addRowTotalsTable"), "info"); return ""; }
    // pusty wiersz roboczy na końcu (np. po kliknięciu niżej) nie jest danymi — liczymy od treści
    if (gap) baseRows.forEach((x) => { if (x.isDraftRow && isEmptyRow(x) && x.rowIndex0 === lastData) lastData = lastRowOf(workbook.Sheets[name]); });
    const r = lastData + 1 + gap;
    const n = currentDisplayModel.headers.length;
    const tpl = gap ? null : (baseRows.find((x) => x.rowIndex0 === lastData) || tail || null);
    const values = new Array(n).fill(null);
    const row = {
      values,
      rawValues: values,
      display: new Array(n).fill(""),
      rowIndex0: r,
      // wygląd wiersza nad nim od razu (kopia — format nowej komórki nie może zmienić wiersza wyżej)
      cellStyles: tpl && Array.isArray(tpl.cellStyles) ? tpl.cellStyles.map((s) => (s ? cloneStyle(s) : s)) : new Array(n).fill(null),
      isDraftRow: true,
    };
    baseRows.push(row);
    viewRows.push(row);
    if (!gap) drafts.add(keyOf(name, r)); // formuły tylko w wierszu przylegającym do danych
    justAdded = row; // kursor ustawiamy po rysowaniu — do tego czasu wiersz chroniony przed sprzątaniem
    try { renderActiveTable(); } finally { justAdded = null; }
    const k = getRowSelectionKey(row);
    setFocusedCell(k, Math.max(0, Math.min(col, n - 1)), { scroll: true });
    return k;
  }
  function totalsTableEndingAt(sheetName, r) {
    if (typeof currentTables === "undefined" || !currentTables) return false;
    return Object.values(currentTables).some((m) => {
      if (!m || m.sheetName !== sheetName || !(m.totalsRows > 0) || !m.ref) return false;
      try { return XLSX.utils.decode_range(m.ref).e.r === r; } catch { return false; }
    });
  }
  function cloneStyle(s) {
    try { return typeof structuredClone === "function" ? structuredClone(s) : JSON.parse(JSON.stringify(s)); } catch { return s; }
  }

  // Formuła względna w tej samej postaci (angielska, znormalizowana) — do porównania wierszy.
  function shiftedF(f, dr) { try { return FC.serialize(FC.shift(FC.parse(f, { lang: "en" }), dr, 0), { lang: "en" }); } catch { return null; } }
  // Kolumna z „ciągłą” formułą — reguła Excela „rozszerzaj formaty i formuły zakresu danych”:
  // co najmniej 3 z 5 wierszy nad nowym mają tę samą formułę względną. Zwraca odległość do
  // najbliższego takiego wiersza (1…5) albo 0. Dawniej wymagane były dokładnie 3 wiersze tuż nad —
  // jeden wiersz bez formuły (np. dopisany inną drogą) zrywał łańcuch dla wszystkich następnych.
  function formulaSource(ws, r, c) {
    if (typeof FC === "undefined" || !FC) return 0;
    const cand = [];
    for (let d = 1; d <= 5 && r - d >= 0; d++) {
      const cell = ws[enc(r - d, c)];
      const sh = cell && cell.f ? shiftedF(cell.f, d) : null;
      if (sh) cand.push({ d, sh });
    }
    if (cand.length < 3) return 0;
    return cand.filter((x) => x.sh === cand[0].sh).length >= 3 ? cand[0].d : 0;
  }

  function rowHasData(ws, r) {
    let rg;
    try { rg = XLSX.utils.decode_range(ws["!ref"]); } catch { return false; }
    for (let c = rg.s.c; c <= rg.e.c; c++) {
      const cell = ws[enc(r, c)];
      if (cell && (typeof isMeaningfulSheetCell === "function" ? isMeaningfulSheetCell(cell) : cell.v != null)) return true;
    }
    return false;
  }

  function rowIsEmpty(ws, r, n) {
    for (let i = 0; i < n; i++) {
      const cell = ws[enc(r, currentStartCol + i)];
      if (cell && (typeof isMeaningfulSheetCell === "function" ? isMeaningfulSheetCell(cell) : cell.v != null)) return false;
    }
    return true;
  }

  // Woła updateSheetCell / updateSheetFormulaCell PRZED zapisem komórki (r, absC) arkusza.
  // „Nowy wiersz” = za danymi z otwartego pliku i (jeszcze) pusty — po Cofnij wiersz znów jest
  // pusty, więc kolejny wpis znów dostaje wygląd i formuły (dawniej: tylko pierwszy raz).
  let preparing = false; // zapis formuł poniżej też woła beforeWrite — bez rekurencji
  function beforeWrite(sheetName, r, absC) {
    if (preparing || !origLast.has(sheetName) || typeof workbook === "undefined" || !workbook) return;
    if (!(r > origLast.get(sheetName))) return;
    const ws = workbook.Sheets[sheetName];
    if (!ws) return;
    const n0 = currentDisplayModel && Array.isArray(currentDisplayModel.headers) ? currentDisplayModel.headers.length : 0;
    if (!rowIsEmpty(ws, r, n0)) return;
    const k = keyOf(sheetName, r);
    const last = liveLast.get(sheetName);
    if (r > last) liveLast.set(sheetName, r);
    preparing = true;
    try {
      if (typeof undoAutoPart === "function") undoAutoPart(() => prepareRow(sheetName, ws, r, absC, k, last));
      else prepareRow(sheetName, ws, r, absC, k, last);
    } finally { preparing = false; }
  }

  function prepareRow(sheetName, ws, r, absC, k, last) {
    const styleMap = styleMapOf(sheetName);
    const xfMap = xfMapOf(sheetName);
    const tplR = r - 1;
    // wiersz za pustym odstępem: wygląd też nie „spływa” z góry (Excel rozszerza formaty tylko w ciągu)
    if (!rowHasData(ws, tplR)) return;
    const n = currentDisplayModel && Array.isArray(currentDisplayModel.headers) ? currentDisplayModel.headers.length : 0;
    const realCols = [];
    for (let i = 0; i < n; i++) if (typeof cfmtIsRealCol !== "function" || cfmtIsRealCol(i)) realCols.push(i);
    const auto = typeof swbAutoFillOn === "undefined" || swbAutoFillOn; // Widok → Tabela: „Automatycznie uzupełniaj…”
    // 1. styl z wiersza nad nim — dla każdej kolumny (pusta też ma ramkę/tło jak w Excelu)
    if (auto) realCols.forEach((i) => {
      const c = currentStartCol + i;
      const ref = enc(r, c);
      const xf = styleMap ? styleMap.get(enc(tplR, c)) : undefined;
      if (!(Number.isInteger(xf) && xf > 0)) return;
      if (styleMap && !styleMap.has(ref)) styleMap.set(ref, xf);
      xfMap.set(ref, xf);
      if (c === absC || ws[ref]) return;
      if (typeof undoCaptureCell === "function") undoCaptureCell(r, i);
      ws[ref] = { t: "z" };
      recordPendingEdit(sheetName, ref, { xf });
    });
    if (typeof swbEnsureSheetRef === "function") swbEnsureSheetRef(ws, r, absC);
    // 2. zakresy w aplikacji rosną jak w pliku (xlsx-grow.js) — PRZED formułami: [#This Row]
    //    w nowym wierszu liczy się z rozszerzonej tabeli (inaczej #REF!), lista wyboru i kolory
    //    formatowania warunkowego działają od razu, bez ponownego otwierania pliku
    if (r === last + 1) growInApp(sheetName, last, r);
    // 3. formuły kolumn (tylko wiersz dodany przyciskiem / Enterem — nie np. wiersz Autosumy)
    if (!auto || !drafts.has(k) || typeof swbFormulaCommitAt !== "function" || typeof swbFormulaShifted !== "function") return;
    const filled = [];
    realCols.forEach((i) => {
      const c = currentStartCol + i;
      if (c === absC) return; // tu użytkownik właśnie coś wpisuje
      const ref = enc(r, c);
      if (ws[ref] && ws[ref].f) return;
      const d = formulaSource(ws, r, c);
      if (!d) return;
      const text = swbFormulaShifted(sheetName, enc(r - d, c), d, 0);
      if (text && swbFormulaCommitAt(r, i, text, { silent: true }) === "ok") filled.push(XLSX.utils.encode_col(c));
    });
    // dymek tylko za pierwszym razem w pliku — przy każdym wierszu byłby natrętny (Excel: mała ikonka)
    if (filled.length && !rowToastShown) {
      rowToastShown = true;
      autoToast(t("autoFillRowDone", { cols: filled.join(", ") }));
    }
  }

  // Dymek po automatycznym uzupełnieniu (zamiennik ikonki z błyskawicą w Excelu): „Cofnij uzupełnienie”
  // cofa TYLKO część automatyczną (krok z undo.js oznaczony auto), „Nie uzupełniaj automatycznie”
  // wyłącza przełącznik w Widok → Tabela. Krok Cofnij zamyka się w mikrozadaniu — dymek po nim.
  let rowToastShown = false;
  function autoToast(msg) {
    queueMicrotask(() => {
      const step = typeof undoStack !== "undefined" ? undoStack[undoStack.length - 1] : null;
      toast(msg, "info", {
        action: t("autoFillUndo"),
        onAction: () => {
          if (typeof undoCloseStep === "function") undoCloseStep();
          if (step && step.auto && undoStack[undoStack.length - 1] === step) undoLast();
          else toast(t("undoStepGone"), "info");
        },
        action2: t("autoFillStop"),
        onAction2: () => {
          if (typeof autoFillFormulasEl !== "undefined" && autoFillFormulasEl) autoFillFormulasEl.checked = false;
          if (typeof syncCellStyleFlags === "function") syncCellStyleFlags();
          if (typeof saveCellStylePreferences === "function") saveCellStylePreferences();
          toast(t("autoFillStopped"), "info");
        },
      });
    });
  }

  // ── kolumna obliczeniowa (Excel: „Wypełnij formuły w tabelach, aby utworzyć kolumny obliczeniowe”) ──
  // Wołane z edytora PO zatwierdzeniu formuły wpisanej ręcznie w komórkę (r, col) — w tym samym
  // zadaniu, jako część automatyczna (pierwsze Cofnij zdejmuje wypełnienie, formuła w komórce zostaje).
  // Komórka w danych tabeli Excela (bez nagłówka i wiersza sumy) i pozostałe komórki kolumny:
  //   • wszystkie puste → formuła wypełnia całą kolumnę,
  //   • wszystkie z tą samą formułą względną co ta komórka PRZED zmianą (już kolumna obliczeniowa)
  //     → nowa formuła zastępuje ją w całej kolumnie.
  // W każdym innym przypadku (ręczne wartości, różne formuły) nic — jak Excel bez pytania.
  function calcColumn(r, col, prevF) {
    if (!(typeof swbAutoFillOn === "undefined" || swbAutoFillOn)) return 0;
    if (typeof currentTables === "undefined" || !currentTables || typeof swbFormulaShifted !== "function" || typeof FC === "undefined" || !FC) return 0;
    const sheetName = currentSheetName;
    const ws = workbook.Sheets[sheetName];
    const c = currentStartCol + col;
    const cell = ws && ws[enc(r, c)];
    if (!cell || !cell.f) return 0;
    const meta = Object.values(currentTables).find((m) => {
      if (!m || m.sheetName !== sheetName || !m.ref) return false;
      try { const g = XLSX.utils.decode_range(m.ref); return c >= g.s.c && c <= g.e.c && r >= g.s.r + (m.headerRows ?? 1) && r <= g.e.r - (m.totalsRows || 0); } catch { return false; }
    });
    if (!meta) return 0;
    const g = XLSX.utils.decode_range(meta.ref);
    const r1 = g.s.r + (meta.headerRows ?? 1), r2 = g.e.r - (meta.totalsRows || 0);
    if (r2 - r1 > 20000) return 0; // bezpiecznik
    const prevShape = prevF ? shiftedF(prevF, 0) : null;
    let allEmpty = true, allPrev = !!prevShape;
    for (let rr = r1; rr <= r2; rr++) {
      if (rr === r) continue;
      const x = ws[enc(rr, c)];
      const filled = x && (x.f || (typeof isMeaningfulSheetCell === "function" ? isMeaningfulSheetCell(x) : x.v != null && x.v !== ""));
      if (filled) allEmpty = false;
      if (allPrev && !(x && x.f && shiftedF(x.f, r - rr) === prevShape)) allPrev = false;
      if (!allEmpty && !allPrev) return 0;
    }
    if (r2 <= r1) return 0; // tabela z jednym wierszem danych — nie ma czego wypełniać
    let count = 0;
    undoAutoPart(() => {
      for (let rr = r1; rr <= r2; rr++) {
        if (rr === r) continue;
        const text = swbFormulaShifted(sheetName, enc(r, c), rr - r, 0);
        if (text && swbFormulaCommitAt(rr, col, text, { silent: true }) === "ok") count++;
      }
    });
    if (count) {
      requestAnimationFrame(() => renderActiveTable()); // wyniki w wierszach, których edytor nie rysuje
      autoToast(t("autoFillColDone", { count }));
    }
    return count;
  }

  // Ta sama reguła co xlsx-grow.js (zakres kończący się na wierszu `prev` i mający ≥ 2 wiersze
  // rośnie do `next`), na modelach w pamięci. Wiersze 0-based.
  function growRng(rg, prev, next) {
    if (!rg || !rg.s || !rg.e || rg.e.r !== prev || rg.s.r >= rg.e.r) return false;
    rg.e.r = next;
    return true;
  }
  function growInApp(sheetName, prev, next) {
    // tabele Excela (silnik formuł, paski wierszy); z wierszem sumy — bez zmian
    if (typeof currentTables !== "undefined" && currentTables) {
      let changed = false;
      new Set(Object.values(currentTables)).forEach((meta) => {
        if (!meta || meta.sheetName !== sheetName || meta.totalsRows > 0 || !meta.ref) return;
        let rg;
        try { rg = XLSX.utils.decode_range(meta.ref); } catch { return; }
        if (growRng(rg, prev, next)) { meta.ref = XLSX.utils.encode_range(rg); changed = true; }
      });
      // nowy obiekt = nowe granice w pamięci podręcznej silnika (fcTables, fcContext)
      if (changed) currentTables = Object.assign({}, currentTables);
    }
    if (typeof dvRulesBySheet !== "undefined" && dvRulesBySheet && dvRulesBySheet.has(sheetName)) {
      dvRulesBySheet.get(sheetName).forEach((rule) => (rule.ranges || []).forEach((rg) => growRng(rg, prev, next)));
    }
    if (typeof currentCFRules !== "undefined" && currentCFRules && currentCFRules.get && currentCFRules.get(sheetName)) {
      currentCFRules.get(sheetName).forEach((block) => (block.ranges || []).forEach((rg) => growRng(rg, prev, next)));
      if (typeof cfEvalCache !== "undefined" && cfEvalCache) cfEvalCache.delete(sheetName);
    }
    const ws = workbook.Sheets[sheetName];
    const af = ws && ws["!autofilter"];
    if (af && af.ref) {
      try { const rg = XLSX.utils.decode_range(af.ref); if (growRng(rg, prev, next)) af.ref = XLSX.utils.encode_range(rg); } catch { /* */ }
    }
  }

  // Nowa komórka przy zapisie dostaje styl wiersza nad nią (xlsx-patch.js czyta payload.xf).
  function withXf(sheetName, ref, payload) {
    const xf = newXf.has(sheetName) ? newXf.get(sheetName).get(ref) : undefined;
    if (!xf) return payload;
    if (payload === null) return { xf }; // wyczyszczona nowa komórka: zostaje sam wygląd (jak w Excelu)
    if (payload && typeof payload === "object" && !("xf" in payload)) return { ...payload, xf };
    return payload;
  }

  // Dla xlsx-grow.js: { arkusz: { fromRow, toRow, fromCol, toCol, tableCols } } (1-based), gdy dane
  // urosły w dół i/lub w bok. tableCols: nazwy nowych kolumn tabel Excela, czytane z ARKUSZA
  // (wiersz nagłówka tabeli) w chwili zapisu — muszą się zgadzać z nagłówkami, bo tak sprawdza Excel.
  function growInfo() {
    const out = {};
    if (typeof workbook === "undefined" || !workbook) return out;
    origLast.forEach((orig, name) => {
      if (orig < 0) return; // pusty arkusz — nie ma zakresów kończących się na danych
      const ws = workbook.Sheets[name];
      const g = {};
      const cur = lastRowOf(ws);
      if (cur > orig) {
        // tabele, filtr, CF i listy rosną tylko o wiersze PRZYLEGAJĄCE do danych (jak Excel) —
        // wiersz za pustym odstępem (klik w dalszego ducha) to osobny blok; wymiar arkusza do końca
        let cont = orig;
        while (cont < cur && rowHasData(ws, cont + 1)) cont++;
        if (cont > orig) { g.fromRow = orig + 1; g.toRow = cont + 1; }
        if (cur > cont) g.dimRow = cur + 1;
      }
      const oc = origLastCol.get(name);
      const cc = lastColOf(ws);
      if (oc >= 0 && cc > oc) {
        g.fromCol = oc + 1; g.toCol = cc + 1;
        origTables.forEach((t, key) => {
          if (t.sheet !== name) return;
          let rg;
          try { rg = XLSX.utils.decode_range(t.ref); } catch { return; }
          if (rg.e.c !== oc) return;
          const names = [];
          for (let c = oc + 1; c <= cc; c++) {
            const cell = ws[enc(rg.s.r, c)];
            const v = cell && cell.v != null ? String(cell.v).trim() : "";
            if (!v) return; // kolumna bez nazwy — tabela nie rośnie w bok
            names.push(v);
          }
          g.tableCols = g.tableCols || {};
          g.tableCols[key] = names;
        });
      }
      if (Object.keys(g).length) out[name] = g;
    });
    return out;
  }

  // ── Etap 1b: nowa kolumna z prawej ────────────────────────────────────────
  const realCount = () => (typeof swbRealHeaderCount === "function" ? swbRealHeaderCount() : currentHeaders.length);

  function nameProblem(raw) {
    const name = String(raw ?? "").replace(/\s+/g, " ").trim();
    if (!name) return t("hdrRenameEmpty");
    if (typeof SWB_HDR_BAD !== "undefined" && SWB_HDR_BAD.test(name)) return t("hdrRenameBadChars");
    const key = typeof normalizeHeaderKey === "function" ? normalizeHeaderKey(name) : name.toLowerCase();
    const taken = (currentHeaders || []).some((h) => (typeof normalizeHeaderKey === "function" ? normalizeHeaderKey(h) : String(h).toLowerCase()) === key);
    if (taken) return t("hdrRenameDuplicate", { name });
    return "";
  }
  function defaultColName() {
    for (let n = realCount() + 1; n < realCount() + 500; n++) {
      const name = t("addColDefault", { n });
      if (!nameProblem(name)) return name;
    }
    return "";
  }

  // Ta sama reguła w bok (kolumny 0-based): koniec na `prev`, ≥ 2 kolumny → do `next`.
  function growRngCol(rg, prev, next) {
    if (!rg || !rg.s || !rg.e || rg.e.c !== prev || rg.s.c >= rg.e.c) return false;
    rg.e.c = next;
    return true;
  }
  function growColsInApp(sheetName, prev, next, name, hdrR) {
    if (typeof currentTables !== "undefined" && currentTables) {
      let changed = false;
      new Set(Object.values(currentTables)).forEach((meta) => {
        if (!meta || meta.sheetName !== sheetName || !meta.ref) return;
        let rg;
        try { rg = XLSX.utils.decode_range(meta.ref); } catch { return; }
        if (rg.s.r !== hdrR || !growRngCol(rg, prev, next)) return; // nagłówek tabeli = wiersz nagłówka widoku
        meta.ref = XLSX.utils.encode_range(rg);
        if (Array.isArray(meta.colNames)) meta.colNames.push(name);
        if (meta.columns) meta.columns[name.toLowerCase()] = next;
        changed = true;
      });
      if (changed) currentTables = Object.assign({}, currentTables);
    }
    if (typeof dvRulesBySheet !== "undefined" && dvRulesBySheet && dvRulesBySheet.has(sheetName)) {
      dvRulesBySheet.get(sheetName).forEach((rule) => (rule.ranges || []).forEach((rg) => growRngCol(rg, prev, next)));
    }
    if (typeof currentCFRules !== "undefined" && currentCFRules && currentCFRules.get && currentCFRules.get(sheetName)) {
      currentCFRules.get(sheetName).forEach((block) => (block.ranges || []).forEach((rg) => growRngCol(rg, prev, next)));
      if (typeof cfEvalCache !== "undefined" && cfEvalCache) cfEvalCache.delete(sheetName);
    }
    const ws = workbook.Sheets[sheetName];
    const af = ws && ws["!autofilter"];
    if (af && af.ref) {
      try { const rg = XLSX.utils.decode_range(af.ref); if (growRngCol(rg, prev, next)) af.ref = XLSX.utils.encode_range(rg); } catch { /* */ }
    }
  }

  // Zapis nowej kolumny — synchronicznie, jedno zadanie JS = jeden krok Cofnij.
  function writeColumn(name) {
    const sheetName = currentSheetName;
    const ws = workbook.Sheets[sheetName];
    const real = realCount();
    const hdrR = currentHeaderRow - 1;
    const absC = currentStartCol + real;
    const leftC = absC - 1;
    const styleMap = styleMapOf(sheetName);
    const xfMap = xfMapOf(sheetName);
    const inherit = (r) => {
      const ref = enc(r, absC);
      const xf = styleMap ? styleMap.get(enc(r, leftC)) : undefined;
      if (Number.isInteger(xf) && xf > 0) {
        if (styleMap && !styleMap.has(ref)) styleMap.set(ref, xf);
        xfMap.set(ref, xf);
      }
      return ref;
    };
    // nagłówek
    if (typeof undoCaptureCell === "function") undoCaptureCell(hdrR, real);
    const hdrRef = inherit(hdrR);
    ws[hdrRef] = { t: "s", v: name };
    recordPendingEdit(sheetName, hdrRef, { v: name, t: "s" });
    if (typeof swbEnsureSheetRef === "function") swbEnsureSheetRef(ws, hdrR, absC);
    // wygląd w wierszach danych (z komórki po lewej)
    (baseRows || []).forEach((row) => {
      if (!row || row.isPreRow || !Number.isFinite(row.rowIndex0) || row.rowIndex0 <= hdrR) return;
      const ref = inherit(row.rowIndex0);
      if (!xfMap.has(ref) || ws[ref]) return;
      if (typeof undoCaptureCell === "function") undoCaptureCell(row.rowIndex0, real);
      ws[ref] = { t: "z" };
      recordPendingEdit(sheetName, ref, { xf: xfMap.get(ref) });
    });
    growColsInApp(sheetName, leftC, absC, name, hdrR);
    if (typeof bumpSheetDataStamp === "function") bumpSheetDataStamp();
    return real; // indeks nowej kolumny w widoku
  }

  async function addColumn({ rowKey = null } = {}) {
    if (!canAdd()) return false;
    if (typeof swbEditGuard === "function" && !swbEditGuard()) return false;
    if (typeof appPrompt !== "function") return false;
    const raw = await appPrompt(t("addColPrompt"), defaultColName(), {
      title: t("addColTitle"),
      okLabel: t("addColOk"),
      validate: (v) => nameProblem(v),
    });
    if (raw == null) return false;
    const name = String(raw).replace(/\s+/g, " ").trim();
    if (nameProblem(name) || !canAdd()) return false;
    const col = writeColumn(name);
    setDirtyState(true);
    await reloadStructure({ shift: { from: col, by: 1 }, focusCol: col, focusRowKey: rowKey });
    if (typeof updateUndoButtons === "function") updateUndoButtons();
    toast(t("addColDone", { name }), "success", typeof swbUndoToastOpts === "function" ? swbUndoToastOpts() : null);
    return true;
  }

  // Liczba kolumn arkusza ≠ liczba kolumn widoku (np. Cofnij dodania kolumny)?
  function columnsChanged() {
    if (typeof workbook === "undefined" || !workbook || !currentSheetName) return false;
    const ws = workbook.Sheets[currentSheetName];
    const last = lastColOf(ws);
    return last >= 0 && last - currentStartCol + 1 !== realCount();
  }

  // Ponowne wczytanie układu arkusza „po cichu” (bez dymka „Wczytano…”, bez skoku do widoku
  // startowego z Excela), z zachowaniem filtrów/sortowania (po nazwach), szerokości kolumn
  // (indeksy za `shift.from` przesunięte o `shift.by`) i przewinięcia.
  let quiet = false;
  async function reloadStructure({ shift = null, focusCol = null, focusRowKey = null } = {}) {
    if (typeof loadBtn === "undefined" || !loadBtn || typeof sheetSelect === "undefined") return;
    const vs = typeof captureViewState === "function" ? captureViewState() : null;
    const widths = {};
    Object.entries(typeof manualColumnWidths !== "undefined" ? manualColumnWidths : {}).forEach(([k, w]) => {
      const i = Number(k);
      widths[shift && i >= shift.from ? i + shift.by : i] = w;
    });
    const wrap = typeof tableWrapEl !== "undefined" ? tableWrapEl : null;
    const scroll = wrap ? { top: wrap.scrollTop, left: wrap.scrollLeft } : null;
    const focus = typeof focusedCellState !== "undefined" && focusedCellState ? { ...focusedCellState } : null;
    const auto = typeof autoHeaderRowEl !== "undefined" && autoHeaderRowEl ? autoHeaderRowEl.checked : null;
    if (auto) autoHeaderRowEl.checked = false; // wiersz nagłówka zostaje ten sam
    quiet = true;
    try {
      sheetSelect.value = currentSheetName;
      loadBtn.click();
      const t0 = performance.now();
      await new Promise((resolve) => {
        const tick = () => (document.body.getAttribute("aria-busy") !== "true" || performance.now() - t0 > 15000 ? resolve() : setTimeout(tick, 40));
        setTimeout(tick, 80);
      });
    } finally {
      quiet = false;
      if (auto) autoHeaderRowEl.checked = true;
    }
    if (vs && typeof applyViewState === "function") applyViewState(vs);
    if (typeof manualColumnWidths !== "undefined") { Object.keys(manualColumnWidths).forEach((k) => delete manualColumnWidths[k]); Object.assign(manualColumnWidths, widths); }
    renderActiveTable();
    if (wrap && scroll) { wrap.scrollTop = scroll.top; wrap.scrollLeft = scroll.left; }
    // Filtry, szerokości i przewinięcie — od razu (bez mignięcia tabeli bez filtra). Kursor i pole pisania
    // dopiero po animowanym przejściu wczytania (klasa vt-scene, rysuje tabelę PÓŹNIEJ niż koniec ładowania):
    // inaczej spóźnione rysowanie zamykało otwarte pole w nowej kolumnie (ghost-rows w tłoku, 2026-11-11).
    await new Promise((resolve) => {
      const t1 = performance.now();
      const tick = () => (!document.documentElement.classList.contains("vt-scene") || performance.now() - t1 > 3000
        ? requestAnimationFrame(() => requestAnimationFrame(resolve)) : setTimeout(tick, 30));
      tick();
    });
    const first = (currentDisplayModel && currentDisplayModel.rows || []).find((r) => r && !r.isPreRow);
    if (Number.isInteger(focusCol) && first) {
      const target = focusRowKey && (currentDisplayModel.rows || []).some((r) => getRowSelectionKey(r) === focusRowKey) ? focusRowKey : getRowSelectionKey(first);
      setFocusedCell(target, focusCol, { scroll: true });
      if (typeof giveGridDomFocus === "function") giveGridDomFocus(); // pisanie od razu trafia do nowej kolumny
      // klik w ducha-kolumnę w konkretnym wierszu = od razu pisanie w tej komórce
      if (focusRowKey === target) requestAnimationFrame(() => {
        const cell = typeof findCellElement === "function" ? findCellElement(focusedCellState) : null;
        if (cell && typeof openCellEditor === "function") openCellEditor(cell);
      });
    }
    else if (focus && focus.rowKey) setFocusedCell(focus.rowKey, Math.min(focus.colIndex0, Math.max(0, currentDisplayModel.headers.length - 1)), { scroll: false });
  }

  // Cofnij/Ponów, po którym arkusz ma inną liczbę kolumn niż widok — układ od nowa.
  // Ponów wpisu w wierszu, którego już nie ma w widoku (pusty wiersz sprzątnięty po Cofnij) —
  // dane z arkusza od nowa, żeby wiersz wrócił do tabeli.
  function afterUndo(step) {
    if (!step || step.sheet !== currentSheetName) return false;
    const ws = workbook.Sheets[currentSheetName];
    const shown = new Set((baseRows || []).map((r) => r.rowIndex0));
    const lost = step.changes.some((ch) => !shown.has(ch.rowIndex0) && ch.rowIndex0 >= currentHeaderRow
      && typeof isMeaningfulSheetCell === "function" && isMeaningfulSheetCell(ws[ch.cellRef]));
    if (lost && typeof rebuildCurrentSheetData === "function") rebuildCurrentSheetData();
    const hdrR = currentHeaderRow - 1;
    if (!step.changes.some((ch) => ch.rowIndex0 === hdrR)) return false;
    if (!columnsChanged()) return false;
    const real = realCount();
    const grew = lastColOf(workbook.Sheets[currentSheetName]) - currentStartCol + 1 > real;
    reloadStructure({ shift: { from: grew ? real : real - 1, by: grew ? 1 : -1 } });
    return true;
  }

  let justAdded = null;
  // Puste wiersze robocze na końcu (oprócz tego z kursorem) — precz z widoku przed rysowaniem.
  function pruneEmptyDrafts() {
    if (typeof baseRows === "undefined" || !Array.isArray(baseRows)) return false;
    const focusKey = typeof focusedCellState !== "undefined" && focusedCellState ? focusedCellState.rowKey : "";
    let changed = false;
    while (baseRows.length) {
      const r = baseRows[baseRows.length - 1];
      if (!r || !r.isDraftRow || !isEmptyRow(r) || r === justAdded || getRowSelectionKey(r) === focusKey) break;
      baseRows.pop();
      const i = viewRows.indexOf(r);
      if (i >= 0) viewRows.splice(i, 1);
      changed = true;
    }
    return changed;
  }

  // ── „duchy”: blade puste wiersze pod danymi ────────────────────────────────
  const GHOSTS = 3;
  let ghostBody = null;
  function onGhostClick(e) {
    const td = e.target && e.target.closest ? e.target.closest("td") : null;
    if (!td || !ghostBody || !ghostBody.contains(td)) return;
    const col = Number(td.dataset.ghostCol);
    const tr = td.parentElement;
    const gap = tr && tr.parentElement === ghostBody ? Math.max(0, Array.prototype.indexOf.call(ghostBody.rows, tr)) : 0;
    if (!addDraft({ col: Number.isInteger(col) ? col : 0, gap })) return;
    requestAnimationFrame(() => {
      const cell = typeof findCellElement === "function" ? findCellElement(focusedCellState) : null;
      if (cell && typeof openCellEditor === "function") openCellEditor(cell);
    });
  }
  function syncGhosts() {
    const table = document.getElementById("dataTable");
    const tb = typeof tbodyEl !== "undefined" && tbodyEl ? tbodyEl : (table ? table.tBodies[0] : null);
    if (!table || !tb) return;
    if (!ghostBody) {
      ghostBody = document.createElement("tbody");
      ghostBody.className = "ghost-body";
      ghostBody.setAttribute("aria-hidden", "true"); // tylko podpowiedź wzrokowa; klawiatura: Enter / „＋ Nowy wiersz”
      ghostBody.addEventListener("click", onGhostClick);
    }
    if (!canAdd() || filtered() || !baseRows.length) { ghostBody.remove(); ghostBody.dataset.key = ""; return; }
    if (ghostBody.previousElementSibling !== tb) tb.after(ghostBody);
    const n = currentDisplayModel.headers.length;
    let last = -1;
    baseRows.forEach((r) => { if (Number.isFinite(r.rowIndex0) && r.rowIndex0 > last) last = r.rowIndex0; });
    const sample = tb.querySelector("tr[data-row-key]");
    const h = sample ? Math.round(sample.getBoundingClientRect().height) : 0;
    const key = `${n}|${last}|${h}`;
    if (ghostBody.dataset.key === key) return;
    ghostBody.dataset.key = key;
    const frag = document.createDocumentFragment();
    for (let g = 1; g <= GHOSTS; g++) {
      const tr = document.createElement("tr");
      tr.className = "ghost-row";
      if (h) tr.style.height = `${h}px`;
      const head = document.createElement("td");
      head.className = "row-head ghost-head";
      head.textContent = String(last + 1 + g); // numer wiersza w Excelu (rowIndex0 + 1)
      tr.appendChild(head);
      for (let i = 0; i < n; i++) {
        const td = document.createElement("td");
        td.dataset.ghostCol = String(i);
        tr.appendChild(td);
      }
      frag.appendChild(tr);
    }
    ghostBody.replaceChildren(frag);
  }

  // ── „duchy-kolumny”: 2 blade kolumny z prawej ──────────────────────────────
  const GHOST_COLS = 2;
  const GHOST_COL_W = 96;
  function syncGhostCols() {
    const table = document.getElementById("dataTable");
    if (!table) return;
    table.querySelectorAll("col.ghost-col, .ghost-col-cell").forEach((el) => el.remove());
    const derived = typeof dcAppliedCount === "number" ? dcAppliedCount : 0;
    if (!canAdd() || filtered() || derived > 0 || !baseRows.length) return;
    const cg = table.querySelector("colgroup");
    const head = typeof theadEl !== "undefined" ? theadEl : table.tHead;
    if (!cg || !head) return;
    const n = currentDisplayModel.headers.length;
    for (let g = 1; g <= GHOST_COLS; g++) {
      const col = document.createElement("col");
      col.className = "ghost-col";
      col.style.width = `${GHOST_COL_W}px`;
      cg.appendChild(col);
    }
    Array.from(head.rows).forEach((tr) => {
      for (let g = 1; g <= GHOST_COLS; g++) {
        const isGuide = tr.classList.contains("guide-row");
        const isHeader = tr.classList.contains("header-row");
        const cell = document.createElement(isGuide || isHeader ? "th" : "td");
        cell.className = "ghost-col-cell";
        cell.dataset.ghostNewCol = String(g);
        if (isGuide) cell.textContent = XLSX.utils.encode_col(currentStartCol + n + g - 1);
        else if (isHeader && g === 1) { cell.textContent = "＋"; cell.title = t("addColHint"); }
        tr.appendChild(cell);
      }
    });
    if (ghostBody) Array.from(ghostBody.rows).forEach((tr) => {
      for (let g = 1; g <= GHOST_COLS; g++) {
        const td = document.createElement("td");
        td.className = "ghost-col-cell";
        td.dataset.ghostNewCol = String(g);
        tr.appendChild(td);
      }
    });
  }
  // Klik w ducha-kolumnę: komórka-duch (nagłówek, wiersz-duch) albo puste pole za ostatnią komórką
  // wiersza danych (tam wiersz po prostu się kończy) → nowa kolumna i pisanie w tym wierszu.
  function onGhostColClick(e) {
    if (!canAdd() || filtered()) return; // przy filtrze duchów-kolumn nie ma — pusty margines nie dodaje kolumny
    const t0 = e.target;
    if (!t0 || !t0.closest) return;
    let rowKey = null;
    const cell = t0.closest(".ghost-col-cell");
    if (cell) {
      rowKey = cell.parentElement && cell.parentElement.dataset.rowKey ? cell.parentElement.dataset.rowKey : null;
    } else {
      // puste pole za ostatnią komórką: przeglądarka nie przypisuje kliku do wiersza (cel = tabela
      // albo tbody) — wiersz po wysokości, kolumna-duch = na prawo od ostatniego nagłówka danych
      if (t0.closest("td, th, input, button")) return;
      const lastTh = document.querySelector(`#dataTable thead tr.header-row th[data-col-index="${currentDisplayModel.headers.length - 1}"]`);
      if (!lastTh) return;
      const edge = lastTh.getBoundingClientRect().right;
      if (e.clientX <= edge) return;
      const hit = document.elementFromPoint(edge - 4, e.clientY);
      const tr = hit && hit.closest ? hit.closest("#dataTable tbody tr[data-row-key]") : null;
      if (!tr) return;
      rowKey = tr.dataset.rowKey;
    }
    e.stopPropagation();
    e.preventDefault();
    addColumn({ rowKey });
  }

  // ── przycisk „＋ Nowy wiersz” w pasku pod tabelą (obok sum) ─────────────────
  // POZA przewijaną tabelą: na dole przewijania ostatni wiersz ma stać przy krawędzi (pilnują
  // tego testy virt-*), a przycisk widać zawsze — nowy wiersz i tak przewija się do widoku.
  let btn = null;
  let colBtn = null;
  const label = (el, long, short) => {
    el.replaceChildren(
      Object.assign(document.createElement("span"), { className: "lbl-long", textContent: long }),
      Object.assign(document.createElement("span"), { className: "lbl-short", textContent: short }),
    );
  };
  function syncButton() {
    const can = canAdd();
    const f = can && filtered();
    if (btn) { btn.hidden = !can; btn.disabled = f; label(btn, t("addRowBtn"), t("addRowShort")); btn.title = t(f ? "addRowFiltered" : "addRowHint"); }
    if (colBtn) { colBtn.hidden = !can; label(colBtn, t("addColBtn"), t("addColShort")); colBtn.title = t("addColHint"); }
  }

  function init() {
    const foot = document.getElementById("tableFoot");
    if (foot) {
      btn = document.createElement("button");
      btn.type = "button";
      btn.id = "addRowBtn";
      btn.className = "add-row-btn";
      btn.hidden = true;
      btn.addEventListener("click", () => {
        if (!addDraft({ col: 0 })) return;
        // od razu pisanie w pierwszej kolumnie nowego wiersza (jak „Dodaj wiersz” w Arkuszach Google)
        requestAnimationFrame(() => {
          const td = typeof findCellElement === "function" ? findCellElement(focusedCellState) : null;
          if (td && typeof openCellEditor === "function") openCellEditor(td);
        });
      });
      colBtn = document.createElement("button");
      colBtn.type = "button";
      colBtn.id = "addColBtn";
      colBtn.className = "add-row-btn add-col-btn";
      colBtn.hidden = true;
      colBtn.addEventListener("click", () => { addColumn(); });
      foot.prepend(btn, colBtn);
    }
    const table = document.getElementById("dataTable");
    if (table) table.addEventListener("click", onGhostColClick, true);
    if (typeof window.handleFile === "function") {
      const orig = window.handleFile;
      window.handleFile = async function handleFileRowGrow(...args) {
        const r = await orig.apply(this, args);
        if (r && workbook) snapshot();
        return r;
      };
    }
    if (typeof window.renderActiveTable === "function") {
      const origRender = window.renderActiveTable;
      window.renderActiveTable = function renderActiveTableRowGrow(...args) {
        pruneEmptyDrafts();
        const r = origRender.apply(this, args);
        syncButton();
        syncGhosts(); // zwykle już po renderTable (niżej) — tu na wypadek rysowania bez renderTable
        syncGhostCols();
        return r;
      };
    }
    // Duchy dopisujemy ZARAZ po narysowaniu tabeli, nie dopiero po renderActiveTable: animacja
    // sortowania/filtra (flipRows w core.js) mierzy wiersze i klonuje <colgroup> tuż po renderTable —
    // bez duchów-kolumn w tej chwili kopie miały inny układ kolumn niż tabela (flip-ghost, 2026-11-11).
    if (typeof window.renderTable === "function") {
      const origTable = window.renderTable;
      window.renderTable = function renderTableRowGrow(...args) {
        const r = origTable.apply(this, args);
        syncGhosts();
        syncGhostCols();
        return r;
      };
    }
    syncButton();
  }

  return {
    init, snapshot, canAdd, filtered, addDraft, calcColumn, beforeWrite, withXf, growInfo, syncButton,
    addColumn, writeColumn, reloadStructure, afterUndo, isQuietReload: () => quiet, syncGhosts,
    _state: () => ({ origLast: [...origLast], liveLast: [...liveLast], drafts: [...drafts], origLastCol: [...origLastCol] }),
  };
})();

document.addEventListener("DOMContentLoaded", () => swbRowGrow.init());
