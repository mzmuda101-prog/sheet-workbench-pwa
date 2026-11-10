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
// Cofnij zostawia pusty wiersz w widoku (jak w Excelu); znika przy przebudowie danych.

const swbRowGrow = (() => {
  const origLast = new Map(); // arkusz → ostatni wiersz z treścią w otwartym pliku (0-based, -1 = pusty)
  const liveLast = new Map(); // arkusz → ostatni wiersz, za którym zaczynają się „nowe” (rośnie)
  const drafts = new Set();   // `${arkusz}\t${r}` — wiersze z przycisku/Entera (dostają formuły)
  const newXf = new Map();    // arkusz → Map(ref → indeks stylu) dla komórek, których nie było w pliku

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

  // Nowy plik: zapamiętaj, gdzie kończyły się dane każdego arkusza (zanim cokolwiek dopiszemy).
  function snapshot() {
    origLast.clear(); liveLast.clear(); drafts.clear(); newXf.clear();
    if (typeof workbook === "undefined" || !workbook) return;
    (workbook.SheetNames || []).forEach((name) => {
      const last = lastRowOf(workbook.Sheets[name]);
      origLast.set(name, last);
      liveLast.set(name, last);
    });
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

  const isEmptyRow = (row) => !row || !Array.isArray(row.values) || row.values.every((v) => v === null || v === undefined || v === "");

  // Wiersz roboczy za ostatnim wierszem danych arkusza. Zwraca klucz wiersza (do fokusu) albo "".
  function addDraft({ col = 0 } = {}) {
    if (!canAdd()) return "";
    const name = currentSheetName;
    const tail = baseRows[baseRows.length - 1];
    // pusty wiersz roboczy już czeka na końcu — nie mnożymy pustych
    if (tail && tail.isDraftRow && isEmptyRow(tail)) {
      if (!viewRows.includes(tail)) { viewRows.push(tail); renderActiveTable(); }
      const k = getRowSelectionKey(tail);
      setFocusedCell(k, Math.max(0, Math.min(col, currentDisplayModel.headers.length - 1)), { scroll: true });
      return k;
    }
    let lastData = liveLast.has(name) ? liveLast.get(name) : lastRowOf(workbook.Sheets[name]);
    baseRows.forEach((r) => { if (Number.isFinite(r.rowIndex0) && r.rowIndex0 > lastData) lastData = r.rowIndex0; });
    // Tabela Excela z wierszem sumy na końcu danych: nowy wiersz pod sumą wypadłby z tabeli.
    // Excel wstawia wtedy NAD sumą — to wstawianie w środku (Etap 4); na razie mówimy wprost.
    if (totalsTableEndingAt(name, lastData)) { toast(t("addRowTotalsTable"), "info"); return ""; }
    const r = lastData + 1;
    const n = currentDisplayModel.headers.length;
    const tpl = baseRows.find((x) => x.rowIndex0 === lastData) || tail || null;
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
    drafts.add(keyOf(name, r));
    renderActiveTable();
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
  function normF(f) { try { return FC.serialize(FC.parse(f, { lang: "en" }), { lang: "en" }); } catch { return null; } }
  function shiftedF(f, dr) { try { return FC.serialize(FC.shift(FC.parse(f, { lang: "en" }), dr, 0), { lang: "en" }); } catch { return null; } }
  // Kolumna z „ciągłą” formułą: 3 wiersze nad nowym mają formułę, każda = poprzednia przesunięta o 1.
  function consistentFormula(ws, r, c) {
    const f = [1, 2, 3].map((d) => (r - d >= 0 ? ws[enc(r - d, c)] : null)).map((cell) => (cell && cell.f ? cell.f : null));
    if (f.some((x) => !x) || typeof FC === "undefined" || !FC) return false;
    const a = shiftedF(f[2], 1); const b = normF(f[1]);
    const c2 = shiftedF(f[1], 1); const d = normF(f[0]);
    return !!a && a === b && !!c2 && c2 === d;
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
    try { prepareRow(sheetName, ws, r, absC, k, last); } finally { preparing = false; }
  }

  function prepareRow(sheetName, ws, r, absC, k, last) {
    const styleMap = styleMapOf(sheetName);
    const xfMap = xfMapOf(sheetName);
    const tplR = r - 1;
    const n = currentDisplayModel && Array.isArray(currentDisplayModel.headers) ? currentDisplayModel.headers.length : 0;
    const realCols = [];
    for (let i = 0; i < n; i++) if (typeof cfmtIsRealCol !== "function" || cfmtIsRealCol(i)) realCols.push(i);
    // 1. styl z wiersza nad nim — dla każdej kolumny (pusta też ma ramkę/tło jak w Excelu)
    realCols.forEach((i) => {
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
    if (!drafts.has(k) || typeof swbFormulaCommitAt !== "function" || typeof swbFormulaShifted !== "function") return;
    realCols.forEach((i) => {
      const c = currentStartCol + i;
      if (c === absC) return; // tu użytkownik właśnie coś wpisuje
      const ref = enc(r, c);
      if (ws[ref] && ws[ref].f) return;
      if (!consistentFormula(ws, r, c)) return;
      const text = swbFormulaShifted(sheetName, enc(r - 1, c), 1, 0);
      if (text) swbFormulaCommitAt(r, i, text, { silent: true });
    });
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

  // Dla xlsx-grow.js: { arkusz: { fromRow, toRow } } (1-based), gdy dane urosły w dół.
  function growInfo() {
    const out = {};
    if (typeof workbook === "undefined" || !workbook) return out;
    origLast.forEach((orig, name) => {
      if (orig < 0) return; // pusty arkusz — nie ma zakresów kończących się na danych
      const cur = lastRowOf(workbook.Sheets[name]);
      if (cur > orig) out[name] = { fromRow: orig + 1, toRow: cur + 1 };
    });
    return out;
  }

  // ── przycisk „＋ Nowy wiersz” w pasku pod tabelą (obok sum) ─────────────────
  // POZA przewijaną tabelą: na dole przewijania ostatni wiersz ma stać przy krawędzi (pilnują
  // tego testy virt-*), a przycisk widać zawsze — nowy wiersz i tak przewija się do widoku.
  let btn = null;
  function syncButton() {
    if (!btn) return;
    btn.hidden = !canAdd();
    btn.textContent = t("addRowBtn");
    btn.title = t("addRowHint");
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
      foot.prepend(btn);
    }
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
        const r = origRender.apply(this, args);
        syncButton();
        return r;
      };
    }
    syncButton();
  }

  return { init, snapshot, canAdd, addDraft, beforeWrite, withXf, growInfo, syncButton, _state: () => ({ origLast: [...origLast], liveLast: [...liveLast], drafts: [...drafts] }) };
})();

document.addEventListener("DOMContentLoaded", () => swbRowGrow.init());
