// =====================================================================
// Narzędzia edycji — masowe czyszczenie/przekształcanie wartości w kolumnie
// lub zaznaczeniu. Operacje:
//   (1) zdejmij śmieć wzorcem — reużywa compileGroupPattern / fuzzyGroupTransform
//       z analysis.js (selektory * ** # @ ? oraz = jako rdzeń);
//   (2) znajdź i zamień (dosłownie / regex);
//   (3) zmiana wielkości liter;
//   (4) przytnij / spacje (końce / zwiń wielokrotne / twarde spacje→zwykłe);
//   (5) prefiks / sufiks (doklejenie tekstu z przodu/z tyłu);
//   (6) wyrównaj długość (padStart/padEnd, np. zera wiodące w kodach);
//   (7) konwersja typu (tekst↔liczba↔data).
// Zapis idzie przez updateSheetCell -> pendingEdits -> ZIP-patch (zachowuje plik).
// Działa tylko w trybie "wide". "Znajdź i zamień" działa na tekście, datach
// (po WYŚWIETLANEJ wartości — user wpisuje to, co widzi) i liczbach (po WARTOŚCI
// SUROWEJ — String(raw), nie po formacie locale); po podmianie odtwarza typ.
// "Wzorzec", "wielkość liter", "przytnij", "prefiks/sufiks" i "wyrównaj" ruszają
// wyłącznie komórki tekstowe. Zakres "Kolumna" + przełącznik "tylko przefiltrowane
// wiersze" zawęża działanie do viewRows (widoczny po filtrze podzbiór).
// =====================================================================

const editScopeEl = document.getElementById("editScope");
const editColumnFieldEl = document.getElementById("editColumnField");
const editColumnSelectEl = document.getElementById("editColumnSelect");
const editOpEl = document.getElementById("editOp");
const editPatternFieldsEl = document.getElementById("editPatternFields");
const editPatternModeEl = document.getElementById("editPatternMode");
const editPatternInputFieldEl = document.getElementById("editPatternInputField");
const editPatternInputEl = document.getElementById("editPatternInput");
const editReplaceFieldsEl = document.getElementById("editReplaceFields");
const editFindEl = document.getElementById("editFind");
const editReplaceEl = document.getElementById("editReplace");
const editRegexEl = document.getElementById("editRegex");
const editCaseFieldsEl = document.getElementById("editCaseFields");
const editCaseModeEl = document.getElementById("editCaseMode");
const editTrimFieldsEl = document.getElementById("editTrimFields");
const editTrimModeEl = document.getElementById("editTrimMode");
const editAffixFieldsEl = document.getElementById("editAffixFields");
const editPrefixEl = document.getElementById("editPrefix");
const editSuffixEl = document.getElementById("editSuffix");
const editPadFieldsEl = document.getElementById("editPadFields");
const editPadLenEl = document.getElementById("editPadLen");
const editPadCharEl = document.getElementById("editPadChar");
const editPadSideEl = document.getElementById("editPadSide");
const editConvertFieldsEl = document.getElementById("editConvertFields");
const editConvertToEl = document.getElementById("editConvertTo");
const editFilteredOnlyFieldEl = document.getElementById("editFilteredOnlyField");
const editFilteredOnlyEl = document.getElementById("editFilteredOnly");
const applyEditToolBtnEl = document.getElementById("applyEditToolBtn");

// Lista kolumn (value = indeks w currentHeaders).
function populateEditColumnSelect() {
  if (!editColumnSelectEl) return;
  editColumnSelectEl.replaceChildren();
  if (!currentHeaders.length) {
    const opt = document.createElement("option");
    opt.value = "";
    opt.textContent = t("firstLoadSheet");
    editColumnSelectEl.appendChild(opt);
    editColumnSelectEl.disabled = true;
    return;
  }
  editColumnSelectEl.disabled = false;
  currentHeaders.forEach((header, i) => {
    const opt = document.createElement("option");
    opt.value = String(i);
    opt.textContent = header || `(${i + 1})`;
    editColumnSelectEl.appendChild(opt);
  });
}

// Pokazuje pola właściwe dla wybranej operacji / zakresu.
function syncEditToolFields() {
  if (!editOpEl) return;
  const op = editOpEl.value;
  editPatternFieldsEl.classList.toggle("hidden", op !== "pattern");
  editReplaceFieldsEl.classList.toggle("hidden", op !== "replace");
  editCaseFieldsEl.classList.toggle("hidden", op !== "case");
  if (editTrimFieldsEl) editTrimFieldsEl.classList.toggle("hidden", op !== "trim");
  if (editAffixFieldsEl) editAffixFieldsEl.classList.toggle("hidden", op !== "affix");
  if (editPadFieldsEl) editPadFieldsEl.classList.toggle("hidden", op !== "pad");
  if (editConvertFieldsEl) editConvertFieldsEl.classList.toggle("hidden", op !== "convert");
  if (op === "pattern") {
    editPatternInputFieldEl.classList.toggle("hidden", editPatternModeEl.value !== "pattern");
  }
  const isColumn = editScopeEl.value === "column";
  editColumnFieldEl.classList.toggle("hidden", !isColumn);
  // „Tylko przefiltrowane wiersze" ma sens jedynie dla zakresu kolumny
  // (zaznaczenie i tak jest już jawnym podzbiorem komórek).
  if (editFilteredOnlyFieldEl) editFilteredOnlyFieldEl.classList.toggle("hidden", !isColumn);
}

// Buduje funkcję transformującą tekst wg wybranej operacji.
// Zwraca { ok, fn } albo { ok:false, err } (klucz i18n błędu).
function buildEditTransform() {
  const op = editOpEl.value;
  if (op === "pattern") {
    if (editPatternModeEl.value === "fuzzy") return { ok: true, fn: fuzzyGroupTransform };
    const pat = (editPatternInputEl.value || "").trim();
    if (!pat) return { ok: false, err: "editErrNoPattern" };
    return { ok: true, fn: compileGroupPattern(pat) };
  }
  if (op === "replace") {
    const find = editFindEl.value;
    if (!find) return { ok: false, err: "editErrNoFind" };
    const repl = editReplaceEl.value;
    if (editRegexEl.checked) {
      let re;
      try { re = new RegExp(find, "gu"); } catch { return { ok: false, err: "editErrBadRegex" }; }
      return { ok: true, fn: (s) => s.replace(re, repl) };
    }
    return { ok: true, fn: (s) => s.split(find).join(repl) };
  }
  if (op === "case") {
    const locale = (typeof I18N !== "undefined" && I18N[currentLang] && I18N[currentLang].locale) || "pl-PL";
    const m = editCaseModeEl.value;
    if (m === "upper") return { ok: true, fn: (s) => s.toLocaleUpperCase(locale) };
    if (m === "lower") return { ok: true, fn: (s) => s.toLocaleLowerCase(locale) };
    return {
      ok: true,
      fn: (s) => s.replace(/\p{L}[\p{L}\p{M}]*/gu, (w) => w[0].toLocaleUpperCase(locale) + w.slice(1).toLocaleLowerCase(locale)),
    };
  }
  if (op === "trim") {
    const m = editTrimModeEl.value;
    if (m === "collapse") return { ok: true, fn: (s) => s.replace(/\s+/gu, " ").trim() };
    // „twarde spacje": NBSP / wąska NBSP / figure space → zwykła spacja, potem przytnij końce
    if (m === "hard") return { ok: true, fn: (s) => s.replace(/[\u00A0\u2007\u202F]/g, " ").trim() };
    return { ok: true, fn: (s) => s.trim() }; // „ends"
  }
  if (op === "affix") {
    const pre = editPrefixEl.value;
    const suf = editSuffixEl.value;
    if (!pre && !suf) return { ok: false, err: "editErrNoAffix" };
    return { ok: true, fn: (s) => pre + s + suf };
  }
  if (op === "pad") {
    const len = parseInt(editPadLenEl.value, 10);
    if (!Number.isFinite(len) || len < 1) return { ok: false, err: "editErrBadPadLen" };
    const ch = (editPadCharEl.value || " ").slice(0, 1) || " ";
    const side = editPadSideEl.value;
    return { ok: true, fn: (s) => (side === "end" ? s.padEnd(len, ch) : s.padStart(len, ch)) };
  }
  return { ok: false, err: "editToolNoChange" };
}

// Zbiera komórki docelowe {row, col} wg zakresu (kolumna / zaznaczenie).
function collectEditTargets() {
  const targets = [];
  if (editScopeEl.value === "selection") {
    const rect = getSelectionRectangle();
    if (!rect) return { err: "editErrNoSelection" };
    for (let r = rect.rowStart; r <= rect.rowEnd; r++) {
      const row = rect.model.rows[r];
      if (!row || row.isLongViewRow || row.isSubheader) continue;
      for (let c = rect.colMin; c <= rect.colMax; c++) targets.push({ row, col: c });
    }
    return { targets };
  }
  // kolumna — wszystkie wiersze danych arkusza (baseRows), albo — gdy zaznaczono
  // „tylko przefiltrowane" — wiersze widoczne po aktualnym filtrze (viewRows, te same
  // obiekty wierszy, tylko przefiltrowany podzbiór).
  const colIdx = parseInt(editColumnSelectEl.value, 10);
  if (!Number.isFinite(colIdx)) return { err: "editErrNoColumn" };
  const rowsSource = editFilteredOnlyEl && editFilteredOnlyEl.checked ? viewRows : baseRows;
  rowsSource.forEach((row) => {
    if (!row || row.isLongViewRow || row.isSubheader) return;
    targets.push({ row, col: colIdx });
  });
  return { targets };
}

function applyEditTool() {
  if (!workbook) { toast(t("noFileToSave"), "warning"); return; }
  if (!currentDisplayModel || currentDisplayModel.mode !== "wide") {
    toast(t("editWideOnly"), "info");
    return;
  }
  const op = editOpEl.value;
  const tg = collectEditTargets();
  if (tg.err) { toast(t(tg.err), "warning"); return; }

  let changed = 0;

  // Konwersja typu — osobna ścieżka, bo zmienia TYP komórki (nie transformuje stringa).
  if (op === "convert") {
    const to = editConvertToEl.value;
    tg.targets.forEach(({ row, col }) => {
      const raw = Array.isArray(row.values) ? row.values[col] : undefined;
      const shown = getDisplayValue(row, col);
      if (shown === "" || shown == null) return; // pustych nie ruszamy
      let newVal, newType;
      if (to === "number") {
        if (typeof raw === "number") return; // już liczba
        const n = parseLooseNumber(shown);
        if (n === null) return; // nie da się jednoznacznie — pomiń
        newVal = n; newType = "number";
      } else if (to === "date") {
        if (raw instanceof Date) return; // już data
        const d = parseDateFlexible(shown);
        if (!(d instanceof Date) || Number.isNaN(d.getTime())) return;
        newVal = d; newType = "date";
      } else { // text
        if (typeof raw === "string") return; // już tekst
        newVal = shown; newType = "string";
      }
      updateSheetCell(row.rowIndex0, col, { value: newVal, type: newType });
      row.values[col] = newVal;
      if (Array.isArray(row.rawValues)) row.rawValues[col] = newVal;
      if (Array.isArray(row.display)) row.display[col] = newVal == null ? "" : toDisplay(newVal);
      changed++;
    });
    finishEditTool(changed);
    return;
  }

  const tr = buildEditTransform();
  if (!tr.ok) { toast(t(tr.err), "warning"); return; }

  tg.targets.forEach(({ row, col }) => {
    const raw = Array.isArray(row.values) ? row.values[col] : undefined;
    const isDateCell = raw instanceof Date;
    const isNumberCell = typeof raw === "number" && Number.isFinite(raw);
    // "Znajdź i zamień" działa na:
    //   • tekście — po wartości surowej,
    //   • datach — po tym, co widać w siatce (dd-mm-yy / format z pliku), żeby
    //     user wpisywał to, co widzi,
    //   • liczbach — po WARTOŚCI SUROWEJ (String(raw), z kropką dziesiętną, bez
    //     separatorów i symboli waluty/%), a NIE po sformatowanym tekście — bo
    //     format locale (spacje, przecinek, „zł", „%") nie da się jednoznacznie
    //     re-sparsować z powrotem do liczby.
    // "wzorzec" / "wielkość liter": tylko czysty tekst (nie ruszamy liczb/dat).
    let source;
    if (op === "replace") {
      if (isDateCell) source = getDisplayValue(row, col);
      else if (isNumberCell) source = String(raw);
      else if (typeof raw === "string") source = raw;
      else return;
    } else {
      if (typeof raw !== "string") return; // tylko tekst — nie ruszamy liczb/dat
      source = raw;
    }
    let next;
    try { next = tr.fn(source); } catch { return; }
    if (typeof next !== "string" || next === source) return;

    if (op === "replace" && (isDateCell || isNumberCell)) {
      // Edycja daty/liczby — odtwórz typ z nowego tekstu (data→data, liczba→liczba);
      // jeśli się nie uda, zostaw zwykły tekst. Formuła przez "=..." zablokowana.
      const parsed = parseInputValue(next);
      if (parsed && parsed.type === "formula") return;
      const newVal = parsed ? parsed.value : next;
      const newType = parsed ? parsed.type : "string";
      updateSheetCell(row.rowIndex0, col, { value: newVal, type: newType });
      row.values[col] = newVal;
      if (Array.isArray(row.rawValues)) row.rawValues[col] = newVal;
      if (Array.isArray(row.display)) row.display[col] = newVal == null ? "" : toDisplay(newVal);
    } else {
      // Tekst (i wzorzec/wielkość liter) — wynik zostaje tekstem, by nie gubić
      // np. zer wiodących w kodach typu "00123".
      updateSheetCell(row.rowIndex0, col, { value: next, type: "string" });
      row.values[col] = next;
      if (Array.isArray(row.rawValues)) row.rawValues[col] = next;
      if (Array.isArray(row.display)) row.display[col] = toDisplay(next);
    }
    changed++;
  });

  finishEditTool(changed);
}

// Wspólne domknięcie po operacji: odśwież widok + komunikat z liczbą zmian.
function finishEditTool(changed) {
  if (changed > 0) {
    setDirtyState(true);
    renderActiveTable();
    toast(t("editToolApplied", { count: changed }), "success");
    log(`Narzedzia edycji: zmieniono ${changed} komorek`, "success");
  } else {
    toast(t("editToolNoChange"), "info");
  }
}

// Tolerancyjny parser liczby dla „Konwersji typu" (tekst→liczba): zdejmuje spacje
// i twarde spacje, normalizuje przecinek dziesiętny na kropkę. Odrzuca wszystko,
// co nie jest czystą liczbą (separatory tysięcy, symbole waluty/%, jednostki).
function parseLooseNumber(s) {
  const t = String(s).replace(/\s/g, "").replace(",", ".");
  if (!/^-?\d+(\.\d+)?$/.test(t)) return null;
  return Number(t);
}

if (applyEditToolBtnEl && editOpEl && editScopeEl) {
  editScopeEl.addEventListener("change", syncEditToolFields);
  editOpEl.addEventListener("change", syncEditToolFields);
  editPatternModeEl.addEventListener("change", syncEditToolFields);
  applyEditToolBtnEl.addEventListener("click", applyEditTool);
  syncEditToolFields();
}
