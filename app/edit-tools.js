// =====================================================================
// Narzędzia edycji — masowe czyszczenie wartości w kolumnie lub zaznaczeniu.
// Operacje: (1) zdejmij śmieć wzorcem — reużywa silnik compileGroupPattern /
// fuzzyGroupTransform z analysis.js (selektory * ** # @ ? oraz = jako rdzeń);
// (2) znajdź i zamień (dosłownie / regex); (3) zmiana wielkości liter.
// Zapis idzie przez updateSheetCell -> pendingEdits -> ZIP-patch (zachowuje plik).
// Działa tylko w trybie "wide". "Znajdź i zamień" działa na tekście, datach
// (po WYŚWIETLANEJ wartości — user wpisuje to, co widzi) i liczbach (po WARTOŚCI
// SUROWEJ — String(raw), nie po formacie locale); po podmianie odtwarza typ.
// "Wzorzec" i "wielkość liter" ruszają wyłącznie komórki tekstowe.
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
  if (op === "pattern") {
    editPatternInputFieldEl.classList.toggle("hidden", editPatternModeEl.value !== "pattern");
  }
  editColumnFieldEl.classList.toggle("hidden", editScopeEl.value !== "column");
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
  // kolumna — wszystkie wiersze danych arkusza (baseRows)
  const colIdx = parseInt(editColumnSelectEl.value, 10);
  if (!Number.isFinite(colIdx)) return { err: "editErrNoColumn" };
  baseRows.forEach((row) => {
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
  const tr = buildEditTransform();
  if (!tr.ok) { toast(t(tr.err), "warning"); return; }
  const tg = collectEditTargets();
  if (tg.err) { toast(t(tg.err), "warning"); return; }

  let changed = 0;
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

  if (changed > 0) {
    setDirtyState(true);
    renderActiveTable();
    toast(t("editToolApplied", { count: changed }), "success");
    log(`Narzedzia edycji: zmieniono ${changed} komorek`, "success");
  } else {
    toast(t("editToolNoChange"), "info");
  }
}

if (applyEditToolBtnEl && editOpEl && editScopeEl) {
  editScopeEl.addEventListener("change", syncEditToolFields);
  editOpEl.addEventListener("change", syncEditToolFields);
  editPatternModeEl.addEventListener("change", syncEditToolFields);
  applyEditToolBtnEl.addEventListener("click", applyEditTool);
  syncEditToolFields();
}
