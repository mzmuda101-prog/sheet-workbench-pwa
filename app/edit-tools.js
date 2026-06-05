// =====================================================================
// Narzędzia edycji — masowe czyszczenie wartości w kolumnie lub zaznaczeniu.
// Operacje: (1) zdejmij śmieć wzorcem — reużywa silnik compileGroupPattern /
// fuzzyGroupTransform z analysis.js (selektory * ** # @ ? oraz = jako rdzeń);
// (2) znajdź i zamień (dosłownie / regex); (3) zmiana wielkości liter.
// Zapis idzie przez updateSheetCell -> pendingEdits -> ZIP-patch (zachowuje plik).
// Działa tylko w trybie "wide" i tylko na komórkach tekstowych (nie rusza liczb/dat).
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
  const tr = buildEditTransform();
  if (!tr.ok) { toast(t(tr.err), "warning"); return; }
  const tg = collectEditTargets();
  if (tg.err) { toast(t(tg.err), "warning"); return; }

  let changed = 0;
  tg.targets.forEach(({ row, col }) => {
    const raw = Array.isArray(row.values) ? row.values[col] : undefined;
    if (typeof raw !== "string") return; // tylko tekst — nie ruszamy liczb/dat
    let next;
    try { next = tr.fn(raw); } catch { return; }
    if (typeof next !== "string" || next === raw) return;
    updateSheetCell(row.rowIndex0, col, { value: next, type: "string" });
    row.values[col] = next;
    if (Array.isArray(row.display)) row.display[col] = toDisplay(next);
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
