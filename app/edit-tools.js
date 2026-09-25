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
//   (7) konwersja typu (tekst↔liczba↔data);
//   (8) ujednolić warianty — znajduje wartości różniące się tylko pisownią
//       (wielkość liter / spacje / . - / _ / polskie znaki) i scala je do jednej,
//       wybranej przez usera pisowni (propozycja = najczęstsza).
// Każda operacja najpierw buduje PLAN zmian (planEditChanges) — ten sam plan
// pokazuje „Podgląd" i wykonuje „Zastosuj", więc podgląd nie może się rozjechać
// z tym, co faktycznie zostanie zapisane.
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
const editIgnoreCaseEl = document.getElementById("editIgnoreCase");
const editWholeCellEl = document.getElementById("editWholeCell");
const editUnifyFieldsEl = document.getElementById("editUnifyFields");
const editUnifyModeEl = document.getElementById("editUnifyMode");
const editUnifyScanBtnEl = document.getElementById("editUnifyScanBtn");
const editUnifyListEl = document.getElementById("editUnifyList");
const previewEditToolBtnEl = document.getElementById("previewEditToolBtn");
const editPreviewEl = document.getElementById("editPreview");
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
  if (editUnifyFieldsEl) editUnifyFieldsEl.classList.toggle("hidden", op !== "unify");
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
    const ignoreCase = !!(editIgnoreCaseEl && editIgnoreCaseEl.checked);
    const whole = !!(editWholeCellEl && editWholeCellEl.checked);
    const flags = ignoreCase ? "giu" : "gu";
    if (editRegexEl.checked) {
      let re;
      // „Tylko cała komórka" = wzorzec zakotwiczony do całej wartości.
      try { re = new RegExp(whole ? `^(?:${find})$` : find, flags); } catch { return { ok: false, err: "editErrBadRegex" }; }
      return { ok: true, fn: (s) => s.replace(re, repl) };
    }
    if (whole) {
      const locale = editLocale();
      const want = ignoreCase ? find.toLocaleLowerCase(locale) : find;
      return { ok: true, fn: (s) => ((ignoreCase ? s.toLocaleLowerCase(locale) : s) === want ? repl : s) };
    }
    if (ignoreCase) {
      const re = new RegExp(escapeEditRegex(find), flags);
      return { ok: true, fn: (s) => s.replace(re, () => repl) }; // funkcja: „$" w zamienniku dosłownie
    }
    return { ok: true, fn: (s) => s.split(find).join(repl) };
  }
  if (op === "case") {
    const locale = editLocale();
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

// Wiarygodny zakres numeru seryjnego daty przy „Konwersji → Data": 1950-01-01 … 2100-12-31.
const EDIT_SERIAL_MIN = 18264;
const EDIT_SERIAL_MAX = 73415;

function editLocale() {
  return (typeof I18N !== "undefined" && I18N[currentLang] && I18N[currentLang].locale) || "pl-PL";
}

function escapeEditRegex(str) {
  return String(str).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

// Buduje plan zmian dla bieżących ustawień panelu, NICZEGO nie zmieniając.
// Zwraca { changes:[{row,col,before,value,type}] } albo { err } (klucz i18n).
function planEditChanges() {
  if (!workbook) return { err: "noFileToSave" };
  if (!currentDisplayModel || currentDisplayModel.mode !== "wide") return { err: "editWideOnly", level: "info" };
  const op = editOpEl.value;
  const tg = collectEditTargets();
  if (tg.err) return { err: tg.err };
  const changes = [];

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
        // Goła liczba to numer seryjny Excela TYLKO w wiarygodnym zakresie dat —
        // inaczej kwota 1500 zamieniłaby się w 1904-02-08, a rok 2024 w 1905-07-16.
        const bare = typeof raw === "number" ? raw : (/^\d+(\.\d+)?$/.test(String(shown).trim()) ? Number(shown) : null);
        if (bare !== null && !(bare >= EDIT_SERIAL_MIN && bare <= EDIT_SERIAL_MAX)) return;
        const d = parseDateFlexible(shown);
        if (!(d instanceof Date) || Number.isNaN(d.getTime())) return;
        newVal = d; newType = "date";
      } else { // text
        if (typeof raw === "string") return; // już tekst
        newVal = shown; newType = "string";
      }
      changes.push({ row, col, before: shown, value: newVal, type: newType });
    });
    return { changes };
  }

  // Ujednolić warianty — mapa klucz→wybrana pisownia pochodzi z listy grup,
  // którą user zbudował przyciskiem „Znajdź warianty" (i ewentualnie poprawił).
  if (op === "unify") {
    const choice = readUnifyChoices();
    if (!choice) return { err: "editErrUnifyScan" };
    if (!choice.map.size) return { changes };
    tg.targets.forEach(({ row, col }) => {
      const raw = Array.isArray(row.values) ? row.values[col] : undefined;
      if (typeof raw !== "string") return; // tylko tekst
      const target = choice.map.get(unifyKey(raw, choice.mode));
      if (target === undefined || target === raw) return;
      changes.push({ row, col, before: raw, value: target, type: "string" });
    });
    return { changes };
  }

  const tr = buildEditTransform();
  if (!tr.ok) return { err: tr.err };

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
      changes.push({
        row, col, before: source,
        value: parsed ? parsed.value : next,
        type: parsed ? parsed.type : "string",
      });
    } else {
      // Tekst (i wzorzec/wielkość liter) — wynik zostaje tekstem, by nie gubić
      // np. zer wiodących w kodach typu "00123".
      changes.push({ row, col, before: source, value: next, type: "string" });
    }
  });
  return { changes };
}

function applyEditTool() {
  const plan = planEditChanges();
  if (plan.err) { toast(t(plan.err), plan.level || "warning"); return; }
  plan.changes.forEach(({ row, col, value, type }) => {
    updateSheetCell(row.rowIndex0, col, { value, type });
    row.values[col] = value;
    if (Array.isArray(row.rawValues)) row.rawValues[col] = value;
    if (Array.isArray(row.display)) row.display[col] = value == null ? "" : toDisplay(value);
  });
  hideEditPreview();
  if (editOpEl.value === "unify" && plan.changes.length) clearUnifyList();
  finishEditTool(plan.changes.length);
}

// ── Podgląd ─────────────────────────────────────────────────────────────
// Pokazuje sumę i zgrupowane pary „przed → po" (z liczbą wystąpień).
const EDIT_PREVIEW_LIMIT = 8;

function showEditValue(v) {
  // Spacje na brzegach / podwójne są niewidoczne — oznacz je, bo przy unifikacji
  // i przycinaniu to często JEDYNA różnica między wariantami.
  const s = v instanceof Date ? toDisplay(v) : String(v ?? "");
  if (s === "") return t("editEmptyValue");
  return `„${s.replace(/^ +| +$| {2,}/g, (m) => "␣".repeat(m.length))}”`;
}

function previewEditTool() {
  if (!editPreviewEl) return;
  const plan = planEditChanges();
  if (plan.err) { hideEditPreview(); toast(t(plan.err), plan.level || "warning"); return; }
  editPreviewEl.replaceChildren();
  const sum = document.createElement("div");
  sum.className = "edit-preview-sum";
  sum.textContent = plan.changes.length
    ? t("editPreviewSum", { count: plan.changes.length })
    : t("editToolNoChange");
  editPreviewEl.appendChild(sum);
  const pairs = new Map();
  plan.changes.forEach((c) => {
    const from = showEditValue(c.before);
    const to = showEditValue(c.value);
    const k = `${from}\u0000${to}`;
    const p = pairs.get(k);
    if (p) p.n++; else pairs.set(k, { from, to, n: 1 });
  });
  const list = [...pairs.values()].sort((a, b) => b.n - a.n);
  list.slice(0, EDIT_PREVIEW_LIMIT).forEach((p) => {
    const rowEl = document.createElement("div");
    rowEl.className = "edit-preview-row";
    const from = document.createElement("span"); from.className = "from"; from.textContent = p.from;
    const arrow = document.createElement("span"); arrow.textContent = "→";
    const to = document.createElement("span"); to.textContent = p.to;
    const n = document.createElement("span"); n.className = "n"; n.textContent = `${p.n}×`;
    rowEl.append(from, arrow, to, n);
    editPreviewEl.appendChild(rowEl);
  });
  if (list.length > EDIT_PREVIEW_LIMIT) {
    const more = document.createElement("div");
    more.className = "unify-more";
    more.textContent = t("editPreviewMore", { count: list.length - EDIT_PREVIEW_LIMIT });
    editPreviewEl.appendChild(more);
  }
  editPreviewEl.classList.remove("hidden");
}

function hideEditPreview() {
  if (!editPreviewEl) return;
  editPreviewEl.classList.add("hidden");
  editPreviewEl.replaceChildren();
}

// ── Ujednolić warianty ─────────────────────────────────────────────────
// Klucz porównania: „case" = wielkość liter + spacje (zwinięte, przycięte);
// „loose" dodatkowo pomija spacje w ogóle, . - / _ , ; : oraz polskie znaki.
function unifyKey(s, mode) {
  let k = String(s).toLocaleLowerCase(editLocale()).replace(/[   ]/g, " ");
  if (mode === "case") return k.replace(/\s+/g, " ").trim();
  return k.normalize("NFD").replace(/[̀-ͯ]/g, "").replace(/ł/g, "l")
    .replace(/[\s.\-\/_,;:]+/g, "");
}

let editUnifyState = null; // { mode, groups:[{ key, variants:[{ value, count }] }] }
const UNIFY_GROUP_LIMIT = 150;

function scanUnifyVariants() {
  if (!workbook) { toast(t("noFileToSave"), "warning"); return; }
  if (!currentDisplayModel || currentDisplayModel.mode !== "wide") { toast(t("editWideOnly"), "info"); return; }
  const tg = collectEditTargets();
  if (tg.err) { toast(t(tg.err), "warning"); return; }
  const mode = editUnifyModeEl.value === "case" ? "case" : "loose";
  const byKey = new Map();
  tg.targets.forEach(({ row, col }) => {
    const raw = Array.isArray(row.values) ? row.values[col] : undefined;
    if (typeof raw !== "string" || !raw.trim()) return;
    const key = unifyKey(raw, mode);
    if (!key) return; // sama interpunkcja — nie ma czego scalać
    let g = byKey.get(key);
    if (!g) { g = new Map(); byKey.set(key, g); }
    g.set(raw, (g.get(raw) || 0) + 1);
  });
  const groups = [];
  byKey.forEach((variants, key) => {
    if (variants.size < 2) return;
    const list = [...variants].map(([value, count]) => ({ value, count }))
      .sort((a, b) => b.count - a.count || unifyNeatness(b.value) - unifyNeatness(a.value));
    groups.push({ key, variants: list, total: list.reduce((n, v) => n + v.count, 0) });
  });
  groups.sort((a, b) => b.total - a.total);
  editUnifyState = { mode, groups };
  hideEditPreview();
  renderUnifyList();
}

// Remis w liczbie wystąpień → wolimy „schludniejszą" pisownię: bez spacji na
// brzegach/podwójnych, z wielką literą na początku.
function unifyNeatness(v) {
  let n = 0;
  if (v === v.trim()) n += 2;
  if (!/ {2,}/.test(v)) n += 1;
  if (/^\p{Lu}/u.test(v)) n += 1;
  return n;
}

function renderUnifyList() {
  if (!editUnifyListEl) return;
  editUnifyListEl.replaceChildren();
  if (!editUnifyState) return;
  const { groups } = editUnifyState;
  if (!groups.length) {
    const p = document.createElement("p");
    p.className = "field-note";
    p.textContent = t("editUnifyNone");
    editUnifyListEl.appendChild(p);
    return;
  }
  groups.slice(0, UNIFY_GROUP_LIMIT).forEach((g, gi) => {
    const box = document.createElement("div");
    box.className = "unify-group";
    box.dataset.group = String(gi);
    const head = document.createElement("label");
    head.className = "unify-head";
    const cb = document.createElement("input");
    cb.type = "checkbox";
    cb.checked = true;
    cb.className = "unify-on";
    cb.addEventListener("change", () => { box.classList.toggle("off", !cb.checked); hideEditPreview(); });
    const title = document.createElement("span");
    title.textContent = t("editUnifyGroupHead", { variants: g.variants.length, count: g.total });
    head.append(cb, title);
    box.appendChild(head);
    g.variants.forEach((v, vi) => {
      const opt = document.createElement("label");
      opt.className = "unify-opt";
      const radio = document.createElement("input");
      radio.type = "radio";
      radio.name = `unify-g-${gi}`;
      radio.value = String(vi);
      radio.checked = vi === 0;
      radio.addEventListener("change", hideEditPreview);
      const val = document.createElement("span");
      val.className = "unify-val";
      val.textContent = showEditValue(v.value);
      const cnt = document.createElement("span");
      cnt.className = "unify-count";
      cnt.textContent = `${v.count}×`;
      opt.append(radio, val, cnt);
      box.appendChild(opt);
    });
    editUnifyListEl.appendChild(box);
  });
  if (groups.length > UNIFY_GROUP_LIMIT) {
    const more = document.createElement("p");
    more.className = "unify-more";
    more.textContent = t("editUnifyMoreGroups", { count: groups.length - UNIFY_GROUP_LIMIT });
    editUnifyListEl.appendChild(more);
  }
}

// Czyta z listy: które grupy są włączone i która pisownia ma zostać.
// null = lista jeszcze nie zbudowana (albo unieważniona zmianą zakresu/trybu).
function readUnifyChoices() {
  if (!editUnifyState || !editUnifyListEl) return null;
  const map = new Map();
  editUnifyListEl.querySelectorAll(".unify-group").forEach((box) => {
    const g = editUnifyState.groups[Number(box.dataset.group)];
    if (!g) return;
    const on = box.querySelector(".unify-on");
    if (on && !on.checked) return;
    const picked = box.querySelector('input[type="radio"]:checked');
    const v = g.variants[picked ? Number(picked.value) : 0];
    if (v) map.set(g.key, v.value);
  });
  return { mode: editUnifyState.mode, map };
}

function clearUnifyList() {
  editUnifyState = null;
  if (editUnifyListEl) editUnifyListEl.replaceChildren();
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
  if (previewEditToolBtnEl) previewEditToolBtnEl.addEventListener("click", previewEditTool);
  if (editUnifyScanBtnEl) editUnifyScanBtnEl.addEventListener("click", scanUnifyVariants);
  // Każda zmiana ustawień unieważnia podgląd; zmiana ZAKRESU lub trybu porównania
  // unieważnia też listę wariantów (była policzona dla innego zbioru komórek).
  const panelEl = document.getElementById("panel-edit-tools");
  if (panelEl) {
    panelEl.addEventListener("input", (e) => {
      if (e.target && e.target.closest && e.target.closest("#editUnifyList")) return;
      hideEditPreview();
    });
    panelEl.addEventListener("change", (e) => {
      const id = e.target && e.target.id;
      if (id === "editScope" || id === "editColumnSelect" || id === "editFilteredOnly" || id === "editUnifyMode") clearUnifyList();
      if (!(e.target && e.target.closest && e.target.closest("#editUnifyList"))) hideEditPreview();
    });
  }
  syncEditToolFields();
}
