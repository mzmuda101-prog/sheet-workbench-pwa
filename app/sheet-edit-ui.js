// sheet-edit-ui.js — rama edycji „jak w Excelu” (wzór: Documents Workbench).
//
// 1. Tryb Edycja / Podgląd. Domyślnie Edycja (tak jak dotąd — komórki zmienia się od
//    razu). Podgląd = tylko oglądanie: blokuje każdą drogę zmiany (edytor, wklejanie,
//    wypełnianie, Delete, format, narzędzia edycji) i chowa paski edycji. Wybór
//    pamiętany na urządzeniu. Wskaźnik pod aktywnym trybem jeździ jak w Documents
//    (transform na kompozytorze, bez ruchu przy „Ogranicz ruch”).
// 2. Pasek „Narzędzia główne” (format → plik przy „Zapisz”, logika w cell-format.js).
// 3. Pasek formuły: adres aktywnej komórki (wpisz B12 + Enter = skok) i pełna
//    zawartość — formuła z pliku albo wartość — do edycji jak w Excelu.
// 4. Skróty: Ctrl/⌘+B/I/U, Ctrl+5, Delete/Backspace (wyczyść zawartość), F2 (edytuj).

const SWB_MODE_KEY = "swb-edit-mode-v1";
const SWB_LAST_COLOR_KEY = "swb-fmt-last-color-v1";
const SWB_LAST_FILL_KEY = "swb-fmt-last-fill-v1";
const SWB_FONT_SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 36];

let swbEditMode = (() => {
  try { return localStorage.getItem(SWB_MODE_KEY) === "view" ? "view" : "edit"; } catch { return "edit"; }
})();
let swbGuardToastAt = 0;

// Jedyna bramka dla wszystkich zmian w arkuszu. W Podglądzie — krótki komunikat
// (najwyżej co 2,5 s, żeby pisanie po klawiaturze nie sypało toastami) i odmowa.
function swbEditGuard() {
  if (swbEditMode !== "view") return true;
  const now = Date.now();
  if (now - swbGuardToastAt > 2500) {
    swbGuardToastAt = now;
    toast(t("viewModeBlocked"), "info");
  }
  return false;
}

function swbIsEditMode() { return swbEditMode !== "view"; }

var sheetEditUi = (() => {
  const barsEl = document.getElementById("editBars");
  const ribbonEl = document.getElementById("formatRibbon");
  const modeSwitchEl = document.getElementById("sheetModeSwitch");
  const modeBtns = modeSwitchEl ? Array.from(modeSwitchEl.querySelectorAll(".mode-btn")) : [];
  const nameBox = document.getElementById("nameBox");
  const fxInput = document.getElementById("formulaInput");
  const fxOk = document.getElementById("formulaOk");
  const fxCancel = document.getElementById("formulaCancel");
  const sizeSel = document.getElementById("fmtSize");
  const numSel = document.getElementById("fmtNumFormat");
  const colorBtn = document.getElementById("fmtColorBtn");
  const colorMore = document.getElementById("fmtColorMore");
  const colorSwatch = document.getElementById("fmtColorSwatch");
  const fillBtn = document.getElementById("fmtFillBtn");
  const fillMore = document.getElementById("fmtFillMore");
  const fillSwatch = document.getElementById("fmtFillSwatch");
  const borderBtn = document.getElementById("fmtBorderBtn");
  const painterBtn = document.getElementById("fmtPainterBtn");
  const clearBtn = document.getElementById("fmtClearBtn");
  const decLess = document.getElementById("fmtDecLess");
  const decMore = document.getElementById("fmtDecMore");
  if (!barsEl || !ribbonEl) return { sync() {}, setMode() {}, get mode() { return swbEditMode; } };

  const readLS = (k, d) => { try { return localStorage.getItem(k) || d; } catch { return d; } };
  const writeLS = (k, v) => { try { localStorage.setItem(k, v); } catch { /* prywatne okno */ } };
  let lastColor = readLS(SWB_LAST_COLOR_KEY, "C00000");
  let lastFill = readLS(SWB_LAST_FILL_KEY, "FFFF00");

  // ── tryb: wskaźnik jak w Documents ─────────────────────────────────────────
  const modePill = modeSwitchEl ? document.createElement("span") : null;
  if (modePill) {
    modePill.className = "mode-pill";
    modePill.setAttribute("aria-hidden", "true");
    modeSwitchEl.prepend(modePill);
    modeSwitchEl.classList.add("has-pill");
  }
  let pillAt = null;
  function placeModePill(animate) {
    const btn = modeBtns.find((b) => b.classList.contains("is-active"));
    if (!modePill || !btn || !btn.offsetWidth) return;
    const at = { x: btn.offsetLeft, w: btn.offsetWidth };
    modePill.style.width = `${at.w}px`;
    modePill.style.transform = `translateX(${at.x}px)`;
    const from = pillAt;
    pillAt = at;
    if (!animate || !from || (from.x === at.x && from.w === at.w) || !modePill.animate) return;
    if (matchMedia("(prefers-reduced-motion: reduce)").matches) return;
    modePill.getAnimations().forEach((a) => a.cancel());
    // w spoczynku aktywny przycisk ma własne tło (to samo co wskaźnik); w trakcie ruchu
    // przezroczyste, żeby było widać jadący wskaźnik
    modeSwitchEl.classList.add("pill-moving");
    const anim = modePill.animate([
      { transform: `translateX(${from.x}px) scaleX(${from.w / at.w})` },
      { transform: `translateX(${at.x}px)` },
    ], { duration: 320, easing: "cubic-bezier(0.22, 1, 0.36, 1)" });
    const done = () => modeSwitchEl.classList.remove("pill-moving");
    anim.onfinish = done;
    anim.oncancel = done;
  }
  if (modePill && typeof ResizeObserver === "function") new ResizeObserver(() => placeModePill(false)).observe(modeSwitchEl);

  function setMode(mode, { announce = false } = {}) {
    const next = mode === "view" ? "view" : "edit";
    const changed = next !== swbEditMode;
    if (next === "view" && typeof flushCellEditor === "function") flushCellEditor(); // otwarte pole — zatwierdź, nie gub
    swbEditMode = next;
    writeLS(SWB_MODE_KEY, next);
    document.documentElement.classList.toggle("swb-view-mode", next === "view");
    modeBtns.forEach((b) => {
      const on = b.dataset.mode === next;
      b.classList.toggle("is-active", on);
      b.setAttribute("aria-pressed", String(on));
    });
    placeModePill(changed);
    if (next === "view") stopPainter();
    sync();
    if (changed && announce) toast(t(next === "view" ? "viewModeOn" : "editModeOn"), "info");
  }
  modeBtns.forEach((b) => b.addEventListener("click", () => setMode(b.dataset.mode, { announce: true })));

  // ── etykiety (język) ───────────────────────────────────────────────────────
  let labelsLang = "";
  function applyLabels() {
    if (labelsLang === currentLang) return;
    labelsLang = currentLang;
    modeBtns.forEach((b) => { b.textContent = t(b.dataset.mode === "view" ? "modeView" : "modeEdit"); });
    if (modeSwitchEl) modeSwitchEl.setAttribute("aria-label", t("modeGroupAria"));
    ribbonEl.setAttribute("aria-label", t("fmtRibbonAria"));
    const aria = {
      '[data-fmt="b"]': "fmtBold", '[data-fmt="i"]': "fmtItalic", '[data-fmt="u"]': "fmtUnderline",
      '[data-fmt="strike"]': "fmtStrike", '[data-fmt="wrap"]': "fmtWrap",
      '[data-align="left"]': "fmtAlignLeft", '[data-align="center"]': "fmtAlignCenter", '[data-align="right"]': "fmtAlignRight",
      "#fmtColorBtn": "fmtFontColor", "#fmtColorMore": "fmtFontColorPick", "#fmtFillBtn": "fmtFillColor",
      "#fmtFillMore": "fmtFillColorPick", "#fmtBorderBtn": "fmtBorders", "#fmtPainterBtn": "fmtPainter",
      "#fmtClearBtn": "fmtClear", "#fmtDecLess": "fmtDecLess", "#fmtDecMore": "fmtDecMore",
      "#fmtSize": "fmtSize", "#fmtNumFormat": "fmtNumFormat", "#nameBox": "fmtNameBox",
      "#formulaInput": "fmtFormulaInput", "#formulaOk": "fmtFormulaOk", "#formulaCancel": "fmtFormulaCancel",
    };
    Object.entries(aria).forEach(([sel, key]) => {
      const el = barsEl.querySelector(sel);
      if (el) el.setAttribute("aria-label", t(key));
    });
    if (nameBox) nameBox.placeholder = t("fmtNameBoxPh");
    // listy wyboru
    if (sizeSel) {
      const cur = sizeSel.value;
      sizeSel.replaceChildren(...[["", "—"], ...SWB_FONT_SIZES.map((n) => [String(n), String(n)])].map(([v, l]) => {
        const o = document.createElement("option");
        o.value = v; o.textContent = l;
        return o;
      }));
      sizeSel.value = cur;
    }
    if (numSel && typeof CFMT_NUMBER_FORMATS !== "undefined") {
      numSel.replaceChildren(...CFMT_NUMBER_FORMATS.map((f) => {
        const o = document.createElement("option");
        o.value = f.id; o.textContent = t(`numFmt_${f.id}`);
        return o;
      }));
      const custom = document.createElement("option");
      custom.value = "custom"; custom.hidden = true;
      numSel.appendChild(custom);
    }
    paintSwatches();
  }

  function paintSwatches() {
    if (colorSwatch) colorSwatch.style.background = `#${lastColor}`;
    if (fillSwatch) fillSwatch.style.background = `#${lastFill}`;
  }

  // ── stan przycisków + pasek formuły ────────────────────────────────────────
  let fxBaseline = "";
  let fxDirty = false;
  let fxCellKey = "";

  function hasData() {
    return !!(typeof workbook !== "undefined" && workbook && currentHeaders && currentHeaders.length);
  }

  function cellContentText(a) {
    if (!a) return "";
    if (a.cell && a.cell.f) return typeof swbFormulaText === "function" ? swbFormulaText(currentSheetName, a.ref) : `=${a.cell.f}`;
    return typeof cellEditString === "function" ? cellEditString(a.row, a.col) : String(a.row.values?.[a.col] ?? "");
  }

  function setPressed(sel, on) {
    const el = ribbonEl.querySelector(sel);
    if (el) el.setAttribute("aria-pressed", on ? "true" : "false");
  }

  function sync() {
    applyLabels();
    const trOpen = !!document.querySelector("#transcribeOverlay:not(.hidden)");
    const show = swbIsEditMode() && hasData() && !trOpen;
    if (barsEl.hidden === show || ribbonEl.hidden === show) {
      barsEl.hidden = !show;
      ribbonEl.hidden = !show; // pasek formatu mieszka w rzędzie 2 (#tableActions)
      if (typeof syncTableViewportHeight === "function") syncTableViewportHeight();
    }
    if (typeof frameCompact !== "undefined") frameCompact.sync();
    if (!show) return;
    const wide = currentDisplayModel && currentDisplayModel.mode === "wide";
    barsEl.classList.toggle("is-long-view", !wide);
    const a = wide && typeof cfmtActiveCell === "function" ? cfmtActiveCell() : null;
    const s = (a && a.style) || {};
    setPressed('[data-fmt="b"]', !!s.bold);
    setPressed('[data-fmt="i"]', !!s.italic);
    setPressed('[data-fmt="u"]', !!s.underline);
    setPressed('[data-fmt="strike"]', !!s.strike);
    setPressed('[data-fmt="wrap"]', !!s.userWrap);
    const h = String(s.horizontal || "").toLowerCase();
    setPressed('[data-align="left"]', h === "left");
    setPressed('[data-align="center"]', h === "center" || h === "centercontinuous");
    setPressed('[data-align="right"]', h === "right");
    ribbonEl.classList.toggle("is-disabled", !a || !a.real);
    // bez aktywnej komórki pasek jest nieaktywny także dla czytnika ekranu (i dla WCAG)
    if (!a || !a.real) ribbonEl.setAttribute("aria-disabled", "true"); else ribbonEl.removeAttribute("aria-disabled");
    if (sizeSel && document.activeElement !== sizeSel) {
      const base = typeof getWorkbookDefaultFontSize === "function" ? getWorkbookDefaultFontSize(workbook) : 11;
      const pt = s.fontSizePt || (s.fontScale ? Math.round(s.fontScale * base * 2) / 2 : (a ? base : 0));
      const v = pt ? String(pt) : "";
      if (v && !Array.from(sizeSel.options).some((o) => o.value === v)) {
        const o = document.createElement("option");
        o.value = v; o.textContent = v;
        sizeSel.appendChild(o);
      }
      sizeSel.value = v;
    }
    if (numSel && document.activeElement !== numSel) {
      const code = a ? a.numFmt : "General";
      const known = (typeof CFMT_NUMBER_FORMATS !== "undefined" ? CFMT_NUMBER_FORMATS : []).find((f) => f.code === code);
      const custom = numSel.querySelector('option[value="custom"]');
      if (known) {
        if (custom) custom.hidden = true;
        numSel.value = known.id;
      } else if (custom) {
        custom.hidden = false;
        custom.textContent = String(code);
        numSel.value = "custom";
      }
    }
    if (painterBtn) painterBtn.setAttribute("aria-pressed", painterFmt ? "true" : "false");

    // pasek formuły (komórki nad nagłówkiem to zwykłe „pre:N” — cfmtActiveCell je zna)
    const ref = a ? a.ref : "";
    const key = ref ? `${currentSheetName}|${ref}` : "";
    if (nameBox && document.activeElement !== nameBox) {
      nameBox.value = ref;
    }
    if (fxInput) {
      const editorOpen = !!document.querySelector("#dataTable input.cell-editor");
      if (key !== fxCellKey || (!fxDirty && document.activeElement !== fxInput && !editorOpen)) {
        if (key !== fxCellKey) fxDirty = false;
        fxCellKey = key;
        fxBaseline = cellContentText(a);
        if (!editorOpen) fxInput.value = fxBaseline;
        setFxDirty(false);
      }
      fxInput.disabled = !a;
      fxInput.readOnly = !!(a && !a.real);
      fxInput.placeholder = !wide ? t("fmtLongViewNote") : (a && !a.real ? t("fmtDerivedNote") : "");
    }
  }

  function setFxDirty(on) {
    fxDirty = !!on;
    if (fxOk) fxOk.hidden = !fxDirty;
    if (fxCancel) fxCancel.hidden = !fxDirty;
    if (barsEl) barsEl.classList.toggle("fx-editing", fxDirty);
  }

  // ── akcje paska ────────────────────────────────────────────────────────────
  // Klik myszą/palcem w przycisk paska oddaje fokus siatce (jak wstążka Excela) — Delete,
  // strzałki i pisanie działają dalej na komórce. Z klawiatury (detail 0) fokus zostaje.
  ribbonEl.addEventListener("click", (e) => {
    const btn = e.target.closest("button");
    if (!btn || !ribbonEl.contains(btn)) return;
    if (e.detail > 0 && !btn.hasAttribute("aria-haspopup") && btn.id !== "fmtSumBtn") setTimeout(returnFocusToGrid, 0);
    if (btn.dataset.fmt) {
      if (btn.dataset.fmt === "wrap") {
        const a = cfmtActiveCell();
        cfmtApply({ wrap: !(a && a.style && a.style.userWrap) });
      } else cfmtToggle(btn.dataset.fmt);
      return;
    }
    if (btn.dataset.align) {
      const a = cfmtActiveCell();
      const cur = String(a?.style?.horizontal || "").toLowerCase();
      cfmtApply({ h: cur === btn.dataset.align ? null : btn.dataset.align }); // drugi klik = „Ogólne”
    }
  });
  if (colorBtn) colorBtn.addEventListener("click", () => cfmtApply({ color: lastColor }));
  if (fillBtn) fillBtn.addEventListener("click", () => cfmtApply({ fill: lastFill }));
  if (colorMore) colorMore.addEventListener("click", () => togglePop(colorMore, "color"));
  if (fillMore) fillMore.addEventListener("click", () => togglePop(fillMore, "fill"));
  if (borderBtn) borderBtn.addEventListener("click", () => togglePop(borderBtn, "border"));
  if (clearBtn) clearBtn.addEventListener("click", () => cfmtApply({ reset: true }));
  if (decLess) decLess.addEventListener("click", () => changeDecimals(-1));
  if (decMore) decMore.addEventListener("click", () => changeDecimals(1));
  if (sizeSel) sizeSel.addEventListener("change", () => {
    const n = Number(sizeSel.value);
    if (n > 0) cfmtApply({ sz: n });
    returnFocusToGrid();
  });
  if (numSel) numSel.addEventListener("change", () => {
    const f = CFMT_NUMBER_FORMATS.find((x) => x.id === numSel.value);
    if (f) cfmtApply({ numFmt: f.code });
    returnFocusToGrid();
  });

  function changeDecimals(delta) {
    const a = cfmtActiveCell();
    if (!a) { cfmtApply({}); return; }
    // jak w Excelu: kod liczony z aktywnej komórki, ten sam trafia do całego zaznaczenia
    cfmtApply({ numFmt: cfmtChangeDecimals(a.numFmt, delta) });
  }

  function returnFocusToGrid() {
    const td = typeof findCellElement === "function" ? findCellElement(focusedCellState) : null;
    if (td) td.focus({ preventScroll: true });
  }

  // ── wyskakujące: kolory i obramowania ──────────────────────────────────────
  const pop = document.createElement("div");
  pop.className = "app-menu fmt-pop";
  pop.id = "fmtPop";
  pop.hidden = true;
  pop.setAttribute("role", "dialog");
  document.body.appendChild(pop);
  let popAnchor = null;

  function closePop(returnFocus) {
    if (pop.hidden) return;
    pop.hidden = true;
    if (popAnchor) {
      popAnchor.setAttribute("aria-expanded", "false");
      if (returnFocus) popAnchor.focus();
    }
    popAnchor = null;
  }

  function placePop(anchor) {
    const r = anchor.getBoundingClientRect();
    const w = pop.offsetWidth || 240;
    const left = Math.max(8, Math.min(r.left, window.innerWidth - w - 8));
    pop.style.left = `${Math.round(left)}px`;
    pop.style.top = `${Math.round(r.bottom + 6)}px`;
  }

  function swatchButton(hex, onPick) {
    const b = document.createElement("button");
    b.type = "button";
    b.className = "fmt-pop-swatch";
    b.style.background = `#${hex}`;
    b.setAttribute("aria-label", `#${hex}`);
    b.title = `#${hex}`;
    b.addEventListener("click", () => onPick(hex));
    return b;
  }

  function buildColorPop(kind) {
    const isFill = kind === "fill";
    const pick = (hex) => {
      if (hex) {
        if (isFill) { lastFill = hex; writeLS(SWB_LAST_FILL_KEY, hex); } else { lastColor = hex; writeLS(SWB_LAST_COLOR_KEY, hex); }
        paintSwatches();
      }
      closePop(false);
      cfmtApply(isFill ? { fill: hex || null } : { color: hex || null });
      returnFocusToGrid();
    };
    const head = document.createElement("button");
    head.type = "button";
    head.className = "app-menu-item fmt-pop-auto";
    head.textContent = t(isFill ? "fmtNoFill" : "fmtAutoColor");
    head.addEventListener("click", () => pick(null));
    const grid = document.createElement("div");
    grid.className = "fmt-pop-grid";
    CFMT_PALETTE.forEach((hex) => grid.appendChild(swatchButton(hex, pick)));
    const more = document.createElement("label");
    more.className = "app-menu-item fmt-pop-more";
    const input = document.createElement("input");
    input.type = "color";
    input.value = `#${isFill ? lastFill : lastColor}`.toLowerCase();
    input.addEventListener("change", () => pick(input.value.replace(/^#/, "").toUpperCase()));
    const span = document.createElement("span");
    span.textContent = t("fmtMoreColors");
    more.append(input, span);
    pop.setAttribute("aria-label", t(isFill ? "fmtFillColor" : "fmtFontColor"));
    pop.replaceChildren(head, grid, more);
  }

  const BORDER_PRESETS = [
    ["all", "M3 3h18v18H3zM3 12h18M12 3v18"],
    ["outer", "M3 3h18v18H3z"],
    ["thickOuter", "M3 3h18v18H3z"],
    ["bottom", "M3 21h18"],
    ["top", "M3 3h18"],
    ["left", "M3 3v18"],
    ["right", "M21 3v18"],
    ["none", "M3 3h18v18H3z"],
  ];
  function buildBorderPop() {
    pop.setAttribute("aria-label", t("fmtBorders"));
    pop.replaceChildren(...BORDER_PRESETS.map(([id, d]) => {
      const b = document.createElement("button");
      b.type = "button";
      b.className = "app-menu-item fmt-pop-border";
      const sw = id === "thickOuter" ? 3 : 2;
      const dash = id === "none" ? ' stroke-dasharray="2 3"' : "";
      b.innerHTML = `<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="${sw}"${dash} aria-hidden="true"><path d="${d}"/></svg>`;
      const label = document.createElement("span");
      label.className = "app-menu-item-text";
      label.textContent = t(`fmtBorder_${id}`);
      b.appendChild(label);
      b.addEventListener("click", () => {
        closePop(false);
        cfmtApply((ctx) => cfmtBorderDelta(id, ctx, null));
        returnFocusToGrid();
      });
      return b;
    }));
  }

  function togglePop(anchor, kind) {
    if (!pop.hidden && popAnchor === anchor) { closePop(true); return; }
    closePop(false);
    if (kind === "border") buildBorderPop(); else buildColorPop(kind);
    pop.dataset.kind = kind;
    pop.hidden = false;
    popAnchor = anchor;
    anchor.setAttribute("aria-expanded", "true");
    placePop(anchor);
    const first = pop.querySelector("button, input");
    if (first && document.documentElement.classList.contains("kbd-focus")) first.focus();
  }
  document.addEventListener("pointerdown", (e) => {
    if (pop.hidden) return;
    if (pop.contains(e.target) || (popAnchor && popAnchor.contains(e.target))) return;
    closePop(false);
  }, true);
  pop.addEventListener("keydown", (e) => {
    if (e.key === "Escape") { e.preventDefault(); e.stopPropagation(); closePop(true); }
  });
  window.addEventListener("resize", () => { if (!pop.hidden && popAnchor) placePop(popAnchor); }, { passive: true });

  // ── malarz formatów ────────────────────────────────────────────────────────
  let painterFmt = null;
  function stopPainter() {
    painterFmt = null;
    document.documentElement.classList.remove("swb-painting");
    if (painterBtn) painterBtn.setAttribute("aria-pressed", "false");
  }
  if (painterBtn) painterBtn.addEventListener("click", () => {
    if (painterFmt) { stopPainter(); return; }
    if (!swbEditGuard()) return;
    const f = cfmtCaptureFormat();
    if (!f) { toast(t("fmtPickCell"), "info"); return; }
    painterFmt = f;
    document.documentElement.classList.add("swb-painting");
    painterBtn.setAttribute("aria-pressed", "true");
    toast(t("fmtPainterArmed"), "info");
  });
  // Klik/tap w komórkę (albo koniec przeciągania zakresu) = nałóż i wyłącz.
  if (typeof tbodyEl !== "undefined" && tbodyEl) {
    tbodyEl.addEventListener("pointerup", (e) => {
      if (!painterFmt || !e.target.closest("td") || e.target.closest("td.row-head")) return;
      const f = painterFmt;
      setTimeout(() => {
        stopPainter();
        if (cfmtApply(f)) toast(t("fmtPainterDone"), "success");
      }, 0);
    });
  }

  // ── pasek formuły: edycja ──────────────────────────────────────────────────
  function activeForCommit() {
    const a = typeof cfmtActiveCell === "function" ? cfmtActiveCell() : null;
    if (!a || !a.real) return null;
    return a;
  }

  function commitFx(move) {
    const a = activeForCommit();
    if (!a) { setFxDirty(false); return true; }
    const text = fxInput.value;
    if (text === fxBaseline) { setFxDirty(false); afterCommitMove(move); return true; }
    if (!swbEditGuard()) { revertFx(); return false; }
    if (a.row.isSubheader) { toast(t("editBlockedRow"), "info"); revertFx(); return false; }
    const parsed = typeof parseCellInput === "function" ? parseCellInput(text, a.row.rowIndex0, a.col) : null;
    if (parsed && parsed.type === "formula" && typeof swbFormulaCommit !== "function") {
      toast(t("formulaEditBlocked"), "warning");
      return false;
    }
    // Lista z Data Validation — ta sama reguła co w edytorze w komórce.
    let dvRule = null;
    try {
      if (typeof getCellDataValidation === "function") dvRule = getCellDataValidation(currentSheetName, a.row.rowIndex0, currentStartCol + a.col);
    } catch { dvRule = null; }
    if (dvRule && !(parsed && parsed.type === "formula")) {
      const str = parsed && parsed.value != null ? String(parsed.value).trim() : "";
      const ok = str === "" ? dvRule.allowBlank : dvRule.values.some((v) => String(v).trim().toLowerCase() === str.toLowerCase());
      if (!ok) {
        if (dvRule.mode === "stop") { toast(t("dvRejected"), "warning"); return false; }
        if (dvRule.mode === "warning") toast(t("dvWarning"), "warning");
      }
    }
    let res;
    if (parsed && parsed.type === "formula") res = swbFormulaCommit(a.row, a.col, text);
    else res = writeCellFromInput(currentDisplayModel, a.row, a.col, text);
    if (res !== "ok") {
      if (res === "formula") toast(t("formulaEditBlocked"), "warning");
      return false;
    }
    setDirtyState(true);
    setFxDirty(false);
    fxBaseline = text;
    const td = findCellElement(focusedCellState);
    if (td) {
      td.textContent = typeof getDisplayValue === "function" ? getDisplayValue(a.row, a.col) : String(a.row.display?.[a.col] ?? "");
      if (typeof flashCellSaved === "function") flashCellSaved(td);
    }
    if (typeof rememberColumnValue === "function") rememberColumnValue(a.col, a.row.values?.[a.col]);
    afterCommitMove(move);
    updateCellStats();
    return true;
  }

  function afterCommitMove(move) {
    if (!move) return;
    if (move.row || move.col) moveFocusedCell(move.row || 0, move.col || 0);
    returnFocusToGrid();
  }

  function revertFx() {
    fxInput.value = fxBaseline;
    setFxDirty(false);
  }

  if (fxInput) {
    fxInput.addEventListener("input", () => setFxDirty(fxInput.value !== fxBaseline));
    fxInput.addEventListener("focus", () => {
      if (!activeForCommit() && fxInput.readOnly) return;
      if (swbEditMode === "view") fxInput.readOnly = true;
    });
    fxInput.addEventListener("keydown", (e) => {
      e.stopPropagation(); // to pole tekstowe — skróty tabeli nie dotyczą
      if (e.key === "Enter") { e.preventDefault(); commitFx({ row: e.shiftKey ? -1 : 1 }); }
      else if (e.key === "Tab") { e.preventDefault(); commitFx({ col: e.shiftKey ? -1 : 1 }); }
      else if (e.key === "Escape") { e.preventDefault(); revertFx(); returnFocusToGrid(); }
    });
    fxInput.addEventListener("blur", () => {
      // klik w ✓/✕ obsługują przyciski (pointerdown zatrzymuje blur-commit niżej)
      if (fxActionPressed) return;
      if (fxDirty) commitFx(null);
    });
  }
  let fxActionPressed = false;
  [fxOk, fxCancel].forEach((b) => b && b.addEventListener("pointerdown", () => { fxActionPressed = true; setTimeout(() => { fxActionPressed = false; }, 400); }));
  if (fxOk) fxOk.addEventListener("click", () => commitFx({}));
  if (fxCancel) fxCancel.addEventListener("click", () => { revertFx(); returnFocusToGrid(); });

  // Pisanie w komórce widać też w pasku formuły (jak w Excelu).
  document.addEventListener("input", (e) => {
    const el = e.target;
    if (fxInput && el && el.classList && el.classList.contains("cell-editor")) fxInput.value = el.value;
  });
  // Zamknięty edytor (Esc / zatwierdzenie) → pasek wraca do stanu komórki.
  document.addEventListener("focusout", (e) => {
    if (e.target && e.target.classList && e.target.classList.contains("cell-editor")) setTimeout(() => { fxCellKey = ""; sync(); }, 0);
  });

  // ── pole nazwy: skok do adresu ─────────────────────────────────────────────
  function goToRef(text) {
    const m = String(text || "").trim().toUpperCase().replace(/\$/g, "").match(/^([A-Z]{1,3})(\d{1,7})(?::([A-Z]{1,3})(\d{1,7}))?$/);
    if (!m || !currentDisplayModel || currentDisplayModel.mode !== "wide") { toast(t("fmtBadRef"), "warning"); return false; }
    const a = XLSX.utils.decode_cell(m[1] + m[2]);
    const b = m[3] ? XLSX.utils.decode_cell(m[3] + m[4]) : null;
    const find = (cell) => {
      const col = cell.c - currentStartCol;
      if (col < 0 || col >= currentDisplayModel.headers.length) return { err: "fmtRefOutside" };
      const row = currentDisplayModel.rows.find((r) => r && r.rowIndex0 === cell.r && !r.isLongViewRow);
      if (!row) {
        const exists = baseRows.some((r) => r && r.rowIndex0 === cell.r);
        return { err: exists ? "fmtRefHidden" : (cell.r < currentHeaderRow ? "fmtRefHeader" : "fmtRefEmpty") };
      }
      return { key: getRowSelectionKey(row), col };
    };
    const A = find(a);
    if (A.err) { toast(t(A.err, { ref: m[1] + m[2] }), "info"); return false; }
    setFocusedCell(A.key, A.col, { scroll: true });
    if (b) {
      const B = find(b);
      if (!B.err && typeof setSelectedCell === "function") setSelectedCell(B.key, B.col, { scroll: false });
    }
    returnFocusToGrid();
    return true;
  }
  if (nameBox) {
    nameBox.addEventListener("focus", () => nameBox.select());
    nameBox.addEventListener("keydown", (e) => {
      e.stopPropagation();
      if (e.key === "Enter") { e.preventDefault(); goToRef(nameBox.value); }
      else if (e.key === "Escape") { e.preventDefault(); sync(); returnFocusToGrid(); }
    });
    nameBox.addEventListener("blur", () => setTimeout(sync, 0));
  }

  // ── skróty w tabeli ────────────────────────────────────────────────────────
  document.addEventListener("keydown", (e) => {
    if (!hasData() || !focusedCellState) return;
    if (typeof shouldIgnoreTableArrowNavigation === "function" && shouldIgnoreTableArrowNavigation()) return;
    const trOverlay = document.getElementById("transcribeOverlay");
    if (trOverlay && !trOverlay.classList.contains("hidden")) return;
    if (document.querySelector(".modal:not(.hidden)")) return;
    const meta = e.ctrlKey || e.metaKey;
    if (e.key === "Escape" && painterFmt) { e.preventDefault(); stopPainter(); return; }
    // Ctrl+D / Ctrl+R = wypełnij w dół / w prawo (kopia, jak w Excelu; autofill.js).
    // Na Macu też ⌘D (Excel tak ma); ⌘R zostaje przeglądarce (odświeżanie).
    if (!e.altKey && !e.shiftKey && (e.ctrlKey || (e.metaKey && e.code === "KeyD")) && (e.code === "KeyD" || e.code === "KeyR")) {
      e.preventDefault();
      if (typeof swbFillCopy === "function") swbFillCopy(e.code === "KeyD" ? "down" : "right");
      return;
    }
    if (meta && !e.altKey && !e.shiftKey) {
      const code = e.code || "";
      const map = { KeyB: "b", KeyI: "i", KeyU: "u", Digit5: "strike" };
      const key = map[code];
      if (key) {
        e.preventDefault();
        cfmtToggle(key);
        return;
      }
    }
    if (!meta && !e.altKey && (e.key === "Delete" || e.key === "Backspace")) {
      e.preventDefault();
      cfmtClearContents();
      return;
    }
    if (e.key === "F2" && !meta && !e.altKey) {
      const td = findCellElement(focusedCellState);
      if (td && typeof openCellEditor === "function") {
        e.preventDefault();
        openCellEditor(td);
        setTimeout(() => {
          const inp = td.querySelector("input.cell-editor");
          if (inp) { const end = inp.value.length; inp.setSelectionRange(end, end); }
        }, 0);
      }
    }
  });

  if (ribbonEl && typeof attachOverflowFade === "function") attachOverflowFade(ribbonEl);

  // Sekcje panelu: treść wsuwa się przy rozwinięciu (jak w Documents). Bez animowania
  // wysokości — to przeliczałoby układ panelu co klatkę. Start apki bez ruchu.
  let revealReady = false;
  setTimeout(() => { revealReady = true; }, 1200);
  document.querySelectorAll(".sidebar details.panel").forEach((d) => {
    d.addEventListener("toggle", () => {
      if (!d.open || !revealReady) return;
      d.classList.remove("is-revealing");
      void d.offsetWidth; // szybkie zamknij/otwórz — animacja od nowa
      d.classList.add("is-revealing");
    });
    d.addEventListener("animationend", (e) => {
      if (e.target.parentElement === d) d.classList.remove("is-revealing");
    });
  });
  setMode(swbEditMode);
  // Zawartość komórki zmienia się też „z boku” (Cofnij, wklejenie) — po każdym renderze tabeli.
  return {
    sync,
    setMode,
    stopPainter,
    goToRef,
    get mode() { return swbEditMode; },
  };
})();
