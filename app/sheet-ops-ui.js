// sheet-ops-ui.js — „Nowy skoroszyt” i arkusze: dodaj / zmień nazwę / usuń (app/sheet-ops.js).
//
// Gdzie (docs/OD-ZERA.md, „Dom poleceń”):
//   • Nowy skoroszyt — panel „Plik i arkusz” (obok wyboru pliku); skrót na ekranie startowym.
//   • Arkusze — lista pod nazwą arkusza w nagłówku (▾): na dole „Nowy arkusz”, „Zmień nazwę…”,
//     „Usuń arkusz…” dla bieżącego arkusza. Telefon: ta sama lista od dołu ekranu.
//
// Jak: bieżące poprawki trafiają najpierw do pliku w pamięci (ten sam ZIP-patch co przy zapisie),
// potem operacja na pliku i ponowne wczytanie skoroszytu — bez okna „porzucić zmiany?” i bez
// wpisu w „Ostatnio otwierane”. Plik zostaje niezapisany (Zapisz jak zawsze). Jak w Excelu
// operacji na arkuszach nie da się cofnąć — historia Cofnij zaczyna się od nowa.

// Zapamiętane PRZED owinięciem przez drafts/recent-files/session-resume (DOMContentLoaded):
// ponowne wczytanie po operacji na arkuszu to ten sam plik, nie „otwarcie innego”.
const swbHandleFileCore = handleFile;

const swbSheetOpsUi = (() => {
  let bookCounter = 0;
  let busy = false;

  const SVG = {
    plus: '<path d="M12 5v14M5 12h14"/>',
    pencil: '<path d="M17 3a2.85 2.83 0 1 1 4 4L7.5 20.5 2 22l1.5-5.5Z"/>',
    trash: '<path d="M3 6h18M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6M8 6V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/>',
  };
  function icon(name) {
    const s = document.createElementNS("http://www.w3.org/2000/svg", "svg");
    s.setAttribute("viewBox", "0 0 24 24");
    s.setAttribute("width", "16");
    s.setAttribute("height", "16");
    s.setAttribute("fill", "none");
    s.setAttribute("stroke", "currentColor");
    s.setAttribute("stroke-width", "2");
    s.setAttribute("stroke-linecap", "round");
    s.setAttribute("stroke-linejoin", "round");
    s.setAttribute("aria-hidden", "true");
    s.innerHTML = SVG[name];
    return s;
  }

  // ── Nowy skoroszyt ───────────────────────────────────────────────────────────
  async function newWorkbook() {
    if (busy) return false;
    if (!(await ensureXlsxLibs(true))) return false;
    const zip = await swbSheetOps.newWorkbook(JSZip, { sheetName: `${t("sheetBase")}1`, headingSheets: t("sheetHeadingPair") });
    const bytes = await zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
    const file = new File([bytes], `${t("newWorkbookBase")}${++bookCounter}.xlsx`, { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    if (typeof swbRecent !== "undefined") swbRecent.skipNextOpen(); // pusty zeszyt nie trafia do „Ostatnio otwierane”
    // window.handleFile: pyta o niezapisane zmiany poprzedniego pliku (drafts.js)
    const opened = await window.handleFile(file);
    if (!opened || !workbook) { bookCounter--; return false; }
    if (typeof sheetEditUi !== "undefined") sheetEditUi.setMode("edit"); // od razu do pisania
    openSheet(workbook.SheetNames[0]);
    return true;
  }

  function openSheet(name) {
    if (autoHeaderRowEl) autoHeaderRowEl.checked = true;
    sheetSelect.value = name;
    loadBtn.click();
  }

  // ── operacje na arkuszach ────────────────────────────────────────────────────
  // Poprawki do pliku → operacja na ZIP → ponowne wczytanie, arkusz `target` na wierzchu.
  async function apply(op, okMessage) {
    if (busy || !workbook || !originalFileBytes) return false;
    if (typeof swbEditGuard === "function" && !swbEditGuard()) return false;
    busy = true;
    try {
      if (typeof flushCellEditor === "function") flushCellEditor();
      const grow = typeof swbRowGrow !== "undefined" ? swbRowGrow.growInfo() : null;
      const base = await buildPatchedXlsx(originalFileBytes, pendingEdits, buildSheetFormulaMaps(), { grow, layout: pendingLayout });
      const zip = await JSZip.loadAsync(base);
      const target = await op(zip);
      const bytes = await zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
      const name = currentFileName;
      const handle = currentFileHandle;
      const mode = typeof swbEditMode !== "undefined" ? swbEditMode : null;
      const ok = await swbHandleFileCore(new File([bytes], name), handle, { reopen: true });
      if (!ok) throw new Error("reopen");
      if (mode && typeof sheetEditUi !== "undefined") sheetEditUi.setMode(mode);
      if (typeof swbRowGrow !== "undefined") swbRowGrow.snapshot(); // dopisane wiersze są już w pliku
      undoStack = [];
      redoStack = [];
      if (typeof updateUndoButtons === "function") updateUndoButtons();
      swbStructureDirty = true;
      setDirtyState(true);
      openSheet(target);
      toast(okMessage(target), "success");
      return true;
    } catch (err) {
      const key = err && /^sheet[A-Z]/.test(err.message) ? err.message : "sheetOpFailed";
      toast(t(key), key === "sheetOpFailed" ? "error" : "warning");
      log(`Operacja na arkuszu nie powiodła się: ${(err && err.message) || err}`, "error");
      return false;
    } finally {
      busy = false;
    }
  }

  async function addSheet() {
    if (!workbook) return false;
    const name = swbSheetOps.nextName(workbook.SheetNames, t("sheetBase"));
    return apply((zip) => swbSheetOps.addSheet(zip, { name, afterName: currentSheetName }), (n) => t("sheetAdded", { name: n }));
  }

  async function renameSheet(from = currentSheetName) {
    if (!workbook || !from) return false;
    if (typeof swbEditGuard === "function" && !swbEditGuard()) return false;
    const names = workbook.SheetNames;
    const raw = await appPrompt(t("sheetRenamePrompt", { name: from }), from, {
      title: t("sheetRenameTitle"),
      okLabel: t("sheetRenameOk"),
      maxLength: 31,
      validate: (v) => { const k = swbSheetOps.nameProblem(v, names, from); return k ? t(k) : ""; },
    });
    if (raw == null || raw === from) return false;
    return apply((zip) => swbSheetOps.renameSheet(zip, { from, to: raw }), (n) => t("sheetRenamed", { name: n }));
  }

  async function deleteSheet(name = currentSheetName) {
    if (!workbook || !name) return false;
    if (typeof swbEditGuard === "function" && !swbEditGuard()) return false;
    const meta = (workbook.Workbook && workbook.Workbook.Sheets) || [];
    const visible = workbook.SheetNames.filter((n, i) => !meta[i] || !meta[i].Hidden);
    if (visible.length <= 1 && visible.includes(name)) { toast(t("sheetDeleteLast"), "info"); return false; }
    // jak Excel: pytamy, gdy w arkuszu coś jest (pusty znika od razu)
    const ws = workbook.Sheets[name];
    const hasData = ws && Object.keys(ws).some((k) => k[0] !== "!" && (typeof isMeaningfulSheetCell !== "function" || isMeaningfulSheetCell(ws[k])));
    if (hasData && !(await appConfirm(t("sheetDeleteMsg"), { title: t("sheetDeleteTitle", { name }), okLabel: t("sheetDeleteOk"), danger: true }))) return false;
    return apply((zip) => swbSheetOps.deleteSheet(zip, { name }), () => t("sheetDeleted", { name }));
  }

  // Pozycje na dole listy arkuszy (menu pod nazwą arkusza i lista od dołu na telefonie).
  function menuActions(close, itemClass) {
    const make = (id, label, ico, run, danger) => {
      const b = document.createElement("button");
      b.type = "button";
      b.className = `${itemClass} sheet-op-item`;
      b.dataset.sheetOp = id;
      if (danger) b.classList.add("is-danger");
      const text = document.createElement("span");
      text.className = "app-menu-item-text";
      text.textContent = label;
      b.append(icon(ico), text);
      b.addEventListener("click", () => { close(); run(); });
      return b;
    };
    const sep = document.createElement("div");
    sep.className = "app-menu-sep";
    sep.setAttribute("role", "separator");
    return [
      sep,
      make("add", t("sheetAdd"), "plus", addSheet),
      make("rename", t("sheetRenameItem"), "pencil", () => renameSheet()),
      make("delete", t("sheetDeleteItem"), "trash", () => deleteSheet(), true),
    ];
  }

  function init() {
    ["newWorkbookBtn", "emptyNewBtn"].forEach((id) => {
      const b = document.getElementById(id);
      if (b) b.addEventListener("click", () => { newWorkbook(); });
    });
  }

  return { newWorkbook, addSheet, renameSheet, deleteSheet, menuActions, init, isBusy: () => busy };
})();

document.addEventListener("DOMContentLoaded", () => swbSheetOpsUi.init());
