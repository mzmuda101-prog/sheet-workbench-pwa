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
      // formuły z odwołaniem do usuniętego arkusza mają w pliku „#REF!”, a w komórce starą wartość — przelicz
      if (typeof swbFormulaRecalcCells === "function") {
        const keys = [];
        workbook.SheetNames.forEach((sn) => { const ws = workbook.Sheets[sn]; Object.keys(ws).forEach((ref) => { if (ref[0] !== "!" && ws[ref] && ws[ref].f && ws[ref].f.includes("#REF!")) keys.push(`${sn}\u0000${ref}`); }); });
        swbFormulaRecalcCells(keys);
      }
      openSheet(target);
      // dymek po wczytaniu arkusza — inaczej „Wczytano …” przychodzi później i go zastępuje
      const t0 = performance.now();
      await new Promise((resolve) => {
        const tick = () => (document.body.getAttribute("aria-busy") !== "true" || performance.now() - t0 > 15000 ? resolve() : setTimeout(tick, 40));
        setTimeout(tick, 80);
      });
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
    const tableBtn = document.getElementById("fmtTableBtn");
    if (tableBtn) tableBtn.addEventListener("click", () => { formatAsTable(); });
    ["newWorkbookBtn", "emptyNewBtn"].forEach((id) => {
      const b = document.getElementById(id);
      if (b) b.addEventListener("click", () => { newWorkbook(); });
    });
  }

  // Formatuj jako tabelę: wiersz nagłówka + wszystkie wiersze danych; kolumny zaznaczone (gdy zaznaczono
  // kilka) albo wszystkie. Nazwy kolumn jak Excel trafiają do komórek nagłówka, potem tabela w pliku.
  async function formatAsTable() {
    if (busy || !workbook || !currentSheetName) return false;
    if (typeof swbEditGuard === "function" && !swbEditGuard()) return false;
    if (!currentDisplayModel || currentDisplayModel.mode !== "wide") { toast(t("tableNotWide"), "info"); return false; }
    const real = typeof swbRealHeaderCount === "function" ? swbRealHeaderCount() : currentHeaders.length;
    if (!real) return false;
    const rect = typeof getSelectionRectangle === "function" ? getSelectionRectangle() : null;
    const c0 = rect && rect.colCount > 1 ? Math.min(rect.colMin, real - 1) : 0;
    let c1 = rect && rect.colCount > 1 ? Math.min(rect.colMax, real - 1) : real - 1;
    // bez zaznaczenia: puste kolumny z prawej (pusta siatka nowego skoroszytu) nie wchodzą — jak bieżący obszar w Excelu
    const wsH = workbook.Sheets[currentSheetName];
    const colEmpty = (c) => { const h = wsH[XLSX.utils.encode_cell({ r: currentHeaderRow - 1, c: currentStartCol + c })]; return !(h && h.v != null && h.v !== "") && !baseRows.some((row) => row && Array.isArray(row.values) && row.values[c] != null && row.values[c] !== ""); };
    if (!(rect && rect.colCount > 1)) while (c1 > c0 && colEmpty(c1)) c1--;
    const r0 = currentHeaderRow - 1;
    let r1 = r0 + 1; // tabela bez danych: nagłówek + jeden pusty wiersz (jak Excel)
    baseRows.forEach((row) => {
      if (!row || !Number.isFinite(row.rowIndex0) || row.rowIndex0 <= r1) return;
      if (Array.isArray(row.values) && row.values.slice(c0, c1 + 1).some((v) => v != null && v !== "")) r1 = row.rowIndex0;
    });
    const range = { s: { r: r0, c: currentStartCol + c0 }, e: { r: r1, c: currentStartCol + c1 } };
    const cross = (a) => a.s.r <= range.e.r && a.e.r >= range.s.r && a.s.c <= range.e.c && a.e.c >= range.s.c;
    const hit = [...new Set(Object.values(currentTables || {}))].find((tb) => {
      if (!tb || tb.sheetName !== currentSheetName) return false;
      try { return cross(XLSX.utils.decode_range(tb.ref)); } catch { return false; }
    });
    if (hit) { toast(t("tableOverlap", { name: hit.displayName || "" }), "info"); return false; }
    if ((currentMerges || []).some((m) => m && m.s && m.e && cross(m))) { toast(t("tableMerged"), "info"); return false; }
    const ws = workbook.Sheets[currentSheetName];
    const raw = [];
    for (let c = range.s.c; c <= range.e.c; c++) {
      const cell = ws[XLSX.utils.encode_cell({ r: r0, c })];
      raw.push(cell && cell.v != null && cell.v !== "" ? String(cell.w != null ? cell.w : cell.v) : null);
    }
    const columns = swbSheetOps.tableColumnNames(raw, t("tableEmptyCol"));
    // Excel: nagłówek tabeli to zawsze tekst równy nazwie kolumny (formuła i liczba też stają się tekstem)
    columns.forEach((name, i) => {
      const cell = ws[XLSX.utils.encode_cell({ r: r0, c: range.s.c + i })];
      if (!cell || cell.f || cell.t !== "s" || cell.v !== name) updateSheetCell(r0, c0 + i, { value: name, type: "string" });
    });
    const taken = new Set(Object.keys(currentTables || {}).concat(((workbook.Workbook && workbook.Workbook.Names) || []).map((n) => String(n.Name || "").toLowerCase())));
    let k = 1;
    while (taken.has(`${t("tableBase")}${k}`.toLowerCase())) k++;
    const name = `${t("tableBase")}${k}`;
    const ref = XLSX.utils.encode_range(range);
    return apply((zip) => swbSheetOps.addTable(zip, { sheet: currentSheetName, ref, name, columns }).then(() => currentSheetName), () => t("tableCreated", { name, ref }));
  }

  return { newWorkbook, addSheet, renameSheet, deleteSheet, formatAsTable, menuActions, init, isBusy: () => busy };
})();

document.addEventListener("DOMContentLoaded", () => swbSheetOpsUi.init());
