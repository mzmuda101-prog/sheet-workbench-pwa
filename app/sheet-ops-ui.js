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
    flag: '<path d="M4 22V4a1 1 0 0 1 1-1h12l-2 4 2 4H5"/>',
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
  // Plik z bieżącymi poprawkami (pendingEdits, dopisane wiersze, szerokości) — baza operacji na ZIP
  // i „zdjęcie” stanu do Cofnij.
  function currentBytes() {
    const grow = typeof swbRowGrow !== "undefined" ? swbRowGrow.growInfo() : null;
    return buildPatchedXlsx(originalFileBytes, pendingEdits, buildSheetFormulaMaps(), { grow, layout: pendingLayout });
  }
  // Wczytanie pliku od nowa (po operacji albo przy Cofnij/Ponów kroku „cały plik”) z zachowaniem
  // nazwy, uchwytu, trybu i stanu „niezapisane”; arkusz `sheet` na wierzchu.
  async function reopen(bytes, sheet, { model = false } = {}) {
    const name = currentFileName;
    const handle = currentFileHandle;
    const mode = typeof swbEditMode !== "undefined" ? swbEditMode : null;
    const ok = await swbHandleFileCore(new File([bytes], name), handle, { reopen: true });
    if (!ok) throw new Error("reopen");
    if (mode && typeof sheetEditUi !== "undefined") sheetEditUi.setMode(mode);
    if (typeof swbRowGrow !== "undefined") swbRowGrow.snapshot(); // dopisane wiersze są już w pliku
    swbStructureDirty = true;
    setDirtyState(true);
    // formuły z odwołaniem do usuniętego arkusza mają w pliku „#REF!”, a w komórce starą wartość — przelicz
    if (typeof swbFormulaRecalcCells === "function") {
      const keys = [];
      workbook.SheetNames.forEach((sn) => { const ws = workbook.Sheets[sn]; Object.keys(ws).forEach((ref) => { if (ref[0] !== "!" && ws[ref] && ws[ref].f && ws[ref].f.includes("#REF!")) keys.push(`${sn}\u0000${ref}`); }); });
      swbFormulaRecalcCells(keys);
    }
    openSheet(sheet);
    // dymek po wczytaniu arkusza — inaczej „Wczytano …” przychodzi później i go zastępuje. Czekamy też
    // na model tabeli: bez niego kolejna operacja (np. „Tylko zaznaczenie” zaraz po Cofnij) widziała
    // „nie tryb szeroki” i kończyła się po cichu (table-undo 3, Chromium).
    const t0 = performance.now();
    await new Promise((resolve) => {
      // model tylko na żądanie (Cofnij/Ponów tabeli) — pusty arkusz („Nowy arkusz”) modelu nie ma
      const ready = () => document.body.getAttribute("aria-busy") !== "true" && (!model || !!currentDisplayModel);
      const tick = () => (ready() || performance.now() - t0 > 15000 ? resolve() : setTimeout(tick, 40));
      setTimeout(tick, 80);
    });
  }

  // Poprawki do pliku → operacja na ZIP → ponowne wczytanie, arkusz `target` na wierzchu.
  // Historia Cofnij zaczyna się od nowa (kroki komórek odnoszą się do starego pliku). Z `opts.undo`
  // (nazwa) operacja sama staje się krokiem Cofnij „cały plik”: przed = `opts.before` albo baza.
  async function apply(op, okMessage, opts = {}) {
    if (busy || !workbook || !originalFileBytes) return false;
    if (typeof swbEditGuard === "function" && !swbEditGuard()) return false;
    busy = true;
    try {
      if (typeof flushCellEditor === "function") flushCellEditor();
      const base = await currentBytes();
      const zip = await JSZip.loadAsync(base);
      const target = await op(zip);
      const bytes = await zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
      await reopen(bytes, target);
      undoStack = [];
      redoStack = [];
      if (opts.undo) undoStack.push({ file: true, workbook, sheet: target, changes: [], label: opts.undo, bytes: opts.before || base, other: bytes });
      if (typeof updateUndoButtons === "function") updateUndoButtons();
      toast(okMessage(target), "success", opts.toast ? opts.toast(target) : null);
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

  // Cofnij / Ponów kroku „cały plik” (undo.js undoApplyStep): wczytuje zapisany plik, a historia
  // (same kroki „cały plik” — kroki komórek zniknęły przy operacji) zostaje. Zwraca krok odwrotny od
  // razu; wczytanie biegnie potem (undoRedo kończy się synchronicznie, kolejne Cofnij czeka na isBusy).
  function fileStep(step) {
    const inverse = { ...step, bytes: step.other, other: step.bytes };
    busy = true;
    inverse.done = Promise.resolve().then(async () => {
      const u = undoStack;
      const r = redoStack;
      let ok = true;
      try { await reopen(step.bytes, step.sheet, { model: true }); } catch (err) {
        ok = false;
        toast(t("sheetOpFailed"), "error");
        log(`Cofnij/Ponów operacji na pliku nie powiodło się: ${(err && err.message) || err}`, "error");
      } finally {
        undoStack = u;
        redoStack = r;
        // nowy obiekt skoroszytu po wczytaniu; treść arkusza jest ta sama co przy nagraniu kroków
        // (Ponów zmian komórek po tabeli zostaje — undoDropStale by je inaczej wyrzucił)
        for (const s of [...u, ...r]) s.workbook = workbook;
        busy = false;
        if (typeof updateUndoButtons === "function") updateUndoButtons();
      }
      if (!ok) return false;
      const undone = r.includes(inverse);
      toast(t(undone ? "fileStepUndone" : "fileStepRedone", { name: step.label }), "info", {
        action: t(undone ? "redoAction" : "undoAction"),
        onAction: () => (undone ? redoLast() : undoLast()),
      });
      return true;
    });
    return inverse;
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

  // „Ustaw jako arkusz startowy” (Mateusz 2026-11-20): Excel otworzy plik na tym arkuszu (po zapisaniu).
  // W apce zostajemy na bieżącym arkuszu.
  async function setStartSheet(name = currentSheetName) {
    if (!workbook || !name || !originalFileBytes) return false;
    try {
      const zip = await JSZip.loadAsync(await currentBytes());
      if (swbSheetOps.startSheetName(await zip.file("xl/workbook.xml")?.async("string")) === name) {
        toast(t("sheetStartAlready", { name }), "info");
        return false;
      }
    } catch { /* sprawdzenie pomocnicze — operacja i tak sprawdzi plik */ }
    const stay = currentSheetName;
    return apply(async (zip) => { await swbSheetOps.setStartSheet(zip, { name }); return stay; }, () => t("sheetStartSet", { name }));
  }

  // Pozycje na dole listy arkuszy (menu pod nazwą arkusza i lista od dołu na telefonie) oraz menu karty
  // arkusza (prawy klik / przytrzymanie) — wtedy `name` to arkusz z karty, nie bieżący.
  function menuActions(close, itemClass, name = null) {
    // bez `name` (lista arkuszy) — bieżący arkusz W CHWILI KLIKNIĘCIA; lista na telefonie buduje się raz przy wczytaniu
    const target = () => name || currentSheetName;
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
      make("start", t("sheetStartItem"), "flag", () => setStartSheet(target())),
      make("add", t("sheetAdd"), "plus", addSheet),
      make("rename", t("sheetRenameItem"), "pencil", () => renameSheet(target())),
      make("delete", t("sheetDeleteItem"), "trash", () => deleteSheet(target()), true),
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
  // exact = zakres arkusza { s, e } (przycisk „Tylko zaznaczenie” w komunikacie): tabela dokładnie
  // na nim, pierwszy wiersz to nagłówek — jak Ctrl+T w Excelu na zaznaczeniu.
  async function formatAsTable(exact = null) {
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
    let r0 = currentHeaderRow - 1;
    let r1 = r0 + 1; // tabela bez danych: nagłówek + jeden pusty wiersz (jak Excel)
    baseRows.forEach((row) => {
      if (!row || !Number.isFinite(row.rowIndex0) || row.rowIndex0 <= r1) return;
      if (Array.isArray(row.values) && row.values.slice(c0, c1 + 1).some((v) => v != null && v !== "")) r1 = row.rowIndex0;
    });
    let range = { s: { r: r0, c: currentStartCol + c0 }, e: { r: r1, c: currentStartCol + c1 } };
    const auto = range;
    if (exact) { range = exact; r0 = exact.s.r; }
    const hc0 = range.s.c - currentStartCol; // kolumna nagłówka względem widoku (updateSheetCell)
    // zaznaczenie inne niż zakres tabeli: w komunikacie „Tylko zaznaczenie” (wiersze ciągłe, ≥ 2)
    const sel = !exact ? selectionRange(rect) : null;
    const offer = sel && XLSX.utils.encode_range(sel) !== XLSX.utils.encode_range(auto) ? sel : null;
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
    // Cofnij wraca do stanu PRZED wpisaniem nazw kolumn do nagłówka (niżej), nie tylko sprzed tabeli;
    // otwarty edytor komórki zatwierdzamy najpierw, żeby jego wpis był w tym stanie
    if (typeof flushCellEditor === "function") flushCellEditor();
    const before = await currentBytes();
    // Excel: nagłówek tabeli to zawsze tekst równy nazwie kolumny (formuła i liczba też stają się tekstem)
    columns.forEach((name, i) => {
      const cell = ws[XLSX.utils.encode_cell({ r: r0, c: range.s.c + i })];
      if (!cell || cell.f || cell.t !== "s" || cell.v !== name) updateSheetCell(r0, hc0 + i, { value: name, type: "string" });
    });
    const taken = new Set(Object.keys(currentTables || {}).concat(((workbook.Workbook && workbook.Workbook.Names) || []).map((n) => String(n.Name || "").toLowerCase())));
    let k = 1;
    while (taken.has(`${t("tableBase")}${k}`.toLowerCase())) k++;
    const name = `${t("tableBase")}${k}`;
    const ref = XLSX.utils.encode_range(range);
    const col = (c) => XLSX.utils.encode_col(c);
    const cols = range.s.c === range.e.c ? col(range.s.c) : `${col(range.s.c)}–${col(range.e.c)}`;
    const msg = exact ? t("tableCreated", { name, ref }) : t("tableCreatedRange", { name, ref, row: r0 + 1, cols });
    return apply((zip) => swbSheetOps.addTable(zip, { sheet: currentSheetName, ref, name, columns }).then(() => currentSheetName), () => msg, {
      undo: name,
      before,
      toast: () => {
        const step = undoStack[undoStack.length - 1]; // krok tej tabeli (apply go właśnie dodał)
        // przyciski dymka działają tylko na TĘ tabelę — po innej zmianie nie cofają cudzego kroku
        const undoThis = () => {
          if (typeof undoCloseStep === "function") undoCloseStep();
          if (undoStack[undoStack.length - 1] !== step || !undoLast()) { toast(t("undoStepGone"), "info"); return null; }
          return redoStack[redoStack.length - 1];
        };
        return {
          stack: !!offer,
          action: t("undoAction"),
          onAction: () => { undoThis(); },
          ...(offer ? {
            action2: t("tableOnlySel", { ref: XLSX.utils.encode_range(offer) }),
            // cofnij tę tabelę, potem utwórz ją dokładnie na zaznaczeniu
            onAction2: async () => {
              try {
                const inv = undoThis();
                if (!inv || !(await inv.done)) return;
                await formatAsTable(offer);
              } catch (err) {
                toast(t("sheetOpFailed"), "error");
                log(`Tylko zaznaczenie: ${(err && err.message) || err}`, "error");
              }
            },
          } : {}),
        };
      },
    });
  }

  // Zaznaczenie jako zakres arkusza — tylko gdy to ciągłe wiersze danych (bez wierszy nad nagłówkiem,
  // w kolejności arkusza: po sortowaniu albo filtrze widoczne wiersze nie muszą leżeć obok siebie).
  function selectionRange(rect) {
    if (!rect || rect.rowCount < 2) return null;
    const idx = [];
    for (let i = rect.rowStart; i <= rect.rowEnd; i++) {
      const row = rect.model.rows[i];
      if (!row || row.isPreRow || !Number.isFinite(row.rowIndex0)) return null;
      idx.push(row.rowIndex0);
    }
    if (idx.some((r, i) => i && r !== idx[i - 1] + 1)) return null;
    const real = typeof swbRealHeaderCount === "function" ? swbRealHeaderCount() : currentHeaders.length;
    if (rect.colMin > real - 1) return null;
    return { s: { r: idx[0], c: currentStartCol + rect.colMin }, e: { r: idx[idx.length - 1], c: currentStartCol + Math.min(rect.colMax, real - 1) } };
  }

  return { newWorkbook, addSheet, renameSheet, deleteSheet, setStartSheet, formatAsTable, fileStep, menuActions, init, isBusy: () => busy };
})();

document.addEventListener("DOMContentLoaded", () => swbSheetOpsUi.init());
