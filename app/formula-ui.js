// formula-ui.js — pisanie formuł jak w Excelu.
//
// - Σ Autosumowanie (Alt+=): na pustej komórce proponuje =SUMA(blok liczb nad nią)
//   w edytorze — zatwierdzasz Enterem; na zaznaczonej kolumnie liczb wstawia sumę pod nią.
// - Wskazywanie: gdy piszesz formułę (w komórce albo w pasku formuły) i kursor stoi po
//   „=”, „(”, „;”, operatorze — klik/tap w komórkę wstawia jej adres, kolejny klik go
//   podmienia, Shift+klik robi zakres (B2:B9). Nie trzeba znać adresów na pamięć.
// - Podpowiedzi funkcji: po wpisaniu kilku liter lista pasujących (SUMA, SUMA.JEŻELI…)
//   ze składnią; ↑↓ + Tab/Enter wstawia „NAZWA(”. Wewnątrz nawiasu — podpowiedź argumentów.
// - fx = okno „Wstaw funkcję”: kategorie + szukajka, klik wstawia do paska formuły.

let swbFormulaPointingUntil = 0;

(() => {
  if (typeof FormulaEngine === "undefined") return;
  const FE = FormulaEngine;
  const fxInput = document.getElementById("formulaInput");
  const sumBtn = document.getElementById("fmtSumBtn");
  const sumMore = document.getElementById("fmtSumMore");
  const fxBtn = document.getElementById("fxInsertBtn");
  const lang = () => (typeof currentLang === "string" ? currentLang : "pl");
  const fnLabel = (en) => (lang() === "pl" ? FE.PL_NAMES[en] || en : en);
  const ARG_EN = {
    liczba: "number", zakres: "range", kryterium: "criteria", wartość: "value", tekst: "text", tablica: "array",
    szukana: "lookup_value", tabela: "table", nr_kolumny: "col_index", nr_wiersza: "row_index", przybliżone: "approximate",
    warunek: "condition", jeśli_prawda: "value_if_true", jeśli_fałsz: "value_if_false", miejsca: "digits", dzielnik: "divisor",
    potęga: "power", data: "date", miesiące: "months", rok: "year", miesiąc: "month", dzień: "day", typ: "type",
    od: "start", do: "end", jednostka: "unit", święta: "holidays", start: "start", dni: "days", ile: "count", nowy: "new",
    stary: "old", które: "instance", razy: "times", separator: "delimiter", pomiń_puste: "ignore_empty", format: "format",
    szukany: "find", k: "k", logiczna: "logical", wyrażenie: "expression", wynik: "result", domyślnie: "default",
    nr: "index", kolejność: "order", istotność: "significance", wielokrotność: "multiple", dół: "bottom", góra: "top",
    suma_zakres: "sum_range", średnia_zakres: "average_range", maks_zakres: "max_range", min_zakres: "min_range",
    funkcja: "function", odwołanie: "reference", tablica_szukana: "lookup_array", tablica_wynikowa: "return_array",
    jeśli_brak: "if_not_found", tryb: "match_mode", kierunek: "search_mode", wektor: "vector", czas: "time",
    godz: "hour", min: "minute", sek: "second", wartość_jeśli_błąd: "value_if_error", wartość_jeśli_brak: "value_if_na",
  };
  function syntaxOf(en) {
    const row = FE.CATALOG.find((x) => x[0] === en);
    if (!row) return "";
    let args = row[2];
    if (lang() !== "pl") {
      args = args.replace(/;/g, ",").replace(/[\p{L}_]+/gu, (w) => {
        const m = w.match(/^([\p{L}_]+?)(\d*)$/u);
        return m && ARG_EN[m[1]] ? ARG_EN[m[1]] + m[2] : w;
      });
    }
    return `${fnLabel(en)}(${args})`;
  }

  // ── aktywne pole formuły ───────────────────────────────────────────────────
  function activeFormulaInput() {
    const a = document.activeElement;
    if (!a || a.tagName !== "INPUT") return null;
    if (!(a.classList.contains("cell-editor") || a === fxInput)) return null;
    return String(a.value).trimStart().startsWith("=") ? a : null;
  }
  const insertable = (before) => /(^\s*=|[=(,;+\-*/^&<>:%])\s*$/.test(before);
  function setInput(input, value, caret) {
    input.value = value;
    input.setSelectionRange(caret, caret);
    input.dispatchEvent(new Event("input", { bubbles: true }));
  }

  // ── wskazywanie komórek ────────────────────────────────────────────────────
  let point = null; // { input, start, end, anchor:{r,c}, sheet }
  function refFromTd(td) {
    const tr = td.parentElement;
    const key = tr && tr.dataset.rowKey;
    const col = parseInt(td.dataset.colIndex || "", 10);
    if (!key || !key.startsWith("wide:") || !Number.isFinite(col)) return null;
    const r = Number(key.slice(5));
    return { r, c: currentStartCol + col };
  }
  function onPointDown(e) {
    const input = activeFormulaInput();
    if (!input) return;
    const td = e.target.closest && e.target.closest("#dataTable tbody td");
    if (!td || td.classList.contains("row-head") || td.contains(input)) return;
    const cell = refFromTd(td);
    if (!cell) return;
    const pos = input.selectionStart ?? input.value.length;
    const continuing = point && point.input === input && pos === point.end;
    if (!continuing && !insertable(input.value.slice(0, pos))) return;
    e.preventDefault();
    e.stopPropagation();
    swbFormulaPointingUntil = Date.now() + 600;
    swallowClick = true;
    let text = XLSX.utils.encode_cell(cell);
    let start = pos;
    if (continuing) {
      start = point.start;
      if (e.shiftKey && point.anchor) {
        const a = point.anchor;
        const r1 = Math.min(a.r, cell.r), r2 = Math.max(a.r, cell.r), c1 = Math.min(a.c, cell.c), c2 = Math.max(a.c, cell.c);
        text = r1 === r2 && c1 === c2 ? text : `${XLSX.utils.encode_cell({ r: r1, c: c1 })}:${XLSX.utils.encode_cell({ r: r2, c: c2 })}`;
      }
    }
    const v = input.value;
    const endOld = continuing ? point.end : pos;
    const next = v.slice(0, start) + text + v.slice(endOld);
    setInput(input, next, start + text.length);
    point = { input, start, end: start + text.length, anchor: continuing && e.shiftKey ? point.anchor : cell };
    td.classList.add("fx-pointed");
    setTimeout(() => td.classList.remove("fx-pointed"), 500);
    setTimeout(() => input.focus(), 0);
  }
  let swallowClick = false;
  document.addEventListener("pointerdown", onPointDown, true);
  document.addEventListener("mousedown", (e) => { if (swbFormulaPointingUntil > Date.now() && e.target.closest && e.target.closest("#dataTable tbody td")) { e.preventDefault(); e.stopPropagation(); } }, true);
  document.addEventListener("click", (e) => {
    if (!swallowClick) return;
    swallowClick = false;
    if (e.target.closest && e.target.closest("#dataTable tbody")) { e.preventDefault(); e.stopPropagation(); }
  }, true);
  // Każde pisanie (poza naszym wstawieniem) kończy „wskazywanie” — następny klik wstawia nowy adres.
  document.addEventListener("input", (e) => {
    if (point && e.target === point.input && e.isTrusted) point = null;
  }, true);

  // ── podpowiedzi funkcji ────────────────────────────────────────────────────
  const box = document.createElement("div");
  box.className = "fx-suggest";
  box.hidden = true;
  box.setAttribute("role", "listbox");
  document.body.appendChild(box);
  let boxInput = null;
  let items = [];
  let sel = 0;
  const allNames = () => FE.CATALOG.map((x) => x[0]);

  function currentWord(input) {
    const pos = input.selectionStart ?? input.value.length;
    const before = input.value.slice(0, pos);
    const m = before.match(/(^|[=(,;+\-*/^&<> ])([A-Za-zÀ-ɏ_][A-Za-z0-9À-ɏ_.]*)$/);
    if (!m) return null;
    // nie podpowiadaj w środku tekstu w cudzysłowie
    if (((before.match(/"/g) || []).length % 2) === 1) return null;
    return { word: m[2], start: pos - m[2].length, end: pos };
  }
  // najgłębiej otwarta funkcja przed kursorem → podpowiedź składni
  function openCall(input) {
    const before = input.value.slice(0, input.selectionStart ?? input.value.length);
    let depth = 0;
    let inStr = false;
    for (let i = before.length - 1; i >= 0; i--) {
      const ch = before[i];
      if (ch === '"') inStr = !inStr;
      if (inStr) continue;
      if (ch === ")") depth++;
      else if (ch === "(") {
        if (depth === 0) {
          const m = before.slice(0, i).match(/([A-Za-zÀ-ɏ_][A-Za-z0-9À-ɏ_.]*)$/);
          if (!m) return null;
          let en = m[1].toUpperCase();
          if (FE.EN_FROM_PL[en]) en = FE.EN_FROM_PL[en];
          return FE.CATALOG.some((x) => x[0] === en) ? en : null;
        }
        depth--;
      }
    }
    return null;
  }
  function place(input) {
    const r = input.getBoundingClientRect();
    const vv = window.visualViewport;
    const top = r.bottom + 4;
    box.style.left = `${Math.max(8, Math.min(r.left, (vv ? vv.width : window.innerWidth) - 300))}px`;
    box.style.top = `${top}px`;
    box.style.maxHeight = `${Math.max(120, (vv ? vv.height : window.innerHeight) - top - 12)}px`;
  }
  function hide() { box.hidden = true; boxInput = null; items = []; }
  function render(input) {
    const w = currentWord(input);
    const call = openCall(input);
    items = [];
    if (w && w.word.length >= 1) {
      const q = w.word.toUpperCase();
      const qa = q.normalize("NFD").replace(/[̀-ͯ]/g, "").replace(/Ł/g, "L");
      const ascii = (s) => s.normalize("NFD").replace(/[̀-ͯ]/g, "").replace(/Ł/g, "L");
      items = allNames().filter((en) => {
        const pl = FE.PL_NAMES[en] || en;
        return en.startsWith(q) || pl.startsWith(q) || ascii(pl).startsWith(qa);
      }).sort((a, b) => fnLabel(a).length - fnLabel(b).length).slice(0, 8);
      if (items.length === 1 && fnLabel(items[0]) === q) items = [];
    }
    if (!items.length && !call) { hide(); return; }
    sel = Math.min(sel, Math.max(0, items.length - 1));
    box.replaceChildren();
    if (items.length) {
      items.forEach((en, i) => {
        const b = document.createElement("button");
        b.type = "button";
        b.className = "fx-suggest-item";
        b.setAttribute("role", "option");
        b.setAttribute("aria-selected", i === sel ? "true" : "false");
        const name = document.createElement("span");
        name.className = "fx-suggest-name";
        name.textContent = fnLabel(en);
        const syn = document.createElement("span");
        syn.className = "fx-suggest-syntax";
        syn.textContent = syntaxOf(en);
        b.append(name, syn);
        b.addEventListener("pointerdown", (e) => { e.preventDefault(); accept(i); });
        box.appendChild(b);
      });
    } else if (call) {
      const hint = document.createElement("div");
      hint.className = "fx-suggest-hint";
      hint.textContent = syntaxOf(call);
      box.appendChild(hint);
    }
    boxInput = input;
    box.hidden = false;
    place(input);
  }
  function accept(i) {
    const input = boxInput;
    if (!input || !items[i]) return;
    const w = currentWord(input);
    if (!w) return;
    const name = fnLabel(items[i]) + "(";
    const v = input.value;
    setInput(input, v.slice(0, w.start) + name + v.slice(w.end), w.start + name.length);
    sel = 0;
    render(input);
    input.focus();
  }
  document.addEventListener("input", (e) => {
    const input = e.target;
    if (!(input && input.tagName === "INPUT" && (input.classList.contains("cell-editor") || input === fxInput))) return;
    if (!String(input.value).trimStart().startsWith("=")) { hide(); return; }
    render(input);
  });
  document.addEventListener("keyup", (e) => {
    const input = activeFormulaInput();
    if (input && (e.key === "ArrowLeft" || e.key === "ArrowRight" || e.key === "Home" || e.key === "End")) render(input);
  });
  // Klawisze listy idą PRZED obsługą edytora (Enter/Tab zatwierdzałyby komórkę).
  document.addEventListener("keydown", (e) => {
    if (box.hidden || !boxInput || e.target !== boxInput || !items.length) {
      if (!box.hidden && e.key === "Escape" && e.target === boxInput) { hide(); e.stopPropagation(); e.preventDefault(); }
      return;
    }
    if (e.key === "ArrowDown" || e.key === "ArrowUp") {
      sel = (sel + (e.key === "ArrowDown" ? 1 : -1) + items.length) % items.length;
      render(boxInput);
    } else if (e.key === "Tab" || (e.key === "Enter" && !e.shiftKey)) {
      accept(sel);
    } else if (e.key === "Escape") {
      hide();
    } else return;
    e.preventDefault();
    e.stopPropagation();
  }, true);
  document.addEventListener("focusout", (e) => {
    if (e.target === boxInput) setTimeout(() => { if (document.activeElement !== boxInput) hide(); }, 0);
  });
  window.addEventListener("resize", () => { if (!box.hidden && boxInput) place(boxInput); }, { passive: true });

  // ── Autosumowanie ──────────────────────────────────────────────────────────
  const SUM_FNS = ["SUM", "AVERAGE", "COUNT", "MAX", "MIN"];
  function autoSum(fn) {
    if (typeof swbEditGuard === "function" && !swbEditGuard()) return;
    if (!currentDisplayModel || currentDisplayModel.mode !== "wide" || !focusedCellState) { toast(t("autoSumNeedEmpty"), "info"); return; }
    const rect = typeof getSelectionRectangle === "function" ? getSelectionRectangle() : null;
    const multi = rect && (rect.rowCount > 1 || rect.colCount > 1);
    if (!multi) {
      const td = findCellElement(focusedCellState);
      const row = getRowByKey(focusedCellState.rowKey);
      if (!td || !row) return;
      const text = swbAutoSumFormula(fn, row.rowIndex0, currentStartCol + focusedCellState.colIndex0);
      openCellEditor(td);
      const input = td.querySelector("input.cell-editor");
      if (!input) return;
      // kursor między nawiasami, gdy nic nie znaleziono — od razu można wskazać zakres
      const caret = text.endsWith("()") ? text.length - 1 : text.length;
      setInput(input, text, caret);
      input.focus();
      return;
    }
    // zakres: suma pod każdą kolumną zaznaczenia (wiersze muszą być ciągłe w arkuszu)
    const rows = [];
    for (let i = rect.rowStart; i <= rect.rowEnd; i++) { const r = rect.model.rows[i]; if (r && Number.isFinite(r.rowIndex0)) rows.push(r.rowIndex0); }
    rows.sort((a, b) => a - b);
    if (!rows.length || rows[rows.length - 1] - rows[0] + 1 !== rows.length) { toast(t("autoSumNeedEmpty"), "info"); return; }
    const ws = workbook.Sheets[currentSheetName];
    const below = rows[rows.length - 1] + 1;
    let added = 0;
    const name = fnLabel(fn);
    for (let col = rect.colMin; col <= rect.colMax; col++) {
      if (typeof cfmtIsRealCol === "function" && !cfmtIsRealCol(col)) continue;
      const absC = currentStartCol + col;
      const target = ws[XLSX.utils.encode_cell({ r: below, c: absC })];
      if (target && target.v !== undefined && target.v !== null && target.v !== "") continue; // nie nadpisujemy danych
      const a = XLSX.utils.encode_cell({ r: rows[0], c: absC });
      const b = XLSX.utils.encode_cell({ r: rows[rows.length - 1], c: absC });
      if (swbFormulaCommitAt(below, col, `=${name}(${a}:${b})`, { silent: true }) === "ok") added++;
    }
    if (!added) { toast(t("autoSumNeedEmpty"), "info"); return; }
    setDirtyState(true);
    if (typeof rebuildCurrentSheetData === "function") rebuildCurrentSheetData(); // nowy wiersz pod danymi
    const key = `wide:${below}`;
    if (currentDisplayModel.rows.some((r) => getRowSelectionKey(r) === key)) setFocusedCell(key, rect.colMin, { scroll: true });
    toast(t("autoSumDone", { count: added }), "success");
  }
  if (sumBtn) sumBtn.addEventListener("click", () => autoSum("SUM"));
  // lista funkcji pod ▾ — ten sam wygląd co menu obramowań
  const menu = document.createElement("div");
  menu.className = "app-menu fmt-pop fx-sum-menu";
  menu.hidden = true;
  menu.setAttribute("role", "menu");
  document.body.appendChild(menu);
  function closeSumMenu() { menu.hidden = true; if (sumMore) sumMore.setAttribute("aria-expanded", "false"); }
  if (sumMore) sumMore.addEventListener("click", () => {
    if (!menu.hidden) { closeSumMenu(); return; }
    menu.setAttribute("aria-label", t("autoSumPick"));
    menu.replaceChildren(...SUM_FNS.map((fn) => {
      const b = document.createElement("button");
      b.type = "button";
      b.className = "app-menu-item";
      b.setAttribute("role", "menuitem");
      b.textContent = fnLabel(fn);
      b.addEventListener("click", () => { closeSumMenu(); autoSum(fn); });
      return b;
    }));
    menu.hidden = false;
    sumMore.setAttribute("aria-expanded", "true");
    const r = sumMore.getBoundingClientRect();
    menu.style.left = `${Math.max(8, Math.min(r.left - 40, window.innerWidth - 200))}px`;
    menu.style.top = `${r.bottom + 6}px`;
  });
  document.addEventListener("pointerdown", (e) => {
    if (!menu.hidden && !menu.contains(e.target) && e.target !== sumMore) closeSumMenu();
  }, true);
  // Alt+= jak w Excelu (e.code — na polskiej klawiaturze Alt bywa AltGr przy literach, „=” jest bezpieczne)
  document.addEventListener("keydown", (e) => {
    if (!e.altKey || e.ctrlKey || e.metaKey || e.code !== "Equal") return;
    if (!focusedCellState || (typeof shouldIgnoreTableArrowNavigation === "function" && shouldIgnoreTableArrowNavigation())) return;
    e.preventDefault();
    autoSum("SUM");
  });

  // ── fx: Wstaw funkcję ──────────────────────────────────────────────────────
  const dlg = document.createElement("div");
  dlg.className = "app-menu fx-dialog";
  dlg.hidden = true;
  dlg.setAttribute("role", "dialog");
  document.body.appendChild(dlg);
  const CATS = ["all", "math", "stat", "logic", "lookup", "date", "text", "info"];
  let cat = "all";
  function closeDlg(focusBack) {
    if (dlg.hidden) return;
    dlg.hidden = true;
    if (focusBack && fxBtn) fxBtn.focus();
  }
  function insertFn(en) {
    closeDlg(false);
    if (!fxInput || fxInput.disabled) return;
    if (typeof swbEditGuard === "function" && !swbEditGuard()) return;
    const name = fnLabel(en) + "(";
    let v = fxInput.value;
    fxInput.focus();
    if (!v.trimStart().startsWith("=")) {
      setInput(fxInput, "=" + name + ")", name.length + 1);
    } else {
      const pos = fxInput.selectionStart ?? v.length;
      setInput(fxInput, v.slice(0, pos) + name + ")" + v.slice(pos), pos + name.length);
    }
    render(fxInput);
  }
  function buildDlg() {
    dlg.setAttribute("aria-label", t("insertFunction"));
    const head = document.createElement("div");
    head.className = "fx-dialog-head";
    const title = document.createElement("div");
    title.className = "fx-dialog-title";
    title.textContent = t("insertFunction");
    const close = document.createElement("button");
    close.type = "button";
    close.className = "fx-act";
    close.textContent = "✕";
    close.setAttribute("aria-label", t("fmtFormulaCancel"));
    close.addEventListener("click", () => closeDlg(true));
    head.append(title, close);
    const search = document.createElement("input");
    search.type = "search";
    search.className = "fx-dialog-search";
    search.placeholder = t("insertFunctionSearch");
    search.setAttribute("aria-label", t("insertFunctionSearch"));
    const tabs = document.createElement("div");
    tabs.className = "fx-dialog-cats";
    const list = document.createElement("div");
    list.className = "fx-dialog-list";
    const hint = document.createElement("p");
    hint.className = "fx-dialog-hint";
    hint.textContent = t("fnInsertHint");
    const fill = () => {
      const q = search.value.trim().toUpperCase();
      const ascii = (s) => s.normalize("NFD").replace(/[̀-ͯ]/g, "").replace(/Ł/g, "L");
      list.replaceChildren(...FE.CATALOG.filter(([en, c]) => (cat === "all" || c === cat)
        && (!q || en.includes(q) || (FE.PL_NAMES[en] || "").includes(q) || ascii(FE.PL_NAMES[en] || "").includes(ascii(q))))
        .map(([en]) => {
          const b = document.createElement("button");
          b.type = "button";
          b.className = "fx-dialog-item";
          const n = document.createElement("span");
          n.className = "fx-suggest-name";
          n.textContent = fnLabel(en);
          const syn = document.createElement("span");
          syn.className = "fx-suggest-syntax";
          syn.textContent = syntaxOf(en);
          b.append(n, syn);
          b.addEventListener("click", () => insertFn(en));
          return b;
        }));
    };
    tabs.replaceChildren(...CATS.map((c) => {
      const b = document.createElement("button");
      b.type = "button";
      b.className = "fx-cat";
      b.textContent = t(`fnCat_${c}`);
      b.setAttribute("aria-pressed", c === cat ? "true" : "false");
      b.addEventListener("click", () => {
        cat = c;
        tabs.querySelectorAll(".fx-cat").forEach((x) => x.setAttribute("aria-pressed", x === b ? "true" : "false"));
        fill();
      });
      return b;
    }));
    search.addEventListener("input", fill);
    search.addEventListener("keydown", (e) => {
      if (e.key === "Enter") { const first = list.querySelector(".fx-dialog-item"); if (first) first.click(); e.preventDefault(); }
    });
    dlg.replaceChildren(head, search, tabs, list, hint);
    fill();
    if (typeof attachOverflowFade === "function") attachOverflowFade(tabs);
    return search;
  }
  if (fxBtn) fxBtn.addEventListener("click", () => {
    if (!dlg.hidden) { closeDlg(true); return; }
    const search = buildDlg();
    dlg.hidden = false;
    const r = fxBtn.getBoundingClientRect();
    const w = Math.min(420, window.innerWidth - 16);
    dlg.style.width = `${w}px`;
    dlg.style.left = `${Math.max(8, Math.min(r.left, window.innerWidth - w - 8))}px`;
    dlg.style.top = `${r.bottom + 6}px`;
    dlg.style.maxHeight = `${Math.max(240, window.innerHeight - r.bottom - 20)}px`;
    if (!matchMedia("(pointer: coarse)").matches) search.focus();
  });
  dlg.addEventListener("keydown", (e) => { if (e.key === "Escape") { e.preventDefault(); e.stopPropagation(); closeDlg(true); } });
  document.addEventListener("pointerdown", (e) => {
    if (!dlg.hidden && !dlg.contains(e.target) && e.target !== fxBtn) closeDlg(false);
  }, true);
})();
