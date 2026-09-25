// edit-tools-audit-playwright.js — AUDYT wszystkich „Narzędzi edycji" na PRAWDZIWYCH plikach.
//
// Dla każdego pliku (scripts/*.xlsx) i każdej operacji, przez prawdziwe UI panelu:
//   • Podgląd niczego nie zmienia, a liczba w nim = liczbie zmian po „Zastosuj",
//   • wynik = NIEZALEŻNIE policzone oczekiwanie (referencja w teście, nie kod apki),
//   • Cofnij przywraca dokładnie stan sprzed (wartości + brak pendingEdits),
//   • zakres „Zaznaczenie" i „Tylko przefiltrowane" ruszają wyłącznie swoje komórki.
// Na koniec: kilka operacji + edycja komórek w edytorze (daty wpisane „12.05.24”
// i „12-05-24”) → ZAPIS (ZIP-patch) → ponowny odczyt: każda zmiana jest w pliku,
// każda NIEzmieniona komórka identyczna z oryginałem, daty mają format daty,
// XML poprawny.
//
// Uruchom z serwerem na APP_URL (domyślnie http://127.0.0.1:4175/), np. `npm run serve:test`.
// FILES=ścieżka1,ścieżka2 — własne pliki.

const { chromium } = require("playwright");
const fs = require("fs");
const path = require("path");

const APP_URL = process.env.APP_URL || "http://127.0.0.1:4175/";
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
const FILES = process.env.FILES
  ? process.env.FILES.split(",")
  : fs.readdirSync(__dirname).filter((f) => f.endsWith(".xlsx")).map((f) => path.join(__dirname, f));

async function auditFile(browser, file) {
  const res = { fails: [], notes: [] };
  try { return await auditFileInner(browser, file, res); }
  catch (e) { res.fails.push(`WYJĄTEK w kroku „${res.step()}”: ` + e.message.split("\n")[0]); return res; }
}
async function auditFileInner(browser, file, res) {
  const [filePath, sheetName] = file.split("::");
  file = filePath;
  const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1400, height: 950 } });
  await context.addInitScript(() => localStorage.setItem("introPlayed", "true"));
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push("pageerror: " + e.message));
  page.on("console", (m) => { if (m.type() === "error") errors.push("console: " + m.text()); });
  const fails = [];
  const notes = [];
  let step = "wczytanie";
  res.fails = fails; res.notes = notes; res.step = () => step;
  const ok = (name, cond, got) => { if (!cond) fails.push(`${name} — ${JSON.stringify(got).slice(0, 400)}`); };

  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(() => { document.getElementById("heroSplash")?.remove(); try { ensureXlsxLibs(false); } catch {} });
  await page.setInputFiles("#fileInput", file);
  await page.waitForFunction(() => document.getElementById("sheetSelect")?.options?.length > 0, null, { timeout: 30000 });
  await sleep(300);
  // Największy arkusz (np. w RODO to „2025-2026”, a nie „Dashboard”).
  await page.evaluate((wanted) => {
    const selEl = document.getElementById("sheetSelect");
    if (wanted && [...selEl.options].some((o) => o.value === wanted)) {
      selEl.value = wanted; selEl.dispatchEvent(new Event("change", { bubbles: true })); return;
    }
    let bestName = null, bestN = -1;
    [...selEl.options].forEach((o) => {
      const sh = workbook && workbook.Sheets[o.value];
      if (!sh || !sh["!ref"]) return;
      const r = XLSX.utils.decode_range(sh["!ref"]);
      const n = (r.e.r - r.s.r + 1) * (r.e.c - r.s.c + 1);
      if (n > bestN) { bestN = n; bestName = o.value; }
    });
    if (bestName) { selEl.value = bestName; selEl.dispatchEvent(new Event("change", { bubbles: true })); }
  }, sheetName || null);
  await sleep(300);
  await page.click("#loadBtn");
  await page.waitForFunction(() => Array.isArray(baseRows) && baseRows.length > 0 && currentDisplayModel, null, { timeout: 60000 });
  await sleep(800);
  await page.evaluate(() => {
    if (typeof setSidebarOpen === "function") setSidebarOpen(true);
    document.getElementById("panel-edit-tools").open = true;
  });

  // Profil kolumn: ile tekstów / dat / liczb.
  const prof = await page.evaluate(() => {
    const cols = currentHeaders.map((h, i) => {
      let s = 0, d = 0, n = 0, dateText = 0, numText = 0;
      baseRows.forEach((r) => {
        const v = r.values[i];
        if (typeof v === "string" && v.trim()) {
          s++;
          if (/^\d{1,2}[.\-/]\d{1,2}[.\-/]\d{2,4}$/.test(v.trim())) dateText++;
          if (/^-?\d+([.,]\d+)?$/.test(v.trim())) numText++;
        } else if (v instanceof Date) d++;
        else if (typeof v === "number") n++;
      });
      return { i, h, s, d, n, dateText, numText };
    });
    return { sheet: currentSheetName, rows: baseRows.length, cols };
  });
  const best = (k) => prof.cols.slice().sort((a, b) => b[k] - a[k])[0];
  const textCol = best("s"), dateCol = best("d"), numCol = best("n");
  notes.push(`arkusz „${prof.sheet}”, ${prof.rows} wierszy; tekst: [${textCol.i}] ${textCol.h} (${textCol.s}), data: [${dateCol.i}] ${dateCol.h} (${dateCol.d}), liczba: [${numCol.i}] ${numCol.h} (${numCol.n})`);
  const textDateCols = prof.cols.filter((c) => c.dateText > 0);
  const textNumCols = prof.cols.filter((c) => c.numText > 0);
  if (textDateCols.length) notes.push("daty zapisane jako TEKST: " + textDateCols.map((c) => `${c.h}(${c.dateText})`).join(", "));
  if (textNumCols.length) notes.push("liczby zapisane jako TEKST: " + textNumCols.map((c) => `${c.h}(${c.numText})`).join(", "));

  const snapshot = (col) => page.evaluate((col) => baseRows.map((r) => {
    const v = r.values[col];
    return v instanceof Date ? { d: v.getTime() } : v === undefined ? null : v;
  }), col);
  const pendingCount = () => page.evaluate(() => Object.keys(pendingEdits[currentSheetName] || {}).length);
  const sel = (id, v) => page.selectOption(`#${id}`, String(v));
  const setCheck = (id, on) => page.evaluate(([id, on]) => { const el = document.getElementById(id); if (el.checked !== on) el.click(); }, [id, on]);

  async function preview() {
    await page.click("#previewEditToolBtn");
    await sleep(60);
    return page.evaluate(() => {
      const s = document.querySelector("#editPreview .edit-preview-sum")?.textContent || "";
      const m = s.match(/(\d+)/);
      return m ? Number(m[1]) : 0;
    });
  }
  async function apply() { await page.click("#applyEditToolBtn"); await sleep(120); }
  async function undo() { await page.evaluate(() => undoLast()); await sleep(120); }

  // Jeden przypadek: ustaw panel → podgląd → zastosuj → porównaj z referencją → cofnij.
  // expectFn(rawBefore) → oczekiwana wartość (albo undefined = bez zmiany). Referencja działa
  // po stronie PRZEGLĄDARKI (musi znać Date / getDisplayValue), więc podajemy ją jako string.
  const ensurePanel = () => page.evaluate(() => {
    const st = { sidebarWasOpen: typeof isSidebarOpen === "function" ? isSidebarOpen() : null, panelWasOpen: document.getElementById("panel-edit-tools").open };
    if (typeof setSidebarOpen === "function" && !isSidebarOpen()) setSidebarOpen(true);
    document.getElementById("panel-edit-tools").open = true;
    return st;
  });
  const qsFilter = async (query) => {
    await page.evaluate((query) => {
      closeQuickSearchPopup(); openQuickSearchPopup();
      document.getElementById("quickSearchPopupOperators").checked = false;
      document.getElementById("quickSearchPopupAction").value = "filter";
      const input = document.getElementById("quickSearchPopupInput");
      input.value = query; input.focus();
    }, query);
    await page.keyboard.press("Enter");
    await sleep(600);
  };
  async function runCase(name, col, setup, expectSrc) {
    step = name;
    const st = await ensurePanel();
    if (st.sidebarWasOpen === false || !st.panelWasOpen) notes.push(`(przed „${name}” panel był zamknięty: ${JSON.stringify(st)})`);
    await sleep(250);
    await sel("editScope", "column");
    await setCheck("editFilteredOnly", false);
    await sel("editColumnSelect", col);
    await setup();
    const before = await snapshot(col);
    const pend0 = await pendingCount();
    const pv = await preview();
    const afterPv = await snapshot(col);
    ok(`${name}: podgląd nic nie zmienia`, JSON.stringify(before) === JSON.stringify(afterPv) && (await pendingCount()) === pend0, null);
    const expected = await page.evaluate(([col, src]) => {
      const fn = new Function("v", "row", "col", src);
      return baseRows.map((r) => {
        // podnagłówki i wiersze pochodne apka celowo omija (jak przy edycji komórki)
        const out = r.isSubheader || r.isLongViewRow ? undefined : fn(r.values[col], r, col);
        const v = out === undefined ? r.values[col] : out;
        return v instanceof Date ? { d: v.getTime() } : v === undefined ? null : v;
      });
    }, [col, expectSrc]);
    const expChanged = expected.filter((v, i) => JSON.stringify(v) !== JSON.stringify(before[i])).length;
    await apply();
    const after = await snapshot(col);
    const changed = after.filter((v, i) => JSON.stringify(v) !== JSON.stringify(before[i])).length;
    const diffs = [];
    after.forEach((v, i) => { if (JSON.stringify(v) !== JSON.stringify(expected[i])) diffs.push({ row: i, before: before[i], got: v, want: expected[i] }); });
    ok(`${name}: wynik = referencja`, diffs.length === 0, { diffs: diffs.slice(0, 4), n: diffs.length });
    ok(`${name}: podgląd (${pv}) = zmienione (${changed})`, pv === changed, { pv, changed, expChanged });
    const pend1 = await pendingCount();
    ok(`${name}: pendingEdits przybyło o ${changed}`, pend1 - pend0 === changed, { pend0, pend1, changed });
    if (changed) {
      await undo();
      const back = await snapshot(col);
      ok(`${name}: Cofnij przywraca`, JSON.stringify(back) === JSON.stringify(before) && (await pendingCount()) === pend0, { pend: await pendingCount(), pend0 });
    }
    notes.push(`${name}: ${changed} zmian`);
    return changed;
  }

  const tc = textCol.i;
  await sel("editOp", "case"); // żeby pola się pokazały
  // 1. wielkość liter
  await runCase("WIELKIE", tc, async () => { await sel("editOp", "case"); await sel("editCaseMode", "upper"); },
    `return typeof v === "string" ? v.toLocaleUpperCase("pl-PL") : undefined;`);
  await runCase("Jak Nazwa Własna", tc, async () => { await sel("editOp", "case"); await sel("editCaseMode", "title"); },
    `return typeof v === "string" ? v.replace(/\\p{L}[\\p{L}\\p{M}]*/gu, (w) => w[0].toLocaleUpperCase("pl-PL") + w.slice(1).toLocaleLowerCase("pl-PL")) : undefined;`);
  // 2. przytnij
  await runCase("Zwiń spacje", tc, async () => { await sel("editOp", "trim"); await sel("editTrimMode", "collapse"); },
    `return typeof v === "string" ? v.replace(/\\s+/gu, " ").trim() : undefined;`);
  // 3. prefiks/sufiks
  await runCase("Prefiks/sufiks", tc, async () => {
    await sel("editOp", "affix"); await page.fill("#editPrefix", "[P]"); await page.fill("#editSuffix", "#");
  }, `return typeof v === "string" ? "[P]" + v + "#" : undefined;`);
  // 4. wyrównaj
  await runCase("Wyrównaj do 12 zerami", tc, async () => {
    await sel("editOp", "pad"); await page.fill("#editPadLen", "12"); await page.fill("#editPadChar", "0"); await sel("editPadSide", "start");
  }, `return typeof v === "string" ? v.padStart(12, "0") : undefined;`);
  // 5. znajdź i zamień — najczęstsza litera w kolumnie, dosłownie / bez wielkości / cała komórka / regex
  const letter = await page.evaluate((col) => {
    const c = {};
    baseRows.forEach((r) => { const v = r.values[col]; if (typeof v === "string") for (const ch of v.toLowerCase()) if (/\p{L}/u.test(ch)) c[ch] = (c[ch] || 0) + 1; });
    return Object.entries(c).sort((a, b) => b[1] - a[1])[0]?.[0] || "a";
  }, tc);
  const commonVal = await page.evaluate((col) => {
    const c = {};
    baseRows.forEach((r) => { const v = r.values[col]; if (typeof v === "string" && v.trim()) c[v] = (c[v] || 0) + 1; });
    return Object.entries(c).sort((a, b) => b[1] - a[1])[0]?.[0] || "x";
  }, tc);
  const setReplace = (find, repl, { regex = false, ic = false, whole = false } = {}) => async () => {
    await sel("editOp", "replace");
    await page.fill("#editFind", find); await page.fill("#editReplace", repl);
    await setCheck("editRegex", regex); await setCheck("editIgnoreCase", ic); await setCheck("editWholeCell", whole);
  };
  const L = JSON.stringify(letter);
  await runCase(`Zamień „${letter}”→„_” (dosłownie)`, tc, setReplace(letter, "_"),
    `return typeof v === "string" ? v.split(${L}).join("_") : undefined;`);
  await runCase(`Zamień „${letter.toUpperCase()}”→„_” (bez wielkości)`, tc, setReplace(letter.toUpperCase(), "_", { ic: true }),
    `return typeof v === "string" ? v.replace(new RegExp(${JSON.stringify(letter.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"))}, "giu"), "_") : undefined;`);
  const CV = JSON.stringify(commonVal);
  await runCase(`Zamień całą komórkę „${commonVal.slice(0, 20)}” (bez wielkości)`, tc, setReplace(commonVal.toUpperCase(), "ZMIENIONE", { ic: true, whole: true }),
    `return typeof v === "string" && v.toLocaleLowerCase("pl-PL") === ${CV}.toLocaleLowerCase("pl-PL") ? "ZMIENIONE" : undefined;`);
  await runCase("Regex: cyfry → #", tc, setReplace("\\d", "#", { regex: true }),
    `return typeof v === "string" ? v.replace(/\\d/gu, "#") : undefined;`);
  // znajdź i zamień na DATACH (po tym, co widać) i LICZBACH (wartość surowa)
  if (dateCol.d > 0) {
    const shown = await page.evaluate((col) => { const r = baseRows.find((r) => r.values[col] instanceof Date); return r ? String(getDisplayValue(r, col)) : ""; }, dateCol.i);
    notes.push(`data wyświetla się jako „${shown}”`);
    // zamiana roku na inny — bierzemy 2 ostatnie cyfry roku widoczne w tekście
    const yy = (shown.match(/(\d{2})\s*$/) || [])[1];
    if (yy) {
      const to = String((Number(yy) + 1) % 100).padStart(2, "0");
      await runCase(`Daty: rok „${yy}”→„${to}” (znajdź i zamień)`, dateCol.i, setReplace(yy, to), `
        if (!(v instanceof Date)) return undefined;
        const s = String(getDisplayValue(row, col));
        const n = s.split(${JSON.stringify(yy)}).join(${JSON.stringify(to)});
        if (n === s) return undefined;
        const p = parseInputValue(n);
        return p ? p.value : n;`);
    }
  }
  if (numCol.n > 0) {
    await runCase("Liczby: „0”→„9”", numCol.i, setReplace("0", "9"), `
      if (typeof v === "string") { const n = v.split("0").join("9"); return n === v ? undefined : n; }
      if (typeof v !== "number") return undefined;
      const s = String(v); const n = s.split("0").join("9");
      if (n === s) return undefined;
      const p = parseInputValue(n); return p ? p.value : n;`);
  }
  // 6. wzorzec (fuzzy + „=#")
  await runCase("Wzorzec „=#” (zostaw cyfry)", tc, async () => {
    await sel("editOp", "pattern"); await sel("editPatternMode", "pattern"); await page.fill("#editPatternInput", "=#");
  }, `if (typeof v !== "string") return undefined; const out = compileGroupPattern("=#")(v); return typeof out === "string" ? out : undefined;`);
  // 7. konwersja typu
  await runCase("Konwersja: liczby → tekst", numCol.i, async () => { await sel("editOp", "convert"); await sel("editConvertTo", "text"); },
    `if (v == null || typeof v === "string") return undefined; const s = getDisplayValue(row, col); return s === "" ? undefined : s;`);
  for (const c of textNumCols.slice(0, 2)) {
    await runCase(`Konwersja: tekst → liczba [${c.h}]`, c.i, async () => { await sel("editOp", "convert"); await sel("editConvertTo", "number"); },
      `if (typeof v === "number") return undefined; const s = String(getDisplayValue(row, col)).replace(/\\s/g, "").replace(",", "."); return /^-?\\d+(\\.\\d+)?$/.test(s) ? Number(s) : undefined;`);
  }
  for (const c of textDateCols.slice(0, 2)) {
    await runCase(`Konwersja: tekst → data [${c.h}]`, c.i, async () => { await sel("editOp", "convert"); await sel("editConvertTo", "date"); },
      `if (v instanceof Date) return undefined; const s = getDisplayValue(row, col); if (!s) return undefined;
       const bare = typeof v === "number" ? v : (/^\\d+(\\.\\d+)?$/.test(String(s).trim()) ? Number(s) : null);
       if (bare !== null && !(bare >= 18264 && bare <= 73415)) return undefined;
       const d = parseDateFlexible(s); return d instanceof Date && !isNaN(d) ? d : undefined;`);
  }
  // 8. ujednolić warianty — po zastosowaniu ponowne skanowanie nie może nic znaleźć
  for (const c of prof.cols.filter((c) => c.s > 5).slice(0, 6)) {
    step = "Ujednolić " + c.h;
    await ensurePanel();
    await sel("editScope", "column"); await setCheck("editFilteredOnly", false);
    await sel("editOp", "unify"); await sel("editColumnSelect", c.i); await sel("editUnifyMode", "loose");
    await page.click("#editUnifyScanBtn"); await sleep(60);
    const groups = await page.evaluate(() => [...document.querySelectorAll("#editUnifyList .unify-group")].map((g) =>
      [...g.querySelectorAll(".unify-val")].map((v) => v.textContent).join(" | ")));
    if (!groups.length) continue;
    notes.push(`warianty w „${c.h}”: ${groups.length} grup, np. ${groups.slice(0, 3).join("  ;  ")}`);
    const before = await snapshot(c.i);
    const pv = await preview();
    await apply();
    const after = await snapshot(c.i);
    const changed = after.filter((v, i) => JSON.stringify(v) !== JSON.stringify(before[i])).length;
    ok(`Ujednolić [${c.h}]: podgląd (${pv}) = zmienione (${changed})`, pv === changed && changed > 0, { pv, changed });
    await page.click("#editUnifyScanBtn"); await sleep(60);
    const left = await page.evaluate(() => document.querySelectorAll("#editUnifyList .unify-group").length);
    ok(`Ujednolić [${c.h}]: po zastosowaniu brak wariantów`, left === 0, left);
    await undo();
    ok(`Ujednolić [${c.h}]: Cofnij`, JSON.stringify(await snapshot(c.i)) === JSON.stringify(before), null);
  }

  // 9. zakres „Zaznaczenie" — Shift+klik 4 wiersze × 1 kolumna tekstowa
  {
    step = "Zaznaczenie";
    const pos = (nth, col) => page.evaluate(([nth, col]) => {
      const trs = [...document.querySelectorAll("#dataTable tbody tr")].filter((tr) => tr.querySelector(`td[data-col-index="${col}"]`));
      const tr = trs[nth];
      const td = tr && tr.querySelector(`td[data-col-index="${col}"]`);
      if (!td) return null;
      td.scrollIntoView({ block: "nearest", inline: "nearest" });
      const r = td.getBoundingClientRect();
      return { x: r.left + r.width / 2, y: r.top + r.height / 2, key: tr.dataset.rowKey };
    }, [nth, col]);
    await page.evaluate(() => { if (typeof setSidebarOpen === "function") setSidebarOpen(false); });
    await sleep(300);
    const a = await pos(1, tc), b = await pos(4, tc);
    if (a && b) {
      await page.mouse.click(a.x, a.y); await sleep(80);
      await page.keyboard.down("Shift"); await page.mouse.click(b.x, b.y); await page.keyboard.up("Shift"); await sleep(120);
      await page.evaluate(() => { setSidebarOpen(true); document.getElementById("panel-edit-tools").open = true; });
      await sleep(300);
      const rect = await page.evaluate(() => { const r = getSelectionRectangle(); return r && { rows: r.rowCount, cols: r.colCount }; });
      ok("Zaznaczenie: Shift+klik = 4×1 (warunek testu)", rect && rect.rows === 4 && rect.cols === 1, rect);
      const before = await snapshot(tc);
      const keys = await page.evaluate(() => [...getSelectionRectangle().rowKeys]);
      await sel("editScope", "selection");
      await sel("editOp", "affix"); await page.fill("#editPrefix", "SEL:"); await page.fill("#editSuffix", "");
      const pv = await preview();
      await apply();
      const after = await snapshot(tc);
      const rowKeysAll = await page.evaluate(() => baseRows.map((r) => `wide:${r.rowIndex0}`));
      const rowsSub = await page.evaluate(() => baseRows.map((r) => !!(r.isSubheader || r.isLongViewRow)));
      const bad = [];
      after.forEach((v, i) => {
        const inSel = keys.includes(rowKeysAll[i]);
        const b0 = before[i];
        const sub = rowsSub[i];
      const want = inSel && !sub && typeof b0 === "string" ? "SEL:" + b0 : b0;
        if (JSON.stringify(v) !== JSON.stringify(want)) bad.push({ i, inSel, b0, v });
      });
      ok("Zaznaczenie: zmienione tylko zaznaczone", bad.length === 0, bad.slice(0, 3));
      ok("Zaznaczenie: podgląd = zmiany", pv === after.filter((v, i) => JSON.stringify(v) !== JSON.stringify(before[i])).length, pv);
      await undo();
    } else fails.push("Zaznaczenie: nie znalazłem komórek w siatce");
  }

  // 10. „Tylko przefiltrowane" — filtr szybkim szukaniem po wartości z kolumny
  {
    step = "Tylko przefiltrowane";
    await sel("editScope", "column");
    const term = commonVal.slice(0, 6);
    await qsFilter(term);
    await ensurePanel(); await sleep(300);
    const vis = await page.evaluate(() => ({ n: viewRows.length, all: baseRows.length, keys: viewRows.map((r) => r.rowIndex0) }));
    notes.push(`filtr „${term}”: widać ${vis.n} z ${vis.all}`);
    await sel("editOp", "affix"); await page.fill("#editPrefix", "F:"); await page.fill("#editSuffix", "");
    await sel("editColumnSelect", tc);
    await setCheck("editFilteredOnly", true);
    const before = await snapshot(tc);
    const idx = await page.evaluate(() => baseRows.map((r) => r.rowIndex0));
    const subF = await page.evaluate(() => baseRows.map((r) => !!(r.isSubheader || r.isLongViewRow)));
    const pv = await preview();
    await apply();
    const after = await snapshot(tc);
    const bad = [];
    after.forEach((v, i) => {
      const inF = vis.keys.includes(idx[i]);
      const want = inF && !subF[i] && typeof before[i] === "string" ? "F:" + before[i] : before[i];
      if (JSON.stringify(v) !== JSON.stringify(want)) bad.push({ i, inF, b: before[i], v });
    });
    ok("Przefiltrowane: zmienione tylko widoczne wiersze", bad.length === 0 && vis.n < vis.all, { bad: bad.slice(0, 3), vis: vis.n });
    ok("Przefiltrowane: podgląd = zmiany", pv === after.filter((v, i) => JSON.stringify(v) !== JSON.stringify(before[i])).length, pv);
    await undo();
    await setCheck("editFilteredOnly", false);
    await page.evaluate(() => { closeQuickSearchPopup(); document.getElementById("resetFiltersBtn").click(); });
    await sleep(900);
    const cleared = await page.evaluate(() => viewRows.length === baseRows.length);
    ok("Przefiltrowane: filtr wyczyszczony (warunek testu)", cleared, cleared);
    await ensurePanel(); await sleep(300);
  }

  step = "Zapis";
  // 11. ZAPIS: kilka operacji + edycja dat w edytorze komórki → zapis → porównanie całego pliku
  await ensurePanel(); await sleep(250);
  await sel("editScope", "column"); await setCheck("editFilteredOnly", false);
  await runCase("[do zapisu] WIELKIE", tc, async () => { await sel("editOp", "case"); await sel("editCaseMode", "upper"); },
    `return typeof v === "string" ? v.toLocaleUpperCase("pl-PL") : undefined;`).then(() => undefined);
  // runCase cofa — nałóż ponownie bez cofania
  await sel("editOp", "case"); await sel("editCaseMode", "upper"); await sel("editColumnSelect", tc); await apply();
  if (numCol.n > 0) { await sel("editColumnSelect", numCol.i); await sel("editOp", "convert"); await sel("editConvertTo", "text"); await apply(); }
  // tekst → data (jeśli są daty-tekst): zapis musi dać w Excelu DATĘ (format), nie „45424”
  const convertedDateCol = textDateCols[0];
  if (convertedDateCol) { await sel("editColumnSelect", convertedDateCol.i); await sel("editOp", "convert"); await sel("editConvertTo", "date"); await apply(); }
  const cellEdits = [];
  if (dateCol.d > 0) {
    // pierwsza komórka z datą i pierwsza PUSTA w kolumnie dat (w zakresie danych)
    const targets = await page.evaluate((col) => {
      const withDate = baseRows.find((r) => r.values[col] instanceof Date);
      const empty = baseRows.find((r) => r.values[col] == null || r.values[col] === "");
      return { withDate: withDate && withDate.rowIndex0, empty: empty && empty.rowIndex0 };
    }, dateCol.i);
    const typeInto = async (rowIndex0, col, text) => {
      await page.evaluate(({ rowIndex0, col }) => { if (typeof setSidebarOpen === "function") setSidebarOpen(false); }, { rowIndex0, col });
      await sleep(250);
      const okOpen = await page.evaluate(({ rowIndex0, col }) => {
        const td = document.querySelector(`tbody tr[data-row-key="wide:${rowIndex0}"] td[data-col-index="${col}"]`);
        if (!td) return false;
        setFocusedCell(`wide:${rowIndex0}`, col, { scroll: true });
        openCellEditor(td);
        return true;
      }, { rowIndex0, col });
      if (!okOpen) { notes.push(`edytor: wiersz ${rowIndex0} poza wyrenderowanym widokiem — pominięty`); return false; }
      await page.waitForSelector("input.cell-editor", { timeout: 3000 });
      await page.fill("input.cell-editor", text);
      await page.keyboard.press("Enter");
      await sleep(200);
      await page.keyboard.press("Escape").catch(() => {});
      return true;
    };
    if (targets.withDate != null && await typeInto(targets.withDate, dateCol.i, "12.05.24")) cellEdits.push({ rowIndex0: targets.withDate, col: dateCol.i, want: "2024-05-12" });
    if (targets.empty != null && await typeInto(targets.empty, dateCol.i, "13-05-24")) cellEdits.push({ rowIndex0: targets.empty, col: dateCol.i, want: "2024-05-13" });
    if (targets.empty == null) notes.push("brak pustej komórki w kolumnie dat — test wpisania daty w pustą pominięty");
  }
  const saved = await page.evaluate(async (cellEdits) => {
    const bytes = await buildOutputBytes("xlsx");
    const problems = [];
    // XML poprawny
    const zp = await JSZip.loadAsync(bytes);
    for (const nm of Object.keys(zp.files)) {
      if (!/\.(xml|rels)$/i.test(nm)) continue;
      const txt = await zp.file(nm).async("string");
      if (new DOMParser().parseFromString(txt, "application/xml").getElementsByTagName("parsererror").length) problems.push("malformed:" + nm);
    }
    const origWb = XLSX.read(originalFileBytes, { cellNF: true, cellStyles: true });
    const newWb = XLSX.read(bytes, { cellNF: true, cellStyles: true });
    const pend = pendingEdits;
    let changedOk = 0, untouchedOk = 0;
    for (const sn of origWb.SheetNames) {
      const a = origWb.Sheets[sn], b = newWb.Sheets[sn];
      if (!b) { problems.push("zniknął arkusz " + sn); continue; }
      const refs = new Set([...Object.keys(a), ...Object.keys(b)].filter((r) => r[0] !== "!"));
      for (const ref of refs) {
        const p = pend[sn] && Object.prototype.hasOwnProperty.call(pend[sn], ref) ? pend[sn][ref] : undefined;
        const ca = a[ref], cb = b[ref];
        if (p === undefined) {
          const va = ca ? (ca.f ? "f:" + ca.f : ca.v) : undefined, vb = cb ? (cb.f ? "f:" + cb.f : cb.v) : undefined;
          if (String(va) !== String(vb)) { if (problems.length < 15) problems.push(`ZMIENIONA bez edycji ${sn}!${ref}: ${va} → ${vb}`); }
          else untouchedOk++;
        } else if (p === null) {
          if (cb && cb.v != null && cb.v !== "") problems.push(`${sn}!${ref} miała być pusta`);
          else changedOk++;
        } else {
          const want = p.t === "d" ? Math.round((p.v.getTime() - p.v.getTimezoneOffset() * 60000) / 86400000 + 25569) : p.v;
          const got = cb ? cb.v : undefined;
          const same = p.t === "d" ? Math.abs(Number(got) - want) < 1e-6 : String(got) === String(want);
          if (!same) { if (problems.length < 15) problems.push(`${sn}!${ref}: w pliku ${got}, miało być ${want} (${p.t})`); }
          else changedOk++;
        }
      }
    }
    // daty wpisane w edytorze: typ liczba + format daty (z), żeby Excel pokazał datę
    const dates = cellEdits.map(({ rowIndex0, col }) => {
      const ref = XLSX.utils.encode_cell({ r: rowIndex0, c: currentStartCol + col });
      const c = newWb.Sheets[currentSheetName][ref];
      const d = XLSX.read(bytes, { cellDates: true }).Sheets[currentSheetName][ref];
      return { ref, t: c && c.t, z: c && c.z, iso: d && d.v instanceof Date ? d.v.toISOString().slice(0, 10) : d && d.v };
    });
    // KAŻDA zapisana data (z edytora i z konwersji) ma format daty i całkowity serial
    let dateEdits = 0, dateNoFmt = [], dateFrac = [];
    const nb = newWb.Sheets[currentSheetName];
    Object.entries(pend[currentSheetName] || {}).forEach(([ref, p]) => {
      if (!p || p.t !== "d") return;
      dateEdits++;
      const c = nb[ref];
      const z = c && String(c.z || "");
      const isDate = c && c.t === "n" && /[dmy]/i.test(z.replace(/"[^"]*"|\[[^\]]*\]/g, "")) && !/general/i.test(z);
      if (!isDate && dateNoFmt.length < 5) dateNoFmt.push(`${ref}: z=${z} v=${c && c.v}`);
      if (c && p.v.getHours() === 0 && Math.abs(c.v - Math.round(c.v)) > 1e-9 && dateFrac.length < 5) dateFrac.push(`${ref}: ${c.v}`);
    });
    if (dateNoFmt.length) problems.push("daty BEZ formatu daty: " + dateNoFmt.join("; "));
    if (dateFrac.length) problems.push("daty z przesunięciem godzin: " + dateFrac.join("; "));
    return { dateEdits, problems, changedOk, untouchedOk, pendingTotal: Object.values(pend).reduce((n, s) => n + Object.keys(s).length, 0), dates };
  }, cellEdits);
  ok("ZAPIS: brak problemów", saved.problems.length === 0, saved.problems);
  ok("ZAPIS: wszystkie edycje w pliku", saved.changedOk === saved.pendingTotal, saved);
  notes.push(`zapis: ${saved.changedOk}/${saved.pendingTotal} edycji w pliku (w tym dat: ${saved.dateEdits}), ${saved.untouchedOk} komórek bez zmian identycznych`);
  saved.dates.forEach((d, i) => {
    const e = cellEdits[i];
    const dateFmt = d.z && /[dmy]/i.test(String(d.z)) && !/general/i.test(String(d.z));
    ok(`Data z edytora ${d.ref}: wartość ${e.want}`, d.iso && Math.abs(new Date(d.iso) - new Date(e.want)) <= 86400000, d);
    ok(`Data z edytora ${d.ref}: w Excelu wyświetli się jako data (format „${d.z}”)`, d.t === "n" && dateFmt, d);
  });

  ok("brak błędów konsoli", errors.length === 0, errors.slice(0, 5));
  await context.close();
  return { fails, notes };
}

(async () => {
  const browser = await chromium.launch({ headless: true });
  let total = 0;
  for (const f of FILES) {
    const name = path.basename(f.split("::")[0]) + (f.includes("::") ? " → " + f.split("::")[1] : "");
    let res;
    try { res = await auditFile(browser, f); } catch (e) { res = { fails: ["WYJĄTEK: " + e.message.split("\n")[0]], notes: [] }; }
    console.log(`\n══ ${name}`);
    res.notes.forEach((n) => console.log("   · " + n));
    if (res.fails.length) { total += res.fails.length; res.fails.forEach((x) => console.log("   ❌ " + x)); }
    else console.log("   ✅ wszystko OK");
  }
  await browser.close();
  console.log(total ? `\n❌ audyt: ${total} problemów` : "\n✅ audyt narzędzi edycji: wszystko OK");
  process.exit(total ? 1 : 0);
})();
