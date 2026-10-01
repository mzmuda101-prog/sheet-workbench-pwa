// state-survival-playwright.js — czy NIEZAPISANE zmiany przeżywają „zwykłe” rzeczy wokół.
// Przeniesione z Documents Workbench (2026-10-01): tam okno przeciągnięte na inny monitor
// przerysowało dokument i wpisany tekst znikał z ekranu i z zapisu, choć „Zapisz” świeciło.
//
// Dwa warianty:
//   A) zatwierdzona zmiana komórki → czynność → wartość jest w arkuszu, w pendingEdits
//      (zapis), na ekranie, a „Zapisz” świeci,
//   B) komórka W TRAKCIE edycji (wpisane, bez Enter) → czynność → wpis nie może zniknąć
//      po cichu: albo edytor dalej jest otwarty z tym tekstem, albo wpis jest zatwierdzony.
// Plus „Aktualizuj” / „Odśwież aplikację” przy niezapisanych zmianach (iOS nie pyta
// „Opuścić stronę?”). Uruchom z serwerem na APP_URL (domyślnie http://127.0.0.1:4175/).

const { chromium, webkit } = require("playwright");
const path = require("path");

const APP_URL = process.env.APP_URL || "http://127.0.0.1:4175/";
const FILE = path.join(__dirname, "stress-test-workbench.xlsx");
const ENGINE_NAME = process.env.ENGINE === "webkit" ? "webkit" : "chromium";
const ENGINE = ENGINE_NAME === "webkit" ? webkit : chromium;
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

async function run() {
  const browser = await ENGINE.launch({ headless: true });
  const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1280, height: 900 } });
  await context.addInitScript(() => localStorage.setItem("introPlayed", "true"));
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push("pageerror: " + e.message));
  page.on("dialog", (d) => d.accept());

  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(() => { document.getElementById("heroSplash")?.remove(); try { ensureXlsxLibs && ensureXlsxLibs(false); } catch {} });
  await page.setInputFiles("#fileInput", FILE);
  await page.waitForFunction(() => document.getElementById("sheetSelect")?.options?.length > 0, null, { timeout: 15000 });
  await sleep(300);
  await page.click("#loadBtn");
  await sleep(1200);
  await page.evaluate(() => { if (typeof setSidebarOpen === "function") setSidebarOpen(false); });
  await sleep(300);

  const results = [];
  const check = (name, ok, detail) => results.push({ name, ok: !!ok, detail });

  // Komórki do testu: kolejne wiersze z widocznej części (kolumna 2).
  const targets = await page.evaluate(() => currentDisplayModel.rows.slice(0, 40).map((r) => r.rowIndex0));
  let ti = 0;
  const markers = []; // { rowIndex0, col, text }

  const openEditor = async (rowIndex0, col) => {
    await page.evaluate(({ rowIndex0, col }) => {
      setFocusedCell(`wide:${rowIndex0}`, col, { scroll: true });
      const td = document.querySelector(`tbody tr[data-row-key="wide:${rowIndex0}"] td[data-col-index="${col}"]`);
      openCellEditor(td);
    }, { rowIndex0, col });
    const okEd = await page.waitForSelector("input.cell-editor", { timeout: 4000 }).then(() => true).catch(() => false);
    if (!okEd) {
      const dbg = await page.evaluate(({ rowIndex0, col }) => {
        const td = document.querySelector(`tbody tr[data-row-key="wide:${rowIndex0}"] td[data-col-index="${col}"]`);
        return { td: !!td, active: document.activeElement?.tagName + "." + document.activeElement?.className, editors: document.querySelectorAll("input.cell-editor").length, ro: typeof readOnlyMode !== "undefined" ? readOnlyMode : "?", mode: currentDisplayModel?.mode, ace: activeCellEditor ? { tdConnected: activeCellEditor.td.isConnected, inputConnected: activeCellEditor.input.isConnected } : null };
      }, { rowIndex0, col });
      throw new Error(`edytor się nie otworzył: ${JSON.stringify(dbg)}`);
    }
  };
  const commitMarker = async () => {
    const m = { rowIndex0: targets[ti++], col: 2, text: `ZN${ti}X` };
    await openEditor(m.rowIndex0, m.col);
    await page.fill("input.cell-editor", m.text);
    await page.keyboard.press("Enter");
    await sleep(150);
    await page.evaluate(() => { document.activeElement?.blur?.(); });
    markers.push(m);
  };
  // stan wszystkich znaczników: arkusz, pendingEdits (zapis), ekran (jeśli wiersz narysowany)
  const verify = async (label) => {
    await sleep(500);
    const r = await page.evaluate((ms) => {
      const lost = [];
      ms.forEach((m) => {
        const ref = XLSX.utils.encode_cell({ r: m.rowIndex0, c: currentStartCol + m.col });
        const cell = workbook.Sheets[currentSheetName][ref];
        const pending = pendingEdits[currentSheetName] && pendingEdits[currentSheetName][ref];
        const td = document.querySelector(`tbody tr[data-row-key="wide:${m.rowIndex0}"] td[data-col-index="${m.col}"]`);
        const pendingVal = pending && (pending.value ?? pending.v ?? pending.parsed?.value ?? JSON.stringify(pending));
        if (!cell || cell.v !== m.text) lost.push(`${m.text}:arkusz`);
        if (!pending || !String(pendingVal).includes(m.text)) lost.push(`${m.text}:zapis`);
        if (td && !td.querySelector("input.cell-editor") && td.textContent !== m.text) lost.push(`${m.text}:ekran(${td.textContent})`);
      });
      return { lost, dirty: hasUnsavedChanges };
    }, markers);
    check(`${label}: zatwierdzone zmiany zostają (arkusz, zapis, ekran)`, !r.lost.length && r.dirty, JSON.stringify(r));
  };
  // komórka w trakcie edycji: czynność nie może zgubić wpisu po cichu
  const verifyOpenEdit = async (label, action) => {
    const m = { rowIndex0: targets[ti++], col: 2, text: `OTW${ti}X` };
    await openEditor(m.rowIndex0, m.col);
    await page.fill("input.cell-editor", m.text);
    await action();
    await sleep(600);
    const r = await page.evaluate((m) => {
      const ref = XLSX.utils.encode_cell({ r: m.rowIndex0, c: currentStartCol + m.col });
      const cell = workbook.Sheets[currentSheetName][ref];
      const ed = document.querySelector("input.cell-editor");
      return { committed: !!cell && cell.v === m.text, editorKeeps: !!ed && ed.value === m.text && ed.isConnected };
    }, m);
    check(`edycja w toku + ${label}: wpis nie znika`, r.committed || r.editorKeeps, JSON.stringify(r));
    // domknij: zatwierdź, jeśli edytor wciąż otwarty
    await page.evaluate(() => { const ed = document.querySelector("input.cell-editor"); if (ed) ed.dispatchEvent(new KeyboardEvent("keydown", { key: "Enter", bubbles: true })); });
    await sleep(200);
    const done = await page.evaluate((m) => { const ref = XLSX.utils.encode_cell({ r: m.rowIndex0, c: currentStartCol + m.col }); return workbook.Sheets[currentSheetName][ref]?.v === m.text; }, m);
    if (done) markers.push(m);
  };
  const step = async (label, action) => { await commitMarker(); await action(); await verify(label); };

  const resizeAround = async () => {
    await page.setViewportSize({ width: 700, height: 900 }); await sleep(500);
    await page.setViewportSize({ width: 1280, height: 900 });
  };
  const rerender = async () => { await page.evaluate(() => { applyFilters(); sortRows(); renderActiveTable(); }); };

  await step("szerokość: komputer → telefon → komputer (inny monitor)", resizeAround);
  await step("szarpanie rozmiarem okna (5×)", async () => {
    for (let i = 0; i < 5; i++) { await page.setViewportSize({ width: i % 2 ? 1280 : 700, height: 900 }); await sleep(60); }
    await page.setViewportSize({ width: 1280, height: 900 });
  });
  await step("obrót (orientationchange)", async () => { await page.evaluate(() => window.dispatchEvent(new Event("orientationchange"))); });
  await step("język PL → EN → PL", async () => { await page.evaluate(() => document.querySelector('#langSwitch [data-lang="en"]').click()); await sleep(200); await page.evaluate(() => document.querySelector('#langSwitch [data-lang="pl"]').click()); });
  await step("motyw jasny ⇄ ciemny", async () => { await page.evaluate(() => { themeToggle.click(); themeToggle.click(); }); });
  await step("tryb czytania wł./wył.", async () => { await page.evaluate(() => { readingToggle?.click(); readingToggle?.click(); }); });
  await step("przerysowanie tabeli (sort/filtr)", rerender);
  await step("szukanie i wyczyszczenie", async () => { await page.fill("#quickSearch", "a"); await sleep(400); await page.fill("#quickSearch", ""); await sleep(400); });
  await step("panel otwarty / zamknięty", async () => { await page.evaluate(() => { setSidebarOpen(true); setSidebarOpen(false); }); });
  await step("karta w tle i z powrotem", async () => {
    await page.evaluate(() => {
      Object.defineProperty(document, "visibilityState", { configurable: true, get: () => "hidden" });
      document.dispatchEvent(new Event("visibilitychange")); window.dispatchEvent(new Event("pagehide"));
      Object.defineProperty(document, "visibilityState", { configurable: true, get: () => "visible" });
      document.dispatchEvent(new Event("visibilitychange")); window.dispatchEvent(new Event("pageshow"));
    });
  });
  await step("okno traci i odzyskuje fokus", async () => { await page.evaluate(() => { window.dispatchEvent(new Event("blur")); window.dispatchEvent(new Event("focus")); }); });
  await step("druk (beforeprint / afterprint)", async () => { await page.evaluate(() => { window.dispatchEvent(new Event("beforeprint")); window.dispatchEvent(new Event("afterprint")); }); });
  await step("daleko w dół i z powrotem (wirtualizacja wierszy)", async () => {
    await page.evaluate(() => { const vp = document.querySelector(".table-wrap, #tableViewport, .table-viewport") || document.scrollingElement; vp.scrollTop = vp.scrollHeight; });
    await sleep(400);
    await page.evaluate(() => { const vp = document.querySelector(".table-wrap, #tableViewport, .table-viewport") || document.scrollingElement; vp.scrollTop = 0; });
  });

  await step("przejście na inny arkusz i z powrotem (Zapisz nie może zgasnąć)", async () => {
    const back = await page.evaluate(() => currentSheetName);
    const other = await page.evaluate((b) => [...sheetSelect.options].map((o) => o.value).find((v) => v !== b), back);
    if (!other) return;
    const load = async (name) => {
      await page.evaluate((n) => { sheetSelect.value = n; loadBtn.click(); }, name);
      await page.waitForFunction((n) => currentSheetName === n && !!currentDisplayModel, name, { timeout: 15000 });
      await sleep(600);
    };
    await load(other);
    const dirtyOnOther = await page.evaluate(() => hasUnsavedChanges);
    check("na innym arkuszu „Zapisz” dalej świeci (zmiany z poprzedniego czekają)", dirtyOnOther);
    await load(back);
  });

  // B) edycja w toku
  await verifyOpenEdit("zmiana szerokości okna (inny monitor)", resizeAround);
  await verifyOpenEdit("przerysowanie tabeli", rerender);
  await verifyOpenEdit("język PL → EN → PL", async () => { await page.evaluate(() => { document.querySelector('#langSwitch [data-lang="en"]').click(); document.querySelector('#langSwitch [data-lang="pl"]').click(); }); });
  await verifyOpenEdit("karta w tle", async () => {
    await page.evaluate(() => {
      Object.defineProperty(document, "visibilityState", { configurable: true, get: () => "hidden" });
      document.dispatchEvent(new Event("visibilitychange"));
      Object.defineProperty(document, "visibilityState", { configurable: true, get: () => "visible" });
      document.dispatchEvent(new Event("visibilitychange"));
    });
  });
  await verifyOpenEdit("obrót", async () => { await page.evaluate(() => window.dispatchEvent(new Event("orientationchange"))); });
  await verify("po edycjach w toku");

  // „Aktualizuj” / „Odśwież aplikację” przy niezapisanych zmianach
  let upd = null;
  try {
    upd = await page.evaluate(async () => {
      window.__stay = "ta sama strona";
      let asked = "";
      window.confirm = (msg) => { asked = msg; return false; };
      await hardRefreshApp();
      const btn = document.getElementById("appUpdateBtn");
      btn?.classList.remove("hidden");
      btn?.click();
      await new Promise((r) => setTimeout(r, 300));
      return { asked: /niezapisane zmiany/.test(asked), busy: !!btn?.classList.contains("is-busy") };
    });
    await sleep(800);
    upd.still = await page.evaluate(() => window.__stay);
  } catch (e) {
    upd = { reloaded: true, err: String(e.message).slice(0, 80) }; // strona się przeładowała = praca przepadła
  }
  check("„Aktualizuj” przy niezapisanych zmianach pyta, „Anuluj” nie przeładowuje", upd.asked && !upd.busy && upd.still === "ta sama strona", JSON.stringify(upd));
  if (!upd.reloaded) {
    await page.evaluate(() => { window.confirm = () => true; document.getElementById("appUpdateBtn")?.classList.add("hidden"); });
    await verify("po anulowanej aktualizacji");
  }

  if (errors.length) check("brak błędów strony", false, errors.slice(0, 3).join(" | "));
  await browser.close();
  let failed = 0;
  for (const r of results) {
    console.log(`${r.ok ? "✅" : "❌"} ${r.name}${!r.ok && r.detail ? `  (${r.detail})` : ""}`);
    if (!r.ok) failed += 1;
  }
  if (failed) { console.error(`\n[${ENGINE_NAME}] ${failed} z ${results.length} nie przeszło`); process.exit(1); }
  console.log(`\n✅ niezapisane zmiany przeżywają [${ENGINE_NAME}]: ${results.length}/${results.length}`);
}

run().catch((e) => { console.error(e); process.exit(1); });
