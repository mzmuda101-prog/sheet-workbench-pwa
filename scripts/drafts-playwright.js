// drafts-playwright.js — szkic niezapisanej pracy i odzyskiwanie (app/drafts.js).
// Przeniesione z Documents Workbench (2026-10-01).
//
// Symulacja iOS: edycja komórek → aplikacja w tle (visibilitychange) → karta ZABITA bez żadnego
// zdarzenia zamknięcia → nowe otwarcie: karta „Niezapisana praca” → „Przywróć” odtwarza zmienione
// komórki (arkusz, zapis, tabela) i otwarty arkusz. Plus: zapis kasuje szkic, „Odrzuć”, otwarcie
// innego pliku przy niezapisanych zmianach PYTA („Anuluj” zostawia pracę), żyjące drugie okno,
// szkic starszy niż 7 dni znika, bez IndexedDB bez błędów.
// Uruchom z serwerem na APP_URL (domyślnie http://127.0.0.1:4175/). ENGINE=webkit (Safari/iPad).

const { chromium, webkit } = require("playwright");
const path = require("path");

const APP_URL = process.env.APP_URL || "http://127.0.0.1:4175/";
const FILE = path.join(__dirname, "stress-test-workbench.xlsx");
const ENGINE_NAME = process.env.ENGINE === "webkit" ? "webkit" : "chromium";
const ENGINE = ENGINE_NAME === "webkit" ? webkit : chromium;
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
const results = [];
const check = (name, ok, detail) => results.push({ name, ok: !!ok, detail });
const errors = [];
let confirmAnswer = true;

async function openPage(context) {
  const page = await context.newPage();
  page.on("pageerror", (e) => errors.push(e.message));
  page.on("dialog", (d) => (d.type() === "confirm" && !confirmAnswer ? d.dismiss() : d.accept(d.defaultValue())));
  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(() => { document.getElementById("heroSplash")?.remove(); try { ensureXlsxLibs(false); } catch {} });
  await sleep(400);
  return page;
}
async function loadFile(page) {
  await page.setInputFiles("#fileInput", FILE);
  await page.waitForFunction(() => document.getElementById("sheetSelect")?.options?.length > 0 && !!workbook, null, { timeout: 15000 });
  await sleep(300);
  await page.evaluate(() => loadBtn.click()); // przycisk siedzi w panelu bocznym (bywa schowany)
  await page.waitForFunction(() => !!currentDisplayModel && document.querySelectorAll("tbody tr").length > 0, null, { timeout: 15000 });
  await sleep(500);
  await page.evaluate(() => { if (typeof setSidebarOpen === "function") setSidebarOpen(false); });
}
async function editCell(page, idx, text) {
  const rowIndex0 = await page.evaluate((i) => currentDisplayModel.rows[i].rowIndex0, idx);
  await page.evaluate((r) => {
    setFocusedCell(`wide:${r}`, 2, { scroll: true });
    openCellEditor(document.querySelector(`tbody tr[data-row-key="wide:${r}"] td[data-col-index="2"]`));
  }, rowIndex0);
  await page.waitForSelector("input.cell-editor");
  await page.fill("input.cell-editor", text);
  await page.keyboard.press("Enter");
  await sleep(150);
  await page.evaluate(() => document.activeElement?.blur?.());
  return rowIndex0;
}
async function backgroundAndKill(page) {
  await page.evaluate(async () => {
    Object.defineProperty(document, "visibilityState", { configurable: true, get: () => "hidden" });
    document.dispatchEvent(new Event("visibilitychange"));
    await swbDrafts._flush();
  });
  await page.close({ runBeforeUnload: false });
}
const cards = (page) => page.evaluate(async () => {
  await new Promise((r) => setTimeout(r, 700));
  return [...document.querySelectorAll("#draftRecovery .draft-card")].map((c) => ({ name: c.querySelector(".draft-name")?.textContent, meta: c.querySelector(".draft-meta")?.textContent }));
});
const dbCount = (page) => page.evaluate(() => new Promise((resolve) => {
  const r = indexedDB.open("swb-drafts", 1);
  r.onsuccess = () => { const q = r.result.transaction("drafts", "readonly").objectStore("drafts").count(); q.onsuccess = () => { resolve(q.result); r.result.close(); }; };
  r.onerror = () => resolve(-1);
}));
const cellState = (page, rows, texts) => page.evaluate(({ rows, texts }) => rows.map((r, i) => {
  const ref = XLSX.utils.encode_cell({ r, c: currentStartCol + 2 });
  const td = document.querySelector(`tbody tr[data-row-key="wide:${r}"] td[data-col-index="2"]`);
  return { sheet: workbook.Sheets[currentSheetName][ref]?.v === texts[i], save: pendingEdits[currentSheetName]?.[ref]?.v === texts[i], view: !td || td.textContent === texts[i] };
}), { rows, texts });

async function run() {
  const browser = await ENGINE.launch({ headless: true });
  const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1280, height: 900 } });
  await context.addInitScript(() => localStorage.setItem("introPlayed", "true"));

  // ── 1. edycja → tło → zabicie → odzyskanie ─────────────────────────────────
  let page = await openPage(context);
  await loadFile(page);
  const sheetName = await page.evaluate(() => currentSheetName);
  const headerBefore = await page.evaluate(() => { headerRowEl.value = "2"; autoHeaderRowEl.checked = false; loadBtn.click(); return headerRowEl.value; });
  await page.waitForFunction(() => String(currentHeaderRow) === "2" && !!currentDisplayModel, null, { timeout: 15000 });
  await sleep(500);
  const r1 = await editCell(page, 1, "SZKIC-A");
  const r2 = await editCell(page, 2, "SZKIC-B");
  await backgroundAndKill(page);

  page = await openPage(context);
  let c = await cards(page);
  check("po zabiciu aplikacji: karta z nazwą pliku, liczbą komórek i arkuszem", c.length === 1 && c[0].name === "stress-test-workbench.xlsx" && /zmienione komórki: 2/.test(c[0].meta) && c[0].meta.includes(sheetName), JSON.stringify(c));
  await page.click("#draftRecovery .btn.primary");
  await page.waitForFunction(() => !!currentDisplayModel && document.querySelectorAll("tbody tr").length > 0, null, { timeout: 15000 });
  await sleep(800);
  let st = await cellState(page, [r1, r2], ["SZKIC-A", "SZKIC-B"]);
  const misc = await page.evaluate(() => ({ dirty: hasUnsavedChanges, sheet: currentSheetName, card: !document.getElementById("draftRecovery").hidden, header: headerRowEl.value }));
  check("„Przywróć”: komórki wracają (arkusz, zapis, tabela), ten sam arkusz, „Zapisz” świeci",
    st.every((x) => x.sheet && x.save && x.view) && misc.dirty && misc.sheet === sheetName && !misc.card, JSON.stringify({ st, misc }));
  check("arkusz wraca z tym samym wierszem nagłówka co przed zabiciem", misc.header === headerBefore, `${headerBefore} → ${misc.header}`);
  const bytesOk = await page.evaluate(async () => {
    const out = await buildPatchedXlsx(originalFileBytes, pendingEdits, buildSheetFormulaMaps());
    const wb = XLSX.read(out, { cellDates: true });
    return Object.values(wb.Sheets[currentSheetName]).some((cell) => cell && cell.v === "SZKIC-A");
  });
  check("plik do zapisu po odzyskaniu zawiera zmiany", bytesOk);

  // dalsza praca nadpisuje ten sam szkic
  await editCell(page, 3, "SZKIC-C");
  await backgroundAndKill(page);
  page = await openPage(context);
  c = await cards(page);
  check("praca po odzyskaniu → nadal jeden szkic, już z 3 komórkami", c.length === 1 && /zmienione komórki: 3/.test(c[0].meta) && (await dbCount(page)) === 1, JSON.stringify(c));
  await page.click("#draftRecovery .btn.primary");
  await page.waitForFunction(() => !!currentDisplayModel, null, { timeout: 15000 });
  await sleep(800);

  // ── 2. otwarcie innego pliku przy niezapisanych zmianach pyta ──────────────
  confirmAnswer = false;
  await page.evaluate(() => loadSampleFile());
  await sleep(800);
  const kept = await page.evaluate(() => ({ name: currentFileName, dirty: hasUnsavedChanges }));
  check("inny plik przy niezapisanych zmianach: pyta, „Anuluj” zostawia pracę", kept.name === "stress-test-workbench.xlsx" && kept.dirty, JSON.stringify(kept));
  confirmAnswer = true;

  // ── 3. zapis kasuje szkic ──────────────────────────────────────────────────
  if (ENGINE_NAME === "chromium") {
    await page.evaluate(() => { window.showSaveFilePicker = async () => ({ kind: "file", name: "kopia.xlsx", createWritable: async () => ({ write: async () => {}, close: async () => {} }), queryPermission: async () => "granted", requestPermission: async () => "granted" }); });
  }
  await page.evaluate(() => { saveWorkbook(); });
  await sleep(600);
  // Safari (bez File System Access): okienko z nazwą pliku → zatwierdź domyślną
  if (await page.evaluate(() => !document.getElementById("saveAsModal")?.classList.contains("hidden"))) {
    await page.evaluate(() => document.getElementById("saveAsForm").requestSubmit());
  }
  await sleep(1500);
  check("po zapisie: brak niezapisanych zmian i brak szkicu", !(await page.evaluate(() => hasUnsavedChanges)) && (await dbCount(page)) === 0, String(await dbCount(page)));

  // ── 4. „Odrzuć” ────────────────────────────────────────────────────────────
  await editCell(page, 4, "DO-ODRZUCENIA");
  await backgroundAndKill(page);
  page = await openPage(context);
  c = await cards(page);
  await page.click("#draftRecovery .btn:not(.primary)");
  await sleep(500);
  check("„Odrzuć” (po potwierdzeniu) usuwa szkic i kartę", c.length === 1 && (await dbCount(page)) === 0 && (await page.evaluate(() => document.getElementById("draftRecovery").hidden)), JSON.stringify(c));

  // ── 5. otwarcie innego pliku z „OK” porzuca też szkic ──────────────────────
  await loadFile(page);
  await editCell(page, 1, "PORZUC");
  await sleep(1800);
  const before = await dbCount(page);
  await page.evaluate(() => loadSampleFile()); // pyta → OK
  await sleep(1500);
  check("inny plik + „OK”: zmiany porzucone świadomie → szkic usunięty", before === 1 && (await dbCount(page)) === 0, `${before} → ${await dbCount(page)}`);

  // ── 6. żyjące drugie okno nie widzi szkicu pierwszego ──────────────────────
  await loadFile(page);
  await editCell(page, 1, "ZYJE");
  await sleep(1800);
  const other = await openPage(context);
  c = await cards(other);
  check("drugie okno: szkic żyjącego okna nie jest pokazywany do odzyskania", c.length === 0 && (await dbCount(other)) === 1, JSON.stringify(c));
  await other.close();

  // ── 7. starszy niż 7 dni znika ─────────────────────────────────────────────
  await backgroundAndKill(page);
  page = await openPage(context);
  await page.evaluate(() => new Promise((resolve) => {
    const r = indexedDB.open("swb-drafts", 1);
    r.onsuccess = () => {
      const tr = r.result.transaction("drafts", "readwrite");
      const s = tr.objectStore("drafts");
      s.getAll().onsuccess = (e) => { e.target.result.forEach((d) => s.put({ ...d, savedAt: Date.now() - 8 * 86400000 })); };
      tr.oncomplete = () => { r.result.close(); resolve(); };
    };
  }));
  await page.reload({ waitUntil: "load" });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  c = await cards(page);
  check("szkic starszy niż 7 dni: brak karty i sam się usuwa", c.length === 0 && (await dbCount(page)) === 0, JSON.stringify(c));
  await page.close();

  // ── 8. bez IndexedDB aplikacja działa normalnie ───────────────────────────
  const ctx2 = await browser.newContext({ serviceWorkers: "block" });
  await ctx2.addInitScript(() => { localStorage.setItem("introPlayed", "true"); Object.defineProperty(window, "indexedDB", { value: undefined }); });
  const p2 = await ctx2.newPage();
  const err2 = [];
  p2.on("pageerror", (e) => err2.push(e.message));
  await p2.goto(APP_URL, { waitUntil: "load" });
  await p2.evaluate(() => { document.getElementById("heroSplash")?.remove(); try { ensureXlsxLibs(false); } catch {} });
  await loadFile(p2);
  await editCell(p2, 1, "BEZ-IDB");
  await p2.evaluate(async () => { document.dispatchEvent(new Event("visibilitychange")); await swbDrafts._flush(); });
  check("bez IndexedDB: edycja działa, brak błędów", (await p2.evaluate(() => hasUnsavedChanges)) && !err2.length, err2.join(" | "));
  await ctx2.close();

  if (errors.length) check("brak błędów strony", false, errors.slice(0, 3).join(" | "));
  await browser.close();
  let failed = 0;
  for (const r of results) {
    console.log(`${r.ok ? "✅" : "❌"} ${r.name}${!r.ok && r.detail ? `  (${r.detail})` : ""}`);
    if (!r.ok) failed += 1;
  }
  if (failed) { console.error(`\n[${ENGINE_NAME}] ${failed} z ${results.length} nie przeszło`); process.exit(1); }
  console.log(`\n✅ szkic i odzyskiwanie [${ENGINE_NAME}]: ${results.length}/${results.length}`);
}

run().catch((e) => { console.error(e); process.exit(1); });
