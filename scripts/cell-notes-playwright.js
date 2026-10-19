// cell-notes-playwright.js — notatki komórek jak w Excelu (app/cell-notes.js + zapis w xlsx-patch.js).
//
// Prośba Mateusza (2026-10-19): aplikacja tylko liczyła komentarze z pliku. Co musi być prawdą
// (plik z openpyxl: arkusz „Dane”, nagłówek w wierszu 2, notatka w K7):
//   1. notatka z pliku: trójkącik w rogu komórki + treść w podpowiedzi (najechanie myszą),
//   2. menu komórki → „Dodaj notatkę”: okno, Ctrl+Enter zapisuje, znacznik się pojawia,
//      wpis trafia do pendingEdits (to idzie do pliku),
//   3. zmiana WARTOŚCI komórki z notatką i jej wyczyszczenie nie kasują notatki (jak Excel),
//   4. Cofnij/Ponów obejmują notatkę,
//   5. Shift+F2 na zaznaczonej komórce edytuje notatkę,
//   6. menu → „Usuń notatkę” (tu: notatka z pliku, K7),
//   7. ZAPIS: openpyxl widzi dokładnie te notatki, schemat Office bez nowych błędów,
//      ponowne wczytanie w aplikacji pokazuje znaczniki,
//   8. szkic (drafts.js): po zabiciu aplikacji „Przywróć” oddaje dopisaną notatkę,
//   9. komentarz wątkowy (Excel 365): tylko do odczytu, polskie znaki (SheetJS gubi UTF-8).

const pw = require("playwright");
const path = require("path");
const fs = require("fs");
const os = require("os");
const { spawnSync } = require("child_process");
const ooxml = require("./ooxml-validate");
const VENV_PY = path.join(__dirname, "..", ".venv/bin/python");
const PY = fs.existsSync(VENV_PY) ? VENV_PY : "python3";

const APP_URL = process.env.APP_URL || "http://127.0.0.1:4175/";
const ENGINE = process.env.ENGINE || "chromium";
const SLEEP_SCALE = Math.min(1, Math.max(0.1, Number(process.env.SLEEP_SCALE || 1)));
const sleep = (ms) => new Promise((r) => setTimeout(r, Math.round(ms * SLEEP_SCALE)));
const FILE = path.join(__dirname, "fixtures", "openpyxl-features.xlsx");

async function run() {
  const browser = await pw[ENGINE].launch({ headless: true });
  const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1280, height: 860 } });
  await context.addInitScript(() => { localStorage.setItem("introPlayed", "true"); });
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  page.on("dialog", (d) => d.accept());
  const failures = [];
  const ok = (name, cond, got) => { if (!cond) failures.push(`${name} — dostałem: ${JSON.stringify(got)}`); };

  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(() => { document.getElementById("heroSplash")?.remove(); return ensureXlsxLibs(false); });
  await page.setInputFiles("#fileInput", FILE);
  await page.waitForFunction(() => currentHeaderRow === 2 && !!document.querySelector("#dataTable tbody tr[data-row-key]"), null, { timeout: 20000 });
  await sleep(400);

  const td = (r, c) => `#dataTable tbody tr[data-row-key="wide:${r}"] td[data-col-index="${c}"]`;
  const hasNote = (r, c) => page.evaluate((s) => !!document.querySelector(s)?.classList.contains("cell-has-note"), td(r, c));
  const noteAt = (ref) => page.evaluate((ref) => { const n = swbNoteOf(currentSheetName, ref); return n ? n.text : null; }, ref);
  const pending = (ref) => page.evaluate((ref) => JSON.stringify(pendingEdits[currentSheetName]?.[ref] ?? null), ref);
  // Przewinięcie tabeli zamyka menu komórki — najpierw komórka w widoku, potem klik w punkt.
  // (spóźnione zdarzenie scroll po przewinięciu w bok potrafi przyjść już po otwarciu — ponawiamy;
  // toasty z poprzednich kroków piętrzą się w prawym rogu i zasłaniają komórkę — chowamy je).
  const menuItem = async (r, c, cls) => {
    await page.locator(td(r, c)).scrollIntoViewIfNeeded();
    for (let attempt = 0; ; attempt++) {
      await sleep(250);
      const box = await page.locator(td(r, c)).boundingBox();
      await page.evaluate(() => document.querySelectorAll(".toast").forEach((x) => x.remove())); // stos toastów zasłania prawy róg
      await page.mouse.click(box.x + box.width / 2, box.y + box.height / 2, { button: "right" });
      try {
        await page.waitForSelector(`#cellMenu:not(.hidden) .cmi-${cls}`, { timeout: 1500 });
        break;
      } catch (e) {
        if (attempt >= 2) throw e;
      }
    }
    await page.click(`#cellMenu .cmi-${cls}`);
  };
  const modalOpen = () => page.evaluate(() => !!document.querySelector("#noteModal:not(.hidden)"));

  // 1. notatka z pliku (K7 = wiersz 6, kolumna 10)
  const k7 = await page.evaluate(() => ({ start: currentStartCol, text: swbNoteOf(currentSheetName, "K7")?.text }));
  ok("1. notatka K7 wczytana z pliku", k7.start === 0 && /Notatka testowa/.test(k7.text || ""), k7);
  ok("1. znacznik notatki na K7", await hasNote(6, 10), null);
  ok("1. bez znacznika na komórce bez notatki", !(await hasNote(6, 9)), null);
  // najpierw w widok (przewinięcie chowa podpowiedź), potem najechanie
  await page.locator(td(6, 10)).scrollIntoViewIfNeeded();
  await sleep(300);
  const kBox = await page.locator(td(6, 10)).boundingBox();
  await page.mouse.move(kBox.x + kBox.width / 2, kBox.y + kBox.height / 2, { steps: 3 });
  await sleep(200);
  const tip = await page.evaluate(() => ({ shown: !document.getElementById("cellTooltip").classList.contains("hidden"), text: document.querySelector("#cellTooltip .cell-tooltip-usernote")?.textContent || "" }));
  ok("1. najechanie pokazuje treść notatki", tip.shown && /Notatka testowa — ąęł/.test(tip.text), tip);
  await page.mouse.move(5, 5);

  // 2. dodanie przez menu komórki (A4 = wiersz 3, kolumna 0)
  await menuItem(3, 0, "note");
  await page.waitForSelector("#noteModal:not(.hidden) .note-modal-text", { timeout: 3000 });
  const title = await page.evaluate(() => document.getElementById("noteModalTitle").textContent);
  ok("2. okno „Nowa notatka — A4”", /A4/.test(title), title);
  await page.fill("#noteModal .note-modal-author-input", "Mateusz");
  await page.fill("#noteModal .note-modal-text", "Sprawdzić z fakturą\nlinia 2 <&> 🔥");
  await page.keyboard.press("ControlOrMeta+Enter");
  await sleep(200);
  ok("2. okno zamknięte po Ctrl+Enter", !(await modalOpen()), null);
  ok("2. znacznik na A4", await hasNote(3, 0), null);
  ok("2. treść w komórce", (await noteAt("A4")) === "Sprawdzić z fakturą\nlinia 2 <&> 🔥", await noteAt("A4"));
  const p2 = JSON.parse(await pending("A4"));
  ok("2. wpis do zapisu z autorem", p2 && p2.note && p2.note.author === "Mateusz" && !("t" in p2), p2);
  ok("2. autor zapamiętany", (await page.evaluate(() => localStorage.getItem("swb-note-author-v1"))) === "Mateusz", null);

  // 3. zmiana wartości i wyczyszczenie nie kasują notatki
  await page.dblclick(td(3, 0));
  await page.waitForSelector(`${td(3, 0)} input.cell-editor`, { timeout: 3000 });
  await page.keyboard.press("ControlOrMeta+a");
  await page.keyboard.type("Nowa wartość");
  await page.keyboard.press("Enter");
  await sleep(250);
  const p3 = JSON.parse(await pending("A4"));
  ok("3. nowa wartość, notatka została (komórka + zapis)", (await noteAt("A4")) && p3.v === "Nowa wartość" && p3.note && /faktur/.test(p3.note.text) && (await hasNote(3, 0)), p3);
  await page.evaluate(() => { setFocusedCell("wide:3", 0, { scroll: false }); cfmtClearContents(); });
  await sleep(250);
  const p3b = JSON.parse(await pending("A4"));
  ok("3. wyczyszczenie zostawia notatkę", (await noteAt("A4")) && p3b.clear && p3b.note && (await hasNote(3, 0)), p3b);

  // 4. Cofnij/Ponów: cofamy wyczyszczenie i wartość, potem samą notatkę
  await page.evaluate(() => { undoRedo("undo"); undoRedo("undo"); });
  await sleep(200);
  ok("4. po cofnięciu wartości notatka dalej jest", (await noteAt("A4")) !== null && (await hasNote(3, 0)), null);
  await page.evaluate(() => undoRedo("undo"));
  await sleep(250);
  ok("4. Cofnij usuwa dodaną notatkę (komórka, znacznik, zapis)", (await noteAt("A4")) === null && !(await hasNote(3, 0)) && (await pending("A4")) === "null", { note: await noteAt("A4"), pend: await pending("A4") });
  await page.evaluate(() => undoRedo("redo"));
  await sleep(250);
  ok("4. Ponów przywraca notatkę", (await noteAt("A4")) !== null && (await hasNote(3, 0)), null);

  // 5. Shift+F2 edytuje notatkę zaznaczonej komórki
  await page.click(td(3, 0));
  await sleep(100);
  await page.keyboard.press("Shift+F2");
  await page.waitForSelector("#noteModal:not(.hidden)", { timeout: 3000 });
  const pre5 = await page.evaluate(() => ({ text: document.querySelector("#noteModal .note-modal-text").value, authorRo: document.querySelector("#noteModal .note-modal-author-input").readOnly }));
  ok("5. Shift+F2 otwiera notatkę z treścią, autor istniejącej tylko do odczytu", /faktur/.test(pre5.text) && pre5.authorRo, pre5);
  await page.fill("#noteModal .note-modal-text", "Po edycji");
  await page.click("#noteModal .note-modal-save");
  await sleep(200);
  ok("5. zapisana zmieniona treść", (await noteAt("A4")) === "Po edycji", await noteAt("A4"));
  // Esc od razu po otwarciu (wcześniej, przed fokusem w polu, trafiał do siatki)
  await page.keyboard.press("Shift+F2");
  await page.keyboard.press("Escape");
  await sleep(150);
  ok("5. Esc zamyka okno bez zmian", !(await modalOpen()) && (await noteAt("A4")) === "Po edycji", null);

  // 6. usunięcie notatki z pliku przez menu
  await menuItem(6, 10, "note-del");
  await sleep(200);
  ok("6. „Usuń notatkę” kasuje K7", (await noteAt("K7")) === null && !(await hasNote(6, 10)) && JSON.parse(await pending("K7")).note === null, await pending("K7"));
  // druga notatka: komórka bez wartości (może mieć sam styl z pliku)
  const emptyCell = await page.evaluate(() => {
    const sheet = workbook.Sheets[currentSheetName];
    for (let r = 3; r < 40; r++) for (let c = 0; c < 12; c++) {
      const ref = XLSX.utils.encode_cell({ r, c });
      const cell = sheet[ref];
      if ((!cell || (cell.v == null && !cell.f && !cell.c)) && document.querySelector(`#dataTable tbody tr[data-row-key="wide:${r}"] td[data-col-index="${c}"]`)) return { r, c, ref };
    }
    return null;
  });
  ok("6. jest pusta komórka do testu", !!emptyCell, emptyCell);
  if (emptyCell) {
    await menuItem(emptyCell.r, emptyCell.c, "note");
    await page.waitForSelector("#noteModal:not(.hidden)", { timeout: 3000 });
    await page.fill("#noteModal .note-modal-text", "Na pustej");
    await page.keyboard.press("ControlOrMeta+Enter");
    await sleep(200);
    ok("6. notatka na pustej komórce", (await noteAt(emptyCell.ref)) === "Na pustej" && (await hasNote(emptyCell.r, emptyCell.c)), null);
  }

  // 7. ZAPIS
  const tmp = fs.mkdtempSync(path.join(os.tmpdir(), "cell-notes-"));
  const out = path.join(tmp, "notatki.xlsx");
  const b64 = await page.evaluate(async () => {
    const u8 = await buildOutputBytes("xlsx");
    let s = ""; for (let i = 0; i < u8.length; i += 0x8000) s += String.fromCharCode(...u8.subarray(i, i + 0x8000));
    return btoa(s);
  });
  fs.writeFileSync(out, Buffer.from(b64, "base64"));
  const py = spawnSync(PY, ["-c", `
import json, sys
from openpyxl import load_workbook
wb = load_workbook(sys.argv[1])
print(json.dumps({ws.title: {c.coordinate: [c.comment.text, c.comment.author] for row in ws.iter_rows() for c in row if c.comment} for ws in wb}))
`, out], { encoding: "utf8" });
  let f = null;
  try { f = JSON.parse(py.stdout); } catch { failures.push("7. openpyxl nie przeczytał zapisu: " + (py.stderr || "").slice(0, 300)); }
  if (f) {
    const want = { A4: ["Po edycji", "Mateusz"] };
    if (emptyCell) want[emptyCell.ref] = ["Na pustej", "Mateusz"];
    ok("7. plik: dokładnie nasze notatki w „Dane” (K7 usunięta)", JSON.stringify(f.Dane) === JSON.stringify(Object.fromEntries(Object.keys(want).sort((a, b) => a.localeCompare(b, "en", { numeric: true })).map((k) => [k, want[k]]))) || (Object.keys(f.Dane).length === Object.keys(want).length && Object.keys(want).every((k) => JSON.stringify(f.Dane[k]) === JSON.stringify(want[k]))), f.Dane);
    ok("7. plik: inne arkusze bez notatek", Object.entries(f).every(([n, m]) => n === "Dane" || !Object.keys(m).length), f);
  }
  if (ooxml.available()) {
    const [vo, vs] = ooxml.validate([FILE, out]);
    const fresh = ooxml.newErrors(vo, vs).map(ooxml.short);
    ok("7. plik zgodny ze schematem Office (bez nowych błędów)", vs && !vs.error && fresh.length === 0, fresh.slice(0, 5));
  } else {
    console.warn("⚠️  schemat Office pominięty — brak .NET / tools/ooxml-validator");
  }
  // ponowne wczytanie
  await page.evaluate(() => { pendingEdits = {}; setDirtyState(false); theadEl.replaceChildren(); tbodyEl.replaceChildren(); });
  await page.setInputFiles("#fileInput", out);
  await page.waitForFunction(() => currentFileName === "notatki.xlsx" && !!document.querySelector("#dataTable tbody tr[data-row-key]"), null, { timeout: 20000 });
  await sleep(300);
  ok("7. po ponownym wczytaniu: znacznik i treść A4, K7 bez notatki", (await hasNote(3, 0)) && (await noteAt("A4")) === "Po edycji" && !(await hasNote(6, 10)), null);

  // 8. szkic: notatka dopisana, aplikacja zabita w tle → „Przywróć” ją oddaje
  await page.evaluate(() => swbSetNote(4, 1, "Ze szkicu", "M"));
  await sleep(150);
  ok("8. notatka B5 przed zabiciem", (await noteAt("B5")) === "Ze szkicu", null);
  await page.evaluate(async () => {
    Object.defineProperty(document, "visibilityState", { configurable: true, get: () => "hidden" });
    document.dispatchEvent(new Event("visibilitychange"));
    await swbDrafts._flush();
  });
  await page.close({ runBeforeUnload: false });
  const page2 = await context.newPage();
  page2.on("pageerror", (e) => errors.push(e.message));
  page2.on("dialog", (d) => d.accept());
  await page2.goto(APP_URL, { waitUntil: "load" });
  await page2.evaluate(() => document.getElementById("heroSplash")?.remove());
  await page2.waitForSelector("#draftRecovery .btn.primary", { timeout: 10000 });
  await page2.click("#draftRecovery .btn.primary");
  await page2.waitForFunction(() => !!currentDisplayModel && !!document.querySelector("#dataTable tbody tr[data-row-key]"), null, { timeout: 20000 });
  await sleep(800);
  const d8 = await page2.evaluate(() => ({
    b5: swbNoteOf(currentSheetName, "B5")?.text || null,
    a4: swbNoteOf(currentSheetName, "A4")?.text || null,
    mark: !!document.querySelector('#dataTable tbody tr[data-row-key="wide:4"] td[data-col-index="1"]')?.classList.contains("cell-has-note"),
    pend: JSON.stringify(pendingEdits[currentSheetName]?.B5 ?? null),
  }));
  ok("8. „Przywróć” oddaje notatkę ze szkicu (komórka, znacznik, zapis) i te z pliku", d8.b5 === "Ze szkicu" && d8.mark && /Ze szkicu/.test(d8.pend) && d8.a4 === "Po edycji", d8);
  // 9. komentarz WĄTKOWY (Excel 365): tylko do odczytu, polskie znaki poprawne, zapis go nie rusza
  await page2.evaluate(() => { pendingEdits = {}; setDirtyState(false); });
  await page2.setInputFiles("#fileInput", path.join(__dirname, "fixtures", "threaded-comments.xlsx"));
  await page2.waitForFunction(() => currentFileName === "threaded-comments.xlsx" && !!document.querySelector("#dataTable tbody tr[data-row-key]"), null, { timeout: 20000 });
  await sleep(400);
  const t9 = await page2.evaluate(() => {
    const n = swbNoteOf(currentSheetName, "A2");
    const r = XLSX.utils.decode_cell("A2").r;
    const opened = swbNoteModal.open(r, 0, {});
    const m = { ro: document.querySelector("#noteModal .note-modal-text").readOnly, save: document.querySelector("#noteModal .note-modal-save").hidden, del: document.querySelector("#noteModal .note-modal-delete").hidden };
    swbNoteModal.close();
    const items = swbNoteMenuItems({ row: { rowIndex0: r }, col: 0, td: null }).map((i) => i.id);
    return { n, opened, m, items, set: swbSetNote(r, 0, "PSUJĘ", "x") };
  });
  ok("9. wątek: treść z polskimi znakami, autorzy, odpowiedź", t9.n && t9.n.threaded && /Jan: Wątek pierwszy/.test(t9.n.text) && /Jan: Odpowiedź/.test(t9.n.text), t9.n);
  ok("9. wątek: okno tylko do odczytu, bez Zapisz/Usuń; menu „Pokaż komentarz”; zapis odrzucony", t9.opened && t9.m.ro && t9.m.save && t9.m.del && t9.items.join() === "note" && t9.set === false, t9);
  fs.rmSync(tmp, { recursive: true, force: true });

  await browser.close();
  if (errors.length) failures.push("błędy strony: " + errors.join(" | "));
  if (failures.length) {
    console.error(`❌ cell-notes (${ENGINE}):\n  - ` + failures.join("\n  - "));
    process.exit(1);
  }
  console.log(`✅ cell-notes (${ENGINE}): znacznik + podpowiedź, dodaj/edytuj/usuń, wartość nie kasuje notatki, Cofnij, Shift+F2, zapis (openpyxl + schemat)`);
}

run().catch((e) => { console.error(e); process.exit(1); });
