// entry-ui-playwright.js — Wpisywanie: nakładka (cel → pola → podgląd) na pliku z entry-fixture.js.
//
// Co musi być prawdą:
//   1. #entryBtn otwiera nakładkę, fokus w polu celu,
//   2. dwa identyczne klucze („Ewa Lis”) — dwóch kandydatów z różnymi kolumnami rozróżniającymi,
//      nic nie wybiera się samo (Enter bez wyboru nie przechodzi dalej),
//   3. literówka w kluczu znajduje wiersz; „Nowy wiersz” przy podobnym kluczu → okno z pytaniem,
//   4. „12.03” w kolumnie dat — podgląd „Zapisze się jako” = to, co potem pokazuje komórka,
//   5. kolumna obliczeniowa tylko do odczytu,
//   6. nadpisanie bez zaznaczenia blokuje „Wpisz”; z zaznaczeniem zapis + dymek z Cofnij,
//   7. przypięte pole przechodzi na następną kartkę,
//   8. klawisze w karcie nie przeciekają do tabeli; Esc zamyka,
//   9. tryb Podgląd — odmowa z dokładnym tekstem; filtr — okno z „Wyczyść filtr i otwórz”.
// TOUCH=1 — kontekst dotykowy (hasTouch, isMobile).

const pw = require("playwright");
const { autoAnswerAppDialogs } = require("./app-dialog-auto");
const { makeEntryFixture } = require("./entry-fixture");
const path = require("path");
const fs = require("fs");
const os = require("os");

const APP_URL = process.env.APP_URL || "http://127.0.0.1:4175/";
const ENGINE = process.env.ENGINE || "chromium";
const TOUCH = process.env.TOUCH === "1";
const SLEEP_SCALE = Math.min(1, Math.max(0.1, Number(process.env.SLEEP_SCALE || 1)));
const sleep = (ms) => new Promise((r) => setTimeout(r, Math.round(ms * SLEEP_SCALE)));

async function run() {
  const tmp = fs.mkdtempSync(path.join(os.tmpdir(), "swb-entry-ui-"));
  const FILE = path.join(tmp, "entry.xlsx");
  makeEntryFixture(FILE);
  const browser = await pw[ENGINE].launch({ headless: true });
  const ctxOpts = TOUCH
    ? { serviceWorkers: "block", viewport: { width: 820, height: 1100 }, hasTouch: true, isMobile: ENGINE !== "firefox" }
    : { serviceWorkers: "block", viewport: { width: 1280, height: 860 } };
  const context = await browser.newContext(ctxOpts);
  await context.addInitScript(() => {
    localStorage.setItem("introPlayed", "true");
    if (!sessionStorage.getItem("enTestInit")) { localStorage.removeItem("excel-workbench-entry"); sessionStorage.setItem("enTestInit", "1"); }
  });
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (e) => { errors.push(e.message); if (process.env.DEBUG) console.log("PAGEERROR", e.message, e.stack); });
  await autoAnswerAppDialogs(page);
  const failures = [];
  const ok = (name, cond, got) => { if (!cond) failures.push(`${name} — dostałem: ${JSON.stringify(got)}`); };

  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(() => { document.getElementById("heroSplash")?.remove(); return ensureXlsxLibs(false); });
  await page.setInputFiles("#fileInput", FILE);
  await page.waitForFunction(() => currentHeaderRow === 1 && !!document.querySelector("#dataTable tbody tr[data-row-key]"), null, { timeout: 20000 });
  await page.evaluate(() => { sheetEditUi.setMode("edit"); });
  await sleep(300);

  const st = () => page.evaluate(() => window.__entry.state());
  const cellText = (r, c) => page.evaluate(({ r, c }) => {
    const row = baseRows.find((x) => x.rowIndex0 === r);
    return row ? String(getDisplayValue(row, c) ?? "") : null;
  }, { r, c });
  const openEntry = () => page.evaluate(() => document.getElementById("entryBtn").click());
  const typeTarget = async (text) => { await page.fill("#enTargetInput", text); await sleep(150); };
  const field = (col) => `#entryOverlay [data-en-col="${col}"] input`;

  // 1.
  await openEntry();
  await sleep(250);
  const s1 = await st();
  ok("1. nakładka otwarta", s1.open && s1.step === 1, s1);
  ok("1. widoczna", await page.evaluate(() => !document.getElementById("entryOverlay").classList.contains("hidden")), null);
  ok("1. fokus w polu celu", await page.evaluate(() => document.activeElement?.id === "enTargetInput"), null);
  ok("1. kolumna-klucz = Osoba (0)", s1.keyCol === 0, s1.keyCol);

  // 2.
  await typeTarget("ewa lis");
  const opts2 = await page.$$eval("#enCandidates [role=option]", (els) => els.map((e) => ({ text: e.textContent, sel: e.getAttribute("aria-selected") })));
  const ewa = opts2.filter((o) => /Ewa Lis/.test(o.text));
  ok("2. dwóch kandydatów Ewa Lis", ewa.length === 2, opts2);
  ok("2. różne opisy (kolumny rozróżniające / numer wiersza)", ewa.length === 2 && ewa[0].text !== ewa[1].text, ewa);
  ok("2. nic nie wybrane samo", opts2.every((o) => o.sel !== "true"), opts2);
  await page.press("#enTargetInput", "Enter");
  await sleep(150);
  ok("2. Enter bez wyboru zostaje w kroku 1", (await st()).step === 1, await st());

  // 3. literówka + strażnik duplikatu
  await typeTarget("Jan Kowalsky");
  const first3 = await page.$eval("#enCandidates [role=option]", (e) => e.textContent);
  ok("3. literówka — pierwszy kandydat Jan Kowalski", /Jan Kowalski/.test(first3), first3);
  await page.evaluate(() => { window.__appDialogLog = []; window.__appDialogMode = "dismiss"; });
  await page.click("#enCandidates [data-en-new]");
  await sleep(300);
  const log3 = await page.evaluate(() => window.__appDialogLog);
  ok("3. okno strażnika duplikatu", log3.length === 1 && /Jan Kowalski/.test(log3[0].msg) && log3[0].buttons.some((b) => /Uzupełnij wiersz 2/.test(b)), log3);
  ok("3. po odrzuceniu okna dalej krok 1", (await st()).step === 1, await st());
  await page.evaluate(() => { window.__appDialogMode = "accept"; });

  // 4./5. wybór Ewa Lis (wiersz 3), data częściowa, kolumna obliczeniowa
  await typeTarget("ewa lis");
  await page.press("#enTargetInput", "ArrowDown");
  await page.press("#enTargetInput", "Enter");
  await sleep(200);
  const s4 = await st();
  ok("4. krok 2, cel wiersz arkusza 3 (rowIndex0 2)", s4.step === 2 && s4.target && s4.target.rowIndex0 === 2, s4);
  ok("4. fokus w pierwszym polu", await page.evaluate(() => !!document.activeElement?.closest("[data-en-col]")), null);
  await page.fill(field(1), "12.03");
  await sleep(150);
  // „12.03” to w tej aplikacji liczba (jak w edytorze komórki) — w kolumnie dat podgląd musi
  // pokazać to, co pokaże komórka w jej formacie, ostrzec i zaproponować pełną datę
  const prevNum = await page.$eval('#entryOverlay [data-en-col="1"] .en-will', (e) => e.textContent);
  ok("4a. podgląd liczby w formacie daty (nie „12.03”)", prevNum === "Zapisze się jako: 12.01.1900", prevNum);
  const year = new Date().getFullYear();
  // podgląd zmian też ostrzega — kto przeoczy żółty tekst pod polem, nie zatwierdzi 12.01.1900 po cichu
  await page.keyboard.press("ControlOrMeta+Enter"); await sleep(200);
  const w3 = await page.$$eval("#entryOverlay .en-change .en-warn", (els) => els.map((e) => e.textContent));
  ok("4a. krok 3 ostrzega o typie", w3.some((x) => /daty/.test(x)), w3);
  const struck = await page.$$eval("#entryOverlay .en-change .en-before", (els) => els.map((e) => ({ t: e.textContent, empty: e.classList.contains("is-empty") })));
  ok("4a. puste „było” nie jest przekreślane", struck.every((x) => x.t !== "(puste)" || x.empty), struck);
  await page.click("#enBackBtn"); await sleep(200);
  const fixBtn = await page.$('#entryOverlay [data-en-col="1"] .en-date-fix');
  ok("4a. propozycja pełnej daty", !!fixBtn && (await fixBtn.textContent()).includes(`12.03.${year}`), fixBtn && await fixBtn.textContent());
  if (fixBtn) await fixBtn.click();
  await sleep(150);
  ok("4a. po poprawce pole = 12.03.rok", (await page.inputValue(field(1))) === `12.03.${year}`, await page.inputValue(field(1)));
  const preview = await page.$eval('#entryOverlay [data-en-col="1"] .en-will', (e) => e.textContent);
  ok("5. kolumna Razem tylko do odczytu", await page.$eval(field(4), (e) => e.readOnly), null);
  await page.fill(field(2), "zdana");
  await page.keyboard.press("ControlOrMeta+Enter");
  await sleep(200);
  ok("4. Ctrl+Enter → krok 3", (await st()).step === 3, await st());
  await page.click("#enCommitBtn");
  await sleep(400);
  const shown = await cellText(2, 1);
  ok("4. podgląd = komórka", preview === `Zapisze się jako: ${shown}` && shown !== "", { preview, shown });
  ok("4. status zapisany", (await cellText(2, 2)) === "zdana", await cellText(2, 2));
  ok("4. po zapisie karta czysta, krok 1", (await st()).step === 1, await st());

  // 6. nadpisanie
  await typeTarget("anna nowak");
  await page.press("#enTargetInput", "Enter"); // jeden dokładny kandydat — podświetlony
  await sleep(200);
  ok("6. jeden dokładny kandydat → krok 2", (await st()).step === 2, await st());
  await page.fill(field(3), "6");
  await page.keyboard.press("ControlOrMeta+Enter");
  await sleep(200);
  ok("6. Wpisz zablokowane bez zgody", await page.$eval("#enCommitBtn", (b) => b.disabled), null);
  await page.check('#entryOverlay [data-en-overwrite="3"]');
  await sleep(100);
  ok("6. po zgodzie Wpisz aktywne", await page.$eval("#enCommitBtn", (b) => !b.disabled), null);
  await page.click("#enCommitBtn");
  await sleep(400);
  ok("6. D4 = 6", (await cellText(3, 3)) === "6", await cellText(3, 3));
  const toast6 = await page.evaluate(() => [...document.querySelectorAll(".toast")].map((t) => ({ msg: t.querySelector(".toast-msg")?.textContent, btns: [...t.querySelectorAll(".toast-action")].map((b) => b.textContent) })));
  const t6 = toast6.find((x) => /Wpisano do wiersza 4/.test(x.msg || ""));
  ok("6. dymek z Cofnij i Pokaż w tabeli", t6 && t6.btns.includes("Cofnij") && t6.btns.includes("Pokaż w tabeli"), toast6);
  await page.evaluate(() => [...document.querySelectorAll(".toast .toast-action")].find((b) => b.textContent === "Cofnij")?.click());
  await sleep(300);
  ok("6. Cofnij z dymku — D4 wraca do 1", (await cellText(3, 3)) === "1", await cellText(3, 3));

  // 7. przypinanie
  await typeTarget("Piotr Zając");
  await page.press("#enTargetInput", "Enter");
  await sleep(200);
  await page.fill(field(1), "20.03.2026");
  await page.click('#entryOverlay [data-en-col="1"] .en-pin');
  await page.keyboard.press("ControlOrMeta+Enter");
  await sleep(200);
  await page.check('#entryOverlay [data-en-overwrite="1"]');
  await page.click("#enCommitBtn");
  await sleep(400);
  const s7 = await st();
  ok("7. przypięta kolumna 1 zapamiętana", s7.pinned.includes(1) && s7.fields[1] === "20.03.2026", s7);
  await typeTarget("Olga Wrona");
  await page.click("#enCandidates [data-en-new]");
  await sleep(300);
  ok("7. nowa kartka — przypięta data w polu", (await page.inputValue(field(1))) === "20.03.2026", await page.inputValue(field(1)));

  // 8. klawisze nie przeciekają, Esc zamyka
  const before8 = await page.evaluate(() => JSON.stringify(focusedCellState));
  await page.focus(field(0));
  await page.keyboard.press("ArrowDown");
  await page.keyboard.press("Delete");
  const after8 = await page.evaluate(() => JSON.stringify(focusedCellState));
  ok("8. strzałki w karcie nie ruszają tabeli", before8 === after8, { before8, after8 });
  await page.keyboard.type("x"); // pole z treścią → miękkie podpowiedzi mogą się pokazać
  await sleep(150);
  const sugOpen = await page.evaluate(() => !!document.querySelector(".cell-suggest:not(.hidden)"));
  if (sugOpen) {
    await page.keyboard.press("Escape");
    await sleep(150);
    ok("8. 1. Esc zamyka podpowiedzi, karta zostaje", (await st()).open && !(await page.evaluate(() => !!document.querySelector(".cell-suggest:not(.hidden)"))), await st());
  }
  await page.keyboard.press("Escape");
  await sleep(250);
  ok("8. Esc zamyka", !(await st()).open, await st());

  // 9. Podgląd i filtr
  await page.evaluate(() => sheetEditUi.setMode("view"));
  await openEntry();
  await sleep(250);
  const t9 = await page.evaluate(() => [...document.querySelectorAll(".toast .toast-msg")].map((e) => e.textContent));
  ok("9. Podgląd — odmowa z tekstem", t9.includes("Wpisywanie działa w trybie Edycja. Przełącz na Edycja i spróbuj ponownie.") && !(await st()).open, t9);
  await page.evaluate(() => sheetEditUi.setMode("edit"));
  await page.evaluate(() => { quickSearchEl.value = "Nowak"; applyQuickSearch(); window.__appDialogLog = []; window.__appDialogMode = "accept"; });
  await sleep(400);
  await openEntry();
  await sleep(500);
  const log9 = await page.evaluate(() => window.__appDialogLog);
  ok("9. filtr — okno z Wyczyść filtr i otwórz", log9.length === 1 && log9[0].buttons.includes("Wyczyść filtr i otwórz"), log9);
  ok("9. po akceptacji filtr wyczyszczony i nakładka otwarta", (await st()).open && (await page.evaluate(() => !swbRowGrow.filtered())), await st());
  await page.keyboard.press("Escape");

  // ── Task 4: pamięć ─────────────────────────────────────────────────────────
  const reloadFile = async () => {
    await page.reload({ waitUntil: "load" });
    await page.evaluate(() => { document.getElementById("heroSplash")?.remove(); return ensureXlsxLibs(false); });
    await page.setInputFiles("#fileInput", FILE);
    await page.waitForFunction(() => currentHeaderRow === 1 && !!document.querySelector("#dataTable tbody tr[data-row-key]"), null, { timeout: 20000 });
    await page.evaluate(() => { sheetEditUi.setMode("edit"); window.__appDialogMode = "accept"; });
    await sleep(300);
  };
  const toggleSettings = async () => { await page.click("#enSettingsBtn"); await sleep(150); };

  // 11. ustawienia arkusza: klucz, kolejność i wybór pól
  await openEntry(); await sleep(250);
  await toggleSettings();
  ok("11. panel ustawień widoczny", await page.isVisible("#enSettings"), null);
  await page.selectOption("#enKeyCol", "2");
  await page.click('#enSettings [data-en-fieldrow="2"] [data-en-move="up"]');
  await page.uncheck('#enSettings [data-en-fieldrow="3"] [data-en-field-on]');
  await sleep(100);
  const s11 = await st();
  ok("11. klucz = 2, kolejność 0,2,1…, pole 3 wyłączone", s11.keyCol === 2 && s11.fieldOrder.slice(0, 3).join() === "0,2,1" && !s11.selected.includes(3), s11);
  await page.keyboard.press("Escape"); await sleep(150);
  await reloadFile();
  await openEntry(); await sleep(250);
  const s11b = await st();
  ok("11. ustawienia po przeładowaniu", s11b.keyCol === 2 && s11b.fieldOrder.slice(0, 3).join() === "0,2,1" && !s11b.selected.includes(3), s11b);
  await toggleSettings();
  await page.selectOption("#enKeyCol", "0");
  await page.check('#enSettings [data-en-fieldrow="3"] [data-en-field-on]');
  await toggleSettings();

  // 12. opcje silnika
  await toggleSettings();
  for (const o of ["fuzzyTarget", "typeHints", "typoHints", "duplicateGuard"]) await page.uncheck(`#enSettings [data-en-opt="${o}"]`);
  await toggleSettings();
  await typeTarget("Jan Kowalsky");
  const c12 = await page.$$eval("#enCandidates [data-en-row]", (els) => els.length);
  ok("12. fuzzy wyłączone — literówka nie znajduje", c12 === 0, c12);
  await page.evaluate(() => { window.__appDialogLog = []; });
  await page.click("#enCandidates [data-en-new]"); await sleep(250);
  ok("12. strażnik wyłączony — bez okna, krok 2", (await page.evaluate(() => window.__appDialogLog.length)) === 0 && (await st()).step === 2, await st());
  await page.fill(field(1), "12.03"); await sleep(150);
  ok("12. typeHints/typoHints wyłączone — brak podglądu i ostrzeżeń", (await page.$$eval('#entryOverlay [data-en-col="1"] .en-will, #entryOverlay [data-en-col="1"] .en-warn', (e) => e.length)) === 0, null);
  await page.fill(field(1), "");
  await page.click("#enChangeTarget"); await sleep(150);
  await toggleSettings();
  await page.uncheck('#enSettings [data-en-opt="confirmStep"]');
  await toggleSettings();
  await typeTarget("Zofia Kruk");
  await page.click("#enCandidates [data-en-new]"); await sleep(250);
  await page.keyboard.press("ControlOrMeta+Enter"); await sleep(400);
  const kruk = await page.evaluate(() => baseRows.some((r) => r.values[0] === "Zofia Kruk"));
  ok("12. confirmStep wyłączone — zapis od razu", kruk && (await st()).step === 1, { kruk, st: await st() });
  await typeTarget("anna nowak"); await page.press("#enTargetInput", "Enter"); await sleep(200);
  await page.fill(field(3), "9");
  await page.keyboard.press("ControlOrMeta+Enter"); await sleep(250);
  ok("12. confirmStep wyłączone, ale nadpisanie pokazuje krok 3", (await st()).step === 3, await st());
  await page.click("#enBackBtn"); await sleep(100);
  await page.fill(field(3), "");
  await page.click("#enChangeTarget"); await sleep(100);
  await page.keyboard.press("Escape"); await sleep(150);
  await reloadFile();
  await openEntry(); await sleep(250);
  const o12 = (await st()).options;
  ok("12. opcje po przeładowaniu", !o12.fuzzyTarget && !o12.typeHints && !o12.typoHints && !o12.duplicateGuard && !o12.confirmStep && o12.suggestions, o12);
  await toggleSettings();
  for (const o of ["fuzzyTarget", "typeHints", "typoHints", "duplicateGuard", "confirmStep"]) await page.check(`#enSettings [data-en-opt="${o}"]`);
  await toggleSettings();

  // 13. niedokończona kartka
  await typeTarget("Piotr Zając"); await page.press("#enTargetInput", "Enter"); await sleep(200);
  await page.fill(field(2), "wzięta");
  await page.fill(field(3), "8");
  await page.evaluate(() => {
    Object.defineProperty(document, "visibilityState", { configurable: true, get: () => "hidden" });
    document.dispatchEvent(new Event("visibilitychange"));
  });
  const d13 = await page.evaluate(() => { const x = JSON.parse(localStorage.getItem("excel-workbench-entry") || "{}"); return Object.values(x.scopes || {}).map((r) => r.draft).find(Boolean) || null; });
  ok("13. visibilitychange hidden — kartka zapisana od razu", d13 && d13.fields && d13.fields[3] === "8" && d13.target.rowIndex0 === 5, d13);
  await reloadFile();
  await openEntry(); await sleep(300);
  const n13 = await page.isVisible("#enDraft") ? await page.textContent("#enDraft") : "";
  ok("13. komunikat o niedokończonej kartce", /Piotr Zając/.test(n13) && /2/.test(n13), n13);
  await page.click("#enDraftResume"); await sleep(250);
  const s13 = await st();
  ok("13. Wróć do niej — cel i pola", s13.step === 2 && s13.target.rowIndex0 === 5 && s13.fields[2] === "wzięta" && (await page.inputValue(field(3))) === "8", s13);
  await page.click("#enChangeTarget"); await sleep(100);
  await page.keyboard.press("Escape"); await sleep(150);
  await openEntry(); await sleep(250);
  if (await page.isVisible("#enDraft")) await page.click("#enDraftDiscard");
  await sleep(150);
  const d13b = await page.evaluate(() => { const x = JSON.parse(localStorage.getItem("excel-workbench-entry") || "{}"); return Object.values(x.scopes || {}).some((r) => r.draft); });
  ok("13. Odrzuć — kartka usunięta", !d13b && !(await page.isVisible("#enDraft")), d13b);

  // 14. limit zakresów
  await page.evaluate(() => {
    const x = JSON.parse(localStorage.getItem("excel-workbench-entry") || "{}");
    x.scopes = x.scopes || {};
    for (let i = 0; i < 31; i++) x.scopes[`stary${i}.xlsx\tA`] = { at: i + 1, keyCol: 0 };
    localStorage.setItem("excel-workbench-entry", JSON.stringify(x));
  });
  await toggleSettings(); await page.selectOption("#enKeyCol", "0"); await toggleSettings();
  const n14 = await page.evaluate(() => { const x = JSON.parse(localStorage.getItem("excel-workbench-entry")); return { n: Object.keys(x.scopes).length, hasOldest: "stary0.xlsx\tA" in x.scopes }; });
  ok("14. max 30 zakresów, najstarszy usunięty", n14.n === 30 && !n14.hasOldest, n14);

  // 15. dziennik sesji i przypomnienie o zapisie
  await typeTarget("anna nowak"); await page.press("#enTargetInput", "Enter"); await sleep(200);
  await page.fill(field(1), "05.05.2026"); await page.keyboard.press("ControlOrMeta+Enter"); await sleep(200);
  await page.check('#entryOverlay [data-en-overwrite="1"]'); await page.click("#enCommitBtn"); await sleep(400);
  const log15 = await page.textContent("#enLogBtn");
  ok("15. licznik wpisanych w sesji", /1/.test(log15), log15);
  await page.click("#enLogBtn"); await sleep(150);
  await page.evaluate(() => document.querySelectorAll(".toast").forEach((t) => t.remove()));
  await page.click('#enLog [data-en-log-row="3"]'); await sleep(300);
  const f15 = await page.evaluate(() => focusedCellState && focusedCellState.rowKey);
  ok("15. Pokaż w tabeli — nakładka zamknięta, kursor w wierszu", !(await st()).open && f15 === "wide:3", { f15 });
  const t15 = await page.evaluate(() => [...document.querySelectorAll(".toast .toast-msg")].map((e) => e.textContent));
  ok("15. przypomnienie o zapisie przy zamknięciu", t15.some((m) => /nie w pliku/.test(m)), t15);
  await page.evaluate(() => document.querySelectorAll(".toast").forEach((t) => t.remove()));
  await openEntry(); await sleep(200); await page.keyboard.press("Escape"); await sleep(200);
  const t15b = await page.evaluate(() => [...document.querySelectorAll(".toast .toast-msg")].map((e) => e.textContent));
  ok("15. przypomnienie tylko raz", !t15b.some((m) => /nie w pliku/.test(m)), t15b);

  // 16. operatory szukania w polu celu (jak szybkie szukanie)
  await openEntry(); await sleep(250);
  await typeTarget("Ewa Lis && Status:=wzięta");
  const rows16 = await page.$$eval("#enCandidates [data-en-row]", (els) => els.map((e) => ({ r: e.dataset.enRow, sel: e.getAttribute("aria-selected") })));
  ok("16. zapytanie — jeden wiersz (Ewa Lis z wziętą pracą), od razu wybrany", rows16.length === 1 && rows16[0].r === "4" && rows16[0].sel === "true", rows16);
  const hint16 = await page.textContent("#enQueryHint").catch(() => "");
  ok("16. widać, że to zapytanie z operatorami", /operator/i.test(hint16) && /1/.test(hint16), hint16);
  ok("16. bez mylącego „Szukam w kolumnie: …”", !(await page.isVisible("#enKeyHint")), null);
  const new16 = await page.textContent("#enCandidates [data-en-new]");
  ok("16. „Nowy wiersz” bez wklejania zapytania jako nazwy", new16.trim() === "Nowy wiersz", new16);
  await page.press("#enTargetInput", "Enter"); await sleep(200);
  const s16 = await st();
  ok("16. Enter → krok 2, wiersz arkusza 5", s16.step === 2 && s16.target.rowIndex0 === 4 && s16.fields[0] === undefined, s16);
  await page.click("#enChangeTarget"); await sleep(150);
  await typeTarget("Osoba:Nowak && Status:=nic");
  await page.press("#enTargetInput", "Enter"); await sleep(200);
  ok("16. zapytanie bez wyników — Enter nie tworzy nowego wiersza", (await st()).step === 1, await st());
  await toggleSettings();
  await page.uncheck('#enSettings [data-en-opt="queryOperators"]');
  await toggleSettings();
  await typeTarget("Ewa Lis && Status:=wzięta");
  ok("16. operatory wyłączone — zwykłe szukanie (brak trafień)", (await page.$$eval("#enCandidates [data-en-row]", (e) => e.length)) === 0, null);
  await toggleSettings();
  await page.check('#enSettings [data-en-opt="queryOperators"]');
  await toggleSettings();

  // 17. notatki przy polach
  await typeTarget("jan kowalski"); await page.press("#enTargetInput", "Enter"); await sleep(200);
  if (await page.getAttribute('#entryOverlay [data-en-col="1"] .en-pin', "aria-pressed") === "true") await page.click('#entryOverlay [data-en-col="1"] .en-pin');
  await page.fill(field(1), "");
  await page.click('#entryOverlay [data-en-col="2"] .en-note-btn'); await sleep(100);
  ok("17. przycisk Notatka rozwija pole", await page.isVisible("#enN2"), null);
  await page.fill("#enN2", "Zwrot do piątku");
  await page.keyboard.press("ControlOrMeta+Enter"); await sleep(200);
  const ch17 = await page.$$eval("#entryOverlay .en-change", (els) => els.map((e) => e.textContent.replace(/\s+/g, " ")));
  ok("17. podgląd pokazuje notatkę", ch17.some((x) => /Notatka/.test(x) && /Zwrot do piątku/.test(x)), ch17);
  await page.click("#enCommitBtn"); await sleep(400);
  const n17 = await page.evaluate(() => swbNoteOf(currentSheetName, "C2")?.text || null);
  ok("17. notatka zapisana w komórce C2", n17 === "Zwrot do piątku", n17);
  await typeTarget("jan kowalski"); await page.press("#enTargetInput", "Enter"); await sleep(200);
  ok("17. przycisk pokazuje, że notatka już jest", await page.$eval('#entryOverlay [data-en-col="2"] .en-note-btn', (b) => b.classList.contains("has-note")), null);
  await page.click('#entryOverlay [data-en-col="2"] .en-note-btn'); await sleep(100);
  ok("17. pole notatki z obecną treścią", (await page.inputValue("#enN2")) === "Zwrot do piątku", await page.inputValue("#enN2"));
  await page.fill("#enN2", "Zwrot w poniedziałek");
  await page.keyboard.press("ControlOrMeta+Enter"); await sleep(200);
  ok("17. zmiana notatki wymaga zgody", await page.$eval("#enCommitBtn", (b) => b.disabled), null);
  await page.check('#entryOverlay [data-en-overwrite-note="2"]'); await sleep(100);
  await page.click("#enCommitBtn"); await sleep(400);
  ok("17. notatka zmieniona", (await page.evaluate(() => swbNoteOf(currentSheetName, "C2")?.text)) === "Zwrot w poniedziałek", null);
  await page.keyboard.press("Escape"); await sleep(150);

  // 18. klucz liczbowy w formacie z zerami (zgłoszenie: „Nr.” = 32 pokazywane jako „032”, karta proponowała nowy wiersz)
  await openEntry(); await sleep(250);
  await page.evaluate(() => {
    // komórka trzyma liczbę, a pokazuje ją w formacie 000 — tak jak plik Mateusza
    baseRows.forEach((r) => { if (r.rowIndex0 >= 1 && Array.isArray(r.display) && typeof r.values[3] === "number") r.display[3] = String(r.values[3]).padStart(3, "0"); });
    bumpSheetDataStamp();
  });
  await toggleSettings(); await page.selectOption("#enKeyCol", "3"); await toggleSettings();
  const want = await page.evaluate(() => { const r = baseRows.find((x) => x.rowIndex0 === 5); return { shown: r.display[3], raw: r.values[3] }; });
  await typeTarget(want.shown);
  const rows18 = await page.$$eval("#enCandidates [data-en-row]", (els) => els.map((e) => e.dataset.enRow));
  ok("18. „0NN” znajduje wiersz z liczbą NN w formacie 000", rows18.includes("5"), { want, rows18 });
  const newText = await page.textContent("#enCandidates [data-en-new]").catch(() => "");
  await page.evaluate(() => { window.__appDialogLog = []; window.__appDialogMode = "dismiss"; });
  await page.click("#enCandidates [data-en-new]"); await sleep(250);
  ok("18. „Nowy wiersz” z istniejącym numerem pyta o duplikat", (await page.evaluate(() => window.__appDialogLog.length)) === 1, newText);
  await page.evaluate(() => { window.__appDialogMode = "accept"; });
  await toggleSettings(); await page.selectOption("#enKeyCol", "0"); await toggleSettings();
  await page.keyboard.press("Escape"); await sleep(150);

  // 19. zgłoszenie „032||”: wiszący operator nie przełącza w zapytanie; przy zapytaniu widać, DLACZEGO wiersz pasuje
  await openEntry(); await sleep(250);
  await typeTarget("Ewa Lis||");
  ok("19. „Ewa Lis||” — zwykłe szukanie (bez trybu zapytania)", !(await page.isVisible("#enQueryHint")), null);
  ok("19. „Ewa Lis||” — obie Ewy Lis", (await page.$$eval("#enCandidates [data-en-row]", (e) => e.map((x) => x.dataset.enRow))).join() === "2,4", null);
  await typeTarget("Status:=wzięta");
  const why19 = await page.$$eval("#enCandidates [data-en-row] .en-cand-why", (e) => e.map((x) => x.textContent));
  ok("19. zapytanie — przy każdym wierszu kolumna i wartość, która pasuje", why19.length >= 2 && why19.every((w) => /Status/.test(w) && /wzięta/.test(w)), why19);
  await typeTarget("Osoba:Lis");
  const why19b = await page.$$eval("#enCandidates [data-en-row] .en-cand-why", (e) => e.length);
  ok("19. trafienie w samej kolumnie-kluczu — bez powtarzania klucza", why19b === 0, why19b);
  await page.keyboard.press("Escape"); await sleep(150);

  // 20. KOLEJKA: lista w polu celu → ekran kolejki → „wszystkim naraz” / „po kolei”
  const cv = (ref) => page.evaluate((ref) => { const x = workbook.Sheets[currentSheetName][ref]; return x ? x.v : null; }, ref);
  const toastMsgs = () => page.evaluate(() => [...document.querySelectorAll(".toast .toast-msg")].map((e) => e.textContent));
  const clearToasts = () => page.evaluate(() => document.querySelectorAll(".toast").forEach((t) => t.remove()));
  await openEntry(); await sleep(250);
  if (await page.isVisible("#enQueueDraftDiscard")) await page.click("#enQueueDraftDiscard");
  await typeTarget("Ewa Lis, Jan Kowalski, Zenon Nowy, Anna Nowak");
  const qopt = await page.textContent("#enCandidates [data-en-queue]").catch(() => "");
  ok("20. lista rozpoznana — propozycja kolejki z liczbą pozycji", /4/.test(qopt), qopt);
  await page.press("#enTargetInput", "Enter"); await sleep(250);
  const s20 = await st();
  ok("20. Enter → ekran kolejki", s20.step === "queue" && s20.queue && s20.queue.items.length === 4, s20.queue);
  const stat = s20.queue ? s20.queue.items.map((x) => x.status) : [];
  ok("20. statusy: kilka wierszy / ok / brak / ok", stat.join() === "ambiguous,ok,missing,ok", stat);
  ok("20. przyciski zablokowane, dopóki „Ewa Lis” bez wyboru", await page.$eval("#enQWalk", (b) => b.disabled) && await page.$eval("#enQAll", (b) => b.disabled), null);
  ok("20. kolumna-klucz (Osoba) NIE jest polem wspólnym — nadałaby wszystkim ten sam klucz", !(await page.$('#entryOverlay [data-en-qf="0"]')), null);
  await page.selectOption('#entryOverlay [data-en-q-choose="0"]', "4"); await sleep(100);
  ok("20. po wyborze — „Po kolei” aktywne, „naraz” czeka na pole wspólne", !(await page.$eval("#enQWalk", (b) => b.disabled)) && await page.$eval("#enQAll", (b) => b.disabled), null);
  ok("20. wyszarzone „naraz” mówi dlaczego", /pole wspólne/.test(await page.textContent("#enQHint")) && await page.isVisible("#enQHint"), await page.textContent("#enQHint").catch(() => ""));
  await page.fill('#entryOverlay [data-en-qf="2"]', "zdana"); await sleep(100);
  ok("20. pole wspólne wypełnione — „naraz” dla 3 wierszy (bez brakującego)", /3/.test(await page.textContent("#enQAll")) && !(await page.$eval("#enQAll", (b) => b.disabled)), await page.textContent("#enQAll"));
  const before20 = await cv("C5");
  await page.click("#enQAll"); await sleep(250);
  ok("20. podgląd serii", (await st()).step === "queueReview" && (await page.$$eval("#entryOverlay .en-q-row", (e) => e.length)) >= 1, await st());
  if (await page.isVisible("#enQOverwrite")) {
    ok("20. nadpisania — „Wpisz” zablokowane do zgody", await page.$eval("#enCommitBtn", (b) => b.disabled), null);
    await page.check("#enQOverwrite"); await sleep(100);
  }
  await page.click("#enCommitBtn"); await sleep(500);
  ok("20. naraz — C2, C4, C5 = zdana", [await cv("C2"), await cv("C4"), await cv("C5")].every((v) => v === "zdana"), [await cv("C2"), await cv("C4"), await cv("C5")]);
  const t20 = await toastMsgs();
  ok("20. dymek mówi prawdę: 2 wiersze zmienione, 1 już miał tę wartość", t20.some((m) => /2 wierszy/.test(m) && /1 bez zmian/.test(m)), t20);
  await page.evaluate(() => [...document.querySelectorAll(".toast .toast-action")].find((b) => b.textContent === "Cofnij")?.click());
  await sleep(300);
  ok("20. Cofnij z dymku cofa całą serię", (await cv("C5")) === before20, { c5: await cv("C5"), before20 });
  ok("20. po serii — karta wraca do pola celu", (await st()).step === 1 && !(await st()).queue, await st());
  await clearToasts();

  // po kolei
  await typeTarget("Jan Kowalski, Anna Nowak"); await page.press("#enTargetInput", "Enter"); await sleep(250);
  await page.fill('#entryOverlay [data-en-qf="1"]', "01.12.2026");
  await page.click("#enQWalk"); await sleep(250);
  const w1 = await st();
  ok("20. po kolei — 1. karta: Jan (wiersz 2), data wspólna wpisana", w1.step === 2 && w1.target.rowIndex0 === 1 && w1.fields[1] === "01.12.2026", w1);
  ok("20. pasek postępu „1 z 2”", /1 z 2/.test(await page.textContent("#enQProgress").catch(() => "")), await page.textContent("#enQProgress").catch(() => ""));
  await page.fill(field(3), "7");
  await page.keyboard.press("ControlOrMeta+Enter"); await sleep(200);
  for (const cb of await page.$$("#entryOverlay [data-en-overwrite]")) await cb.check();
  await sleep(100);
  await page.click("#enCommitBtn"); await sleep(500);
  const w2 = await st();
  ok("20. „Wpisz i dalej” — od razu 2. karta: Anna (wiersz 4) z datą wspólną", w2.step === 2 && w2.target && w2.target.rowIndex0 === 3 && w2.fields[1] === "01.12.2026", w2);
  ok("20. zapisane w 1. karcie: D2 = 7", (await cv("D2")) === 7, await cv("D2"));
  await clearToasts();
  await page.click("#enQSkip"); await sleep(300);
  const t20b = await toastMsgs();
  ok("20. koniec kolejki — podsumowanie wpisano 1 · pominięto 1", t20b.some((m) => /1/.test(m) && /pomini/i.test(m)), t20b);
  ok("20. po kolejce — pole celu", (await st()).step === 1 && !(await st()).queue, await st());

  // brak w arkuszu → nowy wiersz w trybie po kolei
  await typeTarget("Zenon Nowy, Jan Kowalski"); await page.press("#enTargetInput", "Enter"); await sleep(250);
  await page.selectOption('#entryOverlay [data-en-q-missing="0"]', "new"); await sleep(100);
  await page.click("#enQWalk"); await sleep(250);
  const wq3 = await st();
  ok("20. pozycja spoza arkusza → karta nowego wiersza z kluczem", wq3.step === 2 && wq3.target && wq3.target.newRow && wq3.fields[0] === "Zenon Nowy", wq3);
  await page.click("#enQBack"); await sleep(200);
  ok("20. „Wróć do kolejki”", (await st()).step === "queue", await st());

  // przywrócenie po przeładowaniu
  await page.keyboard.press("Escape"); await sleep(150);
  await reloadFile();
  await openEntry(); await sleep(300);
  ok("20. po przeładowaniu — komunikat o niedokończonej kolejce", await page.isVisible("#enQueueDraft"), null);
  if (await page.isVisible("#enQueueDraftResume")) await page.click("#enQueueDraftResume");
  await sleep(250);
  ok("20. Wróć — ekran kolejki z 2 pozycjami", (await st()).step === "queue" && (await st()).queue && (await st()).queue.items.length === 2, await st());
  if (await page.isVisible("#enQCancel")) await page.click("#enQCancel");
  await sleep(200);
  ok("20. Zakończ kolejkę — pole celu", (await st()).step === 1 && !(await st()).queue, await st());

  // ustawienia kolejki
  await toggleSettings();
  await page.uncheck('#enSettings [data-en-opt="queueLists"]');
  await toggleSettings();
  await typeTarget("Ewa Lis, Jan Kowalski");
  ok("20. listy wyłączone — brak propozycji kolejki", !(await page.isVisible("#enCandidates [data-en-queue]")), null);
  await toggleSettings();
  await page.check('#enSettings [data-en-opt="queueLists"]');
  await page.selectOption("#enQueueMissing", "new");
  await toggleSettings();
  await typeTarget("Zenon Nowy, Jan Kowalski"); await page.press("#enTargetInput", "Enter"); await sleep(250);
  ok("20. ustawienie „brak w arkuszu → nowy wiersz” jako domyślne", (await page.inputValue('#entryOverlay [data-en-q-missing="0"]').catch(() => "")) === "new", null);
  if (await page.isVisible("#enQCancel")) await page.click("#enQCancel");
  await sleep(150);
  await toggleSettings(); await page.selectOption("#enQueueMissing", "skip"); await toggleSettings();
  const fq = await page.evaluate(() => window.__entry.fillQueue(["Jan Kowalski", "Anna Nowak", "jan kowalski"]));
  await sleep(200);
  const sfq = await st();
  ok("20. fillQueue (pod OCR) — kolejka z 2 pozycjami, powtórka wyłapana", fq && sfq.step === "queue" && sfq.queue.items.length === 2, sfq.queue);
  if (await page.isVisible("#enQCancel")) await page.click("#enQCancel");
  await sleep(150);
  await page.keyboard.press("Escape"); await sleep(150);

  ok("10. brak pageerror", errors.length === 0, errors);
  await browser.close();
  const tag = `${ENGINE}${TOUCH ? " touch" : ""}`;
  if (failures.length) {
    console.log(`entry-ui (${tag}): ${failures.length} błędów`);
    failures.forEach((f) => console.log("  ✗", f));
    process.exit(1);
  }
  console.log(`entry-ui (${tag}): OK`);
}

run().catch((e) => { console.error(e); process.exit(1); });
