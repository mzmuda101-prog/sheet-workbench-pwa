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
  page.on("pageerror", (e) => errors.push(e.message));
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
