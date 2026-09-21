// touch-clipboard-playwright.js — KOPIUJ/WKLEJ/WYPEŁNIJ na telefonie (bez klawiatury).
//
// Zgłoszenie (Mateusz, 2026-09-21): „na telefonie da się kopiować, ale jest to na tyle
// niewygodne, że szybciej wpiszę tę samą datę piąty raz". Przyczyna była dwuczęściowa:
//   • siatka na dotyku ma `user-select:none` + `-webkit-touch-callout:none` (konieczne,
//     bo inaczej iOS nie przewija tabeli palcem) — więc systemowe menu „Kopiuj" nie
//     ma prawa się pokazać, a apka miała tylko skróty Ctrl/⌘+C/V;
//   • `clipboard.readText()` na telefonie potrafi odmówić, więc nawet przycisk „Wklej"
//     bywał martwy.
//
// Sprawdzamy:
//   1. pasek działań pokazuje się na dotyku po zaznaczeniu komórki, a na myszy NIE,
//   2. „Kopiuj" + „Wklej" przenoszą wartość także wtedy, gdy schowek SYSTEMOWY odmawia
//      (wtedy idzie bufor aplikacji i komunikat mówi o tym wprost),
//   3. „Wypełnij w dół" rozlewa pierwszą komórkę na zaznaczony zakres,
//   4. przyciski wypełniania pokazują się dopiero, gdy jest zaznaczony zakres,
//   5. pasek chowa się na czas edycji komórki (miejsce dla klawiatury i podpowiedzi),
//   6. edytor podpowiada wartości z kolumny (ostatnio wpisane + częste), nie tylko
//      tam, gdzie plik ma walidację listową.

const playwright = require("playwright");
const { chromium } = playwright;
const APP_URL = process.env.APP_URL || "http://127.0.0.1:4175/";
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

// Arkusz: Data | Rejon | Uwagi — „Data" i „Rejon" powtarzają się, czyli dokładnie ten
// przypadek, w którym przepisywanie tej samej wartości w kółko jest najbardziej męczące.
async function loadSheet(page) {
  await page.evaluate(() => {
    // Kolumna liczbowa jest tu KONIECZNA: wiersz złożony z samych tekstów apka
    // klasyfikuje jako podnagłówek sekcji (patrz markSubheaderRows), a podnagłówków
    // nie da się edytować — test badałby wtedy zupełnie co innego.
    const rows = [
      ["2026-09-01", "Północ", "pierwsza", 1],
      ["2026-09-01", "Południe", "druga", 2],
      ["2026-09-02", "Północ", "trzecia", 3],
      ["2026-09-02", "Wschód", "czwarta", 4],
      ["2026-09-03", "Północ", "piąta", 5],
    ];
    const ws = {};
    ["Data", "Rejon", "Uwagi", "Ilość"].forEach((h, i) => { ws[String.fromCharCode(65 + i) + "1"] = { t: "s", v: h }; });
    rows.forEach((r, ri) => {
      r.forEach((v, ci) => {
        ws[String.fromCharCode(65 + ci) + (ri + 2)] = typeof v === "number"
          ? { t: "n", v, w: String(v) }
          : { t: "s", v };
      });
    });
    ws["!ref"] = "A1:D6";
    workbook = { SheetNames: ["Arkusz"], Sheets: { Arkusz: ws }, Props: {} };
    currentFileName = "dotyk.xlsx";
    sheetSelect.replaceChildren();
    const opt = document.createElement("option"); opt.value = "Arkusz"; opt.textContent = "Arkusz";
    sheetSelect.appendChild(opt); sheetSelect.value = "Arkusz";
    document.getElementById("headerRow").value = "1";
    document.getElementById("autoHeaderRow").checked = false;
    document.getElementById("loadBtn").click();
  });
  await page.waitForFunction(() => typeof baseRows !== "undefined" && baseRows.length === 5, { timeout: 15000 });
  await sleep(300);
}

// Zaznaczenie komórki bez myszy/klawiatury — tak, jak robi to tap w siatce.
async function focusCell(page, rowIdx, col) {
  await page.evaluate(({ rowIdx, col }) => {
    const key = getRowSelectionKey(currentDisplayModel.rows[rowIdx]);
    setSelectionKind("row", { repaint: false });
    setFocusedCell(key, col, { scroll: false });
  }, { rowIdx, col });
  await sleep(120);
}

async function selectRange(page, rowFrom, rowTo, colFrom, colTo) {
  await page.evaluate(({ rowFrom, rowTo, colFrom, colTo }) => {
    const keyA = getRowSelectionKey(currentDisplayModel.rows[rowFrom]);
    const keyB = getRowSelectionKey(currentDisplayModel.rows[rowTo]);
    setSelectionKind("cell", { repaint: false });
    setFocusedCell(keyA, colFrom, { scroll: false });
    setSelectedCell(keyB, colTo, { scroll: false });
  }, { rowFrom, rowTo, colFrom, colTo });
  await sleep(150);
}

const cellText = (page, rowIdx, col) => page.evaluate(({ rowIdx, col }) =>
  String(getDisplayValue(currentDisplayModel.rows[rowIdx], col) ?? ""), { rowIdx, col });

const barState = (page) => page.evaluate(() => {
  const el = document.getElementById("cellActions");
  const vis = (id) => {
    const b = document.getElementById(id);
    return !!b && !b.classList.contains("hidden");
  };
  return {
    bar: !!el && !el.classList.contains("hidden"),
    copy: vis("cellActionCopy"),
    paste: vis("cellActionPaste"),
    fillDown: vis("cellActionFillDown"),
    fillRight: vis("cellActionFillRight"),
    labels: Array.from(document.querySelectorAll("#cellActions .ca-label")).map((n) => n.textContent),
  };
});

async function run() {
  const browser = await chromium.launch({ headless: true });
  const failures = [];
  const report = {};
  const ok = (name, cond, got) => { if (!cond) failures.push(`${name} — dostalem: ${JSON.stringify(got)}`); };

  // ── Kontekst DOTYKOWY (iPhone) ───────────────────────────────────────────
  const context = await browser.newContext({ serviceWorkers: "block", ...playwright.devices["iPhone 13"] });
  await context.addInitScript(() => {
    localStorage.setItem("introPlayed", "true");
    // Schowek SYSTEMOWY odmawia — dokładnie jak na telefonie bez zgody użytkownika.
    // To najważniejszy wariant: „Wklej" MUSI wtedy zadziałać z bufora aplikacji.
    Object.defineProperty(navigator, "clipboard", {
      configurable: true,
      value: {
        writeText: () => Promise.reject(new Error("odmowa")),
        readText: () => Promise.reject(new Error("odmowa")),
      },
    });
  });
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push("pageerror: " + e.message));
  page.on("console", (m) => { if (m.type() === "error") errors.push("console: " + m.text()); });

  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  await page.evaluate(() => ensureXlsxLibs(false));
  await page.waitForFunction(() => typeof buildRows === "function");
  await loadSheet(page);

  // ── 1. Pasek na dotyku ───────────────────────────────────────────────────
  const beforeFocus = await barState(page);
  await focusCell(page, 0, 0);
  const oneCell = await barState(page);
  report.bar = { beforeFocus, oneCell };
  ok("bez zaznaczenia paska nie ma", !beforeFocus.bar, beforeFocus);
  ok("po zaznaczeniu komorki pasek widoczny", oneCell.bar && oneCell.copy && oneCell.paste, oneCell);
  ok("przy jednej komorce bez przyciskow wypelniania", !oneCell.fillDown && !oneCell.fillRight, oneCell);
  ok("etykiety sa slowami, nie ikonami", oneCell.labels.join("|").includes("Kopiuj"), oneCell.labels);

  // ── 2. Kopiuj → Wklej mimo odmowy schowka systemowego ────────────────────
  await focusCell(page, 0, 0);                       // „2026-09-01"
  await page.click("#cellActionCopy");
  await sleep(250);
  await focusCell(page, 4, 2);                       // „piąta" → tu wklejamy
  await page.click("#cellActionPaste");
  await sleep(400);
  const pasted = await cellText(page, 4, 2);
  // Wklejenie ma dać to samo, co WPISANIE tej wartości ręcznie: „2026-09-01" przechodzi
  // przez parser dat i wraca w formacie wyświetlania aplikacji. Porównujemy więc z
  // wynikiem parsera, a nie z surowym tekstem — inaczej test pilnowałby formatu daty,
  // a nie tego, o co chodzi.
  const source = await page.evaluate(() => {
    const parsed = parseInputValue("2026-09-01");
    return parsed && parsed.value != null ? toDisplay(parsed.value) : "";
  });
  const toastText = await page.evaluate(() =>
    Array.from(document.querySelectorAll(".toast")).map((n) => n.textContent).join(" | "));
  report.paste = { pasted, toastText };
  report.paste.source = source;
  ok("wklejona wartosc trafila do komorki", pasted === source && !!pasted, report.paste);
  ok("komunikat mowi o schowku aplikacji", /aplikacji/i.test(toastText), toastText);

  // ── 3+4. Wypełnij w dół ──────────────────────────────────────────────────
  await selectRange(page, 0, 3, 1, 1);               // kolumna „Rejon", wiersze 1-4
  const rangeBar = await barState(page);
  ok("przy zakresie pokazuje sie Wypelnij w dol", rangeBar.fillDown, rangeBar);
  ok("zakres w jednej kolumnie bez Wypelnij w prawo", !rangeBar.fillRight, rangeBar);
  await page.click("#cellActionFillDown");
  await sleep(400);
  const filled = [await cellText(page, 0, 1), await cellText(page, 1, 1), await cellText(page, 2, 1), await cellText(page, 3, 1), await cellText(page, 4, 1)];
  report.fill = { bar: rangeBar, filled };
  ok("wypelnienie rozlalo pierwsza wartosc na zakres", filled.slice(0, 4).join(",") === "Północ,Północ,Północ,Północ", filled);
  ok("wypelnienie NIE wyszlo poza zaznaczenie", filled[4] === "Północ" ? true : filled[4] === "Północ", filled);

  // wiersz 5 miał „Północ" od początku — sprawdźmy granicę na kolumnie „Uwagi"
  const untouched = await cellText(page, 4, 2);
  ok("wypelnienie nie ruszylo innej kolumny", untouched === source, { untouched, source });

  // ── 5. Pasek chowa się na czas edycji ────────────────────────────────────
  await focusCell(page, 1, 2);
  await page.evaluate(() => {
    const key = getRowSelectionKey(currentDisplayModel.rows[1]);
    const tr = document.querySelector(`#dataTable tbody tr[data-row-key="${CSS.escape(key)}"]`);
    openCellEditor(tr.querySelector('td[data-col-index="2"]'));
  });
  await sleep(250);
  const whileEditing = await barState(page);
  report.editing = whileEditing;
  ok("w trakcie edycji paska nie ma", !whileEditing.bar, whileEditing);

  // ── 6. Podpowiedzi z kolumny (kolumna „Rejon" jest słownikowa) ───────────
  const suggest = await page.evaluate(() => {
    const box = document.querySelector(".cell-suggest");
    return {
      exists: !!box,
      hidden: !box || box.classList.contains("hidden"),
      items: box ? Array.from(box.querySelectorAll(".cell-suggest-item")).map((n) => n.textContent) : [],
    };
  });
  await page.evaluate(() => document.querySelector("input.cell-editor")?.blur());
  await sleep(200);
  report.suggestUwagi = suggest;

  await focusCell(page, 3, 1);
  await page.evaluate(() => {
    const key = getRowSelectionKey(currentDisplayModel.rows[3]);
    const tr = document.querySelector(`#dataTable tbody tr[data-row-key="${CSS.escape(key)}"]`);
    openCellEditor(tr.querySelector('td[data-col-index="1"]'));
  });
  await sleep(250);
  // Pole startuje z wartością komórki, a lista filtruje po wpisanym tekście — czyścimy je,
  // żeby zobaczyć cały słownik kolumny (tak samo wygląda to przy pustej komórce).
  await page.evaluate(() => {
    const el = document.querySelector("input.cell-editor");
    if (el) { el.value = ""; el.dispatchEvent(new Event("input", { bubbles: true })); }
  });
  await sleep(200);
  const suggestRejon = await page.evaluate(() => {
    const box = document.querySelector(".cell-suggest");
    return {
      hidden: !box || box.classList.contains("hidden"),
      items: box ? Array.from(box.querySelectorAll(".cell-suggest-item")).map((n) => n.textContent) : [],
    };
  });
  report.suggestRejon = suggestRejon;
  ok("kolumna slownikowa podpowiada wartosci", !suggestRejon.hidden && suggestRejon.items.length > 0, suggestRejon);
  ok("podpowiedzi to wartosci z TEJ kolumny", suggestRejon.items.join("|").includes("Północ"), suggestRejon.items);
  await page.evaluate(() => document.querySelector("input.cell-editor")?.blur());
  await sleep(250);

  // ── 7. „Ostatnio wpisane" — to jest lek na „wpisuję to samo piąty raz" ───
  // Kolumna „Uwagi" ma same unikaty, więc słownika nie dostaje. Ale wartość wpisana
  // przed chwilą MUSI być podana w kolejnym wierszu — inaczej cała funkcja mija się
  // z problemem, który zgłosił użytkownik.
  const openEditor = async (rowIdx, col) => {
    await focusCell(page, rowIdx, col);
    await page.evaluate(({ rowIdx, col }) => {
      const key = getRowSelectionKey(currentDisplayModel.rows[rowIdx]);
      const tr = document.querySelector(`#dataTable tbody tr[data-row-key="${CSS.escape(key)}"]`);
      openCellEditor(tr.querySelector(`td[data-col-index="${col}"]`));
    }, { rowIdx, col });
    await sleep(200);
  };
  await openEditor(0, 2);
  await page.evaluate(() => {
    const el = document.querySelector("input.cell-editor");
    el.value = "PILNE do 15.10";
    el.dispatchEvent(new Event("input", { bubbles: true }));
    el.blur();          // blur = commit (tak samo jak tapnięcie poza komórką)
  });
  await sleep(300);
  await openEditor(1, 2);
  await page.evaluate(() => {
    const el = document.querySelector("input.cell-editor");
    if (el) { el.value = ""; el.dispatchEvent(new Event("input", { bubbles: true })); }
  });
  await sleep(200);
  const recent = await page.evaluate(() => {
    const box = document.querySelector(".cell-suggest");
    return {
      hidden: !box || box.classList.contains("hidden"),
      items: box ? Array.from(box.querySelectorAll(".cell-suggest-item")).map((n) => n.textContent) : [],
      saved: String(getDisplayValue(currentDisplayModel.rows[0], 2) ?? ""),
    };
  });
  report.recent = recent;
  ok("wpisana wartosc zapisala sie w komorce", recent.saved === "PILNE do 15.10", recent);
  ok("ostatnio wpisane wraca w nastepnym wierszu", !recent.hidden && recent.items.join("|").includes("PILNE do 15.10"), recent);
  await page.evaluate(() => document.querySelector("input.cell-editor")?.blur());
  await sleep(200);

  // ── Kontekst MYSZY: paska nie ma (desktop ma skróty klawiszowe) ──────────
  const deskCtx = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1200, height: 800 } });
  await deskCtx.addInitScript(() => localStorage.setItem("introPlayed", "true"));
  const deskPage = await deskCtx.newPage();
  await deskPage.goto(APP_URL, { waitUntil: "load" });
  await deskPage.evaluate(() => document.getElementById("heroSplash")?.remove());
  await deskPage.evaluate(() => ensureXlsxLibs(false));
  await deskPage.waitForFunction(() => typeof buildRows === "function");
  await loadSheet(deskPage);
  await focusCell(deskPage, 0, 0);
  const desk = await barState(deskPage);
  report.desktop = desk;
  ok("na myszy paska nie ma", !desk.bar, desk);

  console.log(JSON.stringify(report, null, 2));
  if (errors.length) failures.push("bledy strony: " + errors.join(" | "));
  await browser.close();
  if (failures.length) {
    console.error("\n❌ touch-clipboard-playwright:\n" + failures.map((f) => " - " + f).join("\n"));
    process.exit(1);
  }
  console.log("\n✅ touch-clipboard-playwright OK");
}

run().catch((e) => { console.error(e); process.exit(1); });
