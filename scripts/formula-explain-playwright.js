// formula-explain-playwright.js — formuły słowami (formula-explain.js).
//
// Co musi być prawdą:
//   1. wyjaśnienie z drzewa formuły: JEŻELI / SUMA / LICZ.WARUNKI / WYSZUKAJ.PIONOWO po polsku,
//      z nazwą kolumny zamiast samego adresu,
//   2. w menu komórki z formułą jest „Wyjaśnij formułę”, w komórce bez formuły — nie ma,
//      a okno pokazuje zdanie, odwołania i wynik,
//   3. „✦ Opisz słowami” (atrapa AI): formuła z odpowiedzi modelu jest wyjęta z ``` i wstępu,
//      sprawdzona i wyjaśniona; „Wstaw do paska” tylko wpisuje ją do paska (komórka bez zmian);
//      odwołanie do samej komórki i brak formuły — odrzucone z komunikatem.
//
// Uruchom z serwerem na APP_URL (domyślnie http://127.0.0.1:4175/).

const { chromium, webkit } = require("playwright");
const path = require("path");

const APP_URL = process.env.APP_URL || "http://127.0.0.1:4175/";
const ENGINE = process.env.ENGINE === "webkit" ? webkit : chromium;
const FILE = path.join(__dirname, "RODO Obieg terenów 2026 V5.xlsx");
const SHEET = "2025-2026";
const results = [];
const check = (name, ok, detail = "") => results.push({ name, ok: !!ok, detail });
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

async function run() {
  const browser = await ENGINE.launch({ headless: true });
  const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1280, height: 900 } });
  await context.addInitScript(() => localStorage.setItem("introPlayed", "true"));
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push("pageerror: " + e.message));
  page.on("console", (m) => { if (m.type() === "error") errors.push("console: " + m.text()); });

  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  await page.setInputFiles("#fileInput", FILE);
  await page.waitForFunction(() => document.getElementById("sheetSelect")?.options?.length > 0, null, { timeout: 20000 });
  await page.evaluate((sheet) => {
    const select = document.getElementById("sheetSelect");
    if ([...select.options].some((o) => o.value === sheet)) select.value = sheet;
    select.dispatchEvent(new Event("change", { bubbles: true }));
  }, SHEET);
  await sleep(400);
  await page.evaluate(() => {
    const auto = document.getElementById("autoHeaderRow");
    if (auto) { auto.checked = true; auto.dispatchEvent(new Event("change", { bubbles: true })); }
    if (typeof applyAutoHeaderRowIfEnabled === "function") applyAutoHeaderRowIfEnabled();
  });
  await sleep(400);
  await page.evaluate(() => loadBtn.click());
  await page.waitForFunction(() => typeof baseRows !== "undefined" && baseRows.length > 0, null, { timeout: 20000 });
  await sleep(1500);

  // 1) czysta funkcja na typowych formułach
  const pure = await page.evaluate(() => {
    const h = (n) => currentHeaders[n] || "";
    const col = (n) => FormulaEngine.indexToCol(currentStartCol + n);
    return {
      header0: h(0),
      ifSum: swbExplainFormula(`=IF(${col(0)}5>0,SUM(B2:B10)/${col(0)}5,"-")`).sentence,
      countifs: swbExplainFormula('=LICZ.WARUNKI(A:A;"W trakcie";B:B;">60")').sentence,
      vlookup: swbExplainFormula("=VLOOKUP(A2,Tabela!$A$2:$C$100,3,FALSE)").sentence,
      datedif: swbExplainFormula('=DATEDIF(C2,TODAY(),"D")').sentence,
      bad: swbExplainFormula("=SUM(A1:").ok,
    };
  });
  check("JEŻELI + SUMA słowami, nazwa kolumny zamiast adresu", /^Jeżeli „/.test(pure.ifSum) && pure.ifSum.includes(`„${pure.header0}”`) && /suma zakres B2:B10|suma kolumna/.test(pure.ifSum) && /w przeciwnym razie „-”/.test(pure.ifSum), pure.ifSum);
  check("LICZ.WARUNKI: „ile wierszy spełnia jednocześnie … równa się „W trakcie” … większe niż 60”", /ile wierszy spełnia jednocześnie/i.test(pure.countifs) && /równa się „W trakcie”/.test(pure.countifs) && /jest większe niż 60/.test(pure.countifs), pure.countifs);
  check("WYSZUKAJ.PIONOWO: „znajdź … dokładne dopasowanie”", /znajdź .* w pierwszej kolumnie zakres A2:C100 w arkuszu „Tabela” i weź wartość z 3\. kolumny \(dokładne dopasowanie\)/i.test(pure.vlookup), pure.vlookup);
  check("DATEDIF z DZIŚ: „liczba dni od … do dziś”", /liczba dni od .* do dziś/i.test(pure.datedif), pure.datedif);
  check("zła składnia → ok: false", pure.bad === false);

  // 2) prawdziwa komórka z formułą: menu komórki + okno
  const real = await page.evaluate(() => {
    const ws = workbook.Sheets[currentSheetName];
    let target = null;
    for (const row of baseRows) {
      for (let col = 0; col < currentHeaders.length && !target; col++) {
        const ref = XLSX.utils.encode_cell({ r: row.rowIndex0, c: currentStartCol + col });
        if (ws[ref]?.f) target = { row, col, ref };
      }
      if (target) break;
    }
    if (!target) return { error: "no formula cell" };
    const item = cellMenuExplainItem({ row: target.row, col: target.col });
    let plain = null;
    for (let col = 0; col < currentHeaders.length && plain == null; col++) {
      const ref = XLSX.utils.encode_cell({ r: target.row.rowIndex0, c: currentStartCol + col });
      if (!ws[ref]?.f) plain = col;
    }
    const noItem = cellMenuExplainItem({ row: target.row, col: plain });
    item?.run();
    const dlg = document.querySelector(".fx-explain-dialog[open]");
    const out = {
      ref: target.ref,
      label: item?.label,
      noItem: noItem === null,
      formula: dlg?.querySelector(".fx-explain-formula")?.textContent,
      sentence: dlg?.querySelector(".fx-explain-sentence")?.textContent,
      steps: [...(dlg?.querySelectorAll(".fx-explain-steps li") || [])].map((li) => li.textContent),
      result: dlg?.querySelector(".fx-explain-result")?.textContent,
      refs: dlg?.querySelectorAll(".fx-explain-refs li").length || 0,
    };
    dlg?.close();
    return out;
  });
  check("menu: „Wyjaśnij formułę” tylko przy komórce z formułą", real.label === "Wyjaśnij formułę" && real.noItem, JSON.stringify(real));
  check("tabela Excela + łańcuch JEŻELI: krok „Jeżeli „od” jest puste → puste”", /^Jeżeli „od” jest puste → puste$/.test(real.steps?.[0] || ""), real.sentence);
  check("okno: formuła, zdanie, odwołania i wynik", /^=/.test(real.formula || "") && (real.sentence || "").length > 10 && real.refs >= 1 && /^Wynik w tej komórce: /.test(real.result || ""), JSON.stringify(real));

  // 3) ✦ Opisz słowami (atrapa AI) w oknie fx
  const writer = await page.evaluate(async () => {
    const row = baseRows[2];
    setSelectionKind("cell", { repaint: false });
    setFocusedCell(getRowSelectionKey(row), currentHeaders.length - 1, { scroll: false });
    const target = swbFormulaTarget();
    const colC = FormulaEngine.indexToCol(currentStartCol);
    const before = workbook.Sheets[currentSheetName][target.ref]?.f || null;
    const fxBtn = document.getElementById("fxInsertBtn");
    const ask = async (answer) => {
      window.__localAiMock = () => answer;
      fxBtn.click();
      if (document.querySelector(".fx-dialog")?.hidden) fxBtn.click();
      const input = document.querySelector(".fx-dialog .fx-write-input");
      input.value = "ile dni minęło od daty";
      document.querySelector(".fx-dialog .fx-write .btn").click();
      for (let i = 0; i < 40 && !document.querySelector(".fx-dialog .fx-write-out .ai-badge"); i++) await new Promise((r) => setTimeout(r, 50));
      const out = document.querySelector(".fx-dialog .fx-write-out");
      return { formula: out?.querySelector(".fx-explain-formula")?.textContent, sentence: out?.querySelector(".fx-explain-sentence")?.textContent, warn: out?.querySelector(".lai-warn")?.textContent, insert: !!out?.querySelector(".fx-write-out > .btn") };
    };
    const ok = await ask(`Oto formuła:\n\`\`\`excel\n=DATEDIF(${colC}${row.rowIndex0 + 1},TODAY(),"D")\n\`\`\`\nGotowe.`);
    document.querySelector(".fx-dialog .fx-write-out > .btn")?.click();
    const fx = document.getElementById("formulaInput").value;
    const after = workbook.Sheets[currentSheetName][target.ref]?.f || null;
    document.getElementById("formulaCancel")?.click();
    const self = await ask(`=${target.ref}+1`);
    const none = await ask("Nie wiem, jak to policzyć.");
    delete window.__localAiMock;
    return { ok, fx, unchanged: before === after, self, none };
  });
  check("odpowiedź AI oczyszczona i zapisana po polsku (DATA.RÓŻNICA, DZIŚ, „;”)", /^=DATA\.RÓŻNICA\([A-Z]+\d+;DZIŚ\(\);"D"\)$/.test(writer.ok.formula || ""), JSON.stringify(writer.ok));
  check("formuła AI wyjaśniona tym samym generatorem", /liczba dni od .* do dziś/i.test(writer.ok.sentence || ""), writer.ok.sentence);
  check("„Wstaw do paska” → pasek formuły, komórka bez zmian", writer.fx === writer.ok.formula && writer.unchanged, JSON.stringify({ fx: writer.fx, unchanged: writer.unchanged }));
  check("odwołanie do własnej komórki → odrzucone", /własnej komórki/.test(writer.self.warn || "") && !writer.self.insert, JSON.stringify(writer.self));
  check("brak formuły w odpowiedzi → komunikat", /nie podał formuły/.test(writer.none.warn || ""), JSON.stringify(writer.none));

  check("brak błędów strony", !errors.length, errors.join(" | "));
  await browser.close();
  let failed = 0;
  results.forEach((r) => { console.log(`${r.ok ? "✅" : "❌"} ${r.name}${r.ok || !r.detail ? "" : ` (${String(r.detail).slice(0, 500)})`}`); if (!r.ok) failed++; });
  if (failed) process.exit(1);
  console.log(`\n✅ formula-explain: ${results.length}/${results.length} (${process.env.ENGINE || "chromium"})`);
}
run().catch((e) => { console.error(e); process.exit(1); });
