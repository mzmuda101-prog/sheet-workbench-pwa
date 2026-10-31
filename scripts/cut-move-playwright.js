// cut-move-playwright.js — WYTNIJ → WKLEJ = PRZENIESIENIE, jak w Excelu (2026-10-31).
//   node scripts/cut-move-playwright.js        ENGINE=webkit node scripts/cut-move-playwright.js
//
// Wyrocznia Excela: A1 „5” (pogrubione), C1 „=A1*2”; wytnij A1 → wklej w B3 ⇒ B3 = 5 pogrubione,
// A1 pusta, C1 „=B3*2”. Pilnuje (na przykładowym pliku, wiersze nad nagłówkiem + dane):
//   1. Ctrl/⌘+X: przerywana ramka na źródle, dane jeszcze bez zmian; Esc ją zdejmuje;
//   2. wklejenie: wartość i format w celu, źródło puste i bez formatu, formuła zależna przepięta;
//   3. jedno Cofnij przywraca wszystko (źródło, cel, formułę);
//   4. menu komórki ma „Wytnij” i „Wyczyść zawartość”; „Wyczyść zawartość” czyści zakres;
//   5. wycinanie w posortowanym/przefiltrowanym widoku (wiersze pliku nie po kolei) — odmowa z wyjaśnieniem.
// Uruchom z serwerem na APP_URL (domyślnie http://127.0.0.1:4175/).

const playwright = require("playwright");

const APP_URL = process.env.APP_URL || "http://127.0.0.1:4175/";
const ENGINE = process.env.ENGINE || "chromium";
const SLEEP_SCALE = Math.min(1, Math.max(0.1, Number(process.env.SLEEP_SCALE || 1)));
const sleep = (ms) => new Promise((r) => setTimeout(r, Math.round(ms * SLEEP_SCALE)));
const MOD = process.platform === "darwin" && ENGINE === "webkit" ? "Meta" : "Control";

async function run() {
  const browser = await playwright[ENGINE].launch({ headless: true });
  const fails = [];
  const check = (name, ok, detail) => {
    console.log(`${ok ? "✅" : "❌"} ${name}${ok ? "" : ` — ${JSON.stringify(detail)}`}`);
    if (!ok) fails.push(name);
  };
  const ctx = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1280, height: 900 } });
  await ctx.addInitScript(() => {
    if (sessionStorage.getItem("__cmInit")) return;
    sessionStorage.setItem("__cmInit", "1");
    localStorage.clear();
    localStorage.setItem("introPlayed", "true");
  });
  const p = await ctx.newPage();
  const errors = [];
  p.on("pageerror", (e) => errors.push(e.message));
  await p.goto(APP_URL, { waitUntil: "load" });
  await p.evaluate(() => { document.getElementById("heroSplash")?.remove(); return loadSampleFile(); });
  await p.waitForFunction(() => document.querySelector("#dataTable thead tr.pre-row td.pre-cell"), null, { timeout: 20000 });
  await p.waitForFunction(() => !document.documentElement.classList.contains("vt-scene"), null, { timeout: 10000 });
  await p.evaluate(() => sheetEditUi.setMode("edit"));
  await sleep(300);

  // przygotowanie (wiersze 3–4 nad nagłówkiem są puste): A3 = 5 pogrubione, A4 = A3*2
  await p.evaluate(() => {
    updateSheetCell(2, 0, { value: 5, type: "number" });
    cfmtApplyAt(swbRowForIndex0(2), 0, { b: true });
    swbFormulaCommitAt(3, 0, "=A3*2", { silent: true });
    bumpSheetDataStamp();
    renderActiveTable();
  });
  await sleep(300);
  const A = (r, c) => `#dataTable tr[data-row-key="pre:${r}"] td[data-col-index="${c}"]`;
  const st = () => p.evaluate(() => {
    const sh = workbook.Sheets[currentSheetName];
    const f = (ref) => sh[ref] ? { v: sh[ref].v ?? null, f: sh[ref].f || null } : null;
    return { A3: f("A3"), C3: f("C3"), A4: f("A4"), boldA3: !!cfmtPendingFmt(currentSheetName, "A3")?.b, boldC3: !!cfmtPendingFmt(currentSheetName, "C3")?.b, cut: document.querySelectorAll("td.cell-cut").length };
  });

  await p.click(A(2, 0));
  await p.keyboard.press(`${MOD}+x`);
  await sleep(300);
  const s1 = await st();
  check("1. Wytnij: przerywana ramka, dane bez zmian", s1.cut === 1 && s1.A3 && s1.A3.v === 5, s1);
  await p.keyboard.press("Escape");
  await sleep(200);
  check("1. Esc zdejmuje wycinanie", (await st()).cut === 0, await st());

  await p.keyboard.press(`${MOD}+x`);
  await sleep(200);
  await p.click(A(2, 2));
  await p.keyboard.press(`${MOD}+v`);
  await sleep(500);
  const s2 = await st();
  check("2. cel: wartość 5 i pogrubienie", s2.C3 && s2.C3.v === 5 && s2.boldC3, s2);
  check("2. źródło: puste i bez pogrubienia", (!s2.A3 || s2.A3.v == null) && !s2.boldA3, s2);
  check("2. formuła zależna przepięta =A3*2 → =C3*2", s2.A4 && /^C3\*2$/i.test(s2.A4.f || "") && s2.A4.v === 10, s2);
  check("2. ramka wycinania zniknęła", s2.cut === 0, s2);

  await p.evaluate(() => undoRedo("undo"));
  await sleep(400);
  const s3 = await st();
  check("3. jedno Cofnij przywraca źródło, cel i formułę", s3.A3 && s3.A3.v === 5 && s3.boldA3 && (!s3.C3 || s3.C3.v == null) && /^A3\*2$/i.test(s3.A4?.f || ""), s3);

  // 4. menu i „Wyczyść zawartość”
  await p.click(A(2, 0));
  await p.click(A(3, 0), { modifiers: ["Shift"] });
  await p.click(A(2, 0), { button: "right" });
  await sleep(300);
  const menu = await p.evaluate(() => [...document.querySelectorAll("#cellMenu .cell-menu-item")].map((b) => b.className.match(/cmi-[\w-]+/)?.[0]));
  check("4. menu zakresu: Wytnij, Kopiuj, Wyczyść zawartość", ["cmi-cut", "cmi-copy", "cmi-clear"].every((x) => menu.includes(x)), menu);
  await p.evaluate(() => document.querySelector("#cellMenu .cmi-clear").click());
  await sleep(300);
  const s4 = await st();
  check("4. Wyczyść zawartość czyści zakres, format zostaje", (!s4.A3 || s4.A3.v == null) && (!s4.A4 || !s4.A4.f) && s4.boldA3, s4);

  // 5. przefiltrowany widok z dziurą
  const refused = await p.evaluate(async () => {
    setPrimarySort("Zakres", "desc"); // sąsiednie wiersze widoku ≠ sąsiednie wiersze pliku
    sortRows();
    renderActiveTable();
    const a = viewRows[0];
    const b = viewRows[1];
    setSelectionKind("cell", { repaint: false });
    setFocusedCell(getRowSelectionKey(a), 1, { scroll: false });
    setSelectedCell(getRowSelectionKey(b), 1, { scroll: false });
    document.querySelectorAll(".toast").forEach((x) => x.remove());
    const ok = await cutSelectionToClipboard();
    return { ok, gap: b.rowIndex0 - a.rowIndex0, toast: [...document.querySelectorAll(".toast")].map((x) => x.textContent).join(" | ") };
  });
  check("5. widok po sortowaniu (wiersze pliku nie po kolei) — odmowa z wyjaśnieniem", refused.gap !== 1 && refused.ok === false && /sąsiednie wiersze/.test(refused.toast), refused);

  check("bez błędów strony", errors.length === 0, errors);
  await browser.close();
  if (fails.length) {
    console.error(`\n${fails.length} błędów (${ENGINE})`);
    process.exit(1);
  }
  console.log(`\nOK (${ENGINE})`);
}

run().catch((e) => { console.error(e); process.exit(1); });
