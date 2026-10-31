// file-colors-playwright.js — KOLORY Z PLIKU JAK W EXCELU (2026-10-31).
//   node scripts/file-colors-playwright.js        ENGINE=webkit node scripts/file-colors-playwright.js
//
// Próbka scripts/excel-theme-colors-2023.xlsx zrobiona w Excelu (motyw Office 2023+: akcent 1
// #156082…). Wyrocznia — RGB odczytane z Excela (AppleScript `color of interior object`):
//   A2 156082 · A3 F7C7AC · A4 12501A · A5 CAEDFB · A6 D86DCD · A7 275317 · A8 E8E8E8 · A9 B8D3EF · A10 FFFF00
// Pilnuje:
//   1. kolor wypełnienia = Excel (±2 na kanał — zaokrąglenia odcienia) i PEŁNE krycie (dawniej 28%);
//   2. paleta motywu z PLIKU, nie sztywna Office 2013 (akcent 1 = #156082, nie #4472C4);
//   3. automatyczny tekst na wypełnieniu czytelny (czarny na jasnym, biały na ciemnym);
//   4. zaznaczenie kładzie się NA kolor (półprzezroczyste), kolor komórki dalej widać;
//   5. panel-nakładka ma pełne tło (nie prześwituje tabela) i długa nazwa pliku nie rozpycha sekcji.
// Uruchom z serwerem na APP_URL (domyślnie http://127.0.0.1:4175/).

const path = require("path");
const playwright = require("playwright");

const APP_URL = process.env.APP_URL || "http://127.0.0.1:4175/";
const ENGINE = process.env.ENGINE || "chromium";
const SLEEP_SCALE = Math.min(1, Math.max(0.1, Number(process.env.SLEEP_SCALE || 1)));
const sleep = (ms) => new Promise((r) => setTimeout(r, Math.round(ms * SLEEP_SCALE)));
const ORACLE = { 2: "156082", 3: "F7C7AC", 4: "12501A", 5: "CAEDFB", 6: "D86DCD", 7: "275317", 8: "E8E8E8", 9: "B8D3EF", 10: "FFFF00" };

async function run() {
  const browser = await playwright[ENGINE].launch({ headless: true });
  const fails = [];
  const check = (name, ok, detail) => {
    console.log(`${ok ? "✅" : "❌"} ${name}${ok ? "" : ` — ${JSON.stringify(detail)}`}`);
    if (!ok) fails.push(name);
  };
  const ctx = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1280, height: 820 } });
  await ctx.addInitScript(() => localStorage.setItem("introPlayed", "true"));
  const p = await ctx.newPage();
  const errors = [];
  p.on("pageerror", (e) => errors.push(e.message));
  await p.goto(APP_URL, { waitUntil: "load" });
  await p.evaluate(() => document.getElementById("heroSplash")?.remove());
  await p.setInputFiles("#fileInput", path.join(__dirname, "excel-theme-colors-2023.xlsx"));
  await p.waitForFunction(() => typeof workbook !== "undefined" && workbook && document.querySelector("#dataTable tbody tr[data-row-index]"), null, { timeout: 30000 });
  await sleep(700);
  await p.evaluate(() => { setSidebarOpen(false); sheetEditUi.setMode("edit"); });
  await sleep(300);

  const cells = await p.evaluate(() => {
    const out = {};
    for (const tr of document.querySelectorAll("#dataTable tbody tr[data-row-index]")) {
      const td = tr.querySelector('td[data-col-index="0"]');
      const cs = getComputedStyle(td);
      out[+tr.dataset.rowIndex + 1] = { bg: cs.backgroundColor, color: cs.color };
    }
    return out;
  });
  const parse = (c) => (c.match(/[\d.]+/g) || []).map(Number);
  const bad = [];
  for (const [row, hex] of Object.entries(ORACLE)) {
    const c = cells[row];
    if (!c) { bad.push({ row, missing: true }); continue; }
    const [r, g, b, a = 1] = parse(c.bg);
    const want = [0, 2, 4].map((i) => parseInt(hex.slice(i, i + 2), 16));
    if (a < 0.99 || [r, g, b].some((v, i) => Math.abs(v - want[i]) > 2)) bad.push({ row, got: c.bg, want: hex });
  }
  check("1. wypełnienia = Excel (±2) i pełne krycie", bad.length === 0, bad);

  const pal = await p.evaluate(() => swbThemeColors(workbook).slice(0, 6));
  check("2. paleta motywu z pliku (Office 2023: dk2 #0E2841, akcent 1 #156082)", pal[3] === "#0E2841" && pal[4] === "#156082", pal);

  const lum = (c) => { const [r, g, b] = parse(c); return (0.299 * r + 0.587 * g + 0.114 * b) / 255; };
  const contrastOk = Object.keys(ORACLE).filter((r) => r !== "10").every((r) => Math.abs(lum(cells[r].bg) - lum(cells[r].color)) > 0.4);
  check("3. automatyczny tekst czytelny na każdym wypełnieniu", contrastOk, Object.fromEntries(Object.keys(ORACLE).map((r) => [r, cells[r] && cells[r].color])));

  const sel = await p.evaluate(() => {
    setSelectionKind("cell", { repaint: false });
    const k = (n) => getRowSelectionKey(viewRows.find((r) => r.rowIndex0 === n - 1));
    setFocusedCell(k(3), 0, { scroll: false });
    setSelectedCell(k(5), 1, { scroll: false });
    const td = document.querySelector('#dataTable tbody tr[data-row-index="3"] td[data-col-index="0"]');
    const cs = getComputedStyle(td);
    return { img: cs.backgroundImage, bg: cs.backgroundColor, inRange: td.classList.contains("cell-in-range") || td.classList.contains("cell-active") };
  });
  check("4. zaznaczenie na kolorze: półprzezroczysta warstwa + kolor komórki pod spodem", sel.inRange && /gradient/.test(sel.img), sel);

  // 5. panel na telefonie/tablecie: pełne tło + długa nazwa pliku
  await p.setViewportSize({ width: 1180, height: 820 });
  await p.evaluate(async () => {
    await ensureXlsxLibs(true);
    const buf = await buildSampleWorkbookArrayBuffer();
    hasUnsavedChanges = false;
    if (await handleFile(new File([buf], "zadania i przywileje na zebraniu - październik-listopad 2026 wersja ostateczna.xlsx", { type: XLSX_MIME }))) openDefaultSheet();
  });
  await sleep(900);
  const panel = await p.evaluate(() => {
    setSidebarOpen(true);
    document.querySelector("#loadBtn").closest("details").open = true;
    const sb = document.querySelector(".sidebar");
    const r = sb.getBoundingClientRect();
    const over = [...sb.querySelectorAll("*")].filter((el) => el.offsetParent && el.getBoundingClientRect().right > r.right + 1).map((el) => el.id || el.className).slice(0, 5);
    return { bg: getComputedStyle(sb).backgroundColor, over };
  });
  await sleep(300);
  const alpha = (parse(panel.bg)[3] ?? 1);
  check("5. panel-nakładka nie prześwituje (pełne tło)", alpha >= 0.999, panel.bg);
  check("5. długa nazwa pliku nie wypycha treści poza panel", panel.over.length === 0, panel.over);

  check("bez błędów strony", errors.length === 0, errors);
  await browser.close();
  if (fails.length) {
    console.error(`\n${fails.length} błędów (${ENGINE})`);
    process.exit(1);
  }
  console.log(`\nOK (${ENGINE})`);
}

run().catch((e) => { console.error(e); process.exit(1); });
