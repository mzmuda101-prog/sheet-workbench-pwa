// cell-menu-paste-playwright.js — „Wklej" w menu komórki (prawy klik / przytrzymanie).
//
// Co musi być prawdą:
//   1. bez niczego skopiowanego w apce menu NIE ma „Wklej",
//   2. po „Kopiuj" menu innej komórki ma „Wklej „wartość”" i wkleja dokładnie ją,
//   3. wklejenie = SAME WARTOŚCI: styl komórki docelowej zostaje, formuła ze schowka
//      się nie wkleja,
//   4. jedna wartość + zaznaczony zakres → „Wklej … do N komórek" wypełnia cały zakres,
//   5. skopiowany blok 2×2 wkleja się od lewego górnego rogu zaznaczenia jako 2×2,
//   6. każde wklejenie = JEDEN krok Cofnij,
//   7. zapis → ponowny odczyt: wklejone wartości w pliku, styl (s=) komórek bez zmian.
//
// Uruchom z serwerem na APP_URL (domyślnie http://127.0.0.1:4175/).

const { chromium } = require("playwright");
const path = require("path");

const APP_URL = process.env.APP_URL || "http://127.0.0.1:4175/";
const FILE = path.join(__dirname, "stress-test-workbench.xlsx");
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

async function run() {
  const browser = await chromium.launch({ headless: true });
  const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1280, height: 900 } });
  await context.addInitScript(() => localStorage.setItem("introPlayed", "true"));
  await context.grantPermissions(["clipboard-read", "clipboard-write"]);
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push("pageerror: " + e.message));
  page.on("console", (m) => { if (m.type() === "error") errors.push("console: " + m.text()); });

  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  await page.evaluate(() => ensureXlsxLibs(false));
  await page.setInputFiles("#fileInput", FILE);
  await page.waitForFunction(() => document.getElementById("sheetSelect")?.options?.length > 0, null, { timeout: 20000 });
  await page.click("#loadBtn");
  await page.waitForFunction(() => document.querySelectorAll("#dataTable tbody tr").length > 10, null, { timeout: 20000 });
  await sleep(1000);
  await page.evaluate(() => setSidebarOpen(false));
  await sleep(400);

  const results = [];
  const check = (name, ok, detail) => results.push({ name, ok: !!ok, detail: detail === undefined ? "" : JSON.stringify(detail).slice(0, 300) });

  const cellPos = (nth, col) => page.evaluate(([nth, col]) => {
    const tr = document.querySelectorAll("#dataTable tbody tr")[nth];
    const td = tr.querySelector(`td[data-col-index="${col}"]`);
    td.scrollIntoView({ block: "nearest", inline: "nearest" });
    const r = td.getBoundingClientRect();
    return { x: r.left + r.width / 2, y: r.top + r.height / 2, key: tr.dataset.rowKey };
  }, [nth, col]);
  const click = async (nth, col, shift = false) => {
    const p = await cellPos(nth, col);
    if (shift) await page.keyboard.down("Shift");
    await page.mouse.click(p.x, p.y);
    if (shift) await page.keyboard.up("Shift");
    await sleep(100);
  };
  const menu = async (nth, col) => {
    const p = await cellPos(nth, col);
    await page.mouse.click(p.x, p.y, { button: "right" });
    await sleep(150);
    return page.evaluate(() => [...document.querySelectorAll("#cellMenu .cell-menu-item")].map((b) => ({ id: [...b.classList].find((c) => c.startsWith("cmi-")), text: b.textContent })));
  };
  const clickItem = async (id) => { await page.click(`#cellMenu .cmi-${id}`); await sleep(250); };
  const val = (nth, col) => page.evaluate(([nth, col]) => {
    const key = document.querySelectorAll("#dataTable tbody tr")[nth].dataset.rowKey;
    const row = currentDisplayModel.rows.find((r) => getRowSelectionKey(r) === key);
    const v = row.values[col];
    return v instanceof Date ? "D:" + v.toISOString().slice(0, 10) : v;
  }, [nth, col]);
  const pending = () => page.evaluate(() => Object.keys(pendingEdits[currentSheetName] || {}).length);

  // kolumna tekstowa: pierwsza, w której pierwsze 8 wierszy ma RÓŻNE niepuste teksty
  const col = await page.evaluate(() => {
    const trs = [...document.querySelectorAll("#dataTable tbody tr")].slice(0, 8);
    const rows = trs.map((tr) => currentDisplayModel.rows.find((r) => getRowSelectionKey(r) === tr.dataset.rowKey));
    for (let c = 0; c < currentHeaders.length; c++) {
      const vs = rows.map((r) => r && r.values[c]);
      if (vs.every((v) => typeof v === "string" && v.trim()) && new Set(vs).size >= 4) return c;
    }
    return 1;
  });

  // 1. nic nie skopiowano
  const m0 = await menu(0, col);
  check("1. bez skopiowania brak „Wklej”", !m0.some((i) => i.id === "cmi-paste"), m0);
  await page.keyboard.press("Escape");

  // 2. kopiuj komórkę 0 → wklej w komórkę 3
  const src = await val(0, col);
  await menu(0, col); await clickItem("copy");
  const m1 = await menu(3, col);
  const pasteItem = m1.find((i) => i.id === "cmi-paste");
  check("2. po Kopiuj jest „Wklej” z wartością", pasteItem && pasteItem.text.includes(String(src).slice(0, 10)), m1);
  const p0 = await pending();
  await clickItem("paste");
  check("2. wklejona dokładnie ta wartość", (await val(3, col)) === src, { got: await val(3, col), src });
  check("6. jedna edycja w pendingEdits", (await pending()) - p0 === 1, await pending());
  await page.evaluate(() => undoLast()); await sleep(150);
  check("6. Cofnij przywraca", (await val(3, col)) !== src && (await pending()) === p0, await val(3, col));

  // 3. formuła ze schowka się nie wkleja
  await page.evaluate(() => { internalClipboard = { tsv: "=1+1", rows: 1, cols: 1, ts: Date.now() }; });
  const before5 = await val(5, col);
  await menu(5, col); await clickItem("paste");
  check("3. formuła nie wklejona", (await val(5, col)) === before5, await val(5, col));

  // 4. jedna wartość + zakres 3×1 → wszystkie 3
  await menu(0, col); await clickItem("copy");
  await sleep(600);
  await click(4, col); await click(6, col, true);
  const m2 = await menu(5, col);
  const fill = m2.find((i) => i.id === "cmi-paste");
  check("4. „Wklej … do 3 komórek”", fill && /3/.test(fill.text), m2);
  const p1 = await pending();
  await clickItem("paste");
  const filled = [await val(4, col), await val(5, col), await val(6, col)];
  check("4. cały zakres wypełniony", filled.every((v) => v === src), filled);
  await page.evaluate(() => undoLast()); await sleep(150);
  check("6. Cofnij cofa cały zakres naraz", (await pending()) === p1, await pending());

  // 5. blok 2×2 → od lewego górnego rogu zaznaczenia (zaznaczamy „od dołu")
  const block = [[await val(0, col), await val(0, col + 1)], [await val(1, col), await val(1, col + 1)]];
  await sleep(600);
  await click(0, col); await click(1, col + 1, true);
  await menu(0, col); await clickItem("copy");
  const m3 = await menu(8, col);
  check("5. „Wklej blok 2×2”", m3.some((i) => i.id === "cmi-paste" && /2×2/.test(i.text)), m3);
  await page.keyboard.press("Escape");
  await sleep(600); // po prawym kliku apka celowo połyka kliknięcia komórek przez 0,5 s (anty-duch po przytrzymaniu)
  await click(10, col + 1); await click(8, col, true); // zaznaczenie od prawego dołu do lewej góry
  await menu(9, col); await clickItem("paste");
  const got = [[await val(8, col), await val(8, col + 1)], [await val(9, col), await val(9, col + 1)]];
  check("5. blok wklejony od lewego górnego rogu", JSON.stringify(got) === JSON.stringify(block), { got, block });

  // 7. zapis: wartości w pliku, styl bez zmian
  const saved = await page.evaluate(async () => {
    const bytes = await buildOutputBytes("xlsx");
    const zipA = await JSZip.loadAsync(originalFileBytes);
    const zipB = await JSZip.loadAsync(bytes);
    const out = [];
    const wb = XLSX.read(bytes, {});
    const sheetIdx = workbook.SheetNames.indexOf(currentSheetName) + 1;
    const path = `xl/worksheets/sheet${sheetIdx}.xml`;
    const xa = await zipA.file(path).async("string");
    const xb = await zipB.file(path).async("string");
    const sAttr = (xml, ref) => ((xml.match(new RegExp(`<c r="${ref}"[^>]*`)) || [""])[0].match(/ s="(\d+)"/) || [])[1] || null;
    for (const ref of Object.keys(pendingEdits[currentSheetName] || {})) {
      out.push({ ref, want: pendingEdits[currentSheetName][ref]?.v, got: wb.Sheets[currentSheetName][ref]?.v, sA: sAttr(xa, ref), sB: sAttr(xb, ref) });
    }
    return out;
  });
  check("7. wklejone wartości są w pliku", saved.length > 0 && saved.every((c) => String(c.got) === String(c.want)), saved.slice(0, 3));
  check("7. styl komórek bez zmian", saved.every((c) => c.sA === c.sB), saved.filter((c) => c.sA !== c.sB).slice(0, 3));

  check("brak błędów konsoli", errors.length === 0, errors);
  await browser.close();

  for (const r of results) console.log(`${r.ok ? "✅" : "❌"} ${r.name}${r.detail && !r.ok ? `  (${r.detail})` : ""}`);
  const failed = results.filter((r) => !r.ok).length;
  if (failed) { console.error(`\n${failed} z ${results.length} nie przeszło`); process.exit(1); }
  console.log(`\n✅ cell-menu-paste: ${results.length}/${results.length}`);
  process.exit(0);
}

run().catch((e) => { console.error(e); process.exit(1); });
