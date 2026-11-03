// Strażnik poprawek z polowania na błędy (2026-11-03) — tabela Sheet Workbench:
//   1. pusty arkusz: „Arkusz bez danych” (nie „Zmień filtry”), bez przycisków ekranu startowego,
//      napis przetrwa zmianę języka; nowy silnik odmontowany, stare wiersze zdjęte,
//   2. wylany tekst: otwarty edytor komórki przeżywa przeliczenie wylewania,
//   3. przeniesienie bloku na komórkę z formułą zależną: wartość wygrywa (dawniej formuła cykliczna),
//   4. „Zablokuj kolumny do tej” przy scalonym nagłówku liczy numer kolumny, nie pozycję komórki.
//   node scripts/bughunt-1103-playwright.js        ENGINE=webkit node scripts/bughunt-1103-playwright.js
const playwright = require("playwright");
const path = require("path");
const fs = require("fs");
const os = require("os");
const XLSX = require(path.join(__dirname, "..", "lib", "xlsx.full.min.js"));

const APP_URL = process.env.APP_URL || "http://127.0.0.1:4175/";
const ENGINE = process.env.ENGINE || "chromium";
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

function buildFixture() {
  const wb = XLSX.utils.book_new();
  const big = [["Nazwa", "N", "Opis", "Puste", "Kwota"]];
  for (let i = 0; i < 400; i++) big.push([`poz ${i}`, i, i % 2 ? `bardzo długi tekst który na pewno wylewa się na sąsiednie puste komórki ${i}` : "", "", i * 2]);
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(big), "Duzy");
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([]), "Pusty");
  const m = XLSX.utils.aoa_to_sheet([["Scalony", null, "C", "D", "E"], [1, 2, 3, 4, 5], [6, 7, 8, 9, 10]]);
  m["!merges"] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 1 } }];
  XLSX.utils.book_append_sheet(wb, m, "Scalenia");
  const file = path.join(fs.mkdtempSync(path.join(os.tmpdir(), "swb-bh-")), "bughunt-1103.xlsx");
  fs.writeFileSync(file, XLSX.write(wb, { type: "buffer", bookType: "xlsx" }));
  return file;
}

async function run() {
  const FILE = buildFixture();
  const browser = await playwright[ENGINE].launch({ headless: true });
  const fails = [];
  const check = (name, ok, detail) => { console.log(`${ok ? "✅" : "❌"} ${name}${ok ? "" : ` — ${JSON.stringify(detail)}`}`); if (!ok) fails.push(name); };
  const ctx = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1280, height: 800 } });
  await ctx.addInitScript(() => localStorage.setItem("introPlayed", "true"));
  const p = await ctx.newPage();
  const errors = [];
  p.on("pageerror", (e) => errors.push(e.message + " " + (e.stack || "").split("\n").slice(1, 4).join(" | ")));
  await p.goto(APP_URL, { waitUntil: "load" });
  await p.evaluate(() => { document.getElementById("heroSplash")?.remove(); try { ensureXlsxLibs(false); } catch (_) {} });
  await p.setInputFiles("#fileInput", FILE);
  await p.waitForFunction(() => typeof workbook !== "undefined" && workbook && baseRows.length >= 400, null, { timeout: 30000 });
  await sleep(1200);
  const sheet = async (name) => { await p.evaluate((n) => { sheetSelect.value = n; loadBtn.click(); }, name); await p.waitForFunction((n) => currentSheetName === n, name); await sleep(900); };

  // 2. edytor w komórce z wylanym tekstem
  const spill = await p.evaluate(async () => {
    const wait = (ms) => new Promise((r) => setTimeout(r, ms));
    const td = document.querySelector("#dataTable tbody td.cell-spill");
    if (!td) return { noSpill: true };
    openCellEditor(td);
    await wait(50);
    swbSpill.refreshRow(td.closest("tr"));
    await wait(300);
    const ok = !!(activeCellEditor && activeCellEditor.input.isConnected);
    if (activeCellEditor) activeCellEditor.close();
    return { ok, virt: swbVirt.isActive() };
  });
  check("2. edytor w wylanej komórce zostaje podłączony", spill.ok && spill.virt, spill);

  // 1. pusty arkusz
  await sheet("Pusty"); // arkusz bez <dimension> (SheetJS) — dawniej wyjątek, przełączenie nie działało
  const empty = await p.evaluate(() => ({
    title: emptyTitleEl.textContent, buttons: !!document.getElementById("emptyOpenBtn").offsetParent,
    virt: swbVirt.isActive(), staleRows: document.querySelectorAll("#dataTable tbody tr").length,
  }));
  check("1. pusty arkusz: „Arkusz bez danych”, bez przycisków, bez starych wierszy", empty.title === "Arkusz bez danych" && !empty.buttons && !empty.virt && empty.staleRows === 0, empty);
  await p.evaluate(() => applyLanguage("en"));
  await sleep(400);
  const en = await p.evaluate(() => emptyTitleEl.textContent);
  check("1. po zmianie języka nadal pusty arkusz (nie ekran startowy)", en === "No data in this sheet", en);
  await p.evaluate(() => applyLanguage("pl"));
  await sleep(300);

  // 3. przeniesienie na komórkę z formułą zależną
  await sheet("Scalenia");
  const moved = await p.evaluate(async () => {
    sheetEditUi.setMode("edit");
    const sh = workbook.Sheets[currentSheetName];
    // wiersz 2 (index0 = 1): A2 = 1, B2 = A2*10 → przenieś A2 na B2
    swbFormulaCommitAt(1, 1, "=A2*10", { silent: true });
    const before = sh.B2 && sh.B2.f;
    swbMoveBlock({ r1: 1, r2: 1, c1: 0, c2: 0 }, 0, 1);
    await new Promise((r) => setTimeout(r, 300));
    return { before, B2: sh.B2 ? { v: sh.B2.v, f: sh.B2.f || null } : null };
  });
  check("3. przeniesiona wartość przykrywa formułę zależną", /A2\*10/i.test(moved.before || "") && moved.B2 && moved.B2.v === 1 && !moved.B2.f, moved);

  // 4. blokada do kolumny C przy scalonym nagłówku A:B
  const fz = await p.evaluate(async () => {
    const thC = document.querySelector('#dataTable thead tr.header-row th[data-col-index="2"]');
    if (!thC) return { noTh: true };
    thC.dispatchEvent(new MouseEvent("contextmenu", { bubbles: true, cancelable: true, clientX: 300, clientY: 200 }));
    await new Promise((r) => setTimeout(r, 300));
    const item = document.querySelector('[data-id="hdr-freeze"], #hdr-freeze, .cmi-hdr-freeze, [id="hdr-freeze"]');
    const items = [...document.querySelectorAll("#cellMenu .cell-menu-item, #cellMenu button")].map((b) => b.id || b.className);
    if (!item) return { noItem: true, items };
    item.click();
    await new Promise((r) => setTimeout(r, 300));
    return { n: getFrozenColCount() };
  });
  check("4. „Zablokuj kolumny do tej” (C) przy scalonym A:B = 3 kolumny", fz.n === 3, fz);
  check("bez błędów strony", !errors.length, errors);
  await browser.close();
  if (fails.length) { console.log(`❌ bughunt-1103 (${ENGINE}): ${fails.length} błędów`); process.exit(1); }
  console.log(`✅ bughunt-1103 (${ENGINE}): wszystko OK`);
}
run().catch((e) => { console.error(e); process.exit(1); });
