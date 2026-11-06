// move-other-sheet-playwright.js — przeniesienie bloku (Wytnij → Wklej) przepina formuły
// w INNYCH arkuszach (lista 3.11; Excel robi to samo). Jedno Cofnij cofa wszystko, Ponów wraca.

const { chromium, webkit } = require("playwright");
const fs = require("fs");
const os = require("os");
const path = require("path");
const XLSX = require(path.join(__dirname, "..", "lib", "xlsx.full.min.js"));

const APP_URL = process.env.APP_URL || "http://127.0.0.1:4175/";
const ENGINE = process.env.ENGINE === "webkit" ? webkit : chromium;

function fixture() {
  const dane = XLSX.utils.aoa_to_sheet([["Kwota", "Opis"], [5, "a"], [7, "b"], [null, null], [null, null], [null, null]]);
  const raport = XLSX.utils.aoa_to_sheet([["Wynik", "Suma"], [null, null]]);
  raport.A2 = { t: "n", v: 10, f: "Dane!A2*2" };
  raport.B2 = { t: "n", v: 12, f: "SUM(Dane!A2:A3)" }; // zakres częściowo w bloku — zostaje
  raport.C2 = { t: "n", v: 7, f: "'Dane'!$A$3" };
  raport["!ref"] = "A1:C2";
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, dane, "Dane");
  XLSX.utils.book_append_sheet(wb, raport, "Raport");
  const file = path.join(fs.mkdtempSync(path.join(os.tmpdir(), "mvx-")), "move-other.xlsx");
  fs.writeFileSync(file, XLSX.write(wb, { type: "buffer", bookType: "xlsx" }));
  return file;
}

(async () => {
  const file = fixture();
  const browser = await ENGINE.launch();
  const ctx = await browser.newContext({ viewport: { width: 1200, height: 800 }, serviceWorkers: "block" });
  await ctx.addInitScript(() => { localStorage.setItem("introPlayed", "true"); });
  const page = await ctx.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  await page.goto(APP_URL);
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  await page.setInputFiles("#fileInput", file);
  await page.waitForFunction(() => document.getElementById("sheetSelect")?.options?.length > 0, null, { timeout: 30000 });
  await page.evaluate(() => loadBtn.click());
  await page.waitForFunction(() => Array.isArray(baseRows) && baseRows.length > 0 && currentSheetName === "Dane", null, { timeout: 20000 });
  const fails = [];
  const ok = (n, c, got) => { if (!c) fails.push(`${n} — ${JSON.stringify(got)}`); };
  const read = () => page.evaluate(() => {
    const r = workbook.Sheets.Raport;
    return { a2: r.A2 && r.A2.f, a2v: r.A2 && r.A2.v, b2: r.B2 && r.B2.f, c2: r.C2 && r.C2.f, c2v: r.C2 && r.C2.v, pend: pendingEdits.Raport ? Object.keys(pendingEdits.Raport) : [], src: workbook.Sheets.Dane.A2 ? workbook.Sheets.Dane.A2.v : null };
  });
  // A2:A3 (5, 7) → B4:B5
  await page.waitForFunction(() => typeof currentDisplayModel !== "undefined" && currentDisplayModel, null, { timeout: 20000 });
  await page.evaluate(async () => { sheetEditUi.setMode("edit"); swbMoveBlock({ r1: 1, r2: 2, c1: 0, c2: 0 }, 2, 1); await new Promise((r) => setTimeout(r, 300)); });
  const m = await read();
  ok("formuła w innym arkuszu idzie za komórką", /^Dane!B4\*2$/i.test(m.a2 || ""), m);
  ok("… i ma wartość", m.a2v === 10, m);
  ok("zakres w całości w bloku też idzie (SUMA)", /^SUM\(Dane!B4:B5\)$/i.test(m.b2 || ""), m);
  ok("odwołanie bezwzględne z apostrofem idzie za komórką", /Dane'?!\$B\$5$/i.test(m.c2 || "") && m.c2v === 7, m);
  ok("zmiana trafi do zapisu", m.pend.includes("A2") && m.pend.includes("C2"), m.pend);
  await page.evaluate(() => undoLast());
  await page.waitForTimeout(200);
  const u = await read();
  ok("jedno Cofnij: formuły w innym arkuszu wracają", u.a2 === "Dane!A2*2" && u.b2 === "SUM(Dane!A2:A3)" && !u.pend.length, u);
  ok("jedno Cofnij: przeniesiona wartość wraca", u.src === 5, u);
  await page.evaluate(() => undoRedo("redo"));
  await page.waitForTimeout(200);
  const rd = await read();
  ok("Ponów: znowu przepięte", /^Dane!B4\*2$/i.test(rd.a2 || "") && rd.src == null, rd);
  if (errors.length) fails.push("błędy strony: " + errors.join(" | "));
  await browser.close();
  if (fails.length) { console.log(`❌ move-other-sheet (${process.env.ENGINE || "chromium"}): ${fails.length} błędów`); fails.forEach((f) => console.log("  ✗", f)); process.exit(1); }
  console.log(`✅ move-other-sheet (${process.env.ENGINE || "chromium"}): wszystko przeszło`);
})().catch((e) => { console.error(e); process.exit(1); });
