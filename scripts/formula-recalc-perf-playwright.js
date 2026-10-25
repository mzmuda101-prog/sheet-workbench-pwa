// formula-recalc-perf-playwright.js — przeliczenie zależnych formuł po edycji bez blokady.
//
// Film Mateusza z iPhone'a (2026-10-25): Wklej / Enter w arkuszu „2025-2026” stało ~2 s.
// Profil: jedna edycja przeliczała 771 formuł z innych arkuszy, w tym 148 × LICZ.JEŻELI po
// tabeli 563×32 i 564 × SUMA.WARUNKÓW po całych kolumnach — każda budowała ten sam zakres
// od nowa, komórka po komórce (≈2,9 mln odczytów). Do tego pierwsza edycja budowała indeks
// zależności 28 tys. formuł.
//
// Plik użytkownika nie jest w repo (*.xlsx w .gitignore) — skoroszyt o tym samym kształcie
// budujemy w przeglądarce. Co musi być prawdą:
//   1. indeks zależności buduje się sam w bezczynności po wczytaniu (bez czekania na edycję),
//   2. edycja komórki w danych przelicza zależne formuły z małą liczbą odczytów komórek
//      (zakres budowany raz na przeliczenie, nie raz na formułę),
//   3. wyniki po edycji są poprawne (LICZ.JEŻELI / SUMA.WARUNKÓW / LICZ.WARUNKI z operatorem /
//      symbolem wieloznacznym / liczbą jako kryterium — skrót „równa się tekst” i zwykła ścieżka),
//   4. wszystkie formuły skoroszytu liczą się identycznie z pamięcią zakresów i bez niej.

const playwright = require("playwright");

const APP_URL = process.env.APP_URL || "http://127.0.0.1:4175/";
const ENGINE = process.env.ENGINE || "chromium";

async function run() {
  const browser = await playwright[ENGINE].launch({ headless: true });
  const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1280, height: 860 } });
  await context.addInitScript(() => localStorage.setItem("introPlayed", "true"));
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push("pageerror: " + e.message));
  const results = [];
  const check = (name, ok, detail = "") => results.push({ name, ok: !!ok, detail: typeof detail === "string" ? detail : JSON.stringify(detail) });

  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(async () => {
    document.getElementById("heroSplash")?.remove();
    await ensureXlsxLibs(false);
    const ROWS = 600, BLOCKS = 8;
    const names = ["Anna Nowak", "Jan Kowalski", "Piotr Wiśniewski", "Ewa Zielińska", "Adam Lis", "Ola Mazur"];
    const head = ["Nr"];
    for (let b = 1; b <= BLOCKS; b++) head.push(`Osoba${b}`, `od${b}`, `do${b}`, `Dni${b}`);
    const data = [head];
    for (let r = 0; r < ROWS; r++) {
      const row = [r + 1];
      for (let b = 0; b < BLOCKS; b++) row.push(names[(r + b) % names.length], 100 + r, 130 + r, (r * 7 + b) % 50);
      data.push(row);
    }
    const ws = XLSX.utils.aoa_to_sheet(data);
    const last = ROWS + 1;
    const an = [["Osoba", "Ile w tabeli", "Suma dni", "Ile (A*)", "Ile >20", "Ile = 7"]];
    names.forEach((n) => an.push([n]));
    const aws = XLSX.utils.aoa_to_sheet(an);
    // 6 osób × 25 kopii = 150 wierszy z tym samym zestawem formuł (jak „Analiza Pracowników”)
    let r = 2;
    for (let k = 0; k < 25; k++) {
      names.forEach((n, i) => {
        aws[`A${r}`] = { t: "s", v: n };
        aws[`B${r}`] = { t: "n", v: 0, f: `COUNTIF(Dane!$B$2:$AF$${last},A${r})` };
        aws[`C${r}`] = { t: "n", v: 0, f: `SUMIFS(Dane!$E:$E,Dane!$B:$B,A${r})+SUMIFS(Dane!$I:$I,Dane!$F:$F,A${r})` };
        aws[`D${r}`] = { t: "n", v: 0, f: `COUNTIF(Dane!$B$2:$B$${last},"A*")` };
        aws[`E${r}`] = { t: "n", v: 0, f: `COUNTIF(Dane!$E$2:$E$${last},">20")` };
        aws[`F${r}`] = { t: "n", v: 0, f: `COUNTIF(Dane!$E$2:$E$${last},7)` };
        r++;
      });
    }
    aws["!ref"] = `A1:F${r - 1}`;
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Dane");
    XLSX.utils.book_append_sheet(wb, aws, "Analiza");
    const buf = XLSX.write(wb, { type: "array", bookType: "xlsx" });
    await handleFile(new File([buf], "perf-recalc.xlsx"));
  });
  await page.waitForFunction(() => document.getElementById("sheetSelect")?.options?.length > 1, null, { timeout: 30000 });
  if (!(await page.evaluate(() => currentSheetName === "Dane" && !!document.querySelector("#dataTable tbody tr[data-row-key]")))) {
    await page.evaluate(() => { const sel = document.getElementById("sheetSelect"); sel.value = "Dane"; loadBtn.click(); });
  }
  await page.waitForFunction(() => currentSheetName === "Dane" && !!document.querySelector("#dataTable tbody tr[data-row-key]"), null, { timeout: 20000 });

  // 1. indeks w tle
  const built = await page.waitForFunction(() => typeof fcIndex !== "undefined" && fcIndex && fcIndex.wb === workbook, null, { timeout: 8000 }).then(() => true).catch(() => false);
  const idxSize = await page.evaluate(() => (fcIndex && fcIndex.entries.size) || 0);
  check("1. indeks zależności zbudowany w bezczynności po wczytaniu", built && idxSize >= 700, { built, idxSize });

  // 2 + 3. edycja w danych: B2 (Osoba1) „Anna Nowak” → „Adam Lis”, E2 (Dni1) → 7
  const edit = await page.evaluate(async () => {
    const ws = workbook.Sheets.Dane;
    let reads = 0;
    fcContext("Dane", 0, 0);
    const origGet = fcBaseCtx.get;
    fcBaseCtx.get = (...a) => { reads++; return origGet(...a); };
    const t0 = performance.now();
    const colOf = (h) => currentDisplayModel.headers.indexOf(h);
    updateSheetCell(1, colOf("Osoba1"), { type: "string", value: "Adam Lis" });
    updateSheetCell(1, colOf("Dni1"), { type: "number", value: 7 });
    await new Promise((r) => setTimeout(r, 0)); // przeliczenie w mikrozadaniu
    const ms = Math.round(performance.now() - t0);
    fcBaseCtx.get = origGet;
    // wzorzec: liczenie „na piechotę” z arkusza
    const cells = (c1, c2, r1, r2) => { const out = []; for (let r = r1; r <= r2; r++) for (let c = c1; c <= c2; c++) { const x = ws[XLSX.utils.encode_cell({ r, c })]; out.push(x ? x.v : null); } return out; };
    const range = XLSX.utils.decode_range(ws["!ref"]);
    const all = cells(1, 32, 1, range.e.r);
    const colVals = (c) => cells(c, c, 1, range.e.r);
    const a = workbook.Sheets.Analiza;
    const val = (ref) => a[ref] && a[ref].v;
    const exp = {};
    ["Anna Nowak", "Adam Lis"].forEach((n) => {
      exp[n] = {
        count: all.filter((v) => typeof v === "string" && v.toLowerCase() === n.toLowerCase()).length,
        sum: colVals(1).reduce((s, v, i) => s + (v === n ? colVals(4)[i] : 0), 0) + colVals(5).reduce((s, v, i) => s + (v === n ? colVals(8)[i] : 0), 0),
      };
    });
    const rowOf = (n) => { for (let r = 2; r < 200; r++) if (val(`A${r}`) === n) return r; return -1; };
    const ra = rowOf("Anna Nowak"), rb = rowOf("Adam Lis");
    const starA = colVals(1).filter((v) => typeof v === "string" && /^a/i.test(v)).length;
    const gt20 = colVals(4).filter((v) => typeof v === "number" && v > 20).length;
    const eq7 = colVals(4).filter((v) => v === 7).length;
    return {
      reads, ms,
      got: { annaCount: val(`B${ra}`), annaSum: val(`C${ra}`), adamCount: val(`B${rb}`), adamSum: val(`C${rb}`), starA: val(`D${ra}`), gt20: val(`E${ra}`), eq7: val(`F${ra}`), lastCopy: val(`B${ra + 6 * 24}`) },
      exp: { annaCount: exp["Anna Nowak"].count, annaSum: exp["Anna Nowak"].sum, adamCount: exp["Adam Lis"].count, adamSum: exp["Adam Lis"].sum, starA, gt20, eq7 },
    };
  });
  // przed poprawką: ~150 × 19 200 odczytów samego LICZ.JEŻELI po tabeli (≈ 3 mln)
  check("2. przeliczenie po edycji czyta zakresy raz (mało odczytów komórek)", edit.reads > 0 && edit.reads < 150000, { reads: edit.reads, ms: edit.ms });
  const g = edit.got, e = edit.exp;
  check("3. LICZ.JEŻELI „równa się tekst” po tabeli — poprawne po edycji", g.annaCount === e.annaCount && g.adamCount === e.adamCount && g.lastCopy === e.annaCount, { g, e });
  check("3. SUMA.WARUNKÓW po całych kolumnach — poprawne po edycji", g.annaSum === e.annaSum && g.adamSum === e.adamSum, { g, e });
  check("3. kryteria spoza skrótu (A*, >20, liczba 7) — poprawne", g.starA === e.starA && g.gt20 === e.gt20 && g.eq7 === e.eq7, { g, e });

  // 4. równoważność: każda formuła z pamięcią zakresów = bez niej
  const eq = await page.evaluate(() => {
    const FE = FormulaEngine;
    const idx = fcEnsureIndex();
    const list = [...idx.entries.values()].filter((x) => x.ok && x.ast);
    const same = (a, b) => a === b || (FE.isErr(a) && FE.isErr(b) && a.err === b.err) || (typeof a === "number" && typeof b === "number" && Math.abs(a - b) < 1e-9);
    fcContext("Dane", 0, 0);
    fcBaseCtx.matrixCache = null;
    const plain = list.map((x) => FE.evaluate(x.ast, fcContext(x.sheet, x.r, x.c)));
    fcBaseCtx.matrixCache = FE.makeMatrixCache();
    let diff = 0;
    list.forEach((x, i) => { if (!same(FE.evaluate(x.ast, fcContext(x.sheet, x.r, x.c)), plain[i])) diff++; });
    fcBaseCtx.matrixCache = null;
    return { n: list.length, diff };
  });
  check("4. wszystkie formuły: z pamięcią zakresów = bez niej", eq.n >= 700 && eq.diff === 0, eq);

  check("brak błędów strony", errors.length === 0, errors.slice(0, 3).join(" | "));
  await browser.close();

  let failed = 0;
  for (const r of results) {
    console.log(`${r.ok ? "✅" : "❌"} ${r.name}${r.detail ? `  (${r.detail})` : ""}`);
    if (!r.ok) failed += 1;
  }
  if (failed) { console.error(`\n${failed} z ${results.length} nie przeszło`); process.exit(1); }
  console.log(`\n✅ formula-recalc-perf (${ENGINE}): ${results.length}/${results.length}`);
}

run().catch((e) => { console.error(e); process.exit(1); });
