// cell-spill-playwright.js — TEKST WYLEWA SIĘ NA PUSTE KOMÓRKI OBOK, jak w Excelu (2026-10-31).
//   node scripts/cell-spill-playwright.js        ENGINE=webkit node scripts/cell-spill-playwright.js
//
// Wyrocznia Excela (arkusz → PDF z Excela): tekst do lewej → w prawo do pierwszej niepustej
// komórki; do prawej → w lewo; wyśrodkowany → w obie; liczba nigdy; spacja i formuła ""
// zatrzymują; zawijanie nie wylewa. Pilnuje:
//   1. tekst do lewej wychodzi na pustą B i kończy się na zajętej C;
//   2. formuła "" i spacja w sąsiedniej komórce zatrzymują;
//   3. liczba się nie wylewa;
//   4. do prawej → w lewo (widać koniec tekstu), wyśrodkowany → w obie strony;
//   5. zawijanie (format komórki) wyłącza wylewanie;
//   6. wpisanie czegoś w pustą komórkę obok obcina tekst; wyczyszczenie znów go wylewa;
//   7. podpowiedź „ucięty tekst” nie pokazuje się, gdy tekst mieści się dzięki wylaniu;
//   8. tabela z samym tekstem (grafik): żaden wiersz nie jest „podnagłówkiem” — wszystko edytowalne;
//      tabela z liczbami: wiersz z samym tekstem dalej jest podnagłówkiem, ale też da się go edytować.
// Uruchom z serwerem na APP_URL (domyślnie http://127.0.0.1:4175/).

const playwright = require("playwright");

const APP_URL = process.env.APP_URL || "http://127.0.0.1:4175/";
const ENGINE = process.env.ENGINE || "chromium";
const SLEEP_SCALE = Math.min(1, Math.max(0.1, Number(process.env.SLEEP_SCALE || 1)));
const sleep = (ms) => new Promise((r) => setTimeout(r, Math.round(ms * SLEEP_SCALE)));
const LONG = "Długi opis zmiany, który na pewno nie mieści się w jednej wąskiej komórce";

async function run() {
  const browser = await playwright[ENGINE].launch({ headless: true });
  const fails = [];
  const check = (name, ok, detail) => {
    console.log(`${ok ? "✅" : "❌"} ${name}${ok ? "" : ` — ${JSON.stringify(detail)}`}`);
    if (!ok) fails.push(name);
  };
  const ctx = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1280, height: 900 } });
  await ctx.addInitScript(() => {
    if (sessionStorage.getItem("__spInit")) return;
    sessionStorage.setItem("__spInit", "1");
    localStorage.clear();
    localStorage.setItem("introPlayed", "true");
  });
  const p = await ctx.newPage();
  const errors = [];
  p.on("pageerror", (e) => errors.push(e.message));
  await p.goto(APP_URL, { waitUntil: "load" });
  await p.evaluate(async (L) => {
    document.getElementById("heroSplash")?.remove();
    await ensureXlsxLibs(true);
    const aoa = [
      ["Kol A", "Kol B", "Kol C", "Kol D", "Kol E", "Kol F"],
      [L, null, "x", null, null, null],          // r2: w prawo do C
      [L, null, null, null, null, null],          // r3: B = formuła ""
      [L, " ", null, null, null, null],           // r4: B = spacja
      [12345678901234567890123, null, null, null, null, null], // r5: liczba
      [null, null, L, null, null, null],          // r6: do prawej (format niżej)
      ["y", null, L, null, null, null],           // r7: wyśrodkowany, A zajęta
      [L, null, null, null, null, null],          // r8: zawijanie
      [L, null, null, null, null, "z"],           // r9: do edycji w B (sam tekst — jak w grafiku)
      ["krótki", null, null, null, null, null],   // r10: krótki — bez wylewania
    ];
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    ws.B3 = { t: "s", v: "", f: '""' };
    ws["!cols"] = Array(6).fill({ wch: 12 });
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Grafik");
    const buf = XLSX.write(wb, { type: "array", bookType: "xlsx" });
    const ok = await handleFile(new File([buf], "spill.xlsx", { type: XLSX_MIME }));
    if (ok) openDefaultSheet();
  }, LONG);
  await p.waitForFunction(() => document.querySelectorAll("#dataTable tbody tr[data-row-key]").length >= 9, null, { timeout: 20000 });
  await p.waitForFunction(() => !document.documentElement.classList.contains("vt-scene"), null, { timeout: 10000 });
  await p.evaluate(() => {
    sheetEditUi.setMode("edit");
    const r = (n) => viewRows.find((x) => x.rowIndex0 === n - 1);
    cfmtApplyAt(r(6), 2, { h: "right" });
    cfmtApplyAt(r(7), 2, { h: "center" });
    cfmtApplyAt(r(8), 0, { wrap: true });
    for (let i = 0; i < 6; i++) manualColumnWidths[i] = 90; // wąskie kolumny, jak w grafiku
    renderActiveTable();
  });
  await sleep(500);

  const sub = await p.evaluate(() => viewRows.filter((r) => r.isSubheader).map((r) => r.rowIndex0 + 1));
  check("8. tabela z samym tekstem: brak „podnagłówków” (szarych, zablokowanych wierszy)", sub.length === 0, sub);
  const subNum = await p.evaluate(() => {
    const mk = (vals) => ({ values: vals });
    const rows = [mk(["Dział A", "opis", null]), mk(["Jan", 12, 3]), mk(["Ewa", 7, 1]), mk(["Ola", 4, 2]), mk(["Dział B", "opis", null]), mk(["Piotr", 9, 5])];
    markSubheaderRows(rows);
    return rows.map((r) => !!r.isSubheader);
  });
  check("8. tabela z liczbami: „Dział A”/„Dział B” dalej są podnagłówkami", JSON.stringify(subNum) === JSON.stringify([true, false, false, false, true, false]), subNum);
  const info = () => p.evaluate(() => {
    const out = {};
    const colW = [...document.querySelectorAll("#dataTable thead tr.guide-row th.guide-cell")].map((th) => th.getBoundingClientRect().width);
    for (const tr of document.querySelectorAll("#dataTable tbody tr[data-row-index]")) {
      const n = +tr.dataset.rowIndex + 1;
      out[n] = [...tr.querySelectorAll("td[data-col-index]")].map((td) => {
        const s = td._swbSpill;
        const range = document.createRange();
        range.selectNodeContents(td);
        const tb = range.getBoundingClientRect();
        const cb = td.getBoundingClientRect();
        return s ? { L: Math.round(s.L), R: Math.round(s.R), a: s.align, textL: Math.round(tb.left - cb.left), textR: Math.round(tb.right - cb.right) } : null;
      });
    }
    return { rows: out, colW: colW.map(Math.round) };
  });
  let s = await info();
  const W = s.colW;
  const near = (a, b) => Math.abs(a - b) <= 2;
  check("1. tekst wychodzi na pustą B i kończy się przed zajętą C", s.rows[2][0] && near(s.rows[2][0].R, W[1]) && s.rows[2][0].L === 0, { r: s.rows[2][0], W });
  check("1. tekst naprawdę jest narysowany dalej niż komórka", s.rows[2][0] && s.rows[2][0].textR > 10, s.rows[2][0]);
  check("2. formuła \"\" w B zatrzymuje", s.rows[3][0] === null, s.rows[3][0]);
  check("2. spacja w B zatrzymuje", s.rows[4][0] === null, s.rows[4][0]);
  check("3. liczba się nie wylewa", s.rows[5][0] === null, s.rows[5][0]);
  check("4. do prawej → w lewo na puste A i B", s.rows[6][2] && s.rows[6][2].a === "right" && near(s.rows[6][2].L, W[0] + W[1]) && s.rows[6][2].R === 0, s.rows[6][2]);
  check("4. …i widać koniec tekstu (tekst wystaje w lewo, nie w prawo)", s.rows[6][2] && s.rows[6][2].textL < -10 && s.rows[6][2].textR <= 1, s.rows[6][2]);
  check("4. wyśrodkowany → w obie strony (A zajęta: tylko B z lewej)", s.rows[7][2] && s.rows[7][2].a === "center" && near(s.rows[7][2].L, W[1]) && s.rows[7][2].R > W[3], s.rows[7][2]);
  check("5. zawijanie wyłącza wylewanie", s.rows[8][0] === null, s.rows[8][0]);
  check("krótki tekst bez wylewania (bez zbędnych warstw)", s.rows[10][0] === null, s.rows[10][0]);
  check("9. tekst do F: B–E puste", s.rows[9][0] && near(s.rows[9][0].R, W[1] + W[2] + W[3] + W[4]), { r: s.rows[9][0], W });

  // 7. podpowiedź
  const trunc = await p.evaluate(() => {
    const td = (n, c) => document.querySelector(`#dataTable tbody tr[data-row-index="${n - 1}"] td[data-col-index="${c}"]`);
    return { fits: isCellTextTruncated(td(9, 0)), cut: isCellTextTruncated(td(4, 0)) };
  });
  check("7. tekst zmieszczony dzięki wylaniu ≠ ucięty; zablokowany = ucięty", trunc.fits === false && trunc.cut === true, trunc);

  // 6. wpis w pustą B9 obcina, wyczyszczenie znów wylewa
  const b9 = '#dataTable tbody tr[data-row-index="8"] td[data-col-index="1"]';
  await p.dblclick(b9);
  await sleep(200);
  await p.keyboard.type("Q");
  await p.keyboard.press("Enter");
  await sleep(400);
  s = await info();
  check("6. wpis w B9 obcina tekst z A9", s.rows[9][0] === null, s.rows[9][0]);
  await p.click(b9);
  await p.keyboard.press("Delete");
  await sleep(400);
  s = await info();
  check("6. wyczyszczenie B9 → tekst znów się wylewa", s.rows[9][0] && s.rows[9][0].R > W[1], s.rows[9][0]);
  await p.evaluate(() => undoRedo("undo"));
  await sleep(400);
  s = await info();
  check("6. Cofnij (Q wraca) → znów obcięte", s.rows[9][0] === null, s.rows[9][0]);

  // 8. edycja w wierszu-podnagłówku (domysł) — nie jest blokowana, jak w Excelu
  const subEdit = await p.evaluate(() => {
    const row = viewRows.find((r) => r.rowIndex0 === 9);
    row.isSubheader = true; // wymuszony domysł
    document.querySelectorAll(".toast").forEach((x) => x.remove());
    const td = document.querySelector('#dataTable tbody tr[data-row-index="9"] td[data-col-index="2"]');
    openCellEditor(td);
    const opened = !!td.querySelector("input.cell-editor");
    if (opened) { const inp = td.querySelector("input.cell-editor"); inp.value = "ok"; inp.dispatchEvent(new KeyboardEvent("keydown", { key: "Enter", bubbles: true })); }
    return { opened, v: workbook.Sheets[currentSheetName].C10?.v, toast: [...document.querySelectorAll(".toast")].map((x) => x.textContent).join("|") };
  });
  check("8. podnagłówek da się edytować (bez komunikatu o blokadzie)", subEdit.opened && subEdit.v === "ok" && !/nie można edytować/.test(subEdit.toast), subEdit);

  // tryb Zawijaj tekst (Widok) wyłącza wszystko
  await p.evaluate(() => { tableEl.classList.add("wrap-cells"); swbSpill.invalidate(); });
  await sleep(200);
  s = await info();
  check("Widok → Zawijaj tekst: nic się nie wylewa", Object.values(s.rows).every((r) => r.every((c) => c === null)), null);

  check("bez błędów strony", errors.length === 0, errors);
  await browser.close();
  if (fails.length) {
    console.error(`\n${fails.length} błędów (${ENGINE})`);
    process.exit(1);
  }
  console.log(`\nOK (${ENGINE})`);
}

run().catch((e) => { console.error(e); process.exit(1); });
