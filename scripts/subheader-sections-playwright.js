// subheader-sections-playwright.js — WIERSZE SEKCJI („podnagłówki”), przebudowa 2026-10-31.
//   node scripts/subheader-sections-playwright.js
//
// Dawniej: tylko pierwsze 10 wierszy, reguła „≥2 teksty bez liczb” — w RODO wiersze 2–11
// (zwykłe dane osób) były szare i zablokowane, a sekcja w środku tabeli niewidoczna. Pilnuje:
//   1. sekcje w CAŁEJ tabeli: pogrubiona „Dział B” w środku, scalony pasek, etykieta „MATERIAŁY:”;
//   2. NIE sekcje: niepełny rekord (samo nazwisko, bez formatu), wiersz sumy, zwykłe dane;
//   3. tabela z samym tekstem (grafik): nic, nawet przy pogrubionych nazwiskach w każdym wierszu;
//   4. bezpiecznik: wzór powtarzający się w >30% wierszy to dane, nie sekcje;
//   5. druga linia nagłówka na górze tabeli z liczbami — dalej rozpoznana;
//   6. w przeładowanej tabeli wiersz sekcji dostaje szary pasek i DA SIĘ go edytować.
// Uruchom z serwerem na APP_URL (domyślnie http://127.0.0.1:4175/).

const playwright = require("playwright");

const APP_URL = process.env.APP_URL || "http://127.0.0.1:4175/";
const SLEEP_SCALE = Math.min(1, Math.max(0.1, Number(process.env.SLEEP_SCALE || 1)));
const sleep = (ms) => new Promise((r) => setTimeout(r, Math.round(ms * SLEEP_SCALE)));

async function run() {
  const browser = await playwright.chromium.launch({ headless: true });
  const fails = [];
  const check = (name, ok, detail) => {
    console.log(`${ok ? "✅" : "❌"} ${name}${ok ? "" : ` — ${JSON.stringify(detail)}`}`);
    if (!ok) fails.push(name);
  };
  const ctx = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1280, height: 900 } });
  await ctx.addInitScript(() => localStorage.setItem("introPlayed", "true"));
  const p = await ctx.newPage();
  const errors = [];
  p.on("pageerror", (e) => errors.push(e.message));
  await p.goto(APP_URL, { waitUntil: "load" });

  const res = await p.evaluate(() => {
    const B = { bold: true };
    const row = (r, values, styles) => ({ rowIndex0: r, values, cellStyles: styles || [] });
    const flagged = (rows, merges) => markSubheaderRows(rows, merges).filter((x) => x.isSubheader).map((x) => x.rowIndex0);
    const out = {};
    // 1–2: kosztorys z sekcjami
    const data = [];
    let r = 1;
    data.push(row(r++, ["Dział A", null, null, null], [B]));
    for (let i = 0; i < 6; i++) data.push(row(r++, [`Poz ${i}`, "szt", i + 1, (i + 1) * 10]));
    data.push(row(r++, ["Kowalski", null, null, null]));            // niepełny rekord
    data.push(row(r++, ["Razem", null, 21, 210], [B]));               // suma — ma liczby
    data.push(row(r++, ["Dział B", null, null, null], [B]));          // sekcja w środku
    for (let i = 0; i < 6; i++) data.push(row(r++, [`Poz B${i}`, "kg", i + 2, (i + 2) * 5]));
    const mergedRow = r;
    data.push(row(r++, ["Usługi zewnętrzne", null, null, null]));     // scalony pasek
    for (let i = 0; i < 4; i++) data.push(row(r++, [`Usł ${i}`, "h", i + 1, 100]));
    data.push(row(r++, ["Styczeń", null, null, null]));               // etykieta z „ń” (JS \b jej nie widział)
    for (let i = 0; i < 3; i++) data.push(row(r++, [`Sty ${i}`, "szt", i + 1, 3]));
    data.push(row(r++, ["MATERIAŁY:", null, null, null]));            // etykieta
    for (let i = 0; i < 4; i++) data.push(row(r++, [`Mat ${i}`, "m", i + 3, 7]));
    const merges = [{ s: { r: mergedRow, c: 0 }, e: { r: mergedRow, c: 3 } }];
    const f1 = flagged(data, merges);
    out.sections = f1.map((i) => data.find((x) => x.rowIndex0 === i).values[0]);
    // 3: grafik z samym tekstem, pogrubione nazwiska
    const grafik = [];
    for (let i = 0; i < 20; i++) grafik.push(row(i + 1, [`Osoba ${i}`, "Pon: rano", "Wt: noc", null], [B]));
    grafik.push(row(30, ["Tydzień 2", null, null, null], [B]));
    out.grafik = flagged(grafik);
    // 4: bezpiecznik — co drugi wiersz wygląda jak sekcja
    const zebra = [];
    for (let i = 0; i < 20; i++) zebra.push(i % 2 ? row(i + 1, [`x${i}`, 1, 2]) : row(i + 1, [`Grupa ${i}`, null, null], [B]));
    out.zebra = flagged(zebra).length;
    // 5: druga linia nagłówka
    const two = [row(1, ["", "netto", "brutto"]), ...Array.from({ length: 8 }, (_, i) => row(i + 2, [`Poz ${i}`, i * 10, i * 12]))];
    two[0].values = ["(jedn.)", "netto", "brutto"];
    out.two = flagged(two);
    return out;
  });
  check("1. sekcje w całej tabeli: Dział A, Dział B (środek), scalony pasek, Styczeń, MATERIAŁY:",
    JSON.stringify(res.sections) === JSON.stringify(["Dział A", "Dział B", "Usługi zewnętrzne", "Styczeń", "MATERIAŁY:"]), res.sections);
  check("2. niepełny rekord i wiersz sumy NIE są sekcjami", !res.sections.includes("Kowalski") && !res.sections.includes("Razem"), res.sections);
  check("3. grafik z samym tekstem: zwykłe wiersze nie są sekcjami (pogrubiona „Tydzień 2” może)", res.grafik.length <= 1 && !res.grafik.some((i) => i < 30), res.grafik);
  check("4. wzór w >30% wierszy to dane, nie sekcje", res.zebra === 0, res.zebra);
  check("5. druga linia nagłówka u góry tabeli z liczbami", JSON.stringify(res.two) === JSON.stringify([1]), res.two);

  // 6. pełny przebieg: plik z sekcją w środku → szary pasek + edycja
  await p.evaluate(async () => {
    document.getElementById("heroSplash")?.remove();
    await ensureXlsxLibs(true);
    const aoa = [["Pozycja", "Jedn.", "Ilość", "Cena"]];
    for (let i = 0; i < 6; i++) aoa.push([`Poz ${i}`, "szt", i + 1, 10]);
    aoa.push(["DZIAŁ B", null, null, null]);
    for (let i = 0; i < 6; i++) aoa.push([`Poz B${i}`, "kg", i + 2, 5]);
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Kosztorys");
    const buf = XLSX.write(wb, { type: "array", bookType: "xlsx" });
    if (await handleFile(new File([buf], "sekcje.xlsx", { type: XLSX_MIME }))) openDefaultSheet();
  });
  await p.waitForFunction(() => document.querySelectorAll("#dataTable tbody tr[data-row-key]").length >= 13, null, { timeout: 20000 });
  await p.waitForFunction(() => !document.documentElement.classList.contains("vt-scene"), null, { timeout: 10000 });
  await sleep(300);
  // wyróżnianie sekcji jest domyślnie WYŁĄCZONE (2026-10-31) — test włącza je jak użytkownik w Widoku
  check("6. (domyślnie) bez szarego paska sekcji", await p.evaluate(() => !showSubheadersEl.checked && !document.querySelector("tr.row-subheader")), null);
  await p.evaluate(() => { showSubheadersEl.checked = true; showSubheadersEl.dispatchEvent(new Event("change", { bubbles: true })); });
  await sleep(400);
  const s6 = await p.evaluate(() => {
    sheetEditUi.setMode("edit");
    const tr = document.querySelector('#dataTable tbody tr[data-row-index="7"]');
    const grey = tr.classList.contains("row-subheader");
    const td = tr.querySelector('td[data-col-index="0"]');
    openCellEditor(td);
    const inp = td.querySelector("input.cell-editor");
    if (inp) { inp.value = "DZIAŁ C"; inp.dispatchEvent(new KeyboardEvent("keydown", { key: "Enter", bubbles: true })); }
    return { grey, opened: !!inp, v: workbook.Sheets[currentSheetName].A8?.v };
  });
  check("6. „DZIAŁ B” w środku tabeli: szary pasek sekcji", s6.grey, s6);
  check("6. …i da się go edytować", s6.opened && s6.v === "DZIAŁ C", s6);

  check("bez błędów strony", errors.length === 0, errors);
  await browser.close();
  if (fails.length) {
    console.error(`\n${fails.length} błędów`);
    process.exit(1);
  }
  console.log("\nOK");
}

run().catch((e) => { console.error(e); process.exit(1); });
