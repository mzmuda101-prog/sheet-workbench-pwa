// derived-columns-playwright.js — test KOLUMN WYLICZANYCH (#6) + WYSZUKAJ cross-sheet (#7).
//
// Symuluje realny scenariusz „zamówienia + cennik": użytkownik wczytuje arkusz zamówień
// i chce 1) policzyć kolumnę z wyrażenia, 2) dociągnąć cenę z osobnego arkusza-cennika
// po kodzie (zamiennik VLOOKUP), 3) złożyć obie kolumny w trzecią. Sprawdza, że:
//  (1) [Ilosc]*2 daje poprawne liczby (kolumna z wyrażenia),
//  (2) WYSZUKAJ([Kod];"Cennik";"Kod";"Cena") dociąga ceny, a brak klucza → #N/D,
//  (3) kolumna wyliczana może odwoływać się do innej wyliczanej ([Ilosc]*[Cena]),
//  (4) wyliczone kolumny trafiają do currentHeaders i selecta sortowania (działają w UI),
//  (5) brak błędów konsoli.
//
// Fixture budowany jest W PRZEGLĄDARCE (XLSX strony) → base64 → setInputFiles bufor,
// więc nie wymaga `xlsx` w Node ani pliku na dysku.
//
// Uruchom z serwerem na APP_URL (domyślnie http://127.0.0.1:4175/), np. `npm run serve` obok.

const { chromium } = require("playwright");

const APP_URL = process.env.APP_URL || "http://127.0.0.1:4175/";
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

async function run() {
  const browser = await chromium.launch({ headless: true });
  const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1280, height: 900 } });
  await context.addInitScript(() => {
    localStorage.setItem("introPlayed", "true");
    localStorage.removeItem("excel-workbench-derived-columns"); // czysty start
  });
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push("pageerror: " + e.message));
  page.on("console", (m) => { if (m.type() === "error") errors.push("console: " + m.text()); });

  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  await page.evaluate(() => { try { ensureXlsxLibs && ensureXlsxLibs(false); } catch {} });
  await page.waitForFunction(() => typeof XLSX !== "undefined" && XLSX.utils, null, { timeout: 15000 });

  // Zbuduj skoroszyt z dwoma arkuszami w przeglądarce i zwróć jako base64.
  const b64 = await page.evaluate(() => {
    const wb = XLSX.utils.book_new();
    const zam = XLSX.utils.aoa_to_sheet([
      ["Kod", "Ilosc", "Klient"],
      ["A1", 3, "Jan"],
      ["B2", 5, "Ola"],
      ["C3", 2, "Jan"],
      ["A1", 10, "Ewa"],
      ["Z9", 1, "Tom"], // brak w cenniku → #N/D
    ]);
    const cennik = XLSX.utils.aoa_to_sheet([
      ["Kod", "Cena", "Nazwa"],
      ["A1", 100, "Alfa"],
      ["B2", 50, "Beta"],
      ["C3", 200, "Gamma"],
    ]);
    XLSX.utils.book_append_sheet(wb, zam, "Zamowienia");
    XLSX.utils.book_append_sheet(wb, cennik, "Cennik");
    return XLSX.write(wb, { type: "base64", bookType: "xlsx" });
  });

  await page.setInputFiles("#fileInput", {
    name: "derived-fixture.xlsx",
    mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    buffer: Buffer.from(b64, "base64"),
  });
  await page.waitForFunction(() => document.getElementById("sheetSelect")?.options?.length > 0, null, { timeout: 15000 });
  // wybierz arkusz Zamowienia (pierwszy) i wczytaj
  await page.evaluate(() => { document.getElementById("sheetSelect").value = "Zamowienia"; });
  await sleep(200);
  await page.click("#loadBtn");
  await sleep(800);

  // (1) Kolumna z wyrażenia: [Ilosc]*2
  await page.evaluate(() => {
    document.getElementById("panel-derived-columns").open = true;
    document.getElementById("dcName").value = "Podwojona";
    document.getElementById("dcExpr").value = "[Ilosc]*2";
    document.getElementById("dcAddBtn").click();
  });
  await sleep(300);

  // (2) Kreator WYSZUKAJ: Kod → Cennik.Cena
  await page.evaluate(() => {
    document.getElementById("dcJoinName").value = "Cena";
    document.getElementById("dcJoinKeyCol").value = "Kod";
    const sh = document.getElementById("dcJoinSheet");
    sh.value = "Cennik";
    sh.dispatchEvent(new Event("change"));
    document.getElementById("dcJoinSheetKeyCol").value = "Kod";
    document.getElementById("dcJoinSheetRetCol").value = "Cena";
    document.getElementById("dcJoinBtn").click();
  });
  await sleep(300);

  // (3) Kolumna wyliczana odwołująca się do innej wyliczanej: [Ilosc]*[Cena]
  await page.evaluate(() => {
    document.getElementById("dcName").value = "Razem";
    document.getElementById("dcExpr").value = "[Ilosc]*[Cena]";
    document.getElementById("dcAddBtn").click();
  });
  await sleep(300);

  const out = await page.evaluate(() => {
    const idx = (name) => currentHeaders.indexOf(name);
    const col = (name) => baseRows.map((r) => r.values[idx(name)]);
    const disp = (name) => baseRows.map((r) => r.display[idx(name)]);
    const sortOpts = [...document.getElementById("sortColumnSelect").options].map((o) => o.value);
    return {
      headers: currentHeaders.slice(),
      podwojona: col("Podwojona"),
      cena: col("Cena"),
      cenaDisp: disp("Cena"),
      razem: col("Razem"),
      applied: dcAppliedCount,
      sortHasDerived: ["Podwojona", "Cena", "Razem"].every((n) => sortOpts.includes(n)),
      rowCount: baseRows.length,
    };
  });

  await browser.close();

  const eq = (a, b) => JSON.stringify(a) === JSON.stringify(b);
  const failures = [];
  if (out.rowCount !== 5) failures.push(`expected 5 rows, got ${out.rowCount}`);
  if (!eq(out.headers, ["Kod", "Ilosc", "Klient", "Podwojona", "Cena", "Razem"]))
    failures.push("headers mismatch: " + JSON.stringify(out.headers));
  if (!eq(out.podwojona, [6, 10, 4, 20, 2]))
    failures.push("[Ilosc]*2 wrong: " + JSON.stringify(out.podwojona));
  if (!eq(out.cena, [100, 50, 200, 100, null]))
    failures.push("WYSZUKAJ Cena wrong: " + JSON.stringify(out.cena));
  if (out.cenaDisp[4] !== "#N/D")
    failures.push("missing lookup should display #N/D, got: " + out.cenaDisp[4]);
  if (!eq(out.razem, [300, 250, 400, 1000, 0]))
    failures.push("[Ilosc]*[Cena] wrong: " + JSON.stringify(out.razem));
  if (out.applied !== 3) failures.push(`dcAppliedCount expected 3, got ${out.applied}`);
  if (!out.sortHasDerived) failures.push("derived columns missing from sort select");
  if (errors.length) failures.push(`console/page errors: ${errors.join(" | ")}`);

  console.log(JSON.stringify(out, null, 2));
  if (failures.length) throw new Error(failures.join("; "));
  console.log("✅ derived-columns-playwright OK");
}

run().catch((error) => {
  console.error(error);
  process.exit(1);
});
