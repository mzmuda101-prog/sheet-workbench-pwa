// cf-dv-slice-playwright.js — czytniki formatowania warunkowego (CF) i walidacji (DV)
// dostają XML arkusza BEZ <sheetData> (sheetXmlWithoutData w core.js), bo parsowanie
// całych danych do DOM tylko po to, by je pominąć, kosztowało ~1 s przy wczytaniu.
//
// Co musi być prawdą: dla KAŻDEGO arkusza każdego pliku wynik parseSheetCF
// i parseSheetDataValidations jest IDENTYCZNY dla pełnego i przyciętego XML-a.
// Pliki: z katalogu scripts/ + opcjonalnie wszystkie .xlsx z XLSX_DIR (np. własne arkusze).

const { chromium } = require("playwright");
const fs = require("fs");
const path = require("path");

const APP_URL = process.env.APP_URL || "http://127.0.0.1:4175/";
const files = fs.readdirSync(__dirname).filter((f) => f.endsWith(".xlsx")).map((f) => path.join(__dirname, f));
if (process.env.XLSX_DIR) {
  for (const f of fs.readdirSync(process.env.XLSX_DIR)) if (f.endsWith(".xlsx") && !f.startsWith("~$")) files.push(path.join(process.env.XLSX_DIR, f));
}

(async () => {
  const browser = await chromium.launch({ headless: true });
  const context = await browser.newContext({ serviceWorkers: "block" });
  await context.addInitScript(() => localStorage.setItem("introPlayed", "true"));
  const page = await context.newPage();
  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(() => ensureXlsxLibs(false));

  let sheets = 0, withCF = 0, withDV = 0;
  const failures = [];
  for (const file of files) {
    const b64 = fs.readFileSync(file).toString("base64");
    const res = await page.evaluate(async (b64) => {
      const bytes = Uint8Array.from(atob(b64), (c) => c.charCodeAt(0));
      const zip = await JSZip.loadAsync(bytes);
      const out = [];
      const strip = (list) => JSON.stringify(list, (k, v) => (k === "_ast" || k === "_values" ? undefined : v));
      for (const name of Object.keys(zip.files)) {
        if (!/^xl\/worksheets\/[^/]+\.xml$/.test(name)) continue;
        const xml = await zip.file(name).async("string");
        const sliced = sheetXmlWithoutData(xml);
        const cfA = strip(parseSheetCF(xml)), cfB = strip(parseSheetCF(sliced));
        const dvA = strip(parseSheetDataValidations(xml)), dvB = strip(parseSheetDataValidations(sliced));
        out.push({ name, cfSame: cfA === cfB, dvSame: dvA === dvB, cf: cfA !== "[]", dv: dvA !== "[]", shrink: +(sliced.length / xml.length).toFixed(3) });
      }
      return out;
    }, b64);
    for (const r of res) {
      sheets += 1;
      if (r.cf) withCF += 1;
      if (r.dv) withDV += 1;
      if (!r.cfSame || !r.dvSame) failures.push(`${path.basename(file)} ${r.name} cf=${r.cfSame} dv=${r.dvSame}`);
    }
  }
  await browser.close();
  console.log(`CF/DV slice: ${files.length} plików, ${sheets} arkuszy (z CF: ${withCF}, z DV: ${withDV})`);
  if (failures.length) {
    console.error("❌ różnice:\n" + failures.join("\n"));
    process.exit(1);
  }
  if (!withCF || !withDV) {
    console.error("❌ korpus nie zawiera arkusza z CF i/lub DV — test niczego by nie sprawdził");
    process.exit(1);
  }
  console.log("✅ wynik identyczny dla pełnego i przyciętego XML");
})().catch((e) => { console.error(e); process.exit(1); });
