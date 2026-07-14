// Weryfikacja regresji na realnym pliku Obieg terenów — drugi zapis nie może psuć XML.
import { readFileSync } from "fs";
import { chromium } from "playwright";

const OK = "/Volumes/KINGSTON128/Arkusze/tereny/Obieg terenów 2026 V5_edited.xlsx";
const APP_URL = process.env.APP_URL || "http://127.0.0.1:4175/";

function hasUtf16SurrogateBytes(arr) {
  for (let i = 0; i < arr.length - 2; i++) {
    if (arr[i] === 0xed && arr[i + 1] >= 0xa0 && arr[i + 1] <= 0xbf && arr[i + 2] >= 0x80) return true;
  }
  return false;
}

const orig = readFileSync(OK);
const browser = await chromium.launch({ headless: true });
const page = await browser.newPage();
await page.goto(APP_URL, { waitUntil: "load" });
await page.waitForFunction(() => typeof buildPatchedXlsx === "function");
await page.evaluate(() => ensureXlsxLibs(false));

const result = await page.evaluate(async ({ bytes, editsList }) => {
  const buildFormulaMaps = (b) => {
    const wb = XLSX.read(new Uint8Array(b), { cellFormula: true, cellStyles: true });
    const maps = {};
    for (const name of wb.SheetNames) {
      const sh = wb.Sheets[name]; const m = {};
      for (const ref in sh) { if (ref[0] === "!" || !sh[ref]?.f) continue; m[ref] = sh[ref].f; }
      if (Object.keys(m).length) maps[name] = m;
    }
    return maps;
  };
  let out = new Uint8Array(bytes);
  for (const edits of editsList) {
    out = await buildPatchedXlsx(out, edits, buildFormulaMaps(out));
  }
  const zip = await JSZip.loadAsync(out);
  const sheetPath = Object.keys(zip.files).find((n) => /xl\/worksheets\/sheet2\.xml$/i.test(n));
  const sxBytes = await zip.file(sheetPath).async("uint8array");
  const sx = new TextDecoder().decode(sxBytes);
  return {
    sheetPath,
    hasSurrogate: (() => {
      for (let i = 0; i < sxBytes.length - 2; i++) {
        if (sxBytes[i] === 0xed && sxBytes[i + 1] >= 0xa0 && sxBytes[i + 1] <= 0xbf && sxBytes[i + 2] >= 0x80) return true;
      }
      return false;
    })(),
    hasH66: /<c r="H66"[^>]*\/>/.test(sx),
    hasShield: sx.includes("🛡") || sx.includes("\uD83D\uDEE1"),
    f66: /Gr8 G\. Choiński/.test(sx),
  };
}, {
  bytes: [...orig],
  editsList: [
    { "2025-2026": { F66: { v: "Gr8 G. Choiński", t: "s" } } },
    { "2025-2026": { F125: { v: "Test125", t: "s" } } },
  ],
});

await browser.close();

const fails = [];
if (result.hasSurrogate) fails.push("UTF-16 surrogate w sheet2.xml");
if (!result.hasH66) fails.push("brak pustej komórki H66");
if (!result.hasShield) fails.push("emoji 🛡️ zniknęło z formuły");
if (!result.f66) fails.push("edycja F66 nie zapisana");

if (fails.length) {
  console.error("❌ real-file regression FAIL:", fails.join("; "));
  process.exit(1);
}
console.log("✅ real-file regression OK (2× zapis na _edited.xlsx, sheet:", result.sheetPath + ")");
