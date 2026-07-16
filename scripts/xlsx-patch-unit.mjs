// Szybki test jednostkowy patchSheetXml (bez Playwright) — emoji + pusta komórka + dataValidations.
import { readFileSync } from "fs";
import { createRequire } from "module";
import vm from "vm";

const require = createRequire(import.meta.url);
const patchSrc = readFileSync(new URL("../app/xlsx-patch.js", import.meta.url), "utf8");

const sandbox = {
  module: { exports: {} },
  exports: {},
  DOMParser: null,
  JSZip: null,
  XLSX: { utils: { decode_cell: (ref) => {
    const m = ref.match(/^([A-Z]+)(\d+)$/);
    let c = 0;
    for (const ch of m[1]) c = c * 26 + (ch.charCodeAt(0) - 64);
    return { r: parseInt(m[2], 10) - 1, c: c - 1 };
  } } },
};
vm.runInNewContext(patchSrc + "\nmodule.exports = { patchSheetXml, splitSheetData, indexCells };", sandbox);
const { patchSheetXml, splitSheetData } = sandbox.module.exports;

const shield = "🛡️";
const fEsc = `IF(A1&gt;0,"ok${shield}","")`;
const xml =
  '<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
  '<dataValidations count="1"><dataValidation type="list" sqref="F66:F100" errorStyle="information">' +
  '<formula1>"Jan,Ola"</formula1></dataValidation></dataValidations>' +
  "<sheetData>" +
  `<row r="66"><c r="F66" s="13" t="s"><v>0</v></c><c r="H66" s="17"/>` +
  `<c r="I66" t="str"><f ca="1">${fEsc}</f><v>x ${shield}</v></c></row>` +
  "</sheetData></worksheet>";

const dvBefore = xml.slice(xml.indexOf("<dataValidations"), xml.indexOf("</dataValidations>") + 18);
let out = xml;
for (let i = 0; i < 2; i++) {
  out = patchSheetXml(out, { F66: { v: i === 0 ? "Anna Nowak" : "Gr8 G. Choiński", t: "s" } });
}
const dvAfter = out.slice(out.indexOf("<dataValidations"), out.indexOf("</dataValidations>") + 18);
const hasH66 = /<c r="H66"[^>]*\/>/.test(out);
const hasSurrogate = /\xed[\xa0-\xbf][\x80-\xbf]/.test(out);
const hasShield = out.includes("🛡") || out.includes("\uD83D\uDEE1");

const fails = [];
if (dvBefore !== dvAfter) fails.push("dataValidations changed");
if (!hasH66) fails.push("H66 missing");
if (hasSurrogate) fails.push("UTF-16 surrogate bytes in output");
if (!hasShield) fails.push("emoji lost in untouched formula");
if (!out.includes("Gr8 G. Choiński")) fails.push("F66 edit missing");

if (fails.length) {
  console.error("❌ xlsx-patch-unit FAIL:", fails.join("; "));
  process.exit(1);
}
console.log("✅ xlsx-patch-unit OK (double save, emoji, H66, dataValidations)");

// Regresja 2026-07-16: edycja pustej samozamykającej H (przed formułą I) nie może
// zostawić <c r="H.." s="17"/><v>…</v></c> ani skasować I.
{
  const rowXml =
    '<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>' +
    '<row r="125">' +
    '<c r="F125" s="13" t="s"><v>152</v></c>' +
    '<c r="G125" s="16"><v>46212</v></c>' +
    '<c r="H125" s="17"/>' +
    '<c r="I125" s="8" t="str"><f ca="1">IF(1,"ok","")</f><v>ok</v></c>' +
    '<c r="J125" s="13"/>' +
    "</row></sheetData></worksheet>";
  const patched = patchSheetXml(rowXml, { H125: { v: 46219, t: "n" } });
  const fails2 = [];
  if (/<c r="H125"[^>]*\/>\s*<v>/.test(patched)) fails2.push("selfclose+v orphan after H125");
  if (!/<c r="H125"[^>]*>\s*<v>46219<\/v>\s*<\/c>/.test(patched)) fails2.push("H125 date not written cleanly");
  if (!/<c r="I125"[^>]*>/.test(patched)) fails2.push("I125 formula cell deleted");
  if (!patched.includes('IF(1,"ok","")')) fails2.push("I125 formula text lost");
  if (fails2.length) {
    console.error("❌ xlsx-patch-unit FAIL (selfclose-H-edit):", fails2.join("; "));
    process.exit(1);
  }
  console.log("✅ xlsx-patch-unit OK (self-closing H edit preserves I)");
}