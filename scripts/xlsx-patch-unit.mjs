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
  TextEncoder,
  Date, // ten sam konstruktor co w teście — inaczej `instanceof Date` w xlsx-patch zawodzi (inny realm)
  XLSX: { utils: { decode_cell: (ref) => {
    const m = ref.match(/^([A-Z]+)(\d+)$/);
    let c = 0;
    for (const ch of m[1]) c = c * 26 + (ch.charCodeAt(0) - 64);
    return { r: parseInt(m[2], 10) - 1, c: c - 1 };
  } } },
};
vm.runInNewContext(patchSrc + "\nmodule.exports = { patchSheetXml, splitSheetData, indexCells, zipWriteUtf8Xml, createDateStyler };", sandbox);
const { patchSheetXml, splitSheetData, zipWriteUtf8Xml, createDateStyler } = sandbox.module.exports;

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

// EDYTOWANA komórka z emoji: emoji (poprawne pary UTF-16) muszą przetrwać zapis,
// a samotna połówka pary (bug XMLSerializer) — zniknąć. Dawniej sanitizeXmlText
// wycinał KAŻDĄ połówkę, więc „EMOJI 🔥🎉” zapisywało się jako „EMOJI ”.
{
  const edited = patchSheetXml(xml, {
    F66: { v: "EMOJI 🔥🎉 🛡️", t: "s" },
    H66: { v: "zly\uD83D koniec", t: "s" },
  });
  if (!edited.includes("EMOJI 🔥🎉 🛡️")) fails.push("emoji lost in EDITED cell");
  if (!edited.includes("zly koniec")) fails.push("lone surrogate not stripped");
  if (/[\uD800-\uDBFF](?![\uDC00-\uDFFF])|(?<![\uD800-\uDBFF])[\uDC00-\uDFFF]/.test(edited)) fails.push("lone surrogate left in output");
}

// DATA zapisana do komórki bez formatu daty (pusta / „Ogólny” / tekst „@”) musi dostać
// styl z formatem daty — inaczej Excel pokaże „45424”. Styl już datowy zostaje bez zmian.
{
  const styles = '<styleSheet><numFmts count="1"><numFmt numFmtId="164" formatCode="[$-415]d\\ mmm\\ yy;@"/></numFmts>' +
    '<cellXfs count="3"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' +
    '<xf numFmtId="164" fontId="1" fillId="0" borderId="2" xfId="0" applyNumberFormat="1"/>' +
    '<xf numFmtId="49" fontId="0" fillId="3" borderId="1" xfId="0" applyNumberFormat="1"><alignment horizontal="center"/></xf></cellXfs></styleSheet>';
  const st = createDateStyler(styles);
  const sheetXml = '<worksheet><sheetData><row r="1"><c r="A1" s="2" t="s"><v>0</v></c><c r="B1" s="1"><v>1</v></c><c r="C1"/></row></sheetData></worksheet>';
  const d = (day) => new Date(2024, 4, day);
  const outXml = patchSheetXml(sheetXml, { A1: { v: d(12), t: "d" }, B1: { v: d(13), t: "d" }, C1: { v: d(14), t: "d" } }, null, st);
  const newStyles = st.apply(styles);
  const sOf = (ref) => Number((outXml.match(new RegExp(`<c r="${ref}"[^>]*`))[0].match(/ s="(\d+)"/) || [0, 0])[1]);
  const xfs = newStyles.match(/<xf\b(?:[^>]*?\/>|[^>]*?>[\s\S]*?<\/xf>)/g);
  const fmt = (i) => Number((xfs[i].match(/numFmtId="(\d+)"/) || [])[1]);
  if (!outXml.includes('<v>45424</v>')) fails.push("date serial wrong (want 45424 for 2024-05-12)");
  if (sOf("B1") !== 1) fails.push("already-date style must stay (B1)");
  if (fmt(sOf("A1")) !== 14 || !/fillId="3"/.test(xfs[sOf("A1")]) || !/alignment/.test(xfs[sOf("A1")])) fails.push("A1 (text style) must get date clone keeping fill/alignment");
  if (fmt(sOf("C1")) !== 14) fails.push("C1 (no style) must get date style");
  if (!/<cellXfs count="5">/.test(newStyles)) fails.push("cellXfs count not updated");
  if (/ t="/.test(outXml.match(/<c r="A1"[^>]*/)[0])) fails.push("A1 still has t= attribute");
}

if (fails.length) {
  console.error("❌ xlsx-patch-unit FAIL:", fails.join("; "));
  process.exit(1);
}
console.log("✅ xlsx-patch-unit OK (double save, emoji, emoji w edytowanej komórce, format daty, H66, dataValidations)");

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

// Regresja 2026-07-16: zipWriteUtf8Xml musi pisać bajty TextEncoder (nie JS string → JSZip CESU-8).
{
  const shield = "🛡️";
  const xml = `<?xml version="1.0"?><worksheet><sheetData><c r="E1"><f>IF(1,"ok${shield}","")</f></c></sheetData></worksheet>`;
  const fake = { stored: null, file(_p, data) { this.stored = data; } };
  zipWriteUtf8Xml(fake, "xl/worksheets/sheet1.xml", xml);
  const fails3 = [];
  if (!(fake.stored instanceof Uint8Array)) fails3.push("zipWriteUtf8Xml did not store Uint8Array");
  else {
    const arr = fake.stored;
    let cesu = false, proper = 0;
    for (let i = 0; i < arr.length - 2; i++) {
      if (arr[i] === 0xed && arr[i + 1] >= 0xa0 && arr[i + 1] <= 0xbf && arr[i + 2] >= 0x80) cesu = true;
    }
    for (let i = 0; i < arr.length - 3; i++) {
      if (arr[i] === 0xf0 && arr[i + 1] === 0x9f && arr[i + 2] === 0x9b && arr[i + 3] === 0xa1) proper++;
    }
    if (cesu) fails3.push("CESU-8/surrogate bytes in TextEncoder output");
    if (proper < 1) fails3.push("missing proper UTF-8 shield bytes");
  }
  if (fails3.length) {
    console.error("❌ xlsx-patch-unit FAIL (zipWriteUtf8Xml):", fails3.join("; "));
    process.exit(1);
  }
  console.log("✅ xlsx-patch-unit OK (zipWriteUtf8Xml TextEncoder, no CESU-8)");
}