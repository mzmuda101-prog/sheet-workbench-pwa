// inventory-xlsx.js — co z Excela używa każdy plik .xlsx (do pytania „czy da się to zrobić od zera
// w aplikacji?”). Tylko odczyt; wynik: cechy na plik + suma. Pliki prywatne zostają na dysku.
//   node tools/od-zera/inventory-xlsx.js plik.xlsx folder/ …
const fs = require("fs");
const path = require("path");
const JSZip = require(path.join(__dirname, "../../lib/jszip.min.js")); // ta sama, którą ma aplikacja

const files = process.argv.slice(2).flatMap((a) => (fs.statSync(a).isDirectory() ? fs.readdirSync(a).filter((f) => /\.xlsx$/i.test(f) && !f.startsWith("~$")).map((f) => path.join(a, f)) : [a]));
const count = (xml, re) => (xml.match(re) || []).length;

async function inspect(file) {
  const zip = await JSZip.loadAsync(fs.readFileSync(file));
  const names = Object.keys(zip.files);
  const read = async (n) => (zip.file(n) ? zip.file(n).async("string") : "");
  const sheetNames = names.filter((n) => /^xl\/worksheets\/sheet\d+\.xml$/.test(n));
  const sheets = (await Promise.all(sheetNames.map(read))).join("");
  const wb = await read("xl/workbook.xml");
  const st = await read("xl/styles.xml");
  const f = {};
  const set = (k, v) => { if (v) f[k] = v; };
  set("arkusze", sheetNames.length > 1 ? sheetNames.length : 0);
  set("ukryte arkusze", count(wb, /<sheet [^>]*state="(hidden|veryHidden)"/g));
  // formuły
  const formulas = (sheets.match(/<f(?: [^>]*)?>([^<]*)<\/f>/g) || []).map((s) => s.replace(/<[^>]*>/g, ""));
  set("formuły", formulas.length);
  const fn = new Set();
  formulas.forEach((x) => (x.match(/\b(?:_xlfn\.)?([A-Z][A-Z0-9.]+)\(/g) || []).forEach((m) => fn.add(m.replace(/_xlfn\.|\(/g, ""))));
  set("funkcje", [...fn].sort().join(", "));
  set("formuły tablicowe / rozlane", count(sheets, /<f [^>]*t="array"/g));
  set("formuły wspólne", count(sheets, /<f [^>]*t="shared"/g));
  set("odwołania do innych arkuszy", formulas.filter((x) => /!/.test(x)).length);
  set("nazwy (zdefiniowane)", count(wb, /<definedName /g) - count(wb, /<definedName [^>]*name="_xlnm\./g));
  // wygląd
  set("scalone komórki", count(sheets, /<mergeCell /g));
  set("formaty liczb (własne)", count(st, /<numFmt /g));
  set("wypełnienia", Math.max(0, count(st, /<fill>/g) - 2));
  set("obramowania", Math.max(0, count(st, /<border>/g) - 1));
  set("czcionki", count(st, /<font>/g) > 1 ? count(st, /<font>/g) : 0);
  set("zawijanie tekstu", count(st, /wrapText="1"/g));
  set("obrót tekstu", count(st, /textRotation="[1-9]/g));
  set("wcięcia w komórce", count(st, /indent="[1-9]/g));
  set("tekst sformatowany w komórce", count(await read("xl/sharedStrings.xml"), /<r>/g));
  set("szerokości kolumn", count(sheets, /<col [^>]*customWidth="1"/g));
  set("wysokości wierszy", count(sheets, /<row [^>]*customHeight="1"/g));
  set("ukryte wiersze / kolumny", count(sheets, /hidden="1"/g));
  set("grupowanie (konspekt)", count(sheets, /outlineLevel="[1-9]"/g));
  // dane
  set("formatowanie warunkowe", count(sheets, /<conditionalFormatting /g));
  const cfTypes = new Set((sheets.match(/<cfRule type="(\w+)"/g) || []).map((s) => s.replace(/.*="|"/g, "")));
  set("rodzaje formatowania warunkowego", [...cfTypes].join(", "));
  set("poprawność danych (listy itp.)", count(sheets, /<dataValidation /g));
  set("filtr", count(sheets, /<autoFilter /g));
  set("tabele (Formatuj jako tabelę)", names.filter((n) => /^xl\/tables\//.test(n)).length);
  set("tabele przestawne", names.filter((n) => /^xl\/pivotTables\//.test(n)).length);
  set("zablokowane okienka", count(sheets, /<pane [^>]*state="frozen"/g));
  set("ochrona arkusza", count(sheets, /<sheetProtection /g));
  set("hiperłącza", count(sheets, /<hyperlink /g));
  set("komentarze / notatki", names.filter((n) => /^xl\/(comments|threadedComments\/)/.test(n)).length);
  // obiekty
  set("wykresy", names.filter((n) => /^xl\/charts\/chart\d+\.xml$/.test(n)).length);
  set("obrazy", names.filter((n) => /^xl\/media\//.test(n)).length);
  set("kształty / rysunki", count((await Promise.all(names.filter((n) => /^xl\/drawings\/drawing\d+\.xml$/.test(n)).map(read))).join(""), /<xdr:sp>/g));
  set("wykresy przebiegu (sparkline)", count(sheets, /<x14:sparklineGroup /g));
  // wydruk
  set("obszar wydruku", count(wb, /name="_xlnm\.Print_Area"/g));
  set("tytuły wydruku", count(wb, /name="_xlnm\.Print_Titles"/g));
  set("ustawienia strony (poziomo / skala)", count(sheets, /<pageSetup [^>]*(orientation="landscape"|scale="|fitToWidth)/g));
  set("nagłówek / stopka wydruku", count(sheets, /<headerFooter/g) && count(sheets, /<(oddHeader|oddFooter)>/g));
  set("podziały stron", count(sheets, /<rowBreaks |<colBreaks /g));
  return f;
}

(async () => {
  const out = {};
  for (const file of files) {
    try { out[path.basename(file)] = await inspect(file); } catch (e) { out[path.basename(file)] = { błąd: String(e.message || e) }; }
  }
  const total = {};
  for (const [name, f] of Object.entries(out)) {
    console.log(`\n## ${name}`);
    for (const [k, v] of Object.entries(f)) { console.log(`  ${k}: ${v}`); total[k] = (total[k] || 0) + 1; }
  }
  console.log(`\n## Ile plików używa (z ${files.length})`);
  Object.entries(total).sort((a, b) => b[1] - a[1]).forEach(([k, n]) => console.log(`  ${n}  ${k}`));
})();
