// xlsx-grow-test.js — zakresy w pliku rosną razem z dopisanymi wierszami (app/xlsx-grow.js).
//   node scripts/xlsx-grow-test.js
//
// Pilnuje (bez przeglądarki, na prawdziwych fixturach):
//   1. rośnie tylko zakres kończący się na ostatnim wierszu danych i mający ≥ 2 wiersze;
//      lista wyboru do wiersza 200, pojedyncza komórka, A:A, zakres kończący się wyżej — bez zmian;
//   2. arkusz: <dimension>, autofiltr, formatowanie warunkowe, listy wyboru, rozszerzenia x14;
//      <sheetData> zostaje bajt w bajt; arkusz z prefiksem przestrzeni nazw (x:);
//   3. tabela Excela: ref i jej autofiltr rosną; z wierszem sumy — bez zmian;
//   4. ścieżki tabel z relacji arkusza;
//   5. kolumny (Etap 1b): zakres kończący się na ostatniej kolumnie danych i mający ≥ 2 kolumny
//      rośnie w bok (także przez Z → AA); jedna kolumna bez zmian; tabela Excela dostaje
//      <tableColumn> z nazwą z arkusza (id = największe + 1, count, znaki & < " escapowane),
//      a bez nazw nie rośnie w bok.

const fs = require("fs");
const path = require("path");
const JSZip = require(path.join(__dirname, "..", "lib", "jszip.min.js")); // ta sama kopia co w aplikacji
const G = require("../app/xlsx-grow.js");

const fails = [];
const check = (name, ok, detail) => {
  console.log(`${ok ? "✅" : "❌"} ${name}${ok ? "" : ` — ${JSON.stringify(detail)}`}`);
  if (!ok) fails.push(name);
};
const eq = (name, got, want) => check(name, got === want, { got, want });

async function run() {
  // 1. pojedyncze zakresy
  eq("1. A3:H42 → A3:H45", G.growRange("A3:H42", 42, 45), "A3:H45");
  eq("1. $A$3:$H$42 → $A$3:$H$45 (zachowuje $)", G.growRange("$A$3:$H$42", 42, 45), "$A$3:$H$45");
  eq("1. lista do 200 bez zmian", G.growRange("D3:D200", 42, 45), "D3:D200");
  eq("1. pojedyncza komórka bez zmian", G.growRange("E42", 42, 45), "E42");
  eq("1. jeden wiersz (A42:H42) bez zmian", G.growRange("A42:H42", 42, 45), "A42:H42");
  eq("1. cała kolumna bez zmian", G.growRange("A:A", 42, 45), "A:A");
  eq("1. kończy się wyżej — bez zmian", G.growRange("A3:H41", 42, 45), "A3:H41");
  eq("1. sqref z kilkoma zakresami", G.growSqref("A3:A42 C3:C10  E5:F42", 42, 44), "A3:A44 C3:C10 E5:F44");

  // 2. arkusz z fixtury
  const zip = await JSZip.loadAsync(fs.readFileSync(path.join(__dirname, "fixtures", "openpyxl-features.xlsx")));
  const sheet = await zip.file("xl/worksheets/sheet1.xml").async("string");
  const grown = G.growSheetXml(sheet, 42, 44);
  const pick = (xml, re) => (xml.match(re) || []).join(" | ");
  eq("2. dimension A1:K42 → A1:K44", pick(grown, /<dimension[^>]*>/g), '<dimension ref="A1:K44" />');
  eq("2. autofiltr A2:H42 → A2:H44", pick(grown, /<autoFilter[^>]*>/g), '<autoFilter ref="A2:H44" />');
  eq("2. formatowanie warunkowe rośnie", pick(grown, /<conditionalFormatting [^>]*>/g),
    pick(sheet, /<conditionalFormatting [^>]*>/g).replace(/42"/g, '44"'));
  eq("2. listy wyboru do wiersza 200 bez zmian", pick(grown, /<dataValidation [^>]*>/g), pick(sheet, /<dataValidation [^>]*>/g));
  const sd = (xml) => xml.slice(xml.indexOf("<sheetData"), xml.indexOf("</sheetData>"));
  check("2. <sheetData> bajt w bajt", sd(grown) === sd(sheet));
  eq("2. bez dopisanych wierszy — bez zmian", G.growSheetXml(sheet, 42, 42), sheet);

  const prefixed = '<x:worksheet xmlns:x="m"><x:dimension ref="A1:C10"/><x:sheetData><x:row r="1"/></x:sheetData>'
    + '<x:autoFilter ref="A1:C10"/><x:conditionalFormatting sqref="B2:B10"><x:cfRule/></x:conditionalFormatting>'
    + '<x:extLst><x:ext><x14:conditionalFormattings><x14:conditionalFormatting><xm:sqref>C2:C10</xm:sqref></x14:conditionalFormatting>'
    + '<x14:dataValidations><x14:dataValidation><xm:sqref>A2:A10 A12</xm:sqref></x14:dataValidation></x14:dataValidations></x14:conditionalFormattings></x:ext></x:extLst></x:worksheet>';
  const pg = G.growSheetXml(prefixed, 10, 12);
  check("2. prefiks x: — dimension, autofiltr, CF", /<x:dimension ref="A1:C12"\/>/.test(pg) && /<x:autoFilter ref="A1:C12"\/>/.test(pg) && /sqref="B2:B12"/.test(pg), pg);
  check("2. rozszerzenia x14 (<xm:sqref>)", /<xm:sqref>C2:C12<\/xm:sqref>/.test(pg) && /<xm:sqref>A2:A12 A12<\/xm:sqref>/.test(pg), pg);
  check("2. pusty arkusz <sheetData/>", G.growSheetXml('<worksheet><dimension ref="A1"/><sheetData/></worksheet>', 1, 3).includes('<dimension ref="A1:A3"/>'));

  // 3. tabele
  const rodo = await JSZip.loadAsync(fs.readFileSync(path.join(__dirname, "RODO Obieg terenów 2026 V5.xlsx")));
  const tpaths = Object.keys(rodo.files).filter((p) => /^xl\/tables\/table\d+\.xml$/.test(p));
  let t1 = "";
  for (const tp of tpaths) { const x = await rodo.file(tp).async("string"); if (/\sref="A2:AG565"/.test(x)) t1 = x; }
  check("3. (warunek) w pliku RODO jest tabela A2:AG565", !!t1, tpaths);
  const tg = G.growTableXml(t1, 565, 567);
  check("3. tabela A2:AG565 → A2:AG567 (ref i autofiltr)", /\sref="A2:AG567"/.test(tg) && /<autoFilter ref="A2:AG567"/.test(tg) && !/A2:AG565/.test(tg), tg.slice(0, 400));
  eq("3. tabela kończąca się wyżej — bez zmian", G.growTableXml(t1, 600, 602), t1);
  const withTotals = t1.replace(/<table /, '<table totalsRowCount="1" ');
  eq("3. tabela z wierszem sumy — bez zmian", G.growTableXml(withTotals, 565, 567), withTotals);

  // 4. relacje
  const sheetPath = "xl/worksheets/sheet1.xml";
  const rels = '<Relationships><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/>'
    + '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://x" TargetMode="External"/>'
    + '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="/xl/tables/table7.xml"/></Relationships>';
  eq("4. ścieżki tabel z relacji", JSON.stringify(G.tablePathsFromRels(rels, sheetPath)), JSON.stringify(["xl/tables/table1.xml", "xl/tables/table7.xml"]));

  // 5. kolumny
  eq("5. A3:H42 w bok → A3:I42", G.growRange("A3:H42", { fromCol: 8, toCol: 9 }), "A3:I42");
  eq("5. w dół i w bok naraz", G.growRange("A3:H42", { fromRow: 42, toRow: 44, fromCol: 8, toCol: 9 }), "A3:I44");
  eq("5. jedna kolumna (E3:E42) bez zmian w bok", G.growRange("E3:E42", { fromCol: 5, toCol: 6 }), "E3:E42");
  eq("5. $ zostają", G.growRange("$A$2:$H$42", { fromCol: 8, toCol: 9 }), "$A$2:$I$42");
  eq("5. Z → AA", G.growRange("A1:Z5", { fromCol: 26, toCol: 27 }), "A1:AA5");
  const cg = G.growSheetXml(sheet, { fromCol: 8, toCol: 9 });
  eq("5. autofiltr A2:H42 → A2:I42", pick(cg, /<autoFilter[^>]*>/g), '<autoFilter ref="A2:I42" />');
  check("5. CF na całe dane rośnie w bok, CF na jedną kolumnę nie",
    /<conditionalFormatting sqref="A3:I42">/.test(cg) && /<conditionalFormatting sqref="E3:E42">/.test(cg), pick(cg, /<conditionalFormatting [^>]*>/g));
  eq("5. dimension sięgająca dalej (K) bez zmian", pick(cg, /<dimension[^>]*>/g), '<dimension ref="A1:K42" />');
  const tcol = G.growTableXml(t1, { fromCol: 33, toCol: 34, tableCols: { opracowanie_terenow: ["Uwagi & <nowe>"] } });
  const ids = [...t1.matchAll(/<tableColumn [^>]*\bid="(\d+)"/g)].map((m) => Number(m[1]));
  const cnt = Number((/<tableColumns count="(\d+)"/.exec(t1) || [])[1]);
  check("5. tabela w bok: ref i autofiltr A2:AH565", /\sref="A2:AH565"/.test(tcol) && /<autoFilter ref="A2:AH565"/.test(tcol), (tcol.match(/ ref="[^"]*"/g) || []).join(" "));
  check("5. tabela w bok: nowy <tableColumn> (id, nazwa z escapowaniem, count)",
    tcol.includes(`<tableColumn id="${Math.max(...ids) + 1}" name="Uwagi &amp; &lt;nowe&gt;"/></tableColumns>`) && tcol.includes(`<tableColumns count="${cnt + 1}"`),
    tcol.slice(tcol.indexOf("</tableColumns>") - 160, tcol.indexOf("</tableColumns>") + 15));
  eq("5. tabela bez podanych nazw — nie rośnie w bok", G.growTableXml(t1, { fromCol: 33, toCol: 34 }), t1);
  const both = G.growTableXml(t1, { fromRow: 565, toRow: 566, fromCol: 33, toCol: 34, tableCols: { opracowanie_terenow: ["X"] } });
  check("5. tabela w dół i w bok naraz: A2:AH566", /\sref="A2:AH566"/.test(both) && /<autoFilter ref="A2:AH566"/.test(both));

  if (fails.length) { console.log(`❌ xlsx-grow: ${fails.length} niezaliczone`); process.exit(1); }
  console.log("✅ xlsx-grow: wszystko przeszło");
}
run().catch((e) => { console.error(e); process.exit(1); });
