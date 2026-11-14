// sheet-ops-test.js — nowy skoroszyt i operacje na arkuszach w pliku (app/sheet-ops.js).
//   node scripts/sheet-ops-test.js
//   EXCEL_DIR=/private/tmp/swb-excel-check node scripts/sheet-ops-test.js   (+ kopie do sprawdzenia
//   prawdziwym Excelem: .venv/bin/python -I scripts/excel-real-check.py)
//
// Pilnuje (bez przeglądarki; zachowanie z prawdziwego Excela, Mac PL, 2026-11-14):
//   1. zasady nazw arkuszy, kolejna nazwa „ArkuszN”, apostrofy wokół nazwy w formułach;
//   2. zmiana nazwy w formule: zwykłe i w apostrofach, zakresy, kolumny, nazwy lokalne; tekst
//      w cudzysłowie, inny arkusz o podobnej nazwie i odwołania do innego pliku — bez zmian;
//      usunięcie arkusza: całe odwołanie → #REF!;
//   3. nowy skoroszyt: SheetJS i openpyxl czytają „Arkusz1”, bez autora, bez błędów schematu;
//   4. na fixturze z nazwami, listą wyboru, tabelą, notatkami i ukrytym arkuszem: dodanie (miejsce,
//      sheetId, aktywna zakładka), zmiana nazwy (nazwy zdefiniowane), usunięcie (części tabeli
//      i notatek znikają, nazwy lokalne przesunięte, #REF!), ostatniego widocznego — odmowa;
//      każdy wynik bez NOWYCH błędów schematu (Open XML SDK) i czytany przez openpyxl.

const fs = require("fs");
const os = require("os");
const path = require("path");
const { execFileSync } = require("child_process");
const JSZip = require(path.join(__dirname, "..", "lib", "jszip.min.js"));
global.FormulaEngine = require("../app/formula-engine.js");
const S = require("../app/sheet-ops.js");
let XLSX = null;
try { XLSX = require(path.join(__dirname, "..", "lib", "xlsx.full.min.js")); } catch { XLSX = null; }
let ooxml = null;
try { ooxml = require("./ooxml-validate"); } catch { ooxml = null; }
const PY = path.join(__dirname, "..", ".venv", "bin", "python");

const fails = [];
const check = (name, ok, detail) => {
  console.log(`${ok ? "✅" : "❌"} ${name}${ok ? "" : ` — ${JSON.stringify(detail)}`}`);
  if (!ok) fails.push(name);
};
const eq = (name, got, want) => check(name, got === want, { got, want });
const tmp = fs.mkdtempSync(path.join(os.tmpdir(), "swb-sheetops-"));
const FIX = path.join(__dirname, "fixtures", "openpyxl-features.xlsx");

async function save(zip, name) {
  const buf = await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });
  const file = path.join(tmp, name);
  fs.writeFileSync(file, buf);
  if (process.env.EXCEL_DIR) { fs.mkdirSync(process.env.EXCEL_DIR, { recursive: true }); fs.copyFileSync(file, path.join(process.env.EXCEL_DIR, `sheet-ops-${name}`)); }
  return file;
}
const wbXml = (zip) => zip.file("xl/workbook.xml").async("string");
const sheetNames = (xml) => [...xml.matchAll(/<sheet\b[^>]*\bname="([^"]*)"/g)].map((m) => m[1]);
function schemaErrors(file) {
  if (!ooxml || !ooxml.available()) return null;
  return ooxml.validate([file])[file] || [];
}
function pyRead(file) {
  if (!fs.existsSync(PY)) return null;
  return JSON.parse(execFileSync(PY, ["-I", "-c", `
import json, sys, openpyxl
wb = openpyxl.load_workbook(sys.argv[1])
out = {"sheets": wb.sheetnames, "names": {k: v.attr_text for k, v in wb.defined_names.items()}, "tables": {ws.title: list(ws.tables.keys()) for ws in wb.worksheets}}
print(json.dumps(out, ensure_ascii=False))`, file], { encoding: "utf8" }));
}

async function run() {
  // 1. nazwy
  const names = ["Arkusz1", "Dane"];
  eq("1. pusta nazwa", S.nameProblem("", names), "sheetNameEmpty");
  eq("1. 31 znaków wolno", S.nameProblem("1234567890123456789012345678901", names), null);
  eq("1. 32 znaki — za długa", S.nameProblem("12345678901234567890123456789012", names), "sheetNameTooLong");
  ["a:b", "a/b", "a\\b", "a?b", "a*b", "a[b", "a]b"].forEach((n) => eq(`1. znak „${n}”`, S.nameProblem(n, names), "sheetNameBadChars"));
  eq("1. apostrof na początku", S.nameProblem("'abc", names), "sheetNameApostrophe");
  eq("1. apostrof na końcu", S.nameProblem("abc'", names), "sheetNameApostrophe");
  eq("1. apostrof w środku wolno", S.nameProblem("a'b", names), null);
  eq("1. Historia zarezerwowana", S.nameProblem("historia", names), "sheetNameReserved");
  eq("1. powtórka bez względu na wielkość liter", S.nameProblem("arkusz1", names), "sheetNameTaken");
  eq("1. zmiana wielkości liter własnej nazwy wolno", S.nameProblem("DANE", names, "Dane"), null);
  eq("1. spacje na brzegach wolno", S.nameProblem(" spacja ", names), null);
  eq("1. kolejna nazwa", S.nextName(["Arkusz1", "ąę ß"], "Arkusz"), "Arkusz2");
  eq("1. kolejna nazwa — najniższa wolna", S.nextName(["arkusz2", "Arkusz1", "Arkusz4"], "Arkusz"), "Arkusz3");
  const q = { "Dane_ąę": "Dane_ąę", 2025: "'2025'", AB12: "'AB12'", "Zł": "Zł", "abc.def": "abc.def", R1C1: "'R1C1'", TRUE: "'TRUE'", "a-b": "'a-b'", _x: "_x", "Ąbc": "Ąbc", ABCD12: "ABCD12", x1y: "x1y", "Moje dane": "'Moje dane'", "It's": "'It''s'" };
  Object.entries(q).forEach(([n, want]) => eq(`1. apostrofy jak Excel: ${n}`, S.quoteName(n), want));

  // 2. formuły
  const ren = (from, to) => (sh) => (sh.toUpperCase() === from.toUpperCase() ? { name: to } : null);
  const del = (name) => (sh) => (sh.toUpperCase() === name.toUpperCase() ? { ref: true } : null);
  eq("2. zmiana nazwy (jak w Excelu)", S.rewriteFormula("Arkusz3!A1+SUM(Arkusz3!B2:C3)", ren("Arkusz3", "Moje dane")), "'Moje dane'!A1+SUM('Moje dane'!B2:C3)");
  eq("2. tekst w cudzysłowie bez zmian", S.rewriteFormula('"Arkusz3!A1"&Arkusz3!A1', ren("Arkusz3", "X")), '"Arkusz3!A1"&X!A1');
  eq("2. w apostrofach i wielkość liter", S.rewriteFormula("'arkusz3'!$A$1", ren("Arkusz3", "Nowy")), "Nowy!$A$1");
  eq("2. apostrof w nazwie", S.rewriteFormula("'It''s'!A1", ren("It's", "Its")), "Its!A1");
  eq("2. podobna nazwa bez zmian", S.rewriteFormula("Arkusz30!A1+XArkusz3!A1", ren("Arkusz3", "Y")), "Arkusz30!A1+XArkusz3!A1");
  eq("2. inny plik bez zmian", S.rewriteFormula("[1]Arkusz3!A1+'[2]Arkusz3'!A1", ren("Arkusz3", "Y")), "[1]Arkusz3!A1+'[2]Arkusz3'!A1");
  eq("2. kolumny, wiersze, nazwa lokalna", S.rewriteFormula("SUM(Dane!A:A,Dane!2:3,Dane!Kwoty)", ren("Dane", "D 2")), "SUM('D 2'!A:A,'D 2'!2:3,'D 2'!Kwoty)");
  eq("2. polskie litery bez apostrofów", S.rewriteFormula("Słowniki!B2", ren("Słowniki", "Listy")), "Listy!B2");
  eq("2. usunięcie → #REF! (jak w pliku z Excela)", S.rewriteFormula("Arkusz3!A1+SUM(Arkusz3!B2:C3)", del("Arkusz3")), "#REF!+SUM(#REF!)");
  eq("2. usunięcie: kolumny i nazwa", S.rewriteFormula("COUNTA('Plan 2026'!$A:$A)+Plan!x", del("Plan 2026")), "COUNTA(#REF!)+Plan!x");
  eq("2. XML: escapowanie i <x:f>", S.rewriteXmlFormulas('<x:f t="shared">IF(A&lt;1,Dane!A1,"&amp;")</x:f><f/>', ren("Dane", "R&D")), '<x:f t="shared">IF(A&lt;1,\'R&amp;D\'!A1,"&amp;")</x:f><f/>');

  // 3. nowy skoroszyt
  const nz = await S.newWorkbook(JSZip, { now: new Date("2026-11-14T10:00:00Z") });
  const nf = await save(nz, "nowy.xlsx");
  const core = await nz.file("docProps/core.xml").async("string");
  check("3. bez autora w docProps", !/creator|lastModifiedBy/.test(core), core);
  if (XLSX) {
    const wb = XLSX.read(fs.readFileSync(nf));
    check("3. SheetJS: jeden pusty Arkusz1", wb.SheetNames.join() === "Arkusz1" && Object.keys(wb.Sheets.Arkusz1).filter((k) => k[0] !== "!").length === 0, wb.SheetNames);
  }
  const e0 = schemaErrors(nf);
  if (e0) check("3. bez błędów schematu", e0.length === 0, e0.slice(0, 3));
  const p0 = pyRead(nf);
  if (p0) check("3. openpyxl czyta", p0.sheets.join() === "Arkusz1", p0);

  // 4. fixtura: Dane, Słowniki, Plan 2026 (tabela), Ukryty (ukryty)
  const base = fs.readFileSync(FIX);
  const baseErr = schemaErrors(FIX);
  const noNewErrors = (label, file) => {
    const e = schemaErrors(file);
    if (e) check(`${label}: bez NOWYCH błędów schematu`, ooxml.newErrors(baseErr, e).length === 0, ooxml.newErrors(baseErr, e).slice(0, 3));
  };
  // dodanie za „Dane”
  let z = await JSZip.loadAsync(base);
  eq("4. nowa nazwa", await S.addSheet(z, { name: S.nextName(sheetNames(await wbXml(z)), "Arkusz"), afterName: "Dane" }), "Arkusz1");
  let w = await wbXml(z);
  eq("4. dodany za bieżącym", sheetNames(w).join("|"), "Dane|Arkusz1|Słowniki|Plan 2026|Ukryty");
  check("4. sheetId = największy + 1", /name="Arkusz1" sheetId="5"/.test(w), w.match(/<sheet name="Arkusz1"[^>]*>/));
  check("4. aktywna zakładka = nowy arkusz", /activeTab="1"/.test(w) || !/<workbookView/.test(w), w.match(/<workbookView[^>]*>/));
  check("4. nazwa lokalna _FilterDatabase dalej przy „Dane” (0)", /localSheetId="0"/.test(w), w.match(/<definedName[^>]*localSheetId[^>]*>/));
  const ct = await z.file("[Content_Types].xml").async("string");
  check("4. typ zawartości nowej części", /sheet5\.xml" ContentType="[^"]*worksheet\+xml"/.test(ct), ct.slice(-300));
  let f = await save(z, "dodany.xlsx");
  noNewErrors("4. dodany", f);
  let py = pyRead(f);
  if (py) check("4. openpyxl: dodany arkusz", py.sheets.join("|") === "Dane|Arkusz1|Słowniki|Plan 2026|Ukryty", py.sheets);
  await assertThrows("4. powtórzona nazwa — odmowa", () => S.addSheet(z, { name: "DANE" }));

  // zmiana nazwy „Słowniki” → „Listy wyboru”
  z = await JSZip.loadAsync(base);
  await S.renameSheet(z, { from: "Słowniki", to: "Listy wyboru" });
  w = await wbXml(z);
  eq("4. zmiana nazwy w workbook.xml", sheetNames(w).join("|"), "Dane|Listy wyboru|Plan 2026|Ukryty");
  check("4. nazwa zdefiniowana poszła za arkuszem", w.includes("<definedName name=\"Statusy\">'Listy wyboru'!$B$2:$B$4</definedName>"), w.match(/<definedName name="Statusy"[^<]*/));
  f = await save(z, "zmiana-nazwy.xlsx");
  noNewErrors("4. zmiana nazwy", f);
  py = pyRead(f);
  if (py) check("4. openpyxl: nazwa i nazwa zdefiniowana", py.sheets[1] === "Listy wyboru" && py.names.Statusy === "'Listy wyboru'!$B$2:$B$4", py);

  // usunięcie „Plan 2026” (tabela) — części tabeli znikają
  z = await JSZip.loadAsync(base);
  const next = await S.deleteSheet(z, { name: "Plan 2026" });
  eq("4. po usunięciu aktywny sąsiad z prawej (widoczny)", next, "Słowniki");
  check("4. część tabeli usunięta", !z.file("xl/tables/table1.xml") && !z.file("xl/worksheets/sheet3.xml") && !z.file("xl/worksheets/_rels/sheet3.xml.rels"), Object.keys(z.files));
  const ct2 = await z.file("[Content_Types].xml").async("string");
  check("4. typy zawartości bez usuniętych części", !/table1\.xml|sheet3\.xml/.test(ct2), ct2);
  check("4. notatki innego arkusza zostają", !!z.file("xl/comments/comment1.xml"), Object.keys(z.files));
  f = await save(z, "usuniety-tabela.xlsx");
  noNewErrors("4. usunięty arkusz z tabelą", f);
  py = pyRead(f);
  if (py) check("4. openpyxl: bez arkusza i tabeli", py.sheets.join("|") === "Dane|Słowniki|Ukryty", py);

  // usunięcie „Dane” (notatki, nazwa lokalna 0, nazwa Kwoty → #REF!)
  z = await JSZip.loadAsync(base);
  await S.deleteSheet(z, { name: "Dane" });
  w = await wbXml(z);
  check("4. nazwa lokalna usuniętego arkusza znika", !/_FilterDatabase/.test(w), w.match(/<definedNames>[\s\S]*<\/definedNames>/));
  check("4. nazwa wskazująca usunięty arkusz → #REF!", w.includes('<definedName name="Kwoty">#REF!</definedName>'), w.match(/<definedName name="Kwoty"[^<]*/));
  check("4. notatki i rysunek usuniętego arkusza znikają", !z.file("xl/comments/comment1.xml") && !z.file("xl/drawings/commentsDrawing1.vml"), Object.keys(z.files));
  f = await save(z, "usuniety-pierwszy.xlsx");
  noNewErrors("4. usunięty pierwszy arkusz", f);
  py = pyRead(f);
  if (py) check("4. openpyxl: bez „Dane”", py.sheets.join("|") === "Słowniki|Plan 2026|Ukryty", py);

  // ostatni widoczny arkusz — odmowa
  z = await JSZip.loadAsync(base);
  await S.deleteSheet(z, { name: "Dane" });
  await S.deleteSheet(z, { name: "Słowniki" });
  await assertThrows("4. ostatni widoczny arkusz — odmowa", () => S.deleteSheet(z, { name: "Plan 2026" }), "sheetDeleteLast");

  // nowy skoroszyt: dodaj, zmień nazwę, usuń pierwszy
  z = await S.newWorkbook(JSZip);
  await S.addSheet(z, { name: "Arkusz2", afterName: "Arkusz1" });
  await S.renameSheet(z, { from: "Arkusz2", to: "Koszty 2026" });
  await S.deleteSheet(z, { name: "Arkusz1" });
  eq("4. nowy skoroszyt po operacjach", sheetNames(await wbXml(z)).join("|"), "Koszty 2026");
  const app = await z.file("docProps/app.xml").async("string");
  check("4. docProps/app.xml: lista arkuszy po operacjach", /<vt:i4>1<\/vt:i4>/.test(app) && /<TitlesOfParts><vt:vector size="1" baseType="lpstr"><vt:lpstr>Koszty 2026<\/vt:lpstr><\/vt:vector>/.test(app), app);
  f = await save(z, "nowy-operacje.xlsx");
  const e5 = schemaErrors(f);
  if (e5) check("4. nowy skoroszyt po operacjach: bez błędów schematu", e5.length === 0, e5.slice(0, 3));

  console.log(fails.length ? `\n❌ ${fails.length} nie przeszło` : "\n✅ wszystko przeszło");
  process.exit(fails.length ? 1 : 0);
}

async function assertThrows(name, fn, msg) {
  try { await fn(); check(name, false, "brak błędu"); } catch (e) { check(name, !msg || e.message === msg, e.message); }
}

run().catch((e) => { console.error(e); process.exit(1); });
