// formula-1900-subtotal-test.mjs — silnik formuł vs Excel (lista 3.11, „dalej”), bez przeglądarki.
//
// 1. Błąd roku 1900 jak w Excelu: seria 60 = 29.02.1900, serie 1–59 przesunięte o dzień, DZIEŃ.TYG
//    z serii. EOMONTH/EDATE w Excelu traktują luty 1900 jako 28 dni.
// 2. SUMY.CZĘŚCIOWE: 1–11 pomijają wiersze odfiltrowane, 101–111 także ukryte ręcznie.
// Wartości oczekiwane z prawdziwego Excela (Mac, AppleScript, 2026-11-06).
import { createRequire } from "node:module";
const require = createRequire(import.meta.url);
const FE = require("../app/formula-engine.js");

const vals = [1, 2, 3, 4, 5, 6]; // B1:B6
const ctx = {
  get: (sh, r, c) => (c === 1 ? vals[r] ?? null : null), formula: () => "", bounds: () => ({ r2: 10, c2: 3 }),
  now: () => 45000, sheetFor: (s) => s, sheetName: "A", row: 0, col: 0, depth: 0, lang: "en",
  rowHidden: (sh, r) => (r === 2 ? 1 : r === 4 ? 2 : 0), // B3 ukryty ręcznie, B5 odfiltrowany
};
const ev = (f) => { try { return FE.evaluate(FE.parse(f, { lang: "en" }), ctx); } catch (e) { return `ERR ${e.message}`; } };
const EXCEL = {
  "DATE(1900,2,29)": 60, "DATE(1900,3,1)": 61, "DATE(1900,1,1)": 1, "DATE(1900,1,60)": 60, "DATE(1900,2,30)": 61,
  "DATE(1899,12,31)": 693962, "DATE(1900,14,1)": 398, 'DATEVALUE("1900-02-28")': 59,
  "WEEKDAY(1)": 1, "WEEKDAY(60)": 4, "WEEKDAY(61)": 5, "WEEKDAY(DATE(1900,1,1),2)": 7, "WEEKNUM(DATE(1900,1,1))": 1,
  "YEAR(0)": 1900, "MONTH(0)": 1, "DAY(0)": 0, "DAY(59)": 28, "MONTH(60)": 2, "DAY(60)": 29, "DAY(61)": 1,
  "EOMONTH(DATE(1900,2,1),0)": 59, "EDATE(DATE(1900,1,31),1)": 59,
  'DATEDIF(DATE(1900,1,1),DATE(1900,3,1),"d")': 60, "NETWORKDAYS(DATE(1900,2,26),DATE(1900,3,2))": 5,
  "DATE(2024,2,29)": 45351, "WEEKDAY(45351)": 5,
  // SUMY.CZĘŚCIOWE (Excel: 9 → z ręcznie ukrytymi 21 bez filtra; tu B5 odfiltrowany)
  "SUBTOTAL(9,B1:B6)": 16, "SUBTOTAL(109,B1:B6)": 13, "SUBTOTAL(102,B1:B6)": 4, "SUBTOTAL(2,B1:B6)": 5, "SUBTOTAL(104,B1:B6)": 6,
};
let bad = 0;
for (const [f, want] of Object.entries(EXCEL)) {
  const got = ev(f);
  if (!(typeof got === "number" && Math.abs(got - want) < 1e-9)) { bad++; console.log(`  ✗ ${f} = ${JSON.stringify(got)}, Excel: ${want}`); }
}
// bez informacji o ukryciu (stary kontekst) — jak dawniej: wszystko liczone
const plain = { ...ctx, rowHidden: undefined };
const g = FE.evaluate(FE.parse("SUBTOTAL(109,B1:B6)", { lang: "en" }), plain);
if (g !== 21) { bad++; console.log(`  ✗ SUBTOTAL bez rowHidden = ${g}, oczekiwano 21`); }
console.log(`${bad ? "❌" : "✅"} formula-1900-subtotal: ${Object.keys(EXCEL).length + 1 - bad} ok, ${bad} błędów`);
process.exit(bad ? 1 : 0);
