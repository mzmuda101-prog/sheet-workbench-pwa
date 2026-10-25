// Strażnik poprawek z polowania na błędy (2026-10-25) — funkcje wołane wprost na stronie.
//   node scripts/bughunt-fixes-playwright.js        ENGINE=webkit node scripts/bughunt-fixes-playwright.js
const playwright = require("playwright");

const APP_URL = process.env.APP_URL || "http://127.0.0.1:4175/";
const ENGINE = process.env.ENGINE || "chromium";

async function run() {
  const browser = await playwright[ENGINE].launch({ headless: true });
  const fails = [];
  const check = (name, ok, detail) => {
    console.log(`${ok ? "✅" : "❌"} ${name}${ok ? "" : ` — ${detail}`}`);
    if (!ok) fails.push(name);
  };

  const ctx = await browser.newContext({ serviceWorkers: "block" });
  await ctx.addInitScript(() => localStorage.setItem("introPlayed", "true"));
  const page = await ctx.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  await page.goto(APP_URL, { waitUntil: "load" });
  await page.waitForFunction(() => typeof FormulaEngine !== "undefined" && typeof cfFunc === "function");
  const r = await page.evaluate(() => {
    const F = FormulaEngine;
    const ev = (src, ctx = {}) => F.evaluate(F.parse(src, ctx.lang ? { lang: ctx.lang } : {}), ctx);
    const out = {};
    // stała tablicowa: „;” = wiersze (EN/plik), „\” = wiersze (PL); zapis bez spłaszczenia
    const arr = F.parse('VLOOKUP(2,{1,"a";2,"b"},2,FALSE)');
    out.arrayRows = ev('VLOOKUP(2,{1,"a";2,"b"},2,FALSE)') === "b" && F.serialize(arr, { file: true }) === 'VLOOKUP(2,{1,"a";2,"b"},2,FALSE)'
      && ev('WYSZUKAJ.PIONOWO(2;{1;"a"\\2;"b"};2;FAŁSZ)', { lang: "pl" }) === "b" ? "" : F.serialize(arr, { file: true });
    // X.DOPASUJ tryb 0 bez symboli wieloznacznych, tryb 2 z nimi
    out.xmatch = ev('XMATCH("a*",{"abc","a*"})') === 2 && ev('XMATCH("a*",{"abc","a*"},2)') === 1 ? "" : "zle";
    // TEKST: kody PL tłumaczone poza cudzysłowem i [ ]
    const seen = [];
    const tctx = { formatNumber: (f) => { seen.push(f); return "ok"; }, now: () => 0 };
    ['rrrr-mm-dd', '[Red]0', '0 "hrs"', 'gg:mm'].forEach((f) => ev(`TEXT(5,"${f.replace(/"/g, '""')}")`, tctx));
    out.textFmt = seen.join("|") === 'yyyy-mm-dd|[Red]0|0 "hrs"|hh:mm' ? "" : seen.join("|");
    // SUMY.CZĘŚCIOWE pomija komórki z własnym SUBTOTAL
    const cells = { "0": 10, "1": 20, "2": 30, "3": 5, "4": 5, "5": 10 };
    const forms = { "2": "SUBTOTAL(9,A1:A2)", "5": "SUBTOTAL(9,A4:A5)" };
    const sctx = { get: (s, rr) => cells[rr] ?? null, bounds: () => ({ r2: 5, c2: 0 }), formula: (s, rr) => forms[rr] || "", now: () => 0 };
    out.subtotal = ev("SUBTOTAL(9,A1:A6)", sctx) === 40 ? "" : String(ev("SUBTOTAL(9,A1:A6)", sctx));
    // MAKS na długiej kolumnie (Chromium: RangeError od ~125 tys.)
    const big = Array.from({ length: 200000 }, (_, i) => i);
    out.bigMax = arrMax(big) === 199999 && arrMin(big) === 0 ? "" : "zle";
    // formatowanie warunkowe: DZIEŃ.TYG jak Excel, DATA.ZA.MIESIĄCE z końcem miesiąca, ZAOKR od zera, Z.WIELKIEJ.LITERY
    const sun = cfFunc("DATE", [2026, 10, 25]);
    const edate = cfSerialToDate(cfFunc("EDATE", [cfFunc("DATE", [2026, 1, 31]), 1])).toISOString().slice(0, 10);
    out.cfDates = sun === 46320 && cfFunc("WEEKDAY", [sun]) === 1 && cfFunc("WEEKDAY", [sun, 2]) === 7 && edate === "2026-02-28" ? "" : `${sun} ${cfFunc("WEEKDAY", [sun])} ${edate}`;
    out.cfRound = cfFunc("ROUND", [-2.5, 0]) === -3 && cfFunc("ROUND", [1.005, 2]) === 1.01 ? "" : String(cfFunc("ROUND", [-2.5, 0]));
    out.cfProper = cfFunc("PROPER", ["łódź ŻÓŁW"]) === "Łódź Żółw" ? "" : cfFunc("PROPER", ["łódź ŻÓŁW"]);
    // seria z długim numerem (IBAN) — kopiowanie zamiast zapisu naukowego
    out.fillLong = swbFillClassify({ v: "PL61109010140000071219812874" }, "pl").kind === "copy" && swbFillClassify({ v: "Nr 7" }, "pl").kind === "textnum" ? "" : "zle";
    // „Pokaż tylko takie” na wartości z „ !” albo cudzysłowem — nie da się wyrazić
    out.cellTerm = buildCellFilterTerm("A", "Uwaga !pilne", false) === null && buildCellFilterTerm("A", 'a"b', false) === null && buildCellFilterTerm("A", "x!", false) === 'A:="x!"' ? "" : "zle";
    // tryb spisywania: liczby zapisane tekstem
    const norm = ["0.125", "1,234.50", "1.234,50", "1.234", "12,50"].map(trNormValue).join("|");
    out.trNorm = norm === "0.125|1234.5|1234.5|1234|12.5" ? "" : norm;
    return out;
  });
  for (const [k, v] of Object.entries(r)) check(k, !v, v);
  check("bez wyjątków na stronie", !errors.length, errors.join(" | "));
  await ctx.close();

  // zablokowana pamięć przeglądarki (Safari „Blokuj wszystkie cookie”) — aplikacja wstaje
  const ctx2 = await browser.newContext({ serviceWorkers: "block" });
  await ctx2.addInitScript(() => {
    Object.defineProperty(window, "localStorage", { get() { throw new DOMException("The operation is insecure.", "SecurityError"); }, configurable: true });
  });
  const p2 = await ctx2.newPage();
  const errs2 = [];
  p2.on("pageerror", (e) => errs2.push(e.message));
  await p2.goto(APP_URL, { waitUntil: "load" });
  await p2.waitForTimeout(800);
  const booted = await p2.evaluate(() => typeof t === "function" && !!document.getElementById("fileInput") && typeof currentLang === "string");
  check("start przy zablokowanym localStorage", booted && !errs2.length, `booted=${booted} ${errs2.join(" | ")}`);
  await ctx2.close();

  await browser.close();
  if (fails.length) { console.error(`❌ bughunt-fixes (${ENGINE}): ${fails.length} błąd(y)`); process.exit(1); }
  console.log(`✅ bughunt-fixes (${ENGINE}) OK`);
}

run().catch((e) => { console.error(e); process.exit(1); });
