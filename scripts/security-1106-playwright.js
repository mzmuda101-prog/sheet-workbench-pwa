// Strażnik poprawek z audytu bezpieczeństwa Sheet (2026-11-06, run-1):
//   1. anonimizacja pliku z prefiksem SpreadsheetML (OpenXML SDK: <x:sst><x:si>, <x:c>, <x:sheet>):
//      arkusze znajdowane, oryginał znika z sharedStrings, test szczelności go widzi (dawniej „czysto”),
//   2. szkic po anonimizacji „w miejscu”: stan czyszczenia przeżywa zapis szkicu i odzyskanie,
//      zapis po odzyskaniu czyści nieużywane teksty,
//   3. arkusz „__proto__”: edycja nie podmienia prototypu pendingEdits,
//   4. CSV: tekst zaczynający się od = + - @ z apostrofem, liczby bez zmian.
//   node scripts/security-1106-playwright.js        ENGINE=webkit node scripts/security-1106-playwright.js
const playwright = require("playwright");
const path = require("path");
const JSZip = require(path.join(__dirname, "..", "lib", "jszip.min.js"));

const APP_URL = process.env.APP_URL || "http://127.0.0.1:4175/";
const ENGINE = process.env.ENGINE || "chromium";
const NS = 'xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"';

async function buildPrefixedXlsx() {
  const z = new JSZip();
  z.file("[Content_Types].xml", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/></Types>');
  z.file("_rels/.rels", '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>');
  z.file("xl/workbook.xml", `<?xml version="1.0" encoding="UTF-8"?><x:workbook ${NS} xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><x:sheets><x:sheet name="Lista" sheetId="1" r:id="rId1"/></x:sheets></x:workbook>`);
  z.file("xl/_rels/workbook.xml.rels", '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/></Relationships>');
  z.file("xl/worksheets/sheet1.xml", `<?xml version="1.0" encoding="UTF-8"?><x:worksheet ${NS}><x:sheetData>`
    + '<x:row r="1"><x:c r="A1" t="s"><x:v>0</x:v></x:c></x:row>'
    + '<x:row r="2"><x:c r="A2" t="s"><x:v>1</x:v></x:c></x:row>'
    + '</x:sheetData></x:worksheet>');
  z.file("xl/sharedStrings.xml", `<?xml version="1.0" encoding="UTF-8"?><x:sst ${NS} count="2" uniqueCount="2"><x:si><x:t>Osoba</x:t></x:si><x:si><x:t>Jan Kowalski</x:t></x:si></x:sst>`);
  return z.generateAsync({ type: "nodebuffer" });
}

async function run() {
  const bytes = [...(await buildPrefixedXlsx())];
  const browser = await playwright[ENGINE].launch({ headless: true });
  const ctx = await browser.newContext({ serviceWorkers: "block" });
  await ctx.addInitScript(() => localStorage.setItem("introPlayed", "true"));
  const page = await ctx.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(() => { try { ensureXlsxLibs(false); } catch (_) {} });
  await page.waitForFunction(() => typeof SwbAnon !== "undefined" && typeof swbAnonCleanZip === "function" && typeof JSZip !== "undefined" && typeof swbAnon !== "undefined", null, { timeout: 20000 });

  const r = await page.evaluate(async (arr) => {
    const src = new Uint8Array(arr);
    const res = {};
    const sstOf = async (b) => (await JSZip.loadAsync(b)).file("xl/sharedStrings.xml").async("string");
    res.paths = await resolveSheetPaths(await JSZip.loadAsync(src));
    const patched = await buildPatchedXlsx(src, { Lista: { A2: { v: "Osoba 001", t: "s" } } }, {});
    const session = () => {
      const s = SwbAnon.createSession({ seed: 5, methods: { person: "pseudo" } });
      s.remember("Jan Kowalski", s.replaceText("person", "Jan Kowalski", "pseudo")); // też leakTerms
      s.freezeKnown();
      return s;
    };
    // 1. kopia: oryginał wyczyszczony z tabeli napisów
    const out = await swbAnonCleanZip(patched, session(), {}, { personMethod: "pseudo" });
    res.copySst = await sstOf(out.bytes);
    res.copyWarnings = out.report.warnings;
    // test szczelności bez czyszczenia napisów — oryginał musi zostać ZGŁOSZONY
    const leaky = await swbAnonCleanZip(patched, session(), { sharedStrings: false }, { personMethod: "pseudo" });
    res.leakTerms = leaky.report.leaks.flatMap((e) => e.items.map((i) => i.term));

    // 2. szkic: stan czyszczenia → structuredClone (IndexedDB) → nowy skoroszyt po odzyskaniu
    const wbBak = workbook;
    try {
      workbook = { Sheets: {}, SheetNames: [] };
      swbAnonPendingCleanup = { session: session(), opts: { ...SWB_ANON_ZIP_DEFAULTS }, workbook, prev: null };
      const saved = structuredClone(swbAnon.draftState());
      swbAnonPendingCleanup = null;
      workbook = { Sheets: {}, SheetNames: [] };
      swbAnon.restoreDraftState(saved);
      res.restored = !!(swbAnonPendingCleanup && swbAnonPendingCleanup.restored && swbAnonPendingCleanup.workbook === workbook);
      res.draftSst = await sstOf(await swbAnon.cleanOnSave(patched));
    } finally {
      swbAnonPendingCleanup = null;
      workbook = wbBak;
    }

    // 3. arkusz „__proto__”
    const peBak = pendingEdits;
    try {
      pendingEdits = {};
      recordPendingEdit("__proto__", "A1", { v: 1, t: "n" });
      res.proto = { inherited: ({}).A1 === undefined && pendingEdits.Inny === undefined && Object.getPrototypeOf(pendingEdits) === Object.prototype, own: Object.keys(pendingEdits), cell: pendingEdits["__proto__"] && pendingEdits["__proto__"].A1 };
    } finally { pendingEdits = peBak; }

    // 4. CSV
    res.csv = ["=HYPERLINK(\"x\")", "@SUMA(A1)", "-12,5", "+48 600", "-", "Jan", "50%"].map((v) => escapeCsvField(v, ";"));
    return res;
  }, bytes);

  const fails = [];
  const check = (name, ok, detail) => { console.log(`${ok ? "✅" : "❌"} ${name}${ok ? "" : ` — ${JSON.stringify(detail)}`}`); if (!ok) fails.push(name); };
  check("prefiks: arkusze znalezione (<x:sheet>)", r.paths.Lista === "xl/worksheets/sheet1.xml", r.paths);
  check("prefiks: oryginał usunięty z sharedStrings", !/Kowalski/.test(r.copySst), r.copySst);
  check("prefiks: bez ostrzeżenia o nieodczytanych arkuszach", !r.copyWarnings.includes("unresolvedSheets"), r.copyWarnings);
  check("prefiks: test szczelności widzi oryginał w sharedStrings", r.leakTerms.some((t) => /Kowalski/.test(t)), r.leakTerms);
  check("szkic: czyszczenie przywrócone dla nowego skoroszytu", r.restored, r.restored);
  check("szkic: zapis po odzyskaniu czyści nieużywane teksty", !/Kowalski/.test(r.draftSst), r.draftSst);
  check("__proto__: bez zmiany prototypu, edycja zapisana", r.proto.inherited && r.proto.own.includes("__proto__") && r.proto.cell && r.proto.cell.v === 1, r.proto);
  check("CSV: formuły z apostrofem, liczby bez zmian", JSON.stringify(r.csv) === JSON.stringify(["'=HYPERLINK(\"\"x\"\")".replace(/^/, '"') + '"', "'@SUMA(A1)", "-12,5", "+48 600", "-", "Jan", "50%"]), r.csv);
  check("bez błędów strony", !errors.length, errors.join(" | "));
  await browser.close();
  if (fails.length) { console.log(`❌ security-1106 (${ENGINE}): ${fails.length} błędów`); process.exit(1); }
  console.log(`✅ security-1106 (${ENGINE}): wszystko OK`);
}
run().catch((e) => { console.error(e); process.exit(1); });
