// save-stress-playwright.js — stress test ścieżki ZAPISU (ZIP-patch, app/xlsx-patch.js).
//
// Po co: zapis metodą ZIP-patch nanosi edycje na surowy .xlsx zachowując tabele/wykresy/style.
// Łatwo tu o plik, który aplikacja czyta luźno, a Excel uzna za uszkodzony ("naprawić?").
// Najczęstsze pułapki: osierocony calcChain po edycji komórki z formułą, niespójne
// [Content_Types].xml / rels, źle escapowane znaki, zepsute XML, zgubione części (tabele).
//
// Test buduje W PRZEGLĄDARCE realistyczny plik (formuła + calcChain + tabela + sharedStrings),
// przepuszcza go przez prawdziwe buildPatchedXlsx() z baterią edycji i waliduje wynik.
// Uruchom z serwerem na APP_URL (domyślnie http://127.0.0.1:4175/), np. `npm run serve` obok.

const { chromium } = require("playwright");

const APP_URL = process.env.APP_URL || "http://127.0.0.1:4175/";

async function run() {
  const browser = await chromium.launch({ headless: true });
  const context = await browser.newContext({ serviceWorkers: "block" });
  await context.addInitScript(() => localStorage.setItem("introPlayed", "true"));
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  page.on("console", (m) => { if (m.type() === "error") errors.push(m.text()); });

  await page.goto(APP_URL, { waitUntil: "load" });
  await page.waitForFunction(() => typeof buildPatchedXlsx === "function" && typeof JSZip !== "undefined");

  const result = await page.evaluate(async () => {
    // ── Zbuduj minimalny, ale realistyczny oryginalny .xlsx (z calcChain + tabelą) ──
    const NS = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"';
    const z = new JSZip();
    z.file("[Content_Types].xml",
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
      `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
      `<Default Extension="xml" ContentType="application/xml"/>` +
      `<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>` +
      `<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>` +
      `<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>` +
      `<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>` +
      `<Override PartName="/xl/tables/table1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>` +
      `<Override PartName="/xl/calcChain.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"/>` +
      `</Types>`);
    z.file("_rels/.rels",
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>` +
      `</Relationships>`);
    z.file("xl/workbook.xml",
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<workbook ${NS} xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
      `<sheets><sheet name="Dane" sheetId="1" r:id="rId1"/></sheets>` +
      `<calcPr calcId="191029"/></workbook>`);
    z.file("xl/_rels/workbook.xml.rels",
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>` +
      `<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>` +
      `<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>` +
      `<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain" Target="calcChain.xml"/>` +
      `</Relationships>`);
    z.file("xl/styles.xml",
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<styleSheet ${NS}><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>` +
      `<fills count="1"><fill><patternFill patternType="none"/></fill></fills>` +
      `<borders count="1"><border/></borders>` +
      `<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>` +
      `<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs></styleSheet>`);
    z.file("xl/sharedStrings.xml",
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<sst ${NS} count="3" uniqueCount="3"><si><t>Nazwa</t></si><si><t>Kwota</t></si><si><t>Razem</t></si></sst>`);
    // Arkusz: nagłówki (shared strings), wiersz z liczbą, komórka z FORMUŁĄ (C2) + tabela
    z.file("xl/worksheets/sheet1.xml",
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<worksheet ${NS} xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
      `<dimension ref="A1:C2"/>` +
      `<sheetData>` +
      `<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c><c r="C1" t="s"><v>2</v></c></row>` +
      `<row r="2"><c r="A2" t="s"><v>0</v></c><c r="B2"><v>10</v></c><c r="C2"><f>B2*2</f><v>20</v></c></row>` +
      `</sheetData>` +
      `<tableParts count="1"><tablePart r:id="rId1"/></tableParts>` +
      `</worksheet>`);
    z.file("xl/worksheets/_rels/sheet1.xml.rels",
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/>` +
      `</Relationships>`);
    z.file("xl/tables/table1.xml",
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<table ${NS} id="1" name="Tabela1" displayName="Tabela1" ref="A1:C2" totalsRowShown="0">` +
      `<autoFilter ref="A1:C2"/>` +
      `<tableColumns count="3"><tableColumn id="1" name="Nazwa"/><tableColumn id="2" name="Kwota"/><tableColumn id="3" name="Razem"/></tableColumns>` +
      `<tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/></table>`);
    z.file("xl/calcChain.xml",
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<calcChain ${NS}><c r="C2" i="1"/></calcChain>`);

    const origBytes = await z.generateAsync({ type: "uint8array", compression: "DEFLATE" });

    // ── Bateria edycji ──
    const edits = { "Dane": {
      "C2": { v: 123, t: "n" },                       // KOMÓRKA Z FORMUŁĄ → wartość (test osieroconego calcChain)
      "B2": { v: 'Tekst & <znaki> "specjalne"', t: "s" }, // escape XML + zmiana typu na inlineStr
      "D5": { v: 7, t: "n" },                          // zupełnie nowa komórka (insert <row>/<c>)
      "A2": null,                                       // usunięcie komórki
    } };

    const patched = await buildPatchedXlsx(origBytes, edits);
    const zp = await JSZip.loadAsync(patched);

    // ── Walidacja ──
    const checks = [];
    const ok = (name, cond) => checks.push({ name, ok: !!cond });

    ok("plik wygenerowany (bajty)", patched && patched.length > 0);
    ok("calcChain.xml usunięty", !zp.file("xl/calcChain.xml"));
    const ct = await zp.file("[Content_Types].xml").async("string");
    ok("Content_Types bez calcChain", !/calcChain\.xml/i.test(ct));
    const rels = await zp.file("xl/_rels/workbook.xml.rels").async("string");
    ok("workbook.rels bez calcChain", !/calcChain\.xml/i.test(rels));

    // wszystkie części XML/rels poprawne składniowo
    const malformed = [];
    for (const nm of Object.keys(zp.files)) {
      if (!/\.(xml|rels)$/i.test(nm)) continue;
      const x = await zp.file(nm).async("string");
      if (new DOMParser().parseFromString(x, "application/xml").getElementsByTagName("parsererror").length) malformed.push(nm);
    }
    ok("wszystkie XML poprawne", malformed.length === 0);

    const sx = await zp.file("xl/worksheets/sheet1.xml").async("string");
    const cellOf = (ref) => { const i = sx.indexOf(`<c r="${ref}"`); return i < 0 ? "" : sx.slice(i, sx.indexOf("</c>", i) + 4); };
    const c2 = cellOf("C2");
    ok("C2: formuła usunięta", c2 && !c2.includes("<f"));
    ok("C2: ma nową wartość 123", c2.includes("<v>123</v>"));
    ok("B2: inlineStr + escape", /t="inlineStr"/.test(cellOf("B2")) && cellOf("B2").includes("&amp;") && cellOf("B2").includes("&lt;znaki&gt;"));
    ok("D5: nowa komórka dodana", cellOf("D5").includes("<v>7</v>"));
    ok("A2: komórka usunięta", !sx.includes('<c r="A2"'));
    ok("tabela zachowana", !!zp.file("xl/tables/table1.xml"));
    ok("styles zachowane", !!zp.file("xl/styles.xml"));
    // liczba arkuszy nie zmieniona (zapis nie dodaje/usuwa arkuszy)
    const sheetParts = (zip) => Object.keys(zip.files).filter((n) => /xl\/worksheets\/sheet\d+\.xml$/i.test(n)).length;
    const zOrig = await JSZip.loadAsync(origBytes);
    ok("liczba arkuszy zachowana", sheetParts(zp) === sheetParts(zOrig));
    const wbXml = await zp.file("xl/workbook.xml").async("string");
    ok("workbook.xml: liczba <sheet> bez zmian", (wbXml.match(/<sheet\b/g) || []).length === 1);

    return { checks, malformed, errorsInPage: [] };
  });

  await browser.close();

  const failures = result.checks.filter((c) => !c.ok).map((c) => c.name);
  if (result.malformed && result.malformed.length) failures.push("zepsute części: " + result.malformed.join(", "));
  if (errors.length) failures.push("błędy konsoli/strony: " + errors.join(" | "));

  console.log("save-stress: " + result.checks.map((c) => `${c.ok ? "✅" : "❌"} ${c.name}`).join("\n             "));
  if (failures.length) {
    console.error("\n❌ save-stress FAIL: " + failures.join("; "));
    process.exit(1);
  }
  console.log("\n✅ save-stress: wszystkie asercje OK");
}

run().catch((error) => { console.error(error); process.exit(1); });
