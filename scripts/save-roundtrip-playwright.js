// save-roundtrip-playwright.js — TWARDY test ścieżki ZAPISU metodą ZIP-patch (app/xlsx-patch.js)
// na PRZERÓŻNYCH konstrukcjach arkusza + wielocyklowym round-tripie "przez samą aplikację".
//
// Po co osobno od save-stress: realny użytkownik robi rzeczy w arkuszu na 1000 sposobów
// (formuły dzielone/tablicowe, tabele, scalenia, daty, błędy, emoji, brak calcChain…),
// a potem: EDYCJA → ZAPIS → wyjście → wejście → wczytanie ZAPISANEGO pliku (którego Excel
// JESZCZE nie "naprawił"/nie przeliczył) → EDYCJA → ZAPIS → i DOPIERO TERAZ otwiera w Excelu.
// Ten test symuluje DOKŁADNIE ten przepływ: każdy cykl czyta bajty z POPRZEDNIEGO zapisu
// (nie oryginał, nie znormalizowane przez Excela), nanosi edycje i zapisuje ponownie.
//
// Walidator jest "excelo-podobny": well-formed XML, brak osieroconych formuł dzielonych (si),
// spójność [Content_Types].xml i rels (brak wiszących części), rosnąca kolejność wierszy/kolumn
// i brak duplikatów komórek, oraz ponowna czytelność pliku przez SheetJS.
//
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
  // xlsx/jszip są dogrywane leniwie (ensureXlsxLibs) — wyzwól je przed testem.
  await page.waitForFunction(() => typeof ensureXlsxLibs === "function" && typeof buildPatchedXlsx === "function");
  await page.evaluate(() => ensureXlsxLibs(false));
  await page.waitForFunction(() => typeof JSZip !== "undefined" && typeof XLSX !== "undefined");

  const out = await page.evaluate(async () => {
    const NS = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"';
    const RNS = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';

    // ── elastyczny builder realistycznego .xlsx ──
    function mk({ sheets, sst, withCalcChain = true, withTable = false, withCalcPr = true, calcChainCells = [], merges = {} }) {
      // sheets: [{ name, xml }]  (xml = zawartość <worksheet>… bez tableParts)
      const z = new JSZip();
      const ov = [`<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>`,
        `<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>`];
      sheets.forEach((s, i) => ov.push(`<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`));
      if (sst) ov.push(`<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>`);
      if (withTable) ov.push(`<Override PartName="/xl/tables/table1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>`);
      if (withCalcChain) ov.push(`<Override PartName="/xl/calcChain.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"/>`);
      z.file("[Content_Types].xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>${ov.join("")}</Types>`);
      z.file("_rels/.rels", `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`);
      const sheetEls = sheets.map((s, i) => `<sheet name="${s.name}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join("");
      z.file("xl/workbook.xml", `<?xml version="1.0"?><workbook ${NS} ${RNS}><sheets>${sheetEls}</sheets>${withCalcPr ? '<calcPr calcId="191029"/>' : ''}</workbook>`);
      const rels = sheets.map((s, i) => `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`);
      let nextId = sheets.length + 1;
      rels.push(`<Relationship Id="rId${nextId++}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>`);
      if (sst) rels.push(`<Relationship Id="rId${nextId++}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>`);
      if (withCalcChain) rels.push(`<Relationship Id="rId${nextId++}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain" Target="calcChain.xml"/>`);
      z.file("xl/_rels/workbook.xml.rels", `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels.join("")}</Relationships>`);
      z.file("xl/styles.xml", `<?xml version="1.0"?><styleSheet ${NS}><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="1"><fill><patternFill patternType="none"/></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs></styleSheet>`);
      if (sst) z.file("xl/sharedStrings.xml", `<?xml version="1.0"?><sst ${NS} count="${sst.length}" uniqueCount="${sst.length}">${sst.map(s => `<si><t xml:space="preserve">${s}</t></si>`).join("")}</sst>`);
      sheets.forEach((s, i) => {
        const tableParts = (withTable && i === 0) ? `<tableParts count="1"><tablePart r:id="rId1"/></tableParts>` : '';
        const mergeXml = merges[s.name] ? `<mergeCells count="${merges[s.name].length}">${merges[s.name].map(r => `<mergeCell ref="${r}"/>`).join("")}</mergeCells>` : '';
        z.file(`xl/worksheets/sheet${i + 1}.xml`, `<?xml version="1.0"?><worksheet ${NS} ${RNS}>${s.xml}${mergeXml}${tableParts}</worksheet>`);
      });
      if (withTable) {
        z.file("xl/worksheets/_rels/sheet1.xml.rels", `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/></Relationships>`);
        z.file("xl/tables/table1.xml", `<?xml version="1.0"?><table ${NS} id="1" name="Tabela1" displayName="Tabela1" ref="A1:C4" totalsRowShown="0"><autoFilter ref="A1:C4"/><tableColumns count="3"><tableColumn id="1" name="A"/><tableColumn id="2" name="B"/><tableColumn id="3" name="C"/></tableColumns></table>`);
      }
      if (withCalcChain) z.file("xl/calcChain.xml", `<?xml version="1.0"?><calcChain ${NS}>${calcChainCells.map(r => `<c r="${r}" i="1"/>`).join("")}</calcChain>`);
      return z.generateAsync({ type: "uint8array", compression: "DEFLATE" });
    }
    const sheet = (xml) => ({ name: "Dane", xml });

    function buildFormulaMaps(bytes) {  // = to, co robi aplikacja (buildSheetFormulaMaps) po wczytaniu
      const wb = XLSX.read(bytes, { cellFormula: true, cellStyles: true });
      const maps = {};
      for (const name of wb.SheetNames) {
        const sh = wb.Sheets[name]; const m = {};
        for (const ref in sh) { if (ref[0] === "!") continue; if (sh[ref] && sh[ref].f) m[ref] = sh[ref].f; }
        if (Object.keys(m).length) maps[name] = m;
      }
      return maps;
    }
    // Round-trip "przez samą aplikację": każdy cykl wczytuje POPRZEDNI zapis (nie oryginał, nie Excel).
    async function appRoundTrip(orig, cycles) {
      let bytes = orig;
      for (const edits of cycles) bytes = await buildPatchedXlsx(bytes, edits, buildFormulaMaps(bytes));
      return bytes;
    }

    // UTF-8 sheet: bajty ED A0–BF 80–BF = CESU-8 / wstawione surogaty UTF-16
    // (Excel: „odzyskaj zawartość"). Lepsze niż osascript — ten dialog i tak mija „open OK".
    function hasUtf16SurrogateBytes(arr) {
      for (let i = 0; i < arr.length - 2; i++) {
        if (arr[i] === 0xed && arr[i + 1] >= 0xa0 && arr[i + 1] <= 0xbf && arr[i + 2] >= 0x80) return true;
      }
      return false;
    }
    function sheetHasShieldEmoji(xml) { return xml.includes("🛡") || xml.includes("\uD83D\uDEE1"); }
    function countProperShieldUtf8(arr) {
      let n = 0;
      for (let i = 0; i < arr.length - 3; i++) {
        if (arr[i] === 0xf0 && arr[i + 1] === 0x9f && arr[i + 2] === 0x9b && arr[i + 3] === 0xa1) n++;
      }
      return n;
    }

    // ── walidator "excelo-podobny" ──
    function normalizePath(base, target) {
      const stack = [];
      for (const p of (base + target).split("/")) { if (!p || p === ".") continue; if (p === "..") stack.pop(); else stack.push(p); }
      return stack.join("/");
    }
    async function validate(bytes) {
      const problems = [];
      const zp = await JSZip.loadAsync(bytes);
      const text = {};
      for (const nm of Object.keys(zp.files)) if (/\.(xml|rels)$/i.test(nm)) text[nm] = await zp.file(nm).async("string");
      for (const nm in text)
        if (new DOMParser().parseFromString(text[nm], "application/xml").getElementsByTagName("parsererror").length) problems.push(`malformed:${nm}`);
      // Content_Types: każdy Override → istniejąca część
      [...(text["[Content_Types].xml"] || "").matchAll(/PartName="([^"]+)"/g)].forEach(m => {
        if (!zp.file(m[1].replace(/^\//, ""))) problems.push(`ct-dangling:${m[1]}`);
      });
      // rels: każdy Target (nie-External) → istniejąca część
      for (const nm in text) {
        if (!/\.rels$/i.test(nm)) continue;
        const base = nm.replace(/_rels\/[^/]+$/i, "");
        [...text[nm].matchAll(/<Relationship\b[^>]*>/g)].forEach(m => {
          const rel = m[0];
          if (/TargetMode="External"/.test(rel)) return;
          const t = (rel.match(/Target="([^"]+)"/) || [])[1]; if (!t) return;
          const resolved = t.startsWith("/") ? t.slice(1) : normalizePath(base, t);
          if (!zp.file(resolved)) problems.push(`rel-dangling:${nm}->${t}`);
        });
      }
      // arkusze: osierocone shared si, kolejność wierszy/kolumn, duplikaty + CESU-8
      for (const nm in text) {
        if (!/xl\/worksheets\/sheet\d+\.xml$/i.test(nm)) continue;
        const sx = text[nm];
        const sxBytes = await zp.file(nm).async("uint8array");
        if (hasUtf16SurrogateBytes(sxBytes)) problems.push(`cesu8-or-surrogate:${nm}`);
        const masters = new Set(), deps = new Set();
        [...sx.matchAll(/<f\b([^>]*?)\/?>/g)].forEach(m => {
          const a = m[1];
          if (!/\bt="shared"/.test(a)) return;
          const si = (a.match(/\bsi="(\d+)"/) || [])[1]; if (si == null) return;
          if (/\bref="/.test(a)) masters.add(si); else deps.add(si);
        });
        deps.forEach(si => { if (!masters.has(si)) problems.push(`orphan-si:${nm}:${si}`); });
        const rows = [...sx.matchAll(/<row\b[^>]*\br="(\d+)"/g)].map(m => +m[1]);
        for (let i = 1; i < rows.length; i++) if (rows[i] <= rows[i - 1]) problems.push(`row-order:${nm}:${rows[i]}`);
        sx.split(/<row\b/).slice(1).forEach(rb => {
          const refs = [...rb.matchAll(/<c r="([A-Z]+\d+)"/g)].map(m => m[1]);
          const seen = new Set(); let prev = -1;
          refs.forEach(r => {
            if (seen.has(r)) problems.push(`dup-cell:${nm}:${r}`);
            seen.add(r);
            const col = XLSX.utils.decode_cell(r).c;
            if (col <= prev) problems.push(`col-order:${nm}:${r}`);
            prev = col;
          });
        });
      }
      try { XLSX.read(bytes, { cellFormula: true, cellStyles: true }); } catch (e) { problems.push("sheetjs-reread:" + e.message); }
      return problems;
    }

    const R = {};

    // KANARZEK: walidator MUSI wykryć celowo osierocone si (inaczej testy są bez wartości)
    {
      const orig = await mk({ sheets: [sheet(`<sheetData><row r="1"><c r="A1"><f t="shared" si="0"/><v>1</v></c></row></sheetData>`)], withCalcChain: false });
      const probs = await validate(orig);
      R["canary-detects-orphan"] = probs.some(p => p.startsWith("orphan-si")) ? [] : ["walidator NIE wykrył osieroconego si!"];
    }

    // 1. formuła tablicowa (CSE) — edycja mastera <f t="array" ref>
    R["array-master-edit"] = await validate(await appRoundTrip(await mk({
      sheets: [sheet(`<dimension ref="A1:C4"/><sheetData>` +
        `<row r="2"><c r="A2"><v>3</v></c><c r="B2"><v>4</v></c><c r="C2"><f t="array" ref="C2:C4">A2:A4*B2:B4</f><v>12</v></c></row>` +
        `<row r="3"><c r="A3"><v>5</v></c><c r="B3"><v>6</v></c><c r="C3"><v>30</v></c></row>` +
        `<row r="4"><c r="A4"><v>7</v></c><c r="B4"><v>8</v></c><c r="C4"><v>56</v></c></row></sheetData>`)],
      calcChainCells: ["C2"],
    }), [{ Dane: { C2: { v: 999, t: "n" } } }]));

    // 2. prostokątny blok shared — usunięcie lewego-górnego mastera (range nie zostaje prostokątem)
    R["shared-rect-master-edit"] = await validate(await appRoundTrip(await mk({
      sheets: [sheet(`<dimension ref="A1:D3"/><sheetData>` +
        `<row r="1"><c r="A1"><v>1</v></c></row>` +
        `<row r="2"><c r="C2"><f t="shared" ref="C2:D3" si="0">A1+1</f><v>2</v></c><c r="D2"><f t="shared" si="0"/><v>2</v></c></row>` +
        `<row r="3"><c r="C3"><f t="shared" si="0"/><v>2</v></c><c r="D3"><f t="shared" si="0"/><v>2</v></c></row></sheetData>`)],
      calcChainCells: ["C2", "D2", "C3", "D3"],
    }), [{ Dane: { C2: { v: 5, t: "n" } } }]));

    // 3. cała grupa shared edytowana naraz (grupa kompletnie się rozpada)
    R["shared-whole-group-collapse"] = await validate(await appRoundTrip(await mk({
      sheets: [sheet(`<dimension ref="A1:B4"/><sheetData>` +
        `<row r="1"><c r="A1"><v>1</v></c></row>` +
        `<row r="2"><c r="B2"><f t="shared" ref="B2:B4" si="0">A1*2</f><v>2</v></c></row>` +
        `<row r="3"><c r="B3"><f t="shared" si="0"/><v>2</v></c></row>` +
        `<row r="4"><c r="B4"><f t="shared" si="0"/><v>2</v></c></row></sheetData>`)],
      calcChainCells: ["B2", "B3", "B4"],
    }), [{ Dane: { B2: { v: 1, t: "n" }, B3: { v: 2, t: "n" }, B4: { v: 3, t: "n" } } }]));

    // 4. WIELOCYKL "edycja→zapis→wejście→edycja→zapis→wejście→edycja" na shared
    R["shared-3cycles-appflow"] = await validate(await appRoundTrip(await mk({
      sheets: [sheet(`<dimension ref="A1:B4"/><sheetData>` +
        `<row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>10</v></c></row>` +
        `<row r="2"><c r="A2"><v>2</v></c><c r="B2"><f t="shared" ref="B2:B4" si="0">A2*10+B1</f><v>30</v></c></row>` +
        `<row r="3"><c r="A3"><v>3</v></c><c r="B3"><f t="shared" si="0"/><v>40</v></c></row>` +
        `<row r="4"><c r="A4"><v>4</v></c><c r="B4"><f t="shared" si="0"/><v>50</v></c></row></sheetData>`)],
      calcChainCells: ["B2", "B3", "B4"],
    }), [
      { Dane: { A1: { v: 111, t: "n" } } },   // cykl 1: zwykłe dane
      { Dane: { B2: { v: 222, t: "n" } } },   // cykl 2: master grupy (de-share)
      { Dane: { B3: { v: 333, t: "n" } } },   // cykl 3: była zależna → teraz samodzielna formuła
    ]));

    // 5. master shared edytowany na TEKST (inlineStr), nie liczbę
    R["shared-master-to-text"] = await validate(await appRoundTrip(await mk({
      sheets: [sheet(`<dimension ref="A1:B3"/><sheetData>` +
        `<row r="1"><c r="A1"><v>1</v></c><c r="B1"><f t="shared" ref="B1:B3" si="0">A1+1</f><v>2</v></c></row>` +
        `<row r="2"><c r="A2"><v>2</v></c><c r="B2"><f t="shared" si="0"/><v>3</v></c></row>` +
        `<row r="3"><c r="A3"><v>3</v></c><c r="B3"><f t="shared" si="0"/><v>4</v></c></row></sheetData>`)],
      calcChainCells: ["B1", "B2", "B3"],
    }), [{ Dane: { B1: { v: "już nie formuła", t: "s" } } }]));

    // 6. znaki specjalne, emoji (pary surogatów), niedozwolony control-char, bardzo długi tekst
    R["special-chars"] = await validate(await appRoundTrip(await mk({
      sheets: [sheet(`<dimension ref="A1:A3"/><sheetData><row r="1"><c r="A1" t="s"><v>0</v></c></row></sheetData>`)],
      sst: ["start"], withCalcChain: false,
    }), [{ Dane: {
      A1: { v: 'Faktura & <VAT> "23%" — €100 😀🇵🇱 ąćęłńóśżź', t: "s" },
      A2: { v: 'znakkontrolny null', t: "s" },
      A3: { v: "x".repeat(40000), t: "s" },
    } }]));

    // 7. różne typy w nietkniętych komórkach (bool/error/inlineStr) muszą przetrwać; data/liczba edytowane
    R["mixed-types-preserved"] = await validate(await appRoundTrip(await mk({
      sheets: [sheet(`<dimension ref="A1:E1"/><sheetData><row r="1">` +
        `<c r="A1" t="b"><v>1</v></c><c r="B1" t="e"><v>#N/A</v></c><c r="C1" t="inlineStr"><is><t>inline</t></is></c>` +
        `<c r="D1" s="1"><v>45000</v></c><c r="E1"><v>3.14159</v></c></row></sheetData>`)],
      withCalcChain: false,
    }), [{ Dane: { D1: { v: new Date(2026, 5, 21), t: "d" }, E1: { v: -2.5, t: "n" } } }]));

    // 8. brak calcChain i calcPr; wstawianie nowych komórek/wierszy poza kolejnością + usuwanie
    R["inserts-deletes-no-calcchain"] = await validate(await appRoundTrip(await mk({
      sheets: [sheet(`<dimension ref="B2:B2"/><sheetData><row r="2"><c r="B2"><v>1</v></c></row></sheetData>`)],
      withCalcChain: false, withCalcPr: false,
    }), [
      { Dane: { A1: { v: 1, t: "n" }, A2: { v: 2, t: "n" }, Z100: { v: 9, t: "n" }, C2: { v: 3, t: "n" } } },
      { Dane: { A1: null, B2: { v: 7, t: "n" } } },
    ]));

    // 9. tabela + autofilter muszą przetrwać edycję komórki shared w tabeli
    {
      const final = await appRoundTrip(await mk({
        sheets: [sheet(`<dimension ref="A1:C4"/><sheetData>` +
          `<row r="1"><c r="A1" t="str"><v>A</v></c><c r="B1" t="str"><v>B</v></c><c r="C1" t="str"><v>C</v></c></row>` +
          `<row r="2"><c r="A2"><v>1</v></c><c r="B2"><v>2</v></c><c r="C2"><f t="shared" ref="C2:C4" si="0">A2+B2</f><v>3</v></c></row>` +
          `<row r="3"><c r="A3"><v>4</v></c><c r="B3"><v>5</v></c><c r="C3"><f t="shared" si="0"/><v>9</v></c></row>` +
          `<row r="4"><c r="A4"><v>6</v></c><c r="B4"><v>7</v></c><c r="C4"><f t="shared" si="0"/><v>13</v></c></row></sheetData>`)],
        withTable: true, calcChainCells: ["C2", "C3", "C4"],
      }), [{ Dane: { C2: { v: 100, t: "n" } } }]);
      const probs = await validate(final);
      if (!(await JSZip.loadAsync(final)).file("xl/tables/table1.xml")) probs.push("table-lost");
      R["table-shared-edit"] = probs;
    }

    // 10. scalone komórki (mergeCells) muszą przetrwać round-trip
    {
      const final = await appRoundTrip(await mk({
        sheets: [sheet(`<dimension ref="A1:C2"/><sheetData>` +
          `<row r="1"><c r="A1" t="inlineStr"><is><t>Tytuł</t></is></c></row>` +
          `<row r="2"><c r="A2"><v>1</v></c><c r="B2"><v>2</v></c></row></sheetData>`)],
        withCalcChain: false, merges: { Dane: ["A1:C1"] },
      }), [{ Dane: { A2: { v: 5, t: "n" } } }]);
      const probs = await validate(final);
      const sx = await (await JSZip.loadAsync(final)).file("xl/worksheets/sheet1.xml").async("string");
      if (!/<mergeCell ref="A1:C1"\/>/.test(sx)) probs.push("merge-lost");
      R["mergecells-preserved"] = probs;
    }

    // 11. DWA arkusze edytowane w jednym zapisie, oba z formułami dzielonymi
    R["multi-sheet-edit"] = await validate(await appRoundTrip(await mk({
      sheets: [
        { name: "Plan", xml: `<dimension ref="A1:B3"/><sheetData>` +
          `<row r="1"><c r="A1"><v>1</v></c><c r="B1"><f t="shared" ref="B1:B3" si="0">A1+1</f><v>2</v></c></row>` +
          `<row r="2"><c r="A2"><v>2</v></c><c r="B2"><f t="shared" si="0"/><v>3</v></c></row>` +
          `<row r="3"><c r="A3"><v>3</v></c><c r="B3"><f t="shared" si="0"/><v>4</v></c></row></sheetData>` },
        { name: "Suma", xml: `<dimension ref="A1:B3"/><sheetData>` +
          `<row r="1"><c r="A1"><v>10</v></c><c r="B1"><f t="shared" ref="B1:B3" si="1">A1*2</f><v>20</v></c></row>` +
          `<row r="2"><c r="A2"><v>20</v></c><c r="B2"><f t="shared" si="1"/><v>40</v></c></row>` +
          `<row r="3"><c r="A3"><v>30</v></c><c r="B3"><f t="shared" si="1"/><v>60</v></c></row></sheetData>` },
      ],
      calcChainCells: [],
    }), [{ Plan: { B1: { v: 100, t: "n" } }, Suma: { A2: { v: 999, t: "n" } } }]));

    // 12. DRUGI zapis (round-trip ×2): emoji w NIEEDYTOWANEJ formule musi przetrwać bez
    // surogatów UTF-16 (regresja: Obieg terenów — XMLSerializer psuł 🛡️ przy 2. patchu).
    {
      const shield = "\uD83D\uDEE1\uFE0F"; // 🛡️
      const emojiFormula = `IF(A1>0,"ok${shield}","")`;
      const orig = await mk({
        sheets: [sheet(
          `<dimension ref="A1:F66"/>` +
          `<dataValidations count="1"><dataValidation type="list" sqref="F66:F100" errorStyle="information"><formula1>"Jan,Ola"</formula1></dataValidation></dataValidations>` +
          `<sheetData>` +
          `<row r="1"><c r="A1"><v>1</v></c><c r="E1" t="str"><f>${emojiFormula.replace(/</g, "&lt;").replace(/>/g, "&gt;")}</f><v>ok ${shield}</v></c></row>` +
          `<row r="66"><c r="F66" s="13" t="s"><v>0</v></c><c r="G66"><v>1</v></c><c r="H66" s="17"/>` +
          `<c r="I66" t="str"><f>${emojiFormula.replace(/</g, "&lt;").replace(/>/g, "&gt;")}</f><v>x</v></c></row>` +
          `</sheetData>`
        )],
        sst: ["Jan"], withCalcChain: false,
      });
      const cycle1 = await buildPatchedXlsx(orig, { Dane: { F212: { v: "Anna Nowak", t: "s" } } }, buildFormulaMaps(orig));
      const cycle2 = await buildPatchedXlsx(cycle1, { Dane: { F66: { v: "Gr8 G. Choiński", t: "s" } } }, buildFormulaMaps(cycle1));
      const probs = await validate(cycle2);
      const sxBytes = await (await JSZip.loadAsync(cycle2)).file("xl/worksheets/sheet1.xml").async("uint8array");
      const outSx = new TextDecoder().decode(sxBytes);
      if (hasUtf16SurrogateBytes(sxBytes)) probs.push("utf16-surrogate-in-xml");
      // Po 2× zapisie liczba poprawnych UTF-8 🛡 nie może spaść (regresja JSZip CESU-8).
      const origBytesSheet = await (await JSZip.loadAsync(orig)).file("xl/worksheets/sheet1.xml").async("uint8array");
      if (countProperShieldUtf8(sxBytes) < countProperShieldUtf8(origBytesSheet)) probs.push("proper-utf8-shield-count-dropped");
      if (!/<c r="H66"[^>]*\/>/.test(outSx)) probs.push("empty-cell-H66-lost");
      const origSx = await (await JSZip.loadAsync(orig)).file("xl/worksheets/sheet1.xml").async("string");
      const origDv = origSx.slice(origSx.indexOf("<dataValidations"), origSx.indexOf("</dataValidations>") + 18);
      const outDv = outSx.slice(outSx.indexOf("<dataValidations"), outSx.indexOf("</dataValidations>") + 18);
      if (origDv !== outDv) probs.push("dataValidations-changed");
      if (!outSx.includes("F66") || !/Gr8 G\. Choiński/.test(outSx)) probs.push("F66-edit-missing");
      R["double-save-emoji-dv"] = probs;
    }

    // 13. Trzeci cykl zapisu na już raz patchowanym pliku (symulacja _edited → _edited_v2).
    {
      const shield = "\uD83D\uDEE1\uFE0F";
      const fXml = `IF(A1&gt;0,"x${shield}","")`;
      const orig = await mk({
        sheets: [sheet(`<dimension ref="E1:E3"/><sheetData>` +
          `<row r="1"><c r="E1" t="str"><f ca="1">${fXml}</f><v>1 ${shield}</v></c></row>` +
          `<row r="2"><c r="E2" t="str"><f ca="1">${fXml}</f><v>2 ${shield}</v></c></row>` +
          `<row r="3"><c r="E3" t="str"><f ca="1">${fXml}</f><v>3 ${shield}</v></c></row></sheetData>`)],
        withCalcChain: false,
      });
      const bytes = await appRoundTrip(orig, [
        { Dane: { A10: { v: "valid1", t: "s" } } },
        { Dane: { A11: { v: "valid2", t: "s" } } },
        { Dane: { A12: { v: "invalid typo", t: "s" } } },
      ]);
      const probs = await validate(bytes);
      const sxBytes = await (await JSZip.loadAsync(bytes)).file("xl/worksheets/sheet1.xml").async("uint8array");
      const outSx = new TextDecoder().decode(sxBytes);
      if (hasUtf16SurrogateBytes(sxBytes)) probs.push("utf16-surrogate-after-3-cycles");
      if (!sheetHasShieldEmoji(outSx)) probs.push("emoji-lost-in-formula");
      R["triple-save-emoji-preserved"] = probs;
    }

    return R;
  });

  await browser.close();

  const lines = Object.keys(out).map(k => `${out[k].length ? "❌" : "✅"} ${k}${out[k].length ? " → " + out[k].join(", ") : ""}`);
  console.log("save-roundtrip:\n  " + lines.join("\n  "));
  const failed = Object.entries(out).filter(([, v]) => v.length).map(([k]) => k);
  if (errors.length) failed.push("błędy konsoli/strony: " + errors.join(" | "));
  if (failed.length) {
    console.error("\n❌ save-roundtrip FAIL: " + failed.join("; "));
    process.exit(1);
  }
  console.log("\n✅ save-roundtrip: wszystkie scenariusze OK (" + lines.length + ")");
}

run().catch((error) => { console.error(error); process.exit(1); });
