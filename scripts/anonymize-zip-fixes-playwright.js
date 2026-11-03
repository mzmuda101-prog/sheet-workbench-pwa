// Strażnik poprawek czyszczenia pliku przy anonimizacji (polowanie na błędy 2026-11-03).
// swbAnonCleanZip wołane wprost na ręcznie złożonej paczce xlsx:
//   1. mailto: w relacjach arkusza zamieniany (dawniej druga gałąź wyrażenia → nic),
//   2. zwykła formuła ze stałą tekstową przepisana RAZ (dawniej drugi raz jak lista walidacji),
//   3. link wewnętrzny do arkusza o nazwie osoby nietknięty (nazw arkuszy nie zmieniamy),
//   4. nazwisko za kodem formatu stopki („&BJan Kowalski”) zamienione,
//   5. po autorze notatki (remember po freezeKnown) znane oryginały dalej zamieniane.
//   node scripts/anonymize-zip-fixes-playwright.js        ENGINE=webkit node scripts/anonymize-zip-fixes-playwright.js
const playwright = require("playwright");
const path = require("path");
const JSZip = require(path.join(__dirname, "..", "lib", "jszip.min.js"));

const APP_URL = process.env.APP_URL || "http://127.0.0.1:4175/";
const ENGINE = process.env.ENGINE || "chromium";

async function buildXlsx() {
  const z = new JSZip();
  z.file("[Content_Types].xml", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/comments1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"/></Types>');
  z.file("_rels/.rels", '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>');
  z.file("xl/workbook.xml", '<?xml version="1.0" encoding="UTF-8"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Lista" sheetId="1" r:id="rId1"/></sheets></workbook>');
  z.file("xl/_rels/workbook.xml.rels", '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>');
  z.file("xl/worksheets/sheet1.xml", '<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData>'
    + '<row r="1"><c r="A1" t="inlineStr"><is><t>Osoba</t></is></c><c r="B1" t="str"><f>IF(A2="Jan Kowalski","tak","nie")</f><v>tak</v></c></row>'
    + '</sheetData><hyperlinks><hyperlink ref="A2" r:id="rId9"/><hyperlink ref="A3" location="\'Jan Kowalski\'!A1" display="Karta"/></hyperlinks>'
    + '<headerFooter><oddFooter>&amp;BJan Kowalski&amp;R&amp;P</oddFooter></headerFooter></worksheet>');
  z.file("xl/worksheets/_rels/sheet1.xml.rels", '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="mailto:jan.kowalski@firma.pl" TargetMode="External"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../comments1.xml"/></Relationships>');
  z.file("xl/comments1.xml", '<?xml version="1.0" encoding="UTF-8"?><comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><authors><author>Ewa Mazur</author></authors><commentList><comment ref="A1" authorId="0"><text><r><t>Zapytać: Remontex Hurt</t></r></text></comment></commentList></comments>');
  return z.generateAsync({ type: "nodebuffer" });
}

async function run() {
  const bytes = [...(await buildXlsx())];
  const browser = await playwright[ENGINE].launch({ headless: true });
  const ctx = await browser.newContext({ serviceWorkers: "block" });
  await ctx.addInitScript(() => localStorage.setItem("introPlayed", "true"));
  const page = await ctx.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(() => { try { ensureXlsxLibs(false); } catch (_) {} });
  await page.waitForFunction(() => typeof SwbAnon !== "undefined" && typeof swbAnonCleanZip === "function" && typeof JSZip !== "undefined", null, { timeout: 20000 });
  const r = await page.evaluate(async (arr) => {
    const s = SwbAnon.createSession({ seed: 5, methods: { person: "pseudo", email: "pseudo" } });
    const jan = s.replaceText("person", "Jan Kowalski", "pseudo");
    s.remember("Jan Kowalski", jan);
    s.remember("Remontex Hurt", "[FIRMA]");
    s.remember("jan.kowalski@firma.pl", "osoba@example.com");
    s.freezeKnown();
    const out = await swbAnonCleanZip(new Uint8Array(arr), s, {}, { personMethod: "pseudo" });
    const z = await JSZip.loadAsync(out.bytes);
    const txt = async (p) => (z.file(p) ? z.file(p).async("string") : "");
    return { jan, sheet: await txt("xl/worksheets/sheet1.xml"), rels: await txt("xl/worksheets/_rels/sheet1.xml.rels"), comments: await txt("xl/comments1.xml") };
  }, bytes);

  const fails = [];
  const check = (name, ok, detail) => { console.log(`${ok ? "✅" : "❌"} ${name}${ok ? "" : ` — ${detail}`}`); if (!ok) fails.push(name); };
  const formula = (r.sheet.match(/<f>([^<]*)<\/f>/) || [])[1] || "";
  check("mailto: zamieniony", !/kowalski/i.test(r.rels) && /mailto:/.test(r.rels), r.rels.match(/Target="mailto:[^"]*"/));
  check("formuła przepisana raz (ta sama zamiana co w komórkach)", formula.includes(`&quot;${r.jan}&quot;`) || formula.includes(`"${r.jan}"`), `${formula} (osoba: ${r.jan})`);
  check("link do arkusza nietknięty", r.sheet.includes(`location="'Jan Kowalski'!A1"`), r.sheet.match(/location="[^"]*"/));
  const footer = (r.sheet.match(/<oddFooter>([^<]*)<\/oddFooter>/) || [])[1] || "";
  check("stopka: nazwisko za „&B” zamienione, kody zostają", !/Kowalski/.test(footer) && footer.startsWith("&amp;B") && footer.includes("&amp;P"), footer);
  check("po autorze notatki znane oryginały dalej zamieniane", !/Remontex/.test(r.comments) && !/Ewa Mazur/.test(r.comments), r.comments.match(/<t>[^<]*<\/t>|<author>[^<]*/g));
  check("bez błędów strony", !errors.length, errors.join(" | "));
  await browser.close();
  if (fails.length) { console.log(`❌ anonymize-zip-fixes (${ENGINE}): ${fails.length} błędów`); process.exit(1); }
  console.log(`✅ anonymize-zip-fixes (${ENGINE}): wszystko OK`);
}
run().catch((e) => { console.error(e); process.exit(1); });
