// anonymize-playwright.js — Anonimizacja danych (anonymize-core/-zip/-ui.js).
//
// Co musi być prawdą:
//   1. Skan całego pliku wykrywa osoby, PESEL, e-mail, telefon, adres (z nagłówków i z treści),
//      także w wolnym tekście („Uwagi”).
//   2. „Nowa kopia”: oryginał w apce bez zmian; w kopii nie ma ŻADNEGO oryginału w żadnej
//      części ZIP (komórki, sharedStrings, wyniki formuł, notatki i autor, właściwości pliku,
//      wykres?), test szczelności w raporcie = 0; plik otwiera SheetJS i openpyxl.
//   3. Fałszywe dane spójne: ta sama osoba w dwóch arkuszach i w formie „A. Kowalski” dostaje
//      tę samą fałszywą osobę; fałszywy PESEL ma poprawną sumę kontrolną.
//   4. Odznaczona wartość zostaje; wyjątek „Nigdy nie ruszaj” też.
//   5. „Zmień otwarty plik”: zmiana w tabeli, JEDNO Cofnij przywraca wszystkie arkusze, Ponów
//      wraca; zapis czyści ślady (sharedStrings).
//   6. Klucz spójności: dwa przebiegi z tym samym hasłem = identyczne zamiany.
//   7. Prawdziwy plik RODO (FILE_RODO): cały plik, osoby → pseudonim; test szczelności 0.
//
// Uruchom z serwerem na APP_URL (domyślnie http://127.0.0.1:4175/), np. `npm run serve:test`.

const { chromium, webkit } = require("playwright");
const fs = require("fs");
const os = require("os");
const path = require("path");
const { execFileSync } = require("child_process");

const APP_URL = process.env.APP_URL || "http://127.0.0.1:4175/";
const ENGINE = process.env.ENGINE === "webkit" ? webkit : chromium;
const SLEEP_SCALE = Math.min(1, Math.max(0.1, Number(process.env.SLEEP_SCALE || 1)));
const sleep = (ms) => new Promise((r) => setTimeout(r, Math.round(ms * SLEEP_SCALE)));
const ROOT = path.join(__dirname, "..");
const XLSX = require(path.join(ROOT, "lib", "xlsx.full.min.js"));
const JSZip = require(path.join(ROOT, "lib", "jszip.min.js"));
const A = require(path.join(ROOT, "app", "anonymize-core.js"));
const FILE_RODO = path.join(__dirname, "RODO Obieg terenów 2026 V5.xlsx");

const PEOPLE = [
  ["Jan Kowalski", "44051401359", "jan.kowalski@firma.pl", "600 123 456", "ul. Lipowa 12/3, 62-200 Gniezno", "Dzwonił Jan Kowalski, prosi o fakturę"],
  ["Anna Nowak", "02270803624", "anna.nowak@poczta.pl", "+48 501 222 333", "ul. Polna 7, 00-950 Warszawa", "Spotkanie z Anną — kontakt 501 222 333"],
  ["Piotr Zieliński", "85010112347", "p.zielinski@example.org", "512-345-678", "al. Jana Pawła 15, 30-001 Kraków", "Brak uwag"],
  ["Maria Wiśniewska", "90020254321", "maria.w@gmail.com", "690 000 111", "os. Słoneczne 4 m. 12, 61-100 Poznań", "Klientka Maria Wiśniewska — VIP"],
];

function buildFixture() {
  const rows = [["Lp", "Imię i nazwisko", "PESEL", "E-mail", "Telefon", "Adres", "Uwagi", "Kwota"]];
  PEOPLE.forEach((p, i) => rows.push([i + 1, p[0], p[1], p[2], p[3], p[4], p[5], 1000 * (i + 1)]));
  const ws = XLSX.utils.aoa_to_sheet(rows);
  ws.B2.c = [{ a: "Ewa Mazur", t: "Sprawdzić u Jan Kowalski" }]; // notatka z autorem
  const ws2 = XLSX.utils.aoa_to_sheet([["Pracownik", "Zlecenia", "Z arkusza"], ["J. Kowalski", 3, ""], ["A. Nowak", 5, ""], ["Kamil Testowy", 1, ""]]);
  ws2.C2 = { t: "s", v: "Jan Kowalski", f: "Klienci!B2" }; // wynik formuły zapamiętany w pliku
  ws2.C3 = { t: "s", v: "Anna Nowak", f: "Klienci!B3" };
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Klienci");
  XLSX.utils.book_append_sheet(wb, ws2, "Podsumowanie");
  wb.Props = { Author: "Mateusz Testowy", LastAuthor: "Mateusz Testowy", Company: "Firma Sp. z o.o.", Title: "Lista klientów Kowalski" };
  const file = path.join(fs.mkdtempSync(path.join(os.tmpdir(), "anon-")), "anon-test.xlsx");
  fs.writeFileSync(file, XLSX.write(wb, { type: "buffer", bookType: "xlsx" }));
  return file;
}

async function zipTexts(bytes) {
  const zip = await JSZip.loadAsync(Buffer.from(bytes));
  const out = {};
  for (const p of Object.keys(zip.files)) if (!zip.files[p].dir && /\.(xml|rels|vml)$/i.test(p)) out[p] = await zip.file(p).async("string");
  return out;
}
const decode = (s) => s.replace(/&amp;/g, "&").replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&apos;/g, "'");
function leaksIn(texts, terms) {
  const found = [];
  for (const [p, x] of Object.entries(texts)) {
    const low = decode(x).toLocaleLowerCase("pl");
    // całe słowo (jak test szczelności w apce): „Nowakowski” to nie „Nowak”
    terms.forEach((t) => { const re = new RegExp(`(?<![\\p{L}\\p{N}])${t.toLocaleLowerCase("pl").replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}(?![\\p{L}\\p{N}])`, "u"); if (re.test(low)) found.push(`${p}: ${t}`); });
  }
  return found;
}

async function openFile(page, file) {
  await page.evaluate(() => { try { ensureXlsxLibs && ensureXlsxLibs(false); } catch {} if (typeof setDirtyState === "function") setDirtyState(false); });
  const name = path.basename(file);
  await page.setInputFiles("#fileInput", file);
  await page.waitForFunction((n) => currentFileName === n && document.getElementById("sheetSelect")?.options?.length > 0, name, { timeout: 30000 });
  await sleep(300);
  await page.evaluate(() => loadBtn.click());
  await page.waitForFunction(() => Array.isArray(baseRows) && baseRows.length > 0, null, { timeout: 20000 });
  await sleep(300);
}

// Okno: zakres → skan; zwraca podsumowanie planu.
async function scanWith(page, scope, setup) {
  await page.evaluate(() => swbAnon.open());
  await page.click(`#anonModal .anon-seg-btn[data-scope="${scope}"]`);
  if (setup) await page.evaluate(setup);
  await page.click("#anonScanBtn");
  await page.waitForFunction(() => swbAnon.state.plan, null, { timeout: 30000 });
  return page.evaluate(() => ({
    changes: swbAnon.state.plan.changes.length,
    types: swbAnon.state.plan.typeCounts,
    cols: swbAnon.state.scan.cols.map((c) => `${c.sheet}|${c.header}|${c.profile.type || ""}`),
  }));
}
async function runCopy(page) {
  await page.evaluate(() => {
    window.__anonSaved = null;
    window.showSaveFilePicker = async (o) => ({ name: o.suggestedName, createWritable: async () => ({ write: async (b) => { window.__anonSaved = Array.from(new Uint8Array(await b.arrayBuffer())); }, close: async () => {} }) });
    document.querySelector('#anonModal input[name="anonOut"][value="copy"]').click();
  });
  await page.click("#anonRunBtn");
  await page.waitForFunction(() => document.getElementById("anonReport") && (window.__anonSaved || swbAnon.lastBytes), null, { timeout: 60000 });
  return page.evaluate(() => ({ bytes: window.__anonSaved || Array.from(swbAnon.lastBytes), report: swbAnon.lastReport, reportText: document.getElementById("anonReport").textContent }));
}

async function run() {
  const FILE = buildFixture();
  const browser = await ENGINE.launch({ headless: true });
  const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1280, height: 900 } });
  await context.addInitScript(() => { localStorage.setItem("introPlayed", "true"); localStorage.setItem("swb-panel-pinned-v1", "1"); localStorage.removeItem("swb-anon-prefs-v1"); });
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push("pageerror: " + e.message));
  page.on("dialog", (d) => d.accept());
  page.on("console", (m) => { if (m.type() === "error") errors.push("console: " + m.text()); });
  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(() => { document.getElementById("heroSplash")?.remove(); });

  const failures = [];
  const ok = (name, cond, got) => { if (!cond) failures.push(`${name} — dostałem: ${JSON.stringify(got)}`); };

  await openFile(page, FILE);

  // ── 1. skan całego pliku ──
  const sc = await scanWith(page, "file");
  ["person", "pesel", "email", "phone", "address"].forEach((t) => ok(`wykryto typ ${t}`, (sc.types[t] || 0) > 0, sc.types));
  ok("kolumna Imię i nazwisko = osoby", sc.cols.includes("Klienci|Imię i nazwisko|person"), sc.cols);
  ok("kolumna Pracownik = osoby", sc.cols.includes("Podsumowanie|Pracownik|person"), sc.cols);
  ok("kolumna Kwota nietknięta", !sc.cols.some((c) => /\|Kwota\|.+/.test(c)), sc.cols);

  // ── 4. odznaczenie wartości + wyjątek ──
  await page.evaluate(() => {
    const key = [...swbAnon.state.plan.values.values()].find((v) => v.value === "Kamil Testowy").key;
    swbAnon.state.excluded.add(key);
    swbAnon.state.methods.person = "fake";
    swbAnon.state.methods.pesel = "fake";
  });
  await page.evaluate(() => { swbAnon.state.plan = swbAnon.buildPlan(); });

  // ── 2. kopia ──
  const before = await page.evaluate(() => workbook.Sheets.Klienci.B2.v);
  const copy = await runCopy(page);
  ok("oryginał w apce bez zmian", (await page.evaluate(() => workbook.Sheets.Klienci.B2.v)) === before, before);
  ok("raport: test szczelności 0", copy.report && copy.report.leaks.length === 0, copy.report && copy.report.leaks);
  const texts = await zipTexts(copy.bytes);
  const originals = PEOPLE.flatMap((p) => [p[0], p[1], p[2], p[4].split(",")[0], "Kowalski", "Nowak", "Wiśniewska"]).concat(["Ewa Mazur", "Mateusz Testowy", "Firma Sp. z o.o."]);
  const leaks = leaksIn(texts, originals);
  ok("w kopii brak oryginałów (każda część ZIP)", leaks.length === 0, leaks.slice(0, 8));
  ok("telefony zniknęły", leaksIn(texts, ["600 123 456", "501 222 333", "512-345-678"]).length === 0, leaksIn(texts, ["600 123 456", "501 222 333"]));
  ok("wyjątek (odznaczona wartość) został", leaksIn(texts, ["Kamil Testowy"]).length > 0, "brak Kamil Testowy");
  ok("raport wymienia wyczyszczone ślady", /właściwości pliku|metadata|wyniki formuł/.test(copy.reportText), copy.reportText.slice(0, 200));

  const wb2 = XLSX.read(Buffer.from(copy.bytes), { type: "buffer" });
  const k = wb2.Sheets.Klienci;
  const fakeJan = k.B2.v;
  ok("fałszywa osoba ≠ oryginał", fakeJan && fakeJan !== "Jan Kowalski" && /^\p{Lu}\p{Ll}+ \p{Lu}\p{Ll}+/u.test(fakeJan), fakeJan);
  const pes = String(k.C2.v);
  ok("fałszywy PESEL z poprawną sumą", A.validators.peselValid(pes) && pes !== PEOPLE[0][1], pes);
  const sum = wb2.Sheets.Podsumowanie;
  ok("„J. Kowalski” = ta sama fałszywa osoba", sum.A2.v === `${fakeJan[0]}. ${fakeJan.split(" ")[1]}`, [sum.A2.v, fakeJan]);
  ok("wynik formuły przepisany spójnie", sum.C2.v === fakeJan, [sum.C2.v, fakeJan]);
  ok("formuła została formułą", sum.C2.f === "Klienci!B2", sum.C2.f);
  ok("kwoty bez zmian", k.H2.v === 1000 && k.H5.v === 4000, [k.H2.v, k.H5.v]);
  ok("uwagi: tekst wokół zostaje", /^Dzwonił .+, prosi o fakturę$/.test(k.G2.v) && !/Kowalski/.test(k.G2.v), k.G2.v);
  const core = texts["docProps/core.xml"] || "";
  ok("autor pliku wyczyszczony", !/Mateusz/.test(core), core.slice(0, 300));
  try {
    const tmp = path.join(os.tmpdir(), `anon-copy-${Date.now()}.xlsx`);
    fs.writeFileSync(tmp, Buffer.from(copy.bytes));
    const py = path.join(ROOT, ".venv", "bin", "python");
    if (fs.existsSync(py)) execFileSync(py, ["-I", "-c", "import openpyxl,sys; wb=openpyxl.load_workbook(sys.argv[1]); print(len(wb.sheetnames))", tmp], { stdio: "pipe" });
  } catch (e) { failures.push(`openpyxl nie otwiera kopii: ${String(e.stderr || e.message).slice(0, 300)}`); }
  await page.evaluate(() => swbAnon.close());

  // ── 6. klucz spójności ──
  const withKey = async () => {
    await scanWith(page, "sheet", () => { swbAnon.state.key = "tajne-hasło"; swbAnon.state.methods.person = "fake"; });
    await page.evaluate(() => { swbAnon.state.plan = swbAnon.buildPlan(); });
    const out = await page.evaluate(() => swbAnon.state.plan.changes.filter((c) => c.type === "person").map((c) => c.after && c.after.value).join("|"));
    await page.evaluate(() => { swbAnon.close(); });
    return out;
  };
  const k1 = await withKey();
  const k2 = await withKey();
  ok("ten sam klucz = te same zamiany", k1 && k1 === k2, [k1, k2]);

  // ── 5. w miejscu + Cofnij/Ponów ──
  await scanWith(page, "file", () => { swbAnon.state.key = ""; swbAnon.state.methods.person = "pseudo"; });
  await page.evaluate(() => { document.querySelector('#anonModal input[name="anonOut"][value="inplace"]').click(); });
  await page.click("#anonRunBtn");
  await page.waitForFunction(() => document.getElementById("anonReport"), null, { timeout: 15000 });
  const inPlace = await page.evaluate(() => ({ b2: workbook.Sheets.Klienci.B2.v, a2: workbook.Sheets.Podsumowanie.A2.v, row: baseRows[0].values[currentHeaders.indexOf("Imię i nazwisko")], pend: Object.keys(pendingEdits.Podsumowanie || {}).length, undo: undoStack.length }));
  ok("w miejscu: arkusz 1 zmieniony", /^Osoba \d{3}$/.test(inPlace.b2), inPlace);
  ok("w miejscu: arkusz 2 zmieniony (ta sama osoba)", inPlace.a2 === inPlace.b2, inPlace);
  ok("w miejscu: tabela pokazuje zmianę", inPlace.row === inPlace.b2, inPlace);
  ok("w miejscu: zmiany do zapisu w obu arkuszach", inPlace.pend > 0, inPlace);
  await page.evaluate(() => swbAnon.close());
  await page.evaluate(() => undoLast());
  const undone = await page.evaluate(() => ({ b2: workbook.Sheets.Klienci.B2.v, a2: workbook.Sheets.Podsumowanie.A2.v, row: baseRows[0].values[currentHeaders.indexOf("Imię i nazwisko")], cleanup: !!swbAnonPendingCleanup }));
  ok("jedno Cofnij przywraca oba arkusze", undone.b2 === "Jan Kowalski" && undone.a2 === "J. Kowalski" && undone.row === "Jan Kowalski", undone);
  ok("Cofnij wyłącza czyszczenie przy zapisie", !undone.cleanup, undone);
  await page.evaluate(() => redoLast());
  const redone = await page.evaluate(() => ({ b2: workbook.Sheets.Klienci.B2.v, cleanup: !!swbAnonPendingCleanup }));
  ok("Ponów wraca anonimizację", /^Osoba/.test(redone.b2) && redone.cleanup, redone);
  const saved = await page.evaluate(async () => Array.from(await buildOutputBytes("xlsx")));
  const savedTexts = await zipTexts(saved);
  const savedLeaks = leaksIn(savedTexts, ["Jan Kowalski", "Anna Nowak", "jan.kowalski@firma.pl", "44051401359", "Ewa Mazur"]);
  ok("zapis po anonimizacji w miejscu bez oryginałów", savedLeaks.length === 0, savedLeaks.slice(0, 6));

  // ── 8. tabele, listy wyboru, nazwy, WYSZUKAJ (scripts/fixtures/anon-tables.xlsx) ──
  // Arkusz „bazy” ma DWIE tabele; dolna („zadania”: mikrofony, nagłośnienie…) zaczyna się
  // pod wykrytym wierszem nagłówka arkusza. Jej nagłówki NIE mogą się zmienić (Excel:
  // „zawartość, która powoduje problem” — zgłoszenie 2026-10-31).
  const FIX_TABLES = path.join(__dirname, "fixtures", "anon-tables.xlsx");
  await openFile(page, FIX_TABLES);
  const tsc = await scanWith(page, "file");
  ok("tabele: kolumna tabeli „zadania” ma własny nagłówek", tsc.cols.some((c) => /^bazy\|mikrofony\|person$/.test(c)), tsc.cols);
  const tc = await runCopy(page);
  const tTexts = await zipTexts(tc.bytes);
  const tWb = XLSX.read(Buffer.from(tc.bytes), { type: "buffer" });
  const bz = tWb.Sheets.bazy;
  ok("tabele: nagłówki dolnej tabeli bez zmian", ["A12", "B12", "C12", "D12"].map((r) => bz[r] && bz[r].v).join("|") === "mikrofony|nagłośnienie|porządkowi|scena", ["A12", "B12", "C12", "D12"].map((r) => bz[r] && bz[r].v));
  ok("tabele: nagłówek górnej tabeli bez zmian", bz.A3.v === "Imię i nazwisko", bz.A3.v);
  for (const [p, x] of Object.entries(tTexts)) {
    if (!/^xl\/tables\//.test(p)) continue;
    const ref = x.match(/<table\b[^>]*\bref="([A-Z]+\d+):/)[1];
    const st0 = XLSX.utils.decode_cell(ref);
    const names = [...x.matchAll(/<tableColumn\b[^>]*\bname="([^"]*)"/g)].map((m) => decode(m[1]));
    const cells = names.map((_, i) => { const c = bz[XLSX.utils.encode_cell({ r: st0.r, c: st0.c + i })]; return c ? String(c.v) : ""; });
    ok(`tabele: ${p} — nazwy kolumn = komórki nagłówka`, names.join("|") === cells.join("|"), [names, cells]);
  }
  ok("tabele: dane w dolnej tabeli zanonimizowane", /^Osoba \d{3}$/.test(bz.A13.v), bz.A13.v);
  ok("tabele: test szczelności 0", tc.report.leaks.length === 0, tc.report.leaks);
  ok("tabele: kontrola powiązań — listy wyboru w porządku", /każda wartość nadal jest na swojej liście/.test(tc.reportText), tc.reportText.slice(-400));
  const gl = tWb.Sheets.Glowne;
  ok("tabele: wartość z listy = wartość ze źródła listy", [bz.A13.v, bz.A14.v, bz.A15.v, bz.A16.v, bz.A17.v].includes(gl.B2.v), [gl.B2.v]);
  const glXml = decode(Object.entries(tTexts).find(([p, x]) => /^xl\/worksheets\/sheet\d+\.xml$/.test(p) && x.includes('r="F2"'))?.[1] || "");
  const f2 = (glXml.match(/<c r="F2"[^>]*><f>([^<]*)<\/f>/) || [])[1] || "";
  ok("tabele: stała w formule zamieniona spójnie", new RegExp(`IF\\(B2="${gl.B2.v}"`).test(f2), [f2, gl.B2.v]);
  const V2 = require("./ooxml-validate.js");
  if (V2.available()) {
    const tmp = path.join(os.tmpdir(), `anon-tables-${Date.now()}.xlsx`);
    fs.writeFileSync(tmp, Buffer.from(tc.bytes));
    const [o2, a2] = V2.validate([FIX_TABLES, tmp]);
    const fresh2 = V2.newErrors(o2, a2);
    ok("tabele: Open XML SDK bez nowych błędów", fresh2.length === 0, fresh2.slice(0, 4).map(V2.short));
  }
  // zabezpieczenie zapisu: rozjechany nagłówek tabeli → nazwa kolumny i odwołanie Tabela[Kolumna] dostosowane
  const synced = await page.evaluate(async (bytes) => {
    const zip = await JSZip.loadAsync(new Uint8Array(bytes));
    const sp = await resolveSheetPaths(zip);
    const p = sp.bazy;
    let x = await zip.file(p).async("string");
    x = x.replace(/(<c r="A3"[^>]*>)[\s\S]*?(<\/c>)/, '<c r="A3" t="inlineStr"><is><t>Osoba</t></is></c>');
    zip.file(p, x);
    const n = await swbAnonSyncTables(zip, sp);
    const t = await zip.file(Object.keys(zip.files).find((f) => /tables\/table\d+\.xml$/.test(f) && /bracia/.test(zip.files[f].name) || false) || "xl/tables/table1.xml").async("string");
    const all = await Promise.all(Object.keys(zip.files).filter((f) => /tables\/table\d+\.xml$/.test(f)).map((f) => zip.file(f).async("string")));
    const gx = await zip.file(sp.Glowne).async("string");
    return { n, cols: all.join("").match(/tableColumn id="1" name="[^"]*"/g), f: (gx.match(/VLOOKUP\([^<]*/) || [""])[0] };
  }, tc.bytes);
  ok("tabele: synchronizacja nazwy kolumny", synced.n === 1 && synced.cols.some((c) => /name="Osoba"/.test(c)), synced);
  ok("tabele: odwołanie strukturalne w formule poprawione", /bracia\[\[Osoba\]:\[Telefon\]\]/.test(decode(synced.f)), synced.f);

  // ── 9. raport: pełne wartości i adresy (zakres = jeden arkusz, reszta poza zakresem) ──
  const sc9 = await scanWith(page, "sheet");
  void sc9;
  const r9 = await runCopy(page);
  ok("raport: lista wycieków z pełnymi wartościami i adresami", r9.report.leaks.some((e) => e.kind === "cells" && e.items.some((it) => /\p{Lu}\p{Ll}+ \p{Lu}/u.test(it.term) && it.refs.length)), r9.report.leaks.slice(0, 2));
  ok("raport: bez skracania „…”", !/\p{L}…\p{L}/u.test(r9.reportText.split("Test szczelności")[1] || ""), r9.reportText.slice(0, 300));

  // ── 7. prawdziwy plik RODO ──
  if (fs.existsSync(FILE_RODO) && !process.env.SKIP_RODO) {
    await openFile(page, FILE_RODO);
    const rs = await scanWith(page, "file", () => { swbAnon.state.methods.person = "pseudo"; });
    ok("RODO: wykryto osoby", (rs.types.person || 0) > 100, rs.types);
    const rc = await runCopy(page);
    const rl = rc.report ? rc.report.leaks : null;
    ok("RODO: test szczelności 0", rl && rl.length === 0, rl && rl.slice(0, 6));
    const rt = await zipTexts(rc.bytes);
    ok("RODO: brak „Adam Kamiński”, „Pilipczuk”, „Maślankiewicz”", leaksIn(rt, ["Adam Kamiński", "Pilipczuk", "Maślankiewicz", "Bielak"]).length === 0, leaksIn(rt, ["Adam Kamiński", "Pilipczuk", "Maślankiewicz", "Bielak"]).slice(0, 6));
    // Excel „naprawia” plik, gdy pozycje pola w pamięci tabeli przestawnej się powtarzają
    // (Open XML SDK tego NIE łapie — sprawdzone w prawdziwym Excelu 2026-10-31)
    const dupes = [];
    Object.entries(rt).filter(([p]) => /pivotCacheDefinition/.test(p)).forEach(([p, x]) => {
      for (const m of x.matchAll(/<cacheField\b[^>]*name="([^"]*)"[^>]*>([\s\S]*?)<\/cacheField>/g)) {
        const vals = [...m[2].matchAll(/<s v="([^"]*)"/g)].map((v) => decode(v[1]).toLowerCase());
        if (new Set(vals).size !== vals.length) dupes.push(`${p}:${m[1]}`);
      }
    });
    ok("RODO: pozycje tabeli przestawnej unikalne", dupes.length === 0, dupes);
    // fałszywe dane: fałszywe imię/nazwisko nie może wyglądać jak prawdziwa osoba z pliku
    await page.evaluate(() => swbAnon.close());
    await scanWith(page, "file", () => { swbAnon.state.methods.person = "fake"; });
    const rf = await runCopy(page);
    ok("RODO fałszywe dane: test szczelności 0", rf.report.leaks.length === 0, rf.report.leaks.slice(0, 3));
    const rwb = XLSX.read(Buffer.from(rc.bytes), { type: "buffer" });
    ok("RODO: kopia czyta się (wszystkie arkusze)", rwb.SheetNames.length >= 7, rwb.SheetNames);
    try {
      const tmp = path.join(os.tmpdir(), `anon-rodo-${Date.now()}.xlsx`);
      fs.writeFileSync(tmp, Buffer.from(rc.bytes));
      const py = path.join(ROOT, ".venv", "bin", "python");
      if (fs.existsSync(py)) execFileSync(py, ["-I", "-c", "import openpyxl,sys; openpyxl.load_workbook(sys.argv[1])", tmp], { stdio: "pipe" });
      if (process.env.KEEP_OUT) fs.copyFileSync(tmp, process.env.KEEP_OUT);
      // Open XML SDK (jak Excel): żadnych NOWYCH błędów względem oryginału
      const V = require("./ooxml-validate.js");
      if (V.available()) {
        const [orig, anon] = V.validate([FILE_RODO, tmp]);
        const fresh = V.newErrors(orig, anon);
        ok("RODO: Open XML SDK bez nowych błędów", fresh.length === 0, fresh.slice(0, 4).map(V.short));
      }
    } catch (e) { failures.push(`RODO: openpyxl nie otwiera kopii: ${String(e.stderr || e.message).slice(0, 300)}`); }
  }

  const real = errors.filter((e) => !/favicon|Failed to load resource/.test(e));
  ok("brak błędów w konsoli", real.length === 0, real.slice(0, 5));
  await browser.close();
  if (failures.length) {
    console.log(`❌ anonymize (${process.env.ENGINE || "chromium"}): ${failures.length} błędów`);
    failures.forEach((f) => console.log("  ✗", f));
    process.exit(1);
  }
  console.log(`✅ anonymize (${process.env.ENGINE || "chromium"}): wszystko OK`);
}

run().catch((e) => { console.error(e); process.exit(1); });
