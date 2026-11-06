// anonymize-stages34-playwright.js — anonimizacja, etapy 3–4 (2026-11-06).
//
// Co musi być prawdą:
//   1. Polecenie tekstem ustawia metody i ślady, a okno pokazuje, co zrozumiało (i czego nie).
//   2. Profil: „Zapisz jako profil…” → wybór innego profilu zmienia ustawienia → powrót do
//      własnego przywraca je; profil bez hasła.
//   3. Raport ryzyka: przy unikalnych kombinacjach (kod + data urodzenia + płeć) ostrzega i wskazuje
//      cechę, która daje najwięcej.
//   4. Klucz odwracania: zapis „w miejscu” z hasłem → plik klucza bez danych jawnych; złe hasło =
//      komunikat; dobre = oryginalne wartości wracają (jedno Cofnij), komórki zmienione później zostają.
//   5. Drugi przebieg „w miejscu” przed zapisem nie gubi czyszczenia śladów pierwszego (łańcuch).
//   6. Etap 5 (AI ✦, atrapa modelu): propozycje z wolnego tekstu są odznaczone, zmyślony fragment
//      odrzucony, po zaznaczeniu propozycja trafia do planu.

const { chromium, webkit } = require("playwright");
const fs = require("fs");
const os = require("os");
const path = require("path");

const APP_URL = process.env.APP_URL || "http://127.0.0.1:4175/";
const ENGINE = process.env.ENGINE === "webkit" ? webkit : chromium;
const ROOT = path.join(__dirname, "..");
const XLSX = require(path.join(ROOT, "lib", "xlsx.full.min.js"));
const JSZip = require(path.join(ROOT, "lib", "jszip.min.js"));

const NAMES = ["Jan Kowalski", "Anna Nowak", "Piotr Zieliński", "Maria Wiśniewska", "Tomasz Lewandowski", "Ewa Kamińska", "Marek Wójcik", "Katarzyna Dąbrowska", "Paweł Kaczmarek", "Agnieszka Piotrowska"];
const PESELS = ["44051401359", "85010112347", "90020254321", "02270803624"];
function buildFixture() {
  const rows = [["Imię i nazwisko", "PESEL", "Telefon", "Kod pocztowy", "Data urodzenia", "Płeć", "Kwota", "Uwagi"]];
  const NOTES = ["Ustalono z Zenkiem z Zaułka Rybaków, że odbierze w piątek.", "Klient prosi, żeby Zośka oddzwoniła po 15.", "Faktura wysłana, brak dalszych uwag w tej sprawie."];
  NAMES.forEach((n, i) => rows.push([n, PESELS[i % 4], `600 100 1${String(i).padStart(2, "0")}`, i < 4 ? "62-200" : `0${i}-95${i}`, new Date(Date.UTC(1980 + (i % 3 === 0 ? 0 : i), i % 12, 10 + i)), i % 2 ? "K" : "M", 100 * i, NOTES[i] || ""]));
  const ws = XLSX.utils.aoa_to_sheet(rows, { cellDates: true });
  ws.A2.c = [{ a: "Ewa Mazur", t: "Sprawdzić u Jan Kowalski" }];
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Lista");
  const file = path.join(fs.mkdtempSync(path.join(os.tmpdir(), "anon34-")), "anon34.xlsx");
  fs.writeFileSync(file, XLSX.write(wb, { type: "buffer", bookType: "xlsx", cellDates: true }));
  return file;
}

async function run() {
  const FILE = buildFixture();
  const browser = await ENGINE.launch({ headless: true });
  const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1280, height: 900 } });
  await context.addInitScript(() => { localStorage.setItem("introPlayed", "true"); localStorage.setItem("swb-panel-pinned-v1", "1"); localStorage.removeItem("swb-anon-prefs-v1"); localStorage.removeItem("swb-anon-profiles-v1");
    // zapis plików (kopia / klucz) do pamięci strony — przed startem apki, żeby canFSA było prawdą też w WebKit
    window.__saved = {};
    window.showSaveFilePicker = async (o) => ({ name: o.suggestedName, createWritable: async () => ({ write: async (b) => { window.__saved[o.suggestedName] = await b.text(); }, close: async () => {} }) });
  });
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push("pageerror: " + e.message));
  page.on("console", (m) => { if (m.type() === "error") errors.push("console: " + m.text()); });
  page.on("dialog", (d) => (d.type() === "prompt" ? d.accept("Test miesięczny") : d.accept()));
  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(() => { document.getElementById("heroSplash")?.remove(); });
  const failures = [];
  const ok = (name, cond, got) => { if (!cond) failures.push(`${name} — dostałem: ${JSON.stringify(got)}`); };

  await page.evaluate(() => { if (typeof setDirtyState === "function") setDirtyState(false); });
  await page.setInputFiles("#fileInput", FILE);
  await page.waitForFunction(() => currentFileName === "anon34.xlsx" && document.getElementById("sheetSelect")?.options?.length > 0, null, { timeout: 30000 });
  await page.evaluate(() => loadBtn.click());
  await page.waitForFunction(() => Array.isArray(baseRows) && baseRows.length > 0, null, { timeout: 20000 });

  // ── 1. polecenie tekstem ──
  await page.evaluate(() => swbAnon.open({ scope: "file" }));
  await page.fill("#anonCommand", "PESEL zamaskuj, nazwiska na pseudonimy, daty do miesiąca, telefony zostaw, notatki usuń, zrób coś fajnego");
  await page.press("#anonCommand", "Enter");
  const c1 = await page.evaluate(() => ({ m: swbAnon.state.methods, tr: swbAnon.state.traces.notes, dg: swbAnon.state.dateGeneralize, txt: document.getElementById("anonCommandResult")?.textContent || "" }));
  ok("polecenie: PESEL → maska", c1.m.pesel === "mask", c1.m);
  ok("polecenie: osoby → pseudonim", c1.m.person === "pseudo", c1.m);
  ok("polecenie: telefony zostaw", c1.m.phone === "keep", c1.m);
  ok("polecenie: notatki usuń", c1.tr === "remove", c1.tr);
  ok("polecenie: daty do miesiąca", c1.dg === "month", c1.dg);
  ok("polecenie: okno pokazuje zrozumiane i niezrozumiane", /Zrozumiałem/.test(c1.txt) && /zrób coś fajnego/.test(c1.txt), c1.txt);

  // ── 2. skan + raport ryzyka ──
  await page.click("#anonScanBtn");
  await page.waitForFunction(() => swbAnon.state.plan, null, { timeout: 30000 });
  const risk = await page.evaluate(() => document.getElementById("anonRisk")?.textContent || "");
  ok("ryzyko: ostrzeżenie o unikalnych kombinacjach", /UNIKALN/.test(risk) && /Kod pocztowy/.test(risk) && /Płeć/.test(risk), risk);
  ok("ryzyko: podpowiedź, która cecha daje najwięcej", /Najwięcej daje/.test(risk), risk);
  const dob = await page.evaluate(() => swbAnon.state.plan.changes.filter((ch) => ch.type === "dob").map((ch) => ch.after && ch.after.value));
  ok("daty urodzenia do miesiąca (RRRR-MM)", dob.length > 0 && dob.every((v) => /^\d{4}-\d{2}$/.test(String(v))), dob);

  // ── 3. profile ──
  await page.click("#anonProfileSave");
  await page.waitForFunction(() => (JSON.parse(localStorage.getItem("swb-anon-profiles-v1") || "[]")).length === 1);
  const stored = await page.evaluate(() => localStorage.getItem("swb-anon-profiles-v1"));
  ok("profil zapisany pod nazwą", /Test miesięczny/.test(stored), stored);
  await page.fill("#anonKey", "tajne-haslo-spojnosci");
  await page.selectOption("#anonProfile", "b:test-data");
  const afterPreset = await page.evaluate(() => swbAnon.state.methods.pesel);
  ok("profil wbudowany zmienia metody", afterPreset === "fake", afterPreset);
  const own = await page.evaluate(() => [...document.querySelectorAll("#anonProfile option")].find((o) => o.textContent === "Test miesięczny")?.value);
  await page.selectOption("#anonProfile", own);
  const back = await page.evaluate(() => ({ pesel: swbAnon.state.methods.pesel, phone: swbAnon.state.methods.phone, notes: swbAnon.state.traces.notes, dg: swbAnon.state.dateGeneralize }));
  ok("własny profil przywraca ustawienia", back.pesel === "mask" && back.phone === "keep" && back.notes === "remove" && back.dg === "month", back);
  const stored2 = await page.evaluate(() => localStorage.getItem("swb-anon-profiles-v1"));
  ok("profil bez hasła", !/tajne-haslo/.test(stored2), stored2);
  await page.fill("#anonKey", "");

  // ── 4. klucz odwracania + przebieg „w miejscu” ──
  const orig = await page.evaluate(() => ({ a2: workbook.Sheets.Lista.A2.v, b2: workbook.Sheets.Lista.B2.v }));
  await page.click("#anonScanBtn");
  await page.waitForFunction(() => swbAnon.state.plan && swbAnon.state.plan.changes.length > 0);
  await page.evaluate(() => document.querySelector('#anonModal input[name="anonOut"][value="inplace"]').click());
  await page.locator("#anonModal .anon-key-opt input[type=checkbox]").check();
  await page.fill("#anonKeyPass", "krotkie");
  await page.fill("#anonKeyPass2", "krotkie");
  await page.click("#anonRunBtn");
  const weak = await page.evaluate(() => workbook.Sheets.Lista.A2.v);
  ok("za krótkie hasło klucza blokuje anonimizację", weak === orig.a2, weak);
  await page.fill("#anonKeyPass", "Klucz-Testowy-2026");
  await page.fill("#anonKeyPass2", "Klucz-Testowy-2026");
  await page.click("#anonRunBtn");
  await page.waitForFunction(() => Object.keys(window.__saved).some((n) => /klucz odwracania/.test(n)), null, { timeout: 30000 });
  const after1 = await page.evaluate(() => ({ a2: workbook.Sheets.Lista.A2.v, b2: workbook.Sheets.Lista.B2.v, cleanup: !!swbAnonPendingCleanup }));
  ok("w miejscu: osoba zamieniona", after1.a2 !== orig.a2 && /^Osoba \d+/.test(after1.a2), after1);
  const keyName = await page.evaluate(() => Object.keys(window.__saved).find((n) => /klucz odwracania/.test(n)));
  const keyText = await page.evaluate((n) => window.__saved[n], keyName);
  ok("plik klucza bez jawnych danych", /swb-anon-key/.test(keyText) && !/Kowalski|44051401359|anon34/.test(keyText), keyText.slice(0, 120));

  // drugi przebieg przed zapisem: telefony usuń, notatki zostaw — czyszczenie 1. przebiegu (notatki usuń) zostaje
  await page.evaluate(() => swbAnon.open({ scope: "file" }));
  await page.fill("#anonCommand", "telefony usuń, resztę zostaw, notatki zostaw");
  await page.press("#anonCommand", "Enter");
  await page.click("#anonScanBtn");
  await page.waitForFunction(() => swbAnon.state.plan && swbAnon.state.plan.changes.length > 0);
  await page.evaluate(() => document.querySelector('#anonModal input[name="anonOut"][value="inplace"]').click());
  await page.locator("#anonModal .anon-key-opt input[type=checkbox]").uncheck();
  await page.click("#anonRunBtn");
  await page.waitForFunction(() => !workbook.Sheets.Lista.C2 || workbook.Sheets.Lista.C2.t === "z");
  const chain = await page.evaluate(async () => {
    const out = await swbAnon.cleanOnSave(new Uint8Array(originalFileBytes));
    const zip = await JSZip.loadAsync(out);
    const parts = Object.keys(zip.files);
    let text = "";
    for (const p of parts) if (/\.(xml|vml)$/.test(p)) text += await zip.file(p).async("string");
    return { prev: !!(swbAnonPendingCleanup && swbAnonPendingCleanup.prev), mazur: /Ewa Mazur/.test(text), comments: parts.some((p) => /comments/i.test(p)) };
  });
  ok("drugi przebieg: łańcuch czyszczeń", chain.prev, chain);
  ok("drugi przebieg: notatki z 1. przebiegu usunięte przy zapisie", !chain.mazur && !chain.comments, chain);
  await page.evaluate(() => swbAnon.close());

  // ── 5. przywracanie z klucza ──
  await page.click("#anonRestoreBtn").catch(async () => { await page.evaluate(() => swbAnon.openRestore()); });
  await page.waitForSelector("#anonRestoreModal:not(.hidden)");
  await page.setInputFiles("#anonRestoreFile", { name: keyName, mimeType: "application/json", buffer: Buffer.from(keyText) });
  await page.fill("#anonRestorePass", "zle-haslo-123");
  await page.click("#anonRestoreRun");
  await page.waitForFunction(() => /Złe hasło/.test(document.getElementById("anonRestoreResult").textContent), null, { timeout: 20000 });
  await page.fill("#anonRestorePass", "Klucz-Testowy-2026");
  await page.click("#anonRestoreRun");
  await page.waitForFunction(() => /Przywrócono/.test(document.getElementById("anonRestoreResult").textContent), null, { timeout: 20000 });
  const restored = await page.evaluate(() => ({ a2: workbook.Sheets.Lista.A2.v, b2: workbook.Sheets.Lista.B2.v, txt: document.getElementById("anonRestoreResult").textContent, cleanup: swbAnonPendingCleanup }));
  ok("przywrócono osobę", restored.a2 === orig.a2, restored);
  ok("przywrócono PESEL", String(restored.b2) === String(orig.b2), restored);
  ok("po przywróceniu zapis nie czyści śladów", restored.cleanup === null, restored.cleanup);
  // Cofnij przywracanie → wraca stan zanonimizowany i jego czyszczenie
  await page.evaluate(() => { document.getElementById("anonRestoreModal").classList.add("hidden"); undoLast(); });
  const undone = await page.evaluate(() => ({ a2: workbook.Sheets.Lista.A2.v, cleanup: !!swbAnonPendingCleanup }));
  ok("Cofnij przywracanie: znów pseudonim i czyszczenie", /^Osoba/.test(undone.a2) && undone.cleanup, undone);

  // ── 5b. złośliwy klucz: arkusz „__proto__” / „constructor” nie zapisuje do Object.prototype ──
  const proto = await page.evaluate(async () => {
    const txt = await SwbAnonTools.encryptReverseKey([
      { sheet: "__proto__", r: 0, c: 0, before: "WSTRZYK", after: null },
      { sheet: "constructor", r: 1, c: 1, before: "WSTRZYK", after: null },
    ], "Klucz-Zly-2026");
    const data = await SwbAnonTools.decryptReverseKey(txt, "Klucz-Zly-2026");
    swbAnon.restoreFromKey(data);
    return { a1: ({}).A1, b2: ({}).B2, own: Object.prototype.hasOwnProperty.call(Object.prototype, "A1") };
  });
  ok("złośliwy klucz nie zapisuje do Object.prototype", proto.a1 === undefined && proto.b2 === undefined && !proto.own, proto);

  // ── 6. AI ✦ w wolnym tekście (atrapa modelu) ──
  await page.evaluate(() => {
    window.__aiCalls = 0;
    window.__localAiMock = (messages) => {
      window.__aiCalls++;
      const user = messages[messages.length - 1].content;
      const found = [];
      user.split("\n").forEach((line) => {
        const m = line.match(/^(\d+)\. (.*)$/);
        if (!m) return;
        if (/Zenkiem/.test(m[2])) found.push({ nr: +m[1], fragment: "Zenkiem", typ: "osoba" }, { nr: +m[1], fragment: "Zaułka Rybaków", typ: "adres" }, { nr: +m[1], fragment: "Jan Nowakowski", typ: "osoba" });
        if (/Zośka/.test(m[2])) found.push({ nr: +m[1], fragment: "Zośka", typ: "osoba" });
      });
      return JSON.stringify({ znalezione: found });
    };
  });
  await page.evaluate(() => swbAnon.open({ scope: "file" }));
  await page.fill("#anonCommand", "nazwiska na pseudonimy, adresy na pseudonimy");
  await page.press("#anonCommand", "Enter");
  await page.click("#anonScanBtn");
  await page.waitForFunction(() => swbAnon.state.plan);
  ok("AI: sekcja wolnego tekstu jest", await page.isVisible("#anonAi"), null);
  await page.click("#anonAiRun");
  await page.waitForFunction(() => document.getElementById("anonAiMsg"), null, { timeout: 20000 });
  const ai = await page.evaluate(() => ({
    msg: document.getElementById("anonAiMsg").textContent, calls: window.__aiCalls,
    terms: [...swbAnon.state.aiTerms.values()].map((t) => `${t.type}:${t.value}`),
    aiVals: [...document.querySelectorAll("#anonModal .anon-value.is-ai")].map((l) => ({ t: l.querySelector(".anon-value-text").textContent, on: l.querySelector("input").checked })),
    planHasZenek: swbAnon.state.plan.changes.some((ch) => /Zenkiem/.test(String(ch.before)) && !/Zenkiem/.test(String(ch.after && ch.after.value))),
  }));
  ok("AI: model wywołany", ai.calls >= 1, ai);
  ok("AI: propozycje z tekstu", ai.terms.includes("person:Zenkiem") && ai.terms.includes("address:Zaułka Rybaków") && ai.terms.includes("person:Zośka"), ai.terms);
  ok("AI: zmyślony fragment odrzucony", !ai.terms.some((t) => /Nowakowski/.test(t)), ai.terms);
  ok("AI: propozycje odznaczone (tylko proponuje)", ai.aiVals.length >= 3 && ai.aiVals.every((v) => !v.on) && !ai.planHasZenek, ai);
  await page.locator("#anonModal .anon-value.is-ai", { hasText: "Zenkiem" }).locator("input").check();
  const picked = await page.evaluate(() => swbAnon.state.plan.changes.filter((ch) => /Zenkiem/.test(String(ch.before))).map((ch) => ch.after && ch.after.value));
  ok("AI: zaznaczona propozycja w planie", picked.length === 1 && !/Zenkiem/.test(picked[0]) && /Zaułka Rybaków/.test(picked[0]), picked);
  await page.evaluate(() => { delete window.__localAiMock; swbAnon.close(); });

  ok("brak błędów JS", errors.length === 0, errors);
  await browser.close();
  if (failures.length) { console.log(`❌ anonymize-stages34 (${process.env.ENGINE || "chromium"}): ${failures.length} błędów`); failures.forEach((f) => console.log("  ✗", f)); process.exit(1); }
  console.log(`✅ anonymize-stages34 (${process.env.ENGINE || "chromium"}): wszystko przeszło`);
}
run().catch((e) => { console.error(e); process.exit(1); });
