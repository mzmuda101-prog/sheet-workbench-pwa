// qs-tokens-playwright.js — podgląd szybkiego szukania = wynik Entera + podpowiadanie tokenów.
//
// Zgłoszenie Mateusza (2026-10-25): „@notatka” przed Enterem nic nie pokazywało, mimo że po
// Enterze filtr działał — podgląd miał własną logikę i szukał „@notatka” jako tekstu w komórkach.
// Do tego: tokeny trzeba było znać na pamięć, a na iPhonie „✓” nad klawiaturą chował klawiaturę,
// ale panel podpowiedzi wisiał dalej nad tabelą (i był częściowo pod klawiaturą).
//
// Co musi być prawdą:
//   1. podgląd liczy DOKŁADNIE to, co zostanie po Enterze (notatki: @notatka, !@notatka,
//      @notatka:słowo, tekst + token; cykle: @wtoku, @dni>>N, @ostatni),
//   2. przy @notatka pozycja podglądu pokazuje treść notatki (oznaczona jako notatka),
//   3. „@” / „@no” pokazuje listę pasujących tokenów z liczbą wierszy; tylko te, które mają
//      sens w arkuszu (bez notatek — bez @notatka, bez cykli — bez @wtoku),
//   4. Enter / Tab na niedokończonym tokenie wstawia podpowiedź (nie szuka „@no”), dopiero
//      kolejny Enter filtruje; klik w podpowiedź wstawia ją i zostawia fokus w polu,
//   5. wyjście z pola bez dotknięcia (jak „✓” na iPhonie) chowa panel i zwęża pole,
//   6. panel mieści się w widocznym obszarze (niskie okno = jak ekran z klawiaturą).

const playwright = require("playwright");
const path = require("path");

const APP_URL = process.env.APP_URL || "http://127.0.0.1:4175/";
const ENGINE = process.env.ENGINE || "chromium";
const SLEEP_SCALE = Math.min(1, Math.max(0.1, Number(process.env.SLEEP_SCALE || 1)));
const sleep = (ms) => new Promise((r) => setTimeout(r, Math.round(ms * SLEEP_SCALE)));
const sleepFixed = (ms) => new Promise((r) => setTimeout(r, ms));

async function open(browser, file, viewport = { width: 1280, height: 860 }) {
  const context = await browser.newContext({ serviceWorkers: "block", viewport });
  await context.addInitScript(() => localStorage.setItem("introPlayed", "true"));
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push("pageerror: " + e.message));
  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(() => { document.getElementById("heroSplash")?.remove(); return ensureXlsxLibs(false); });
  await page.setInputFiles("#fileInput", file);
  await page.waitForFunction(() => !!document.querySelector("#dataTable tbody tr[data-row-key]") || document.getElementById("sheetSelect")?.options?.length > 1, null, { timeout: 30000 });
  await sleep(600);
  return { context, page, errors };
}

const live = (page) => page.evaluate(() => {
  const el = document.getElementById("qsLiveResults");
  const items = [...el.querySelectorAll(".qs-live-item")];
  const more = (el.querySelector(".qs-live-more")?.textContent || "").match(/\d+/);
  return {
    hidden: el.classList.contains("hidden"),
    n: items.length + (more ? +more[0] : 0),
    tokens: items.filter((x) => x.classList.contains("qs-token-item")).map((x) => ({ tok: x.querySelector(".qs-token-name")?.textContent, count: x.querySelector(".qs-token-count")?.textContent || "" })),
    notes: items.filter((x) => x.classList.contains("is-note")).map((x) => x.textContent),
    first: items[0]?.textContent || "",
  };
});
const enterCount = (page) => page.evaluate(() => viewRows.length);

async function run() {
  const browser = await playwright[ENGINE].launch({ headless: true });
  const results = [];
  const check = (name, ok, detail = "") => results.push({ name, ok: !!ok, detail: typeof detail === "string" ? detail : JSON.stringify(detail) });
  const allErrors = [];

  // ── arkusz z notatką (K7) ──
  let { context, page, errors } = await open(browser, path.join(__dirname, "fixtures", "openpyxl-features.xlsx"));
  allErrors.push(errors);
  const typeQuery = async (q) => { await page.fill("#quickSearch", q); await sleepFixed(450); };
  for (const q of ["@notatka", "!@notatka", "@notatka:testowa", "@notatka:brak-takiego", "Nowak @notatka", "Nowak"]) {
    await page.click("#quickSearch");
    await typeQuery(q);
    const l = await live(page);
    await page.press("#quickSearch", "Enter");
    await sleep(400);
    const n = await enterCount(page);
    const shown = l.hidden ? 0 : l.n;
    check(`1. podgląd = Enter: „${q}” (${shown} / ${n})`, shown === n, { l, n });
    if (q === "@notatka") check("2. @notatka: podgląd pokazuje treść notatki", l.notes.length === 1 && /Notatka testowa/.test(l.notes[0]), l);
    await page.fill("#quickSearch", "");
    await page.press("#quickSearch", "Enter");
    await sleep(250);
  }

  // 3. podpowiedzi po „@”
  await page.click("#quickSearch");
  await typeQuery("@no");
  const t3 = await live(page);
  check("3. „@no” → podpowiedź @notatka z liczbą wierszy", t3.tokens.some((x) => x.tok === "@notatka" && /^1\b/.test(x.count)), t3);
  await typeQuery("@");
  const t3b = await live(page);
  check("3. arkusz bez cykli: bez @wtoku / @dni w podpowiedziach", t3b.tokens.length >= 3 && !t3b.tokens.some((x) => /wtoku|dni/.test(x.tok)), t3b);

  // 4. Enter na niedokończonym tokenie wstawia, kolejny filtruje
  await typeQuery("@not");
  await page.press("#quickSearch", "Enter");
  await sleep(200);
  const v4 = await page.inputValue("#quickSearch");
  check("4. Enter na „@not” wstawia „@notatka ” (bez szukania)", v4 === "@notatka ", v4);
  await page.press("#quickSearch", "Enter");
  await sleep(400);
  check("4. drugi Enter filtruje (1 wiersz z notatką)", (await enterCount(page)) === 1, await enterCount(page));
  await page.fill("#quickSearch", "");
  await page.press("#quickSearch", "Enter");
  await sleep(250);
  await page.click("#quickSearch");
  await typeQuery("Nowak !@n");
  await page.press("#quickSearch", "Tab");
  await sleep(200);
  const v4b = await page.inputValue("#quickSearch");
  const f4b = await page.evaluate(() => document.activeElement?.id);
  check("4. Tab dokańcza token w środku zapytania („Nowak !@notatka ”), fokus zostaje", v4b === "Nowak !@notatka " && f4b === "quickSearch", { v4b, f4b });
  await typeQuery("@n");
  await page.locator("#qsLiveResults .qs-token-item").filter({ has: page.locator(".qs-token-name", { hasText: /^@notatka:$/ }) }).click();
  await sleep(200);
  const v4c = await page.inputValue("#quickSearch");
  const f4c = await page.evaluate(() => document.activeElement?.id);
  check("4. klik w „@notatka:” wstawia token bez spacji (na słowo), fokus w polu", v4c === "@notatka:" && f4c === "quickSearch", { v4c, f4c });

  // 5. wyjście z pola bez dotknięcia (jak „✓” nad klawiaturą iPhone'a)
  await typeQuery("Nowak");
  const before5 = await page.evaluate(() => ({ live: !document.getElementById("qsLiveResults").classList.contains("hidden"), wide: document.querySelector(".table-toolbar").classList.contains("qs-wide") }));
  await page.evaluate(() => document.getElementById("quickSearch").blur());
  await sleep(150);
  const after5 = await page.evaluate(() => ({ live: !document.getElementById("qsLiveResults").classList.contains("hidden"), wide: document.querySelector(".table-toolbar").classList.contains("qs-wide") }));
  check("5. „✓” (blur bez dotknięcia): panel znika, pole wraca do zwykłej szerokości", before5.live && before5.wide && !after5.live && !after5.wide, { before5, after5 });
  await context.close();

  // 6. niskie okno: panel w widocznym obszarze
  ({ context, page, errors } = await open(browser, path.join(__dirname, "fixtures", "openpyxl-features.xlsx"), { width: 390, height: 430 }));
  allErrors.push(errors);
  await page.click("#quickSearch");
  await page.fill("#quickSearch", "20");
  await sleepFixed(500);
  const fit = await page.evaluate(() => { const r = document.getElementById("qsLiveResults").getBoundingClientRect(); const vv = window.visualViewport; return { bottom: Math.round(r.bottom), h: Math.round(r.height), visible: Math.round(vv ? vv.height : innerHeight), scrolls: document.getElementById("qsLiveResults").scrollHeight > document.getElementById("qsLiveResults").clientHeight }; });
  check("6. panel mieści się w widocznym obszarze (przewija się w środku)", fit.bottom <= fit.visible && fit.h >= 100, fit);
  await context.close();

  // ── arkusz z cyklami (tryby auto) ──
  ({ context, page, errors } = await open(browser, path.join(__dirname, "stress-test-workbench.xlsx")));
  allErrors.push(errors);
  await page.evaluate(() => { const sel = document.getElementById("sheetSelect"); sel.value = "MergedHeaders_Okresy"; loadBtn.click(); });
  await page.waitForFunction(() => currentSheetName === "MergedHeaders_Okresy" && sheetLoadedName === "MergedHeaders_Okresy" && !!getSmartModel(), null, { timeout: 30000 });
  await sleep(500);
  for (const q of ["@wtoku", "@dni>>60", "@ostatni @zakonczone"]) {
    await page.click("#quickSearch");
    await page.fill("#quickSearch", q);
    await sleepFixed(500);
    const l = await live(page);
    await page.press("#quickSearch", "Enter");
    await sleep(500);
    const n = await enterCount(page);
    check(`1. podgląd = Enter (cykle): „${q}” (${l.hidden ? 0 : l.n} / ${n})`, (l.hidden ? 0 : l.n) === n && n > 0, { l, n });
    await page.fill("#quickSearch", "");
    await page.press("#quickSearch", "Enter");
    await sleep(300);
  }
  await page.click("#quickSearch");
  await page.fill("#quickSearch", "@w");
  await sleepFixed(500);
  const t7 = await live(page);
  const exp7 = await page.evaluate(() => { const f = swbBuildRowFilter({ query1: "@wtoku", preview: true }); return baseRows.filter(f.rowPasses).length; });
  check("3. arkusz z cyklami: „@w” → @wtoku z liczbą = wynik filtra", t7.tokens.some((x) => x.tok === "@wtoku" && x.count.replace(/\D/g, "") === String(exp7)) && !t7.tokens.some((x) => /notatk/.test(x.tok)), { t7, exp7 });
  await context.close();

  const errs = allErrors.flat();
  check("brak błędów strony", errs.length === 0, errs.slice(0, 3).join(" | "));
  await browser.close();

  let failed = 0;
  for (const r of results) {
    console.log(`${r.ok ? "✅" : "❌"} ${r.name}${!r.ok && r.detail ? `  (${r.detail})` : ""}`);
    if (!r.ok) failed += 1;
  }
  if (failed) { console.error(`\n${failed} z ${results.length} nie przeszło`); process.exit(1); }
  console.log(`\n✅ qs-tokens (${ENGINE}): ${results.length}/${results.length}`);
}

run().catch((e) => { console.error(e); process.exit(1); });
