// nl-query-playwright.js — „✦ Zapytaj zwykłymi słowami” (nl-query.js) w oknie szybkiego szukania.
//
// Co musi być prawdą:
//   1. reguły bez AI: otwarte / niezakończone → @wtoku (nie @zamkniete), zamknięte, bez startu,
//      ostatni cykl, „ponad 2 tygodnie” → @dni>>14,
//   2. odmiana nazwiska („Kowalskiej”) trafia w dane jako początek wyrazu („Kowalsk”),
//   3. „Zastosuj” wpisuje zwykłe zapytanie aplikacji i daje TE SAME wiersze co ręczne
//      ustawienie trybów (liczy silnik arkusza, nie model),
//   4. „ile … na osobę” ustawia grupowanie wg kolumny osoby,
//   5. odpowiedź AI (atrapa) przechodzi przez sito: nieznany tryb, nieistniejąca kolumna i
//      słowo spoza danych są odrzucane.
//
// Uruchom z serwerem na APP_URL (domyślnie http://127.0.0.1:4175/).

const { chromium, webkit } = require("playwright");
const path = require("path");

const APP_URL = process.env.APP_URL || "http://127.0.0.1:4175/";
const ENGINE = process.env.ENGINE === "webkit" ? webkit : chromium;
const FILE = path.join(__dirname, "RODO Obieg terenów 2026 V5.xlsx");
const SHEET = "2025-2026";
const results = [];
const check = (name, ok, detail = "") => results.push({ name, ok: !!ok, detail });
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

async function run() {
  const browser = await ENGINE.launch({ headless: true });
  const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1280, height: 900 } });
  await context.addInitScript(() => localStorage.setItem("introPlayed", "true"));
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push("pageerror: " + e.message));
  page.on("console", (m) => { if (m.type() === "error") errors.push("console: " + m.text()); });

  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  await page.setInputFiles("#fileInput", FILE);
  await page.waitForFunction(() => document.getElementById("sheetSelect")?.options?.length > 0, null, { timeout: 20000 });
  await page.evaluate((sheet) => {
    const select = document.getElementById("sheetSelect");
    if ([...select.options].some((o) => o.value === sheet)) select.value = sheet;
    select.dispatchEvent(new Event("change", { bubbles: true }));
  }, SHEET);
  await sleep(400);
  await page.evaluate(() => {
    const auto = document.getElementById("autoHeaderRow");
    if (auto) { auto.checked = true; auto.dispatchEvent(new Event("change", { bubbles: true })); }
    if (typeof applyAutoHeaderRowIfEnabled === "function") applyAutoHeaderRowIfEnabled();
  });
  await sleep(400);
  await page.evaluate(() => loadBtn.click());
  await page.waitForFunction(() => typeof baseRows !== "undefined" && baseRows.length > 0 && typeof getSmartModel === "function" && getSmartModel(), null, { timeout: 20000 });
  // po wykryciu nagłówka arkusz buduje się drugi raz (setTimeout) — czekamy, aż się uspokoi,
  // inaczej spóźniona przebudowa wyzeruje filtr nałożony w teście
  await sleep(1500);

  // 1) reguły — czysta funkcja na stałym korpusie
  const rules = await page.evaluate(() => {
    const corpus = "kowalska\njan nowak";
    const p = (q) => { const r = parseNlQuery(q, { corpus }); return { tokens: r.tokens.map((x) => x.token), terms: r.terms, unknown: r.unknown }; };
    return {
      open: p("pokaż otwarte sprawy"),
      notClosed: p("niezakończone"),
      closedLast: p("zamknięte w ostatnim cyklu"),
      weeks: p("trwają ponad 2 tygodnie"),
      noStart: p("bez startu"),
      surname: p("sprawy Kowalskiej"),
      unknown: p("sprawy Wiśniewskiego"),
    };
  });
  check("„otwarte” → @wtoku", JSON.stringify(rules.open.tokens) === '["@wtoku"]', JSON.stringify(rules.open));
  check("„niezakończone” → @wtoku (nie @zamkniete)", JSON.stringify(rules.notClosed.tokens) === '["@wtoku"]', JSON.stringify(rules.notClosed));
  check("„zamknięte w ostatnim cyklu” → @zamkniete + @ostatni", rules.closedLast.tokens.includes("@zamkniete") && rules.closedLast.tokens.includes("@ostatni"), JSON.stringify(rules.closedLast));
  check("„ponad 2 tygodnie” → @dni>>14", rules.weeks.tokens.includes("@dni>>14"), JSON.stringify(rules.weeks));
  check("„bez startu” → @bezstartu", rules.noStart.tokens.includes("@bezstartu"), JSON.stringify(rules.noStart));
  check("„Kowalskiej” → szukaj „Kowalsk” (odmiana)", rules.surname.terms[0] === "Kowalsk" && !rules.surname.unknown.length, JSON.stringify(rules.surname));
  check("słowo spoza danych → „nie rozpoznałem”", rules.unknown.unknown.includes("Wiśniewskiego") && !rules.unknown.terms.length, JSON.stringify(rules.unknown));

  // 2) prawdziwy arkusz: osoba z otwartym cyklem, nazwisko w odmianie
  const probe = await page.evaluate(() => {
    const m = getSmartModel();
    for (const row of baseRows) {
      for (const block of m.blocks) {
        const from = String(getDisplayValue(row, block.startIndex + m.startIdx) ?? "").trim();
        const to = String(getDisplayValue(row, block.startIndex + m.endIdx) ?? "").trim();
        const who = String(getDisplayValue(row, block.startIndex + m.entityIdx) ?? "").trim();
        if (from && !to && /\s/.test(who)) {
          const surname = who.split(/\s+/).pop();
          const inflected = /a$/.test(surname) ? surname.replace(/a$/, "iej") : `${surname}a`;
          return { who, surname, inflected, entityHeader: currentHeaders[block.startIndex + m.entityIdx] };
        }
      }
    }
    return null;
  });
  check("w arkuszu jest osoba z otwartym cyklem", !!probe, JSON.stringify(probe));
  if (probe) {
    const flow = await page.evaluate(async (probe) => {
      openQuickSearchPopup();
      const input = document.getElementById("nlqInput");
      input.value = `pokaż otwarte sprawy ${probe.inflected} dłużej niż 30 dni`;
      document.getElementById("nlqBtn").click();
      const chips = [...document.querySelectorAll("#nlqResult .nlq-chip")].map((c) => c.textContent);
      const apply = [...document.querySelectorAll("#nlqResult .nlq-actions .btn")].find((b) => b.textContent === "Zastosuj");
      apply.click();
      await new Promise((r) => setTimeout(r, 600));
      const query = document.getElementById("quickSearchPopupInput").value;
            const viaNl = new Set(viewRows.map((r) => r.rowIndex0));
      // to samo ręcznie: zwykłe zapytanie z tokenami w polu szukania
      const term = query.split(" @")[0];
      searchQueryEl.value = `${term} @wtoku @dni>>30`;
      smartFilterState = { states: [], rowFlags: [], only: "", minDays: null };
      filtersCommitted = true;
      applyFilters();
      const manual = new Set(viewRows.map((r) => r.rowIndex0));
      const same = viaNl.size === manual.size && [...viaNl].every((x) => manual.has(x));
      return { chips, query, rows: viaNl.size, manual: manual.size, same };
    }, probe);
    check("chipy: szukaj + w toku + dłużej niż 30 dni", flow.chips.some((c) => /^szukaj „/.test(c)) && flow.chips.includes("w toku") && flow.chips.includes("dłużej niż 30 dni"), JSON.stringify(flow.chips));
    check("zapytanie w polu: „<nazwisko> @wtoku @dni>>30”", new RegExp(`^${probe.surname.slice(0, 4)}\\S* @wtoku @dni>>30$`, "i").test(flow.query), flow.query);
    check("te same wiersze co ręcznie ustawione tryby", flow.same, JSON.stringify(flow));

    const group = await page.evaluate(() => {
      const input = document.getElementById("nlqInput");
      input.value = "ile otwartych spraw na osobę";
      document.getElementById("nlqBtn").click();
      const chips = [...document.querySelectorAll("#nlqResult .nlq-chip")].map((c) => c.textContent);
      [...document.querySelectorAll("#nlqResult .nlq-actions .btn")].find((b) => b.textContent === "Zastosuj").click();
      return { chips, groupBy: aggregationWorkbenchState.groupBy };
    });
    check("„ile … na osobę” → grupowanie wg kolumny osoby", group.groupBy === probe.entityHeader && group.chips.some((c) => /^grupuj wg/.test(c)), JSON.stringify({ group, expected: probe.entityHeader }));
  }

  // 3) AI (atrapa) przez sito
  const ai = await page.evaluate(async () => {
    window.__localAiMock = () => '```json\n{"szukaj": ["Xyzzyq"], "tryby": ["@wtoku", "@hack"], "dni_ponad": 60, "grupuj": "Nieistniejąca"}\n```';
    const input = document.getElementById("nlqInput");
    input.value = "zaległe sprawy";
    document.getElementById("nlqBtn").click();
    const aiBtn = [...document.querySelectorAll("#nlqResult .nlq-actions .btn")].find((b) => b.textContent === "Zapytaj AI");
    if (!aiBtn) return { error: "no ai button" };
    aiBtn.click();
    for (let i = 0; i < 40 && !document.querySelector("#nlqResult .ai-badge"); i++) await new Promise((r) => setTimeout(r, 50));
    const chips = [...document.querySelectorAll("#nlqResult .nlq-chip")].map((c) => c.textContent);
    const note = document.querySelector("#nlqResult .nlq-note")?.textContent || "";
    delete window.__localAiMock;
    return { chips, note, badge: !!document.querySelector("#nlqResult .ai-badge") };
  });
  check("AI: znany tryb + dni przechodzą, „@hack” i nieistniejąca kolumna odrzucone", ai.badge && ai.chips.includes("w toku") && ai.chips.includes("dłużej niż 60 dni") && !ai.chips.some((c) => /grupuj|hack/.test(c)), JSON.stringify(ai));
  check("AI: słowo spoza danych → „nie rozpoznałem”", /Xyzzyq/.test(ai.note), ai.note);

  check("brak błędów strony", !errors.length, errors.join(" | "));
  await browser.close();
  let failed = 0;
  results.forEach((r) => { console.log(`${r.ok ? "✅" : "❌"} ${r.name}${r.ok || !r.detail ? "" : ` (${String(r.detail).slice(0, 500)})`}`); if (!r.ok) failed++; });
  if (failed) process.exit(1);
  console.log(`\n✅ nl-query: ${results.length}/${results.length} (${process.env.ENGINE || "chromium"})`);
}
run().catch((e) => { console.error(e); process.exit(1); });
