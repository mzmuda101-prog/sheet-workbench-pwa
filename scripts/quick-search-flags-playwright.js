// quick-search-flags-playwright.js — ikonki dopasowania w szybkim szukaniu (≠ = a… .*),
// składnia „Kolumna:wartość" i wykluczanie frazy !"…".
//
// Co musi być prawdą (każdy wynik porównujemy z NIEZALEŻNYM przeliczeniem w teście):
//   1. „Kolumna:wartość" (przy operatorach) patrzy tylko w tę kolumnę,
//   2. „Kolumna:!wartość" = dokładne dopełnienie pkt 1,
//   3. prefiks, który NIE jest nazwą nagłówka, zostaje zwykłym tekstem,
//   4. !"fraza ze spacją" wyklucza CAŁĄ frazę (dawniej tylko pierwsze słowo),
//   5. ≠ odwraca szukanie, = / a… / .* przełączają tryb Filtra 1 (wykluczają się),
//   6. regex zachowuje wielkość liter wzorca (\D ≠ \d),
//   7. live-podgląd liczy tak samo jak filtr (także z ≠ i z kolumną),
//   8. ikonka na PASKU przy już zastosowanym zapytaniu przefiltrowuje od razu.
//
// Uruchom z serwerem na APP_URL (domyślnie http://127.0.0.1:4175/).

const { chromium, webkit } = require("playwright");
const path = require("path");

const APP_URL = process.env.APP_URL || "http://127.0.0.1:4175/";
const FILE = path.join(__dirname, "stress-test-workbench.xlsx");
const ENGINE = process.env.ENGINE === "webkit" ? webkit : chromium;
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
  await page.evaluate(() => { try { ensureXlsxLibs && ensureXlsxLibs(false); } catch {} });
  await page.setInputFiles("#fileInput", FILE);
  await page.waitForFunction(() => document.getElementById("sheetSelect")?.options?.length > 0, null, { timeout: 15000 });
  await sleep(300);
  await page.click("#loadBtn");
  await sleep(1200);

  const failures = [];
  const report = {};
  const ok = (name, cond, got) => { if (!cond) failures.push(`${name} — dostałem: ${JSON.stringify(got)}`); };

  // Dane do scenariuszy: kolumna tekstowa z wartością, która występuje TEŻ w innych
  // kolumnach (żeby zawężenie do kolumny dało realną różnicę), i fraza ze spacją.
  const data = await page.evaluate(() => {
    const H = currentHeaders;
    const rows = baseRows;
    const disp = (r, i) => String(getDisplayValue(r, i) ?? "");
    let pick = null;
    for (let c = 0; c < H.length && !pick; c++) {
      if (!String(H[c] || "").trim()) continue;
      const seen = new Set();
      for (const r of rows) {
        const v = disp(r, c).trim();
        if (v.length < 3 || /^[\d\s.,:-]+$/.test(v) || seen.has(v)) continue;
        seen.add(v);
        const needle = v.slice(0, 3).toLowerCase();
        const inCol = rows.filter((x) => disp(x, c).toLowerCase().includes(needle)).length;
        const any = rows.filter((x) => H.some((_, i) => disp(x, i).toLowerCase().includes(needle))).length;
        if (inCol > 0 && any > inCol && inCol < rows.length) { pick = { col: c, header: H[c], needle, inCol, any }; break; }
      }
    }
    let phrase = null;
    outer: for (const r of rows) {
      for (let i = 0; i < H.length; i++) {
        const v = disp(r, i).trim();
        const words = v.split(/\s+/);
        if (words.length >= 2 && words[0].length >= 3 && !/\d/.test(v)) {
          const p = v.toLowerCase();
          const first = words[0].toLowerCase();
          const withPhrase = rows.filter((x) => H.some((_, k) => disp(x, k).toLowerCase().includes(p))).length;
          const withFirst = rows.filter((x) => H.some((_, k) => disp(x, k).toLowerCase().includes(first))).length;
          if (withFirst > withPhrase) { phrase = { text: v, withPhrase }; break outer; }
        }
      }
    }
    return { total: rows.length, pick, phrase };
  });
  report.data = data;
  ok("jest kolumna do testu zawężenia", !!data.pick, data);

  // Wyszukanie przez okno (Cmd+Shift+F): wpisz, Enter, zwróć liczbę wierszy widoku.
  const commit = async (query, { ops = true } = {}) => {
    await page.evaluate(({ query, ops }) => {
      closeQuickSearchPopup();
      openQuickSearchPopup();
      document.getElementById("quickSearchPopupOperators").checked = ops;
      document.getElementById("quickSearchPopupAction").value = "filter";
      const input = document.getElementById("quickSearchPopupInput");
      input.value = query;
      input.focus();
    }, { query, ops });
    await page.keyboard.press("Enter");
    await sleep(350);
    return page.evaluate(() => viewRows.length);
  };
  const setFlags = (mode, negated) => page.evaluate(({ mode, negated }) => {
    filterModeEl.value = mode;
    filterNegateEl.checked = negated;
    syncQuickSearchModeControls();
  }, { mode, negated });

  if (data.pick) {
    const { header, needle, inCol, any } = data.pick;
    await setFlags("contains", false);

    // 1. Kolumna:wartość
    const scoped = await commit(`${header}:${needle}`);
    report.scoped = { q: `${header}:${needle}`, got: scoped, expected: inCol };
    ok("Kolumna:wartość patrzy tylko w tę kolumnę", scoped === inCol, report.scoped);
    // nagłówek wielkimi literami i z nadmiarowymi spacjami — ten sam wynik
    const scopedCase = await commit(`${String(header).toUpperCase()}:${needle}`);
    ok("nazwa kolumny bez rozróżniania wielkości liter", scopedCase === inCol, scopedCase);

    // 2. Kolumna:!wartość
    const scopedNeg = await commit(`${header}:!${needle}`);
    report.scopedNeg = { got: scopedNeg, expected: data.total - inCol };
    ok("Kolumna:!wartość = dopełnienie", scopedNeg === data.total - inCol, report.scopedNeg);

    // bez operatorów „Kolumna:" jest zwykłym tekstem (nic się nie zmienia względem dawnej apki)
    const literal = await commit(`${header}:${needle}`, { ops: false });
    const literalExpected = await page.evaluate((q) => baseRows.filter((r) =>
      currentHeaders.some((_, i) => String(getDisplayValue(r, i) ?? "").toLowerCase().includes(q))).length,
      `${String(header).toLowerCase()}:${needle}`);
    ok("bez operatorów prefiks jest dosłownym tekstem", literal === literalExpected, { literal, literalExpected });

    // 3. prefiks spoza nagłówków → zwykły tekst
    const bogus = await commit(`zzqnieistnieje:${needle}`);
    ok("nieznany prefiks = zwykły tekst (0 trafień)", bogus === 0, bogus);

    // 5. ≠ klikiem w ikonkę okna + zwykłe zapytanie
    await page.evaluate(() => { closeQuickSearchPopup(); openQuickSearchPopup(); });
    await page.click('#qsFlagsPopup [data-qs-flag="negate"]');
    const negState = await page.evaluate(() => ({
      pressed: document.querySelector('#qsFlagsPopup [data-qs-flag="negate"]').getAttribute("aria-pressed"),
      barPressed: document.querySelector('#qsFlags [data-qs-flag="negate"]').getAttribute("aria-pressed"),
      panel: filterNegateEl.checked,
      amber: document.getElementById("quickSearchPopup").classList.contains("qs-negated"),
    }));
    ok("≠ włącza Odwróć w panelu i świeci w obu paskach", negState.pressed === "true" && negState.barPressed === "true" && negState.panel && negState.amber, negState);
    const neg = await commit(needle);
    report.neg = { got: neg, expected: data.total - any };
    ok("≠ odwraca szukanie", neg === data.total - any, report.neg);

    // 7. live-podgląd z ≠ liczy dopełnienie
    await page.evaluate((q) => {
      closeQuickSearchPopup(); openQuickSearchPopup();
      document.getElementById("quickSearchPopupOperators").checked = true;
      const input = document.getElementById("quickSearchPopupInput");
      input.value = q; input.dispatchEvent(new Event("input", { bubbles: true }));
    }, needle);
    await sleep(500);
    const liveNeg = await page.evaluate(() => ({
      items: document.querySelectorAll("#qsLiveResultsPopup .qs-live-item").length,
      empty: !!document.querySelector("#qsLiveResultsPopup .qs-live-empty"),
    }));
    ok("live-podgląd działa przy ≠", (data.total - any > 0) ? liveNeg.items > 0 : liveNeg.empty, liveNeg);
    await page.click('#qsFlagsPopup [data-qs-flag="negate"]'); // wyłącz

    // live z kolumną: liczba = inCol (sprawdzamy przez qsScanSheet — ten sam kod co lista)
    const liveScoped = await page.evaluate(({ header, needle }) =>
      qsScanSheet(currentSheetName, `${header}:${needle}`.toLowerCase(), { mode: "contains", negated: false, operators: true }, 3).count,
      { header, needle });
    ok("live-podgląd rozumie Kolumna:wartość", liveScoped === inCol, { liveScoped, inCol });

    // 5b. tryby wykluczają się; ponowny klik = powrót do „Zawiera"
    await page.click('#qsFlagsPopup [data-qs-flag="equals"]');
    await page.click('#qsFlagsPopup [data-qs-flag="starts_with"]');
    const modeA = await page.evaluate(() => ({
      mode: getNormalizedSelectValue(filterModeEl),
      pressed: [...document.querySelectorAll('#qsFlagsPopup .qs-flag[aria-pressed="true"]')].map((b) => b.dataset.qsFlag),
    }));
    ok("= potem a… → tylko „Zaczyna się”", modeA.mode === "starts_with" && modeA.pressed.join() === "starts_with", modeA);
    await page.click('#qsFlagsPopup [data-qs-flag="starts_with"]');
    const modeB = await page.evaluate(() => getNormalizedSelectValue(filterModeEl));
    ok("ponowny klik wraca do „Zawiera”", modeB === "contains", modeB);

    // zmiana w panelu → ikonki nadążają
    const fromPanel = await page.evaluate(() => {
      filterModeEl.value = "regex";
      filterModeEl.dispatchEvent(new Event("change", { bubbles: true }));
      return document.querySelector('#qsFlags [data-qs-flag="regex"]').getAttribute("aria-pressed");
    });
    ok("tryb ustawiony w panelu świeci się w szybkim szukaniu", fromPanel === "true", fromPanel);
  }

  // 6. regex: \D (wielka litera) ≠ \d
  if (data.pick) {
    await setFlags("regex", false);
    const col = data.pick.header;
    const re = await commit(`${col}:^\\D+$`);
    const reExpected = await page.evaluate((c) => {
      const idx = currentHeaders.map((h, i) => (String(h).trim().toLowerCase() === String(c).trim().toLowerCase() ? i : -1)).filter((i) => i >= 0);
      return baseRows.filter((r) => idx.some((i) => /^\D+$/i.test(String(getDisplayValue(r, i) ?? "")))).length;
    }, col);
    report.regex = { got: re, expected: reExpected };
    ok("regex zachowuje \\D (nie zamienia w \\d)", re === reExpected && re > 0, report.regex);
    await setFlags("contains", false);
  }

  // 4. !"fraza"
  if (data.phrase) {
    const got = await commit(`!"${data.phrase.text}"`);
    report.phrase = { q: `!"${data.phrase.text}"`, got, expected: data.total - data.phrase.withPhrase };
    ok("!\"fraza\" wyklucza całą frazę", got === data.total - data.phrase.withPhrase, report.phrase);
  }

  // 8. pasek: ikonka przy zastosowanym zapytaniu przefiltrowuje od razu
  if (data.pick) {
    await page.evaluate(() => closeQuickSearchPopup());
    const barVisible = await page.evaluate(() => {
      const wrap = document.getElementById("quickSearchWrap");
      wrap.classList.remove("hidden");
      return getComputedStyle(wrap).display !== "none";
    });
    if (barVisible) {
      const { needle, any } = data.pick;
      await page.evaluate((q) => {
        document.getElementById("quickSearchOperators").checked = true;
        quickSearchEl.value = q;
        commitQuickSearch();
      }, needle);
      await sleep(300);
      const before = await page.evaluate(() => viewRows.length);
      // Panel boczny (scrim) może zasłaniać pasek — klik programowy, ten sam handler.
      await page.evaluate(() => document.querySelector('#qsFlags [data-qs-flag="negate"]').click());
      await sleep(300);
      const after = await page.evaluate(() => viewRows.length);
      report.bar = { before, after, any };
      ok("≠ na pasku przefiltrowuje od razu", before === any && after === data.total - any, report.bar);
      await page.evaluate(() => document.querySelector('#qsFlags [data-qs-flag="negate"]').click());
    } else {
      report.bar = "pasek niewidoczny — pominięte";
    }
  }

  ok("brak błędów w konsoli", errors.length === 0, errors);
  await browser.close();

  console.log(JSON.stringify(report, null, 2));
  if (failures.length) {
    console.log("\n❌ " + failures.length + " błędów:\n - " + failures.join("\n - "));
    process.exit(1);
  }
  console.log("\n✅ quick-search-flags: wszystko OK");
}

run().catch((e) => { console.error(e); process.exit(1); });
