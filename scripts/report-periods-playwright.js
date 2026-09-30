// report-periods-playwright.js — kąt „Porównanie okresów”: elastyczny, ale nie naiwny.
//
// Stałe daty z 2025 r. (wynik nie zależy od dnia uruchomienia testu).
// Sprawdzamy:
//   1. kolumna z datą: czerwiec vs maj (+50%), vs średnia 3 poprzednich (1,6×), rok wcześniej,
//      suma kwot, zmiana składu (Status), nowe wartości (Klient), trend z pustymi miesiącami = 0,
//   2. ręczne poprawki: okres (kwartał), okres do porównania, źródło czasu,
//   3. dane urwane na początku okresu → domyślnie ostatni pełny + notka (także na kartce);
//      ręcznie wybrany niepełny okres → „do tego samego dnia”,
//   4. Rok + Miesiąc (nazwy miesięcy) — bez tygodni,
//   5. liczby Excela w kolumnie „Data” = daty; w kolumnie „Kwota” — NIE,
//   6. brak czasu → kąt niedostępny z powodem; kilka wpisów → bez procentów,
//   7. zakres raportu obcinający przeszłość → ostrzeżenie w panelu,
//   8. łamanie stron = PDF; (jeśli jest plik) RODO: cykle, kwiecień vs marzec.

const fs = require("fs");
const { chromium } = require("playwright");
const APP_URL = process.env.APP_URL || "http://127.0.0.1:4175/";
const SLEEP_SCALE = Math.min(1, Math.max(0.1, Number(process.env.SLEEP_SCALE || 1)));
const sleep = (ms) => new Promise((r) => setTimeout(r, Math.round(ms * SLEEP_SCALE)));
const RODO = "scripts/RODO Obieg terenów 2026 V5.xlsx";

async function loadRows(page, { headers, rows, file }) {
  await page.evaluate(({ headers, rows, file }) => {
    const ws = {};
    const col = (i) => String.fromCharCode(65 + i);
    headers.forEach((h, i) => { ws[col(i) + "1"] = { t: "s", v: h }; });
    rows.forEach((r, ri) => r.forEach((v, ci) => {
      if (v === null || v === "") return;
      ws[col(ci) + (ri + 2)] = typeof v === "number" ? { t: "n", v, w: String(v) } : { t: "s", v };
    }));
    ws["!ref"] = `A1:${col(headers.length - 1)}${rows.length + 1}`;
    workbook = { SheetNames: ["Arkusz"], Sheets: { Arkusz: ws }, Props: {} };
    currentFileName = file;
    sheetSelect.replaceChildren();
    const opt = document.createElement("option"); opt.value = "Arkusz"; opt.textContent = "Arkusz";
    sheetSelect.appendChild(opt); sheetSelect.value = "Arkusz";
    document.getElementById("headerRow").value = "1";
    document.getElementById("autoHeaderRow").checked = false;
    document.getElementById("loadBtn").click();
  }, { headers, rows, file });
  await page.waitForFunction(({ n, file }) => typeof baseRows !== "undefined" && baseRows.length === n && currentFileName === file, { n: rows.length, file }, { timeout: 15000 });
  await sleep(300);
}

const pad = (n) => String(n).padStart(2, "0");
// Wiersze: [Klient, Status, Kwota, Data]
function monthRows(y, m, n, { status = "W toku", kwota = 100, clientBase = "Klient", day = (i) => 1 + (i % 28) } = {}) {
  return Array.from({ length: n }, (_, i) => [`${clientBase} ${i % 6}`, typeof status === "function" ? status(i) : status, kwota, `${y}-${pad(m)}-${pad(day(i))}`]);
}

const state = (page) => page.evaluate(() => {
  const m = window.__report.data().periods;
  return {
    m: m && { src: m.src.id, gran: m.gran, cur: m.cur.label, prev: m.prev.label, n: m.cur.n, pn: m.prev.n, avg: m.avg && m.avg.n, yoy: m.yoy && m.yoy.n, partial: m.partial, autoNote: m.autoNote, trend: m.trend.map((x) => x.n), sources: m.sources.map((x) => x.id), allowed: m.allowed },
    f: window.__report.findings(),
    sec: (document.querySelector('#rpPage [data-section="periods"]')?.innerText || "").replace(/ | /g, " "),
    info: document.querySelector("#rpAngles .rp-angle-info")?.textContent || "",
    infoWarn: !!document.querySelector("#rpAngles .rp-angle-info.is-warn"),
  };
});

async function run() {
  const browser = await chromium.launch({ headless: true });
  const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1200, height: 900 } });
  await context.addInitScript(() => {
    localStorage.setItem("introPlayed", "true");
    localStorage.setItem("swb-report-prefs", JSON.stringify({ preset: "short", angle: "periods" }));
  });
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push("pageerror: " + e.message));
  page.on("console", (m) => { if (m.type() === "error") errors.push("console: " + m.text()); });
  const results = [];
  const check = (name, cond, got) => {
    results.push({ name, ok: !!cond });
    console.log(`${cond ? "✅" : "❌"} ${name}${cond ? "" : ` — dostałem: ${JSON.stringify(got)}`}`);
  };

  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  await page.evaluate(() => ensureXlsxLibs(false));
  await page.waitForFunction(() => typeof buildRows === "function" && window.__report);

  // ── 1. Kolumna z datą ──
  // styczeń–czerwiec 2025: 10, 12, 8, 10, 10, 15; czerwiec 2024: 5; lipiec–grudzień 2024 puste.
  // Czerwiec: 60% „Zakończone” (maj: 0%), dwóch nowych klientów; ostatni wpis 30 czerwca (pełny okres).
  const rowsA = [
    ...monthRows(2024, 6, 5, { kwota: 50 }),
    ...monthRows(2025, 1, 10), ...monthRows(2025, 2, 12), ...monthRows(2025, 3, 8),
    ...monthRows(2025, 4, 10), ...monthRows(2025, 5, 10),
    ...monthRows(2025, 6, 15, { status: (i) => (i < 9 ? "Zakończone" : "W toku"), kwota: 200, day: (i) => (i === 14 ? 30 : 1 + i) }),
    ["Nowy Klient A", "W toku", 10, "2025-06-10"], ["Nowy Klient B", "W toku", 10, "2025-06-11"],
  ];
  await loadRows(page, { headers: ["Klient", "Status", "Kwota", "Data"], rows: rowsA, file: "okresy.xlsx" });
  await page.evaluate(() => { window.__report.open(); document.getElementById("rpContentBtn").click(); });
  await sleep(250);
  let s = await state(page);
  check("źródło = kolumna „Data”, okres = miesiąc, czerwiec 2025 vs maj 2025", s.m && s.m.src.startsWith("date:") && s.m.gran === "month" && s.m.cur === "czerwiec 2025" && s.m.prev === "maj 2025", s.m);
  check("czerwiec 17 wpisów vs maj 10; średnia 3 poprzednich 9,33; rok wcześniej 5", s.m.n === 17 && s.m.pn === 10 && Math.abs(s.m.avg - 28 / 3) < 0.01 && s.m.yoy === 5, s.m);
  check("wniosek główny: +70% i 1,8× średniej", /czerwiec 2025 — wpisy: 17, o 70% więcej niż poprzednio \(maj 2025: 10\)\. To 1,8× średniej/.test(s.f[0]), s.f);
  check("rok wcześniej: 5 → +12 (+240%)", s.f.some((x) => /Rok wcześniej \(czerwiec 2024\): 5 — zmiana \+12 \(\+240%\)/.test(x)), s.f);
  check("suma kwot: 3020 vs 1000", s.f.some((x) => /Suma „Kwota”: 3020 wobec 1000 poprzednio/.test(x.replace(/\s(?=\d{3}\b)/g, ""))), s.f);
  check("zmiana składu: „Zakończone” 53% teraz, 0% poprzednio", s.f.some((x) => /„Status” = „Zakończone”: 53% wpisów teraz, 0% poprzednio/.test(x)), s.f);
  check("nowe wartości: 2 nowych klientów", s.f.some((x) => /Nowe w tym okresie \(„Klient”\): 2, np\. „Nowy Klient A”, „Nowy Klient B”/.test(x)), s.f);
  check("trend 12 miesięcy z pustymi = 0 (lip–gru 2024)", JSON.stringify(s.m.trend) === JSON.stringify([0, 0, 0, 0, 0, 0, 10, 12, 8, 10, 10, 17]), s.m.trend);
  check("tabela: miary × okresy (z „Śr. 3 poprz.” i rokiem wcześniej)", /Miara\tczerwiec 2025\tmaj 2025\tZmiana\tŚr\. 3 poprz\.\tZmiana\tczerwiec 2024/.test(s.sec) && /Wpisy\t17\t10\t\+7 \(\+70%\)/.test(s.sec), s.sec.slice(0, 300));
  const picks = await page.evaluate(() => ["rpPeriodSrc", "rpPeriodGran", "rpPeriodAnchor"].map((id) => !!document.getElementById(id)));
  check("panel: trzy ręczne poprawki (Czas wg / Okres / Porównaj)", picks.every(Boolean), picks);

  // ── 2. Ręczne poprawki ──
  await page.selectOption("#rpPeriodGran", "quarter");
  await sleep(200);
  s = await state(page);
  check("okres = kwartał: 2 kw. 2025 (45) vs 1 kw. 2025 (30)", s.m.gran === "quarter" && s.m.cur === "2 kw. 2025" && s.m.n === 37 && s.m.pn === 30, s.m);
  await page.selectOption("#rpPeriodGran", "month");
  await sleep(150);
  const aprilKey = await page.evaluate(() => Array.from(document.querySelectorAll("#rpPeriodAnchor option")).find((o) => o.textContent === "kwiecień 2025").value);
  await page.selectOption("#rpPeriodAnchor", aprilKey);
  await sleep(200);
  s = await state(page);
  check("ręczny okres do porównania: kwiecień 2025 vs marzec 2025", s.m.cur === "kwiecień 2025" && s.m.prev === "marzec 2025" && s.m.n === 10 && s.m.pn === 8, s.m);
  await page.evaluate(() => window.__report.setPeriodPick({ src: "", anchor: "", gran: "" }));

  // ── 7. Zakres obcinający przeszłość ──
  // Zakres tylko z czerwca: automatycznie raport przechodzi na TYGODNIE (sensowne dla 30 dni),
  // a przy ręcznie wybranych miesiącach mówi, że zakres obciął wcześniejsze okresy.
  await page.evaluate(() => window.__report.setScope({ kind: "query", q: "Data:>>=2025-06-01" }));
  await sleep(200);
  s = await state(page);
  check("zakres = tylko czerwiec → auto przechodzi na tygodnie + ostrzeżenie, od kiedy jest zakres", s.m.gran === "week" && s.infoWarn && /Zakres raportu zaczyna się 1 cze 2025 — wcześniejsze dane arkusza nie wchodzą/.test(s.info), s);
  await page.evaluate(() => { window.__report.setPeriodPick({ gran: "" }); window.__report.setScope({ kind: "view" }); });

  // ── 8a. Strony = PDF ──
  await page.evaluate(() => window.__report.setPreset("normal"));
  await sleep(300);
  const pages = await page.evaluate(() => ({ n: window.__report.pageCount(), bad: window.__report.marginViolations() }));
  await page.evaluate(() => document.body.classList.add("rp-printing"));
  await page.emulateMedia({ media: "print" });
  const pdf = await page.pdf({ format: "A4", preferCSSPageSize: true });
  await page.emulateMedia({ media: "screen" });
  await page.evaluate(() => document.body.classList.remove("rp-printing"));
  const realPages = (pdf.toString("latin1").match(/\/Type\s*\/Page[^s]/g) || []).length;
  check("porównanie okresów + normalny: pasek stron = PDF, marginesy OK", pages.n === realPages && !pages.bad.length, { ...pages, realPages });
  await page.evaluate(() => { window.__report.setPreset("short"); window.__report.close(); });

  // ── 3. Dane urwane na początku okresu ──
  await loadRows(page, {
    headers: ["Klient", "Status", "Kwota", "Data"],
    rows: [...monthRows(2025, 3, 20), ...monthRows(2025, 4, 20), ...monthRows(2025, 5, 20, { day: (i) => 1 + (i % 30) }), ["X", "W toku", 1, "2025-06-01"], ["Y", "W toku", 1, "2025-06-02"], ["Z", "W toku", 1, "2025-06-02"]],
    file: "urwane.xlsx",
  });
  await page.evaluate(() => { window.__report.open(); document.getElementById("rpContentBtn").click(); });
  await sleep(250);
  s = await state(page);
  check("dane kończą się 2 czerwca → domyślnie maj (pełny) vs kwiecień + notka w panelu i na kartce",
    s.m.cur === "maj 2025" && s.m.prev === "kwiecień 2025" && !s.m.partial && /Dane kończą się 2 cze 2025 — w „czerwiec 2025” jest ich za mało/.test(s.info) && /Inny okres wybierzesz/.test(s.info) && /Dane kończą się 2 cze 2025/.test(s.sec) && !/Inny okres/.test(s.sec), s);
  const juneKey = await page.evaluate(() => Array.from(document.querySelectorAll("#rpPeriodAnchor option")).find((o) => o.textContent === "czerwiec 2025").value);
  await page.selectOption("#rpPeriodAnchor", juneKey);
  await sleep(200);
  s = await state(page);
  check("ręcznie czerwiec → „do tego samego dnia”: 3 wobec maja 1–2 (2), bez procentów", s.m.partial && s.m.n === 3 && s.m.pn === 2 && s.f.some((x) => /Dane kończą się 2 cze 2025 — „czerwiec 2025” jest niepełny/.test(x)) && s.f.some((x) => /Za mało wpisów na wnioski procentowe/.test(x)), s);
  await page.evaluate(() => window.__report.close());

  // ── 4. Rok + Miesiąc ──
  const months = ["Styczeń", "Luty", "Marzec", "Kwiecień", "Maj", "Czerwiec"];
  await loadRows(page, {
    headers: ["Dział", "Rok", "Miesiąc", "Kwota"],
    rows: [2024, 2025].flatMap((y) => months.flatMap((mn, mi) => Array.from({ length: 6 + mi + (y === 2025 ? 3 : 0) }, (_, i) => [`D${i % 3}`, y, mn, 100]))),
    file: "rokmiesiac.xlsx",
  });
  await page.evaluate(() => { window.__report.open(); document.getElementById("rpContentBtn").click(); });
  await sleep(250);
  s = await state(page);
  check("Rok + Miesiąc (nazwy): czerwiec 2025 (14) vs maj 2025 (13), rok wcześniej 11", s.m && s.m.src === "ym" && s.m.cur === "czerwiec 2025" && s.m.n === 14 && s.m.pn === 13 && s.m.yoy === 11, s.m);
  check("Rok + Miesiąc: bez tygodni (tylko miesiąc/kwartał/rok), bez udawanego „niepełnego”", !s.m.allowed.includes("week") && !s.m.partial, s.m);
  await page.evaluate(() => window.__report.close());

  // ── 5. Liczby Excela ──
  const serial = (y, m, d) => Math.round((Date.UTC(y, m - 1, d) - Date.UTC(1899, 11, 30)) / 86400000);
  await loadRows(page, {
    headers: ["Pozycja", "Kwota", "Data zamówienia"],
    rows: [3, 4, 5, 6].flatMap((m) => Array.from({ length: 8 }, (_, i) => [`P${i}`, 45000 + i, serial(2025, m, 1 + i)])),
    file: "seriale.xlsx",
  });
  await page.evaluate(() => { window.__report.open(); document.getElementById("rpContentBtn").click(); });
  await sleep(250);
  s = await state(page);
  check("liczby Excela w „Data zamówienia” = daty; „Kwota” (45000…) NIE jest źródłem czasu", s.m && s.m.src === "serial:2" && s.m.sources.every((x) => x !== "serial:1") && s.m.cur === "czerwiec 2025", s.m);
  await page.evaluate(() => window.__report.close());

  // ── 6. Brak czasu ──
  await loadRows(page, { headers: ["Region", "Dział", "Kwota"], rows: Array.from({ length: 20 }, (_, i) => [`R${i % 4}`, `D${i % 3}`, 100 + i]), file: "bezczasu.xlsx" });
  await page.evaluate(() => { window.__report.open(); document.getElementById("rpContentBtn").click(); });
  await sleep(200);
  const none = await page.evaluate(() => ({ m: !!window.__report.data().periods, dis: document.querySelector('#rpAngles [data-angle="periods"]').disabled, hint: document.querySelector('#rpAngles [data-angle="periods"]').getAttribute("data-hint"), secs: window.__report.sections() }));
  check("brak czasu: kąt niedostępny, powód w dymku, brak sekcji", !none.m && none.dis && /Brak czasu w danych/.test(none.hint) && !none.secs.includes("periods"), none);
  await page.evaluate(() => { window.__report.setAngle("overview"); window.__report.close(); });

  // ── 8b. RODO ──
  if (fs.existsSync(RODO)) {
    await page.setInputFiles("#fileInput", RODO);
    await page.waitForFunction(() => document.getElementById("sheetSelect")?.options?.length > 1, null, { timeout: 30000 });
    await page.evaluate(() => {
      const el = document.getElementById("sheetSelect"); el.value = "2025-2026"; el.dispatchEvent(new Event("change", { bubbles: true }));
      document.getElementById("autoHeaderRow").checked = false; document.getElementById("headerRow").value = "2";
    });
    await sleep(300);
    await page.evaluate(() => document.getElementById("loadBtn").click());
    await page.waitForFunction(() => currentFileName.startsWith("RODO") && baseRows.length > 100 && currentDisplayModel, null, { timeout: 60000 });
    await sleep(600);
    await page.evaluate(() => { window.__report.open(); window.__report.setAngle("periods"); });
    await sleep(300);
    s = await state(page);
    check("RODO: starty cykli ze wszystkich bloków, kwiecień 2026 vs marzec 2026 (maj urwany)", s.m && s.m.src === "records" && s.m.cur === "kwiecień 2026" && s.m.prev === "marzec 2026" && /Dane kończą się/.test(s.sec), s.m);
    check("RODO: przepływ — rozpoczęte vs zakończone", s.f.some((x) => /Rozpoczęto \d+, zakończono \d+/.test(x)), s.f);
    await page.evaluate(() => { window.__report.setAngle("overview"); window.__report.close(); });
  }

  check("brak błędów w konsoli", errors.length === 0, errors);
  await browser.close();
  const failed = results.filter((r) => !r.ok);
  if (failed.length) {
    console.log(`\n❌ report-periods: ${failed.length} z ${results.length} nie przeszło`);
    process.exit(1);
  }
  console.log(`\n✅ report-periods: ${results.length}/${results.length}`);
}

run().catch((e) => { console.error(e); process.exit(1); });
