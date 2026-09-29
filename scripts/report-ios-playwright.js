// report-ios-playwright.js — druk i pobieranie plików na iPadzie (WebKit udający iPada).
//
// Zgłoszenie Mateusza: na tablecie „klikam i nic się nie dzieje”. Dwie przyczyny w iOS:
//   • w apce z ekranu początkowego window.print() potrafi nic nie zrobić,
//   • arkusz „Udostępnij” wolno otworzyć tylko tuż po dotknięciu — po kilku sekundach
//     tworzenia PDF iOS go odrzuca, a zwykłe „pobierz” w apce często nic nie robi.
// Sprawdzamy:
//   1. „Drukuj / PDF”, gdy okno druku się NIE otworzyło → komunikat + PDF + pasek
//      „Drukuj / udostępnij”; dotknięcie paska otwiera arkusz z plikiem PDF,
//   2. „Drukuj / PDF”, gdy okno druku się otworzyło (beforeprint) → żadnego zapasowego PDF,
//      a klasa druku NIE znika od razu (iPad drukuje chwilę po powrocie z print()),
//   3. „Pobierz PDF” po czasie (brak świeżego dotknięcia) → pasek „Zapisz / udostępnij”,
//   4. eksport CSV na iPadzie też idzie przez arkusz / pasek, nie przez martwe „pobierz”.

const { webkit } = require("playwright");
const APP_URL = process.env.APP_URL || "http://127.0.0.1:4175/";
const SLEEP_SCALE = Math.min(1, Math.max(0.1, Number(process.env.SLEEP_SCALE || 1)));
const sleep = (ms) => new Promise((r) => setTimeout(r, Math.round(ms * SLEEP_SCALE)));
const sleepFixed = (ms) => new Promise((r) => setTimeout(r, ms));
const IPAD_UA = "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/18.0 Safari/605.1.15";

async function run() {
  const browser = await webkit.launch({ headless: true });
  const context = await browser.newContext({
    serviceWorkers: "block",
    viewport: { width: 820, height: 1180 },
    userAgent: IPAD_UA,
    hasTouch: true,
  });
  await context.addInitScript(() => {
    localStorage.setItem("introPlayed", "true");
    // iPadOS udaje Maca — rozpoznanie po ekranie dotykowym.
    Object.defineProperty(navigator, "maxTouchPoints", { get: () => 5 });
    Object.defineProperty(navigator, "standalone", { get: () => true });
    // Podstawione okno „Udostępnij” (headless WebKit go nie ma).
    window.__shared = [];
    window.__printMode = "ignore"; // "ignore" = iOS nie otwiera okna druku; "open" = otwiera
    // Zwykłe przypisanie do navigator.* WebKit ignoruje — trzeba podmienić na prototypie.
    Object.defineProperty(Navigator.prototype, "canShare", { configurable: true, value: (d) => !!(d && d.files && d.files.length) });
    Object.defineProperty(Navigator.prototype, "share", { configurable: true, value: async (d) => {
      // iOS odrzuca share() bez świeżego dotknięcia — test może to wymusić raz.
      if (window.__shareRejectOnce) { window.__shareRejectOnce = false; throw new DOMException("no activation", "NotAllowedError"); }
      window.__shared.push(d.files.map((f) => ({ name: f.name, type: f.type, size: f.size })));
    } });
    window.print = () => {
      window.__printCalls = (window.__printCalls || 0) + 1;
      if (window.__printMode === "open") window.dispatchEvent(new Event("beforeprint"));
    };
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
  await page.evaluate(() => {
    const ws = { A1: { t: "s", v: "Status" }, B1: { t: "s", v: "Kwota" }, C1: { t: "s", v: "Klient" } };
    for (let i = 0; i < 30; i++) {
      ws["A" + (i + 2)] = { t: "s", v: ["W toku", "Zakończone"][i % 2] };
      ws["B" + (i + 2)] = { t: "n", v: 100 + i, w: String(100 + i) };
      ws["C" + (i + 2)] = { t: "s", v: "Klient ąę " + i };
    }
    ws["!ref"] = "A1:C31";
    workbook = { SheetNames: ["S"], Sheets: { S: ws }, Props: {} };
    currentFileName = "ipad.xlsx";
    sheetSelect.replaceChildren();
    const o = document.createElement("option"); o.value = o.textContent = "S";
    sheetSelect.appendChild(o); sheetSelect.value = "S";
    document.getElementById("headerRow").value = "1";
    document.getElementById("autoHeaderRow").checked = false;
    document.getElementById("loadBtn").click();
  });
  await page.waitForFunction(() => typeof baseRows !== "undefined" && baseRows.length === 30);
  await sleep(400);

  const ready = () => page.evaluate(() => {
    const el = document.querySelector(".file-ready");
    return el && !el.classList.contains("hidden")
      ? { text: el.querySelector(".file-ready-text").textContent, btn: el.querySelector(".file-ready-go").textContent }
      : null;
  });

  // ── 1. Druk nie ruszył (iOS zignorował print()) → PDF + pasek „Drukuj / udostępnij” ──
  await page.evaluate(() => window.__report.open());
  await sleep(300);
  await page.tap("#rpPrintBtn");
  // Zapasowa ścieżka rusza po 1,2 s bez „beforeprint”. Jeśli dotknięcie jest jeszcze „świeże”,
  // arkusz „Udostępnij” otwiera się od razu; jeśli nie — pasek z przyciskiem.
  await page.waitForFunction(() => window.__shared.length > 0
    || !!document.querySelector(".file-ready:not(.hidden)"), null, { timeout: 15000 }).catch(() => {});
  const toastTxt = await page.evaluate(() => Array.from(document.querySelectorAll(".toast")).map((t) => t.textContent).join(" | "));
  check("komunikat, że okno druku się nie otworzyło", /nie otworzyło/.test(toastTxt), toastTxt);
  const r1 = await ready();
  if (r1) {
    check("druk zignorowany → pasek „Drukuj / udostępnij” z PDF", /\.pdf/.test(r1.text) && /Drukuj/.test(r1.btn), r1);
    await page.tap(".file-ready-go");
    await sleep(200);
    check("po udostępnieniu pasek znika", !(await ready()));
  }
  const shared1 = await page.evaluate(() => window.__shared.slice());
  check("druk zignorowany → PDF trafia do „Udostępnij” (tam jest „Drukuj”)", shared1.length === 1 && shared1[0][0].type === "application/pdf" && shared1[0][0].size > 5000, shared1);

  // ── 2. Okno druku się otworzyło → bez zapasowego PDF; klasa druku zostaje ──
  await page.evaluate(() => { window.__printMode = "open"; window.__shared = []; });
  await page.tap("#rpPrintBtn");
  const clsNow = await page.evaluate(() => document.body.classList.contains("rp-printing"));
  check("klasa druku zostaje po powrocie z print() (iPad drukuje później)", clsNow, clsNow);
  await sleepFixed(1600);
  check("druk ruszył → brak zapasowego PDF", !(await ready()) && (await page.evaluate(() => window.__shared.length)) === 0);
  await page.evaluate(() => window.dispatchEvent(new Event("afterprint")));
  check("po „afterprint” klasa druku zdjęta", !(await page.evaluate(() => document.body.classList.contains("rp-printing"))));

  // ── 3. Pobierz PDF bez świeżego dotknięcia → pasek „Zapisz / udostępnij” ──
  await page.evaluate(() => { window.__shared = []; });
  // Po kilku sekundach tworzenia PDF iOS odrzuca share() (NotAllowedError) → musi być pasek
  // z przyciskiem, którego dotknięcie otwiera arkusz.
  const act = "NotAllowedError";
  await page.evaluate(() => { window.__shareRejectOnce = true; });
  await page.evaluate(() => window.__report.downloadPdf());
  const r3 = await ready();
  check("Pobierz PDF → pasek „Zapisz / udostępnij”", r3 && /Zapisz/.test(r3.btn), { r3, act });
  await page.tap(".file-ready-go");
  await sleep(200);
  check("pasek → arkusz z PDF", (await page.evaluate(() => window.__shared.length)) === 1);
  await page.evaluate(() => window.__report.close());

  // ── 4. Eksport CSV na iPadzie → przez arkusz „Udostępnij” (albo pasek) ──
  await page.evaluate(() => { window.__shared = []; document.getElementById("exportCsvBtn").click(); });
  await sleep(200);
  await page.tap("#exportRunAction");
  await sleep(300);
  const r4 = await ready();
  const s4 = await page.evaluate(() => window.__shared.slice());
  const viaShare = s4.length === 1 && /\.csv$/.test(s4[0][0].name);
  const viaBar = r4 && /\.csv/.test(r4.text);
  check("CSV na iPadzie: arkusz „Udostępnij” albo pasek (nie martwe pobieranie)", viaShare || viaBar, { s4, r4 });

  check("brak błędów w konsoli", errors.length === 0, errors);
  await browser.close();
  const failed = results.filter((r) => !r.ok);
  if (failed.length) {
    console.log(`\n❌ report-ios: ${failed.length} z ${results.length} nie przeszło`);
    process.exit(1);
  }
  console.log(`\n✅ report-ios: ${results.length}/${results.length}`);
}

run().catch((e) => { console.error(e); process.exit(1); });
