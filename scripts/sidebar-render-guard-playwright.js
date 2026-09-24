// sidebar-render-guard-playwright.js — strażnik: w panelu bocznym NIC nie może mieć
// `content-visibility: auto`.
//
// Dlaczego: na iOS sekcje panelu bywały po wysunięciu BIAŁE (rozwinięte, a puste —
// pomagało dopiero zwinięcie i rozwinięcie sekcji). Panel wjeżdża transformem spoza
// ekranu, a WebKit nie sprawdzał ponownie widoczności sekcji pominiętych przez
// content-visibility i trzymał ich stary, zapamiętany rozmiar. Odtwarza się tylko na
// prawdziwym iOS (symulator), NIE w WebKicie Playwrighta — dlatego pilnujemy przyczyny.
//
// Co musi być prawdą: po wczytaniu pliku i otwarciu panelu żaden element w .sidebar
// nie ma obliczonego content-visibility innego niż „visible".

const { chromium, webkit } = require("playwright");

const APP_URL = process.env.APP_URL || "http://127.0.0.1:4175/";
const ENGINE = process.env.ENGINE === "webkit" ? webkit : chromium;
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

async function run() {
  const browser = await ENGINE.launch({ headless: true });
  const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 390, height: 844 }, deviceScaleFactor: 2, hasTouch: true, isMobile: ENGINE === chromium });
  await context.addInitScript(() => localStorage.setItem("introPlayed", "true"));
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push("pageerror: " + e.message));

  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  await page.evaluate(() => document.getElementById("loadSampleBtn").click());
  await page.waitForFunction(() => document.querySelectorAll("#dataTable tbody tr").length > 5, null, { timeout: 30000 });
  await sleep(1500);
  await page.evaluate(() => setSidebarOpen(true));
  await sleep(800);

  const found = await page.evaluate(() => {
    const out = [];
    for (const el of document.querySelectorAll(".sidebar, .sidebar *")) {
      const cv = getComputedStyle(el).getPropertyValue("content-visibility");
      if (cv && cv !== "visible") out.push(`${el.tagName.toLowerCase()}${el.id ? "#" + el.id : ""}.${String(el.className).split(" ")[0]} = ${cv}`);
    }
    return { out, total: document.querySelectorAll(".sidebar .panel").length };
  });

  const results = [
    { name: "panel ma sekcje (warunek testu)", ok: found.total > 3, detail: `sekcji: ${found.total}` },
    { name: "żaden element panelu nie ma content-visibility ≠ visible", ok: found.out.length === 0, detail: found.out.slice(0, 5).join("; ") },
    { name: "brak błędów strony", ok: errors.length === 0, detail: errors.slice(0, 3).join(" | ") },
  ];
  await browser.close();

  let failed = 0;
  for (const r of results) {
    console.log(`${r.ok ? "✅" : "❌"} ${r.name}${r.detail ? `  (${r.detail})` : ""}`);
    if (!r.ok) failed += 1;
  }
  if (failed) { console.error(`\n${failed} z ${results.length} nie przeszło`); process.exit(1); }
  console.log(`\n✅ sidebar-render-guard (${process.env.ENGINE || "chromium"}): ${results.length}/${results.length}`);
}

run().catch((e) => { console.error(e); process.exit(1); });
