// pwa-offline-playwright.js — zachowanie service workera (sw.js) na PRAWDZIWYM SW.
//
// Własny serwer z logiem żądań i trybem „zawieszonej sieci", adres pwa.localhost (nie
// localhost — tam SW celowo zostawia stale-while-revalidate dla wygody developmentu).
//
// Co musi być prawdą:
//   1. drugie otwarcie NIE pyta serwera o pliki z ?v= (są niezmienne, idą z cache),
//   2. „jest kreska, a nic nie przechodzi" (serwer nie odpowiada na nawigację):
//      apka startuje z cache po ~3 s, a nie wisi na białym ekranie,
//   3. pełny offline: apka startuje, wczytuje plik (xlsx z cache) i pokazuje tabelę,
//   4. worker buildRows ładuje się offline (importScripts z wersją = adres z precache).

const { chromium } = require("playwright");
const http = require("http");
const fs = require("fs");
const path = require("path");

const ROOT = path.join(__dirname, "..");
const PORT = 4190;
const HOST = "pwa.localhost"; // *.localhost = bezpieczny kontekst (SW działa), ale nie „localhost" z sw.js
const ORIGIN = `http://${HOST}:${PORT}`;
const FILE = path.join(__dirname, "stress-test-workbench.xlsx");
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

const TYPES = { ".html": "text/html; charset=utf-8", ".js": "text/javascript", ".css": "text/css", ".json": "application/json", ".png": "image/png", ".webp": "image/webp", ".woff2": "font/woff2", ".mp4": "video/mp4", ".svg": "image/svg+xml" };
let log = [];
let hangNavigations = false;
const hung = [];
const server = http.createServer((req, res) => {
  const url = new URL(req.url, ORIGIN);
  log.push(url.pathname + url.search);
  const isNav = url.pathname === "/" || url.pathname.endsWith(".html");
  if (hangNavigations && isNav) { hung.push(res); return; } // nigdy nie odpowiada
  let file = path.join(ROOT, decodeURIComponent(url.pathname));
  if (url.pathname === "/") file = path.join(ROOT, "index.html");
  if (!file.startsWith(ROOT) || !fs.existsSync(file) || fs.statSync(file).isDirectory()) { res.writeHead(404); res.end(); return; }
  res.writeHead(200, { "Content-Type": TYPES[path.extname(file)] || "application/octet-stream", "Cache-Control": "no-store" });
  fs.createReadStream(file).pipe(res);
});

async function run() {
  await new Promise((r) => server.listen(PORT, "127.0.0.1", r));
  const browser = await chromium.launch({
    headless: true,
    args: [`--host-resolver-rules=MAP ${HOST} 127.0.0.1`],
  });
  const context = await browser.newContext({ viewport: { width: 900, height: 800 } });
  await context.addInitScript(() => localStorage.setItem("introPlayed", "true"));
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push("pageerror: " + e.message));

  const results = [];
  const check = (name, ok, detail) => results.push({ name, ok: !!ok, detail });

  // Pierwsze wejście: instalacja SW + precache powłoki, potem dogranie ciężkich zasobów.
  await page.goto(ORIGIN + "/", { waitUntil: "load" });
  await page.evaluate(() => navigator.serviceWorker.ready);
  await page.reload({ waitUntil: "load" }); // teraz strona jest kontrolowana
  check("SW kontroluje stronę", await page.evaluate(() => !!navigator.serviceWorker.controller));
  // xlsx/jszip trafiają do cache przy pierwszym użyciu (albo po 8 s z activate)
  await page.evaluate(() => ensureXlsxLibs(false));
  await page.evaluate(() => { const wb = XLSX.utils.book_new(); return !!wb; });
  await sleep(1500);

  // 1) Drugie otwarcie: zero zapytań o pliki z ?v=
  log = [];
  await page.reload({ waitUntil: "load" });
  await sleep(1500);
  const versioned = log.filter((u) => /[?&]v=/.test(u) && !u.startsWith("/sw.js"));
  check("ponowne otwarcie nie pyta sieci o pliki z ?v=", versioned.length === 0, versioned.slice(0, 5).join(", ") || `zapytań łącznie: ${log.length}`);

  // 2) Zawieszona sieć przy nawigacji → start z cache po limicie
  hangNavigations = true;
  const t0 = Date.now();
  await page.reload({ waitUntil: "load", timeout: 15000 });
  const waited = Date.now() - t0;
  const booted = await page.evaluate(() => !!document.getElementById("fileInput") && typeof ensureXlsxLibs === "function");
  check("zawieszona sieć: start z cache w < 6 s", booted && waited < 6000, `${waited} ms`);
  hangNavigations = false;
  hung.splice(0).forEach((res) => { try { res.destroy(); } catch (_) {} });

  // 3) + 4) Pełny offline: start, wczytanie pliku, tabela, worker buildRows
  await context.setOffline(true);
  await page.reload({ waitUntil: "load" });
  check("offline: apka startuje", await page.evaluate(() => !!document.getElementById("fileInput")));
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  await page.setInputFiles("#fileInput", FILE);
  await page.waitForFunction(() => document.getElementById("sheetSelect")?.options?.length > 0, null, { timeout: 20000 }).catch(() => {});
  await page.click("#loadBtn").catch(() => {});
  await page.waitForFunction(() => document.querySelectorAll("#dataTable tbody tr").length > 5, null, { timeout: 20000 }).catch(() => {});
  check("offline: plik wczytany, tabela jest", await page.evaluate(() => document.querySelectorAll("#dataTable tbody tr").length > 5));
  const workerOk = await page.evaluate(() => new Promise((resolve) => {
    const v = typeof APP_BUILD_VERSION !== "undefined" ? APP_BUILD_VERSION : "1";
    const w = new Worker(`app/build-rows-worker.js?v=${encodeURIComponent(v)}`);
    const done = (ok, why) => { try { w.terminate(); } catch (_) {} resolve({ ok, why }); };
    w.onerror = (e) => done(false, e.message || "error");
    w.onmessage = (e) => done(!!(e.data && e.data.id === -1), JSON.stringify(e.data).slice(0, 80));
    w.postMessage({ type: "buildRows", id: -1, sheet: "x", headerRow: 1, workbook: { SheetNames: [], Sheets: {} }, options: {} });
    setTimeout(() => done(false, "timeout"), 8000);
  }));
  check("offline: worker buildRows się ładuje", workerOk.ok, workerOk.why);
  await context.setOffline(false);

  check("brak błędów strony", errors.length === 0, errors.slice(0, 3).join(" | "));
  await browser.close();
  server.close();

  let failed = 0;
  for (const r of results) {
    console.log(`${r.ok ? "✅" : "❌"} ${r.name}${r.detail ? `  (${r.detail})` : ""}`);
    if (!r.ok) failed += 1;
  }
  if (failed) { console.error(`\n${failed} z ${results.length} nie przeszło`); process.exit(1); }
  console.log(`\n✅ PWA: ${results.length}/${results.length}`);
}

run().catch((e) => { console.error(e); server.close(); process.exit(1); });
