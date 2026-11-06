// launch-files-playwright.js — „OTWÓRZ ZA POMOCĄ” I „UDOSTĘPNIJ” (2026-10-31, jak w Documents Workbench).
//   node scripts/launch-files-playwright.js
//
// Pilnuje:
//   1. manifest: file_handlers (.xlsx/.xlsm), launch_handler navigate-new, share_target;
//   2. plik z launchQueue (Otwórz za pomocą na Macu/Windows) otwiera się Z UCHWYTEM
//      (Zapisz pisze do oryginału) i pokazuje tabelę;
//   3. przy niezapisanych zmianach apka pyta przed podmianą pliku (odmowa = nic się nie zmienia);
//   4. Android „Udostępnij”: POST na ./share-target → service worker → ./?open=shared otwiera plik.
// Uruchom z serwerem na APP_URL (domyślnie http://127.0.0.1:4175/).

const playwright = require("playwright");

const APP_URL = process.env.APP_URL || "http://127.0.0.1:4175/";
const SLEEP_SCALE = Math.min(1, Math.max(0.1, Number(process.env.SLEEP_SCALE || 1)));
const sleep = (ms) => new Promise((r) => setTimeout(r, Math.round(ms * SLEEP_SCALE)));

async function run() {
  const browser = await playwright.chromium.launch({ headless: true });
  const fails = [];
  const check = (name, ok, detail) => {
    console.log(`${ok ? "✅" : "❌"} ${name}${ok ? "" : ` — ${JSON.stringify(detail)}`}`);
    if (!ok) fails.push(name);
  };

  // 1. manifest
  const ctx0 = await browser.newContext({ serviceWorkers: "block" });
  const p0 = await ctx0.newPage();
  const man = await (await p0.request.get(new URL("manifest.json", APP_URL).href)).json();
  const exts = Object.values(man.file_handlers?.[0]?.accept || {}).flat();
  check("1. manifest: file_handlers dla .xlsx i .xlsm", exts.includes(".xlsx") && exts.includes(".xlsm"), man.file_handlers);
  check("1. manifest: każdy plik we własnym oknie", man.launch_handler?.client_mode === "navigate-new", man.launch_handler);
  check("1. manifest: share_target na ./share-target", man.share_target?.action === "./share-target" && man.share_target?.method === "POST", man.share_target);
  await ctx0.close();

  // 2–3. launchQueue (atrapa: przeglądarka daje ją tylko zainstalowanej aplikacji)
  const ctx = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1280, height: 900 } });
  await ctx.addInitScript(() => {
    localStorage.setItem("introPlayed", "true");
    // Chromium ma prawdziwe launchQueue (nie do podmiany) — przechwytujemy setConsumer
    if (window.LaunchQueue) window.LaunchQueue.prototype.setConsumer = function (fn) { window.__launchConsumer = fn; };
    else {
      window.LaunchParams = function LaunchParams() {};
      window.LaunchParams.prototype.files = [];
      window.launchQueue = { setConsumer(fn) { window.__launchConsumer = fn; } };
    }
  });
  const p = await ctx.newPage();
  const errors = [];
  p.on("pageerror", (e) => errors.push(e.message));
  await p.goto(APP_URL + "?open=file", { waitUntil: "load" });
  await p.waitForFunction(() => typeof window.__launchConsumer === "function", null, { timeout: 10000 });
  check("2. strona odbiera pliki z launchQueue", true);
  const launch = (name) => p.evaluate(async (n) => {
    document.getElementById("heroSplash")?.remove();
    await ensureXlsxLibs(true);
    const buf = await buildSampleWorkbookArrayBuffer();
    const file = new File([buf], n, { type: XLSX_MIME });
    window.__handle = { kind: "file", name: n, getFile: async () => file };
    await window.__launchConsumer({ files: [window.__handle] });
  }, name);
  await launch("Z-Findera.xlsx");
  await p.waitForFunction(() => document.querySelector("#dataTable tbody tr[data-row-key]"), null, { timeout: 20000 });
  const s2 = await p.evaluate(() => ({ name: currentFileName, handle: currentFileHandle === window.__handle, url: location.search, title: document.title }));
  check("2. plik otwarty z uchwytem (Zapisz do oryginału) i tabelą", s2.name === "Z-Findera.xlsx" && s2.handle, s2);
  check("2. ?open=file znika z adresu (odświeżenie nie otwiera drugi raz)", s2.url === "", s2.url);

  await p.evaluate(() => setDirtyState(true));
  p.once("dialog", (d) => d.dismiss());
  await launch("Drugi.xlsx");
  await sleep(300);
  const s3 = await p.evaluate(() => currentFileName);
  check("3. niezapisane zmiany + odmowa → plik zostaje", s3 === "Z-Findera.xlsx", s3);
  check("bez błędów strony (launchQueue)", errors.length === 0, errors);
  await ctx.close();

  // 4. share_target przez service workera
  const ctx2 = await browser.newContext({ viewport: { width: 400, height: 800 } });
  await ctx2.addInitScript(() => localStorage.setItem("introPlayed", "true"));
  const q = await ctx2.newPage();
  const errors2 = [];
  q.on("pageerror", (e) => errors2.push(e.message));
  await q.goto(APP_URL, { waitUntil: "load" });
  await q.waitForFunction(() => navigator.serviceWorker && navigator.serviceWorker.controller, null, { timeout: 20000 }).catch(() => {});
  if (!(await q.evaluate(() => !!navigator.serviceWorker.controller))) {
    await q.reload({ waitUntil: "load" });
    await q.waitForFunction(() => navigator.serviceWorker.controller, null, { timeout: 20000 });
  }
  const shared = await q.evaluate(async () => {
    await ensureXlsxLibs(true);
    const buf = await buildSampleWorkbookArrayBuffer();
    const fd = new FormData();
    fd.append("file", new File([buf], "Udostępniony.xlsx", { type: XLSX_MIME }));
    const res = await fetch("./share-target", { method: "POST", body: fd });
    return { url: res.url, redirected: res.redirected };
  });
  check("4. POST na share-target → przekierowanie na ?open=shared", shared.redirected && /\?open=shared$/.test(shared.url), shared);
  // pytanie przed otwarciem (POST może wysłać też obca strona — audyt 6.11)
  const asked = [];
  q.once("dialog", (d) => { asked.push(d.message()); d.accept(); });
  await q.goto(APP_URL + "?open=shared", { waitUntil: "load" });
  await q.waitForFunction(() => typeof currentFileName === "string" && currentFileName === "Udostępniony.xlsx", null, { timeout: 20000 }).catch(() => {});
  const s4 = await q.evaluate(async () => ({ name: currentFileName, handle: currentFileHandle, left: (await caches.keys()).includes("excel-wb-share") ? !!(await (await caches.open("excel-wb-share")).match("./__shared-file")) : false }));
  check("4. udostępniony plik otwarty (kopia, bez uchwytu), cache wyczyszczony", s4.name === "Udostępniony.xlsx" && !s4.handle && !s4.left, s4);
  check("4. pytanie przed otwarciem z nazwą pliku", asked.length === 1 && asked[0].includes("Udostępniony.xlsx"), asked);
  check("bez błędów strony (share)", errors2.length === 0, errors2);

  // 5. Chrome otworzył aplikację „do pliku”, ale pliku nie przekazał (zapamiętane „Nie zezwalaj”) —
  // zamiast pustego ekranu bez słowa: ostrzeżenie z podpowiedzią (zgłoszenie 2026-11-02, nagranie)
  const ctx5 = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1280, height: 900 } });
  await ctx5.addInitScript(() => {
    localStorage.setItem("introPlayed", "true");
    if (window.LaunchQueue) window.LaunchQueue.prototype.setConsumer = function (fn) { window.__launchConsumer = fn; };
  });
  const r = await ctx5.newPage();
  await r.goto(APP_URL + "?open=file", { waitUntil: "load" });
  const warned = await r.waitForFunction(() => /bez pliku|without the file/.test(document.getElementById("toastContainer")?.textContent || ""), null, { timeout: 8000 }).then(() => true).catch(() => false);
  check("5. otwarcie bez pliku → ostrzeżenie zamiast pustego ekranu", warned, "brak komunikatu");
  await ctx5.close();

  await browser.close();
  if (fails.length) {
    console.error(`\n${fails.length} błędów`);
    process.exit(1);
  }
  console.log("\nOK");
}

run().catch((e) => { console.error(e); process.exit(1); });
