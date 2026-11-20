// update-announce-playwright.js — dymek „Zaktualizowano do vX / Co nowego” dopiero PO aktualizacji
// (zgłoszenie Mateusza 2026-11-20: pokazywał się przed kliknięciem „Aktualizuj”).
//
// Przyczyna: index.html idzie z sieci, więc nowy kod startuje jeszcze pod STARYM service workerem
// (obok przycisku „Aktualizuj”). Dymek ma czekać, aż stronę obsługuje SW tej samej wersji.
//
//   1. stary SW steruje stroną (sw.js?v=stara) → brak dymka, wersja nie jest odhaczona,
//   2. po aktualizacji (SW tej wersji) → dymek „Zaktualizowano do vX” z „Co nowego”,
//   3. bez SW (np. przeglądarka bez obsługi) → dymek jak dotąd.
//
// Uruchom z serwerem na APP_URL (domyślnie http://127.0.0.1:4175/). ENGINE=webkit — Safari.

const playwright = require("playwright");

const APP_URL = process.env.APP_URL || "http://127.0.0.1:4175/";
const ENGINE = process.env.ENGINE || "chromium";
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

// controller: null = brak SW; "current" = SW tej wersji; inaczej numer starej wersji
async function open(browser, controller) {
  const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1280, height: 800 } });
  await context.addInitScript((ctl) => {
    localStorage.setItem("introPlayed", "true");
    localStorage.setItem("swb-seen-version", "v0.00.01");
    if (ctl && navigator.serviceWorker) {
      Object.defineProperty(ServiceWorkerContainer.prototype, "controller", {
        configurable: true,
        get() {
          const v = ctl === "current" ? (typeof APP_BUILD_VERSION !== "undefined" ? APP_BUILD_VERSION : "") : ctl;
          return { scriptURL: `${location.origin}/sw.js?v=${v}`, postMessage() {} };
        },
      });
    }
  }, controller);
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  await page.goto(APP_URL, { waitUntil: "load" });
  await sleep(2200); // dymek startuje po 1,2 s
  const res = await page.evaluate(() => ({
    toast: [...document.querySelectorAll(".toast")].map((t) => t.textContent).join(" | "),
    seen: localStorage.getItem("swb-seen-version"),
    v: APP_BUILD_VERSION,
  }));
  await context.close();
  return { ...res, errors };
}

async function run() {
  const browser = await playwright[ENGINE].launch({ headless: true });
  const results = [];
  const check = (name, ok, detail = "") => results.push({ name, ok: !!ok, detail: typeof detail === "string" ? detail : JSON.stringify(detail) });

  const before = await open(browser, "v0.00.01");
  check("1. stary SW: brak dymka „Zaktualizowano”", !/Zaktualizowano|Updated to/.test(before.toast), before.toast);
  check("1. stary SW: wersja nie odhaczona (dymek poczeka na aktualizację)", before.seen === "v0.00.01", before.seen);

  const after = await open(browser, "current");
  check("2. po aktualizacji: dymek „Zaktualizowano do vX”", after.toast.includes(after.v) && /Zaktualizowano|Updated to/.test(after.toast), after.toast);
  check("2. po aktualizacji: wersja odhaczona", after.seen === after.v, after.seen);

  const noSw = await open(browser, null);
  check("3. bez SW: dymek jak dotąd", /Zaktualizowano|Updated to/.test(noSw.toast), noSw.toast);

  check("brak błędów strony", [before, after, noSw].every((x) => x.errors.length === 0), [before, after, noSw].flatMap((x) => x.errors).join(" | "));
  await browser.close();
  const failed = results.filter((r) => !r.ok);
  results.forEach((r) => console.log(`${r.ok ? "✓" : "✗"} ${r.name}${r.ok ? "" : "  — " + r.detail}`));
  if (failed.length) throw new Error(`${failed.length} z ${results.length} asercji nie przeszło (${ENGINE})`);
  console.log(`\n✅ update-announce-playwright OK — ${results.length} asercji (${ENGINE})`);
}

run().catch((e) => { console.error(e.message || e); process.exit(1); });
