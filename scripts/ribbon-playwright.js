// ribbon-playwright.js — strażnik wstążki (app/ribbon.js; docs/WSTAZKA.md, makieta docs/makiety/wstazka-v2.html).
//
// Etap 1 — rama:
//   1. na komputerze (≥ 769 px) body.rb-on, wstążka widoczna, nagłówek „hero” schowany,
//   2. 7 kart w kolejności z makiety, wybrana „Narzędzia główne”,
//   3. Szukaj, Zapisz, ⋯ i Cofnij to TE SAME elementy, przeniesione do wstążki,
//   4. pasek formuły tuż pod wstążką (≤ 82 px od góry),
//   5. strzałki przełączają karty i przenoszą fokus,
//   6. 768 px: wszystko wraca na swoje miejsca; 1440 px: znowu we wstążce,
//   7. 820 px: nic nie wystaje w bok, rząd kart ma strzałkę „+N”,
//   8. telefon: bez wstążki, układ jak dotąd.
//
// Uruchom z serwerem na APP_URL (domyślnie http://127.0.0.1:4175/). ENGINE=webkit — Safari.

const playwright = require("playwright");

const APP_URL = process.env.APP_URL || "http://127.0.0.1:4175/";
const ENGINE = process.env.ENGINE || "chromium";
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

async function open(browser, viewport, extra = {}) {
  const context = await browser.newContext({ serviceWorkers: "block", viewport, ...extra });
  await context.addInitScript(() => {
    if (sessionStorage.getItem("__rbInit")) return;
    sessionStorage.setItem("__rbInit", "1");
    localStorage.clear();
    localStorage.setItem("introPlayed", "true");
  });
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push("pageerror: " + e.message));
  page.on("console", (m) => { if (m.type() === "error") errors.push("console: " + m.text()); });
  await page.goto(APP_URL + "?sample=300", { waitUntil: "load" });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  await page.evaluate(() => document.getElementById("loadSampleBtn")?.click());
  await page.waitForFunction(() => document.querySelector("#dataTable tbody tr[data-row-key]"), null, { timeout: 30000 });
  await page.evaluate(() => { if (typeof setSidebarOpen === "function") setSidebarOpen(false); document.querySelectorAll(".toast").forEach((t) => t.remove()); });
  await sleep(400);
  return { context, page, errors };
}

const inRibbon = (page, sel) => page.evaluate((s) => !!document.querySelector(s)?.closest("#ribbon"), sel);
const shown = (page, sel) => page.evaluate((s) => {
  const el = document.querySelector(s);
  if (!el) return false;
  const r = el.getBoundingClientRect();
  return r.width > 0 && r.height > 0 && getComputedStyle(el).visibility !== "hidden";
}, sel);

async function run() {
  const browser = await playwright[ENGINE].launch({ headless: true });
  const results = [];
  const check = (name, ok, detail = "") => results.push({ name, ok: !!ok, detail: String(detail) });

  // ── komputer ────────────────────────────────────────────────────────────
  {
    const { context, page, errors } = await open(browser, { width: 1440, height: 900 });
    check("1. body.rb-on na komputerze", await page.evaluate(() => document.body.classList.contains("rb-on")));
    check("1. wstążka widoczna", await shown(page, "#ribbon"));
    check("1. nagłówek hero schowany", !(await shown(page, ".hero")));

    const tabs = await page.evaluate(() => Array.from(document.querySelectorAll('#ribbon [role="tab"][data-rtab]')).filter((t) => !t.hidden).map((t) => t.textContent.trim()));
    check("2. 7 kart w kolejności z makiety", JSON.stringify(tabs) === JSON.stringify(["Narzędzia główne", "Wstaw", "Dane", "Formuły", "Analiza", "Widok", "Asystent"]), tabs.join(" | "));
    const sel = await page.evaluate(() => document.querySelector('#ribbon [role="tab"][aria-selected="true"]')?.dataset.rtab);
    check("2. wybrana karta: Narzędzia główne", sel === "home", sel);

    for (const id of ["#quickSearch", "#heroSaveBtn", "#appMenuBtn", "#undoBtn", "#redoBtn"]) {
      check(`3. ${id} we wstążce`, await inRibbon(page, id));
    }
    check("3. Szukaj widoczne", await shown(page, "#quickSearch"));

    const fbTop = await page.evaluate(() => Math.round(document.getElementById("formulaBar").getBoundingClientRect().top));
    check("4. pasek formuły ≤ 82 px od góry", fbTop <= 82, fbTop + " px");

    await page.focus('#ribbon [role="tab"][aria-selected="true"]');
    await page.keyboard.press("ArrowRight");
    const after = await page.evaluate(() => ({ sel: document.querySelector('#ribbon [role="tab"][aria-selected="true"]')?.dataset.rtab, focus: document.activeElement?.dataset?.rtab }));
    check("5. strzałka w prawo → Wstaw (wybór i fokus)", after.sel === "ins" && after.focus === "ins", JSON.stringify(after));
    const panelOk = await page.evaluate(() => {
      const ps = Array.from(document.querySelectorAll("#ribbon .rb-panel[data-rtab]"));
      return ps.length === 7 && ps.filter((p) => !p.hidden).map((p) => p.dataset.rtab).join() === "ins";
    });
    check("5. widać tylko panel karty Wstaw", panelOk);

    await page.setViewportSize({ width: 768, height: 900 });
    await sleep(300);
    check("6. 768 px: bez rb-on", !(await page.evaluate(() => document.body.classList.contains("rb-on"))));
    check("6. 768 px: Szukaj wróciło na miejsce", !(await inRibbon(page, "#quickSearch")) && (await page.evaluate(() => !!document.getElementById("quickSearch").closest("#quickSearchWrap")?.closest(".table-panel"))));
    check("6. 768 px: Zapisz wróciło do nagłówka", await page.evaluate(() => !!document.getElementById("heroSaveBtn").closest(".hero")));
    check("6. 768 px: nagłówek widoczny", await shown(page, ".hero"));
    await page.setViewportSize({ width: 1440, height: 900 });
    await sleep(300);
    check("6. z powrotem 1440 px: Szukaj we wstążce", await inRibbon(page, "#quickSearch"));

    await page.setViewportSize({ width: 820, height: 900 });
    await sleep(400);
    const ov = await page.evaluate(() => document.documentElement.scrollWidth - innerWidth);
    check("7. 820 px: nic nie wystaje w bok", ov <= 0, ov + " px");
    const nudge = await page.evaluate(() => {
      const n = document.querySelector('#ribbon [data-strip="tabs"] .rb-nudge-r');
      return n ? { vis: n.getBoundingClientRect().width > 0, txt: n.textContent.trim() } : null;
    });
    check("7. 820 px: rząd kart ma strzałkę „+N”", nudge && nudge.vis && /^\+\d+/.test(nudge.txt), JSON.stringify(nudge));

    check("brak błędów strony (komputer)", errors.length === 0, errors.join(" | "));
    await context.close();
  }

  // ── telefon ─────────────────────────────────────────────────────────────
  {
    const { context, page, errors } = await open(browser, { width: 390, height: 844 }, { hasTouch: true, isMobile: ENGINE === "chromium" });
    check("8. telefon: bez rb-on", !(await page.evaluate(() => document.body.classList.contains("rb-on"))));
    check("8. telefon: wstążka ukryta", !(await shown(page, "#ribbon")));
    check("8. telefon: Szukaj na swoim miejscu", !(await inRibbon(page, "#quickSearch")));
    check("brak błędów strony (telefon)", errors.length === 0, errors.join(" | "));
    await context.close();
  }

  await browser.close();
  const failed = results.filter((r) => !r.ok);
  results.forEach((r) => console.log(`${r.ok ? "✓" : "✗"} ${r.name}${r.ok ? "" : "  — " + r.detail}`));
  if (failed.length) throw new Error(`${failed.length} z ${results.length} asercji nie przeszło (${ENGINE})`);
  console.log(`\n✅ ribbon-playwright OK — ${results.length} asercji (${ENGINE})`);
}

run().catch((e) => { console.error(e.message || e); process.exit(1); });
