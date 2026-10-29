// flip-ghost-playwright.js — strażnik animacji sortowania/filtra przez „duszki” (core.js, flipGhostAnimate).
//
// Wiersze w ruchu to kopie w warstwie nad tabelą; prawdziwe stoją już na miejscu, przezroczyste.
// Najważniejsza gwarancja: KONIEC animacji jest niewidoczny — kopie w stanie końcowym dają
// obraz taki sam jak prawdziwa tabela (inaczej na końcu byłoby widać przeskok); dopuszczalne
// jest tylko inne wygładzanie liter na osobnej warstwie GPU (≤ 8/255 na ≤ 15 % obrazu).
// Do tego: animacja naprawdę idzie przez kopie (i bez przeliczeń stylu co klatkę w Chromium),
// klik / klawisz / przebudowa tabeli kończą ją od razu, nic nie zostaje po czasie, ruch
// wyłączony przy prefers-reduced-motion, brak błędów strony.
//
// Przypadki obrazu: jasny i ciemny motyw, telefon, zawijanie tekstu, zamrożona kolumna
// przewinięta w bok, zaznaczona komórka w przesuwanym wierszu, filtr (wiersze wjeżdżające).
//
//   node scripts/flip-ghost-playwright.js        ENGINE=webkit node scripts/flip-ghost-playwright.js
const playwright = require("playwright");
const path = require("path");

const APP_URL = process.env.APP_URL || "http://127.0.0.1:4175/";
const ENGINE = process.env.ENGINE || "chromium";
const FILE = path.join(__dirname, "RODO Obieg terenów 2026 V5.xlsx");
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

const fails = [];
const check = (name, ok, detail = "") => {
  console.log(`${ok ? "✅" : "❌"} ${name}${ok ? "" : ` — ${detail}`}`);
  if (!ok) fails.push(name);
};

async function openApp(browser, { mobile = false, reducedMotion = "no-preference" } = {}) {
  const ctx = await browser.newContext({
    serviceWorkers: "block",
    viewport: mobile ? { width: 390, height: 844 } : { width: 1280, height: 900 },
    deviceScaleFactor: 1,
    hasTouch: mobile && ENGINE !== "firefox",
    isMobile: mobile && ENGINE === "chromium",
    reducedMotion,
  });
  await ctx.addInitScript(() => localStorage.setItem("introPlayed", "true"));
  const page = await ctx.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  await page.setInputFiles("#fileInput", FILE);
  await page.waitForFunction(() => document.getElementById("sheetSelect")?.options?.length > 0, null, { timeout: 30000 });
  await page.evaluate(() => loadBtn.click());
  await page.waitForFunction(() => document.querySelectorAll("#dataTable tbody tr").length > 20, null, { timeout: 30000 });
  await sleep(1500);
  await page.evaluate(() => { try { setSidebarOpen(false); } catch (_) {} });
  await sleep(500);
  return { ctx, page, errors };
}

// Uruchom akcję, zatrzymaj animację kopii w stanie końcowym, zrób zrzut; potem zakończ i zrzut drugi.
async function endFrameMatches(page, action) {
  await page.evaluate(() => {
    // tylko animacje kopii lecą przez Web Animations — inne (CSS) dobijamy, żeby nie mieszały w obrazie
    window.__ghostSeen = 0;
  });
  await page.evaluate(() => { const t = document.getElementById("toastContainer"); if (t) t.style.visibility = "hidden"; });
  await action();
  await page.waitForFunction(() => typeof _flipGhost !== "undefined" && !!_flipGhost, null, { timeout: 1500 }).catch(() => {});
  const info = await page.evaluate(async () => {
    const g = typeof _flipGhost !== "undefined" ? _flipGhost : null;
    if (!g) return { ghosts: 0 };
    clearTimeout(g.timer);
    g.anims.forEach((a) => { a.pause(); a.currentTime = a.effect.getComputedTiming().endTime; });
    document.getAnimations().forEach((a) => { if (!g.anims.includes(a)) { try { a.finish(); } catch (_) {} } });
    await new Promise((r) => requestAnimationFrame(() => requestAnimationFrame(r)));
    return { ghosts: g.layer.querySelectorAll(".flip-ghost").length, hidden: g.rows.length };
  });
  if (!info.ghosts) return { info, same: null };
  // przyklejone podpowiedzi / kursor mogą migać — chowamy je na czas obu zrzutów
  await page.mouse.move(2, 2);
  const clip = await page.evaluate(() => { const r = tableWrapEl.getBoundingClientRect(); return { x: Math.max(0, r.left), y: Math.max(0, r.top), width: Math.min(r.width, innerWidth - Math.max(0, r.left)), height: Math.min(r.height, innerHeight - Math.max(0, r.top)) }; });
  const a = await page.screenshot({ clip, animations: "allow" });
  await page.evaluate(() => { flipGhostFinish(); });
  await page.evaluate(() => new Promise((r) => requestAnimationFrame(() => requestAnimationFrame(r))));
  const b = await page.screenshot({ clip, animations: "allow" });
  if (a.equals(b)) return { info, same: true, a, b, stat: "identyczne bajt w bajt" };
  // Kopia leży na osobnej warstwie GPU, więc wygładzanie liter może się różnić o 1–5/255 na
  // ułamku pikseli (niewidoczne). Prawdziwa usterka — inne tło, brak kreski, przesunięcie —
  // daje różnice dziesiątek/setek jednostek. Próg: kanał ≤ 8/255 i różnice na ≤ 15 % obrazu (zły odcień tła kopii = lekka różnica na CAŁYM obszarze; wygładzanie liter: 0,2–6 %).
  const stat = await page.evaluate(async ([a64, b64]) => {
    const img = (src) => new Promise((res, rej) => { const i = new Image(); i.onload = () => res(i); i.onerror = rej; i.src = "data:image/png;base64," + src; });
    const [ia, ib] = await Promise.all([img(a64), img(b64)]);
    const px = (i) => { const c = document.createElement("canvas"); c.width = i.width; c.height = i.height; const x = c.getContext("2d"); x.drawImage(i, 0, 0); return x.getImageData(0, 0, i.width, i.height).data; };
    const da = px(ia), db = px(ib);
    let diff = 0, max = 0;
    for (let k = 0; k < da.length; k += 4) {
      const d = Math.max(Math.abs(da[k] - db[k]), Math.abs(da[k + 1] - db[k + 1]), Math.abs(da[k + 2] - db[k + 2]));
      if (d) { diff++; if (d > max) max = d; }
    }
    return { diff, max, total: da.length / 4 };
  }, [a.toString("base64"), b.toString("base64")]);
  const same = stat.max <= 8 && stat.diff <= stat.total * 0.15;
  return { info, same, a, b, stat: `${stat.diff} pikseli różni się o maks. ${stat.max}/255` };
}

const sortClick = (page, n = 2) => page.evaluate((n) => { document.querySelectorAll("#dataTable thead th[data-sort-header]")[n].click(); }, n);

async function scenarioImages(browser, label, opts, setup) {
  const { ctx, page, errors } = await openApp(browser, opts);
  if (setup) await setup(page);
  await sleep(400);
  for (const [name, act] of [["sort", () => sortClick(page)], ["sort odwrotnie", () => sortClick(page)]]) {
    const r = await endFrameMatches(page, act);
    check(`${label}: ${name} — animacja przez kopie (${r.info.ghosts || 0} wierszy)`, r.info.ghosts > 0, "brak warstwy kopii");
    if (r.same !== null) {
      if (!r.same && process.env.DUMP) {
        require("fs").writeFileSync(`/tmp/flip-${label.replace(/\W+/g, "_")}-${name.replace(/\W+/g, "_")}-ghost.png`, r.a);
        require("fs").writeFileSync(`/tmp/flip-${label.replace(/\W+/g, "_")}-${name.replace(/\W+/g, "_")}-real.png`, r.b);
      }
      check(`${label}: ${name} — koniec animacji = tabela (${r.stat})`, r.same, "kopie różnią się od prawdziwych wierszy (DUMP=1 zapisze obrazy w /tmp)");
    }
    await sleep(300);
  }
  check(`${label}: bez błędów strony`, !errors.length, errors.join(" | "));
  await ctx.close();
}

(async () => {
  const browser = await playwright[ENGINE].launch({ headless: true });

  await scenarioImages(browser, "desktop", {});
  await scenarioImages(browser, "ciemny motyw", {}, (p) => p.evaluate(() => document.documentElement.setAttribute("data-theme", "dark")));
  await scenarioImages(browser, "telefon", { mobile: true });
  await scenarioImages(browser, "zawijanie tekstu", {}, (p) => p.evaluate(() => { const el = document.getElementById("wrapCells") || document.querySelector("#wrapCellsToggle, input[id*=wrap]"); tableEl.classList.add("wrap-cells"); if (el) { el.checked = true; } }));
  await scenarioImages(browser, "zamrożona kolumna, przewinięta w bok", {}, async (p) => {
    await p.evaluate(() => { tableWrapEl.classList.add("freeze-first-col"); tableWrapEl.scrollLeft = 260; });
  });
  await scenarioImages(browser, "zaznaczona komórka", {}, async (p) => {
    await p.click("#dataTable tbody tr:nth-child(5) td:nth-child(4)");
    await sleep(300);
  });

  // filtr: wiersze wjeżdżają (opacity + przesunięcie) — koniec też musi być identyczny
  {
    const { ctx, page, errors } = await openApp(browser, {});
    await page.fill("#quickSearch", "ann");
    await page.press("#quickSearch", "Enter");
    await sleep(1200);
    const r = await endFrameMatches(page, async () => { await page.fill("#quickSearch", ""); await page.press("#quickSearch", "Enter"); });
    check(`filtr wyczyszczony — wiersze wjeżdżają przez kopie (${r.info.ghosts || 0})`, r.info.ghosts > 0, "brak warstwy kopii");
    if (r.same !== null) check(`filtr — koniec animacji = tabela (${r.stat})`, r.same, "różnica obrazu");

    // klik w trakcie animacji kończy ją od razu, wiersze wracają
    await sortClick(page);
    const during = await page.evaluate(() => !!_flipGhost && document.querySelectorAll(".flip-ghosts").length);
    await page.click("#dataTable tbody tr:nth-child(3) td:nth-child(3)");
    const after = await page.evaluate(() => ({ ghost: !!_flipGhost, layers: document.querySelectorAll(".flip-ghosts").length, hidden: document.querySelectorAll("tr.flip-hidden").length }));
    check("klik w trakcie animacji kończy ją od razu", during && !after.ghost && !after.layers && !after.hidden, JSON.stringify({ during, after }));

    // klawisz w trakcie
    await sortClick(page);
    await page.keyboard.press("ArrowDown");
    const afterKey = await page.evaluate(() => ({ ghost: !!_flipGhost, hidden: document.querySelectorAll("tr.flip-hidden").length }));
    check("klawisz w trakcie animacji kończy ją od razu", !afterKey.ghost && !afterKey.hidden, JSON.stringify(afterKey));

    // samo zakończenie: po czasie nic nie zostaje
    await sortClick(page);
    await sleep(1200);
    const left = await page.evaluate(() => ({ ghost: !!_flipGhost, layers: document.querySelectorAll(".flip-ghosts").length, hidden: document.querySelectorAll("tr.flip-hidden").length, inline: [...document.querySelectorAll("#dataTable tbody tr")].filter((tr) => tr.style.transform || tr.style.opacity).length }));
    check("po animacji: bez kopii, bez ukrytych wierszy, bez resztek stylu", !left.ghost && !left.layers && !left.hidden && !left.inline, JSON.stringify(left));

    // szybkie dwa sortowania pod rząd — drugie sprząta po pierwszym
    await sortClick(page);
    await sleep(60);
    await sortClick(page);
    await sleep(1200);
    const twice = await page.evaluate(() => ({ layers: document.querySelectorAll(".flip-ghosts").length, hidden: document.querySelectorAll("tr.flip-hidden").length }));
    check("dwa sortowania pod rząd — nic nie zostaje", !twice.layers && !twice.hidden, JSON.stringify(twice));

    // Chromium: w trakcie animacji kopii nie ma przeliczeń stylu co klatkę (to był cały koszt)
    if (ENGINE === "chromium") {
      const cdp = await ctx.newCDPSession(page);
      const events = [];
      cdp.on("Tracing.dataCollected", (d) => events.push(...d.value));
      const done = new Promise((res) => cdp.once("Tracing.tracingComplete", res));
      await cdp.send("Tracing.start", { categories: "devtools.timeline,disabled-by-default-devtools.timeline.invalidationTracking", transferMode: "ReportEvents" });
      await sortClick(page);
      const nAnims = await page.evaluate(() => (_flipGhost ? _flipGhost.anims.length : 0));
      await sleep(900);
      await cdp.send("Tracing.end");
      await done;
      const animInv = events.filter((e) => e.name === "StyleRecalcInvalidationTracking" && e.args?.data?.reason === "Animation").length;
      // stary FLIP na <tr>: ~60 unieważnień na animację (co klatkę); na kompozytorze kilka (start/koniec)
      const perAnim = nAnims ? animInv / nAnims : Infinity;
      check(`Chromium: animacja bez przeliczeń stylu co klatkę (${animInv} unieważnień / ${nAnims} animacji)`, perAnim < 12, "animacja wróciła na wątek główny");
    }
    check("filtr/przerwania: bez błędów strony", !errors.length, errors.join(" | "));
    await ctx.close();
  }

  // prefers-reduced-motion: bez animacji
  {
    const { ctx, page } = await openApp(browser, { reducedMotion: "reduce" });
    await sortClick(page);
    const any = await page.evaluate(() => ({ ghost: typeof _flipGhost !== "undefined" && !!_flipGhost, layers: document.querySelectorAll(".flip-ghosts").length }));
    check("prefers-reduced-motion — bez animacji", !any.ghost && !any.layers, JSON.stringify(any));
    await ctx.close();
  }

  await browser.close();
  if (fails.length) { console.log(`\n[${ENGINE}] ${fails.length} błędów`); process.exit(1); }
  console.log(`\n[${ENGINE}] OK`);
})().catch((e) => { console.error(e); process.exit(1); });
