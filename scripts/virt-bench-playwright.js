// virt-bench-playwright.js — POMIAR punktu wyjścia pod wirtualizację wierszy (etap F0).
//
// Nie jest testem (nie ma progów, nie siedzi w `npm test`). Ten sam skrypt odpalamy
// przed i po wirtualizacji (VIRT_QUERY="&virt=1") i porównujemy liczby.
//
// Dla kilku limitów wierszy (ROWS_LIST, domyślnie 200,600,2000,5000) mierzy na
// Chromium ze spowolnionym CPU (CPU, domyślnie ×4 ≈ średni telefon):
//   - wczytanie: klik „przykład” → pierwsza klatka z tabelą,
//   - sort / filtr / przebudowa: start operacji → po następnej klatce (JS + styl +
//     layout + malowanie), mediana z 5 powtórzeń, ścieżką użytkownika (z FLIP),
//   - przewijanie: 90 klatek po 90 px; średni i p95 czas klatki, klatki > 33 ms,
//   - liczba węzłów DOM po wczytaniu.
//
// Użycie:  node scripts/virt-bench-playwright.js
//          CPU=6 ROWS_LIST=200,2000 node scripts/virt-bench-playwright.js
//          OUT=docs/notes/virt-baseline.json node scripts/virt-bench-playwright.js

const { chromium } = require("playwright");
const fs = require("fs");

const APP_URL = process.env.APP_URL || "http://127.0.0.1:4175/";
const CPU = parseFloat(process.env.CPU || "4");
const ROWS_LIST = (process.env.ROWS_LIST || "200,600,2000,5000").split(",").map((s) => parseInt(s, 10));
const SAMPLE = parseInt(process.env.SAMPLE || "5000", 10);
const VIRT_QUERY = process.env.VIRT_QUERY || "";
const REPEAT = 5;

async function measureOne(browser, rows) {
  const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1100, height: 800 } });
  await context.addInitScript((rows) => {
    localStorage.setItem("introPlayed", "true");
    localStorage.setItem("excel-workbench-max-rows", String(rows));
  }, rows);
  const page = await context.newPage();
  const cdp = await context.newCDPSession(page);
  await page.goto(`${APP_URL}?sample=${SAMPLE}${VIRT_QUERY}`, { waitUntil: "load" });
  await page.evaluate(() => document.getElementById("heroSplash")?.remove());
  await page.evaluate(() => { try { if (typeof setSidebarOpen === "function") setSidebarOpen(false); } catch (_) {} });
  await cdp.send("Emulation.setCPUThrottlingRate", { rate: CPU });

  const load = await page.evaluate(() => new Promise((resolve) => {
    const t0 = performance.now();
    document.getElementById("loadSampleBtn")?.click();
    const poll = () => {
      if (document.querySelector("#dataTable tbody tr[data-row-key]")) {
        requestAnimationFrame(() => { const ch = new MessageChannel(); ch.port1.onmessage = () => resolve(performance.now() - t0); ch.port2.postMessage(0); });
      } else setTimeout(poll, 5);
    };
    poll();
  }));
  await page.waitForTimeout(1500); // prewarm analiz itp. niech się wyciszy
  await page.evaluate(() => { try { if (typeof setSidebarOpen === "function") setSidebarOpen(false); } catch (_) {} try { heroUserCollapsed = true; } catch (_) {} });

  const res = await page.evaluate(async (REPEAT) => {
    const afterFrame = (t0) => new Promise((resolve) => requestAnimationFrame(() => {
      const ch = new MessageChannel(); ch.port1.onmessage = () => resolve(performance.now() - t0); ch.port2.postMessage(0);
    }));
    const idle = () => new Promise((r) => setTimeout(r, 250));
    const median = (a) => { const s = [...a].sort((x, y) => x - y); return Math.round(s[Math.floor(s.length / 2)]); };
    const run = async (fn) => {
      const out = [];
      for (let i = 0; i < REPEAT; i++) { await idle(); const t0 = performance.now(); fn(i); out.push(await afterFrame(t0)); }
      return median(out);
    };
    const col = currentHeaders[1];
    const sort = await run((i) => { setPrimarySort(col, i % 2 ? "desc" : "asc"); sortRows(); renderActiveTable(); });
    const filter = await run((i) => { searchQueryEl.value = i % 2 ? "" : "a"; filtersCommitted = true; applyFilters(); sortRows(); renderActiveTable(); });
    searchQueryEl.value = ""; applyFilters(); sortRows(); renderActiveTable();
    await idle();
    const rerender = await run(() => { renderActiveTable(); });

    // przewijanie: stały krok na klatkę, mierzymy odstępy rAF
    tableWrapEl.scrollTop = 0;
    await idle();
    const frames = [];
    await new Promise((resolve) => {
      let last = performance.now(); let n = 0;
      const step = (t) => {
        frames.push(t - last); last = t;
        if (++n >= 90) return resolve();
        tableWrapEl.scrollTop += 90;
        requestAnimationFrame(step);
      };
      requestAnimationFrame((t) => { last = t; tableWrapEl.scrollTop += 90; requestAnimationFrame(step); });
    });
    const f = frames.slice(1).sort((a, b) => a - b);
    const avg = f.reduce((s, v) => s + v, 0) / f.length;
    return {
      shown: document.querySelectorAll("#dataTable tbody tr[data-row-key]").length,
      cols: currentHeaders.length,
      domNodes: document.getElementsByTagName("*").length,
      sort, filter, rerender,
      scrollAvg: Math.round(avg * 10) / 10,
      scrollP95: Math.round(f[Math.floor(f.length * 0.95)] * 10) / 10,
      scrollJank: f.filter((v) => v > 33).length,
    };
  }, REPEAT);
  await context.close();
  return { rows, load: Math.round(load), ...res };
}

(async () => {
  const browser = await chromium.launch({ headless: true });
  const out = [];
  for (const rows of ROWS_LIST) {
    const r = await measureOne(browser, rows);
    out.push(r);
    console.log(`limit ${String(rows).padStart(5)} | w DOM ${String(r.shown).padStart(5)} × ${r.cols} kol | węzły ${String(r.domNodes).padStart(6)} | wczytanie ${r.load} ms | sort ${r.sort} | filtr ${r.filter} | przebudowa ${r.rerender} ms | scroll śr ${r.scrollAvg} p95 ${r.scrollP95} ms, >33ms: ${r.scrollJank}/89`);
  }
  await browser.close();
  if (process.env.OUT) {
    fs.writeFileSync(process.env.OUT, JSON.stringify({ date: new Date().toISOString(), cpu: CPU, sample: SAMPLE, virt: VIRT_QUERY || "off", results: out }, null, 2));
    console.log(`zapisano ${process.env.OUT}`);
  }
})().catch((e) => { console.error(e); process.exit(1); });
