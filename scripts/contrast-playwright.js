// contrast-playwright.js — strażnik KONTRASTU tekstu (WCAG AA) w jasnym i ciemnym motywie.
//
// Dla każdego widocznego tekstu liczymy kontrast koloru względem RZECZYWISTEGO tła
// (półprzezroczyste warstwy składane w dół aż do nieprzezroczystej, z opacity przodków).
// Próg: 4,5:1, a dla dużego tekstu (≥ 24 px albo ≥ 18,66 px pogrubiony) 3:1.
// Tekst na kartce raportu mierzymy w rozmiarze EKRANOWYM (kartka jest skalowana).
//
// Pomijamy tylko to, czego WCAG nie wymaga: elementy wyłączone (disabled), niewidoczne
// (opacity ≈ 0, visibility, poza ekranem jak skip-link przed fokusem).
//
// Miejsca: raport (3 style × panel „Zawartość” z kątami × pasek zakresu), okno Eksport,
// główny ekran z wczytanym arkuszem.
//
// Po co: w ciemnym motywie komórki tabel raportu dziedziczyły jasny kolor apki
// (1,2:1 na białej kartce), a drobne szare opisy miały ~3,4:1.

const { chromium } = require("playwright");
const APP_URL = process.env.APP_URL || "http://127.0.0.1:4175/";
const SLEEP_SCALE = Math.min(1, Math.max(0.1, Number(process.env.SLEEP_SCALE || 1)));
const sleep = (ms) => new Promise((r) => setTimeout(r, Math.round(ms * SLEEP_SCALE)));

const AUDIT = (root) => {
  const parse = (c) => {
    const cs = String(c).match(/color\(srgb ([^)]+)\)/);
    if (cs) {
      const p = cs[1].split(/[ /]+/).filter(Boolean).map(Number);
      return { r: p[0] * 255, g: p[1] * 255, b: p[2] * 255, a: p.length > 3 ? p[3] : 1 };
    }
    const m = String(c).match(/rgba?\(([^)]+)\)/);
    if (!m) return null;
    const p = m[1].split(/[ ,/]+/).filter(Boolean).map(Number);
    return { r: p[0], g: p[1], b: p[2], a: p.length > 3 ? p[3] : 1 };
  };
  const lum = ({ r, g, b }) => {
    const f = (v) => { v /= 255; return v <= 0.03928 ? v / 12.92 : Math.pow((v + 0.055) / 1.055, 2.4); };
    return 0.2126 * f(r) + 0.7152 * f(g) + 0.0722 * f(b);
  };
  const blend = (top, bottom) => ({ r: top.r * top.a + bottom.r * (1 - top.a), g: top.g * top.a + bottom.g * (1 - top.a), b: top.b * top.a + bottom.b * (1 - top.a), a: 1 });
  const bgOf = (el) => {
    const stack = [];
    for (let e = el; e; e = e.parentElement) {
      const c = parse(getComputedStyle(e).backgroundColor);
      if (c && c.a > 0) { stack.push(c); if (c.a >= 1) break; }
    }
    let base = { r: 255, g: 255, b: 255, a: 1 };
    if (stack.length && stack[stack.length - 1].a >= 1) base = stack.pop();
    for (let i = stack.length - 1; i >= 0; i--) base = blend(stack[i], base);
    return base;
  };
  const out = [];
  document.querySelectorAll(`${root} *`).forEach((el) => {
    if (!Array.from(el.childNodes).some((n) => n.nodeType === 3 && n.textContent.trim())) return;
    if (el.closest(".hidden, [hidden], [inert], :disabled, [aria-disabled='true'], option, script, style")) return;
    const cs = getComputedStyle(el);
    if (cs.display === "none" || cs.visibility !== "visible" || !el.getClientRects().length) return;
    const r = el.getBoundingClientRect();
    if (r.right <= 0 || r.bottom <= 0 || r.left >= innerWidth || r.top >= innerHeight + 4000 || r.width < 2) return;
    let op = 1;
    for (let e = el; e; e = e.parentElement) op *= Number(getComputedStyle(e).opacity);
    if (op < 0.2) return; // w trakcie animacji / schowane przez opacity
    const fg0 = parse(cs.color);
    if (!fg0) return;
    const bg = bgOf(el);
    const fg = blend({ ...fg0, a: fg0.a * op }, bg);
    const L1 = lum(fg);
    const L2 = lum(bg);
    const ratio = (Math.max(L1, L2) + 0.05) / (Math.min(L1, L2) + 0.05);
    const scale = el.closest("#rpPage") ? (r.height / (el.offsetHeight || r.height || 1)) || 1 : 1;
    const px = parseFloat(cs.fontSize) * scale;
    const large = px >= 24 || (px >= 18.66 && Number(cs.fontWeight) >= 700);
    const need = large ? 3 : 4.5;
    if (ratio < need - 0.02) out.push(`${Math.round(ratio * 100) / 100}:1 (≥${need}, ${Math.round(px * 10) / 10}px) .${String(el.className || el.tagName).trim().split(/\s+/).join(".")} „${el.textContent.trim().slice(0, 30)}”`);
  });
  return Array.from(new Set(out));
};

async function loadSheet(page) {
  await page.evaluate(() => {
    const now = new Date();
    const iso = (d) => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
    const ws = {};
    ["Klient", "Status", "Miasto", "Kwota", "Data przyjęcia", "Termin"].forEach((h, i) => { ws[String.fromCharCode(65 + i) + "1"] = { t: "s", v: h }; });
    const st = ["W toku", "Zakończone", "Nowe", "Anulowane"];
    const mi = ["Kraków", "Warszawa", "Gdańsk"];
    for (let i = 0; i < 40; i++) {
      const r = i + 2;
      ws["A" + r] = { t: "s", v: `Klient ${i + 1}` };
      ws["B" + r] = { t: "s", v: st[i % 4] };
      ws["C" + r] = { t: "s", v: mi[i % 3] };
      ws["D" + r] = { t: "n", v: 100 + ((i * 37) % 400), w: String(100 + ((i * 37) % 400)) };
      ws["E" + r] = { t: "s", v: iso(new Date(now.getFullYear(), now.getMonth(), now.getDate() - i * 3)) };
      ws["F" + r] = { t: "s", v: iso(new Date(now.getFullYear(), now.getMonth(), now.getDate() + 20 - i * 2)) };
    }
    ws["!ref"] = "A1:F41";
    workbook = { SheetNames: ["Zlecenia"], Sheets: { Zlecenia: ws }, Props: {} };
    currentFileName = "kontrast.xlsx";
    sheetSelect.replaceChildren();
    const opt = document.createElement("option"); opt.value = "Zlecenia"; opt.textContent = "Zlecenia";
    sheetSelect.appendChild(opt); sheetSelect.value = "Zlecenia";
    document.getElementById("headerRow").value = "1";
    document.getElementById("autoHeaderRow").checked = false;
    document.getElementById("loadBtn").click();
  });
  await page.waitForFunction(() => typeof baseRows !== "undefined" && baseRows.length === 40 && currentDisplayModel, null, { timeout: 15000 });
  await sleep(800);
}

async function run() {
  const browser = await chromium.launch({ headless: true });
  const results = [];
  const check = (name, list) => {
    results.push({ name, ok: !list.length });
    console.log(`${list.length ? "❌" : "✅"} ${name}${list.length ? `\n   ${list.slice(0, 12).join("\n   ")}` : ""}`);
  };
  const errors = [];

  for (const scheme of ["light", "dark"]) {
    const context = await browser.newContext({ serviceWorkers: "block", colorScheme: scheme, viewport: { width: 1280, height: 900 } });
    await context.addInitScript(() => {
      localStorage.setItem("introPlayed", "true");
      localStorage.setItem("swb-report-prefs", JSON.stringify({ preset: "detailed", angle: "state" }));
    });
    const page = await context.newPage();
    page.on("pageerror", (e) => errors.push(`${scheme}: ${e.message}`));
    await page.goto(APP_URL, { waitUntil: "load" });
    await page.evaluate(() => document.getElementById("heroSplash")?.remove());
    await page.evaluate(() => ensureXlsxLibs(false));
    await page.waitForFunction(() => typeof buildRows === "function" && window.__report);
    await loadSheet(page);
    const theme = await page.evaluate(() => document.documentElement.getAttribute("data-theme") || "light");
    const label = scheme === "dark" ? "ciemny" : "jasny";
    if (scheme === "dark") check("motyw ciemny faktycznie włączony", theme === "dark" ? [] : [`data-theme=${theme}`]);

    await page.evaluate(() => document.querySelectorAll(".toast").forEach((el) => el.remove()));
    check(`${label}: główny ekran`, await page.evaluate(AUDIT, "body > .app"));

    await page.evaluate(() => openExportModal());
    await sleep(700);
    check(`${label}: okno Eksport`, await page.evaluate(AUDIT, "#exportModal"));
    await page.evaluate(() => document.querySelector("#exportModal [data-close], #exportModal .modal-close, #exportCancelBtn")?.click());
    await page.evaluate(() => document.getElementById("exportModal").classList.add("hidden"));

    await page.evaluate(() => {
      window.__report.open();
      window.__report.setScope({ kind: "query", q: 'Status:="W toku"' });
      document.getElementById("rpContentBtn").click();
    });
    await sleep(300);
    for (const style of ["modern", "classic", "ink"]) {
      await page.selectOption("#rpStyle", style);
      await sleep(150);
      check(`${label}: raport „${style}” (pasek, zakres, kąty, kartka)`, await page.evaluate(AUDIT, "#reportOverlay"));
    }
    await page.evaluate(() => window.__report.setAngle("compare"));
    await sleep(150);
    check(`${label}: raport — porównanie grup`, await page.evaluate(AUDIT, "#reportOverlay"));
    await page.evaluate(() => { window.__report.setAngle("overview"); window.__report.setScope({ kind: "view" }); window.__report.close(); });
    await context.close();
  }

  check("brak błędów w konsoli", errors);
  await browser.close();
  const failed = results.filter((r) => !r.ok);
  if (failed.length) {
    console.log(`\n❌ contrast: ${failed.length} z ${results.length} nie przeszło`);
    process.exit(1);
  }
  console.log(`\n✅ contrast: ${results.length}/${results.length}`);
}

run().catch((e) => { console.error(e); process.exit(1); });
