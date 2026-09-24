// edit-unify-playwright.js — Narzędzia edycji: „Ujednolić warianty", „Znajdź i zamień"
// (bez wielkości liter / tylko cała komórka) i „Podgląd".
//
// Co musi być prawdą:
//   1. „Znajdź warianty" (szeroko) grupuje pisownie różniące się wielkością liter,
//      spacjami, . - / _ i polskimi znakami; propozycja = najczęstsza pisownia,
//   2. tryb „wąsko" NIE łączy „C-03" z „c03" (różnica to myślnik, nie wielkość liter),
//   3. odznaczona grupa zostaje nietknięta; inne kolumny też,
//   4. „Podgląd" niczego nie zmienia i pokazuje dokładnie tyle zmian, ile potem „Zastosuj",
//   5. całe zastosowanie = JEDEN krok Cofnij,
//   6. „Tylko cała komórka" + „bez wielkości liter": „a1" → zmienia „a1" i „A1", nie „a10",
//   7. ZAPIS (ZIP-patch) → ponowny odczyt: nowe wartości są w pliku, daty zostały datami,
//      nieedytowane komórki bez zmian.
//
// Uruchom z serwerem na APP_URL (domyślnie http://127.0.0.1:4175/), np. `npm run serve:test`.

const { chromium } = require("playwright");
const fs = require("fs");
const os = require("os");
const path = require("path");

const APP_URL = process.env.APP_URL || "http://127.0.0.1:4175/";
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

// Plik testowy budujemy sami (SheetJS z lib/), żeby znać każdy wariant.
function buildFixture() {
  const XLSX = require(path.join(__dirname, "..", "lib", "xlsx.full.min.js"));
  const d = (y, m, dd) => new Date(Date.UTC(y, m - 1, dd));
  const rows = [
    ["Nr", "Teren", "Kod", "Data"],
    [1, "Teren C-03", "a1", d(2024, 5, 12)],
    [2, "Teren C-03", "A1", d(2024, 5, 13)],
    [3, "teren c03", "a10", d(2024, 5, 14)],
    [4, "Teren C 03 ", "B-2", d(2024, 5, 15)],
    [5, "Teren C-03", "b2", d(2024, 5, 16)],
    [6, "Łąka", "C3", d(2024, 5, 17)],
    [7, "Łąka", "c3", d(2024, 5, 18)],
    [8, "Laka", "D4", d(2024, 5, 19)],
    [9, "Inny", "E5", d(2024, 5, 20)],
    [10, "Inny", "F6", d(2024, 5, 21)],
  ];
  const ws = XLSX.utils.aoa_to_sheet(rows, { cellDates: true, dateNF: "dd-mm-yy" });
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Dane");
  const file = path.join(fs.mkdtempSync(path.join(os.tmpdir(), "unify-")), "unify-test.xlsx");
  fs.writeFileSync(file, XLSX.write(wb, { type: "buffer", bookType: "xlsx", cellDates: false }));
  return file;
}

async function run() {
  const FILE = buildFixture();
  const browser = await chromium.launch({ headless: true });
  const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1280, height: 900 } });
  await context.addInitScript(() => localStorage.setItem("introPlayed", "true"));
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push("pageerror: " + e.message));
  page.on("console", (m) => { if (m.type() === "error") errors.push("console: " + m.text()); });

  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(() => { document.getElementById("heroSplash")?.remove(); });
  await page.evaluate(() => { try { ensureXlsxLibs && ensureXlsxLibs(false); } catch {} });
  await page.setInputFiles("#fileInput", FILE);
  await page.waitForFunction(() => document.getElementById("sheetSelect")?.options?.length > 0, null, { timeout: 15000 });
  await sleep(300);
  await page.click("#loadBtn");
  await page.waitForFunction(() => Array.isArray(baseRows) && baseRows.length === 10, null, { timeout: 15000 });
  await page.evaluate(() => { document.getElementById("panel-edit-tools").open = true; });
  await sleep(200);

  const failures = [];
  const ok = (name, cond, got) => { if (!cond) failures.push(`${name} — dostałem: ${JSON.stringify(got)}`); };

  const col = (name) => page.evaluate((name) => currentHeaders.indexOf(name), name);
  const colValues = (name) => page.evaluate((name) => {
    const i = currentHeaders.indexOf(name);
    return baseRows.map((r) => r.values[i]);
  }, name);
  const setSel = (id, value) => page.selectOption(`#${id}`, String(value));
  const groups = () => page.evaluate(() => [...document.querySelectorAll("#editUnifyList .unify-group")].map((g) => ({
    variants: [...g.querySelectorAll(".unify-opt")].map((o) => ({
      text: o.querySelector(".unify-val").textContent,
      checked: o.querySelector("input").checked,
    })),
  })));

  const tCol = await col("Teren");
  await setSel("editScope", "column");
  await setSel("editColumnSelect", tCol);
  await setSel("editOp", "unify");

  // 2. wąsko — nic do scalenia
  await setSel("editUnifyMode", "case");
  await page.click("#editUnifyScanBtn");
  const narrow = await groups();
  const noneNote = await page.evaluate(() => document.querySelector("#editUnifyList .field-note")?.textContent || "");
  ok("2. wąsko: brak grup", narrow.length === 0 && noneNote.length > 0, { narrow, noneNote });

  // 1. szeroko — dwie grupy, propozycja = najczęstsza
  await setSel("editUnifyMode", "loose");
  const clearedOnModeChange = await page.evaluate(() => document.getElementById("editUnifyList").children.length === 0);
  ok("zmiana trybu czyści listę", clearedOnModeChange, clearedOnModeChange);
  await page.click("#editUnifyScanBtn");
  const g = await groups();
  ok("1. dwie grupy", g.length === 2, g);
  ok("1. grupa C-03: 3 warianty, propozycja „Teren C-03”",
    g[0] && g[0].variants.length === 3 && g[0].variants[0].checked && g[0].variants[0].text === "„Teren C-03”", g[0]);
  ok("1. spacja na końcu jest widoczna (␣)", g[0] && g[0].variants.some((v) => v.text.includes("␣")), g[0]);
  ok("1. grupa Łąka/Laka, propozycja „Łąka”",
    g[1] && g[1].variants.length === 2 && g[1].variants[0].text === "„Łąka”", g[1]);

  // 3. odznacz grupę Łąka
  await page.evaluate(() => { document.querySelectorAll("#editUnifyList .unify-group .unify-on")[1].click(); });

  // 4. podgląd — nic nie zmienia, pokazuje 2 zmiany
  const before = await colValues("Teren");
  await page.click("#previewEditToolBtn");
  const prev = await page.evaluate(() => ({
    visible: !document.getElementById("editPreview").classList.contains("hidden"),
    sum: document.querySelector("#editPreview .edit-preview-sum")?.textContent || "",
    rows: document.querySelectorAll("#editPreview .edit-preview-row").length,
    pending: Object.keys(pendingEdits[currentSheetName] || {}).length,
  }));
  const afterPreview = await colValues("Teren");
  ok("4. podgląd widoczny z 2 zmianami", prev.visible && /2/.test(prev.sum) && prev.rows === 2, prev);
  ok("4. podgląd nic nie zmienił", JSON.stringify(before) === JSON.stringify(afterPreview) && prev.pending === 0, { afterPreview, prev });

  await page.click("#applyEditToolBtn");
  await sleep(200);
  const teren = await colValues("Teren");
  ok("3. C-03 ujednolicone, Łąka/Laka nietknięte",
    JSON.stringify(teren) === JSON.stringify(["Teren C-03", "Teren C-03", "Teren C-03", "Teren C-03", "Teren C-03", "Łąka", "Łąka", "Laka", "Inny", "Inny"]), teren);
  const kod = await colValues("Kod");
  ok("3. inna kolumna nietknięta", kod.join("|") === "a1|A1|a10|B-2|b2|C3|c3|D4|E5|F6", kod);
  const pend = await page.evaluate(() => Object.keys(pendingEdits[currentSheetName] || {}).sort());
  ok("4. zastosowano tyle, ile pokazał podgląd", pend.length === 2, pend);
  const listCleared = await page.evaluate(() => document.getElementById("editUnifyList").children.length === 0);
  ok("lista wyczyszczona po zastosowaniu", listCleared, listCleared);

  // 5. jeden krok Cofnij
  await page.evaluate(() => undoLast());
  await sleep(150);
  const undone = await colValues("Teren");
  ok("5. Cofnij przywraca obie komórki naraz", JSON.stringify(undone) === JSON.stringify(before), undone);
  await page.evaluate(() => redoLast());
  await sleep(150);

  // 6. znajdź i zamień — cała komórka + bez wielkości liter
  await setSel("editOp", "replace");
  await setSel("editColumnSelect", await col("Kod"));
  await page.fill("#editFind", "a1");
  await page.fill("#editReplace", "A-1");
  await page.check("#editIgnoreCase");
  await page.check("#editWholeCell");
  await page.click("#applyEditToolBtn");
  await sleep(150);
  const kod2 = await colValues("Kod");
  ok("6. „a1” i „A1” → „A-1”, „a10” zostaje", kod2.slice(0, 3).join("|") === "A-1|A-1|a10", kod2);
  // bez „całej komórki": b2/B-2 — zamiana podciągu bez wielkości liter
  await page.uncheck("#editWholeCell");
  await page.fill("#editFind", "c");
  await page.fill("#editReplace", "K$1"); // „$" ma zostać dosłownie
  await page.click("#applyEditToolBtn");
  await sleep(150);
  const kod3 = await colValues("Kod");
  ok("6. podciąg bez wielkości liter + „$” dosłownie", kod3[5] === "K$13" && kod3[6] === "K$13", kod3);

  // 7. zapis → ponowny odczyt
  const saved = await page.evaluate(async () => {
    const bytes = await buildOutputBytes("xlsx");
    const wb = XLSX.read(bytes, { cellDates: true, cellNF: true });
    const sh = wb.Sheets[wb.SheetNames[0]];
    const out = {};
    ["B2", "B4", "B5", "B7", "B9", "C2", "C3", "C4", "D2", "D11"].forEach((r) => {
      const c = sh[r];
      out[r] = c ? { t: c.t, v: c.v instanceof Date ? c.v.toISOString().slice(0, 10) : c.v } : null;
    });
    return out;
  });
  ok("7. zapisane wartości", saved.B4.v === "Teren C-03" && saved.B5.v === "Teren C-03" && saved.B7.v === "Łąka"
    && saved.B9.v === "Laka" && saved.C2.v === "A-1" && saved.C3.v === "A-1" && saved.C4.v === "a10", saved);
  ok("7. daty zostały datami", saved.D2.t === "d" && saved.D2.v === "2024-05-12" && saved.D11.t === "d", saved);

  ok("brak błędów konsoli", errors.length === 0, errors);
  await browser.close();

  if (failures.length) {
    console.error("❌ edit-unify: " + failures.length + " błędów\n  - " + failures.join("\n  - "));
    process.exit(1);
  }
  console.log("✅ edit-unify: wszystko OK");
  process.exit(0);
}

run().catch((e) => { console.error(e); process.exit(1); });
