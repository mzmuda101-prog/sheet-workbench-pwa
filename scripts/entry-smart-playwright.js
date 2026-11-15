// entry-smart-playwright.js — Wpisywanie, tryb smart (cykle od/do) na prawdziwym pliku RODO (arkusz „2025-2026”, klucz Nr.).
//
// Wiersze z pomiaru (Nr → wiersz Excela): 001 → 3 (cykl 1 otwarty), 036 → 38 (cykl 1 oddany, 2 otwarty),
// 004 → 6 (oddana), 002 → 4 (pusty), 293 → 295 (dwa otwarte), 043 → 45 (osoba bez od).
// Co musi być prawdą:
//   A. silnik w aplikacji: 8 cykli × 4, role osoba/od/do/formuła; decyzje per wiersz,
//   B. kolejka z polami per wiersz: różne kolumny, jeden zapis, jedno Cofnij,
//   C. karta 001: panel smart zamiast 33 pól, domyślnie Zdaj z datą dziś → podgląd „do” + cykl → zapis D3,
//   D. Wydaj na 001 (otwarty) → okno z wyborem; „Zamknij poprzedni i wydaj” → D3 + F3 + G3,
//   E. Zdaj na 004 (oddana) → komunikat, nic się nie zapisuje,
//   F. kolejka „001, 036, 004, 293” + Zdaj: 001 → do, 036 → do2, 004 nie wydana, 293 do sprawdzenia (wybór),
//      „Wszystkim naraz” = D3 + H38, jedno Cofnij,
//   G. Ustawienia → Tryb smart: nazwa akcji per arkusz, zostaje po przeładowaniu; wyłączenie = zwykła karta,
//   H. telefon (MOBILE=1: WebKit, dotyk, 382 px): bez poziomego przewijania karty i kolejki,
//   I. zapis pliku: openpyxl widzi daty w D3/H38 i osobę w F4, OOXML bez nowych błędów; EXCEL_DIR → kopia dla Excela.
const pw = require("playwright");
const { autoAnswerAppDialogs } = require("./app-dialog-auto");
const path = require("path");
const fs = require("fs");
const os = require("os");
const { execFileSync } = require("child_process");
let ooxml = null;
try { ooxml = require("./ooxml-validate"); } catch { ooxml = null; }

const APP_URL = process.env.APP_URL || "http://127.0.0.1:4175/";
const MOBILE = process.env.MOBILE === "1";
const ENGINE = process.env.ENGINE || (MOBILE ? "webkit" : "chromium");
const EXCEL_DIR = process.env.EXCEL_DIR || "";
const RODO = path.join(__dirname, "RODO Obieg terenów 2026 V5.xlsx");
const PY = fs.existsSync(path.join(__dirname, "..", ".venv/bin/python")) ? path.join(__dirname, "..", ".venv/bin/python") : "python3";
const SLEEP_SCALE = Math.min(1, Math.max(0.1, Number(process.env.SLEEP_SCALE || 1)));
const sleep = (ms) => new Promise((r) => setTimeout(r, Math.round(ms * SLEEP_SCALE)));

const failures = [];
async function run() {
  if (!fs.existsSync(RODO)) { console.log("⏭  entry-smart: brak pliku RODO — pominięte"); return; }
  const tmp = fs.mkdtempSync(path.join(os.tmpdir(), "swb-entry-smart-"));
  const FILE = path.join(tmp, "RODO Obieg terenów 2026 V5.xlsx");
  fs.copyFileSync(RODO, FILE);
  const browser = await pw[ENGINE].launch({ headless: true });
  const context = await browser.newContext(MOBILE
    ? { serviceWorkers: "block", viewport: { width: 382, height: 800 }, hasTouch: true, isMobile: true, deviceScaleFactor: 3 }
    : { serviceWorkers: "block", viewport: { width: 1280, height: 860 } });
  await context.addInitScript(() => {
    localStorage.setItem("introPlayed", "true");
    if (!sessionStorage.getItem("enTestInit")) { localStorage.removeItem("excel-workbench-entry"); sessionStorage.setItem("enTestInit", "1"); }
  });
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (e) => { errors.push(e.message); if (process.env.DEBUG) console.log("PAGEERROR", e.stack); });
  await autoAnswerAppDialogs(page);
  const ok = (name, cond, got) => { if (!cond) failures.push(`${name} — dostałem: ${JSON.stringify(got)}`); };

  const load = async () => {
    await page.goto(APP_URL, { waitUntil: "load" });
    await page.evaluate(() => { document.getElementById("heroSplash")?.remove(); window.__appDialogMode = "off"; return ensureXlsxLibs(false); });
    await page.setInputFiles("#fileInput", FILE);
    await page.waitForFunction(() => typeof workbook !== "undefined" && workbook && sheetSelect.options.length > 0, null, { timeout: 60000 });
    await page.evaluate(() => { sheetSelect.value = "2025-2026"; loadBtn.click(); });
    await page.waitForFunction(() => currentSheetName === "2025-2026" && document.querySelector("#dataTable tbody tr[data-row-key]"), null, { timeout: 60000 });
    await page.evaluate(() => { sheetEditUi.setMode("edit"); });
    await sleep(800);
  };
  await load();

  const today = await page.evaluate(() => SwbEntryCore.todayText());
  const todayYmd = await page.evaluate(() => { const d = new Date(); return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`; });
  const cell = (ref) => page.evaluate((ref) => {
    const x = workbook.Sheets[currentSheetName][ref];
    if (!x || x.v == null || x.v === "") return null;
    const d = x.v;
    return d instanceof Date ? `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}` : d;
  }, ref);
  const st = () => page.evaluate(() => window.__entry.state());
  const undoCard = () => page.evaluate(() => window.__entry.undoCard());
  const clickDlg = async (value) => {
    await page.waitForSelector(`.app-dialog .app-dialog-actions button[data-value="${value}"]`, { timeout: 5000 });
    await page.click(`.app-dialog .app-dialog-actions button[data-value="${value}"]`);
    await sleep(250);
  };
  const noHScroll = (label) => page.evaluate(() => {
    const ov = document.getElementById("entryOverlay");
    const wide = [...ov.querySelectorAll("*")].filter((el) => { const r = el.getBoundingClientRect(); return r.width && (r.right > innerWidth + 1 || r.left < -1); }).slice(0, 4).map((el) => `${el.tagName}.${el.className}#${el.id}`);
    return { doc: document.documentElement.scrollWidth, vw: innerWidth, ov: ov.scrollWidth - ov.clientWidth, wide };
  }).then((m) => ok(`H. ${label}: bez poziomego przewijania`, m.doc <= m.vw && m.ov <= 1 && !m.wide.length, m));

  // ── A. silnik w aplikacji ──
  const a = await page.evaluate(() => {
    const src = window.__entry.cycleSource();
    const d = "16.11.2026";
    const sp = (r, act, o) => window.__entry.smartPlan(r, act, { date: d, ...o });
    return {
      src: src && { count: src.count, span: src.span, first: src.first, roles: src.roles },
      c001: sp(2, "close"), c036: sp(37, "close"), c004: sp(5, "close"), c293: sp(294, "close"), c043: sp(44, "issue"),
      i002: sp(3, "issue", { data: { 0: "Anna Testowa" } }), i001: sp(2, "issue", { data: { 0: "Anna Testowa" } }),
    };
  });
  ok("A. 8 cykli × 4 od kolumny 1", a.src && a.src.count === 8 && a.src.span === 4 && a.src.first === 1, a.src);
  ok("A. role osoba/od/do/formuła", a.src && a.src.roles.join() === "data,start,end,skip", a.src);
  ok("A. 001 Zdaj → do (kol. 3)", a.c001.status === "ok" && a.c001.writes[3] === "16.11.2026", a.c001);
  ok("A. 036 Zdaj → do2 (kol. 7)", a.c036.status === "ok" && a.c036.cycle === 2 && a.c036.writes[7], a.c036);
  ok("A. 004 Zdaj → nie wydana", a.c004.status === "not-out", a.c004);
  ok("A. 293 → do sprawdzenia", a.c293.status === "check", a.c293);
  ok("A. 043 → do sprawdzenia", a.c043.status === "check", a.c043);
  ok("A. 002 Wydaj → cykl 1 osoba + od", a.i002.status === "ok" && a.i002.writes[1] === "Anna Testowa" && a.i002.writes[2], a.i002);
  ok("A. 001 Wydaj → pytanie (poprzedni otwarty)", a.i001.status === "open-prev" && a.i001.openCycle === 1 && a.i001.cycle === 2, a.i001);

  // ── B. kolejka z polami per wiersz ──
  const b = await page.evaluate(async () => {
    const r = await window.__entry.commitMany({ items: [{ rowIndex0: 2, fields: { 3: "16.11.2026" } }, { rowIndex0: 37, fields: { 7: "16.11.2026" } }] });
    return r;
  });
  ok("B. commitMany z items ok", b.ok && b.rows.length === 2 && b.undoSteps === 1, b);
  ok("B. D3 i H38 = 16.11.2026", (await cell("D3")) === "2026-11-16" && (await cell("H38")) === "2026-11-16", [await cell("D3"), await cell("H38")]);
  ok("B. nic poza tym (F3, D38 bez zmian)", (await cell("F3")) === null && (await cell("D38")) !== null, [await cell("F3"), await cell("D38")]);
  ok("B. jedno Cofnij zdejmuje oba", (await undoCard()) && (await cell("D3")) === null && (await cell("H38")) === null, [await cell("D3"), await cell("H38")]);

  // ── C. karta 001: Zdaj ──
  const openEntry = async () => { await page.evaluate(() => document.getElementById("entryBtn").click()); await sleep(300); };
  const pick = async (q) => {
    await page.fill("#enTargetInput", q); await sleep(200);
    await page.press("#enTargetInput", "Enter"); await sleep(300);
  };
  await openEntry();
  await pick("001");
  const c1 = await page.evaluate(() => ({
    smart: !!document.querySelector("#enSmart"), fields: document.querySelectorAll("#entryOverlay .en-field[data-en-col]").length,
    state: document.querySelector("#enSmartState")?.textContent || "",
    close: document.querySelector('[data-en-action="close"]')?.getAttribute("aria-pressed"),
    date: document.querySelector("#enSmartDate")?.value,
  }));
  ok("C. panel smart zamiast pól", c1.smart && c1.fields === 0, c1);
  ok("C. stan: cykl 1 otwarty + osoba", /1/.test(c1.state) && /Adam Kamiński/.test(c1.state), c1.state);
  ok("C. domyślnie Zdaj (otwarty cykl)", c1.close === "true", c1);
  ok("C. data domyślnie dziś", c1.date === today, c1.date);
  await page.click("#enReviewBtn"); await sleep(300);
  const c2 = await page.evaluate(() => ({ step: window.__entry.state().step, txt: document.querySelector("#enStage").textContent, cyc: document.querySelector("#enSmartCycle")?.textContent || "" }));
  ok("C. podgląd: krok 3, kolumna „do”, cykl 1", c2.step === 3 && /\bdo\b/.test(c2.txt) && /1/.test(c2.cyc), c2);
  const afterTxt = await page.evaluate(() => document.querySelector("#enStage .en-after")?.textContent || "");
  await page.click("#enCommitBtn"); await sleep(400);
  const shown = await page.evaluate(() => String(getDisplayValue(baseRows.find((r) => r.rowIndex0 === 2), 3) ?? ""));
  ok("C. podgląd daty = to, co pokazuje komórka (d mmm yy)", afterTxt && afterTxt === shown && !/-/.test(afterTxt), { afterTxt, shown });
  ok("C. D3 = dziś, F3 puste", (await cell("D3")) === todayYmd && (await cell("F3")) === null, [await cell("D3"), await cell("F3")]);
  ok("C. po zapisie krok 1", (await st()).step === 1, (await st()).step);
  ok("C. Cofnij kartki", (await undoCard()) && (await cell("D3")) === null, await cell("D3"));
  if (MOBILE) { await pick("001"); await noHScroll("karta smart"); await page.click("#enChangeTarget").catch(() => {}); await sleep(200); }

  // ── D. Wydaj przy otwartym poprzednim ──
  await page.fill("#enTargetInput", ""); await pick("001");
  await page.click('[data-en-action="issue"]'); await sleep(150);
  await page.fill('[data-en-smart-data="0"]', "Anna Testowa"); await sleep(100);
  await page.click("#enReviewBtn"); await sleep(300);
  const dlg = await page.evaluate(() => ({ msg: document.querySelector(".app-dialog .app-dialog-msg")?.textContent || "", btns: [...document.querySelectorAll(".app-dialog .app-dialog-actions button")].map((b) => b.dataset.value) }));
  ok("D. okno: stan poprzedniego cyklu", /Adam Kamiński/.test(dlg.msg) && /1/.test(dlg.msg), dlg);
  ok("D. okno: cztery wybory", ["close-issue", "issue-anyway", "note"].every((v) => dlg.btns.includes(v)) && dlg.btns.length === 4, dlg.btns);
  await clickDlg("close-issue");
  ok("D. po wyborze — podgląd", (await st()).step === 3, (await st()).step);
  await page.click("#enCommitBtn"); await sleep(400);
  ok("D. D3 = dziś, F3 = Anna Testowa, G3 = dziś", (await cell("D3")) === todayYmd && (await cell("F3")) === "Anna Testowa" && (await cell("G3")) === todayYmd,
    [await cell("D3"), await cell("F3"), await cell("G3")]);
  ok("D. Cofnij kartki zdejmuje wszystkie trzy", (await undoCard()) && (await cell("D3")) === null && (await cell("F3")) === null && (await cell("G3")) === null, null);
  // Tylko notatka: nic nie wydaje, notatka w komórce Nr.
  await pick("001");
  await page.click('[data-en-action="issue"]'); await sleep(150);
  await page.click("#enReviewBtn"); await sleep(300);
  await clickDlg("note");
  await page.waitForSelector(".app-dialog .app-dialog-field input", { timeout: 5000 });
  await page.fill(".app-dialog .app-dialog-field input", "Dzwonić: karta nie wróciła");
  await page.click('.app-dialog .app-dialog-actions button[data-kind="primary"]'); await sleep(300);
  await page.click("#enCommitBtn").catch(() => {}); await sleep(400);
  const n1 = await page.evaluate(() => swbNoteOf(currentSheetName, "A3")?.text || null);
  ok("D. Tylko notatka: notatka w A3, F3 puste", n1 === "Dzwonić: karta nie wróciła" && (await cell("F3")) === null, { n1, F3: await cell("F3") });
  await undoCard();

  // ── E. Zdaj na oddanej ──
  await pick("004");
  await page.click('[data-en-action="close"]'); await sleep(150);
  await page.click("#enReviewBtn"); await sleep(300);
  const e1 = await page.evaluate(() => ({ step: window.__entry.state().step, msg: document.querySelector("#enSmartMsg")?.textContent || "" }));
  ok("E. 004 Zdaj: zostaje na karcie z komunikatem", e1.step === 2 && e1.msg.length > 5, e1);
  ok("E. nic nie zapisane (I6 puste)", (await cell("I6")) === null && (await cell("H6")) === null, null);
  await page.click("#enChangeTarget"); await sleep(200);

  // ── F. kolejka ze Zdaj ──
  await page.fill("#enTargetInput", "001, 036, 004, 293"); await sleep(200);
  await page.press("#enTargetInput", "Enter"); await sleep(400);
  await page.click('[data-en-q-action="close"]'); await sleep(250);
  const f1 = await page.evaluate(() => ({
    s: [0, 1, 2, 3].map((i) => document.querySelector(`[data-en-q-smart="${i}"]`)?.textContent || ""),
    sel3: !!document.querySelector('[data-en-q-resolve="3"]'),
    all: document.querySelector("#enQAll")?.disabled, date: document.querySelector("#enQSmartDate")?.value, hint: document.querySelector("#enQHint")?.textContent || "",
  }));
  ok("F. 001 → do, 036 → do2", /\bdo\b/.test(f1.s[0]) && /do2/.test(f1.s[1]), f1.s);
  ok("F. 004 nie wydana, 293 do sprawdzenia z wyborem", f1.s[2].length > 3 && !/do\d?\b/.test(f1.s[2].replace(/\d{3}/, "")) && f1.sel3, f1);
  ok("F. bez decyzji przy 293 — „naraz” zablokowane; data dziś", f1.all === true && f1.date === today, f1);
  ok("F. podpowiedź mówi o „do sprawdzenia”, nie o „pasuje kilka wierszy”", /sprawdzenia/.test(f1.hint) && !/kilka wierszy/.test(f1.hint), f1.hint);
  if (MOBILE) await noHScroll("kolejka smart");
  // F2. pominięta pozycja (004, brak otwartego cyklu): wybór zamiast samego „pomijam”; „Popraw ręcznie” → wiersz → z powrotem
  const f004 = await page.evaluate(() => { const s = document.querySelector('[data-en-q-resolve="2"]'); return s ? { v: s.value, opts: [...s.options].map((o) => o.value) } : null; });
  ok("F2. 004: wybór z domyślnym „Pomiń”, notatka i poprawa ręczna", f004 && f004.v === "skip" && ["note", "fix", "skip"].every((x) => f004.opts.includes(x)) && !f004.opts.includes(""), f004);
  await page.selectOption('[data-en-q-resolve="2"]', "fix"); await sleep(300);
  const fx = await page.evaluate(() => {
    const inp = document.querySelector('#entryOverlay [data-en-col="2"] input');
    return { step: window.__entry.state().step, fields: document.querySelectorAll("#entryOverlay .en-field[data-en-col]").length, back: !!document.querySelector("#enQFixBack"),
      ph: inp ? inp.placeholder : null, phColor: inp ? getComputedStyle(inp, "::placeholder").color : null, ink: inp ? getComputedStyle(inp).color : null };
  });
  ok("F2. Popraw ręcznie: zwykła karta wiersza 6 z powrotem do kolejki", fx.step === 2 && fx.fields > 10 && fx.back, fx);
  ok("G2. obecna wartość w polu (placeholder) czarna, nie szara", fx.ph && fx.phColor === fx.ink, fx);
  await page.fill('#entryOverlay [data-en-col="6"] input', "10.11.2026"); await sleep(100);
  await page.click("#enReviewBtn"); await sleep(300); await page.click("#enCommitBtn"); await sleep(500);
  const f3 = await page.evaluate(() => ({ step: window.__entry.state().step, s2: document.querySelector('[data-en-q-smart="2"]')?.textContent || "" }));
  ok("F2. po poprawie wraca do kolejki, 004 gotowa: → do2", f3.step === "queue" && /do2/.test(f3.s2), f3);
  ok("F2. poprawka zapisana (G6)", (await cell("G6")) === "2026-11-10", await cell("G6"));
  await page.selectOption('[data-en-q-resolve="3"]', "skip"); await sleep(150);
  ok("F. po decyzji „naraz” dostępne", await page.evaluate(() => !document.querySelector("#enQAll").disabled), null);
  await page.click("#enQAll"); await sleep(400);
  const f2 = await page.evaluate(() => ({ step: window.__entry.state().step, rows: document.querySelectorAll(".en-q-row").length, txt: document.querySelector("#enStage").textContent }));
  ok("F. podgląd: 3 wiersze, do i do2", f2.step === "queueReview" && f2.rows === 3 && /do2/.test(f2.txt), f2);
  await page.click("#enCommitBtn"); await sleep(500);
  ok("F. po naraz — wybór Zakończ / uzupełnij resztę", await page.evaluate(() => window.__entry.state().step === "queueAfter" && !!document.querySelector("#enQAfterWalk")), null);
  await page.click("#enQAfterEnd"); await sleep(200);
  ok("F. D3, H38 i H6 (po poprawie) = dziś", (await cell("D3")) === todayYmd && (await cell("H38")) === todayYmd && (await cell("H6")) === todayYmd, [await cell("D3"), await cell("H38"), await cell("H6")]);
  ok("F. jedno Cofnij", (await undoCard()) && (await cell("D3")) === null && (await cell("H38")) === null && (await cell("H6")) === null, null);
  await page.evaluate(() => undoLast()); await sleep(200); // poprawka G6 to osobny krok
  ok("F2. Cofnij poprawki", (await cell("G6")) === null, await cell("G6"));

  // ── G. Ustawienia ──
  await page.click("#enSettingsBtn"); await sleep(200);
  ok("G. sekcja Tryb smart", await page.evaluate(() => !!document.querySelector("#enSmartOnOpt") && !!document.querySelector("#enSmartCloseName")), null);
  await page.fill("#enSmartCloseName", "Oddaj"); await page.dispatchEvent("#enSmartCloseName", "change"); await sleep(150);
  await page.click("#enSettingsBtn"); await sleep(200);
  await pick("001");
  ok("G. nowa nazwa akcji na karcie", /Oddaj/.test(await page.textContent('[data-en-action="close"]')), await page.textContent('[data-en-action="close"]'));
  await page.evaluate(() => window.__entry.close()); await sleep(200);
  await page.evaluate(() => { hasUnsavedChanges = false; setDirtyState(false); });
  await load();
  await openEntry();
  await pick("001");
  ok("G. nazwa zostaje po przeładowaniu", /Oddaj/.test((await page.textContent('[data-en-action="close"]').catch(() => "")) || ""), null);
  await page.click("#enSettingsBtn"); await sleep(200);
  await page.uncheck("#enSmartOnOpt"); await sleep(150);
  await page.click("#enSettingsBtn"); await sleep(200);
  await page.click("#enChangeTarget"); await sleep(150); await pick("001");
  ok("G. wyłączony tryb = zwykłe pola", await page.evaluate(() => !document.querySelector("#enSmart") && document.querySelectorAll("#entryOverlay .en-field[data-en-col]").length > 10), null);
  await page.click("#enSettingsBtn"); await sleep(200);
  await page.check("#enSmartOnOpt"); await sleep(150);
  await page.click("#enSettingsBtn"); await sleep(200);
  await page.click("#enChangeTarget"); await sleep(150);

  // ── I. plik do Excela: Zdaj 001 i 036, Wydaj 002 ──
  await page.fill("#enTargetInput", "001, 036"); await sleep(200); await page.press("#enTargetInput", "Enter"); await sleep(400);
  await page.click('[data-en-q-action="close"]'); await sleep(200);
  await page.click("#enQAll"); await sleep(300); await page.click("#enCommitBtn"); await sleep(500);
  await page.click("#enQAfterEnd"); await sleep(200);
  await pick("002");
  await page.click('[data-en-action="issue"]'); await sleep(100);
  await page.fill('[data-en-smart-data="0"]', "Anna Testowa"); await sleep(100);
  await page.click("#enReviewBtn"); await sleep(300); await page.click("#enCommitBtn"); await sleep(500);
  ok("I. B4 = Anna Testowa, C4 = dziś", (await cell("B4")) === "Anna Testowa" && (await cell("C4")) === todayYmd, [await cell("B4"), await cell("C4")]);
  const out = path.join(tmp, "smart.xlsx");
  fs.writeFileSync(out, Buffer.from(await page.evaluate(async () => Array.from(await buildOutputBytes("xlsx")))));
  if (EXCEL_DIR) { fs.mkdirSync(EXCEL_DIR, { recursive: true }); fs.copyFileSync(out, path.join(EXCEL_DIR, "entry-smart.xlsx")); }
  let o = null;
  try {
    o = JSON.parse(execFileSync(PY, ["-I", "-c", "import openpyxl,sys,json\nws=openpyxl.load_workbook(sys.argv[1])['2025-2026']\nf=lambda c:(ws[c].value.strftime('%Y-%m-%d') if hasattr(ws[c].value,'strftime') else ws[c].value)\nprint(json.dumps({k:f(k) for k in ['D3','H38','B4','C4','D4','F3']}|{'fmt':ws['D3'].number_format,'E3':str(ws['E3'].value)[:3],'ref':ws.tables['opracowanie_terenow'].ref}))", out], { encoding: "utf8" }));
  } catch (e) { o = { err: String(e.stderr || e.message).slice(0, 300) }; }
  ok("I. openpyxl: daty i osoba w swoich kolumnach", o && o.D3 === todayYmd && o.H38 === todayYmd && o.B4 === "Anna Testowa" && o.C4 === todayYmd && o.D4 == null && o.F3 == null, o);
  ok("I. format daty i formuła Długość zostają, tabela bez zmian", o && /mmm/.test(o.fmt) && o.E3 === "=IF" && o.ref === "A2:AG565", o);
  if (ooxml && !MOBILE) {
    try {
      const [b0, a0] = ooxml.validate([RODO, out]);
      const fresh = ooxml.newErrors(b0, a0);
      ok("I. OOXML bez nowych błędów", fresh.length === 0, fresh.slice(0, 3));
    } catch (e) { if (process.env.DEBUG) console.log("ooxml pominięty:", e.message); }
  }

  ok("bez błędów strony", errors.length === 0, errors.slice(0, 5));
  await browser.close();
  const tag = `${ENGINE}${MOBILE ? " 382px dotyk" : ""}`;
  if (failures.length) { console.log(`❌ entry-smart (${tag}): ${failures.length} niezaliczone`); failures.forEach((f) => console.log("  ✗", f)); process.exit(1); }
  console.log(`✅ entry-smart (${tag}): wszystko przeszło`);
}
run().catch((e) => { failures.forEach((f) => console.log("  ✗", f)); console.error(e); process.exit(1); });
