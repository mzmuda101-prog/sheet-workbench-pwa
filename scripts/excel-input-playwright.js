// excel-input-playwright.js — wpisywanie i wyświetlanie wartości jak w POLSKIM Excelu.
//
// Wzorzec: wyniki z Excela (Microsoft Excel na Macu, `formula local` przez AppleScript,
// 2026-10-17) i parser OnlyOffice (sdkjs FormatParser.parse). Co musi być prawdą:
//   1. widok: liczby z formatem po polsku („1 234,50”, „12,5%”, „1,50E+05”), a format
//      „dd.mm.yyyy hh:mm” — na którym SheetJS się wywraca — pokazuje datę Z GODZINĄ,
//   2. parser: 12% → 0,12 [0%], 8:30 → h:mm, PRAWDA → wartość logiczna, '007 → tekst „007”,
//      (100) → −100, 12 zł → 12 [zł], 1E5, data z godziną, daty bez roku (12-03, mar 2025)…,
//   3. wpisanie w komórkę Ogólny nadaje format jak Excel; komórka z własnym formatem go
//      zachowuje (12% w „0.00” = 0,12); w komórce Tekst („@”) „007” zostaje tekstem,
//   4. Cofnij zdejmuje i wartość, i nadany format,
//   5. zapis: openpyxl widzi liczby, formaty, wartość logiczną i godzinę w dacie.

const pw = require("playwright");
const path = require("path");
const fs = require("fs");
const os = require("os");
const { spawnSync } = require("child_process");
const VENV_PY = path.join(__dirname, "..", ".venv/bin/python");
const PY = fs.existsSync(VENV_PY) ? VENV_PY : "python3";

const APP_URL = process.env.APP_URL || "http://127.0.0.1:4175/";
const ENGINE = process.env.ENGINE || "chromium";
const SLEEP_SCALE = Math.min(1, Math.max(0.1, Number(process.env.SLEEP_SCALE || 1)));
const sleep = (ms) => new Promise((r) => setTimeout(r, Math.round(ms * SLEEP_SCALE)));

const tmp = fs.mkdtempSync(path.join(os.tmpdir(), "excel-input-"));
const FILE = path.join(tmp, "wpisy.xlsx");

// Arkusz „Wpisy”: A = opis, B = liczby w różnych formatach (do widoku), C = Ogólny (puste),
// D = Tekst „@”, E = „0.00”. Wiersze 2–15.
const gen = spawnSync(PY, ["-c", `
import sys, datetime
from openpyxl import Workbook
wb = Workbook(); ws = wb.active; ws.title = "Wpisy"
ws.append(["Opis", "Wartosc", "Ogolny", "Tekst", "Dziesietne"])
rows = [("tys2", 1234.5, '#,##0.00'), ("tys0", 1234567, '#,##0'), ("proc", 0.125, '0.0%'), ("zl", 1234.5, '#,##0.00 "zł"'),
  ("czas", 0.3541666666666667, 'h:mm'), ("dt", datetime.datetime(2025, 3, 12, 10, 30), 'dd.mm.yyyy hh:mm'),
  ("data", datetime.date(2025, 3, 12), 'dd.mm.yyyy'), ("dec", 3.14159, '0.00'), ("ogolny", 12.5, 'General'),
  ("neg", -1234.5, '#,##0.00;[Red]-#,##0.00'), ("sci", 150000, '0.00E+00'), ("godz", 1.0625, '[h]:mm'), ("bool", True, 'General'), ("x", 1, 'General')]
for d, v, f in rows:
    ws.append([d, v, None, None, None]); r = ws.max_row
    ws.cell(r, 2).number_format = f
    ws.cell(r, 4).number_format = "@"
    ws.cell(r, 5).number_format = "0.00"
    ws.cell(r, 5).value = 1
wb.save(sys.argv[1])
`, FILE], { encoding: "utf8" });
if (gen.status !== 0) { console.error("❌ excel-input: fixture —", gen.stderr); process.exit(1); }

async function run() {
  const browser = await pw[ENGINE].launch({ headless: true });
  const context = await browser.newContext({ serviceWorkers: "block", viewport: { width: 1280, height: 860 } });
  await context.addInitScript(() => { localStorage.setItem("introPlayed", "true"); });
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  page.on("dialog", (d) => d.accept());
  const failures = [];
  const ok = (name, cond, got) => { if (!cond) failures.push(`${name} — dostałem: ${JSON.stringify(got)}`); };

  await page.goto(APP_URL, { waitUntil: "load" });
  await page.evaluate(() => { document.getElementById("heroSplash")?.remove(); return ensureXlsxLibs(false); });
  await page.setInputFiles("#fileInput", FILE);
  await page.waitForFunction(() => currentFileName === "wpisy.xlsx" && !!document.querySelector("#dataTable tbody tr[data-row-key]"), null, { timeout: 20000 });
  await sleep(400);

  const td = (r, c) => `#dataTable tbody tr[data-row-key="wide:${r}"] td[data-col-index="${c}"]`;
  const text = (r, c) => page.evaluate((sel) => document.querySelector(sel)?.textContent ?? null, td(r, c));

  // 1. widok — to, co pokazuje polski Excel (te same formaty, sprawdzone w Excelu)
  const shown = await page.evaluate(() => {
    const out = {};
    document.querySelectorAll("#dataTable tbody tr[data-row-key]").forEach((tr) => {
      const k = tr.querySelector('td[data-col-index="0"]')?.textContent;
      out[k] = tr.querySelector('td[data-col-index="1"]')?.textContent.replace(/ /g, " ");
    });
    return out;
  });
  const wantShown = {
    tys2: "1 234,50", tys0: "1 234 567", proc: "12,5%", zl: "1 234,50 zł", czas: "8:30", dt: "12.03.2025 10:30",
    data: "12.03.2025", dec: "3,14", ogolny: "12,5", neg: "-1 234,50", sci: "1,50E+05", godz: "25:30", bool: "PRAWDA",
  };
  const badShown = Object.keys(wantShown).filter((k) => shown[k] !== wantShown[k]).map((k) => `${k}: „${shown[k]}” (Excel: „${wantShown[k]}”)`);
  ok("1. liczby, daty i PRAWDA/FAŁSZ wyglądają jak w polskim Excelu", badShown.length === 0, badShown);

  // 2. parser — wynik Excela dla każdego wpisu (komórka w formacie Ogólny)
  const parsed = await page.evaluate(() => {
    const y = new Date().getFullYear();
    const fmtD = (d) => `${d.getFullYear() === y ? "Y" : d.getFullYear()}-${d.getMonth() + 1}-${d.getDate()}` + (d.getHours() || d.getMinutes() || d.getSeconds() ? ` ${d.getHours()}:${d.getMinutes()}:${d.getSeconds()}` : "");
    return Object.fromEntries(["12%", "12,5%", "-12%", "1 234", "1 234,5", "(100)", "12 zł", "12zł", "(12 zł)", "12,50 zł", "zł 12",
      "8:30", "8:30:15", "25:30", "8:30 PM", "12.03.2025 10:30", "2025-03-12 10:30", "PRAWDA", "fałsz", "TRUE", "1E5", "1,5E3", "'007", "'=A1",
      "007", "5,", "1.", ",5", "--5", "1/2", "12-03", "3/2025", "mar 2025", "marzec 2025", "12 mar", "12-mar-2025", "31.02.2025", "(-5)", "12e"].map((s) => {
      const r = parseInputValue(s);
      const v = !r ? "null" : r.value instanceof Date ? fmtD(r.value) : String(r.value);
      return [s, r ? `${r.type}:${v}${r.numFmt ? ` [${r.numFmt}]` : ""}` : "null"];
    }));
  });
  const wantParsed = {
    "12%": "number:0.12 [0%]", "12,5%": "number:0.125 [0.00%]", "-12%": "number:-0.12 [0%]", "1 234": "number:1234 [#,##0]",
    "1 234,5": "number:1234.5 [#,##0.00]", "(100)": "number:-100", "12 zł": 'number:12 [#,##0 "zł"]', "12zł": 'number:12 [#,##0 "zł"]',
    "(12 zł)": 'number:-12 [#,##0 "zł"]', "12,50 zł": 'number:12.5 [#,##0.00 "zł"]', "zł 12": "string:zł 12",
    "8:30": `number:${8.5 / 24} [h:mm]`, "8:30:15": `number:${(8 * 3600 + 30 * 60 + 15) / 86400} [h:mm:ss]`, "25:30": "number:1.0625 [[h]:mm:ss]",
    "8:30 PM": `number:${20.5 / 24} [h:mm AM/PM]`, "12.03.2025 10:30": "date:2025-3-12 10:30:0 [dd.mm.yyyy hh:mm]",
    "2025-03-12 10:30": "date:2025-3-12 10:30:0 [dd.mm.yyyy hh:mm]", "PRAWDA": "boolean:true", "fałsz": "boolean:false", "TRUE": "string:TRUE",
    "1E5": "number:100000 [0.00E+00]", "1,5E3": "number:1500 [0.00E+00]", "'007": "string:007", "'=A1": "string:=A1",
    "007": "number:7", "5,": "number:5", "1.": "string:1.", ",5": "number:0.5", "--5": "string:--5",
    "1/2": "date:Y-2-1", "12-03": "date:Y-3-12", "3/2025": "date:2025-3-1", "mar 2025": "date:2025-3-1", "marzec 2025": "date:2025-3-1",
    "12 mar": "date:Y-3-12", "12-mar-2025": "date:2025-3-12", "31.02.2025": "string:31.02.2025", "(-5)": "string:(-5)", "12e": "string:12e",
  };
  const badParsed = Object.keys(wantParsed).filter((k) => parsed[k] !== wantParsed[k]).map((k) => `${k} → ${parsed[k]} (Excel: ${wantParsed[k]})`);
  ok("2. wpisy rozpoznane jak w polskim Excelu", badParsed.length === 0, badParsed);

  // 3. wpisywanie w komórki: C = Ogólny, D = Tekst, E = 0.00
  const type = async (r, c, s) => {
    await page.dblclick(td(r, c));
    await page.waitForSelector(`${td(r, c)} input.cell-editor`, { timeout: 3000 });
    await page.keyboard.press("ControlOrMeta+a");
    await page.keyboard.type(s);
    await page.keyboard.press("Enter");
    await sleep(200);
  };
  const cellAt = (ref) => page.evaluate((ref) => { const c = workbook.Sheets[currentSheetName][ref]; return c ? { t: c.t, v: c.v instanceof Date ? c.v.toISOString() : c.v, z: c.z } : null; }, ref);
  await type(1, 2, "12%");
  await type(2, 2, "8:30");
  await type(3, 2, "PRAWDA");
  await type(4, 2, "'007");
  await type(5, 2, "12.03.2025 10:30");
  await type(6, 2, "1 234");
  await type(7, 2, "12 zł");
  await type(1, 3, "007");
  await type(1, 4, "12%");
  const typed = {
    c2: [await text(1, 2), await cellAt("C2")], c3: [await text(2, 2), await cellAt("C3")], c4: [await text(3, 2), await cellAt("C4")],
    c5: [await text(4, 2), await cellAt("C5")], c6: [await text(5, 2), await cellAt("C6")], c7: [(await text(6, 2) || "").replace(/ /g, " "), await cellAt("C7")],
    c8: [await text(7, 2), await cellAt("C8")], d2: [await text(1, 3), await cellAt("D2")], e2: [await text(1, 4), await cellAt("E2")],
  };
  ok("3. „12%” w Ogólnym → 0,12 w formacie 0%, widać „12%”", typed.c2[0] === "12%" && typed.c2[1].v === 0.12 && typed.c2[1].z === "0%", typed.c2);
  ok("3. „8:30” → godzina (h:mm), widać „8:30”", typed.c3[0] === "8:30" && Math.abs(typed.c3[1].v - 8.5 / 24) < 1e-12 && typed.c3[1].z === "h:mm", typed.c3);
  ok("3. „PRAWDA” → wartość logiczna, widać PRAWDA", typed.c4[0] === "PRAWDA" && typed.c4[1].t === "b" && typed.c4[1].v === true, typed.c4);
  ok("3. „'007” → tekst „007” bez apostrofu", typed.c5[0] === "007" && typed.c5[1].t === "s" && typed.c5[1].v === "007", typed.c5);
  ok("3. „12.03.2025 10:30” → data z godziną, widać godzinę", typed.c6[0] === "12.03.2025 10:30" && typed.c6[1].t === "d" && typed.c6[1].z === "dd.mm.yyyy hh:mm", typed.c6);
  ok("3. „1 234” → 1234 w formacie #,##0, widać „1 234”", typed.c7[0] === "1 234" && typed.c7[1].v === 1234 && typed.c7[1].z === "#,##0", typed.c7);
  ok("3. „12 zł” → 12 w formacie złotowym", typed.c8[0] === "12 zł" && typed.c8[1].v === 12 && /zł/.test(typed.c8[1].z || ""), typed.c8);
  ok("3. w kolumnie Tekst „007” zostaje tekstem", typed.d2[0] === "007" && typed.d2[1].t === "s" && typed.d2[1].v === "007", typed.d2);
  ok("3. „12%” w formacie 0.00 → 0,12 i format 0.00 zostaje (jak Excel)", typed.e2[0] === "0,12" && typed.e2[1].v === 0.12 && typed.e2[1].z === "0.00", typed.e2);

  // 4. Cofnij zdejmuje wartość i nadany format
  await type(8, 2, "50%");
  const before = await page.evaluate(() => ({ p: JSON.stringify(pendingEdits[currentSheetName].C9 || null) }));
  await page.keyboard.press("ControlOrMeta+z");
  await sleep(250);
  const after = await page.evaluate(() => ({ cell: workbook.Sheets[currentSheetName].C9 || null, p: pendingEdits[currentSheetName].C9 || null, numFmt: cfmtNumFmtAt(8, 2) }));
  ok("4. Cofnij po „50%” — komórka znowu pusta i w formacie Ogólny", /0%/.test(before.p) && !(after.cell && after.cell.v != null) && after.numFmt === "General", { before, after });

  // 5. zapis i drugi czytnik (openpyxl)
  const out = path.join(tmp, "zapis.xlsx");
  const b64 = await page.evaluate(async () => {
    const u8 = await buildOutputBytes("xlsx");
    let s = ""; for (let i = 0; i < u8.length; i += 0x8000) s += String.fromCharCode(...u8.subarray(i, i + 0x8000));
    return btoa(s);
  });
  fs.writeFileSync(out, Buffer.from(b64, "base64"));
  const py = spawnSync(PY, ["-c", `
import json, sys
from openpyxl import load_workbook
ws = load_workbook(sys.argv[1])["Wpisy"]
def c(ref):
    x = ws[ref]
    return {"v": None if x.value is None else str(x.value), "type": type(x.value).__name__, "fmt": x.number_format}
print(json.dumps({k: c(k) for k in ["C2", "C3", "C4", "C5", "C6", "C7", "C8", "D2", "E2", "C9"]}))
`, out], { encoding: "utf8" });
  let f = null;
  try { f = JSON.parse(py.stdout); } catch { failures.push("5. openpyxl nie przeczytał zapisu: " + (py.stderr || "").slice(0, 300)); }
  if (f) {
    ok("5. plik: 12% = 0.12 w formacie 0%", f.C2.v === "0.12" && f.C2.fmt === "0%", f.C2);
    ok("5. plik: 8:30 w formacie h:mm", f.C3.fmt === "h:mm" && f.C3.v != null, f.C3);
    ok("5. plik: PRAWDA jako wartość logiczna", f.C4.type === "bool" && f.C4.v === "True", f.C4);
    ok("5. plik: '007 jako tekst", f.C5.type === "str" && f.C5.v === "007", f.C5);
    ok("5. plik: data z godziną 10:30", f.C6.type === "datetime" && /2025-03-12 10:30/.test(f.C6.v) && /hh:mm/.test(f.C6.fmt), f.C6);
    ok("5. plik: 1 234 w formacie #,##0", f.C7.v === "1234" && f.C7.fmt === "#,##0", f.C7);
    ok("5. plik: 12 zł w formacie złotowym", f.C8.v === "12" && /zł/.test(f.C8.fmt), f.C8);
    ok("5. plik: „007” w kolumnie Tekst jako tekst", f.D2.type === "str" && f.D2.v === "007" && f.D2.fmt === "@", f.D2);
    ok("5. plik: 0,12 w formacie 0.00", f.E2.v === "0.12" && f.E2.fmt === "0.00", f.E2);
    ok("5. plik: cofnięte 50% nie zostało w pliku", f.C9.v === null && f.C9.fmt === "General", f.C9);
  }
  fs.rmSync(tmp, { recursive: true, force: true });

  await browser.close();
  if (errors.length) failures.push("błędy strony: " + errors.join(" | "));
  if (failures.length) {
    console.error(`❌ excel-input (${ENGINE}):\n  - ` + failures.join("\n  - "));
    process.exit(1);
  }
  console.log(`✅ excel-input (${ENGINE}): widok liczb/dat jak w polskim Excelu, wpisy %, zł, godziny, PRAWDA, ' i format Tekst, Cofnij, zapis`);
}

run().catch((e) => { console.error(e); process.exit(1); });
