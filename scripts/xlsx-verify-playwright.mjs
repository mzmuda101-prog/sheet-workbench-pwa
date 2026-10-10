// xlsx-verify-playwright.mjs — zapis aplikacji sprawdzany DRUGIM parserem (openpyxl).
//
// save-roundtrip/save-stress sprawdzają wynik tym samym SheetJS, który bierze udział
// w zapisie. Tu: plik → 2 cykle edycji przez buildPatchedXlsx (jak w aplikacji: każdy
// cykl czyta POPRZEDNI zapis) → scripts/py/xlsx_verify.py porównuje oryginał z wynikiem
// komórka po komórce (wartości, formuły, style, DV, CF, scalenia, tabele, nazwy…).
//
// Pliki: syntetyczny z gen_fixtures.py (zawsze) + stress-test + prywatne pliki ze scripts/,
// jeśli leżą na dysku. FILES="a.xlsx,b.xlsx" podmienia listę.
// Wymaga: npm run py:setup (jednorazowo) i serwera na APP_URL (npm test startuje go sam).

import { execFileSync, spawnSync } from "node:child_process";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { chromium } from "playwright";
import { createRequire } from "node:module";

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const APP_URL = process.env.APP_URL || "http://127.0.0.1:4175/";
const VENV_PY = path.join(ROOT, ".venv/bin/python");
const PY = fs.existsSync(VENV_PY) ? VENV_PY : "python3";
const VERIFY = path.join(ROOT, "scripts/py/xlsx_verify.py");

if (spawnSync(PY, ["-c", "import openpyxl"]).status !== 0) {
  console.error("❌ xlsx-verify: brak openpyxl — uruchom raz: npm run py:setup");
  process.exit(1);
}

const fixture = execFileSync(PY, [path.join(ROOT, "scripts/py/gen_fixtures.py")], { encoding: "utf8" }).trim();
const files = process.env.FILES
  ? process.env.FILES.split(",").map((f) => path.resolve(ROOT, f.trim()))
  : [fixture, ...["stress-test-workbench.xlsx", "Kosztorys Renault Laguna 3 .xlsx", "RODO Obieg terenów 2026 V5.xlsx"]
      .map((f) => path.join(ROOT, "scripts", f))
      .filter((f) => fs.existsSync(f))];

// OUT_DIR=… zostawia zapisane pliki (np. do otwarcia w prawdziwym Excelu).
const tmp = process.env.OUT_DIR ? (fs.mkdirSync(process.env.OUT_DIR, { recursive: true }), process.env.OUT_DIR)
  : fs.mkdtempSync(path.join(os.tmpdir(), "xlsx-verify-"));
const browser = await chromium.launch({ headless: true });
const context = await browser.newContext({ serviceWorkers: "block" });
await context.addInitScript(() => localStorage.setItem("introPlayed", "true"));
const page = await context.newPage();
const pageErrors = [];
page.on("pageerror", (e) => pageErrors.push(e.message));
await page.goto(APP_URL, { waitUntil: "load" });
await page.waitForFunction(() => typeof ensureXlsxLibs === "function" && typeof buildPatchedXlsx === "function");
await page.evaluate(() => ensureXlsxLibs(false));
await page.waitForFunction(() => typeof JSZip !== "undefined" && typeof XLSX !== "undefined");

// Cykle zapisu dokładnie jak w aplikacji: mapy formuł z bieżących bajtów → patch → kolejny cykl.
async function appSave(bytes, cycles) {
  const b64 = await page.evaluate(async ({ b64, cycles }) => {
    const toBytes = (s) => Uint8Array.from(atob(s), (ch) => ch.charCodeAt(0));
    const toB64 = (u8) => { let s = ""; for (let i = 0; i < u8.length; i += 0x8000) s += String.fromCharCode(...u8.subarray(i, i + 0x8000)); return btoa(s); };
    const formulaMaps = (b) => {
      const wb = XLSX.read(b, { cellFormula: true, cellStyles: true });
      const maps = {};
      for (const name of wb.SheetNames) {
        const sh = wb.Sheets[name]; const m = {};
        for (const ref in sh) if (ref[0] !== "!" && sh[ref] && sh[ref].f) m[ref] = sh[ref].f;
        if (Object.keys(m).length) maps[name] = m;
      }
      return maps;
    };
    let out = toBytes(b64);
    for (const cycle of cycles) {
      const edits = {};
      for (const [sheet, cells] of Object.entries(cycle)) {
        edits[sheet] = {};
        for (const [ref, p] of Object.entries(cells)) {
          if (p && p.t === "d") {  // data z JSON (ISO, czas lokalny) → Date jak w edytorze komórki
            const [y, mo, d] = p.v.slice(0, 10).split("-").map(Number);
            edits[sheet][ref] = { t: "d", v: new Date(y, mo - 1, d) };
          } else edits[sheet][ref] = p;
        }
      }
      out = await buildPatchedXlsx(out, edits, formulaMaps(out));
    }
    return toB64(out);
  }, { b64: Buffer.from(bytes).toString("base64"), cycles });
  return Buffer.from(b64, "base64");
}

// Schemat Office (Open XML SDK, tools/ooxml-validator) — łapie to, czego openpyxl nie widzi:
// złą kolejność elementów w <font>/<border>/<xf>, złe atrybuty. Liczą się tylko NOWE błędy
// (prawdziwe pliki z innych programów bywają już nie w pełni zgodne). Bez .NET — pomijane
// z ostrzeżeniem; walidator, który się nie uruchomił, to BŁĄD (nie „brak nowych błędów”).
const ooxml = createRequire(import.meta.url)("./ooxml-validate.js");
const hasOoxml = ooxml.available();
function newSchemaErrors(orig, saved) {
  const [vo, vs] = ooxml.validate([orig, saved]);
  if (vs.error) return [`zapis nie otwiera się w Open XML SDK: ${vs.error}`];
  return ooxml.newErrors(vo, vs).map(ooxml.short);
}

function pyVerify(orig, saved, cyclesFile) {
  const r = spawnSync(PY, [VERIFY, "verify", orig, saved, cyclesFile], { encoding: "utf8" });
  return { ok: r.status === 0, out: (r.stdout || "") + (r.stderr || "") };
}

const failed = [];
let canaryDone = false;
for (const file of files) {
  const name = path.basename(file);
  const cycles = JSON.parse(execFileSync(PY, [VERIFY, "plan", file], { encoding: "utf8" }));
  const cyclesFile = path.join(tmp, name + ".edits.json");
  fs.writeFileSync(cyclesFile, JSON.stringify(cycles));
  const saved = path.join(tmp, name.replace(/\.xlsx$/i, "") + ".saved.xlsx");
  fs.writeFileSync(saved, await appSave(fs.readFileSync(file), cycles));

  const res = pyVerify(file, saved, cyclesFile);
  const nEdits = cycles.reduce((n, c) => n + Object.values(c).reduce((m, cells) => m + Object.keys(cells).length, 0), 0);
  console.log(`${res.ok ? "✅" : "❌"} ${name} — ${Object.keys(cycles[0]).length} ark., ${nEdits} edycji w 2 cyklach`);
  if (res.out.trim()) console.log(res.out.replace(/\s+$/, ""));
  if (!res.ok) failed.push(name);
  if (hasOoxml) {
    const schema = newSchemaErrors(file, saved);
    if (schema.length) {
      console.log(`  ❌ schemat Office: ${schema.length} nowych błędów`);
      schema.slice(0, 8).forEach((e) => console.log("     " + e));
      failed.push(name + " (schemat)");
    }
  }

  // KANARZEK: ten sam zapis porównany BEZ listy edycji musi zostać odrzucony —
  // inaczej weryfikator niczego nie widzi i zielony wynik nic nie znaczy.
  if (!canaryDone) {
    const empty = path.join(tmp, "empty.json");
    fs.writeFileSync(empty, "[]");
    if (pyVerify(file, saved, empty).ok) failed.push("kanarek: weryfikator nie zauważył edycji");
    canaryDone = true;
  }
}

await browser.close();
if (!process.env.OUT_DIR) fs.rmSync(tmp, { recursive: true, force: true });
if (pageErrors.length) failed.push("błędy strony: " + pageErrors.join(" | "));
if (failed.length) {
  console.error("\n❌ xlsx-verify FAIL: " + failed.join("; "));
  process.exit(1);
}
console.log(`\n✅ xlsx-verify: ${files.length} plik(ów) zgodnych wg openpyxl${hasOoxml ? " i schematu Office" : " (schemat Office pominięty — brak tools/ooxml-validator)"}`);
