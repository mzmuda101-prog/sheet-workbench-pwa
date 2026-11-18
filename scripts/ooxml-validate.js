// ooxml-validate.js — walidacja .xlsx ze schematem Office przez Open XML SDK Microsoftu
// (tools/ooxml-validator, .NET; kopia z Documents Workbench). Narzędzie testowe na komputerze dewelopera — nie jest częścią
// aplikacji (nie trafia do dist/), pliki nigdzie nie wychodzą.
//
//   node scripts/ooxml-validate.js plik.xlsx [plik2.xlsx …]     (wypisuje błędy, kod 1 przy błędach)
//
// W testach: const { available, validate, newErrors } = require("./ooxml-validate");
// Brak .NET → available() === false, a testy pomijają walidację z ostrzeżeniem (instalacja:
// brew install dotnet).

const { execFileSync, spawnSync } = require("child_process");
const crypto = require("crypto");
const fs = require("fs");
const path = require("path");

const PROJECT = path.resolve(__dirname, "../tools/ooxml-validator");
const DLL = path.join(PROJECT, "bin/Release/net10.0/OoxmlValidator.dll");

let ready = null;
function available() {
  if (ready !== null) return ready;
  const dotnet = spawnSync("dotnet", ["--version"], { encoding: "utf8" });
  if (dotnet.status !== 0) return (ready = false);
  const src = ["Program.cs", "OoxmlValidator.csproj"].map((f) => fs.statSync(path.join(PROJECT, f)).mtimeMs);
  if (!fs.existsSync(DLL) || Math.max(...src) > fs.statSync(DLL).mtimeMs) {
    const b = spawnSync("dotnet", ["build", "-c", "Release", "--nologo", "-v", "q"], { cwd: PROJECT, encoding: "utf8" });
    if (b.status !== 0) { console.error(b.stdout, b.stderr); return (ready = false); }
  }
  return (ready = true);
}

// Wynik dla wzorców z scripts/ (oryginały porównywane w każdym teście, np. RODO ~1 s) pamiętamy
// po hashu pliku i walidatora — wyniki zapisów z testów (pliki tymczasowe) zawsze liczymy od nowa.
const CACHE_DIR = path.resolve(__dirname, "../.test-cache/ooxml");
const SCRIPTS_DIR = __dirname + path.sep;
const cacheKey = (file) => {
  if (!path.resolve(file).startsWith(SCRIPTS_DIR)) return null;
  try {
    const h = crypto.createHash("sha1").update(fs.readFileSync(file)).update(String(fs.statSync(DLL).mtimeMs));
    return path.join(CACHE_DIR, h.digest("hex") + ".json");
  } catch { return null; }
};
const readCache = (p) => { try { return p ? JSON.parse(fs.readFileSync(p, "utf8")) : null; } catch { return null; } };

// [{ file, ok, error?, errors: [{ id, type, part, path, description }] }]
function validate(files) {
  if (!files.length) return [];
  const keys = files.map(cacheKey);
  const res = keys.map((k, i) => { const r = readCache(k); return r && { ...r, file: files[i] }; });
  const todo = files.filter((_, i) => !res[i]);
  if (!todo.length) return res;
  let out;
  try {
    out = execFileSync("dotnet", [DLL, ...todo], { encoding: "utf8", maxBuffer: 64 * 1024 * 1024 });
  } catch (e) {
    out = e.stdout; // kod 1/2 = są błędy; wynik i tak na stdout
  }
  let fresh;
  try {
    fresh = JSON.parse(out);
  } catch (_) {
    throw new Error(`walidator nie zwrócił wyniku (dotnet: ${String(out || "").slice(0, 200) || "pusto"})`);
  }
  let j = 0;
  return res.map((r, i) => {
    if (r) return r;
    const f = fresh[j++];
    if (keys[i] && f && !f.error) try { fs.mkdirSync(CACHE_DIR, { recursive: true }); fs.writeFileSync(keys[i], JSON.stringify(f)); } catch { /* bez pamięci — następnym razem policzy */ }
    return f;
  });
}

// Błędy, których nie było w oryginale (porównanie po części pliku i opisie, z krotnością —
// ścieżki XPath się przesuwają, gdy przybywa akapitów).
function newErrors(before, after) {
  const key = (e) => `${e.part} | ${e.description}`;
  const left = new Map();
  (before?.errors || []).forEach((e) => left.set(key(e), (left.get(key(e)) || 0) + 1));
  return (after?.errors || []).filter((e) => {
    const n = left.get(key(e)) || 0;
    if (n > 0) { left.set(key(e), n - 1); return false; }
    return true;
  });
}

const short = (e) => `${e.part} ${e.path || ""} — ${e.description}`;

module.exports = { available, validate, newErrors, short };

if (require.main === module) {
  const files = process.argv.slice(2);
  if (!files.length) { console.log("Użycie: node scripts/ooxml-validate.js plik.xlsx [...]"); process.exit(2); }
  if (!available()) { console.error("Brak .NET (brew install dotnet) albo walidator się nie zbudował."); process.exit(2); }
  let bad = 0;
  for (const r of validate(files)) {
    console.log(`${r.ok ? "✅" : "❌"} ${path.basename(r.file)}${r.error ? `  (${r.error})` : r.ok ? "" : `  — ${r.errors.length} błędów`}`);
    r.errors.slice(0, 20).forEach((e) => console.log(`     ${short(e)}`));
    if (!r.ok) bad++;
  }
  process.exit(bad ? 1 : 0);
}
