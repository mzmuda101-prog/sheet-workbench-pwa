// fetch-accuracy-suite.mjs — jednorazowe pobranie danych „Spreadsheet Accuracy Suite” (MIT,
// github.com/MacJedi42/Spreadsheet-Accuracy-Suite) do .cache/accuracy-suite (poza repozytorium).
// Używa ich scripts/formula-accuracy-suite.mjs. ~17 MB: formula_golden.tsv + zasady + licencja.
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const DIR = path.join(ROOT, ".cache/accuracy-suite");
const BASE = "https://raw.githubusercontent.com/MacJedi42/Spreadsheet-Accuracy-Suite/main/";
const FILES = ["goldens/formula/formula_golden.tsv", "EXCEL-BEHAVIOUR.md", "LICENSE", "README.md"];
fs.mkdirSync(DIR, { recursive: true });
for (const f of FILES) {
  const res = await fetch(BASE + f);
  if (!res.ok) { console.error(`❌ ${f}: HTTP ${res.status}`); process.exit(1); }
  fs.writeFileSync(path.join(DIR, path.basename(f)), Buffer.from(await res.arrayBuffer()));
  console.log(`✅ ${f}`);
}
console.log(`Gotowe: ${DIR}`);
