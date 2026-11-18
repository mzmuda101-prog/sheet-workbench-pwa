// build-release-test.mjs — strażnik buildu produkcji (dist/, Vercel).
//
// Po `npm run build` każdy plik, do którego odwołuje się dist/index.html (<link href>, <script src>)
// i każdy wpis list SHELL_ASSETS / HEAVY_ASSETS w dist/sw.js musi istnieć w dist/.
// Powód (2026-11-18): build kopiował tylko styles/app.css, więc nowy styles/ribbon.css nie trafił
// na produkcję — wstążka bez stylów, a precache service workera dostawał 404.

import { execFileSync } from "node:child_process";
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const DIST = path.join(ROOT, "dist");

execFileSync(process.execPath, [path.join(ROOT, "scripts", "build-release.mjs")], { cwd: ROOT, stdio: "ignore" });

const local = (u) => {
  const p = u.split(/[?#]/)[0].replace(/^\.\//, "");
  return p && !/^(https?:)?\/\//.test(p) && !p.startsWith("data:") ? p : null;
};
const refs = new Set();
const html = fs.readFileSync(path.join(DIST, "index.html"), "utf8");
for (const m of html.matchAll(/<(?:link|script)\b[^>]*\b(?:href|src)="([^"]+)"/g)) {
  const p = local(m[1]);
  if (p) refs.add(p);
}
const sw = fs.readFileSync(path.join(DIST, "sw.js"), "utf8");
for (const list of ["SHELL_ASSETS", "HEAVY_ASSETS"]) {
  const block = sw.match(new RegExp(`const ${list} = \\[([\\s\\S]*?)\\];`));
  if (!block) throw new Error(`Brak ${list} w dist/sw.js`);
  for (const m of block[1].matchAll(/[`"']([^`"']+)[`"']/g)) {
    const p = local(m[1].replace(/\$\{[^}]+\}/g, ""));
    if (p && p !== "/") refs.add(p);
  }
}

const missing = [...refs].filter((p) => !fs.existsSync(path.join(DIST, p)));
if (missing.length) {
  console.error(`❌ build: ${missing.length} plików, do których odwołuje się index.html / sw.js, nie ma w dist/:\n  ${missing.join("\n  ")}`);
  process.exit(1);
}
console.log(`✅ build-release-test OK — ${refs.size} odwołań, wszystkie pliki są w dist/`);
