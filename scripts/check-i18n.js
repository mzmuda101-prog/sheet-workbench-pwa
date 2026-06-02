#!/usr/bin/env node
/**
 * check-i18n.js — walidator dev-time tłumaczeń.
 *
 * Łapie całą klasę bugów „surowy klucz / undefined" wynikającą z dwóch słowników
 * (I18N via t(), STATIC_TRANSLATIONS via copy.X w applyStaticTranslations):
 *   1) każdy literalny `t("klucz")` w app/*.js ma pokrycie w którymkolwiek słowniku,
 *   2) każdy `copy.klucz` w language.js ma pokrycie,
 *   3) parytet PL/EN — klucz obecny w jednym języku, brakuje w drugim.
 *
 * Dynamicznie sklejane klucze (np. t("aggGroupKind_" + kind)) są pomijane —
 * regex celuje tylko w literały. Uruchamiane w `npm test` przed smoke-testem.
 *
 * Exit 1 przy błędach krytycznych (brak pokrycia); parytet to ostrzeżenie.
 */

const fs = require("fs");
const path = require("path");

const ROOT = path.resolve(__dirname, "..");
const APP_DIR = path.join(ROOT, "app");
const LANG_FILE = path.join(APP_DIR, "language.js");

// ── Wyłuskanie literału obiektu po nazwie const (dopasowanie nawiasów) ──────────
function extractObjectLiteral(source, constName) {
  const marker = `const ${constName} = {`;
  const start = source.indexOf(marker);
  if (start < 0) throw new Error(`Nie znaleziono ${constName} w language.js`);
  let i = source.indexOf("{", start);
  let depth = 0;
  let inStr = null;
  for (let j = i; j < source.length; j++) {
    const ch = source[j];
    const prev = source[j - 1];
    if (inStr) {
      if (ch === inStr && prev !== "\\") inStr = null;
      continue;
    }
    if (ch === '"' || ch === "'" || ch === "`") { inStr = ch; continue; }
    if (ch === "{") depth++;
    else if (ch === "}") {
      depth--;
      if (depth === 0) {
        const literal = source.slice(i, j + 1);
        // eslint-disable-next-line no-new-func
        return Function(`"use strict"; return (${literal});`)();
      }
    }
  }
  throw new Error(`Nie udało się dopasować nawiasów dla ${constName}`);
}

const langSrc = fs.readFileSync(LANG_FILE, "utf8");
const I18N = extractObjectLiteral(langSrc, "I18N");
const STATIC = extractObjectLiteral(langSrc, "STATIC_TRANSLATIONS");

const keysOf = (dict, lang) => new Set(Object.keys((dict && dict[lang]) || {}));
const i18nPl = keysOf(I18N, "pl");
const i18nEn = keysOf(I18N, "en");
const staticPl = keysOf(STATIC, "pl");
const staticEn = keysOf(STATIC, "en");

const available = new Set([...i18nPl, ...i18nEn, ...staticPl, ...staticEn]);

// ── Zbierz użycia ───────────────────────────────────────────────────────────────
const appFiles = fs.readdirSync(APP_DIR).filter((f) => f.endsWith(".js")).map((f) => path.join(APP_DIR, f));

const tUses = new Map();    // key -> [files]
const copyUses = new Map();

const T_RE = /\bt\(\s*(["'`])([A-Za-z0-9_]+)\1/g;
const COPY_RE = /\bcopy\.([A-Za-z0-9_]+)/g;

for (const file of appFiles) {
  const src = fs.readFileSync(file, "utf8");
  const name = path.basename(file);
  let m;
  while ((m = T_RE.exec(src))) {
    // Pomiń dynamiczne prefiksy: t("profType_" + kind) — po stringu jest `+`.
    let p = T_RE.lastIndex;
    while (p < src.length && /\s/.test(src[p])) p++;
    if (src[p] === "+") continue;
    const k = m[2];
    if (!tUses.has(k)) tUses.set(k, new Set());
    tUses.get(k).add(name);
  }
}
// copy.X istnieje tylko w language.js (zmienna lokalna applyStaticTranslations)
{
  let m;
  while ((m = COPY_RE.exec(langSrc))) {
    const k = m[1];
    if (!copyUses.has(k)) copyUses.set(k, new Set());
    copyUses.get(k).add("language.js");
  }
}
// copy.title/description itd. to zwykłe pola — sprawdzamy je tak samo.

// ── Walidacja ─────────────────────────────────────────────────────────────────
const errors = [];
const warnings = [];

for (const [key, files] of [...tUses].sort()) {
  if (!available.has(key)) {
    errors.push(`t("${key}") — brak w jakimkolwiek słowniku (użyte w: ${[...files].join(", ")})`);
  }
}
for (const [key, files] of [...copyUses].sort()) {
  if (!available.has(key)) {
    errors.push(`copy.${key} — brak w jakimkolwiek słowniku (użyte w: ${[...files].join(", ")})`);
  }
}

// Parytet PL/EN (ostrzeżenia) — osobno dla każdego słownika.
const parity = (plSet, enSet, label) => {
  for (const k of [...plSet].sort()) if (!enSet.has(k)) warnings.push(`${label}: "${k}" jest w PL, brak w EN`);
  for (const k of [...enSet].sort()) if (!plSet.has(k)) warnings.push(`${label}: "${k}" jest w EN, brak w PL`);
};
parity(i18nPl, i18nEn, "I18N");
parity(staticPl, staticEn, "STATIC_TRANSLATIONS");

// ── Raport ──────────────────────────────────────────────────────────────────────
console.log(`i18n: I18N(${i18nPl.size} pl / ${i18nEn.size} en), STATIC(${staticPl.size} pl / ${staticEn.size} en)`);
console.log(`Użycia: ${tUses.size} literalnych t(), ${copyUses.size} copy.X`);

if (warnings.length) {
  console.log(`\n⚠️  Parytet PL/EN (${warnings.length}):`);
  warnings.forEach((w) => console.log(`   • ${w}`));
}
if (errors.length) {
  console.error(`\n❌  Braki pokrycia (${errors.length}):`);
  errors.forEach((e) => console.error(`   • ${e}`));
  console.error("\nNaprawa: dodaj klucz do I18N lub STATIC_TRANSLATIONS (oba języki).");
  process.exit(1);
}
console.log("\n✅  Wszystkie literalne klucze t()/copy.X mają pokrycie.");
