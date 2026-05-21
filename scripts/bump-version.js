#!/usr/bin/env node
/**
 * bump-version.js
 * Aktualizuje wersję cache/buildu we wszystkich plikach projektu jedną komendą.
 *
 * Użycie:
 *   npm run release            → auto-generuje datę dzisiejszą jako "YYYYMMDD-01"
 *   npm run release 20260601-02  → użyje podanej wersji
 */

const fs = require("fs");
const path = require("path");

// ─── Ustal nową wersję ────────────────────────────────────────────────────────

function todayVersion() {
  const now = new Date();
  const y = now.getFullYear();
  const m = String(now.getMonth() + 1).padStart(2, "0");
  const d = String(now.getDate()).padStart(2, "0");
  return `${y}${m}${d}-01`;
}

const newVersion = process.argv[2] || todayVersion();
if (!/^\d{8}-\d{2}$/.test(newVersion)) {
  console.error(`❌  Nieprawidłowy format wersji: "${newVersion}". Oczekiwany format: YYYYMMDD-NN (np. 20260601-01)`);
  process.exit(1);
}

// ─── Pomocnicze ───────────────────────────────────────────────────────────────

const ROOT = path.resolve(__dirname, "..");

function read(relPath) {
  return fs.readFileSync(path.join(ROOT, relPath), "utf8");
}

function write(relPath, content) {
  fs.writeFileSync(path.join(ROOT, relPath), content, "utf8");
}

/**
 * Zastępuje PIERWSZE dopasowanie wzorca w tekście.
 * Rzuca błąd jeśli wzorzec nie został znaleziony.
 */
function replaceFirst(content, pattern, replacement, label) {
  if (!pattern.test(content)) {
    throw new Error(`Nie znaleziono wzorca w ${label}`);
  }
  return content.replace(pattern, replacement);
}

// ─── Zbierz aktualną wersję (z sw.js) ────────────────────────────────────────

const swRaw = read("sw.js");
const currentMatch = swRaw.match(/const CACHE_VERSION = "(\d{8}-\d{2})"/);
const oldVersion = currentMatch ? currentMatch[1] : null;

if (!oldVersion) {
  console.error("❌  Nie znaleziono CACHE_VERSION w sw.js");
  process.exit(1);
}

if (oldVersion === newVersion) {
  console.log(`ℹ️  Wersja jest już aktualna: ${newVersion}. Nic nie zmieniono.`);
  process.exit(0);
}

console.log(`\n🔄  Aktualizacja wersji: ${oldVersion} → ${newVersion}\n`);

const versionRegex = new RegExp(oldVersion.replace("-", "\\-"), "g");
const changes = [];

// ─── 1. sw.js ─────────────────────────────────────────────────────────────────

{
  const file = "sw.js";
  const before = read(file);
  const after = before.replace(versionRegex, newVersion);
  write(file, after);
  const count = (before.match(versionRegex) || []).length;
  changes.push(`✅  ${file} — zaktualizowano ${count} wystąpień`);
}

// ─── 2. app/core.js ───────────────────────────────────────────────────────────

{
  const file = "app/core.js";
  const before = read(file);
  const after = replaceFirst(
    before,
    /const APP_BUILD_VERSION = "\d{8}-\d{2}"/,
    `const APP_BUILD_VERSION = "${newVersion}"`,
    file
  );
  write(file, after);
  changes.push(`✅  ${file} — APP_BUILD_VERSION zaktualizowany`);
}

// ─── 3. index.html ────────────────────────────────────────────────────────────

{
  const file = "index.html";
  const before = read(file);
  const after = before.replace(versionRegex, newVersion);
  const count = (before.match(versionRegex) || []).length;
  write(file, after);
  changes.push(`✅  ${file} — zaktualizowano ${count} wystąpień`);
}

// ─── 4. scripts/smoke-playwright.js ──────────────────────────────────────────

{
  const file = "scripts/smoke-playwright.js";
  const before = read(file);
  const after = before.replace(versionRegex, newVersion);
  const count = (before.match(versionRegex) || []).length;
  if (count > 0) {
    write(file, after);
    changes.push(`✅  ${file} — zaktualizowano ${count} wystąpień`);
  } else {
    changes.push(`⚠️   ${file} — brak wystąpień starej wersji (pominięto)`);
  }
}

// ─── Podsumowanie ─────────────────────────────────────────────────────────────

console.log(changes.join("\n"));
console.log(`\n🎉  Gotowe! Nowa wersja: ${newVersion}`);
console.log(`\nNastępne kroki:`);
console.log(`  git add -A`);
console.log(`  git commit -m "release: ${newVersion}"`);
console.log(`  git push`);
