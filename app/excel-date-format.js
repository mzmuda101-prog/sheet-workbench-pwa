// Daty z nazwami miesięcy wg WŁASNEGO formatu z pliku — tak jak pokazuje je Excel.
// Ładowany w wątku głównym (index.html) i w wątku roboczym (build-rows-worker.js).
//
// Wyrocznia Excela (Mac, system PL, 2026-11-01, data 2026-09-08):
//   d/mmm/yyyy → 8.wrz.2026        [$-415]d/mmm/yyyy;@ → 8.wrz.2026     dd/mmm/yy → 08.wrz.26
//   d mmmm yyyy → 8 wrzesień 2026   dddd, d mmmm yyyy → wtorek, 8 wrzesień 2026
//   ddd d/mm → wt 8.09             d.mmmm → 8.wrzesień                  mmmmm → w
//   d "r." mmm → 8 r. wrz          yyyy-mmm-dd → 2026-wrz-08 („-” zostaje)
//   [$-409]d/mmm/yyyy → 8.Sep.2026 (nazwy z języka formatu, separator „/” z systemu)
// Formaty wbudowane (d-mmm-yy, d-mmm, mmm-yy) Excel tłumaczy po swojemu — tu ich nie ruszamy
// (wtedy null i zostaje dotychczasowa ścieżka).
(function (global) {
  "use strict";

  const BUILTIN = new Set(["d-mmm-yy", "d-mmm", "mmm-yy", "m/d/yy", "m/d/yy h:mm", "mm-dd-yy", "dd-mmm-yy"]);
  const LCID_LOCALE = { 0x415: "pl-PL", 0x409: "en-US", 0x809: "en-GB", 0x407: "de-DE", 0x40c: "fr-FR", 0x410: "it-IT", 0xc0a: "es-ES", 0x419: "ru-RU", 0x422: "uk-UA", 0x405: "cs-CZ", 0x41b: "sk-SK" };
  const cache = new Map();

  function namesFor(locale) {
    let n = cache.get(locale);
    if (n) return n;
    const fmt = (opts) => { try { return new Intl.DateTimeFormat(locale, opts); } catch { return new Intl.DateTimeFormat("pl-PL", opts); } };
    const mShort = fmt({ month: "short" }), mLong = fmt({ month: "long" });
    const wShort = fmt({ weekday: "short" }), wLong = fmt({ weekday: "long" });
    const clean = (s) => s.replace(/\.$/, ""); // Intl: „wt.”, Excel: „wt”
    n = {
      monthShort: Array.from({ length: 12 }, (_, i) => clean(mShort.format(new Date(2026, i, 1)))),
      monthLong: Array.from({ length: 12 }, (_, i) => mLong.format(new Date(2026, i, 1))), // mianownik: „wrzesień”
      // 2026-02-01 to niedziela → indeks = getDay()
      dayShort: Array.from({ length: 7 }, (_, i) => clean(wShort.format(new Date(2026, 1, 1 + i)))),
      dayLong: Array.from({ length: 7 }, (_, i) => wLong.format(new Date(2026, 1, 1 + i))),
    };
    cache.set(locale, n);
    return n;
  }

  function dateSeparator(locale) {
    const key = "sep|" + locale;
    if (cache.has(key)) return cache.get(key);
    let sep = "/";
    try {
      const lit = new Intl.DateTimeFormat(locale, { day: "numeric", month: "numeric", year: "numeric" })
        .formatToParts(new Date(2026, 8, 8)).find((p) => p.type === "literal");
      if (lit) sep = lit.value.trim() || sep;
    } catch {}
    cache.set(key, sep);
    return sep;
  }

  // date — Date w czasie lokalnym; code — cell.z; appLocale — język aplikacji (separator „/”).
  // Zwraca tekst albo null, gdy format nie jest prostą datą z nazwą miesiąca/dnia.
  function swbExcelMonthDate(date, code, appLocale) {
    if (!(date instanceof Date) || Number.isNaN(date.getTime()) || typeof code !== "string") return null;
    let section = code.split(";")[0];
    if (BUILTIN.has(section.toLowerCase())) return null;
    let locale = appLocale || "pl-PL";
    const tag = /\[\$[^\]-]*-([0-9a-f]+)\]/i.exec(section);
    if (tag) {
      const id = parseInt(tag[1], 16) & 0xffff;
      if (id >= 0xf000) return null; // [$-F800] itp. = długa data systemu
      locale = LCID_LOCALE[id] || locale;
    }
    section = section.replace(/\[\$[^\]]*\]/g, "");
    const bare = section.replace(/"[^"]*"|\\./g, "");
    if (/\[|[hs0#?@]|am\/pm|a\/p|[eg]/i.test(bare)) return null; // czas, liczby, era — nie ta ścieżka
    if (!/m{3,}|d{3,}/i.test(bare)) return null;
    const names = namesFor(locale);
    const sep = dateSeparator(appLocale || "pl-PL");
    const pad = (v) => String(v).padStart(2, "0");
    let out = "";
    for (let i = 0; i < section.length;) {
      const ch = section[i];
      if (ch === '"') {
        const end = section.indexOf('"', i + 1);
        out += section.slice(i + 1, end < 0 ? section.length : end);
        i = end < 0 ? section.length : end + 1;
        continue;
      }
      if (ch === "\\") { out += section[i + 1] || ""; i += 2; continue; }
      const lower = ch.toLowerCase();
      if (lower === "d" || lower === "m" || lower === "y") {
        let j = i;
        while (j < section.length && section[j].toLowerCase() === lower) j++;
        const n = j - i;
        if (lower === "d") {
          out += n === 1 ? date.getDate() : n === 2 ? pad(date.getDate()) : n === 3 ? names.dayShort[date.getDay()] : names.dayLong[date.getDay()];
        } else if (lower === "m") {
          const mi = date.getMonth();
          out += n === 1 ? mi + 1 : n === 2 ? pad(mi + 1) : n === 3 ? names.monthShort[mi] : n === 4 ? names.monthLong[mi] : names.monthLong[mi].charAt(0);
        } else {
          out += n <= 2 ? pad(date.getFullYear() % 100) : String(date.getFullYear());
        }
        i = j;
        continue;
      }
      out += ch === "/" ? sep : ch;
      i++;
    }
    return out;
  }

  global.swbExcelMonthDate = swbExcelMonthDate;
})(typeof self !== "undefined" ? self : globalThis);
