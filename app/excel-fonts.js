// Czcionki z pliku → stos CSS, jak rysuje je Excel. Ładowany w wątku głównym (index.html)
// i w wątku roboczym (build-rows-worker.js) — jedna wersja zamiast dwóch kopii.
//
// Wyrocznia Excela (Mac, 2026-11-02, czcionki „Czcionka tekstu podstawowego” i „Zupelnie Nieznana XYZ”):
// czcionki, której nie ma w systemie, Excel NIE zastępuje jednym krojem — zwykły tekst rysuje
// szeryfowym (Times), a pogrubiony, kursywę i pogrubioną kursywę bezszeryfowym (Helvetica).
// Gdy czcionka jest zainstalowana, przeglądarka weźmie ją z pierwszego miejsca stosu (jak Excel).
(function (global) {
  "use strict";

  // Znane kroje Office → najbliższe zamienniki na Macu/iOS (gdy oryginału nie ma).
  const FALLBACKS = {
    "calibri": "'Segoe UI', Candara, Optima, sans-serif",
    "calibri light": "'Segoe UI Light', 'Segoe UI', sans-serif",
    "aptos": "'Segoe UI', sans-serif",
    "aptos narrow": "'Segoe UI', sans-serif",
    "arial": "Helvetica, 'Helvetica Neue', sans-serif",
    "arial narrow": "'Arial Narrow', Arial, sans-serif",
    "arial black": "Arial, sans-serif",
    "helvetica": "'Helvetica Neue', Arial, sans-serif",
    "tahoma": "Geneva, Verdana, sans-serif",
    "verdana": "Geneva, sans-serif",
    "segoe ui": "Candara, sans-serif",
    "trebuchet ms": "'Lucida Grande', sans-serif",
    "century gothic": "'Apple Gothic', sans-serif",
    "times new roman": "Times, serif",
    "times": "serif",
    "cambria": "Georgia, serif",
    "georgia": "'Times New Roman', serif",
    "garamond": "'Apple Garamond', Georgia, serif",
    "book antiqua": "Palatino, 'Palatino Linotype', serif",
    "palatino linotype": "Palatino, serif",
    "courier new": "Courier, monospace",
    "consolas": "'Lucida Console', Monaco, monospace",
    "lucida console": "Monaco, monospace",
    "comic sans ms": "'Comic Sans', cursive",
  };
  const UNKNOWN_REGULAR = "'Times New Roman', Times, serif";
  const UNKNOWN_STYLED = "Helvetica, Arial, sans-serif";

  function swbIsKnownExcelFont(name) {
    return Object.prototype.hasOwnProperty.call(FALLBACKS, String(name || "").trim().toLowerCase());
  }

  function swbExcelFontStack(name, bold, italic) {
    const n = String(name || "").trim();
    if (!n) return "";
    const fallback = FALLBACKS[n.toLowerCase()] || (bold || italic ? UNKNOWN_STYLED : UNKNOWN_REGULAR);
    return `"${n.replace(/"/g, "")}", ${fallback}`;
  }

  global.swbIsKnownExcelFont = swbIsKnownExcelFont;
  global.swbExcelFontStack = swbExcelFontStack;
})(typeof self !== "undefined" ? self : globalThis);
