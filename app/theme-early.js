// theme-early.js — motyw ustawiany przed pierwszym malowaniem (anty-FOUC), ładowany w <head>
// bez defer. Klucz „excel-workbench-theme”: "light" | "dark" | brak = Systemowy (śledzi
// prefers-color-scheme; nasłuch na żywo robi ui-controls.js → initTheme()).
(function () {
  var pref = null;
  try { pref = localStorage.getItem("excel-workbench-theme"); } catch (e) { /* tryb prywatny */ }
  var dark = pref === "dark" || (pref !== "light" && !!(window.matchMedia && window.matchMedia("(prefers-color-scheme: dark)").matches));
  document.documentElement.setAttribute("data-theme", dark ? "dark" : "light");
})();
