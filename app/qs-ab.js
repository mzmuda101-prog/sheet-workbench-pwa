/* ============================================================================
   qs-ab.js — A/B na telefonie: co w „Trybie szybkie szukanie" zacina przewijanie
   (DEBUG ONLY, ładowany przez debug-loaders.js przy ?qsab).

   Po co (2026-09-24): na fizycznym iPhonie przewijanie tabeli laguje, gdy widać
   pasek szybkiego szukania. Nie odtwarza się w Chromium/WebKicie/symulatorze.
   Ustalone na telefonie: pasek zwinięty strzałką = brak laga; mniejszy pasek
   (więcej miejsca na tabelę) = lag dalej; nic nie wpisane w pole = lag dalej.
   Zwinięcie chowało też aktywny przycisk trybu (z cieniem), więc podejrzanych jest
   kilku. Panel pozwala wyłączać ich po kolei bez przeładowania strony.
   ========================================================================== */
(function () {
  "use strict";
  if (window.__QS_AB__) return;
  window.__QS_AB__ = true;

  var MODES = [
    ["nobar", "bez paska", "#quickSearchWrap{display:none!important}"],
    ["ghost", "pasek niewidoczny", "#quickSearchWrap{opacity:0!important}"],
    ["noshadow", "bez cienia przycisku", ".toggle-pill[aria-pressed=\"true\"],.qs-scope-toggle[aria-pressed=\"true\"]{box-shadow:none!important}"],
    ["noinput", "bez pola", "#quickSearch{visibility:hidden!important}"],
    ["noflags", "bez ikonek", "#qsFlags{visibility:hidden!important}"],
    ["nobtns", "bez przycisków", "#quickSearchBtn,#qsMoreBtn{visibility:hidden!important}"],
    ["noframe", "bez ramki", "#quickSearchWrap{border-color:transparent!important;background:transparent!important}"],
  ];
  var root = document.documentElement;
  var css = MODES.map(function (m) {
    return m[2].replace(/(^|\})([^{]+)\{/g, function (_, brace, sel) {
      return brace + sel.split(",").map(function (s) { return "html.qsab-" + m[0] + " " + s.trim(); }).join(",") + "{";
    });
  }).join("\n");

  var saved = [];
  try { saved = JSON.parse(localStorage.getItem("qsabModes") || "[]"); } catch (e) {}

  function mount() {
    var style = document.createElement("style");
    style.textContent = css;
    document.head.appendChild(style);

    // Panel celowo „płaski": bez cienia, rozmycia i animacji — sam nie może dokładać
    // pracy kompozytorowi, bo zafałszowałby wynik.
    var box = document.createElement("div");
    box.id = "qsAbPanel";
    box.style.cssText = "position:fixed;left:8px;right:8px;bottom:calc(8px + env(safe-area-inset-bottom));z-index:9998;"
      + "display:flex;flex-wrap:wrap;gap:4px;padding:6px;border-radius:10px;background:#0c120e;color:#e8f5ee;"
      + "font:11px/1.2 ui-monospace,Menlo,monospace";
    MODES.forEach(function (m) {
      var b = document.createElement("button");
      b.type = "button";
      b.textContent = m[1];
      var on = saved.indexOf(m[0]) >= 0;
      var paint = function () {
        b.style.cssText = "padding:7px 8px;border-radius:7px;border:1px solid #3a5a46;font:inherit;"
          + (on ? "background:#e8a33d;color:#111;border-color:#e8a33d" : "background:transparent;color:#e8f5ee");
      };
      root.classList.toggle("qsab-" + m[0], on);
      paint();
      b.addEventListener("click", function () {
        on = !on;
        root.classList.toggle("qsab-" + m[0], on);
        paint();
        var list = MODES.filter(function (x) { return root.classList.contains("qsab-" + x[0]); }).map(function (x) { return x[0]; });
        try { localStorage.setItem("qsabModes", JSON.stringify(list)); } catch (e) {}
      });
      box.appendChild(b);
    });
    var hide = document.createElement("button");
    hide.type = "button";
    hide.textContent = "×";
    hide.setAttribute("aria-label", "Schowaj panel");
    hide.style.cssText = "margin-left:auto;padding:7px 10px;border-radius:7px;border:1px solid #3a5a46;background:transparent;color:#e8f5ee;font:inherit";
    hide.addEventListener("click", function () { box.style.display = "none"; });
    box.appendChild(hide);
    document.body.appendChild(box);
  }
  if (document.body) mount();
  else document.addEventListener("DOMContentLoaded", mount, { once: true });
})();
