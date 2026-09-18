/* ============================================================================
   perf-hud.js — licznik płynności NA URZĄDZENIU (DEBUG ONLY, ładowany na żądanie).

   Po co: zacinanie przewijania tabeli zgłoszone na fizycznym iPhonie nie odtwarza
   się ani w Chromium z dławieniem CPU, ani w WebKicie, ani w symulatorze iOS na
   Macu — bo tam sprzęt jest po prostu za mocny. Jedyne miejsce, gdzie da się to
   zmierzyć, to telefon użytkownika. Ten moduł pokazuje więc liczby wprost na ekranie.

   Włączenie: dopisz ?perfhud do adresu (albo #perfhud, albo localStorage.perfhud=1).
   Wyłączenie: usuń parametr i odśwież (lub stuknij w licznik → chowa się).

   Co pokazuje:
     • FPS i czasy klatek (mediana / 95. percentyl / najgorsza) z ostatnich ~3 s,
     • ZACIĘCIA: ile klatek przekroczyło 32 ms (czyli zgubiona klatka przy 60 Hz),
     • rozmiar tabeli w DOM (wiersze × komórki) — główny podejrzany,
     • profil urządzenia: rdzenie, pamięć, tryb „low-power" apki, PWA standalone,
       wsparcie content-visibility — czyli wszystko, czego nie da się zgadnąć zdalnie.
   ========================================================================== */
(function () {
  "use strict";

  function enabled() {
    try {
      return /(\?|&)perfhud\b/.test(location.search)
        || /perfhud/.test(location.hash)
        || localStorage.getItem("perfhud") === "1";
    } catch (e) {
      return /perfhud/.test(location.search + location.hash);
    }
  }
  if (!enabled() || window.__PERF_HUD__) return;
  window.__PERF_HUD__ = true;

  var box = document.createElement("div");
  box.id = "perfHud";
  box.setAttribute("aria-hidden", "true");
  box.style.cssText = [
    // Górna krawędź, nie dolna: na dole siedzą główne przyciski i pasek Safari,
    // a licznik ma nie zasłaniać tego, co się właśnie testuje.
    "position:fixed", "left:8px", "right:8px",
    "top:calc(4px + env(safe-area-inset-top))",
    "z-index:9999", "padding:8px 10px", "border-radius:10px",
    "background:rgba(12,18,14,.9)", "color:#e8f5ee",
    "font:10.5px/1.4 ui-monospace,SFMono-Regular,Menlo,monospace",
    "white-space:pre-wrap", "pointer-events:auto", "cursor:pointer",
    "box-shadow:0 6px 20px rgba(0,0,0,.35)",
  ].join(";");
  box.addEventListener("click", function () { box.style.display = "none"; });

  function mount() {
    if (document.body) document.body.appendChild(box);
    else document.addEventListener("DOMContentLoaded", mount, { once: true });
  }
  mount();

  // ── zbieranie klatek ───────────────────────────────────────────────────────
  // UWAGA METODYCZNA: requestAnimationFrame mierzy GŁÓWNY WĄTEK (JavaScript).
  // Na iOS przewijaniem zajmuje się osobny wątek kompozytora — gdy to ON się dławi
  // malowaniem kafelków, obraz szarpie, a rAF dalej tyka równiutko 60 razy na sekundę.
  // Użytkownik zgłosił dokładnie to: „licznik pokazuje 60 fps, a i tak się zacina".
  // Dlatego oprócz klatek mierzymy RUCH OBRAZU: próbkujemy pozycję przewijania w każdej
  // klatce i patrzymy, czy przesuwa się równo. Zamrożony obraz pod jadącym palcem to
  // ciąg klatek z zerowym przyrostem pozycji — i to widać w liczbie „zamrożone".
  var frames = [];          // czasy klatek z okna ~3 s
  var jank = 0;             // klatki > 32 ms w tym oknie
  var worstEver = 0;        // najgorsza klatka od startu (nie kasowana)
  var scrolling = false;
  var scrollEndTimer = 0;
  var last = performance.now();

  var wrap = document.getElementById("tableWrap");
  var lastTop = 0, lastLeft = 0;
  var moveDeltas = [];      // przyrosty pozycji przewijania (px/klatkę) w oknie
  var frozenFrames = 0;     // klatki BEZ ruchu obrazu, mimo trwającego przewijania
  var worstFreeze = 0;      // najdłuższa taka seria (w klatkach)
  var freezeRun = 0;

  if (wrap) {
    wrap.addEventListener("scroll", function () {
      scrolling = true;
      clearTimeout(scrollEndTimer);
      scrollEndTimer = setTimeout(function () {
        scrolling = false;
        freezeRun = 0;
      }, 220);
    }, { passive: true });
    lastTop = wrap.scrollTop;
    lastLeft = wrap.scrollLeft;
  }

  function tick(now) {
    var dt = now - last;
    last = now;
    if (dt > 0 && dt < 2000) {
      frames.push(dt);
      if (dt > 32) jank += 1;
      if (dt > worstEver) worstEver = dt;
      if (frames.length > 180) { // ~3 s przy 60 Hz
        var dropped = frames.shift();
        if (dropped > 32) jank = Math.max(0, jank - 1);
      }
    }
    if (wrap && scrolling) {
      var moved = Math.abs(wrap.scrollTop - lastTop) + Math.abs(wrap.scrollLeft - lastLeft);
      lastTop = wrap.scrollTop;
      lastLeft = wrap.scrollLeft;
      moveDeltas.push(moved);
      if (moveDeltas.length > 180) moveDeltas.shift();
      if (moved < 0.5) {
        frozenFrames += 1;
        freezeRun += 1;
        if (freezeRun > worstFreeze) worstFreeze = freezeRun;
      } else {
        freezeRun = 0;
      }
    } else if (wrap) {
      lastTop = wrap.scrollTop;
      lastLeft = wrap.scrollLeft;
    }
    requestAnimationFrame(tick);
  }
  requestAnimationFrame(tick);

  function pct(sorted, q) {
    if (!sorted.length) return 0;
    return sorted[Math.min(sorted.length - 1, Math.floor(sorted.length * q))];
  }

  function deviceLine() {
    var cores = navigator.hardwareConcurrency || "?";
    var mem = navigator.deviceMemory || "?";
    var lowPower = document.documentElement.classList.contains("low-power") ? "TAK" : "nie";
    var standalone = (window.navigator.standalone === true
      || (window.matchMedia && matchMedia("(display-mode: standalone)").matches)) ? "TAK" : "nie";
    var cv = (window.CSS && CSS.supports && CSS.supports("content-visibility", "auto")) ? "tak" : "NIE";
    return "rdzenie " + cores + " · pamięć " + mem + " · low-power " + lowPower
      + "\nPWA " + standalone + " · content-visibility " + cv + " · dpr " + (window.devicePixelRatio || 1);
  }

  function tableLine() {
    var rows = document.querySelectorAll("#dataTable tbody tr").length;
    var cells = document.querySelectorAll("#dataTable tbody td").length;
    var frozen = document.getElementById("tableWrap");
    var flags = [];
    if (frozen && frozen.classList.contains("freeze-first-col")) flags.push("blokada kol.");
    if (frozen && frozen.classList.contains("freeze-headers")) flags.push("blokada nagł.");
    return "tabela " + rows + " wierszy · " + cells + " komórek"
      + (flags.length ? "\n" + flags.join(" · ") : "");
  }

  setInterval(function () {
    if (box.style.display === "none") return;
    var sorted = frames.slice().sort(function (a, b) { return a - b; });
    var p50 = pct(sorted, 0.5);
    var p95 = pct(sorted, 0.95);
    var fps = p50 > 0 ? Math.round(1000 / p50) : 0;
    var ver = (typeof APP_BUILD_VERSION !== "undefined") ? APP_BUILD_VERSION : "?";
    var moves = moveDeltas.slice().sort(function (a, b) { return a - b; });
    var moveMed = pct(moves, 0.5);
    var ruch = moves.length
      ? "ruch obrazu: mediana " + moveMed.toFixed(0) + " px/klatkę · zamrożone "
        + frozenFrames + " kl. (najdłuższa seria " + worstFreeze + ")"
      : "ruch obrazu: — (jeszcze nie przewijano)";
    box.textContent = [
      (scrolling ? "▶ PRZEWIJANIE" : "• spokój") + "   " + fps + " fps (główny wątek)",
      "klatka: " + p50.toFixed(1) + " ms  p95 " + p95.toFixed(1) + "  max " + worstEver.toFixed(0),
      "zacięcia JS (>32 ms): " + jank + " / " + frames.length,
      ruch,
      tableLine(),
      deviceLine(),
      "build " + ver + "  (stuknij, by schować)",
    ].join("\n");
  }, 400);
})();
