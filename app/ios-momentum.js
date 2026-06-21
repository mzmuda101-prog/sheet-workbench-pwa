/* ============================================================================
   ios-momentum.js — inercyjne (fling) przewijanie tabeli dla iOS/iPadOS.
   ----------------------------------------------------------------------------
   PO CO: Safari/WebKit daje WEWNĘTRZNYM kontenerom overflow:auto marny rozpęd
   (fling) — na iPadzie tabela „zatrzymuje się pod palcem", brak inercji jak na
   telefonie/Androidzie czy przy przewijaniu całej strony. Zmierzone nagrywarką
   palca (?scrolltest): momentum ~66px na iOS vs ~1079px na telefonie, mimo że
   programowe przewijanie scrollTop na iPadzie jest płynne (60fps).

   ROZWIĄZANIE: na iOS przejmujemy gest na #tableWrap — palec przewija 1:1
   (preventDefault na natywnym), a po puszczeniu animujemy zjazd z decelaracją
   (rAF). Z blokadą osi (czysty pion albo poziom) i wzmocnieniem poziomu (poziome
   machnięcia są krótsze, więc rozpęd byłby za mały).

   ZAKRES: WYŁĄCZNIE iOS/iPadOS. Na desktopie i Androidzie natywne przewijanie
   działa świetnie — tam moduł NIE robi NIC. Pod diagnostyką (?scrolltest) też się
   nie włącza (tam tuninguje się ręcznie przez scroll-diagnostics.js).

   Strojenie na żywo (konsola): window.__iosMomCfg.friction = 0.95; itd.
   ============================================================================ */
(function () {
  "use strict";

  // ── Tylko iOS/iPadOS (iPadOS 13+ podaje się jako "MacIntel" + dotyk) ──
  var ua = navigator.userAgent || "";
  var isIOS = /iPad|iPhone|iPod/.test(ua) ||
    (navigator.platform === "MacIntel" && (navigator.maxTouchPoints || 0) > 1);
  if (!isIOS) return;
  // pod diagnostyką nie włączaj (kolidowałby z ręcznym JS-momentum w teście)
  if (/scrolltest/.test(location.search + location.hash)) return;

  var cfg = {
    friction: 0.93,   // wytracanie prędkości / klatkę (mniejsze = krótszy zjazd)
    xBoost: 1.7,      // poziomo machamy krócej → wzmacniamy rozpęd w poziomie
    maxVpf: 90,       // limit px/klatkę (anti-overshoot)
    axisLock: 1.6     // gdy jedna oś dominuje > tyle× — druga oś zjazdu = 0
  };
  window.__iosMomCfg = cfg;

  var el = null, tracking = false, lastT = 0, lastX = 0, lastY = 0,
    vX = 0, vY = 0, raf = null, moved = false;

  function cancelAnim() { if (raf) { cancelAnimationFrame(raf); raf = null; } }

  function onStart(e) {
    if (e.touches && e.touches.length > 1) { tracking = false; return; } // pinch → natywne
    var t = e.touches && e.touches[0]; if (!t) return;
    cancelAnim();
    tracking = true; moved = false;
    lastT = performance.now(); lastX = t.clientX; lastY = t.clientY; vX = 0; vY = 0;
  }

  function onMove(e) {
    if (!tracking) return;
    if (e.touches && e.touches.length > 1) { tracking = false; return; } // pinch → puść natywne
    var t = e.touches && e.touches[0]; if (!t) return;
    var dx = t.clientX - lastX, dy = t.clientY - lastY;
    if (!moved && Math.abs(dx) < 3 && Math.abs(dy) < 3) return; // próg — pozwól na tap
    moved = true;
    if (e.cancelable) e.preventDefault(); // przejmujemy przewijanie
    var now = performance.now(), dt = (now - lastT) || 16;
    el.scrollLeft -= dx; el.scrollTop -= dy;
    vX = dx / dt; vY = dy / dt; // px/ms
    lastX = t.clientX; lastY = t.clientY; lastT = now;
  }

  function onEnd() {
    if (!tracking) return; tracking = false;
    if (!moved) return;
    var fX = vX * 16 * cfg.xBoost, fY = vY * 16; // px/klatkę
    var ax = Math.abs(fX), ay = Math.abs(fY);
    if (ax > ay * cfg.axisLock) fY = 0; else if (ay > ax * cfg.axisLock) fX = 0;
    var cap = cfg.maxVpf;
    fX = Math.max(-cap, Math.min(cap, fX)); fY = Math.max(-cap, Math.min(cap, fY));
    function step() {
      fX *= cfg.friction; fY *= cfg.friction;
      if (Math.abs(fX) < 0.15 && Math.abs(fY) < 0.15) { raf = null; return; }
      el.scrollLeft -= fX; el.scrollTop -= fY;
      raf = requestAnimationFrame(step);
    }
    raf = requestAnimationFrame(step);
  }

  function attach() {
    el = document.getElementById("tableWrap");
    if (!el) return;
    // #tableWrap jest stałym kontenerem (re-render podmienia tylko <table> w środku),
    // więc listenery przeżywają przerysowania tabeli.
    el.addEventListener("touchstart", onStart, { passive: true });
    el.addEventListener("touchmove", onMove, { passive: false });
    el.addEventListener("touchend", onEnd, { passive: true });
    el.addEventListener("touchcancel", onEnd, { passive: true });
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", attach);
  } else {
    attach();
  }
})();
