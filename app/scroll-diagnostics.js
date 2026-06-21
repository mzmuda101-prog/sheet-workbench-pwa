/* ============================================================================
   scroll-diagnostics.js — TRYB DIAGNOSTYCZNY przewijania tabeli
   ----------------------------------------------------------------------------
   Po co: „shutter" przy przewijaniu tabeli pojawia się TYLKO na iPadzie, a nie
   da się go odtworzyć w Chrome na desktopie. Ten moduł pozwala zmierzyć płynność
   przewijania NA PRAWDZIWYM URZĄDZENIU (iPad/telefon) bez podłączania do Maca —
   wynik pokazuje się w alert() + w panelu Log.

   Aktywacja: dodaj do adresu `?scrolltest` (np. .../index.html?scrolltest)
   albo `#scrolltest`. Pojawi się pływający przycisk „▶ Test scrolla" (lewy dół).
   Można też wywołać z konsoli: window.runScrollDiag().

   Co robi: symuluje serię gestów przewijania w różnych kierunkach, prędkościach
   i profilach gestu (wolny drag / szybki flick / seria flicków / przekątna),
   napędzając realnie scrollLeft/scrollTop kontenera tabeli klatka-po-klatce
   (requestAnimationFrame) i mierząc czas każdej klatki. Z tego liczy FPS,
   medianę/p95/max czasu klatki oraz % „janku" (klatek dłuższych niż ~budżet 60Hz).

   UWAGA metodyczna: napędzany scroll biegnie po głównym wątku (jak na iOS, gdy
   jest jank), więc dobrze pokazuje koszt MALOWANIA nowo odsłanianej treści —
   główny podejrzany „shuttera". Jeśli ten test też szarpie na iPadzie → przyczyną
   jest paint/ilość komórek. Jeśli test jest gładki, a palcem dalej szarpie →
   problem leży po stronie kompozytora/obsługi dotyku.

   Plik jest bezpieczny w produkcji: NIC nie robi, dopóki nie włączysz `scrolltest`.
   ============================================================================ */
(function () {
  "use strict";

  function diagEnabled() {
    try {
      return /(\?|&)scrolltest\b/.test(location.search) ||
             /scrolltest/.test(location.hash) ||
             localStorage.getItem("scrolltest") === "1";
    } catch (e) {
      return /scrolltest/.test(location.search + location.hash);
    }
  }
  if (!diagEnabled()) return;

  // ── Drobne narzędzia ──────────────────────────────────────────────────────
  var raf = window.requestAnimationFrame
    ? window.requestAnimationFrame.bind(window)
    : function (cb) { return setTimeout(function () { cb(performance.now()); }, 16); };
  var sleep = function (ms) { return new Promise(function (r) { setTimeout(r, ms); }); };
  var nextFrame = function () { return new Promise(function (r) { raf(function (t) { r(t); }); }); };

  function easeLinear(t) { return t; }
  function easeOutCubic(t) { return 1 - Math.pow(1 - t, 3); }

  function getTableWrap() {
    return document.getElementById("tableWrap") || document.querySelector(".table-wrap");
  }
  function rowCount() {
    return document.querySelectorAll("#dataTable tbody tr").length;
  }

  // Statystyki z listy odstępów między klatkami (ms).
  function summarize(gaps, maxGap) {
    if (!gaps.length) return null;
    var sorted = gaps.slice().sort(function (a, b) { return a - b; });
    var sum = sorted.reduce(function (a, b) { return a + b; }, 0);
    var pick = function (p) { return sorted[Math.min(sorted.length - 1, Math.floor(sorted.length * p))]; };
    var avg = sum / sorted.length;
    return {
      frames: sorted.length,
      avgMs: +avg.toFixed(1),
      medianMs: +pick(0.5).toFixed(1),
      p95Ms: +pick(0.95).toFixed(1),
      maxMs: +maxGap.toFixed(1),
      fps: +(1000 / avg).toFixed(0),
      // budżet 60Hz ≈ 16.7ms; klatka >24ms = opóźniona, >34ms = w praktyce zgubiona
      jankPct: +((sorted.filter(function (g) { return g > 24; }).length / sorted.length) * 100).toFixed(0),
      dropped: sorted.filter(function (g) { return g > 34; }).length
    };
  }

  // Napędza scroll od (fromX,fromY) do (toX,toY) w czasie duration, mierząc klatki.
  function animate(el, fromX, toX, fromY, toY, duration, ease) {
    return new Promise(function (resolve) {
      var start = performance.now(), last = start, frames = 0, maxGap = 0, gaps = [];
      function frame(now) {
        var gap = now - last; last = now;
        if (frames > 0) { gaps.push(gap); if (gap > maxGap) maxGap = gap; }
        frames++;
        var t = (now - start) / duration; if (t > 1) t = 1;
        var e = ease(t);
        if (toX !== fromX) el.scrollLeft = fromX + (toX - fromX) * e;
        if (toY !== fromY) el.scrollTop = fromY + (toY - fromY) * e;
        if (t < 1) raf(frame);
        else resolve({ gaps: gaps, maxGap: maxGap });
      }
      raf(frame);
    });
  }

  // Łączy kilka odcinków animacji w jeden pomiar (np. seria flicków).
  async function runSegments(el, segments) {
    var allGaps = [], maxGap = 0;
    for (var i = 0; i < segments.length; i++) {
      var s = segments[i];
      var r = await animate(el, s.fromX, s.toX, s.fromY, s.toY, s.duration, s.ease);
      allGaps = allGaps.concat(r.gaps);
      if (r.maxGap > maxGap) maxGap = r.maxGap;
      if (s.pause) await sleep(s.pause);
    }
    return summarize(allGaps, maxGap);
  }

  async function settle(el) {
    el.scrollTop = 0; el.scrollLeft = 0;
    await nextFrame(); await sleep(200); await nextFrame();
  }

  // ── Definicje gestów ──────────────────────────────────────────────────────
  function buildGestures(el) {
    var maxY = Math.max(0, el.scrollHeight - el.clientHeight);
    var maxX = Math.max(0, el.scrollWidth - el.clientWidth);
    var vh = el.clientHeight, vw = el.clientWidth;
    var g = [];

    // PION — wolne, równomierne przeciągnięcie przez całą tabelę
    if (maxY > 4) {
      g.push({ name: "Pion · wolny drag", from: { x: 0, y: 0 }, run: function () {
        return runSegments(el, [{ fromX: 0, toX: 0, fromY: 0, toY: maxY, duration: 2200, ease: easeLinear }]);
      }});
      // PION — szybki flick (duża prędkość, krótki czas, profil wyhamowania)
      g.push({ name: "Pion · szybki flick", from: { x: 0, y: 0 }, run: function () {
        return runSegments(el, [{ fromX: 0, toX: 0, fromY: 0, toY: maxY, duration: 420, ease: easeOutCubic }]);
      }});
      // PION — seria 5 szybkich flicków (jak wielokrotne machnięcia palcem)
      g.push({ name: "Pion · seria flicków", from: { x: 0, y: 0 }, run: function () {
        var step = Math.max(vh * 1.1, maxY / 5), segs = [], y = 0;
        for (var i = 0; i < 5; i++) {
          var ny = Math.min(maxY, y + step);
          segs.push({ fromX: 0, toX: 0, fromY: y, toY: ny, duration: 260, ease: easeOutCubic, pause: 110 });
          y = ny; if (y >= maxY) y = 0; // zawiń, żeby było co przewijać
        }
        return runSegments(el, segs);
      }});
    }

    // POZIOM — tylko jeśli tabela jest szersza niż widok
    if (maxX > 4) {
      g.push({ name: "Poziom · wolny drag", from: { x: 0, y: 0 }, run: function () {
        return runSegments(el, [{ fromX: 0, toX: maxX, fromY: 0, toY: 0, duration: 1800, ease: easeLinear }]);
      }});
      g.push({ name: "Poziom · szybki flick", from: { x: 0, y: 0 }, run: function () {
        return runSegments(el, [{ fromX: 0, toX: maxX, fromY: 0, toY: 0, duration: 420, ease: easeOutCubic }]);
      }});
    }

    // PRZEKĄTNA — jednoczesny ruch w obu osiach (najcięższy dla paintu)
    if (maxY > 4 && maxX > 4) {
      g.push({ name: "Przekątna · flick", from: { x: 0, y: 0 }, run: function () {
        return runSegments(el, [{ fromX: 0, toX: maxX, fromY: 0, toY: maxY, duration: 600, ease: easeOutCubic }]);
      }});
    }

    return g;
  }

  // ── Uruchomienie pełnego testu ────────────────────────────────────────────
  var running = false;
  async function runScrollDiag() {
    if (running) return;
    var el = getTableWrap();
    if (!el) { alert("Brak kontenera tabeli (#tableWrap)."); return; }
    if (rowCount() === 0) { alert("Najpierw wczytaj plik / arkusz, żeby było co przewijać."); return; }

    running = true;
    var btn = document.getElementById("scrollDiagBtn");
    if (btn) { btn.disabled = true; btn.textContent = "⏳ Mierzę…"; }

    try {
      var rows = rowCount();
      var cols = document.querySelectorAll("#dataTable thead .header-row th").length ||
                 document.querySelectorAll("#dataTable thead th").length;
      var dpr = window.devicePixelRatio || 1;
      var cells = rows * cols;

      // Bonus: policz „long tasks" w trakcie (gdzie Safari wspiera PerformanceObserver).
      var longTasks = 0, po = null;
      try {
        po = new PerformanceObserver(function (list) { longTasks += list.getEntries().length; });
        po.observe({ entryTypes: ["longtask"] });
      } catch (e) { po = null; }

      var gestures = buildGestures(el);
      var results = [];
      for (var i = 0; i < gestures.length; i++) {
        await settle(el);
        var stat = await gestures[i].run();
        results.push({ name: gestures[i].name, stat: stat });
      }
      await settle(el);
      if (po) try { po.disconnect(); } catch (e) {}

      // Werdykt: najgorszy gest po max czasie klatki.
      var worst = null;
      results.forEach(function (r) {
        if (r.stat && (!worst || r.stat.maxMs > worst.stat.maxMs)) worst = r;
      });

      // Tekstowy raport
      var head = "SCROLL TEST — " + rows + " wierszy × " + cols + " kol (" + cells + " komórek)\n" +
                 "Widok: " + el.clientWidth + "×" + el.clientHeight + " @" + dpr + "x" +
                 (longTasks ? "  · longtasks: " + longTasks : "") + "\n" +
                 "(fps / mediana / p95 / max ms / jank%)\n";
      var lines = results.map(function (r) {
        var s = r.stat;
        if (!s) return "• " + r.name + ": —";
        return "• " + r.name + ":  " + s.fps + "fps  " +
               s.medianMs + " / " + s.p95Ms + " / " + s.maxMs + "ms  jank " + s.jankPct + "%" +
               (s.dropped ? "  (zgubione " + s.dropped + ")" : "");
      });
      var verdict = worst
        ? "\nNAJGORZEJ: " + worst.name + " — max klatka " + worst.stat.maxMs +
          "ms, jank " + worst.stat.jankPct + "%"
        : "";

      var report = head + lines.join("\n") + verdict;

      // Do panelu Log + konsoli (pełny szczegół), i alert (czytelne na iPadzie).
      if (typeof window.log === "function") {
        window.log("📊 " + head.replace(/\n/g, " | "), "info");
        results.forEach(function (r) { if (r.stat) window.log("• " + r.name + ": " + JSON.stringify(r.stat), "info"); });
      }
      try { console.table(results.map(function (r) { return Object.assign({ gest: r.name }, r.stat || {}); })); } catch (e) {}
      console.log(report);
      alert(report);
    } catch (err) {
      alert("Błąd testu: " + (err && err.message ? err.message : err));
      console.error(err);
    } finally {
      running = false;
      if (btn) { btn.disabled = false; btn.textContent = "▶ Test scrolla"; }
    }
  }
  window.runScrollDiag = runScrollDiag;

  /* ==========================================================================
     NAGRYWARKA PALCA — mierzy REALNE gesty dotyku na #tableWrap.
     Po co: programowy test (wyżej) omija dotyk, a problem „hamuje pod palcem"
     siedzi właśnie w warstwie dotyku. Tu, gdy WŁĄCZYSZ nagrywanie i sam machasz
     palcem po tabeli, dla KAŻDEGO gestu liczymy:
       • ruch palca (px) vs realny scroll (px) → „nadążanie" (followRatio).
         <90% = treść hamuje pod palcem (scroll nie nadąża za palcem).
       • liczbę/tempo zdarzeń touchmove (Hz) → niskie = główny wątek dławi dotyk.
       • klatki w trakcie gestu (mediana/max ms, jank%).
       • momentum po puszczeniu (px w 700ms) → 0 = brak rozpędu/flinga.
       • czy gest był „martwy" (palec ruszył, scroll ~0) oraz defaultPrevented.
     Każdy gest pokazuje się na żywo na nakładce + na końcu pełny raport (alert/Log).
     ========================================================================== */
  var rec = {
    on: false, el: null, gestures: [], g: null,
    fingerDown: false, frameGaps: [], lastFrameT: 0, rafActive: false
  };

  function recHud(line) {
    var hud = document.getElementById("scrollDiagHud");
    if (!hud) return;
    var row = document.createElement("div");
    row.textContent = line;
    row.style.cssText = "padding:2px 0;border-top:1px solid rgba(255,255,255,.12)";
    hud.appendChild(row);
    while (hud.childNodes.length > 7) hud.removeChild(hud.firstChild);
    hud.scrollTop = hud.scrollHeight;
  }

  function recFrameLoop() {
    if (!rec.fingerDown) { rec.rafActive = false; return; }
    rec.rafActive = true;
    raf(function (now) {
      if (rec.lastFrameT) rec.frameGaps.push(now - rec.lastFrameT);
      rec.lastFrameT = now;
      recFrameLoop();
    });
  }

  function onRecStart(e) {
    var t = e.touches && e.touches[0]; if (!t) return;
    rec.frameGaps = []; rec.lastFrameT = 0; rec.fingerDown = true;
    rec.g = {
      t0: performance.now(),
      sx: t.clientX, sy: t.clientY, lx: t.clientX, ly: t.clientY,
      pathX: 0, pathY: 0, moves: 0, maxV: 0, prevent: false,
      startTop: rec.el.scrollTop, startLeft: rec.el.scrollLeft
    };
    if (!rec.rafActive) recFrameLoop();
  }
  function onRecMove(e) {
    if (!rec.g) return;
    var t = e.touches && e.touches[0]; if (!t) return;
    var dx = t.clientX - rec.g.lx, dy = t.clientY - rec.g.ly;
    var now = performance.now();
    var dt = now - (rec.g.lastMoveT || rec.g.t0); rec.g.lastMoveT = now;
    var dist = Math.sqrt(dx * dx + dy * dy);
    if (dt > 0) { var v = dist / dt * 1000; if (v > rec.g.maxV) rec.g.maxV = v; }
    rec.g.pathX += Math.abs(dx); rec.g.pathY += Math.abs(dy);
    rec.g.lx = t.clientX; rec.g.ly = t.clientY; rec.g.moves++;
    if (e.defaultPrevented) rec.g.prevent = true;
  }
  function onRecEnd() {
    if (!rec.g) return;
    rec.fingerDown = false;
    var g = rec.g; rec.g = null;
    var dur = performance.now() - g.t0;
    var relTop = rec.el.scrollTop, relLeft = rec.el.scrollLeft;
    var frames = summarize(rec.frameGaps, rec.frameGaps.reduce(function (a, b) { return Math.max(a, b); }, 0));
    // momentum: ile jeszcze przejedzie w 700ms po puszczeniu
    setTimeout(function () {
      var momTop = rec.el.scrollTop - relTop, momLeft = rec.el.scrollLeft - relLeft;
      var fingerNetY = Math.abs(g.sy - g.ly), fingerNetX = Math.abs(g.sx - g.lx);
      var fingerNet = Math.max(fingerNetX, fingerNetY, Math.sqrt(fingerNetX * fingerNetX + fingerNetY * fingerNetY));
      var scrollNet = Math.sqrt(Math.pow(relTop - g.startTop, 2) + Math.pow(relLeft - g.startLeft, 2));
      var follow = fingerNet > 4 ? Math.round((scrollNet / fingerNet) * 100) : null;
      var momentum = Math.round(Math.sqrt(momTop * momTop + momLeft * momLeft));
      var dir = (g.pathY >= g.pathX ? "pion" : "poziom");
      if (g.pathX > 8 && g.pathY > 8 && Math.min(g.pathX, g.pathY) / Math.max(g.pathX, g.pathY) > 0.4) dir = "przekątna";
      var dead = (fingerNet > 20 && scrollNet < 5);
      var entry = {
        dir: dir,
        palecPx: Math.round(fingerNet),
        scrollPx: Math.round(scrollNet),
        nadazaPct: follow,
        moves: g.moves,
        moveHz: dur > 0 ? Math.round(g.moves / (dur / 1000)) : 0,
        klMedMs: frames ? frames.medianMs : null,
        klMaxMs: frames ? frames.maxMs : null,
        jankPct: frames ? frames.jankPct : null,
        momentumPx: momentum,
        peakVpx_s: Math.round(g.maxV),
        martwy: dead,
        prevented: g.prevent
      };
      rec.gestures.push(entry);
      recHud(
        (dir === "pion" ? "↕" : dir === "poziom" ? "↔" : "✚") +
        " palec " + entry.palecPx + "→scroll " + entry.scrollPx +
        " (" + (follow == null ? "-" : follow + "%") + ")" +
        " · mv " + entry.moves + "@" + entry.moveHz + "Hz" +
        " · kl " + (entry.klMedMs || "-") + "/" + (entry.klMaxMs || "-") + "ms" +
        " · mom " + momentum +
        (dead ? " · MARTWY" : "") + (entry.prevented ? " · PREVENTED" : "")
      );
    }, 720);
  }

  function startFingerRec() {
    var el = getTableWrap();
    if (!el) { alert("Brak #tableWrap."); return; }
    if (rowCount() === 0) { alert("Najpierw wczytaj plik/arkusz."); return; }
    rec.el = el; rec.on = true; rec.gestures = [];
    el.addEventListener("touchstart", onRecStart, { passive: true });
    el.addEventListener("touchmove", onRecMove, { passive: true });
    el.addEventListener("touchend", onRecEnd, { passive: true });
    el.addEventListener("touchcancel", onRecEnd, { passive: true });
    var hud = document.getElementById("scrollDiagHud");
    if (hud) { hud.style.display = "block"; hud.innerHTML = ""; }
    recHud("● NAGRYWAM — machaj palcem po tabeli (różne kierunki/prędkości). Stop = ✋.");
  }
  function stopFingerRec() {
    if (!rec.el) return;
    rec.on = false;
    rec.el.removeEventListener("touchstart", onRecStart);
    rec.el.removeEventListener("touchmove", onRecMove);
    rec.el.removeEventListener("touchend", onRecEnd);
    rec.el.removeEventListener("touchcancel", onRecEnd);
    // poczekaj aż dojdą ostatnie pomiary momentum, potem raport
    setTimeout(function () {
      var gs = rec.gestures.slice();
      if (!gs.length) { alert("Brak nagranych gestów."); return; }
      var avg = function (key) {
        var v = gs.map(function (x) { return x[key]; }).filter(function (x) { return x != null; });
        return v.length ? Math.round(v.reduce(function (a, b) { return a + b; }, 0) / v.length) : null;
      };
      var dead = gs.filter(function (x) { return x.martwy; }).length;
      var prevented = gs.filter(function (x) { return x.prevented; }).length;
      var head = "FINGER REC — " + gs.length + " gestów\n" +
        "ŚREDNIE: nadążanie " + avg("nadazaPct") + "%  · moveHz " + avg("moveHz") +
        "  · klatka med " + avg("klMedMs") + "/max " + avg("klMaxMs") + "ms  · jank " + avg("jankPct") + "%\n" +
        "momentum śr " + avg("momentumPx") + "px  · peakV śr " + avg("peakVpx_s") + "px/s\n" +
        "martwe gesty: " + dead + "/" + gs.length + (prevented ? "  · PREVENTED: " + prevented : "") + "\n" +
        "— gesty —\n";
      var lines = gs.map(function (x, i) {
        return (i + 1) + ". " + x.dir + "  palec " + x.palecPx + "→" + x.scrollPx + "px (" +
          (x.nadazaPct == null ? "-" : x.nadazaPct + "%") + ")  mv " + x.moves + "@" + x.moveHz +
          "Hz  kl " + x.klMedMs + "/" + x.klMaxMs + "ms j" + x.jankPct + "%  mom " + x.momentumPx +
          (x.martwy ? " MARTWY" : "");
      });
      var report = head + lines.join("\n");
      if (typeof window.log === "function") {
        window.log("✋ " + head.replace(/\n/g, " | "), "info");
        gs.forEach(function (x) { window.log("• " + JSON.stringify(x), "info"); });
      }
      try { console.table(gs); } catch (e) {}
      console.log(report);
      alert(report);
    }, 900);
  }
  window.startFingerRec = startFingerRec;
  window.stopFingerRec = stopFingerRec;

  /* ==========================================================================
     PRZEŁĄCZNIK WARIANTÓW CSS — testuje na żywo iOS-owe przełączniki przewijania
     bezpośrednio na #tableWrap, żeby empirycznie znaleźć, który przywraca
     kompozytorowe/inercyjne przewijanie (momentum + niski czas klatki).
     Cyklujesz wariant → nagrywasz FINGER REC → porównujesz „momentum" i „klatka".
     ========================================================================== */
  var cssVariants = [
    { name: "baseline (bez zmian)", apply: function (s) {
      s.webkitOverflowScrolling = ""; s.overscrollBehavior = ""; s.willChange = ""; s.transform = "";
    }},
    { name: "-webkit-overflow-scrolling: touch", apply: function (s) {
      s.webkitOverflowScrolling = "touch"; s.overscrollBehavior = ""; s.willChange = ""; s.transform = "";
    }},
    { name: "overscroll-behavior: auto", apply: function (s) {
      s.webkitOverflowScrolling = ""; s.overscrollBehavior = "auto"; s.willChange = ""; s.transform = "";
    }},
    { name: "will-change: auto (zdjęte)", apply: function (s) {
      s.webkitOverflowScrolling = ""; s.overscrollBehavior = ""; s.willChange = "auto"; s.transform = "";
    }},
    { name: "WSZYSTKO (touch+overscroll auto+willchange auto)", apply: function (s) {
      s.webkitOverflowScrolling = "touch"; s.overscrollBehavior = "auto"; s.willChange = "auto"; s.transform = "";
    }},
    { name: "WSZYSTKO + translateZ(0)", apply: function (s) {
      s.webkitOverflowScrolling = "touch"; s.overscrollBehavior = "auto"; s.willChange = "auto"; s.transform = "translateZ(0)";
    }}
  ];
  var cssVariantIdx = 0;
  function applyCssVariant(idx) {
    var el = getTableWrap();
    if (!el) { alert("Brak #tableWrap."); return null; }
    cssVariantIdx = ((idx % cssVariants.length) + cssVariants.length) % cssVariants.length;
    cssVariants[cssVariantIdx].apply(el.style);
    var name = cssVariants[cssVariantIdx].name;
    recHud("🔧 Wariant CSS: " + name + "  → teraz nagraj FINGER REC i porównaj momentum/klatkę.");
    return name;
  }
  window.applyCssVariant = applyCssVariant;

  /* ==========================================================================
     PROTOTYP: JS-MOMENTUM — własna inercja przewijania (omija słaby fling iOS).
     iOS daje wewnętrznym overflow-scrollerom marny rozpęd. Tu PRZEJMUJEMY gest:
     palec = przewijanie 1:1 (preventDefault na native), puszczenie = animowany
     fling z decelaracją (rAF), aż prędkość spadnie. Programowy scrollTop na iPadzie
     jest płynny (60fps), więc to powinno dać rozpęd jak na telefonie.
     Tap (bez ruchu) NIE jest blokowany → zaznaczanie/podpowiedzi działają.
     Friction regulowalny — większy = dłuższy zjazd. ========================== */
  var jsMom = {
    on: false, el: null, tracking: false, lastT: 0, lastX: 0, lastY: 0,
    vX: 0, vY: 0, raf: null, moved: false, friction: 0.93, maxVpf: 70
  };
  function jsMomCancelAnim() { if (jsMom.raf) { cancelAnimationFrame(jsMom.raf); jsMom.raf = null; } }
  function jsMomStart(e) {
    var t = e.touches && e.touches[0]; if (!t) return;
    jsMomCancelAnim();
    jsMom.tracking = true; jsMom.moved = false;
    jsMom.lastT = performance.now(); jsMom.lastX = t.clientX; jsMom.lastY = t.clientY;
    jsMom.vX = 0; jsMom.vY = 0;
  }
  function jsMomMove(e) {
    if (!jsMom.tracking) return;
    var t = e.touches && e.touches[0]; if (!t) return;
    var dx = t.clientX - jsMom.lastX, dy = t.clientY - jsMom.lastY;
    if (!jsMom.moved && Math.abs(dx) < 3 && Math.abs(dy) < 3) return; // pozwól na tap
    jsMom.moved = true;
    if (e.cancelable) e.preventDefault(); // przejmujemy przewijanie
    var now = performance.now(), dt = (now - jsMom.lastT) || 16;
    jsMom.el.scrollLeft -= dx; jsMom.el.scrollTop -= dy;
    jsMom.vX = dx / dt; jsMom.vY = dy / dt;        // px/ms
    jsMom.lastX = t.clientX; jsMom.lastY = t.clientY; jsMom.lastT = now;
  }
  function jsMomEnd() {
    if (!jsMom.tracking) return; jsMom.tracking = false;
    if (!jsMom.moved) return;
    var vX = jsMom.vX * 16, vY = jsMom.vY * 16; // px/klatkę
    // Blokada osi: jeśli jeden kierunek wyraźnie dominuje, drugi zerujemy —
    // zjazd leci czysto w pionie ALBO poziomie (bez „ukosu", lepszy feeling poziomu).
    var ax = Math.abs(vX), ay = Math.abs(vY);
    if (ax > ay * 1.6) vY = 0; else if (ay > ax * 1.6) vX = 0;
    // Ogranicz prędkość startową (krótka zawartość poziomo nie „wystrzeliwuje").
    var cap = jsMom.maxVpf;
    vX = Math.max(-cap, Math.min(cap, vX)); vY = Math.max(-cap, Math.min(cap, vY));
    var el = jsMom.el, fr = jsMom.friction;
    function step() {
      vX *= fr; vY *= fr;
      if (Math.abs(vX) < 0.15 && Math.abs(vY) < 0.15) { jsMom.raf = null; return; }
      el.scrollLeft -= vX; el.scrollTop -= vY;
      jsMom.raf = requestAnimationFrame(step);
    }
    jsMom.raf = requestAnimationFrame(step);
  }
  function jsMomEnable() {
    var el = getTableWrap(); if (!el) { alert("Brak #tableWrap."); return false; }
    if (jsMom.on) return true;
    jsMom.el = el; jsMom.on = true;
    el.addEventListener("touchstart", jsMomStart, { passive: true });
    el.addEventListener("touchmove", jsMomMove, { passive: false });
    el.addEventListener("touchend", jsMomEnd, { passive: true });
    el.addEventListener("touchcancel", jsMomEnd, { passive: true });
    recHud("🚀 JS-momentum WŁ. (friction " + jsMom.friction + ") — machaj palcem, sprawdź rozpęd.");
    return true;
  }
  function jsMomDisable() {
    if (!jsMom.el) return;
    jsMom.on = false; jsMomCancelAnim();
    jsMom.el.removeEventListener("touchstart", jsMomStart);
    jsMom.el.removeEventListener("touchmove", jsMomMove);
    jsMom.el.removeEventListener("touchend", jsMomEnd);
    jsMom.el.removeEventListener("touchcancel", jsMomEnd);
    recHud("🚀 JS-momentum WYŁ. (powrót do natywnego).");
  }
  window.jsMomEnable = jsMomEnable;
  window.jsMomDisable = jsMomDisable;
  window.jsMomSetFriction = function (f) { jsMom.friction = f; recHud("friction = " + f); };

  // ── Pływające przyciski + nakładka na żywo ─────────────────────────────────
  function mkBtn(id, text, bg, onClick) {
    var b = document.createElement("button");
    b.id = id; b.type = "button"; b.textContent = text;
    b.style.cssText = [
      "padding:10px 14px", "border-radius:999px", "border:1px solid rgba(0,0,0,.15)",
      "background:" + bg, "color:#fff", "font:600 14px/1 system-ui,sans-serif",
      "box-shadow:0 6px 18px rgba(0,0,0,.25)", "cursor:pointer", "-webkit-tap-highlight-color:transparent"
    ].join(";");
    b.addEventListener("click", onClick);
    return b;
  }

  function mountButton() {
    if (document.getElementById("scrollDiagBar")) return;

    var hud = document.createElement("div");
    hud.id = "scrollDiagHud";
    hud.style.cssText = [
      "position:fixed", "left:12px", "bottom:108px", "z-index:99999", "display:none",
      "max-width:min(94vw,560px)", "max-height:40vh", "overflow:auto",
      "padding:8px 10px", "border-radius:12px", "background:rgba(15,25,20,.92)",
      "color:#dfeee7", "font:500 12px/1.35 ui-monospace,Menlo,monospace",
      "box-shadow:0 8px 24px rgba(0,0,0,.35)", "white-space:pre-wrap"
    ].join(";");
    document.body.appendChild(hud);

    var bar = document.createElement("div");
    bar.id = "scrollDiagBar";
    bar.style.cssText = [
      "position:fixed", "left:12px", "bottom:12px", "z-index:99999",
      "display:flex", "gap:8px", "flex-wrap:wrap"
    ].join(";");

    bar.appendChild(mkBtn("scrollDiagBtn", "▶ Auto-test", "#2f6f5c", runScrollDiag));

    var recBtn = mkBtn("scrollDiagRecBtn", "✋ Nagraj palec", "#b4532a", function () {
      if (rec.on) {
        recBtn.textContent = "✋ Nagraj palec"; recBtn.style.background = "#b4532a";
        stopFingerRec();
      } else {
        startFingerRec();
        recBtn.textContent = "⏹ Stop + raport"; recBtn.style.background = "#c0392b";
      }
    });
    bar.appendChild(recBtn);

    var varBtn = mkBtn("scrollDiagVarBtn", "🔧 Wariant: baseline", "#3a4a8c", function () {
      var hud = document.getElementById("scrollDiagHud");
      if (hud) hud.style.display = "block";
      var name = applyCssVariant(cssVariantIdx + 1);
      if (name) varBtn.textContent = "🔧 Wariant: " + name.split(" ")[0];
    });
    bar.appendChild(varBtn);

    var momBtn = mkBtn("scrollDiagMomBtn", "🚀 JS-momentum: WYŁ", "#7a3aa0", function () {
      var hud = document.getElementById("scrollDiagHud");
      if (hud) hud.style.display = "block";
      if (jsMom.on) { jsMomDisable(); momBtn.textContent = "🚀 JS-momentum: WYŁ"; momBtn.style.background = "#7a3aa0"; }
      else if (jsMomEnable()) { momBtn.textContent = "🚀 JS-momentum: WŁ (" + jsMom.friction + ")"; momBtn.style.background = "#9b59b6"; }
    });
    bar.appendChild(momBtn);

    var friBtn = mkBtn("scrollDiagFriBtn", "zjazd −/+", "#555", function () {
      // cykl długości zjazdu: krótszy ↔ dłuższy
      var steps = [0.93, 0.95, 0.965, 0.98, 0.99];
      var i = steps.indexOf(jsMom.friction);
      jsMom.friction = steps[(i + 1) % steps.length];
      window.jsMomSetFriction(jsMom.friction);
      if (jsMom.on) document.getElementById("scrollDiagMomBtn").textContent = "🚀 JS-momentum: WŁ (" + jsMom.friction + ")";
    });
    bar.appendChild(friBtn);

    document.body.appendChild(bar);
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", mountButton);
  } else {
    mountButton();
  }
  if (typeof window.log === "function") {
    window.log("🧪 Tryb diagnostyczny przewijania aktywny (?scrolltest) — przycisk w lewym dolnym rogu.", "info");
  }
})();
