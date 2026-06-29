/* ============================================================================
   ipad-scroll-debug.js — diagnostyka hamowania scrolla na iPadzie (DEBUG ONLY).
   Aktywacja: ?scrolldebug lub #scrolldebug lub localStorage.scrolldebug=1
   Działa OFFLINE na iPadzie — logi trafiają do panelu Log + localStorage + fetch
   (fetch działa tylko przy podłączonym Web Inspectorze do Maca).
   ============================================================================ */
(function () {
  "use strict";

  function enabled() {
    try {
      return /(\?|&)scrolldebug\b/.test(location.search) ||
        /scrolldebug/.test(location.hash) ||
        localStorage.getItem("scrolldebug") === "1";
    } catch (e) {
      return /scrolldebug/.test(location.search + location.hash);
    }
  }
  if (!enabled()) return;

  window.__IPAD_SCROLL_DEBUG__ = true;
  var SESSION = "5a2667";
  var ENDPOINT = "http://127.0.0.1:7291/ingest/6970f7d9-a205-4b03-ac05-a2517eaa0a5f";
  var LS_KEY = "__ipadScrollDbg_v1";

  function emit(hypothesisId, location, message, data) {
    var payload = {
      sessionId: SESSION,
      hypothesisId: hypothesisId,
      location: location,
      message: message,
      data: data || {},
      timestamp: Date.now(),
      runId: "ipad-pre"
    };
    // #region agent log
    fetch(ENDPOINT, {
      method: "POST",
      headers: { "Content-Type": "application/json", "X-Debug-Session-Id": SESSION },
      body: JSON.stringify(payload)
    }).catch(function () {});
    // #endregion
    try {
      var buf = JSON.parse(localStorage.getItem(LS_KEY) || "[]");
      buf.push(payload);
      if (buf.length > 80) buf = buf.slice(-80);
      localStorage.setItem(LS_KEY, JSON.stringify(buf));
    } catch (e) {}
    if (typeof window.log === "function") {
      window.log("[scrolldbg] " + message + " " + JSON.stringify(data || {}), "info");
    }
  }

  function deviceInfo() {
    return {
      ua: (navigator.userAgent || "").slice(0, 120),
      platform: navigator.platform,
      maxTouchPoints: navigator.maxTouchPoints || 0,
      coarse: !!(window.matchMedia && window.matchMedia("(pointer: coarse)").matches),
      w: window.innerWidth,
      h: window.innerHeight,
      dpr: window.devicePixelRatio || 1
    };
  }

  function attach() {
    var wrap = document.getElementById("tableWrap");
    var tbody = document.querySelector("#dataTable tbody");
    if (!wrap) {
      emit("H0", "ipad-scroll-debug.js:attach", "brak tableWrap", deviceInfo());
      return;
    }

    emit("H0", "ipad-scroll-debug.js:attach", "monitor start", Object.assign(deviceInfo(), {
      isIpad: !!(document.documentElement && document.documentElement.classList.contains("device-ipad")),
      runId: "post-fix"
    }));

    var touching = false;
    var gesture = null;
    var lastScrollT = 0;
    var lastScrollTop = 0;
    var lastScrollLeft = 0;
    var scrollBurst = 0;
    var pointerEnterDuringTouch = 0;

    // H3: pointerenter podczas gestu (iPad potrafi syntetyzować hover)
    if (tbody) {
      tbody.addEventListener("pointerenter", function (e) {
        if (!touching) return;
        pointerEnterDuringTouch++;
        if (pointerEnterDuringTouch <= 3) {
          emit("H3", "ipad-scroll-debug.js:pointerenter", "pointerenter w trakcie touch", {
            n: pointerEnterDuringTouch,
            tag: e.target && e.target.tagName
          });
        }
      }, true);
    }

    wrap.addEventListener("touchstart", function (e) {
      var t = e.touches && e.touches[0];
      if (!t) return;
      touching = true;
      pointerEnterDuringTouch = 0;
      scrollBurst = 0;
      gesture = {
        t0: performance.now(),
        sx: t.clientX, sy: t.clientY,
        startTop: wrap.scrollTop, startLeft: wrap.scrollLeft,
        moves: 0, prevented: false,
        maxScrollV: 0
      };
      emit("H1", "ipad-scroll-debug.js:touchstart", "gest start", {
        scrollTop: wrap.scrollTop,
        scrollLeft: wrap.scrollLeft
      });
    }, { passive: true });

    wrap.addEventListener("touchmove", function (e) {
      if (!gesture) return;
      gesture.moves++;
      if (e.defaultPrevented) gesture.prevented = true;
    }, { passive: true });

    wrap.addEventListener("touchend", function () {
      if (!gesture) { touching = false; return; }
      var g = gesture;
      gesture = null;
      touching = false;
      var relTop = wrap.scrollTop;
      var relLeft = wrap.scrollLeft;
      var fingerNet = Math.sqrt(
        Math.pow(g.sx - (g.lx || g.sx), 2) + Math.pow(g.sy - (g.ly || g.sy), 2)
      );
      var scrollNet = Math.sqrt(
        Math.pow(relTop - g.startTop, 2) + Math.pow(relLeft - g.startLeft, 2)
      );
      var dur = performance.now() - g.t0;
      emit("H1", "ipad-scroll-debug.js:touchend", "gest end (sync)", {
        durMs: Math.round(dur),
        moves: g.moves,
        scrollPx: Math.round(scrollNet),
        prevented: g.prevented,
        pointerEnters: pointerEnterDuringTouch,
        scrollBurst: scrollBurst,
        maxScrollV: Math.round(g.maxScrollV)
      });
      // H1/H5: momentum 700ms po puszczeniu — niski = brak rozpędu (hamowanie)
      setTimeout(function () {
        var mom = Math.round(Math.sqrt(
          Math.pow(wrap.scrollTop - relTop, 2) + Math.pow(wrap.scrollLeft - relLeft, 2)
        ));
        emit("H1", "ipad-scroll-debug.js:momentum", "momentum 700ms", {
          momentumPx: mom,
          scrollPx: Math.round(scrollNet),
          moves: g.moves,
          prevented: g.prevented
        });
      }, 700);
    }, { passive: true });

    wrap.addEventListener("touchmove", function (e) {
      if (!gesture) return;
      var t = e.touches && e.touches[0];
      if (t) { gesture.lx = t.clientX; gesture.ly = t.clientY; }
    }, { passive: true });

    // H2/H4: scroll burst + prędkość scrolla podczas gestu
    wrap.addEventListener("scroll", function () {
      var now = performance.now();
      if (touching) scrollBurst++;
      if (lastScrollT) {
        var dt = now - lastScrollT;
        if (dt > 0 && dt < 200) {
          var dTop = wrap.scrollTop - lastScrollTop;
          var dLeft = wrap.scrollLeft - lastScrollLeft;
          var v = Math.sqrt(dTop * dTop + dLeft * dLeft) / dt;
          if (gesture && v > gesture.maxScrollV) gesture.maxScrollV = v;
          // nagła zmiana prędkości = „hamulec"
          if (touching && dt > 32 && Math.abs(dTop) + Math.abs(dLeft) < 2) {
            emit("H2", "ipad-scroll-debug.js:scroll-stall", "scroll stall during touch", {
              dtMs: Math.round(dt),
              scrollTop: wrap.scrollTop
            });
          }
        }
      }
      lastScrollT = now;
      lastScrollTop = wrap.scrollTop;
      lastScrollLeft = wrap.scrollLeft;
    }, { passive: true });

    // H2: long tasks podczas przewijania
    if (typeof PerformanceObserver !== "undefined") {
      try {
        var po = new PerformanceObserver(function (list) {
          if (!touching) return;
          list.getEntries().forEach(function (entry) {
            if (entry.duration > 50) {
              emit("H2", "ipad-scroll-debug.js:longtask", "long task during touch", {
                durationMs: Math.round(entry.duration),
                name: entry.name
              });
            }
          });
        });
        po.observe({ entryTypes: ["longtask"] });
      } catch (e) {}
    }

    // Przycisk: eksport bufora z localStorage (na iPadzie bez Maca)
    var fab = document.createElement("button");
    fab.type = "button";
    fab.textContent = "📋 Scroll dbg";
    fab.style.cssText = "position:fixed;left:8px;bottom:72px;z-index:99999;padding:8px 10px;font-size:12px;border-radius:8px;border:1px solid #888;background:#1a1a1a;color:#fff;opacity:.9";
    fab.addEventListener("click", function () {
      try {
        var buf = JSON.parse(localStorage.getItem(LS_KEY) || "[]");
        var summary = buf.slice(-15).map(function (x) {
          return x.message + " " + JSON.stringify(x.data);
        }).join("\n");
        alert("Ostatnie logi (" + buf.length + "):\n\n" + summary);
      } catch (err) {
        alert("Brak logów: " + err);
      }
    });
    document.body.appendChild(fab);

    if (typeof window.log === "function") {
      window.log("Scroll debug ON — machaj po tabeli, potem tap 📋 Scroll dbg", "info");
    }
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", attach);
  } else {
    attach();
  }
})();
