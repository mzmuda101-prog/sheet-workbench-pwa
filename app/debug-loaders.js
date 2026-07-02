// [EN] Lazy loader for scroll-diagnostics / ipad-scroll-debug — external file
// (not inline) so CSP can stay script-src 'self'. Same activation rules as before:
// query param, hash or localStorage flag. No flag = zero extra JS loaded.
(function () {
  function loadScript(src) {
    var s = document.createElement("script");
    s.src = src;
    s.defer = true;
    document.head.appendChild(s);
  }
  var v = "20260629-06";
  try {
    if (/(\?|&)scrolltest\b/.test(location.search) ||
        /scrolltest/.test(location.hash) ||
        localStorage.getItem("scrolltest") === "1") {
      loadScript("app/scroll-diagnostics.js?v=" + v);
    }
    if (/(\?|&)scrolldebug\b/.test(location.search) ||
        /scrolldebug/.test(location.hash) ||
        localStorage.getItem("scrolldebug") === "1") {
      loadScript("app/ipad-scroll-debug.js?v=" + v);
    }
  } catch (e) {
    // [EN] localStorage may throw (private mode) — fall back to URL flags only
    if (/scrolltest/.test(location.search + location.hash)) {
      loadScript("app/scroll-diagnostics.js?v=" + v);
    }
    if (/scrolldebug/.test(location.search + location.hash)) {
      loadScript("app/ipad-scroll-debug.js?v=" + v);
    }
  }
})();
