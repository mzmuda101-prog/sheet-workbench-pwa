// Cursor-following contextual hints shared by interactive controls.

window.MateuszCursorHint = (() => {
  function createCursorHintController({ cursorHint, prefersReducedMotion = false, getFallbackHint = () => "" }) {
    let cursorHintX = -999;
    let cursorHintY = -999;
    let cursorHintTargetX = -999;
    let cursorHintTargetY = -999;
    let cursorHintFrame = null;
    let cursorHintTimer = null;
    let activeHintEl = null;
    let activePointerType = "";

    function allowsTouchHint(el) {
      const value = String(el.dataset.hintTouch || "").toLowerCase();
      return value === "on" || value === "true" || value === "1" || value === "yes";
    }

    function isDisabled(el, pointerType = "") {
      if (!cursorHint || prefersReducedMotion) return true;
      if (pointerType === "touch") return !allowsTouchHint(el);
      return pointerType !== "mouse" && window.matchMedia("(pointer: coarse)").matches;
    }

    function getCurrentLang() {
      return (document.documentElement.lang || "pl").toLowerCase().startsWith("en") ? "en" : "pl";
    }

    function getHintText(el) {
      const lang = getCurrentLang();
      const langKey = lang === "en" ? "hintEn" : "hintPl";
      return el.dataset[langKey]
        || el.dataset.hint
        || getFallbackHint()
        || "";
    }

    function getHintDelayMs(el, pointerType = "") {
      const rawDelay = el.dataset.hintDelay || "";
      const parsedDelay = parseFloat(rawDelay.replace(",", "."));
      if (Number.isFinite(parsedDelay) && parsedDelay >= 0) return parsedDelay * 1000;
      return pointerType === "touch" ? 650 : 0;
    }

    function moveCursorHint(x, y) {
      if (!cursorHint) return;
      const style = getComputedStyle(cursorHint);
      const parsedOffsetX = parseInt(style.getPropertyValue("--hint-offset-x"), 10);
      const parsedOffsetY = parseInt(style.getPropertyValue("--hint-offset-y"), 10);
      const offsetX = Number.isNaN(parsedOffsetX) ? 22 : parsedOffsetX;
      const offsetY = Number.isNaN(parsedOffsetY) ? 18 : parsedOffsetY;

      cursorHintTargetX = x + offsetX;
      cursorHintTargetY = y - offsetY;
      if (cursorHintFrame !== null) return;

      const animateHint = () => {
        cursorHintX += (cursorHintTargetX - cursorHintX) * 0.24;
        cursorHintY += (cursorHintTargetY - cursorHintY) * 0.24;

        if (Math.abs(cursorHintTargetX - cursorHintX) < 0.2) cursorHintX = cursorHintTargetX;
        if (Math.abs(cursorHintTargetY - cursorHintY) < 0.2) cursorHintY = cursorHintTargetY;

        cursorHint.style.transform = `translate3d(${cursorHintX}px, ${cursorHintY}px, 0)`;

        if (cursorHintX !== cursorHintTargetX || cursorHintY !== cursorHintTargetY) {
          cursorHintFrame = window.requestAnimationFrame(animateHint);
        } else {
          cursorHintFrame = null;
        }
      };

      cursorHintFrame = window.requestAnimationFrame(animateHint);
    }

    function clearCursorHintTimer() {
      if (cursorHintTimer === null) return;
      window.clearTimeout(cursorHintTimer);
      cursorHintTimer = null;
    }

    function setHintText(hintContent) {
      const span = cursorHint && cursorHint.querySelector("span");
      if (!span) return;

      span.textContent = "";
      String(hintContent).split("/|").forEach((line, index) => {
        if (index > 0) span.appendChild(document.createElement("br"));
        span.appendChild(document.createTextNode(line.trim()));
      });
    }

    function showCursorHint(el, x, y) {
      if (isDisabled(el, activePointerType)) return;
      setHintText(getHintText(el));
      cursorHint.className = `cursor-hint is-visible ${el.dataset.hintClass || ""}`.trim();
      cursorHint.style.transformOrigin = "left bottom";
      moveCursorHint(x, y);
    }

    function scheduleCursorHint(el, event) {
      clearCursorHintTimer();
      activeHintEl = el;
      activePointerType = event.pointerType || "mouse";

      const delayMs = getHintDelayMs(el, activePointerType);
      const x = event.clientX;
      const y = event.clientY;

      if (delayMs <= 0) {
        showCursorHint(el, x, y);
        return;
      }

      cursorHintTimer = window.setTimeout(() => {
        cursorHintTimer = null;
        if (activeHintEl !== el) return;
        showCursorHint(el, x, y);
      }, delayMs);
    }

    function hideCursorHint() {
      clearCursorHintTimer();
      activeHintEl = null;
      activePointerType = "";
      if (!cursorHint) return;
      cursorHint.classList.remove("is-visible");
      cursorHintTargetX = -999;
      cursorHintTargetY = -999;
      moveCursorHint(-999, -999);
    }

    function setupCursorHint(elements, clickCallback = null) {
      elements.forEach((el) => {
        if (!el || el.dataset.cursorHintBound === "1") return;
        el.dataset.cursorHintBound = "1";

        el.addEventListener("pointerenter", (event) => {
          event.stopPropagation();
          if (event.pointerType === "touch" || isDisabled(el, event.pointerType || "mouse")) return;
          scheduleCursorHint(el, event);
        });

        el.addEventListener("pointermove", (event) => {
          event.stopPropagation();
          const pointerType = event.pointerType || activePointerType || "mouse";
          if (isDisabled(el, pointerType)) return;
          if (cursorHint && cursorHint.classList.contains("is-visible")) {
            moveCursorHint(event.clientX, event.clientY);
          } else if (activeHintEl === el) {
            scheduleCursorHint(el, event);
          }
        });

        el.addEventListener("pointerdown", (event) => {
          if (event.pointerType !== "touch" || isDisabled(el, "touch")) return;
          event.stopPropagation();
          scheduleCursorHint(el, event);
        });

        el.addEventListener("pointerleave", () => {
          hideCursorHint();
        });

        el.addEventListener("pointerup", hideCursorHint);
        el.addEventListener("pointercancel", hideCursorHint);

        if (clickCallback) {
          el.addEventListener("click", () => clickCallback(el));
        }
      });
    }

    return { setupCursorHint };
  }

  function initCursorHints({
    selector = "[data-hint], [data-hint-pl], [data-hint-en]",
    fallbackHint = "",
  } = {}) {
    const cursorHint = document.getElementById("cursorHint") || document.getElementById("cursor-hint");
    const prefersReducedMotion = window.matchMedia("(prefers-reduced-motion: reduce)").matches;
    const controller = createCursorHintController({
      cursorHint,
      prefersReducedMotion,
      getFallbackHint: () => fallbackHint,
    });
    const bindTargets = () => controller.setupCursorHint(document.querySelectorAll(selector));
    bindTargets();
    if ("MutationObserver" in window) {
      const observer = new MutationObserver(bindTargets);
      observer.observe(document.body, { childList: true, subtree: true });
    }
    return controller;
  }

  return { createCursorHintController, initCursorHints };
})();
