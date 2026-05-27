// =============================================================================
// cursor-hint.js — podążający za kursorem dymek z podpowiedzią
// =============================================================================
//
// WSZYSTKIE ATRYBUTY HTML (data-*) — ściągawka
// -----------------------------------------------------------------------------
//
// data-hint
//   Główny atrybut — jego obecność na elemencie włącza hint.
//   Trzy warianty użycia:
//
//   a) data-hint="Zapisz plik"
//      Tekst podany wprost. Używaj gdy hint nie musi się tłumaczyć.
//
//   b) data-hint (bez wartości) lub data-hint=""
//      Pusty atrybut = "chcę hint, ale tekst pochodzi skądinąd".
//      Wyświetli fallbackHint ("Kliknij" / "Click") chyba że JS
//      dynamicznie wpisze tekst przez setAttribute("data-hint", "..."),
//      np. jak robi to language.js przez setHint().
//
//   c) data-hint + data-hint-pl + data-hint-en  ← PREFEROWANE dla i18n
//      Samo data-hint jako "znacznik aktywacji", a teksty w osobnych
//      atrybutach językowych (patrz niżej). Wtedy data-hint może być puste.
//
// data-hint-pl="..."
//   Tekst hinta po polsku. Używany gdy lang aplikacji = "pl".
//   Jeśli istnieje, ma pierwszeństwo nad data-hint.
//   Przykład: data-hint-pl="Zmień motyw jasny / ciemny"
//
// data-hint-en="..."
//   Tekst hinta po angielsku. Używany gdy lang aplikacji = "en".
//   Jeśli istnieje, ma pierwszeństwo nad data-hint.
//   Przykład: data-hint-en="Switch light / dark theme"
//
//   UWAGA: data-hint-pl i data-hint-en działają niezależnie —
//   możesz podać tylko jeden z nich, drugi nie jest wymagany.
//   Jeśli brakuje wersji dla aktualnego języka, system spada
//   na data-hint, a potem na fallback.
//
// data-hint-delay="1.1"
//   Opóźnienie pojawienia się hinta przy kursorem myszy, w sekundach.
//   Timer startuje gdy mysz wejdzie na element i resetuje się gdy
//   mysz ruszy się o więcej niż ~4px (MOVE_THRESHOLD).
//   Hint pojawia się dopiero po zatrzymaniu myszy na zadany czas.
//   Wartość: liczba dziesiętna, akceptuje przecinek i kropkę.
//   Domyślnie: 0 (pojawia się od razu).
//   Przykład: data-hint-delay="1.1"  ← pojawi się po 1.1 sekundy bezruchu
//
// data-hint-touchdelay="0.65"
//   Opóźnienie dla dotyku (przytrzymanie palca), w sekundach.
//   NIEZALEŻNE od data-hint-delay — dotyk i mysz mają osobne timery.
//   Wymaga data-hint-touch="on" żeby hint w ogóle działał na dotyk.
//   Domyślnie: 0.65 (650ms).
//   Przykład: data-hint-touchdelay="0.8"
//
// data-hint-touch="on"
//   Włącza hint dla urządzeń dotykowych. Domyślnie hint działa tylko
//   dla myszy — dodaj ten atrybut jeśli chcesz obsługę dotyku.
//   Akceptowane wartości: "on", "true", "1", "yes".
//   Przykład: data-hint-touch="on"
//
// data-hint-class="nazwa-klasy"
//   Dodatkowa klasa CSS doklejana do dymka gdy jest widoczny.
//   Przydatne do stylowania konkretnych hintów inaczej niż reszta,
//   np. inny kolor, rozmiar, wariant. Klasa jest usuwana gdy hint znika.
//   Przykład: data-hint-class="smaller"  ← jest już taka klasa w app.css
//
// -----------------------------------------------------------------------------
// SEPARATOR LINII W TEKŚCIE HINTA
// -----------------------------------------------------------------------------
//
//   /|  (ukośnik + pipe) — ręczny podział na nową linię w treści hinta.
//   Możesz go użyć w każdym atrybucie tekstowym (data-hint, data-hint-pl itd.).
//   Przykład: data-hint="Filtruj dane /| i sortuj wyniki"
//             wyświetli się jako dwie linie:
//               Filtruj dane
//               i sortuj wyniki
//
// -----------------------------------------------------------------------------
// KOMPLETNY PRZYKŁAD — element z pełnym zestawem atrybutów
// -----------------------------------------------------------------------------
//
//   <button
//     data-hint
//     data-hint-pl="Pobierz kopię pliku /| jako XLSX"
//     data-hint-en="Download a copy /| as XLSX"
//     data-hint-delay="1.1"
//     data-hint-touch="on"
//     data-hint-touchdelay="0.8"
//     data-hint-class="smaller"
//   >
//     Zapisz
//   </button>
//
// -----------------------------------------------------------------------------
// CSS — customizacja offsetu dymka przez zmienne
// -----------------------------------------------------------------------------
//
//   --hint-offset-x  (domyślnie 22px) — poziome przesunięcie od kursora
//   --hint-offset-y  (domyślnie 18px) — pionowe przesunięcie od kursora
//
//   Ustaw na elemencie #cursorHint lub przez klasę z data-hint-class:
//   .cursor-hint.smaller { --hint-offset-x: 18; --hint-offset-y: 22; }
//
// =============================================================================

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

      // data-hint-pl / data-hint-en — zawsze mają pierwszeństwo jeśli istnieją
      if (el.dataset[langKey] !== undefined) return el.dataset[langKey];

      // data-hint="tekst" — użyj tekstu; data-hint="" lub samo data-hint — fallback
      if (el.dataset.hint !== undefined && el.dataset.hint !== "") return el.dataset.hint;

      // brak wartości lub pusty atrybut → domyślny hint (np. "Kliknij")
      return getFallbackHint() || "";
    }

    function getHintDelayMs(el, pointerType = "") {
      if (pointerType === "touch") {
        // Dla dotyku używamy data-hint-touchdelay (osobny atrybut),
        // data-hint-delay na dotyk nie ma wpływu.
        const rawTouch = el.dataset.hintTouchdelay || "";
        const parsedTouch = parseFloat(rawTouch.replace(",", "."));
        if (Number.isFinite(parsedTouch) && parsedTouch >= 0) return parsedTouch * 1000;
        return 650; // domyślny delay dotyku
      }
      const rawDelay = el.dataset.hintDelay || "";
      const parsedDelay = parseFloat(rawDelay.replace(",", "."));
      if (Number.isFinite(parsedDelay) && parsedDelay >= 0) return parsedDelay * 1000;
      return 0;
    }

    // Oblicza docelową pozycję hinta z uwzględnieniem granic ekranu.
    // Hint jest najpierw renderowany (niewidoczny) żeby znać jego rozmiar,
    // potem pozycjonowany tak by nie wychodził poza viewport.
    function computeHintPosition(x, y) {
      if (!cursorHint) return { tx: x + 22, ty: y - 18, originX: "left", originY: "bottom" };

      const style = getComputedStyle(cursorHint);
      const parsedOffsetX = parseInt(style.getPropertyValue("--hint-offset-x"), 10);
      const parsedOffsetY = parseInt(style.getPropertyValue("--hint-offset-y"), 10);
      const offsetX = Number.isNaN(parsedOffsetX) ? 22 : parsedOffsetX;
      const offsetY = Number.isNaN(parsedOffsetY) ? 18 : parsedOffsetY;

      const W = window.innerWidth;
      const H = window.innerHeight;
      const MARGIN = 8; // minimalna odległość od krawędzi ekranu

      const rect = cursorHint.getBoundingClientRect();
      const hintW = rect.width || 0;
      const hintH = rect.height || 0;

      // Domyślna pozycja: prawo-góra względem kursora
      let tx = x + offsetX;
      let ty = y - offsetY - hintH;
      let originX = "left";
      let originY = "bottom";

      // Wychodzi poza prawą krawędź → przesuń w lewo od kursora
      if (tx + hintW + MARGIN > W) {
        tx = x - offsetX - hintW;
        originX = "right";
      }

      // Wychodzi poza górną krawędź → pokaż pod kursorem
      if (ty < MARGIN) {
        ty = y + offsetY;
        originY = "top";
      }

      // Zabezpieczenie przed lewą krawędzią
      if (tx < MARGIN) tx = MARGIN;

      // Zabezpieczenie przed dolną krawędzią
      if (ty + hintH + MARGIN > H) ty = H - hintH - MARGIN;

      return { tx, ty, originX, originY };
    }

    function moveCursorHint(x, y) {
      if (!cursorHint) return;

      const { tx, ty, originX, originY } = computeHintPosition(x, y);

      cursorHint.style.transformOrigin = `${originX} ${originY}`;
      cursorHintTargetX = tx;
      cursorHintTargetY = ty;

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
      // transformOrigin zostanie ustawiony przez moveCursorHint po obliczeniu pozycji
      moveCursorHint(x, y);
    }

    let pendingHintX = 0;
    let pendingHintY = 0;
    let lastMoveX = 0;
    let lastMoveY = 0;
    const MOVE_THRESHOLD = 4; // px — ruch poniżej tego progu nie resetuje timera

    function scheduleCursorHint(el, event) {
      clearCursorHintTimer();
      activeHintEl = el;
      activePointerType = event.pointerType || "mouse";
      pendingHintX = event.clientX;
      pendingHintY = event.clientY;

      const delayMs = getHintDelayMs(el, activePointerType);

      if (delayMs <= 0) {
        showCursorHint(el, pendingHintX, pendingHintY);
        return;
      }

      cursorHintTimer = window.setTimeout(() => {
        cursorHintTimer = null;
        if (activeHintEl !== el) return;
        showCursorHint(el, pendingHintX, pendingHintY);
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
          lastMoveX = event.clientX;
          lastMoveY = event.clientY;
          scheduleCursorHint(el, event);
        });

        el.addEventListener("pointermove", (event) => {
          event.stopPropagation();
          const pointerType = event.pointerType || activePointerType || "mouse";
          if (isDisabled(el, pointerType)) return;
          if (cursorHint && cursorHint.classList.contains("is-visible")) {
            moveCursorHint(event.clientX, event.clientY);
          } else if (activeHintEl === el) {
            const dx = event.clientX - lastMoveX;
            const dy = event.clientY - lastMoveY;
            if (Math.sqrt(dx * dx + dy * dy) >= MOVE_THRESHOLD) {
              // Mysz rzeczywiście się ruszyła — resetuj timer
              lastMoveX = event.clientX;
              lastMoveY = event.clientY;
              scheduleCursorHint(el, event);
            }
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