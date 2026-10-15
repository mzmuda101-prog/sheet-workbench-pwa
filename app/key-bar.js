// key-bar.js — pasek cyfr i operatorów nad klawiaturą telefonu.
//
// Po co (prośba Mateusza 2026-10-15): klawiatura iPhone'a nie ma ani rzędu cyfr, ani
// znaków potrzebnych w formułach (; : ( ) =) i w szukaniu (>> << && || !) — każdy
// wymaga przełączania klawiatury, czasem dwa razy. Tablety mają cyfry pod ręką, więc
// pasek jest TYLKO na telefonie (dotyk + wąski albo niski ekran) i tylko wtedy, gdy
// klawiatura ekranowa jest naprawdę otwarta.
//
// Zestaw zależy od pola:
//   • komórka w edycji / pasek formuły — cyfry, potem , - = ( ) ; : + * / %
//   • szukanie i filtry — operatory wyszukiwania, potem cyfry; pierwszy operator włącza
//     „Operatory wyszukiwania” dla tego pola (bez tego wpisane >> nic by nie znaczyło).
//
// Jak: przycisk wstawia znak w miejscu kursora i NIE zabiera fokusu z pola (preventDefault
// na mousedown) — klawiatura zostaje, edytor komórki nie zatwierdza się blur-em.
// Pozycja: tuż nad klawiaturą, liczona z visualViewport (GOTCHA iOS: position:fixed żyje
// w layout viewport, więc do top dodajemy visualViewport.offsetTop).

(function () {
  const vv = window.visualViewport;
  const phoneMq = window.matchMedia ? window.matchMedia("(pointer: coarse) and (max-width: 600px), (pointer: coarse) and (max-height: 500px)") : null;
  const CELL_KEYS = ["1", "2", "3", "4", "5", "6", "7", "8", "9", "0", ",", "-", "=", "(", ")", ";", ":", "+", "*", "/", "%"];
  const SEARCH_KEYS = [">>", "<<", ">>=", "=<<", "&&", "||", "!", "\"", ":", "=", "1", "2", "3", "4", "5", "6", "7", "8", "9", "0", ",", "-"];
  const SEARCH_OPS = new Set([">>", "<<", ">>=", "=<<", "&&", "||", "!", "\"", ":", "="]);
  // pole szukania → jego przełącznik „Operatory wyszukiwania”
  const SEARCH_FIELDS = {
    quickSearch: "quickSearchOperators",
    quickSearchPopupInput: "quickSearchPopupOperators",
    searchQuery: "filterOperators",
    searchQuery2: "filterOperators2",
  };

  let bar = null;
  let row = null;
  let field = null;      // pole, do którego piszemy
  let mode = "";         // "cell" | "search"
  let watchFrame = 0;
  let lastTop = NaN;

  function fieldMode(el) {
    if (!el || el.tagName !== "INPUT" || el.readOnly || el.disabled) return "";
    if (el.classList.contains("cell-editor") || el.id === "formulaInput") return "cell";
    if (SEARCH_FIELDS[el.id]) return "search";
    return "";
  }

  function keyboardOpen() {
    if (window.__swbKeyBarForce) return true; // testy (w przeglądarce bez okna nie ma klawiatury)
    if (!vv) return false;
    return window.innerHeight - vv.height > 120;
  }

  function build() {
    bar = document.createElement("div");
    bar.className = "key-bar";
    bar.hidden = true;
    bar.setAttribute("role", "toolbar");
    row = document.createElement("div");
    row.className = "key-bar-row";
    bar.appendChild(row);
    document.body.appendChild(bar);
    // Fokus zostaje w polu: Safari (też iOS) i Chrome przenoszą fokus przy mousedown —
    // blokujemy tylko to. NIE touchstart: tamto zablokowałoby przewijanie paska palcem.
    bar.addEventListener("mousedown", (e) => e.preventDefault());
    bar.addEventListener("click", (e) => {
      const key = e.target.closest(".key-bar-key");
      if (key) insert(key.dataset.key);
    });
    if (typeof attachOverflowFade === "function") attachOverflowFade(row);
  }

  function fill(nextMode) {
    if (nextMode === mode && row.childElementCount) return;
    mode = nextMode;
    bar.dataset.mode = mode;
    bar.setAttribute("aria-label", typeof t === "function" ? t(mode === "search" ? "keyBarSearchAria" : "keyBarCellAria") : "");
    row.replaceChildren(...(mode === "search" ? SEARCH_KEYS : CELL_KEYS).map((k) => {
      const b = document.createElement("button");
      b.type = "button";
      b.className = "key-bar-key" + (/^\d$/.test(k) ? " is-digit" : " is-op");
      b.dataset.key = k;
      b.textContent = k;
      b.tabIndex = -1; // klawiatura sprzętowa i tak tego nie potrzebuje
      return b;
    }));
    row.scrollLeft = 0;
  }

  function insert(text) {
    const el = field;
    if (!el || !el.isConnected || document.activeElement !== el) return;
    if (mode === "search" && SEARCH_OPS.has(text)) enableOperators(el);
    let start = el.value.length;
    let end = start;
    try { start = el.selectionStart ?? start; end = el.selectionEnd ?? end; } catch (_) { /* typ bez zaznaczenia */ }
    el.setRangeText(text, start, end, "end");
    el.dispatchEvent(new Event("input", { bubbles: true }));
  }

  function enableOperators(el) {
    const box = document.getElementById(SEARCH_FIELDS[el.id] || "");
    if (!box || box.checked) return;
    box.checked = true;
    box.dispatchEvent(new Event("change", { bubbles: true }));
    if (typeof toast === "function") toast(t("keyBarOpsOn"), "info");
  }

  // Pozycja nad klawiaturą. Klawiatura wjeżdża animacją, a iOS nie zawsze daje przy tym
  // zdarzenie — póki pasek jest widoczny, pilnujemy pozycji w rAF (tani odczyt vv).
  function place() {
    if (!bar || bar.hidden) return;
    const h = bar.offsetHeight || 48;
    const top = vv ? vv.offsetTop + vv.height - h : window.innerHeight - h;
    if (top !== lastTop) {
      lastTop = top;
      bar.style.transform = `translate3d(0, ${Math.round(top)}px, 0)`;
    }
  }
  function watch() {
    watchFrame = 0;
    if (!bar || bar.hidden) return;
    if (!keyboardOpen() || !field || document.activeElement !== field) { hide(); return; }
    place();
    watchFrame = requestAnimationFrame(watch);
  }

  function show(el, m) {
    if (!bar) build();
    field = el;
    fill(m);
    if (!keyboardOpen()) { hide(); return; }
    if (bar.hidden) { bar.hidden = false; lastTop = NaN; }
    place();
    if (!watchFrame) watchFrame = requestAnimationFrame(watch);
  }
  function hide() {
    if (watchFrame) { cancelAnimationFrame(watchFrame); watchFrame = 0; }
    if (bar) bar.hidden = true;
  }

  function sync() {
    const el = document.activeElement;
    const m = phoneMq && phoneMq.matches ? fieldMode(el) : "";
    if (m) show(el, m); else { field = null; hide(); }
  }

  document.addEventListener("focusin", () => setTimeout(sync, 0));
  document.addEventListener("focusout", () => setTimeout(sync, 0));
  // klawiatura wjeżdża PO focusin — dopiero wtedy wiadomo, że jest otwarta
  if (vv) vv.addEventListener("resize", () => { if (field || fieldMode(document.activeElement)) sync(); });

  window.swbKeyBar = { sync, get visible() { return !!(bar && !bar.hidden); }, get mode() { return mode; } };
})();
