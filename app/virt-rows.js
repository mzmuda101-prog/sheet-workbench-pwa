// virt-rows.js — „Nowy silnik tabeli”: wirtualizacja wierszy.
//
// Dlaczego: koszt każdej akcji (sort, filtr, przebudowa, przewijanie) rósł liniowo
// z liczbą wierszy w DOM — przy 5000 wierszach sort ~3 s na średnim telefonie
// (docs/notes/virt-baseline-cpu4.json). Excel/Google Sheets rysują tylko to, co widać.
// Tu robimy to samo: w <tbody> żyje okno widocznych wierszy + zapas (overscan)
// po obu stronach, a resztę wysokości udają dwa wiersze-wypełniacze (.vr-spacer).
// Przewijanie zostaje NATYWNE (ta sama fizyka iOS), zmienia się tylko zawartość.
//
// Zasady (każda z czegoś wynika):
// - Wiersze buduje ten sam buildTableRow() co stary render → identyczny wygląd.
// - Przejmujemy render tylko gdy: przełącznik włączony, wierszy > VIRT_MIN_ROWS,
//   brak scaleń (rowspan przecięty krawędzią okna) i brak zawijania tekstu (wysokości
//   nieprzewidywalne) — wtedy rysujemy wszystko po staremu, z limitem wierszy.
// - Wysokości: znane z góry (ręczne / z Excela) albo zmierzone po narysowaniu
//   (pamiętane po kluczu wiersza), reszta = typowa wysokość zmierzona na starcie.
// - KOTWICA: gdy dokładamy wiersze NAD widokiem, a ich prawdziwa wysokość różni się
//   od szacunku, górny wypełniacz koryguje się tak, by widoczny wiersz nie drgnął.
//   Różnicę (drift) wyrównujemy dopiero w spoczynku, razem ze scrollTop — w tej samej
//   klatce, więc niewidocznie. W trakcie gestu NIE ruszamy scrollTop (na iOS zabija
//   to rozpęd).
// - Wypełniacze mają w tle linie siatki: przy bardzo szybkim machnięciu, zanim
//   dorysujemy wiersze, widać pustą siatkę (jak w Excelu), a nie białą dziurę.
//
// Przełącznik: Widok → „Nowy silnik tabeli (beta)” (zapamiętany na urządzeniu).
// `?virt=1` / `?virt=off` w adresie wymusza na tę sesję (testy, A/B).

const VIRT_PREF_KEY = "swb-virt-rows-v1";
const VIRT_MIN_ROWS = 200;       // do tylu wierszy stary render i tak jest tani
const VIRT_OVERSCAN = 1.0;       // zapas po każdej stronie, w wysokościach okna
const VIRT_MARGIN = 0.8;         // dorysuj, gdy do krawędzi okna zostało mniej niż tyle ekranów
const VIRT_POOL_MAX = 600;       // odłączone wiersze trzymane do ponownego użycia
const VIRT_IDLE_MS = 220;        // spoczynek po przewijaniu → wyrównanie driftu

const swbVirt = (() => {
  const urlForced = (() => {
    try {
      const v = new URLSearchParams(location.search).get("virt");
      if (v === "1" || v === "on") return true;
      if (v === "0" || v === "off") return false;
    } catch (_) {}
    return null;
  })();
  let enabledPref = (() => { try { return localStorage.getItem(VIRT_PREF_KEY) === "1"; } catch (_) { return false; } })();

  let st = null;           // stan zamontowanego okna (null = stary render)
  const heightCache = new Map(); // rowKey -> { h, f } (f = wysokość wymuszona przy pomiarze)
  let heightEnv = "";
  let defaultH = 28;       // typowa wysokość wiersza w tym środowisku
  let defaultMeasured = false;
  let touchActive = false;
  let lastScrollAt = 0;
  let idleTimer = 0;
  // W ruchu (palec na ekranie albo rozpęd po jego zdjęciu) NIE piszemy scrollTop —
  // na iOS każdy zapis zatrzymuje rozpęd. Rozpęd = zdarzenia scroll wciąż napływają.
  const inMotion = () => touchActive || performance.now() - lastScrollAt < VIRT_IDLE_MS;

  function enabled() { return urlForced !== null ? urlForced : enabledPref; }
  function setEnabled(v) {
    enabledPref = !!v;
    try { localStorage.setItem(VIRT_PREF_KEY, v ? "1" : "0"); } catch (_) {}
  }
  function isForcedByUrl() { return urlForced !== null; }
  function isActive() { return !!st; }

  function fixedHeightOf(row, useExcelLayout) {
    let h = manualRowHeights[row.rowIndex0] || (manualRowHeightAll > 0 ? manualRowHeightAll : 0);
    if (!h && useExcelLayout) h = toPixelHeight(currentSheetRowHeights[row.rowIndex0]) || 0;
    return h || 0;
  }

  function hasMerges(rows, colCount, model) {
    if (model.mode !== "wide" || !Array.isArray(currentMerges) || !currentMerges.length) return false;
    const m = computeMergeLayout(rows, colCount);
    return !!(m && (m.anchors.size || m.covered.size));
  }

  function canVirtualize(rows, headers, model) {
    return enabled()
      && rows.length > VIRT_MIN_ROWS
      && !(tableEl && tableEl.classList.contains("wrap-cells"))
      && !hasMerges(rows, headers.length, model);
  }

  // ── wysokości i pozycje (prefix[i] = górna krawędź wiersza i w „modelu”) ──────
  // st.h[i] = najlepsza wiedza o wysokości wiersza i; st.known[i] = 1 gdy zmierzona
  // albo wymuszona (reszta = defaultH). Tablice typowane → przeliczenie pozycji przy
  // 100 000 wierszy to ułamek milisekundy.
  function heightAt(i) { return st.h[i]; }
  function fillHeights() {
    for (let i = 0; i < st.n; i++) {
      const c = heightCache.get(st.keys[i]);
      const f = st.fixed[i];
      if (c && c.f === f) { st.h[i] = c.h; st.known[i] = 1; }
      else if (f) { st.h[i] = f; st.known[i] = 1; }
      else { st.h[i] = defaultH; st.known[i] = 0; }
    }
  }
  function rebuildPrefix() {
    const n = st.n;
    const p = st.prefix;
    const h = st.h;
    let acc = 0;
    for (let i = 0; i < n; i++) { p[i] = acc; acc += h[i]; }
    p[n] = acc;
  }
  // największe i, dla którego prefix[i] <= y
  function indexAt(y) {
    const p = st.prefix;
    let lo = 0; let hi = st.n - 1;
    if (y <= 0) return 0;
    while (lo < hi) {
      const mid = (lo + hi + 1) >> 1;
      if (p[mid] <= y) lo = mid; else hi = mid - 1;
    }
    return lo;
  }

  function makeSpacer(cls) {
    const tr = document.createElement("tr");
    tr.className = `vr-spacer ${cls}`;
    tr.setAttribute("aria-hidden", "true");
    const td = document.createElement("td");
    td.colSpan = st.colCount + 1;
    tr.appendChild(td);
    return tr;
  }
  function setSpacer(which, px) {
    const tr = which === "top" ? st.topSpacer : st.bottomSpacer;
    const h = Math.max(0, Math.round(px * 100) / 100);
    if (h <= 0.5) {
      if (tr.isConnected) tr.remove();
    } else {
      tr.firstChild.style.height = `${h}px`;
      if (!tr.isConnected) {
        if (which === "top") tbodyEl.insertBefore(tr, tbodyEl.firstChild);
        else tbodyEl.appendChild(tr); // ostatni — żeby `tr:last-child` łapał go, a nie wiersz danych
      }
    }
    if (which === "top") st.topPx = h <= 0.5 ? 0 : h;
  }

  function buildRow(i) {
    return buildTableRow(st.buildCtx, st.rows[i], i);
  }

  function prunePool() {
    const pool = st.buildCtx.nextRowNodes;
    if (pool.size <= VIRT_POOL_MAX) return;
    for (const [key, tr] of pool) {
      if (pool.size <= VIRT_POOL_MAX * 0.8) break;
      if (!tr.isConnected) pool.delete(key);
    }
  }

  // Mierzy świeżo wstawione wiersze (jeden odczyt układu) i zapisuje w pamięci wysokości.
  // Zwraca true, gdy coś się różniło od założeń (trzeba przeliczyć pozycje).
  function measure(list) {
    let changed = false;
    for (let k = 0; k < list.length; k++) {
      const i = list[k];
      const tr = st.nodes[i - st.s];
      if (!tr) continue;
      const h = tr.getBoundingClientRect().height;
      if (!(h > 0)) continue;
      heightCache.set(st.keys[i], { h, f: st.fixed[i] });
      if (Math.abs(st.h[i] - h) > 0.01) changed = true;
      st.h[i] = h;
      st.known[i] = 1;
    }
    return changed;
  }

  // Ustawia DOM na wiersze [ns, ne]. anchor = { i, domTop } — wiersz, który ma zostać
  // dokładnie tam, gdzie był (y względem góry <tbody>), albo null (skok).
  function place(ns, ne, anchor) {
    const oldS = st.s; const oldE = st.e;
    const added = [];
    const overlap = st.nodes.length && ns <= oldE && ne >= oldS;
    if (!overlap) {
      st.nodes.forEach((tr) => tr.remove());
      const frag = document.createDocumentFragment();
      const nodes = [];
      for (let i = ns; i <= ne; i++) { const tr = buildRow(i); nodes.push(tr); frag.appendChild(tr); added.push(i); }
      if (st.bottomSpacer.isConnected) tbodyEl.insertBefore(frag, st.bottomSpacer);
      else tbodyEl.appendChild(frag);
      st.nodes = nodes;
    } else {
      // zdejmij to, co wyjeżdża (okna się nakładają, więc coś zawsze zostaje)
      while (st.s < ns) { st.nodes.shift().remove(); st.s++; }
      while (st.s + st.nodes.length - 1 > ne) st.nodes.pop().remove();
      // dołóż z góry
      if (ns < st.s) {
        const frag = document.createDocumentFragment();
        const fresh = [];
        for (let i = ns; i < st.s; i++) { const tr = buildRow(i); fresh.push(tr); frag.appendChild(tr); added.push(i); }
        tbodyEl.insertBefore(frag, st.nodes[0]);
        st.nodes = fresh.concat(st.nodes);
        st.s = ns;
      }
      // dołóż z dołu
      const curE = st.s + st.nodes.length - 1;
      if (ne > curE) {
        const frag = document.createDocumentFragment();
        for (let i = curE + 1; i <= ne; i++) { const tr = buildRow(i); st.nodes.push(tr); frag.appendChild(tr); added.push(i); }
        if (st.bottomSpacer.isConnected) tbodyEl.insertBefore(frag, st.bottomSpacer);
        else tbodyEl.appendChild(frag);
      }
    }
    st.s = ns; st.e = ne;

    if (added.length && measure(added)) {
      if (!defaultMeasured) {
        const t = typicalHeight();
        if (t) {
          defaultMeasured = true;
          if (Math.abs(t - defaultH) > 0.01) {
            defaultH = t;
            for (let i = 0; i < st.n; i++) if (!st.known[i]) st.h[i] = defaultH;
            tbodyEl.style.setProperty("--vr-row-h", `${defaultH}px`);
          }
        }
      }
      rebuildPrefix();
    }

    // górny wypełniacz: kotwica trzyma widoczny wiersz w miejscu
    let top = st.prefix[ns];
    if (anchor && anchor.i >= ns && anchor.i <= ne) top = anchor.domTop - (st.prefix[anchor.i] - st.prefix[ns]);
    let shift = 0;
    // Ujemnego wypełniacza nie ma — wtedy (i na samej górze poza ruchem) oddajemy
    // różnicę przez scrollTop. Na górze W RUCHU zostaje chwilowa szczelina z siatką,
    // domknie ją settle() po zatrzymaniu.
    if (top < 0 || (ns === 0 && Math.abs(top) > 0.5 && !inMotion())) {
      shift = st.prefix[ns] - top;
      top = st.prefix[ns];
    }
    setSpacer("top", top);
    st.drift = st.topPx - st.prefix[ns];
    setSpacer("bottom", st.prefix[st.n] - st.prefix[ne + 1]);
    if (shift) tableWrapEl.scrollTop += shift;
    prunePool();
  }

  // Typowa wysokość = mediana zmierzonych wierszy bez wymuszonej wysokości.
  function typicalHeight() {
    const hs = [];
    for (let i = st.s; i <= st.e && hs.length < 200; i++) {
      if (st.fixed[i]) continue;
      const c = heightCache.get(st.keys[i]);
      if (c) hs.push(c.h);
    }
    if (!hs.length) return 0;
    hs.sort((a, b) => a - b);
    return hs[hs.length >> 1];
  }

  function viewport() {
    const scrollTop = tableWrapEl.scrollTop;
    const clientH = tableWrapEl.clientHeight || 1;
    return { top: scrollTop - st.bodyTop, clientH };
  }

  function desiredRange(modelTop, clientH) {
    const ns = indexAt(modelTop - clientH * VIRT_OVERSCAN);
    let ne = indexAt(modelTop + clientH * (1 + VIRT_OVERSCAN));
    return [ns, Math.min(st.n - 1, ne)];
  }

  // Wiersz z otwartym edytorem nie może zniknąć z DOM (edytor = <input> w komórce).
  function editorIndex() {
    if (typeof activeCellEditor === "undefined" || !activeCellEditor || !activeCellEditor.td) return -1;
    const tr = activeCellEditor.td.closest("tr[data-row-key]");
    return tr ? st.keyIndex.get(tr.dataset.rowKey) ?? -1 : -1;
  }

  function update(force = false) {
    if (!st || !tableWrapEl) return;
    const { top: domTop, clientH } = viewport();
    const covTop = st.topPx;
    const covBottom = st.topPx + (st.prefix[st.e + 1] - st.prefix[st.s]);
    const margin = clientH * VIRT_MARGIN;
    const needTop = st.s > 0 && domTop - margin < covTop;
    const needBottom = st.e < st.n - 1 && domTop + clientH + margin > covBottom;
    if (!force && !needTop && !needBottom) return;

    const modelTop = domTop - st.drift;
    let [ns, ne] = desiredRange(modelTop, clientH);
    const ei = editorIndex();
    if (ei >= 0 && (ei < ns || ei > ne)) {
      if (Math.abs(ei < ns ? ns - ei : ei - ne) < 300) { ns = Math.min(ns, ei); ne = Math.max(ne, ei); }
      else if (activeCellEditor.input) activeCellEditor.input.blur(); // za daleko — zatwierdź (blur = commit)
    }
    const a = indexAt(Math.max(0, modelTop));
    const anchor = (a >= st.s && a <= st.e) ? { i: a, domTop: st.topPx + (st.prefix[a] - st.prefix[st.s]) } : null;
    place(ns, ne, anchor);
    afterWindowChange();
  }

  // Nowe wiersze w oknie → te same stany co przy pełnym renderze (fokus, zakres).
  function afterWindowChange() {
    if (typeof syncFocusedCellInDom === "function") syncFocusedCellInDom({ clearMissing: false, passive: true });
    if (typeof syncSelectedCellInDom === "function") syncSelectedCellInDom({ clearMissing: false });
    if (typeof syncRangeHighlightInDom === "function") syncRangeHighlightInDom();
  }

  // W spoczynku zerujemy drift: wypełniacz i scrollTop w tej samej klatce = bez ruchu na ekranie.
  function settle() {
    if (!st || inMotion() || Math.abs(st.drift) < 0.5) return;
    const d = st.drift;
    setSpacer("top", st.topPx - d);
    st.drift = st.topPx - st.prefix[st.s];
    tableWrapEl.scrollTop -= d;
  }

  function onScroll() {
    if (!st) return;
    lastScrollAt = performance.now();
    update(false);
    clearTimeout(idleTimer);
    idleTimer = setTimeout(settle, VIRT_IDLE_MS);
  }

  let listenersOn = false;
  function ensureListeners() {
    if (listenersOn || !tableWrapEl) return;
    listenersOn = true;
    tableWrapEl.addEventListener("scroll", onScroll, { passive: true });
    tableWrapEl.addEventListener("touchstart", () => { touchActive = true; }, { passive: true });
    const end = () => { touchActive = false; clearTimeout(idleTimer); idleTimer = setTimeout(settle, VIRT_IDLE_MS); };
    tableWrapEl.addEventListener("touchend", end, { passive: true });
    tableWrapEl.addEventListener("touchcancel", end, { passive: true });
    if (typeof ResizeObserver === "function") {
      let lastH = 0;
      new ResizeObserver(() => {
        if (!st) return;
        const h = tableWrapEl.clientHeight;
        if (h === lastH) return;
        lastH = h;
        update(true);
      }).observe(tableWrapEl);
    }
  }

  function unmount() {
    if (!st) return;
    st = null;
    if (tbodyEl) tbodyEl.classList.remove("is-virtual");
    if (tableWrapEl) tableWrapEl.classList.remove("virt-on");
  }

  // Wołane z renderTable() po zbudowaniu nagłówka, gdy <tbody> jest puste.
  // Zwraca true, gdy przejęliśmy render wierszy.
  function mount(ctx, rows, envKey) {
    if (!tbodyEl || !tableWrapEl || !canVirtualize(rows, ctx.headers, ctx.model)) { unmount(); return false; }
    ensureListeners();
    if (envKey !== heightEnv) {
      heightCache.clear();
      heightEnv = envKey;
      defaultMeasured = false;
      defaultH = Math.round(28 * (parseFloat(zoomLevelEl && zoomLevelEl.value) || 1));
    }
    const n = rows.length;
    const keys = new Array(n);
    const keyIndex = new Map();
    const fixed = new Float32Array(n);
    for (let i = 0; i < n; i++) {
      const k = getRowSelectionKey(rows[i]);
      keys[i] = k;
      keyIndex.set(k, i);
      fixed[i] = fixedHeightOf(rows[i], ctx.useExcelLayout);
    }
    // Wiersze dobudowywane w trakcie przewijania: w tym samym środowisku, więc zawsze
    // wolno użyć ponownie tych, które już raz narysowaliśmy (wracają przy przewinięciu w górę).
    const buildCtx = { ...ctx, canReuse: true, reuseMap: ctx.nextRowNodes };
    st = {
      rows, n, keys, keyIndex, fixed,
      prefix: new Float64Array(n + 1),
      h: new Float32Array(n),
      known: new Uint8Array(n),
      colCount: ctx.headers.length,
      s: 0, e: -1, nodes: [], topPx: 0, drift: 0, bodyTop: 0,
      buildCtx,
    };
    st.topSpacer = makeSpacer("vr-top");
    st.bottomSpacer = makeSpacer("vr-bottom");
    tbodyEl.classList.add("is-virtual");
    // Własne kotwiczenie przewijania przeglądarki (Chromium) poprawiałoby scrollTop
    // RAZEM z naszą kotwicą → podwójna korekta. Wyłączone, gdy silnik działa.
    tableWrapEl.classList.add("virt-on");
    tbodyEl.style.setProperty("--vr-row-h", `${defaultH}px`);
    fillHeights();
    rebuildPrefix();

    // Pierwsze okno budujemy z WŁAŚCIWEGO ctx (ponowne użycie wierszy z poprzedniego renderu).
    const wrapRect = tableWrapEl.getBoundingClientRect();
    st.bodyTop = tbodyEl.getBoundingClientRect().top - wrapRect.top - tableWrapEl.clientTop + tableWrapEl.scrollTop;
    const clientH = tableWrapEl.clientHeight || 600;
    const maxTop = Math.max(0, st.prefix[n] - clientH);
    const modelTop = Math.min(Math.max(0, tableWrapEl.scrollTop - st.bodyTop), maxTop);
    const [ns, ne] = desiredRange(modelTop, clientH);
    const a = indexAt(modelTop);
    const anchor = { i: a, domTop: st.prefix[a] };
    st.buildCtx = ctx; // pierwszy przebieg: jak stary render (canReuse wg envKey)
    place(ns, ne, anchor);
    st.buildCtx = buildCtx;
    return true;
  }

  // Dla kodu, który chce DOTKNĄĆ wiersza spoza okna (strzałki, skok z analiz,
  // „następne trafienie”): przewiń tak, by był na środku, i dorysuj okno od razu.
  function ensureRowRendered(key) {
    if (!st) return null;
    const i = st.keyIndex.get(key);
    if (i == null) return null;
    if (i >= st.s && i <= st.e) return st.nodes[i - st.s];
    // Najpierw bez skoku: przewinięcie (np. scrollIntoView przy strzałkach) mogło już
    // się stać, tylko zdarzenie scroll jeszcze nie dotarło. Dorysuj okno dla BIEŻĄCEJ
    // pozycji — jeśli wiersz jest tuż za krawędzią, trafi do okna, a widok nie drgnie
    // (dalej zadziała zwykłe scrollIntoView „nearest”, o jeden wiersz, jak w Excelu).
    update(true);
    if (i >= st.s && i <= st.e) return st.nodes[i - st.s];
    settleNow();
    const clientH = tableWrapEl.clientHeight || 600;
    const target = st.bodyTop + st.drift + st.prefix[i] - Math.max(0, (clientH - heightAt(i)) / 2);
    tableWrapEl.scrollTop = Math.max(0, target);
    update(true);
    return (i >= st.s && i <= st.e) ? st.nodes[i - st.s] : null;
  }
  function settleNow() {
    if (!st || Math.abs(st.drift) < 0.5) return;
    // (celowo bez inMotion: wołane tylko przy jawnym skoku, który i tak ustawia scrollTop)
    const d = st.drift;
    setSpacer("top", st.topPx - d);
    st.drift = st.topPx - st.prefix[st.s];
    tableWrapEl.scrollTop -= d;
  }
  function ensureRowByIndex0(rowIndex0) {
    if (!st) return null;
    const i = st.rows.findIndex((r) => r.rowIndex0 === rowIndex0);
    return i < 0 ? null : ensureRowRendered(st.keys[i]);
  }

  function hasKey(key) { return !!(st && st.keyIndex.has(key)); }
  function indexOfKey(key) { return st ? (st.keyIndex.get(key) ?? -1) : -1; }
  function renderedRange() { return st ? { s: st.s, e: st.e, n: st.n, drift: st.drift } : null; }

  return {
    enabled, setEnabled, isForcedByUrl, isActive, mount, unmount, update,
    ensureRowRendered, ensureRowByIndex0, hasKey, indexOfKey, renderedRange,
  };
})();
