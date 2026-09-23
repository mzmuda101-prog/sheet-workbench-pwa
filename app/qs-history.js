// qs-history.js — historia zapytań szybkiego szukania.
//
// Puste pole szukania (fokus, wyczyszczenie tekstu albo ↓) pokazuje w miejscu
// live-podglądu listę: „Przypięte" + „Ostatnie". Klik = to samo zapytanie z tymi
// samymi ikonkami (≠ = a… .*) i operatorami, od razu zastosowane. ☆ przypina,
// × usuwa z historii. Wpisy to te same .qs-live-item co trafienia, więc ↓/↑/Enter
// działają bez żadnej nowej obsługi klawiatury.
//
// Historia jest preferencją URZĄDZENIA (localStorage), nie pliku: te same zapytania
// wracają przy kolejnym arkuszu. Nic nie wychodzi poza przeglądarkę.

const QS_HISTORY_KEY = "excel-workbench-qs-history";
const QS_HISTORY_MAX_RECENT = 12;
const QS_HISTORY_MAX_PINNED = 10;

function qsHistoryLoad() {
  try {
    const raw = JSON.parse(localStorage.getItem(QS_HISTORY_KEY) || "null");
    if (raw && Array.isArray(raw.recent) && Array.isArray(raw.pinned)) return raw;
  } catch { /* prywatne okno / uszkodzony wpis */ }
  return { recent: [], pinned: [] };
}

function qsHistorySave(h) {
  try { localStorage.setItem(QS_HISTORY_KEY, JSON.stringify(h)); } catch { /* brak miejsca / prywatne okno */ }
}

// Ten sam tekst z innymi ikonkami to INNE zapytanie („faktura" vs „≠ faktura").
function qsHistoryKey(e) {
  return `${e.mode || "contains"}|${e.neg ? 1 : 0}|${e.ops ? 1 : 0}|${e.q}`;
}

function rememberQsQuery(query) {
  const q = String(query || "").trim();
  if (!q || q.length > 500) return;
  const flags = typeof getQuickSearchFlags === "function" ? getQuickSearchFlags() : { mode: "contains", negated: false };
  const entry = { q, mode: flags.mode, neg: !!flags.negated, ops: !!quickSearchOperatorsEnabled };
  const key = qsHistoryKey(entry);
  const h = qsHistoryLoad();
  h.recent = [entry, ...h.recent.filter((e) => qsHistoryKey(e) !== key)].slice(0, QS_HISTORY_MAX_RECENT);
  qsHistorySave(h);
}

function qsHistoryTogglePin(entry) {
  const key = qsHistoryKey(entry);
  const h = qsHistoryLoad();
  if (h.pinned.some((e) => qsHistoryKey(e) === key)) {
    h.pinned = h.pinned.filter((e) => qsHistoryKey(e) !== key);
  } else {
    h.pinned = [entry, ...h.pinned].slice(0, QS_HISTORY_MAX_PINNED);
  }
  qsHistorySave(h);
}

function qsHistoryRemove(entry) {
  const key = qsHistoryKey(entry);
  const h = qsHistoryLoad();
  h.recent = h.recent.filter((e) => qsHistoryKey(e) !== key);
  h.pinned = h.pinned.filter((e) => qsHistoryKey(e) !== key);
  qsHistorySave(h);
}

// Przywraca zapytanie RAZEM z ikonkami i operatorami, potem zwykły commit
// (ten sam co Enter — także zakres „wszystkie arkusze").
function qsHistoryApply(ctx, entry) {
  if (filterModeEl) filterModeEl.value = entry.mode || "contains";
  if (filterNegateEl) filterNegateEl.checked = !!entry.neg;
  [filterOperatorsEl, quickSearchOperatorsEl, quickSearchPopupOperatorsEl].forEach((el) => { if (el) el.checked = !!entry.ops; });
  quickSearchOperatorsEnabled = !!entry.ops;
  if (quickSearchEl) quickSearchEl.value = entry.q;
  if (quickSearchPopupInput) quickSearchPopupInput.value = entry.q;
  syncQuickSearchModeControls();
  hideQsLive(ctx);
  commitQuickSearch();
}

const QS_FLAG_GLYPH = { equals: "=", starts_with: "a…", regex: ".*" };

function qsHistoryRow(ctx, entry, pinned) {
  const row = document.createElement("div");
  row.className = "qs-hist-row";

  const item = document.createElement("button");
  item.type = "button";
  item.className = "qs-live-item qs-hist-item";
  item.setAttribute("role", "option");
  item.tabIndex = -1;
  if (entry.neg) {
    const n = document.createElement("span");
    n.className = "qs-hist-flag is-neg";
    n.textContent = "≠";
    item.appendChild(n);
  }
  if (QS_FLAG_GLYPH[entry.mode]) {
    const m = document.createElement("span");
    m.className = "qs-hist-flag";
    m.textContent = QS_FLAG_GLYPH[entry.mode];
    item.appendChild(m);
  }
  const text = document.createElement("span");
  text.className = "qs-hist-text";
  text.textContent = entry.q;
  item.appendChild(text);
  item.addEventListener("click", () => qsHistoryApply(ctx, entry));

  const pin = document.createElement("button");
  pin.type = "button";
  pin.className = `qs-hist-btn qs-hist-pin${pinned ? " is-pinned" : ""}`;
  pin.tabIndex = -1;
  pin.textContent = pinned ? "★" : "☆";
  pin.setAttribute("aria-label", t(pinned ? "qsHistoryUnpin" : "qsHistoryPin"));
  pin.setAttribute("aria-pressed", String(pinned));
  pin.addEventListener("click", (e) => {
    e.stopPropagation();
    qsHistoryTogglePin(entry);
    renderQsHistory(ctx);
    ctx.inputEl?.focus({ preventScroll: true });
  });

  const del = document.createElement("button");
  del.type = "button";
  del.className = "qs-hist-btn qs-hist-del";
  del.tabIndex = -1;
  del.textContent = "×";
  del.setAttribute("aria-label", t("qsHistoryRemove"));
  del.addEventListener("click", (e) => {
    e.stopPropagation();
    qsHistoryRemove(entry);
    renderQsHistory(ctx);
    ctx.inputEl?.focus({ preventScroll: true });
  });

  row.append(item, pin, del);
  return row;
}

// Zwraca true, gdy lista jest pokazana (ui-controls wtedy nie chowa podglądu).
function renderQsHistory(ctx) {
  if (!ctx || !ctx.liveEl || !ctx.inputEl) return false;
  if ((ctx.inputEl.value || "").trim()) return false;
  const h = qsHistoryLoad();
  const pinnedKeys = new Set(h.pinned.map(qsHistoryKey));
  const recent = h.recent.filter((e) => !pinnedKeys.has(qsHistoryKey(e)));
  if (!h.pinned.length && !recent.length) { hideQsLive(ctx); return false; }

  const frag = document.createDocumentFragment();
  const section = (label, list, pinned) => {
    if (!list.length) return;
    const title = document.createElement("div");
    title.className = "qs-live-group";
    title.textContent = label;
    frag.appendChild(title);
    list.forEach((e) => frag.appendChild(qsHistoryRow(ctx, e, pinned)));
  };
  section(t("qsHistoryPinned"), h.pinned, true);
  section(t("qsHistoryRecent"), recent, false);

  if (recent.length) {
    const clear = document.createElement("button");
    clear.type = "button";
    clear.className = "qs-hist-clear";
    clear.tabIndex = -1;
    clear.textContent = t("qsHistoryClear");
    clear.addEventListener("click", (e) => {
      e.stopPropagation();
      const cur = qsHistoryLoad();
      cur.recent = [];
      qsHistorySave(cur);
      renderQsHistory(ctx);
      ctx.inputEl?.focus({ preventScroll: true });
    });
    frag.appendChild(clear);
  }

  ctx.liveEl.replaceChildren(frag);
  ctx.liveEl.classList.add("is-history");
  ctx.liveEl.classList.remove("hidden");
  return true;
}
