// backstage.js — ekran „Plik” jak w Excelu (etap 2 wstążki; docs/WSTAZKA.md, makieta v2 .bs).
//
// Wjeżdża pod rzędem kart (przenikanie + 12 px, 280 ms --ease-drawer; przy „ogranicz ruch” samo
// przenikanie). Po lewej: Wróć, Nowy, Otwórz, Informacje, Zapisz jako, Eksportuj · Opcje, Pomoc.
// Akcje klikają TE SAME przyciski co dotąd (pusty ekran, panel Plik); Eksport…, motyw, język, kopia
// ustawień i skróty to przeniesione kontrolki z menu ⋯ / „Więcej” (gniazda [data-rb] w ribbon.js —
// na telefonie wracają na swoje miejsca). Bez dubli: Anonimizuj jest na karcie Dane.

const swbBackstage = (() => {
  const el = document.getElementById("rbBackstage");
  const fileBtn = document.getElementById("rbFileBtn");
  if (!el || !fileBtn) return null;
  const nav = el.querySelector(".rb-bs-nav");
  const secs = Array.from(el.querySelectorAll("[data-bs-sec-body]"));
  const infoEl = document.getElementById("rbInfo");
  const recentEl = document.getElementById("rbRecent");
  const reduceMq = matchMedia("(prefers-reduced-motion: reduce)");
  const tr = (k, v) => (typeof t === "function" ? t(k, v) : k);
  let current = "new";
  let closeTimer = 0;

  const isOpen = () => !el.hidden && el.classList.contains("is-open");
  const inertTargets = () => [document.querySelector(".workspace"), document.querySelector(".rb-r2")].filter(Boolean);

  function place() {
    const r1 = document.querySelector(".rb-r1");
    if (r1) el.style.top = `${Math.round(r1.getBoundingClientRect().bottom)}px`;
  }

  function show(sec, { focus = false } = {}) {
    current = sec;
    secs.forEach((s) => { s.hidden = s.dataset.bsSecBody !== sec; });
    nav.querySelectorAll("[data-bs-sec]").forEach((b) => {
      if (b.dataset.bsSec === sec) b.setAttribute("aria-current", "true"); else b.removeAttribute("aria-current");
    });
    if (sec === "info") renderInfo();
    // znacznik języka liczy się z rozmiarów przycisków — gdy był schowany (menu ⋯), miał 0 px
    if (sec === "opts" && typeof updateLangSwitchIndicator === "function") requestAnimationFrame(() => updateLangSwitchIndicator());
    if (sec === "new" && recentEl && typeof swbRecent !== "undefined") swbRecent.render(recentEl);
    if (focus) nav.querySelector(`[data-bs-sec="${sec}"]`)?.focus();
  }

  function open(sec = "new") {
    if (!document.body.classList.contains("rb-on")) return;
    clearTimeout(closeTimer);
    if (typeof swbRibbon !== "undefined" && swbRibbon?.closeMenus) swbRibbon.closeMenus();
    place();
    el.hidden = false;
    fileBtn.setAttribute("aria-expanded", "true");
    inertTargets().forEach((x) => { x.inert = true; });
    show(sec);
    // klasa w następnej klatce = przejście CSS od stanu początkowego (bez @starting-style w starszym Safari)
    requestAnimationFrame(() => requestAnimationFrame(() => el.classList.add("is-open")));
    nav.querySelector(`[data-bs-sec="${sec}"]`)?.focus();
  }

  function close({ refocus = true } = {}) {
    if (el.hidden) return;
    el.classList.remove("is-open");
    fileBtn.setAttribute("aria-expanded", "false");
    inertTargets().forEach((x) => { x.inert = false; });
    // wyjście szybsze niż wejście (Emil): 200 ms, potem hidden
    clearTimeout(closeTimer);
    closeTimer = setTimeout(() => { el.hidden = true; }, reduceMq.matches ? 0 : 200);
    if (refocus) fileBtn.focus();
  }

  // ── Informacje: to, co już wiadomo o otwartym pliku ──────────────────────────
  function renderInfo() {
    if (!infoEl) return;
    const wb = typeof workbook !== "undefined" ? workbook : null;
    infoEl.replaceChildren();
    if (!wb) {
      infoEl.append(Object.assign(document.createElement("p"), { className: "rb-bs-lead", textContent: tr("rbBsInfoEmpty") }));
      return;
    }
    const rows = typeof viewRows !== "undefined" && Array.isArray(viewRows) ? viewRows.length : 0;
    const cols = typeof currentHeaders !== "undefined" && Array.isArray(currentHeaders) ? currentHeaders.length : 0;
    const items = [
      [currentFileName || "—", tr("rbBsInfoFile"), true],
      [typeof currentSheetName === "string" ? currentSheetName : "—", tr("rbBsInfoSheet")],
      [String(wb.SheetNames?.length || 0), tr("rbBsInfoSheets")],
      [String(rows), tr("rbBsInfoRows")],
      [String(cols), tr("rbBsInfoCols")],
    ];
    for (const [val, label, lead] of items) {
      const box = Object.assign(document.createElement("div"), { className: lead ? "lead" : "" });
      box.append(Object.assign(document.createElement("b"), { textContent: val }), Object.assign(document.createElement("small"), { textContent: label }));
      infoEl.append(box);
    }
  }

  // ── akcje: te same przyciski co dotąd ────────────────────────────────────────
  const ACTS = {
    back: () => close(),
    open: () => { close({ refocus: false }); document.getElementById("emptyOpenBtn")?.click(); },
    new: () => { close({ refocus: false }); document.getElementById("emptyNewBtn")?.click(); },
    sample: () => { close({ refocus: false }); document.getElementById("loadSampleBtn")?.click(); },
    saveas: () => { close({ refocus: false }); document.getElementById("saveAsBtn")?.click(); },
  };
  el.addEventListener("click", (e) => {
    const sec = e.target.closest("[data-bs-sec]");
    if (sec && nav.contains(sec)) { show(sec.dataset.bsSec); return; }
    const a = e.target.closest("[data-bs-act]");
    if (a) { ACTS[a.dataset.bsAct]?.(); return; }
    // przyciski, które otwierają okno albo panel (Eksport…, Skróty, panel boczny) — ekran ustępuje
    const leaves = e.target.closest("#exportCsvBtn, #shortcutsMenuItem, #settingsBackupItem, [data-rb-panel]");
    if (leaves) close({ refocus: false });
  });
  // strzałki w pionie po liście po lewej
  nav.addEventListener("keydown", (e) => {
    if (e.key !== "ArrowDown" && e.key !== "ArrowUp") return;
    const items = Array.from(nav.querySelectorAll("button"));
    const i = items.indexOf(document.activeElement);
    if (i < 0) return;
    e.preventDefault();
    items[(i + (e.key === "ArrowDown" ? 1 : -1) + items.length) % items.length].focus();
  });
  fileBtn.addEventListener("click", () => (isOpen() ? close() : open()));
  document.addEventListener("keydown", (e) => {
    // Esc należy do okna nad ekranem Plik (eksport, skróty), jeśli jakieś WIDAĆ — wtedy nie zamykamy
    const overlay = () => Array.from(document.querySelectorAll('.modal, [aria-modal="true"]'))
      .some((m) => m !== el && !el.contains(m) && m.getClientRects().length > 0 && getComputedStyle(m).visibility !== "hidden");
    if (e.key === "Escape" && !el.hidden && !overlay()) {
      e.preventDefault();
      close();
    }
  });
  // telefon / zmiana progu: ekran znika (na telefonie nie ma wstążki)
  document.addEventListener("swb:ribbon", (e) => { if (!e.detail?.on) close({ refocus: false }); });
  window.addEventListener("resize", () => { if (!el.hidden) place(); }, { passive: true });

  return { open, close, isOpen, show };
})();
