// app/cell-spill.js — TEKST WYLEWA SIĘ NA PUSTE KOMÓRKI OBOK, jak w Excelu (2026-10-31).
//
// Wyrocznia Excela (arkusz testowy → PDF z Excela, 2026-10-31):
//   • wylewa się tylko TEKST (także wynik formuły i liczba zapisana jako tekst); liczba,
//     data, PRAWDA/FAŁSZ i błąd (#N/D) — nigdy;
//   • do lewej / ogólne → w prawo; do prawej → w lewo (widać KONIEC tekstu); wyśrodkowane →
//     w obie strony, nadal wyśrodkowane na swojej komórce (zajęty bok obcina, środek zostaje);
//   • zatrzymuje się na pierwszej NIEPUSTEJ komórce — także spacji i formule dającej "";
//   • przechodzi nad wypełnieniem pustej komórki (tekst na żółtym tle sąsiada);
//   • nie wylewa się: zawijanie, scalenie, wyrównanie „Wypełnij” / „Wyjustuj” / „Rozłożone”;
//   • ukryte kolumny przeskakuje (u nas: kolumny poza widokiem po prostu nie są sąsiadami).
//
// Jak: komórka dostaje .cell-spill (overflow: visible) i clip-path przycięty do pustych
// sąsiadów — szerokości kolumn z JEDNEGO pomiaru rzędu liter (A, B, C…) po renderze, nie
// z mierzenia tekstu. Tekst siedzi w <span class="spill-t"> (flex:
// flex-end / center), bo zwykły tekst za szeroki na komórkę zawsze wystaje w prawo;
// span ma pointer-events: none, żeby wylany tekst nie zabierał kliknięć pustym sąsiadom.
// clip-path robi z komórki osobny kontekst warstw → malowana PO tłach sąsiadów (żółte
// wypełnienie, zaznaczenie), czyli tekst leży na wierzchu, jak w Excelu.

const swbSpill = (() => {
  if (typeof document === "undefined" || typeof tableEl === "undefined" || !tableEl) {
    return { schedule() {}, invalidate() {}, refreshRow() {}, visibleWidth() { return 0; } };
  }
  const NO_SPILL_ALIGN = new Set(["fill", "justify", "distributed"]);
  const ERROR_RE = /^#(?:N\/[AD]|DZIEL\/0!|DIV\/0!|ARG!|VALUE!|ADR!|REF!|NAZWA\?|NAME\?|LICZBA!|NUM!|ZERO!|NULL!|SPILL!|CALC!|ROZLANIE!|OBL!)$/;
  const MAX_REACH = 4000; // px — dalej i tak nikt nie czyta jednej linii
  let widths = null; // szerokości kolumn widoku (px), indeks = data-col-index
  let fontPx = 13;
  let raf = 0;

  function measure() {
    const guide = tableEl.querySelector("thead tr.guide-row");
    if (!guide) return false;
    widths = [];
    for (const th of guide.children) {
      const i = parseInt(th.dataset.colIndex, 10);
      if (Number.isFinite(i)) widths[i] = th.getBoundingClientRect().width;
    }
    // litery nie mają data-col-index? — pozycja w rzędzie (0 = narożnik)
    if (!widths.length) [...guide.children].slice(1).forEach((th, i) => { widths[i] = th.getBoundingClientRect().width; });
    const sample = tableEl.querySelector("tbody td:not(.row-head)");
    if (sample) fontPx = parseFloat(getComputedStyle(sample).fontSize) || fontPx;
    return widths.length > 0;
  }

  const isText = (v) => typeof v === "string" && v !== "" && !ERROR_RE.test(v);
  const sheetNow = () => (typeof workbook !== "undefined" && workbook && workbook.Sheets ? workbook.Sheets[currentSheetName] : null);
  function emptyAt(row, td, i, sheet) {
    if (!td || td.classList.contains("cell-merged") || td.classList.contains("cell-editing")) return false;
    const v = row.values ? row.values[i] : null;
    if (v !== null && v !== undefined && v !== "") return false;
    if (td.textContent !== "") return false; // coś jednak widać (np. świeżo wpisane, model jeszcze stary)
    // formuła dająca "" też zatrzymuje (wyrocznia)
    if (sheet && typeof row.rowIndex0 === "number") {
      const cell = sheet[XLSX.utils.encode_cell({ r: row.rowIndex0, c: currentStartCol + i })];
      if (cell && cell.f) return false;
    }
    return true;
  }

  function clear(td) {
    if (!td._swbSpill) return;
    td._swbSpill = null;
    td.classList.remove("cell-spill", "spill-right", "spill-center");
    td.style.removeProperty("--spill-l");
    td.style.removeProperty("--spill-r");
    const span = td.firstElementChild;
    if (span && span.classList.contains("spill-t")) td.textContent = span.textContent;
  }

  function apply(td, align, L, R) {
    const key = `${align}|${L}|${R}`;
    const prev = td._swbSpill;
    if (prev && prev.key === key && td.firstElementChild && td.firstElementChild.classList.contains("spill-t")) return;
    if (prev) clear(td);
    td._swbSpill = { key, align, L, R };
    td.classList.add("cell-spill");
    if (align !== "left") td.classList.add(align === "right" ? "spill-right" : "spill-center");
    // tekst w <span> bez pointer-events: klik w pustą komórkę obok trafia w NIĄ, nie w tę
    // (clip-path rozszerza też obszar trafień, a ta komórka leży warstwą wyżej)
    const span = document.createElement("span");
    span.className = "spill-t";
    span.textContent = td.textContent;
    td.replaceChildren(span);
    td.style.setProperty("--spill-l", `${Math.round(L)}px`);
    td.style.setProperty("--spill-r", `${Math.round(R)}px`);
  }

  function spillRow(tr, row, opts) {
    const { sheet, frozen, off } = opts;
    const byCol = [];
    for (const td of tr.children) {
      const i = td.dataset.colIndex;
      if (i !== undefined && !td.classList.contains("row-head")) byCol[+i] = td;
    }
    for (let i = 0; i < byCol.length; i++) {
      const td = byCol[i];
      if (!td) continue;
      const v = row && row.values ? row.values[i] : null;
      const style = row && row.cellStyles ? row.cellStyles[i] : null;
      const h = String((style && style.horizontal) || "").toLowerCase();
      if (off || !row || !isText(v) || (frozen && i === 0)
        || td.classList.contains("cell-merged") || td.classList.contains("cell-wrap") || td.classList.contains("cell-editing")
        || (style && style.userWrap) || NO_SPILL_ALIGN.has(h)) { clear(td); continue; }
      const text = td.textContent;
      const inner = (widths[i] || 0) - 24 * (parseFloat(getComputedStyleZoom()) || 1);
      // tani szacunek bez mierzenia tekstu: krótki tekst i tak się mieści
      if (!text || text.length * fontPx * 0.42 <= inner) { clear(td); continue; }
      const align = h === "right" ? "right" : (h === "center" || h === "centercontinuous") ? "center" : "left";
      let L = 0;
      let R = 0;
      if (align !== "right") for (let j = i + 1; j < byCol.length && R < MAX_REACH && emptyAt(row, byCol[j], j, sheet); j++) R += widths[j] || 0;
      if (align !== "left") for (let j = i - 1; j >= 0 && L < MAX_REACH && !(frozen && j === 0) && emptyAt(row, byCol[j], j, sheet); j--) L += widths[j] || 0;
      if (!L && !R) clear(td);
      else apply(td, align, L, R);
    }
  }

  let zoomCache = null;
  function getComputedStyleZoom() {
    if (zoomCache === null) zoomCache = getComputedStyle(tableEl).getPropertyValue("--table-zoom") || "1";
    return zoomCache;
  }

  function context() {
    return {
      sheet: sheetNow(),
      frozen: typeof tableWrapEl !== "undefined" && tableWrapEl && tableWrapEl.classList.contains("freeze-first-col"),
      off: tableEl.classList.contains("wrap-cells")
        || !(typeof currentDisplayModel !== "undefined" && currentDisplayModel && currentDisplayModel.mode === "wide"),
    };
  }

  function rowOf(tr) {
    return tr._swbRow || (typeof getRowByKey === "function" ? getRowByKey(tr.dataset.rowKey) : null);
  }

  function runAll() {
    raf = 0;
    if (!widths && !measure()) return;
    const opts = context();
    for (const tr of tableEl.querySelectorAll("tbody > tr[data-row-key], thead > tr.pre-row[data-row-key]")) {
      spillRow(tr, rowOf(tr), opts);
    }
  }

  function schedule() {
    if (!raf) raf = requestAnimationFrame(runAll);
  }
  // nowy układ kolumn (render, zoom, zmiana szerokości, obrót ekranu) → zmierz od nowa
  function invalidate() {
    widths = null;
    zoomCache = null;
    schedule();
  }
  // jedna komórka zmieniła treść w miejscu (edycja, Cofnij) — przelicz jej wiersz
  function refreshRow(tr) {
    if (!tr) return;
    if (tr._swbRow && typeof getRowByKey === "function") tr._swbRow = getRowByKey(tr.dataset.rowKey) || tr._swbRow;
    schedule();
  }

  // Szerokość, na której tekst komórki jest naprawdę widoczny (podpowiedź „ucięty tekst”).
  function visibleWidth(td) {
    const s = td && td._swbSpill;
    if (!s) return 0;
    if (s.align === "center") return 2 * Math.min(s.L, s.R);
    return s.L + s.R;
  }

  if (typeof ResizeObserver === "function") {
    let lastW = 0;
    new ResizeObserver((entries) => {
      const w = Math.round(entries[0].contentRect.width);
      if (w !== lastW) { lastW = w; invalidate(); }
    }).observe(tableEl);
  }

  return { schedule, invalidate, refreshRow, visibleWidth };
})();
