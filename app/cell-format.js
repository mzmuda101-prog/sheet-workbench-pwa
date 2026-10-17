// cell-format.js — formatowanie komórek jak na karcie „Narzędzia główne” w Excelu
// (pogrubienie, kolory, wyrównanie, format liczb, obramowanie…), zapisywane do pliku.
//
// Model: format jest DELTĄ trzymaną w tym samym miejscu co edycje wartości —
// pendingEdits[arkusz][A1].fmt. Dzięki temu bez osobnej hydrauliki działają:
// licznik na „Zapisz”, Cofnij/Ponów (undo.js robi migawkę wpisu), szkic po awarii
// (drafts.js zapisuje pendingEdits) i sam zapis (xlsx-patch.js: styl z pliku + delta).
//   { v, t }            — nowa wartość (jak dotąd)
//   { v, t, fmt }       — nowa wartość + format
//   { fmt }             — tylko format, wartość z pliku zostaje
//   { clear: true, fmt} — wyczyszczona zawartość, format zostaje
//   null                — wyczyszczona zawartość (bez zmian formatu)
// Delta: { b, i, u, strike, sz, color, fill, h, v, wrap, numFmt, border:{top,…}, reset }.
// Klucz obecny = ustawiony (także na false/null — „wyłącz”, „brak koloru”).
//
// W tabeli format jest NAKŁADANY na styl odczytany z pliku (row.cellStyles) — zarówno
// od razu po kliknięciu, jak i po każdej przebudowie wierszy (cfmtOverlayRows).

const CFMT_FONT_KEYS = ["b", "i", "u", "strike", "sz", "color"];

// Paleta jak w Excelu: kolory motywu (rząd 1) + kolory standardowe (rząd 2).
const CFMT_PALETTE = [
  "000000", "FFFFFF", "44546A", "E7E6E6", "4472C4", "ED7D31", "A5A5A5", "FFC000", "5B9BD5", "70AD47",
  "C00000", "FF0000", "FFC000", "FFFF00", "92D050", "00B050", "00B0F0", "0070C0", "002060", "7030A0",
  "F2F2F2", "D9D9D9", "DDEBF7", "FCE4D6", "EDEDED", "FFF2CC", "DDEBF7", "E2EFDA", "FFE699", "C6E0B4",
];

// Formaty liczb z listy (kod = to, co trafia do pliku).
const CFMT_NUMBER_FORMATS = [
  { id: "general", code: "General" },
  { id: "number", code: "#,##0.00" },
  { id: "integer", code: "#,##0" },
  { id: "negred", code: "#,##0.00;[Red]-#,##0.00" },
  { id: "pln", code: '#,##0.00 "zł"' },
  { id: "eur", code: '#,##0.00 "€"' },
  { id: "percent", code: "0%" },
  { id: "fraction", code: "# ?/?" },
  { id: "date", code: "yyyy-mm-dd" },
  { id: "datetime", code: "yyyy-mm-dd hh:mm" },
  { id: "text", code: "@" },
];

function cfmtRealColCount() {
  const dc = typeof dcAppliedCount === "number" ? dcAppliedCount : 0;
  return Math.max(0, (currentHeaders?.length || 0) - dc);
}

// Kolumna istnieje w pliku (nie jest wirtualną kolumną wyliczaną dopisaną na końcu).
function cfmtIsRealCol(col) {
  return Number.isInteger(col) && col >= 0 && col < cfmtRealColCount();
}

function cfmtRef(rowIndex0, col) {
  return XLSX.utils.encode_cell({ r: rowIndex0, c: currentStartCol + col });
}

function cfmtPendingFmt(sheetName, ref) {
  const p = pendingEdits?.[sheetName]?.[ref];
  return p && p.fmt ? p.fmt : null;
}

function cfmtMerge(prev, delta) {
  if (!delta) return prev ? { ...prev } : {};
  if (delta.reset) {
    const out = { ...delta };
    if (delta.border) out.border = { ...delta.border };
    return out;
  }
  const out = { ...(prev || {}) };
  Object.keys(delta).forEach((k) => {
    if (k === "border") out.border = { ...(out.border || {}), ...delta.border };
    else out[k] = delta[k];
  });
  return out;
}

// Zapis delty do pendingEdits (z zachowaniem wartości, jeśli była edytowana).
function cfmtRecord(sheetName, ref, delta) {
  if (!pendingEdits[sheetName]) pendingEdits[sheetName] = {};
  const cur = pendingEdits[sheetName][ref];
  if (cur === null) pendingEdits[sheetName][ref] = { clear: true, fmt: cfmtMerge(null, delta) };
  else if (cur === undefined) pendingEdits[sheetName][ref] = { fmt: cfmtMerge(null, delta) };
  else pendingEdits[sheetName][ref] = { ...cur, fmt: cfmtMerge(cur.fmt, delta) };
}

function cfmtHex(c) {
  const h = String(c || "").replace(/^#/, "").toUpperCase();
  if (/^[0-9A-F]{8}$/.test(h)) return `#${h.slice(2)}`;
  return /^[0-9A-F]{6}$/.test(h) ? `#${h}` : null;
}

// Styl do rysowania w tabeli = styl z pliku + delta. Idempotentne: ta sama delta
// nałożona drugi raz daje to samo (każdy klucz ustawia wartość wprost).
function cfmtStyleWithFmt(base, fmt) {
  if (!fmt) return base;
  const s = fmt.reset
    ? { fillColor: null, hasCustomFill: false, fontColor: null, hasCustomFontColor: false, fontColorIsDefaultLike: true,
      fontFamily: "", fontScale: 0, bold: false, italic: false, underline: false, horizontal: "", vertical: "", wrapText: false,
      hasBorder: false, border: null }
    : { ...(base || {}) };
  if ("b" in fmt) s.bold = !!fmt.b;
  if ("i" in fmt) s.italic = !!fmt.i;
  if ("u" in fmt) s.underline = !!fmt.u;
  if ("strike" in fmt) s.strike = !!fmt.strike;
  if ("sz" in fmt && Number(fmt.sz) > 0 && typeof getWorkbookDefaultFontSize === "function") {
    s.fontScale = Number(fmt.sz) / getWorkbookDefaultFontSize(workbook);
    s.fontSizePt = Number(fmt.sz);
  }
  if ("fill" in fmt) {
    s.fillColor = cfmtHex(fmt.fill);
    s.hasCustomFill = !!s.fillColor;
  }
  if ("color" in fmt) s.fontColor = cfmtHex(fmt.color);
  if ("color" in fmt || "fill" in fmt) {
    s.fontColorIsDefaultLike = typeof isDefaultLikeFontColor === "function" ? isDefaultLikeFontColor(s.fontColor) : !s.fontColor;
    s.hasCustomFontColor = !!s.fontColor && (!s.fontColorIsDefaultLike || !!s.hasCustomFill);
    // Kolor wybrany ręcznie pokazujemy pełnym tłem (jak w Excelu), nie przygaszonym.
    if ("fill" in fmt) s.userFill = !!s.fillColor;
  }
  if ("h" in fmt) s.horizontal = fmt.h || "";
  if ("v" in fmt) s.vertical = fmt.v || "";
  if ("wrap" in fmt) { s.wrapText = !!fmt.wrap; s.userWrap = !!fmt.wrap; }
  if (fmt.border) {
    const border = { ...(s.border && typeof s.border === "object" ? s.border : {}) };
    ["top", "bottom", "left", "right"].forEach((side) => {
      if (!(side in fmt.border)) return;
      const e = fmt.border[side];
      if (!e || !e.style || e.style === "none") delete border[side];
      else border[side] = { style: e.style, color: e.color ? { rgb: `FF${String(e.color).replace(/^#/, "").toUpperCase()}` } : null };
      delete border[side[0].toUpperCase() + side.slice(1)];
    });
    s.border = border;
    s.hasBorder = typeof hasCustomBorder === "function" ? hasCustomBorder(border) : true;
  }
  if ("numFmt" in fmt) s.numFmt = fmt.numFmt;
  return s;
}

// Wartość w formacie liczbowym (SheetJS SSF — ten sam silnik, którym czytamy plik,
// więc po zapisie i ponownym otwarciu komórka wygląda tak samo).
function cfmtFormatValue(value, code) {
  if (value == null || value === "") return "";
  if (!code || code === "General") return null;
  if (code === "@") return value instanceof Date ? toDisplay(value) : String(value);
  const ssf = typeof XLSX !== "undefined" && XLSX.SSF;
  if (!ssf) return null;
  let v = value;
  if (v instanceof Date) v = typeof dateToExcelSerial === "function" ? dateToExcelSerial(v) : null;
  if (typeof v === "boolean") return null;
  if (typeof v !== "number" || !Number.isFinite(v)) return null; // tekst: format liczb nie dotyczy
  const out = swbSsfFormat(code, v);
  return out == null ? null : swbLocalizeCellText(out, { t: "n", z: code }, currentLang);
}

// Tekst komórki po zmianie wartości lub formatu: format z delty, a gdy go nie ma —
// format zapisany w pliku (cell.z), żeby nowa liczba w kolumnie „zł” też miała „zł”.
function cfmtDisplayFor(row, col) {
  const v = Array.isArray(row.values) ? row.values[col] : null;
  const ref = cfmtRef(row.rowIndex0, col);
  const sheetCell = workbook?.Sheets?.[currentSheetName]?.[ref];
  // Błąd (wpisany „#DZIEL/0!” albo z pliku): format liczbowy go nie dotyczy, nazwa po polsku.
  if (sheetCell && sheetCell.t === "e") return swbLocalizeCellText(String(sheetCell.w || sheetCell.v), sheetCell, currentLang);
  const fmt = cfmtPendingFmt(currentSheetName, ref);
  let code = fmt && "numFmt" in fmt ? fmt.numFmt : null;
  if (code == null) {
    const cell = sheetCell;
    if (cell && cell.z && cell.z !== "General" && !(v instanceof Date)) code = cell.z;
  }
  if (code != null) {
    const out = cfmtFormatValue(v, code);
    if (out != null) return out;
  }
  return v == null ? "" : swbGeneralText(v);
}

function cfmtRefreshDisplay(row, col) {
  if (!row || !Array.isArray(row.display)) return;
  row.display[col] = cfmtDisplayFor(row, col);
  cfmtSyncNumFmtColor(row, col);
}

// ── Kolor z formatu liczbowego: „0,00;[Czerwony]-0,00” ──────────────────────────
// Excel PL pokazuje [Czerwony], [Niebieski]…, a w pliku zapisuje [Red], [Blue]…
// (sprawdzone: number format local → number format). Sekcje: dodatnie;ujemne;zero;tekst,
// z warunkami [<=100] itp. SheetJS formatuje tekst, ale kolor gubi — wcześniej ujemne
// kwoty z formatów księgowych były w tabeli czarne. Kolor liczymy przy RYSOWANIU
// z bieżącej wartości (wynik formuły może zmienić znak), a wiersz trzyma tylko kod
// formatu (row.nfFmt[kol]) i tylko dla formatów, które w ogóle mają kolor.
const SWB_NF_COLOR_NAMES = {
  black: "black", blue: "blue", cyan: "cyan", green: "green", magenta: "magenta", red: "red", white: "white", yellow: "yellow",
  czarny: "black", niebieski: "blue", zielony: "green", czerwony: "red", "biały": "white", "żółty": "yellow",
};
// [Color1]…[Color56] = domyślna paleta Excela
const SWB_NF_PALETTE = [
  "#000000", "#FFFFFF", "#FF0000", "#00FF00", "#0000FF", "#FFFF00", "#FF00FF", "#00FFFF",
  "#800000", "#008000", "#000080", "#808000", "#800080", "#008080", "#C0C0C0", "#808080",
  "#9999FF", "#993366", "#FFFFCC", "#CCFFFF", "#660066", "#FF8080", "#0066CC", "#CCCCFF",
  "#000080", "#FF00FF", "#FFFF00", "#00FFFF", "#800080", "#800000", "#008080", "#0000FF",
  "#00CCFF", "#CCFFFF", "#CCFFCC", "#FFFF99", "#99CCFF", "#FF99CC", "#CC99FF", "#FFCC99",
  "#3366FF", "#33CCCC", "#99CC00", "#FFCC00", "#FF9900", "#FF6600", "#666699", "#969696",
  "#003366", "#339966", "#003300", "#333300", "#993300", "#993366", "#333399", "#333333",
];
const swbNfSectionCache = new Map();

function swbNfHasColor(code) {
  return typeof code === "string" && code.indexOf("[") >= 0 && swbNfSections(code) !== null;
}

// Kod → sekcje [{ color, cond, at }] albo null, gdy żadna sekcja nie ma koloru.
function swbNfSections(code) {
  if (swbNfSectionCache.has(code)) return swbNfSectionCache.get(code);
  const parts = [];
  let cur = "";
  for (let i = 0; i < code.length; i++) {
    const ch = code[i];
    if (ch === '"') { const j = code.indexOf('"', i + 1); const end = j < 0 ? code.length : j; cur += code.slice(i, end + 1); i = end; continue; }
    if (ch === "\\") { cur += code.slice(i, i + 2); i += 1; continue; }
    if (ch === "[") { const j = code.indexOf("]", i + 1); const end = j < 0 ? code.length : j; cur += code.slice(i, end + 1); i = end; continue; }
    if (ch === ";") { parts.push(cur); cur = ""; continue; }
    cur += ch;
  }
  parts.push(cur);
  let any = false;
  const secs = parts.map((p) => {
    let color = null;
    let cond = null;
    (p.match(/\[[^\]]*\]/g) || []).forEach((tok) => {
      const inner = tok.slice(1, -1).trim();
      const named = SWB_NF_COLOR_NAMES[inner.toLowerCase()];
      if (named) { color = named; return; }
      const idx = /^(?:color|kolor)\s*(\d{1,2})$/i.exec(inner);
      if (idx && Number(idx[1]) >= 1 && Number(idx[1]) <= 56) { color = SWB_NF_PALETTE[Number(idx[1]) - 1]; return; }
      const c = /^(<=|>=|<>|<|>|=)\s*(-?\d+(?:[.,]\d+)?)$/.exec(inner);
      if (c) cond = { op: c[1], n: Number(c[2].replace(",", ".")) };
    });
    if (color) any = true;
    return { color, cond, at: /@/.test(p.replace(/"[^"]*"/g, "")) };
  });
  const out = any ? secs : null;
  if (swbNfSectionCache.size > 500) swbNfSectionCache.clear();
  swbNfSectionCache.set(code, out);
  return out;
}

function swbNfCondTrue(cond, n) {
  switch (cond.op) {
    case "<": return n < cond.n;
    case "<=": return n <= cond.n;
    case ">": return n > cond.n;
    case ">=": return n >= cond.n;
    case "=": return n === cond.n;
    default: return n !== cond.n;
  }
}

// Kolor sekcji, która obejmuje wartość: "red" / "blue"… albo "#RRGGBB" ([ColorN]); null = brak.
function swbNumFmtColor(code, value) {
  const secs = swbNfSections(code);
  if (!secs || value == null || value === "" || typeof value === "boolean") return null;
  if (typeof value === "string") {
    const sec = secs.length >= 4 ? secs[3] : (secs.length === 1 && secs[0].at ? secs[0] : null);
    return sec ? sec.color : null;
  }
  const n = value instanceof Date ? 1 : value;
  if (typeof n !== "number" || !Number.isFinite(n)) return null;
  const num = secs.slice(0, 3).filter((sec, i) => i === 0 || !sec.at);
  let sec;
  if (num.some((x) => x.cond)) {
    sec = num.find((x) => x.cond && swbNfCondTrue(x.cond, n)) || num.find((x) => !x.cond) || null;
  } else if (num.length === 1) sec = num[0];
  else if (num.length === 2) sec = n < 0 ? num[1] : num[0];
  else sec = n > 0 ? num[0] : (n < 0 ? num[1] : num[2]);
  return sec ? sec.color : null;
}

// Nałożenie koloru na <td> (wołane przy budowie wiersza tabeli).
function swbApplyNumFmtColor(td, row, col) {
  const code = row && row.nfFmt ? row.nfFmt[col] : null;
  if (!code) return;
  const color = swbNumFmtColor(code, Array.isArray(row.values) ? row.values[col] : null);
  if (!color) return;
  if (color[0] === "#") { td.classList.add("nf-color"); td.style.setProperty("--nf-color", color); }
  else td.classList.add(`nf-${color}`);
}

function cfmtCodeAt(row, col) {
  const ref = cfmtRef(row.rowIndex0, col);
  const fmt = cfmtPendingFmt(currentSheetName, ref);
  if (fmt && "numFmt" in fmt) return typeof fmt.numFmt === "string" ? fmt.numFmt : null;
  const cell = workbook?.Sheets?.[currentSheetName]?.[ref];
  return cell && typeof cell.z === "string" ? cell.z : null;
}

function cfmtSyncNumFmtColor(row, col) {
  if (!row || row.isLongViewRow || !Number.isFinite(row.rowIndex0)) return;
  const code = cfmtCodeAt(row, col);
  if (swbNfHasColor(code)) {
    if (!row.nfFmt) row.nfFmt = [];
    row.nfFmt[col] = code;
  } else if (row.nfFmt) {
    row.nfFmt[col] = undefined;
  }
}

// Po przebudowie wierszy: kody formatów z kolorem z pliku. Przejście po komórkach
// arkusza (nie po wierszach×kolumnach) — formatów z „[” jest zwykle garstka.
function cfmtNumFmtColorPass(rows) {
  const sheet = workbook?.Sheets?.[currentSheetName];
  if (!sheet || !Array.isArray(rows) || !rows.length) return;
  let byRow = null;
  for (const ref in sheet) {
    if (ref.charCodeAt(0) === 33) continue; // „!ref”, „!merges”…
    const z = sheet[ref] && sheet[ref].z;
    if (typeof z !== "string" || z.indexOf("[") < 0 || !swbNfHasColor(z)) continue;
    if (!byRow) {
      byRow = new Map();
      rows.forEach((row) => { if (row && Number.isFinite(row.rowIndex0) && !row.isLongViewRow) byRow.set(row.rowIndex0, row); });
    }
    const { r, c } = XLSX.utils.decode_cell(ref);
    const row = byRow.get(r);
    const col = c - currentStartCol;
    if (!row || col < 0 || !Array.isArray(row.values) || col >= row.values.length) continue;
    if (!row.nfFmt) row.nfFmt = [];
    row.nfFmt[col] = z;
  }
}

// Po każdej przebudowie wierszy (wczytanie arkusza, zmiana wiersza nagłówka…) nałóż
// formaty z bieżącej sesji — inaczej znikałyby do czasu zapisu i ponownego otwarcia.
function cfmtOverlayRows(rows) {
  cfmtNumFmtColorPass(rows);
  const cells = pendingEdits?.[currentSheetName];
  if (!cells || !Array.isArray(rows) || !rows.length) return;
  const refs = Object.keys(cells).filter((ref) => cells[ref] && cells[ref].fmt);
  if (!refs.length) return;
  const byRow = new Map();
  rows.forEach((row) => { if (row && Number.isFinite(row.rowIndex0) && !row.isLongViewRow) byRow.set(row.rowIndex0, row); });
  refs.forEach((ref) => {
    const { r, c } = XLSX.utils.decode_cell(ref);
    const row = byRow.get(r);
    const col = c - currentStartCol;
    if (!row || col < 0 || !Array.isArray(row.values) || col >= row.values.length) return;
    if (!Array.isArray(row.cellStyles)) row.cellStyles = [];
    row.cellStyles[col] = cfmtStyleWithFmt(row.cellStyles[col], cells[ref].fmt);
    if ("numFmt" in cells[ref].fmt) cfmtRefreshDisplay(row, col);
  });
}

// Komórki zaznaczenia (albo sama aktywna komórka) jako lista { row, col, rPos, cPos, rect }.
function cfmtSelectionCells() {
  const rect = typeof rangeCellsToCopy === "function" ? rangeCellsToCopy() : null;
  if (!rect || !rect.model || rect.model.mode !== "wide") return { rect: null, cells: [] };
  const out = [];
  for (let r = rect.rowStart; r <= rect.rowEnd; r++) {
    const row = rect.model.rows[r];
    if (!row || row.isLongViewRow || !Number.isFinite(row.rowIndex0)) continue;
    for (let c = rect.colMin; c <= rect.colMax; c++) {
      if (!cfmtIsRealCol(c)) continue;
      out.push({ row, col: c, rPos: r, cPos: c, rect });
    }
  }
  return { rect, cells: out };
}

// Styl aktywnej komórki (do stanu przycisków paska).
function cfmtActiveCell() {
  if (!focusedCellState || !currentDisplayModel || currentDisplayModel.mode !== "wide") return null;
  const row = typeof getRowByKey === "function" ? getRowByKey(focusedCellState.rowKey) : null;
  if (!row || !Number.isFinite(row.rowIndex0)) return null;
  const col = focusedCellState.colIndex0;
  const style = (Array.isArray(row.cellStyles) && row.cellStyles[col]) || {};
  const ref = cfmtRef(row.rowIndex0, col);
  const fmt = cfmtPendingFmt(currentSheetName, ref);
  const cell = workbook?.Sheets?.[currentSheetName]?.[ref] || null;
  let numFmt = fmt && "numFmt" in fmt ? fmt.numFmt : (cell && cell.z) || cfmtFileNumFmt(ref) || "General";
  if (typeof numFmt === "number") numFmt = "General";
  return { row, col, ref, style, fmt, cell, numFmt, real: cfmtIsRealCol(col) };
}

// Bieżący format liczbowy komórki (delta z sesji → plik → xf pustej komórki → Ogólny).
// Parser wpisów (parseInputValue) potrzebuje go, bo komórka w formacie Tekst („@”) nie
// zamienia „007” na 7, a automatyczny format („12%” → 0%) dostaje tylko komórka Ogólny.
function cfmtNumFmtAt(rowIndex0, col) {
  if (!workbook || !currentSheetName || !Number.isFinite(rowIndex0) || !Number.isInteger(col)) return "General";
  const ref = cfmtRef(rowIndex0, col);
  const fmt = cfmtPendingFmt(currentSheetName, ref);
  if (fmt && "numFmt" in fmt) return typeof fmt.numFmt === "string" && fmt.numFmt ? fmt.numFmt : "General";
  const cell = workbook.Sheets?.[currentSheetName]?.[ref];
  if (cell && typeof cell.z === "string" && cell.z) return cell.z;
  return cfmtFileNumFmt(ref) || "General";
}

// Wpis użytkownika → wartość, z uwzględnieniem formatu komórki (patrz parseInputValue).
function parseCellInput(raw, rowIndex0, col) {
  return parseInputValue(raw, { numFmt: cfmtNumFmtAt(rowIndex0, col) });
}

// Format liczbowy z pliku (xf komórki) — dla komórki pustej, ale sformatowanej (np. pusta
// kolumna dat): SheetJS nie tworzy dla niej obiektu komórki, więc bez tego pasek pokazywał
// „Ogólny”, choć Excel pokazałby format daty. Indeks xf mamy z buildStyleIndexMap.
function cfmtFileNumFmt(ref) {
  const idx = currentStyleIndexMap && currentStyleIndexMap.get(currentSheetName)
    ? currentStyleIndexMap.get(currentSheetName).get(ref) : undefined;
  const st = workbook && workbook.Styles;
  const xf = Number.isInteger(idx) && st && Array.isArray(st.CellXf) ? st.CellXf[idx] : null;
  if (!xf) return null;
  const id = Number(xf.numFmtId ?? xf.numfmtid);
  if (!Number.isFinite(id) || id === 0) return null;
  const code = (st.NumberFmt && st.NumberFmt[id])
    || (XLSX.SSF && typeof XLSX.SSF.get_table === "function" ? XLSX.SSF.get_table()[id] : null);
  return typeof code === "string" && code !== "General" ? code : null;
}

// Wypełnianie (autofill.js): format liczbowy komórki źródła przechodzi na cel — jak
// w Excelu, gdzie przeciągnięta data zostaje datą, a kwota kwotą. Tylko gdy się różni.
function cfmtCopyNumFmt(srcRow, srcCol, dstRow, dstCol) {
  if (!srcRow || !dstRow || !Number.isFinite(srcRow.rowIndex0) || !Number.isFinite(dstRow.rowIndex0)) return false;
  if (!cfmtIsRealCol(dstCol)) return false;
  const code = cfmtNumFmtAt(srcRow.rowIndex0, srcCol);
  if (code === cfmtNumFmtAt(dstRow.rowIndex0, dstCol)) return false;
  if (typeof undoCaptureCell === "function") undoCaptureCell(dstRow.rowIndex0, dstCol);
  const ref = cfmtRef(dstRow.rowIndex0, dstCol);
  cfmtRecord(currentSheetName, ref, { numFmt: code });
  if (!Array.isArray(dstRow.cellStyles)) dstRow.cellStyles = [];
  dstRow.cellStyles[dstCol] = cfmtStyleWithFmt(dstRow.cellStyles[dstCol], cfmtPendingFmt(currentSheetName, ref));
  const sheetCell = workbook.Sheets[currentSheetName][ref];
  if (sheetCell) sheetCell.z = code;
  cfmtRefreshDisplay(dstRow, dstCol);
  return true;
}

// Główna akcja: delta (obiekt) albo funkcja (komórka → delta | null) na zaznaczeniu.
// Zwraca liczbę zmienionych komórek.
function cfmtApply(deltaOrFn, { label } = {}) {
  if (typeof swbEditGuard === "function" && !swbEditGuard()) return 0;
  if (!workbook || !currentSheetName) return 0;
  if (!currentDisplayModel || currentDisplayModel.mode !== "wide") {
    toast(t("editWideOnly"), "info");
    return 0;
  }
  const { cells } = cfmtSelectionCells();
  if (!cells.length) {
    toast(t(focusedCellState ? "fmtDerivedCol" : "fmtPickCell"), "info");
    return 0;
  }
  let changed = 0;
  cells.forEach((ctx) => {
    const delta = typeof deltaOrFn === "function" ? deltaOrFn(ctx) : deltaOrFn;
    if (!delta || !Object.keys(delta).length) return;
    if (typeof undoCaptureCell === "function") undoCaptureCell(ctx.row.rowIndex0, ctx.col);
    const ref = cfmtRef(ctx.row.rowIndex0, ctx.col);
    cfmtRecord(currentSheetName, ref, delta);
    if (!Array.isArray(ctx.row.cellStyles)) ctx.row.cellStyles = [];
    ctx.row.cellStyles[ctx.col] = cfmtStyleWithFmt(ctx.row.cellStyles[ctx.col], cfmtPendingFmt(currentSheetName, ref));
    if ("numFmt" in delta || delta.reset) {
      const sheetCell = workbook.Sheets[currentSheetName][ref];
      if (sheetCell) {
        const code = cfmtPendingFmt(currentSheetName, ref)?.numFmt;
        if (code != null) sheetCell.z = code;
        else if (delta.reset) delete sheetCell.z;
      }
      cfmtRefreshDisplay(ctx.row, ctx.col);
    }
    changed += 1;
  });
  if (!changed) return 0;
  bumpSheetDataStamp();
  setDirtyState(true);
  renderActiveTable();
  if (typeof sheetEditUi !== "undefined") sheetEditUi.sync();
  if (label) log(`${label}: ${changed} kom.`, "success");
  return changed;
}

// Przełącznik B/I/U/S: jak w Excelu decyduje aktywna komórka (pogrubiona → zdejmij).
function cfmtToggle(key) {
  const active = cfmtActiveCell();
  const styleKey = { b: "bold", i: "italic", u: "underline", strike: "strike" }[key];
  const on = !!(active && active.style && active.style[styleKey]);
  return cfmtApply({ [key]: !on });
}

// Zmiana liczby miejsc po przecinku w kodzie formatu (każda sekcja „;” osobno).
function cfmtChangeDecimals(code, delta) {
  let c = !code || code === "General" ? "0" : String(code);
  if (c === "@") return c;
  return c.split(";").map((section) => {
    let done = false;
    return section.replace(/("[^"]*")|(0+)(\.0*)?/g, (m, quoted, ints, dec) => {
      if (quoted || done) return m;
      done = true;
      const n = dec ? dec.length - 1 : 0;
      const next = Math.max(0, Math.min(10, n + delta));
      return ints + (next ? "." + "0".repeat(next) : "");
    });
  }).join(";");
}

// Krawędzie dla presetów obramowania (zależne od położenia komórki w zakresie).
function cfmtBorderDelta(preset, ctx, color) {
  const { rect, rPos, cPos } = ctx;
  const top = rPos === rect.rowStart, bottom = rPos === rect.rowEnd;
  const left = cPos === rect.colMin, right = cPos === rect.colMax;
  const thin = { style: "thin", color };
  const thick = { style: "medium", color };
  switch (preset) {
    case "all": return { border: { top: thin, bottom: thin, left: thin, right: thin } };
    case "none": return { border: { top: null, bottom: null, left: null, right: null } };
    case "outer": {
      const b = {};
      if (top) b.top = thin;
      if (bottom) b.bottom = thin;
      if (left) b.left = thin;
      if (right) b.right = thin;
      return Object.keys(b).length ? { border: b } : null;
    }
    case "thickOuter": {
      const b = {};
      if (top) b.top = thick;
      if (bottom) b.bottom = thick;
      if (left) b.left = thick;
      if (right) b.right = thick;
      return Object.keys(b).length ? { border: b } : null;
    }
    case "bottom": return bottom ? { border: { bottom: thin } } : null;
    case "top": return top ? { border: { top: thin } } : null;
    case "left": return left ? { border: { left: thin } } : null;
    case "right": return right ? { border: { right: thin } } : null;
    default: return null;
  }
}

// Malarz formatów: kopiuje KOMPLET formatu aktywnej komórki (jak widać go w tabeli).
function cfmtCaptureFormat() {
  const a = cfmtActiveCell();
  if (!a) return null;
  const s = a.style || {};
  const hexOf = (c) => (c ? String(c).replace(/^#/, "").toUpperCase() : null);
  const edge = (e) => {
    if (!e) return null;
    const style = e.style || e.Style;
    if (!style || style === "none") return null;
    const col = typeof colorFromStyleNode === "function" ? colorFromStyleNode(e.color || e.Color) : null;
    return { style, color: hexOf(col) };
  };
  const b = s.border || {};
  return {
    reset: true,
    b: !!s.bold, i: !!s.italic, u: !!s.underline, strike: !!s.strike,
    color: s.hasCustomFontColor ? hexOf(s.fontColor) : null,
    fill: s.hasCustomFill ? hexOf(s.fillColor) : null,
    h: s.horizontal || null, v: s.vertical || null, wrap: !!s.userWrap,
    numFmt: a.numFmt || "General",
    border: { top: edge(b.top || b.Top), bottom: edge(b.bottom || b.Bottom), left: edge(b.left || b.Left), right: edge(b.right || b.Right) },
    ...(s.fontSizePt ? { sz: s.fontSizePt } : {}),
  };
}

// Delete / Backspace na zaznaczeniu = „Wyczyść zawartość” (format zostaje).
function cfmtClearContents() {
  if (typeof swbEditGuard === "function" && !swbEditGuard()) return 0;
  if (!currentDisplayModel || currentDisplayModel.mode !== "wide") return 0;
  const { cells } = cfmtSelectionCells();
  let changed = 0;
  cells.forEach(({ row, col }) => {
    if (row.isSubheader) return;
    const v = Array.isArray(row.values) ? row.values[col] : null;
    const ref = cfmtRef(row.rowIndex0, col);
    const has = (v != null && v !== "") || !!workbook.Sheets[currentSheetName]?.[ref]?.f;
    if (!has) return;
    updateSheetCell(row.rowIndex0, col, null);
    if (Array.isArray(row.values)) row.values[col] = null;
    if (Array.isArray(row.rawValues) && row.rawValues !== row.values) row.rawValues[col] = null;
    if (Array.isArray(row.display)) row.display[col] = "";
    changed += 1;
  });
  if (changed) {
    setDirtyState(true);
    renderActiveTable();
    scheduleViewRefresh({ analyses: true, filterBadge: true });
    toast(t("cellsCleared", { count: changed }), "success");
    if (typeof sheetEditUi !== "undefined") sheetEditUi.sync();
  }
  return changed;
}
