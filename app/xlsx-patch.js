// =====================================================================
// ZIP-patch zapisu XLSX — zachowuje oryginalny plik (tabele, wykresy, style,
// formuły, rysunki) i podmienia TYLKO wartości edytowanych komórek w sheet.xml.
//
// Dlaczego: xlsx-js-style (jak SheetJS) przy XLSX.write gubi tabele Excela,
// wykresy, rysunki i calcChain — co psuje formuły z odwołaniami strukturalnymi
// (np. SUBTOTAL(9,Dane[Kwota])) i formatowanie tabel. Zamiast przepisywać plik
// od zera, rozpakowujemy oryginalny ZIP (JSZip), nanosimy edycje na XML arkusza
// i pakujemy z powrotem — reszta pliku zostaje nietknięta.
//
// WAŻNE (2026-07): patch dotyka WYŁĄCZNIE wnętrza <sheetData> metodą stringową.
// Pełny DOMParser+XMLSerializer na całym worksheet (4MB+, emoji w formułach)
// psuł plik przy drugim zapisie (UTF-16 surogaty w UTF-8, znikające puste <c/>).
// Sekcje dataValidations / extLst / mergeCells pozostają bajt-identyczne.
// =====================================================================

const SPREADSHEETML_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const OOXML_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

// Regex komórki w sheetData — oba warianty: <c …></c> oraz samozamykające <c …/>.
// WAŻNE: w wariancie otwartym attrs = [^>/]* (NIE [^>]*). Inaczej <c r="H1" s="17"/>
// jest mylone z otwarciem (attrs kończy się na "/"), a ([\s\S]*?)</c> połyka NASTĘPNĄ
// komórkę — przy edycji H powstaje <c r="H1" s="17"/><v>…</v></c> (Excel: uszkodzony).
const CELL_BLOCK_RE = /<c r="([A-Z]+\d+)"([^>/]*)>([\s\S]*?)<\/c>|<c r="([A-Z]+\d+)"([^/>]*)\/>/g;

// Date -> numer seryjny Excela (system 1900, dni od 1899-12-30). Używa komponentów
// lokalnych daty (tak jak SheetJS zwraca Date przy cellDates), niezależnie od strefy.
function dateToExcelSerial(date) {
  const epoch = Date.UTC(1899, 11, 30);
  const utc = Date.UTC(
    date.getFullYear(), date.getMonth(), date.getDate(),
    date.getHours(), date.getMinutes(), date.getSeconds(), date.getMilliseconds()
  );
  return (utc - epoch) / 86400000;
}

function escapeXmlText(s) {
  return String(s)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

// Usuwa znaki niedozwolone w XML 1.0 (tab/LF/CR zostają). SAMOTNE połówki UTF-16
// (bug XMLSerializer przy emoji) też wyrzucamy — inaczej Excel odmawia otwarcia.
// UWAGA: tylko samotne! Poprawna PARA (emoji 🔥, 🎉) zostaje — dawniej zakres
// [\uD800-\uDFFF] bez sprawdzania pary zjadał każde emoji z edytowanej komórki.
function sanitizeXmlText(s) {
  return String(s)
    .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F\uFFFE\uFFFF]/g, "")
    .replace(/[\uD800-\uDBFF](?![\uDC00-\uDFFF])|(?<![\uD800-\uDBFF])[\uDC00-\uDFFF]/g, "");
}

function escapeFormulaXml(text) {
  return escapeXmlText(text);
}

// Mapa: nazwa arkusza -> ścieżka w ZIP (xl/worksheets/sheetN.xml).
async function resolveSheetPaths(zip) {
  const wbFile = zip.file("xl/workbook.xml");
  const relsFile = zip.file("xl/_rels/workbook.xml.rels");
  if (!wbFile || !relsFile) return {};
  const parser = new DOMParser();
  const wb = parser.parseFromString(await wbFile.async("string"), "application/xml");
  const rels = parser.parseFromString(await relsFile.async("string"), "application/xml");

  const idToTarget = {};
  Array.from(rels.getElementsByTagName("Relationship")).forEach((r) => {
    idToTarget[r.getAttribute("Id")] = r.getAttribute("Target");
  });

  const map = {};
  Array.from(wb.getElementsByTagName("sheet")).forEach((s) => {
    const name = s.getAttribute("name");
    const rid =
      s.getAttributeNS(OOXML_REL_NS, "id") || s.getAttribute("r:id");
    if (!name || !rid) return;
    let target = idToTarget[rid];
    if (!target) return;
    if (target.startsWith("/")) target = target.slice(1);
    else target = "xl/" + target.replace(/^\.\//, "");
    map[name] = target;
  });
  return map;
}

// Wycina <sheetData>…</sheetData> bez ruszania reszty arkusza (dataValidations, extLst…).
function splitSheetData(xml) {
  const open = xml.indexOf("<sheetData");
  if (open < 0) return null;
  const openEnd = xml.indexOf(">", open);
  if (openEnd < 0) return null;
  const close = xml.indexOf("</sheetData>", openEnd);
  if (close < 0) return null;
  return {
    before: xml.slice(0, openEnd + 1),
    inner: xml.slice(openEnd + 1, close),
    after: xml.slice(close),
  };
}

// Indeks komórek w sheetData (ref → metadane bloku XML).
function indexCells(inner) {
  const map = new Map();
  CELL_BLOCK_RE.lastIndex = 0;
  let m;
  while ((m = CELL_BLOCK_RE.exec(inner))) {
    if (m[1]) {
      map.set(m[1], {
        ref: m[1], attrs: m[2], inner: m[3], full: m[0],
        start: m.index, end: m.index + m[0].length, selfClose: false,
      });
    } else {
      map.set(m[4], {
        ref: m[4], attrs: m[5], inner: "", full: m[0],
        start: m.index, end: m.index + m[0].length, selfClose: true,
      });
    }
  }
  return map;
}

function sharedFormulaInInner(inner) {
  const m = inner.match(/<f t="shared"([^>/]*)(?:\/>|>([\s\S]*?)<\/f>)/);
  if (!m) return null;
  const siM = m[1].match(/\bsi="(\d+)"/);
  if (!siM) return null;
  return { si: siM[1], text: m[2] || "", fullF: m[0] };
}

function replaceSharedFormulaInInner(inner, formulaText) {
  const re = /<f t="shared"[^>]*(?:\/>|>[\s\S]*?<\/f>)/;
  if (!formulaText) return inner.replace(re, "");
  return inner.replace(re, `<f>${escapeFormulaXml(formulaText)}</f>`);
}

// FORMUŁY DZIELONE — od-dziel grupy dotknięte edycją (stringowo, bez serializera).
function unshareTouchedGroupsString(inner, cells, formulaMap, cellMap) {
  const groups = {};
  cellMap.forEach((cell, ref) => {
    const info = sharedFormulaInInner(cell.inner);
    if (!info) return;
    (groups[info.si] = groups[info.si] || []).push({ ref, cell, text: info.text });
  });

  const replacements = [];
  Object.keys(groups).forEach((si) => {
    const members = groups[si];
    const touched = members.some((m) => Object.prototype.hasOwnProperty.call(cells, m.ref));
    if (!touched) return;
    members.forEach((m) => {
      const text = m.text || (formulaMap && formulaMap[m.ref]) || "";
      const newInner = replaceSharedFormulaInInner(m.cell.inner, text);
      if (newInner === m.cell.inner) return;
      const newFull = `<c r="${m.ref}"${m.cell.attrs}>${newInner}</c>`;
      replacements.push({ start: m.cell.start, len: m.cell.full.length, text: newFull });
    });
  });

  replacements.sort((a, b) => b.start - a.start);
  let out = inner;
  replacements.forEach((r) => {
    out = out.slice(0, r.start) + r.text + out.slice(r.start + r.len);
  });
  return out;
}

const FILE_ERROR_CODES = new Set(["#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A"]);

// Komórka z formułą: <f> + ostatni policzony wynik w <v> (Excel i tak przelicza przy
// otwarciu — fullCalcOnLoad). Typ wyniku: liczba bez t, tekst t="str", logiczna t="b",
// błąd t="e". Brak wyniku (funkcja, której nie liczymy) = sama formuła, bez <v>.
function buildFormulaCellXml(ref, payload, attrs) {
  const a = String(attrs || "").replace(/\/\s*$/g, "").replace(/\s*\bt="[^"]*"/g, "").trim();
  const f = `<f>${escapeFormulaXml(sanitizeXmlText(String(payload.f).replace(/^=/, "")))}</f>`;
  const v = payload.v;
  let tAttr = "";
  let vXml = "";
  if (v instanceof Date) vXml = `<v>${dateToExcelSerial(v)}</v>`;
  else if (typeof v === "number" && Number.isFinite(v)) vXml = `<v>${v}</v>`;
  else if (typeof v === "boolean") { tAttr = ' t="b"'; vXml = `<v>${v ? 1 : 0}</v>`; }
  // Tylko 7 klasycznych kodów: #SPILL!/#CALC! zapisane w <v> Excel „naprawia” na tekst.
  else if (payload.err && FILE_ERROR_CODES.has(payload.err)) { tAttr = ' t="e"'; vXml = `<v>${payload.err}</v>`; }
  else if (typeof v === "string") { tAttr = ' t="str"'; vXml = `<v>${escapeXmlText(sanitizeXmlText(v))}</v>`; }
  return `<c r="${ref}"${a ? " " + a : ""}${tAttr}>${f}${vXml}</c>`;
}

function normalizeCellAttrs(attrs, payload) {
  // [EN] Strip stray "/" left if a self-closing cell was ever mis-parsed as open.
  let a = String(attrs || "").replace(/\/\s*$/g, "").replace(/\s*\bt="[^"]*"/g, "").trim();
  if (payload && payload.t === "s") a = (a ? a + " " : "") + 't="inlineStr"';
  else if (payload && payload.t === "b") a = (a ? a + " " : "") + 't="b"'; // PRAWDA/FAŁSZ: <v>1</v> / <v>0</v>
  else if (payload && payload.t === "e") a = (a ? a + " " : "") + 't="e"'; // wpisany błąd: <v>#DIV/0!</v>
  return a ? " " + a : "";
}

function buildCellXml(ref, payload, attrs, dateStyler) {
  // Data zapisuje się jako liczba seryjna — bez formatu daty w stylu Excel pokazałby
  // „45424”. Styler podmienia s= na wariant tego samego stylu z formatem daty.
  if (dateStyler && payload.t === "d") attrs = withDateStyle(attrs, dateStyler);
  if (payload.f) return buildFormulaCellXml(ref, payload, attrs);
  const attrStr = normalizeCellAttrs(attrs, payload);
  if (payload.t === "s") {
    const t = escapeXmlText(sanitizeXmlText(payload.v));
    return `<c r="${ref}"${attrStr}><is><t xml:space="preserve">${t}</t></is></c>`;
  }
  // Wpisany błąd (#DZIEL/0! → kod angielski). W <v> zawsze angielski kod z listy — polska
  // nazwa czy #SPILL! w pliku to dla Excela uszkodzenie; spoza listy = tekst, nie błąd.
  if (payload.t === "e") {
    if (FILE_ERROR_CODES.has(payload.err)) return `<c r="${ref}"${attrStr}><v>${payload.err}</v></c>`;
    const t = escapeXmlText(sanitizeXmlText(String(payload.err ?? "")));
    return `<c r="${ref}"${normalizeCellAttrs(attrs, { t: "s" })}><is><t xml:space="preserve">${t}</t></is></c>`;
  }
  let num;
  if (payload.t === "d" && payload.v instanceof Date) num = dateToExcelSerial(payload.v);
  else num = Number(payload.v);
  // NaN / Infinity / 1e400 w <v> to dla Excela uszkodzenie (naprawia na tekst) — pusta wartość.
  if (!Number.isFinite(num)) return `<c r="${ref}"${attrStr}/>`;
  return `<c r="${ref}"${attrStr}><v>${num}</v></c>`;
}

function colIndex0(ref) {
  return XLSX.utils.decode_cell(ref).c;
}

function rowNum(ref) {
  return XLSX.utils.decode_cell(ref).r + 1;
}

// Wstawia nową komórkę w istniejącym wierszu (zachowuje rosnącą kolejność kolumn).
function insertCellInRow(rowInner, ref, cellXml) {
  const targetCol = colIndex0(ref);
  const cellRe = /<c r="([A-Z]+\d+)"[^>]*(?:\/>|>[\s\S]*?<\/c>)/g;
  let m;
  while ((m = cellRe.exec(rowInner))) {
    if (colIndex0(m[1]) > targetCol) {
      return rowInner.slice(0, m.index) + cellXml + rowInner.slice(m.index);
    }
  }
  return rowInner + cellXml;
}

// Wstawia nowy wiersz z jedną komórką w sheetData (rosnące r).
function insertRowWithCell(inner, ref, cellXml) {
  const rNum = rowNum(ref);
  const rowRe = /<row r="(\d+)"[^>]*>/g;
  let insertAt = inner.length;
  let m;
  while ((m = rowRe.exec(inner))) {
    if (parseInt(m[1], 10) > rNum) {
      insertAt = m.index;
      break;
    }
  }
  const rowXml = `<row r="${rNum}">${cellXml}</row>`;
  return inner.slice(0, insertAt) + rowXml + inner.slice(insertAt);
}

function insertNewCell(inner, ref, cellXml) {
  const rNum = rowNum(ref);
  const rowRe = new RegExp(`<row r="${rNum}"([^>]*)>([\\s\\S]*?)<\\/row>`);
  const m = inner.match(rowRe);
  if (m) {
    const newRowInner = insertCellInRow(m[2], ref, cellXml);
    const newRow = `<row r="${rNum}"${m[1]}>${newRowInner}</row>`;
    return inner.slice(0, m.index) + newRow + inner.slice(m.index + m[0].length);
  }
  return insertRowWithCell(inner, ref, cellXml);
}

function patchSheetDataInner(inner, cells, formulaMap, dateStyler) {
  if (!cells || Object.keys(cells).length === 0) return inner;

  let cellMap = indexCells(inner);
  inner = unshareTouchedGroupsString(inner, cells, formulaMap, cellMap);
  cellMap = indexCells(inner);

  const replacements = [];
  const inserts = [];

  Object.keys(cells).forEach((ref) => {
    const payload = cells[ref];
    if (payload === undefined) return;

    const existing = cellMap.get(ref);
    if (payload === null) {
      // Jak „Wyczyść zawartość” w Excelu: znika wartość/formuła, styl (s=) zostaje —
      // inaczej czyszczenie komórki w obramowanej tabeli zjadało ramkę, tło i format.
      if (existing) {
        const s = (existing.attrs.match(/\bs="\d+"/) || [])[0];
        replacements.push({ start: existing.start, len: existing.full.length, text: s ? `<c r="${ref}" ${s}/>` : "" });
      }
      return;
    }

    // Zmiana formatu (pasek „Narzędzia główne”): nowy styl = styl z pliku + delta.
    const fmt = payload.fmt && dateStyler && typeof dateStyler.withFmt === "function" ? payload.fmt : null;
    if (fmt && (payload.clear || !("t" in payload))) {
      const curAttrs = existing ? existing.attrs : "";
      const idx = dateStyler.withFmt(cellStyleIdx(curAttrs), fmt);
      let text;
      if (payload.clear || !existing) text = `<c r="${ref}"${withCellStyleIdx(curAttrs.replace(/\s*\bt="[^"]*"/g, "").replace(/\/\s*$/g, ""), idx)}/>`;
      else {
        const a = withCellStyleIdx(existing.attrs.replace(/\/\s*$/g, ""), idx);
        text = existing.selfClose ? `<c r="${ref}"${a}/>` : `<c r="${ref}"${a}>${existing.inner}</c>`;
      }
      if (/^<c r="[A-Z]+\d+"\/>$/.test(text) && !existing) return; // pusta komórka bez stylu — nic do wstawienia
      if (existing) replacements.push({ start: existing.start, len: existing.full.length, text });
      else inserts.push({ ref, cellXml: text });
      return;
    }

    // Bez pisarza stylów (plik bez styles.xml) sama zmiana formatu nie ma czego zapisać;
    // „wyczyść + format” zostaje zwykłym wyczyszczeniem.
    if (payload.clear) {
      if (existing) {
        const s = (existing.attrs.match(/\bs="\d+"/) || [])[0];
        replacements.push({ start: existing.start, len: existing.full.length, text: s ? `<c r="${ref}" ${s}/>` : "" });
      }
      return;
    }
    if (!("t" in payload)) return;

    let attrsIn = existing ? existing.attrs : "";
    let styler = dateStyler;
    if (fmt) {
      attrsIn = withCellStyleIdx(attrsIn, dateStyler.withFmt(cellStyleIdx(attrsIn), fmt, { date: payload.t === "d" }));
      styler = null; // format daty już rozstrzygnięty w withFmt
    }
    const cellXml = buildCellXml(ref, payload, attrsIn, styler);
    if (existing) {
      replacements.push({ start: existing.start, len: existing.full.length, text: cellXml });
    } else {
      inserts.push({ ref, cellXml });
    }
  });

  replacements.sort((a, b) => b.start - a.start);
  let out = inner;
  replacements.forEach((r) => {
    out = out.slice(0, r.start) + r.text + out.slice(r.start + r.len);
  });

  inserts.forEach(({ ref, cellXml }) => {
    out = insertNewCell(out, ref, cellXml);
  });

  return out;
}

// Nanosi edycje na arkusz — poza <sheetData> XML jest bajt-identyczny z oryginałem.
function patchSheetXml(xml, cells, formulaMap, dateStyler) {
  const parts = splitSheetData(xml);
  if (!parts) return xml;
  const patchedInner = patchSheetDataInner(parts.inner, cells, formulaMap, dateStyler);
  if (patchedInner === parts.inner) return xml;
  return parts.before + patchedInner + parts.after;
}

// ─────────────────────────────────────────────────────────────────────────────
// FORMAT DATY dla zapisywanych dat (styles.xml, stringowo — bez serializera).
// Komórka z datą w pliku to liczba + styl z formatem daty. Gdy data trafia do
// komórki, której styl NIE jest datą (pusta, „Ogólny”, tekst po konwersji), dokładamy
// na końcu <cellXfs> KOPIĘ jej stylu z numFmtId="14" (krótka data wg ustawień Excela)
// — obramowanie/czcionka/wypełnienie zostają. Jedna kopia na styl źródłowy.
// Wbudowane formaty dat/czasu Excela (ECMA-376 18.8.30 + warianty azjatyckie).
const BUILTIN_DATE_FMT_IDS = new Set([14, 15, 16, 17, 18, 19, 20, 21, 22, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36,
  45, 46, 47, 50, 51, 52, 53, 54, 55, 56, 57, 58]);
const XF_RE = /<((?:[\w-]+:)?)xf\b(?:[^>]*?\/>|[^>]*?>[\s\S]*?<\/\1xf>)/g;

function isDateFormatCode(code) {
  const c = String(code || "")
    .replace(/"[^"]*"/g, "")      // teksty w cudzysłowie („ d”)
    .replace(/\\./g, "")          // znaki poprzedzone \
    .replace(/\[[^\]]*\]/g, "");   // [$-415], [Red], [h] itp.
  return /[dmyhs]/i.test(c) && !/^general$/i.test(c.trim());
}

function createDateStyler(stylesXml) {
  const m = stylesXml.match(/<((?:[\w-]+:)?)cellXfs\b([^>]*)>([\s\S]*?)<\/\1cellXfs>/);
  if (!m) return null;
  const prefix = m[1];
  const xfs = m[3].match(XF_RE) || [];
  if (!xfs.length) return null;
  const customFmts = {};
  (stylesXml.match(/<(?:[\w-]+:)?numFmt\b[^>]*\/?>/g) || []).forEach((tag) => {
    const id = (tag.match(/\bnumFmtId="(\d+)"/) || [])[1];
    const code = (tag.match(/\bformatCode="([^"]*)"/) || [])[1];
    if (id != null) customFmts[id] = (code || "").replace(/&quot;/g, '"').replace(/&amp;/g, "&");
  });
  const fmtOf = (xf) => Number((xf.match(/\bnumFmtId="(\d+)"/) || [])[1] || 0);
  const isDateXf = (xf) => {
    const id = fmtOf(xf);
    return BUILTIN_DATE_FMT_IDS.has(id) || (customFmts[id] != null && isDateFormatCode(customFmts[id]));
  };
  const added = [];
  const cache = new Map();
  return {
    ensure(sIdx) {
      const i = Number.isFinite(sIdx) && sIdx >= 0 && sIdx < xfs.length ? sIdx : 0;
      if (isDateXf(xfs[i])) return i;
      if (cache.has(i)) return cache.get(i);
      let clone = xfs[i];
      const head = clone.match(/^<(?:[\w-]+:)?xf\b[^>]*?(?=\/?>)/)[0];
      let newHead = /\bnumFmtId="/.test(head) ? head.replace(/\bnumFmtId="\d+"/, 'numFmtId="14"') : head + ' numFmtId="14"';
      newHead = /\bapplyNumberFormat="/.test(newHead) ? newHead.replace(/\bapplyNumberFormat="[^"]*"/, 'applyNumberFormat="1"') : newHead + ' applyNumberFormat="1"';
      clone = newHead + clone.slice(head.length);
      const idx = xfs.length + added.length;
      added.push(clone);
      cache.set(i, idx);
      return idx;
    },
    get changed() { return added.length > 0; },
    apply(xml) {
      if (!added.length) return xml;
      return xml.replace(/<((?:[\w-]+:)?)cellXfs\b([^>]*)>([\s\S]*?)<\/\1cellXfs>/, (all, p, attrs, body) => {
        const total = xfs.length + added.length;
        const newAttrs = /\bcount="/.test(attrs) ? attrs.replace(/\bcount="\d+"/, `count="${total}"`) : `${attrs} count="${total}"`;
        return `<${p}cellXfs${newAttrs}>${body}${added.join("")}</${p}cellXfs>`;
      });
    },
    prefix,
  };
}

function withDateStyle(attrs, dateStyler) {
  const a = String(attrs || "");
  const cur = a.match(/\bs="(\d+)"/);
  const next = dateStyler.ensure(cur ? Number(cur[1]) : 0);
  if (cur) return a.replace(/\bs="\d+"/, `s="${next}"`);
  return next === 0 ? a : `${a} s="${next}"`;
}

// ─────────────────────────────────────────────────────────────────────────────
// FORMATOWANIE KOMÓREK (pasek „Narzędzia główne”) → styles.xml, stringowo.
//
// Edycja niesie „deltę” formatu (fmt), nie gotowy styl: { b, i, u, strike, sz, color,
// fill, h, v, wrap, numFmt, border: { top|bottom|left|right: {style,color}|null }, reset }.
// Przy zapisie bierzemy styl, który komórka ma W PLIKU (s="N"), robimy jego kopię
// z naniesioną deltą i dopisujemy na końcu <fonts>/<fills>/<borders>/<numFmts>/<cellXfs>.
// Wszystko, czego delta nie dotyka (krój, obramowanie, ochrona, wcięcie…), zostaje.
// Identyczne wyniki są scalane (ten sam xf/font/fill nie mnoży się przy 500 komórkach).
// Ten sam obiekt obsługuje też daty (ensure) — dwa osobne pisarze dopisujące do
// <cellXfs> rozjechałyby sobie numerację.
const STYLE_FONT_ORDER = ["b", "i", "strike", "condense", "extend", "outline", "shadow", "u", "vertAlign", "sz", "color", "name", "family", "charset", "scheme"];
const STYLE_BORDER_ORDER = ["start", "end", "left", "right", "top", "bottom", "diagonal", "vertical", "horizontal"];
const BUILTIN_NUMFMT_IDS = {
  General: 0, "0": 1, "0.00": 2, "#,##0": 3, "#,##0.00": 4, "0%": 9, "0.00%": 10, "0.00E+00": 11,
  "# ?/?": 12, "# ??/??": 13, "h:mm": 20, "h:mm:ss": 21, "@": 49,
};

function styleSection(xml, name) {
  const re = new RegExp(`<((?:[\\w-]+:)?)${name}\\b([^>]*?)(?:\\/>|>([\\s\\S]*?)<\\/\\1${name}>)`);
  const m = xml.match(re);
  if (!m) return null;
  return { prefix: m[1], attrs: m[2], body: m[3] || "" };
}

function styleItems(body, name) {
  const re = new RegExp(`<((?:[\\w-]+:)?)${name}\\b(?:[^>]*?\\/>|[^>]*?>[\\s\\S]*?<\\/\\1${name}>)`, "g");
  return body.match(re) || [];
}

function parseXmlAttrs(s) {
  const out = new Map();
  const re = /([\w:.-]+)="([^"]*)"/g;
  let m;
  while ((m = re.exec(s || ""))) out.set(m[1], m[2]);
  return out;
}

function attrsToString(map) {
  let s = "";
  map.forEach((v, k) => { s += ` ${k}="${v}"`; });
  return s;
}

function escapeXmlAttr(s) {
  return escapeXmlText(s).replace(/"/g, "&quot;");
}

// Element złożony z dzieci (font, border): { p, name, attrs, kids:{local:raw}, extra:[raw] }.
function parseStyleElement(xml, name) {
  const open = String(xml || "").match(new RegExp(`^<((?:[\\w-]+:)?)${name}\\b([^>]*?)(\\/?)>`));
  if (!open) return { p: "", name, attrs: "", kids: {}, extra: [] };
  const el = { p: open[1], name, attrs: open[2], kids: {}, extra: [] };
  if (open[3]) return el;
  const inner = xml.slice(open[0].length, xml.lastIndexOf("</"));
  const re = /<((?:[\w-]+:)?)([A-Za-z]+)\b[^>]*?(?:\/>|>[\s\S]*?<\/\1\2>)/g;
  let m;
  while ((m = re.exec(inner))) {
    if (el.kids[m[2]] === undefined) el.kids[m[2]] = m[0];
    else el.extra.push(m[0]);
  }
  return el;
}

function serializeStyleElement(el, order) {
  const known = order.map((k) => el.kids[k] || "").join("");
  const other = Object.keys(el.kids).filter((k) => !order.includes(k)).map((k) => el.kids[k]).join("");
  const body = known + other + el.extra.join("");
  return body ? `<${el.p}${el.name}${el.attrs}>${body}</${el.p}${el.name}>` : `<${el.p}${el.name}${el.attrs}/>`;
}

function normHexColor(c) {
  const h = String(c || "").replace(/^#/, "").toUpperCase();
  if (/^[0-9A-F]{6}$/.test(h)) return `FF${h}`;
  if (/^[0-9A-F]{8}$/.test(h)) return h;
  return null;
}

function createStyleWriter(stylesXml) {
  const xfSec = styleSection(stylesXml, "cellXfs");
  if (!xfSec) return null;
  const P = xfSec.prefix;
  const sections = {
    font: styleSection(stylesXml, "fonts"),
    fill: styleSection(stylesXml, "fills"),
    border: styleSection(stylesXml, "borders"),
  };
  const lists = {
    xf: styleItems(xfSec.body, "xf"),
    font: sections.font ? styleItems(sections.font.body, "font") : [],
    fill: sections.fill ? styleItems(sections.fill.body, "fill") : [],
    border: sections.border ? styleItems(sections.border.body, "border") : [],
  };
  if (!lists.xf.length) return null;
  const base = { xf: lists.xf.length, font: lists.font.length, fill: lists.fill.length, border: lists.border.length };
  const numFmtSec = styleSection(stylesXml, "numFmts");
  const numFmtTags = numFmtSec ? styleItems(numFmtSec.body, "numFmt") : [];
  const baseNumFmts = numFmtTags.length;
  const customFmts = {};
  let maxFmtId = 163;
  numFmtTags.forEach((tag) => {
    const a = parseXmlAttrs(tag);
    const id = Number(a.get("numFmtId"));
    const code = decodeXmlEntitiesSimple(a.get("formatCode") || "");
    if (Number.isFinite(id)) { customFmts[id] = code; maxFmtId = Math.max(maxFmtId, id); }
  });
  const indexOfItem = new Map(); // `${kind}|${xml}` → indeks (scalanie identycznych)
  ["xf", "font", "fill", "border"].forEach((kind) => lists[kind].forEach((x, i) => {
    const key = `${kind}|${x}`;
    if (!indexOfItem.has(key)) indexOfItem.set(key, i);
  }));
  const addItem = (kind, xml) => {
    const key = `${kind}|${xml}`;
    if (indexOfItem.has(key)) return indexOfItem.get(key);
    lists[kind].push(xml);
    const idx = lists[kind].length - 1;
    indexOfItem.set(key, idx);
    return idx;
  };
  const numFmtIdFor = (code) => {
    if (code == null) return 0;
    if (Object.prototype.hasOwnProperty.call(BUILTIN_NUMFMT_IDS, code)) return BUILTIN_NUMFMT_IDS[code];
    for (const id of Object.keys(customFmts)) if (customFmts[id] === code) return Number(id);
    maxFmtId += 1;
    customFmts[maxFmtId] = code;
    numFmtTags.push(`<${P}numFmt numFmtId="${maxFmtId}" formatCode="${escapeXmlAttr(code)}"/>`);
    return maxFmtId;
  };
  const isDateFmtId = (id) => BUILTIN_DATE_FMT_IDS.has(id) || (customFmts[id] != null && isDateFormatCode(customFmts[id]));

  const parseXf = (xml) => {
    const el = parseStyleElement(xml, "xf");
    const attrs = parseXmlAttrs(el.attrs);
    const alignment = el.kids.alignment ? parseXmlAttrs(el.kids.alignment) : new Map();
    return { el, attrs, alignment };
  };
  const num = (v) => { const n = Number(v); return Number.isFinite(n) ? n : 0; };
  const clampXf = (sIdx) => (Number.isFinite(sIdx) && sIdx >= 0 && sIdx < lists.xf.length ? sIdx : 0);

  const fontKeys = ["b", "i", "u", "strike", "sz", "color", "fontName"];
  function patchFont(fontXml, fmt) {
    const el = parseStyleElement(fontXml || `<${P}font/>`, "font");
    if (!el.name) el.name = "font";
    const p = el.p;
    const flag = (key, tag, raw) => {
      if (!(key in fmt)) return;
      if (fmt[key]) el.kids[tag] = raw || `<${p}${tag}/>`;
      else delete el.kids[tag];
    };
    flag("b", "b");
    flag("i", "i");
    flag("strike", "strike");
    flag("u", "u");
    if ("sz" in fmt && Number(fmt.sz) > 0) el.kids.sz = `<${p}sz val="${Number(fmt.sz)}"/>`;
    if ("color" in fmt) {
      const rgb = normHexColor(fmt.color);
      if (rgb) el.kids.color = `<${p}color rgb="${rgb}"/>`;
      else delete el.kids.color; // „Automatyczny”
    }
    if ("fontName" in fmt && fmt.fontName) el.kids.name = `<${p}name val="${escapeXmlAttr(fmt.fontName)}"/>`;
    return serializeStyleElement(el, STYLE_FONT_ORDER);
  }
  function fillXml(color) {
    const rgb = normHexColor(color);
    if (!rgb) return `<${P}fill><${P}patternFill patternType="none"/></${P}fill>`;
    return `<${P}fill><${P}patternFill patternType="solid"><${P}fgColor rgb="${rgb}"/><${P}bgColor indexed="64"/></${P}patternFill></${P}fill>`;
  }
  function patchBorder(borderXml, edges) {
    const el = parseStyleElement(borderXml || `<${P}border/>`, "border");
    const p = el.p;
    ["left", "right", "top", "bottom"].forEach((side) => {
      if (!(side in edges)) return;
      const e = edges[side];
      if (!e || !e.style || e.style === "none") { el.kids[side] = `<${p}${side}/>`; return; }
      const rgb = normHexColor(e.color);
      const color = rgb ? `<${p}color rgb="${rgb}"/>` : `<${p}color indexed="64"/>`;
      el.kids[side] = `<${p}${side} style="${e.style}">${color}</${p}${side}>`;
    });
    // Excel zawsze pisze komplet krawędzi (puste też) — trzymamy się tego kształtu.
    ["left", "right", "top", "bottom", "diagonal"].forEach((side) => { if (!el.kids[side]) el.kids[side] = `<${p}${side}/>`; });
    return serializeStyleElement(el, STYLE_BORDER_ORDER);
  }

  const cache = new Map();
  function withFmt(sIdx, fmt, opts = {}) {
    const i0 = clampXf(sIdx);
    const f = fmt || {};
    const key = `${i0}|${opts.date ? 1 : 0}|${JSON.stringify(f)}`;
    if (cache.has(key)) return cache.get(key);
    const xf = parseXf(lists.xf[f.reset ? 0 : i0]);
    const { attrs, alignment } = xf;
    if (fontKeys.some((k) => k in f)) {
      const fontId = num(attrs.get("fontId"));
      attrs.set("fontId", String(addItem("font", patchFont(lists.font[fontId], f))));
      attrs.set("applyFont", "1");
    }
    if ("fill" in f) {
      attrs.set("fillId", String(addItem("fill", fillXml(f.fill))));
      attrs.set("applyFill", "1");
    }
    if (f.border && typeof f.border === "object") {
      const borderId = num(attrs.get("borderId"));
      attrs.set("borderId", String(addItem("border", patchBorder(lists.border[borderId], f.border))));
      attrs.set("applyBorder", "1");
    }
    if ("numFmt" in f) {
      attrs.set("numFmtId", String(numFmtIdFor(f.numFmt)));
      attrs.set("applyNumberFormat", "1");
    } else if (opts.date && !isDateFmtId(num(attrs.get("numFmtId")))) {
      attrs.set("numFmtId", "14");
      attrs.set("applyNumberFormat", "1");
    }
    let alignChanged = false;
    const setAlign = (key, attr, val) => {
      if (!(key in f)) return;
      alignChanged = true;
      if (val == null || val === "" || val === false) alignment.delete(attr);
      else alignment.set(attr, String(val));
    };
    setAlign("h", "horizontal", f.h);
    setAlign("v", "vertical", f.v);
    setAlign("wrap", "wrapText", f.wrap ? "1" : null);
    if (alignChanged) {
      attrs.set("applyAlignment", "1");
      if (alignment.size) xf.el.kids.alignment = `<${xf.el.p}alignment${attrsToString(alignment)}/>`;
      else delete xf.el.kids.alignment;
    }
    for (const k of ["fontId", "fillId", "borderId", "numFmtId"]) if (!attrs.has(k)) attrs.set(k, "0");
    xf.el.attrs = attrsToString(attrs);
    const idx = addItem("xf", serializeStyleElement(xf.el, ["alignment", "protection", "extLst"]));
    cache.set(key, idx);
    return idx;
  }

  function rewriteSection(xml, name, kind) {
    const total = lists[kind].length;
    if (total === base[kind]) return xml;
    const added = lists[kind].slice(base[kind]).join("");
    const re = new RegExp(`<((?:[\\w-]+:)?)${name}\\b([^>]*?)(?:\\/>|>([\\s\\S]*?)<\\/\\1${name}>)`);
    return xml.replace(re, (all, p, attrs, body) => {
      const a = /\bcount="/.test(attrs) ? attrs.replace(/\bcount="\d+"/, `count="${total}"`) : `${attrs} count="${total}"`;
      return `<${p}${name}${a}>${body || ""}${added}</${p}${name}>`;
    });
  }

  return {
    withFmt,
    // zgodność z dawnym stylerem dat (withDateStyle)
    ensure(sIdx) { return withFmt(sIdx, null, { date: true }); },
    get changed() {
      return lists.xf.length !== base.xf || lists.font.length !== base.font || lists.fill.length !== base.fill
        || lists.border.length !== base.border || numFmtTags.length !== baseNumFmts;
    },
    apply(xml) {
      let out = xml;
      if (numFmtTags.length !== baseNumFmts) {
        const body = numFmtTags.join("");
        const tag = `<${P}numFmts count="${numFmtTags.length}">${body}</${P}numFmts>`;
        if (numFmtSec) out = out.replace(/<((?:[\w-]+:)?)numFmts\b[^>]*?(?:\/>|>[\s\S]*?<\/\1numFmts>)/, tag);
        else out = out.replace(/<((?:[\w-]+:)?)styleSheet\b[^>]*>/, (open) => open + tag);
      }
      if (sections.font) out = rewriteSection(out, "fonts", "font");
      if (sections.fill) out = rewriteSection(out, "fills", "fill");
      if (sections.border) out = rewriteSection(out, "borders", "border");
      out = rewriteSection(out, "cellXfs", "xf");
      return out;
    },
    prefix: P,
  };
}

function decodeXmlEntitiesSimple(s) {
  return String(s).replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&amp;/g, "&");
}

function cellStyleIdx(attrs) {
  const m = String(attrs || "").match(/\bs="(\d+)"/);
  return m ? Number(m[1]) : 0;
}

function withCellStyleIdx(attrs, idx) {
  const a = String(attrs || "");
  if (/\bs="\d+"/.test(a)) return a.replace(/\bs="\d+"/, `s="${idx}"`);
  return idx ? `${a} s="${idx}"` : a;
}

async function resolveStylesPath(zip) {
  const rels = zip.file("xl/_rels/workbook.xml.rels");
  if (rels) {
    const xml = await rels.async("string");
    const rel = (xml.match(/<Relationship\b[^>]*relationships\/styles"[^>]*>/) || [])[0];
    const target = rel && (rel.match(/\bTarget="([^"]+)"/) || [])[1];
    if (target) {
      const path = target.startsWith("/") ? target.slice(1) : "xl/" + target.replace(/^\.\//, "");
      if (zip.file(path)) return path;
    }
  }
  return zip.file("xl/styles.xml") ? "xl/styles.xml" : null;
}

// Wymusza pełne przeliczenie formuł przy otwarciu (Excel pokaże aktualne wyniki
// po zmianie wartości wejściowych). Uwaga: calcChain usuwamy osobno w dropCalcChain
// (wraz z wpisami w [Content_Types].xml i workbook.xml.rels), bo po edycji komórek
// z formułami bywał niespójny i Excel zgłaszał uszkodzenie.
async function forceRecalcOnLoad(zip) {
  const f = zip.file("xl/workbook.xml");
  if (!f) return;
  let xml = await f.async("string");
  if (/<calcPr\b/.test(xml)) {
    if (/\bfullCalcOnLoad=/.test(xml)) {
      xml = xml.replace(/\bfullCalcOnLoad="[^"]*"/, 'fullCalcOnLoad="1"');
    } else {
      xml = xml.replace(/<calcPr\b/, '<calcPr fullCalcOnLoad="1"');
    }
    zipWriteUtf8Xml(zip, "xl/workbook.xml", xml);
  }
}

// Usuwa calcChain.xml WRAZ z jego wpisami w [Content_Types].xml i workbook.xml.rels.
async function dropCalcChain(zip) {
  if (!zip.file("xl/calcChain.xml")) return;
  zip.remove("xl/calcChain.xml");

  const ctFile = zip.file("[Content_Types].xml");
  if (ctFile) {
    let ct = await ctFile.async("string");
    ct = ct.replace(/<Override\b[^>]*calcChain\.xml[^>]*\/>\s*/gi, "");
    zip.file("[Content_Types].xml", ct);
  }

  const relsFile = zip.file("xl/_rels/workbook.xml.rels");
  if (relsFile) {
    let rels = await relsFile.async("string");
    rels = rels.replace(/<Relationship\b[^>]*calcChain\.xml[^>]*\/>\s*/gi, "");
    zip.file("xl/_rels/workbook.xml.rels", rels);
  }
}

// Zapis XML arkusza jako bajty UTF-8 (TextEncoder), NIE jako JS string do JSZip.
// JSZip przy zip.file(path, string) potrafi zapisać emoji jako CESU-8 (ed a0 bd…),
// a Excel wtedy żąda „odzyskania skoroszytu". TextEncoder → poprawne f0 9f ….
function zipWriteUtf8Xml(zip, path, xmlString) {
  const bytes = typeof TextEncoder !== "undefined"
    ? new TextEncoder().encode(String(xmlString))
    : xmlString;
  zip.file(path, bytes);
}

// ─────────────────────────────────────────────────────────────────────────────
// NOTATKI KOMÓREK (cell-notes.js) → xl/commentsN.xml + rysunek VML, stringowo.
//
// Notatka Excela to DWIE części: tekst w commentsN.xml (<comment ref="A1">) i kształt
// w legacy VML (<v:shape> z <x:Row>/<x:Column>) — bez kształtu Excel notatki nie pokaże.
// Edycja niesie pendingEdits[arkusz][A1].note = { text, author } | null (null = usuń).
// Zmieniamy tylko wpisy dotkniętych komórek; reszta obu części zostaje bajt-identyczna
// (cudze kształty w VML, np. pola wyboru, też). Nowe części dostają rel arkusza,
// wpis w [Content_Types].xml i <legacyDrawing> w arkuszu (w kolejności ze schematu).
// Komentarzy WĄTKOWYCH (threadedComments, Excel 365) nie dotykamy — aplikacja pokazuje
// je tylko do odczytu, a ich legacy-kopia w comments musi zostać zgodna z wątkiem.
const REL_TYPE_COMMENTS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";
const REL_TYPE_VML = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing";
const CT_COMMENTS = "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml";
const CT_VML = "application/vnd.openxmlformats-officedocument.vmlDrawing";
const VML_NS = { v: "urn:schemas-microsoft-com:vml", o: "urn:schemas-microsoft-com:office:office", x: "urn:schemas-microsoft-com:office:excel" };
// Elementy arkusza, które w schemacie stoją PO <legacyDrawing> (wstawiamy przed pierwszym z nich).
const AFTER_LEGACY_DRAWING = ["legacyDrawingHF", "drawingHF", "picture", "oleObjects", "controls", "webPublishItems", "tableParts", "extLst", "AlternateContent"];

function xmlAttr(tag, name) {
  const m = String(tag).match(new RegExp(`\\s${name}="([^"]*)"`));
  return m ? m[1] : null;
}

function parseRelsList(xml) {
  const out = [];
  const re = /<Relationship\b[^>]*?\/?>/g;
  let m;
  while ((m = re.exec(xml || ""))) {
    out.push({ id: xmlAttr(m[0], "Id"), type: xmlAttr(m[0], "Type") || "", target: xmlAttr(m[0], "Target") || "", external: xmlAttr(m[0], "TargetMode") === "External", raw: m[0] });
  }
  return out;
}

// Target relacji → ścieżka w ZIP (względem katalogu części źródłowej albo od korzenia).
function resolvePartPath(baseDir, target) {
  const parts = target.startsWith("/") ? [] : baseDir.split("/").filter(Boolean);
  target.replace(/^\//, "").split("/").forEach((seg) => {
    if (seg === "..") parts.pop();
    else if (seg && seg !== ".") parts.push(seg);
  });
  return parts.join("/");
}

function relativePartTarget(fromDir, path) {
  const from = fromDir.split("/").filter(Boolean);
  const to = path.split("/").filter(Boolean);
  let i = 0;
  while (i < from.length && i < to.length - 1 && from[i] === to[i]) i++;
  return "../".repeat(from.length - i) + to.slice(i).join("/");
}

function nextFreePath(zip, pattern) {
  for (let n = 1; n < 10000; n++) {
    const p = pattern.replace("{n}", String(n));
    if (!zip.file(p)) return p;
  }
  throw new Error("brak wolnej nazwy części");
}

function cellRowCol(ref) {
  return { row: rowNum(ref) - 1, col: colIndex0(ref) };
}

function compareRefs(a, b) {
  const ra = cellRowCol(a);
  const rb = cellRowCol(b);
  return ra.row - rb.row || ra.col - rb.col;
}

function noteTextXml(P, text) {
  const clean = sanitizeXmlText(String(text ?? "").replace(/\r\n?/g, "\n"));
  return `<${P}text><${P}t xml:space="preserve">${escapeXmlText(clean)}</${P}t></${P}text>`;
}

// commentsN.xml: { xml } albo { xml: null } gdy lista notatek opustoszała (część do usunięcia).
// notes: { A1: { text, author } | null }, skip: Set refów z komentarzem wątkowym.
function patchCommentsXml(xml, notes, skip) {
  let doc = xml || `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<comments xmlns="${SPREADSHEETML_NS}"><authors/><commentList/></comments>`;
  const P = (doc.match(/<((?:[\w-]+:)?)comments\b/) || [])[1] || "";

  // autorzy: istniejące wpisy zostają co do bajtu, nowi dopisywani na końcu
  const authorsRe = new RegExp(`<${P}authors\\b([^>]*?)(?:\\/>|>([\\s\\S]*?)<\\/${P}authors>)`);
  const am = doc.match(authorsRe);
  const authorItems = [];
  if (am && am[2]) {
    const itemRe = new RegExp(`<${P}author\\b[^>]*?(?:\\/>|>([\\s\\S]*?)<\\/${P}author>)`, "g");
    let m;
    while ((m = itemRe.exec(am[2]))) authorItems.push({ raw: m[0], name: decodeXmlEntitiesSimple(m[1] || "") });
  }
  const baseAuthors = authorItems.length;
  const authorId = (name) => {
    const n = String(name || "");
    let i = authorItems.findIndex((a) => a.name === n);
    if (i < 0) {
      authorItems.push({ raw: `<${P}author>${escapeXmlText(sanitizeXmlText(n))}</${P}author>`, name: n });
      i = authorItems.length - 1;
    }
    return i;
  };

  const listRe = new RegExp(`<${P}commentList\\b([^>]*?)(?:\\/>|>([\\s\\S]*?)<\\/${P}commentList>)`);
  const lm = doc.match(listRe);
  const items = new Map();
  if (lm && lm[2]) {
    const itemRe = new RegExp(`<${P}comment\\b([^>]*?)(?:\\/>|>([\\s\\S]*?)<\\/${P}comment>)`, "g");
    let m;
    while ((m = itemRe.exec(lm[2]))) {
      const ref = xmlAttr(` ${m[1]}`, "ref");
      if (ref) items.set(ref, { raw: m[0], attrs: m[1] });
    }
  }

  Object.keys(notes || {}).forEach((ref) => {
    if (skip && skip.has(ref)) return;
    const note = notes[ref];
    const cur = items.get(ref);
    if (!note) { items.delete(ref); return; }
    if (cur) {
      // edycja: atrybuty (autor, xr:uid…) zostają, zmienia się sam tekst
      const attrs = cur.attrs.replace(/\/\s*$/, "");
      items.set(ref, { raw: `<${P}comment${attrs}>${noteTextXml(P, note.text)}</${P}comment>`, attrs });
    } else {
      items.set(ref, { raw: `<${P}comment ref="${ref}" authorId="${authorId(note.author)}" shapeId="0">${noteTextXml(P, note.text)}</${P}comment>` });
    }
  });

  if (!items.size) return { xml: null };
  const listBody = Array.from(items.keys()).sort(compareRefs).map((ref) => items.get(ref).raw).join("");
  const listTag = `<${P}commentList${lm ? lm[1].replace(/\/\s*$/, "") : ""}>${listBody}</${P}commentList>`;
  doc = lm ? doc.replace(listRe, () => listTag) : doc.replace(new RegExp(`<\\/${P}comments>`), (end) => listTag + end);
  if (authorItems.length !== baseAuthors || !am) {
    const authorsTag = `<${P}authors${am ? am[1].replace(/\/\s*$/, "") : ""}>${authorItems.map((a) => a.raw).join("")}</${P}authors>`;
    doc = am ? doc.replace(authorsRe, () => authorsTag) : doc.replace(/<((?:[\w-]+:)?)comments\b[^>]*>/, (open) => open + authorsTag);
  }
  return { xml: doc };
}

function vmlShapeXml(pf, id, row, col) {
  const { v, o, x } = pf;
  const top = Math.max(0, row - 1);
  return `<${v}:shape id="_x0000_s${id}" type="#_x0000_t202" style="position:absolute;margin-left:59.25pt;margin-top:1.5pt;width:108pt;height:59.25pt;z-index:1;visibility:hidden" fillcolor="#ffffe1" ${o}:insetmode="auto">`
    + `<${v}:fill color2="#ffffe1"/><${v}:shadow on="t" color="black" obscured="t"/><${v}:path ${o}:connecttype="none"/>`
    + `<${v}:textbox style="mso-direction-alt:auto"><div style="text-align:left"></div></${v}:textbox>`
    + `<${x}:ClientData ObjectType="Note"><${x}:MoveWithCells/><${x}:SizeWithCells/>`
    + `<${x}:Anchor>${col + 1}, 15, ${top}, 10, ${col + 3}, 15, ${top + 4}, 4</${x}:Anchor>`
    + `<${x}:AutoFill>False</${x}:AutoFill><${x}:Row>${row}</${x}:Row><${x}:Column>${col}</${x}:Column></${x}:ClientData></${v}:shape>`;
}

// Rysunek VML notatek: kształt dla każdej notatki (dodaj brakujący), usunięte notatki
// tracą swój kształt. Prefiksy przestrzeni nazw bierzemy z pliku (openpyxl pisze ns0/ns1/ns2).
// Zwraca { xml, shapes } — shapes = ile kształtów zostało (0 → część do usunięcia).
function patchVmlXml(xml, notes, skip, idmap) {
  let doc = xml || `<xml xmlns:v="${VML_NS.v}" xmlns:o="${VML_NS.o}" xmlns:x="${VML_NS.x}">`
    + `<o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="${idmap}"/></o:shapelayout></xml>`;
  const rootM = doc.match(/<((?:[\w-]+:)?)xml\b[^>]*>/);
  if (!rootM) throw new Error("VML bez korzenia <xml>");
  const pf = {};
  const declared = new Set();
  doc.replace(/xmlns:([\w-]+)="([^"]*)"/g, (all, p, uri) => {
    declared.add(p);
    Object.keys(VML_NS).forEach((k) => { if (VML_NS[k] === uri && !pf[k]) pf[k] = p; });
    return all;
  });
  const extraNs = [];
  Object.keys(VML_NS).forEach((k) => {
    if (pf[k]) return;
    let p = k;
    while (declared.has(p)) p += "x";
    pf[k] = p;
    declared.add(p);
    extraNs.push(` xmlns:${p}="${VML_NS[k]}"`);
  });
  if (extraNs.length) doc = doc.replace(rootM[0], rootM[0].replace(/\s*\/?>$/, (end) => extraNs.join("") + end));

  const { v, x } = pf;
  const shapeRe = new RegExp(`\\s*<${v}:shape\\b[^>]*?(?:\\/>|>[\\s\\S]*?<\\/${v}:shape>)`, "g");
  const rowRe = new RegExp(`<${x}:Row>\\s*(\\d+)\\s*<\\/${x}:Row>`);
  const colRe = new RegExp(`<${x}:Column>\\s*(\\d+)\\s*<\\/${x}:Column>`);
  const noteShapeKey = (s) => {
    if (!/ObjectType="Note"/.test(s)) return null;
    const r = s.match(rowRe);
    const c = s.match(colRe);
    return r && c ? `${r[1]}:${c[1]}` : null;
  };

  const want = new Map(); // "row:col" → true (ma mieć kształt) | false (bez kształtu)
  Object.keys(notes || {}).forEach((ref) => {
    if (skip && skip.has(ref)) return;
    const { row, col } = cellRowCol(ref);
    want.set(`${row}:${col}`, !!notes[ref]);
  });

  const present = new Set();
  doc = doc.replace(shapeRe, (s) => {
    const key = noteShapeKey(s);
    if (key && want.get(key) === false) return "";
    if (key) present.add(key);
    return s;
  });

  const missing = Array.from(want.entries()).filter(([key, on]) => on && !present.has(key)).map(([key]) => key);
  if (missing.length) {
    if (!/id="_x0000_t202"/.test(doc)) {
      const st = `<${v}:shapetype id="_x0000_t202" coordsize="21600,21600" ${pf.o}:spt="202" path="m,l,21600r21600,l21600,xe"><${v}:stroke joinstyle="miter"/><${v}:path gradientshapeok="t" ${pf.o}:connecttype="rect"/></${v}:shapetype>`;
      doc = doc.replace(new RegExp(`<\\/${rootM[1]}xml>\\s*$`), (end) => st + end);
    }
    let maxId = 0;
    doc.replace(/_x0000_s(\d+)/g, (all, n) => { maxId = Math.max(maxId, Number(n)); return all; });
    const block = Number((doc.match(/\bdata="(\d+)/) || [])[1] || idmap || 1);
    let next = Math.max(maxId + 1, block * 1024 + 1);
    const added = missing.sort((a, b) => {
      const [ra, ca] = a.split(":").map(Number);
      const [rb, cb] = b.split(":").map(Number);
      return ra - rb || ca - cb;
    }).map((key) => {
      const [row, col] = key.split(":").map(Number);
      return vmlShapeXml(pf, next++, row, col);
    }).join("");
    doc = doc.replace(new RegExp(`<\\/${rootM[1]}xml>\\s*$`), (end) => added + end);
  }

  const shapes = (doc.match(new RegExp(`<${v}:shape\\b`, "g")) || []).length;
  return { xml: doc, shapes };
}

// <legacyDrawing r:id="…"/> w arkuszu: wstawienie w miejscu ze schematu / usunięcie.
function sheetRelPrefix(xml) {
  const m = xml.match(new RegExp(`xmlns:([\\w-]+)="${OOXML_REL_NS.replace(/[.]/g, "\\.")}"`));
  return m ? m[1] : null;
}

function insertLegacyDrawing(xml, rid) {
  const rootM = xml.match(/<((?:[\w-]+:)?)worksheet\b[^>]*>/);
  if (!rootM) return xml;
  const P = rootM[1];
  let out = xml;
  let rp = sheetRelPrefix(out);
  if (!rp) {
    rp = "r";
    while (new RegExp(`xmlns:${rp}=`).test(out)) rp += "r";
    out = out.replace(rootM[0], rootM[0].replace(/\s*>$/, ` xmlns:${rp}="${OOXML_REL_NS}">`));
  }
  const tag = `<${P}legacyDrawing ${rp}:id="${rid}"/>`;
  const afterData = Math.max(0, out.search(new RegExp(`<\\/${P}sheetData>|<${P}sheetData\\s*\\/>`)));
  let at = -1;
  AFTER_LEGACY_DRAWING.forEach((name) => {
    const re = new RegExp(`<(?:[\\w-]+:)?${name}\\b`, "g");
    re.lastIndex = afterData;
    const m = re.exec(out);
    if (m && (at < 0 || m.index < at)) at = m.index;
  });
  if (at < 0) at = out.search(new RegExp(`<\\/${P}worksheet>\\s*$`));
  return out.slice(0, at) + tag + out.slice(at);
}

function removeLegacyDrawing(xml) {
  return xml.replace(/<(?:[\w-]+:)?legacyDrawing\b[^>]*?\/>|<((?:[\w-]+:)?)legacyDrawing\b[^>]*>[\s\S]*?<\/\1legacyDrawing>/, "");
}

function ctEnsureDefault(ct, ext, type) {
  if (new RegExp(`<Default\\b[^>]*Extension="${ext}"`, "i").test(ct)) return ct;
  return ct.replace(/<Types\b[^>]*>/, (open) => `${open}<Default Extension="${ext}" ContentType="${type}"/>`);
}
function ctAddOverride(ct, partName, type) {
  if (ct.includes(`PartName="${partName}"`)) return ct;
  return ct.replace(/<\/Types>\s*$/, (end) => `<Override PartName="${partName}" ContentType="${type}"/>${end}`);
}
function ctRemoveOverride(ct, partName) {
  return ct.replace(new RegExp(`<Override\\b[^>]*PartName="${partName.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}"[^>]*\\/>`), "");
}
function relsAdd(xml, type, target) {
  const ids = new Set(parseRelsList(xml).map((r) => r.id));
  let n = 1;
  while (ids.has(`rId${n}`)) n++;
  const id = `rId${n}`;
  return { id, xml: xml.replace(/<\/Relationships>\s*$/, (end) => `<Relationship Id="${id}" Type="${type}" Target="${target}"/>${end}`) };
}
function relsRemove(xml, id) {
  return xml.replace(new RegExp(`<Relationship\\b[^>]*\\bId="${id}"[^>]*\\/>`), "");
}

// Nanosi notatki jednego arkusza. Zwraca nowy XML arkusza (zmienia się tylko <legacyDrawing>).
async function applySheetNotes(zip, sheetPath, sheetXml, notes) {
  const dir = sheetPath.slice(0, sheetPath.lastIndexOf("/"));
  const base = sheetPath.slice(sheetPath.lastIndexOf("/") + 1);
  const relsPath = `${dir}/_rels/${base}.rels`;
  const relsFile = zip.file(relsPath);
  let rels = relsFile ? await relsFile.async("string")
    : `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`;
  const relList = parseRelsList(rels);
  const ctFile = zip.file("[Content_Types].xml");
  let ct = ctFile ? await ctFile.async("string") : null;
  let sheet = sheetXml;

  // komentarze wątkowe — tych komórek nie ruszamy
  const skip = new Set();
  for (const r of relList) {
    if (r.external || !/\/threadedComment$/i.test(r.type)) continue;
    const f = zip.file(resolvePartPath(dir, r.target));
    if (!f) continue;
    (await f.async("string")).replace(/<(?:[\w-]+:)?threadedComment\b[^>]*\bref="([A-Z]+\d+)"/g, (all, ref) => { skip.add(ref); return all; });
  }

  // tekst notatek
  const cRel = relList.find((r) => !r.external && r.type === REL_TYPE_COMMENTS);
  const cPath = cRel ? resolvePartPath(dir, cRel.target) : null;
  const cFile = cPath ? zip.file(cPath) : null;
  const cOut = patchCommentsXml(cFile ? await cFile.async("string") : null, notes, skip);
  if (cOut.xml) {
    const path = cFile ? cPath : nextFreePath(zip, "xl/comments{n}.xml");
    zipWriteUtf8Xml(zip, path, cOut.xml);
    if (!cFile) {
      if (cRel) rels = relsRemove(rels, cRel.id); // rel do nieistniejącej części
      rels = relsAdd(rels, REL_TYPE_COMMENTS, relativePartTarget(dir, path)).xml;
      if (ct) ct = ctAddOverride(ct, `/${path}`, CT_COMMENTS);
    }
  } else if (cRel) {
    if (cFile) zip.remove(cPath);
    rels = relsRemove(rels, cRel.id);
    if (ct && cPath) ct = ctRemoveOverride(ct, `/${cPath}`);
  }

  // kształty VML
  const ldRid = (sheet.match(/<(?:[\w-]+:)?legacyDrawing\b[^>]*?\b(?:[\w-]+:)?id="([^"]+)"/) || [])[1] || null;
  const vRel = ldRid ? relList.find((r) => r.id === ldRid && !r.external) : null;
  const vPath = vRel ? resolvePartPath(dir, vRel.target) : null;
  const vFile = vPath ? zip.file(vPath) : null;
  const anyNote = Object.values(notes || {}).some(Boolean);
  if (vFile || anyNote) {
    let idmap = 1;
    if (!vFile) {
      // nowy rysunek dostaje własny blok identyfikatorów kształtów (o:idmap) za wszystkimi innymi
      for (const p of Object.keys(zip.files)) {
        if (!/\.vml$/i.test(p)) continue;
        const xml = await zip.file(p).async("string");
        (xml.match(/\bdata="([\d,\s]+)"/) || ["", ""])[1].split(",").forEach((n) => { if (+n >= idmap) idmap = +n + 1; });
      }
    }
    const vOut = patchVmlXml(vFile ? await vFile.async("string") : null, notes, skip, idmap);
    if (vOut.shapes > 0) {
      const path = vFile ? vPath : nextFreePath(zip, "xl/drawings/vmlDrawing{n}.vml");
      zipWriteUtf8Xml(zip, path, vOut.xml);
      if (!vFile) {
        if (ldRid) { sheet = removeLegacyDrawing(sheet); if (vRel) rels = relsRemove(rels, vRel.id); }
        const added = relsAdd(rels, REL_TYPE_VML, relativePartTarget(dir, path));
        rels = added.xml;
        sheet = insertLegacyDrawing(sheet, added.id);
        if (ct) ct = ctEnsureDefault(ct, "vml", CT_VML);
      }
    } else if (vFile) {
      zip.remove(vPath);
      rels = relsRemove(rels, vRel.id);
      sheet = removeLegacyDrawing(sheet);
      if (ct) ct = ctRemoveOverride(ct, `/${vPath}`);
      const vRelsPath = `${vPath.slice(0, vPath.lastIndexOf("/"))}/_rels/${vPath.slice(vPath.lastIndexOf("/") + 1)}.rels`;
      if (zip.file(vRelsPath)) zip.remove(vRelsPath);
    }
  }

  if (parseRelsList(rels).length) zipWriteUtf8Xml(zip, relsPath, rels);
  else if (relsFile) zip.remove(relsPath);
  if (ct && ctFile) zipWriteUtf8Xml(zip, "[Content_Types].xml", ct);
  return sheet;
}

async function buildPatchedXlsx(originalBytes, edits, formulaMaps = {}) {
  if (typeof JSZip === "undefined") throw new Error("JSZip niedostępny");
  const zip = await JSZip.loadAsync(originalBytes);
  const sheetPaths = await resolveSheetPaths(zip);

  // Pisarz stylów tylko gdy jakaś edycja zapisuje datę albo zmienia format
  // (zwykły zapis tekstu nie rusza styles.xml).
  const needsStyles = Object.values(edits || {}).some((cells) => cells && Object.values(cells).some((p) => p && (p.t === "d" || p.fmt)));
  let dateStyler = null;
  let stylesPath = null;
  let stylesXml = null;
  if (needsStyles) {
    stylesPath = await resolveStylesPath(zip);
    if (stylesPath) {
      stylesXml = await zip.file(stylesPath).async("string");
      dateStyler = createStyleWriter(stylesXml) || createDateStyler(stylesXml);
    }
  }

  let touched = false;
  for (const sheetName of Object.keys(edits || {})) {
    const cells = edits[sheetName];
    if (!cells || Object.keys(cells).length === 0) continue;
    const path = sheetPaths[sheetName];
    if (!path || !zip.file(path)) continue;
    // Sama notatka nie dotyka <sheetData> (ani formuł dzielonych w tej komórce).
    const valueCells = {};
    const notes = {};
    Object.keys(cells).forEach((ref) => {
      const p = cells[ref];
      if (p && "note" in p) notes[ref] = p.note;
      if (p === null || (p && ("t" in p || p.fmt || p.clear))) valueCells[ref] = p;
    });
    const xml = await zip.file(path).async("string");
    let patched = xml;
    if (Object.keys(valueCells).length) {
      patched = patchSheetXml(xml, valueCells, (formulaMaps && formulaMaps[sheetName]) || null, dateStyler);
      touched = true;
    }
    if (Object.keys(notes).length) patched = await applySheetNotes(zip, path, patched, notes);
    if (patched !== xml) zipWriteUtf8Xml(zip, path, patched);
  }

  if (dateStyler && dateStyler.changed) zipWriteUtf8Xml(zip, stylesPath, dateStyler.apply(stylesXml));

  if (touched) {
    await forceRecalcOnLoad(zip);
    await dropCalcChain(zip);
  }

  return zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
}

// Eksport do testów Node/Playwright (poza przeglądarką).
if (typeof module !== "undefined" && module.exports) {
  module.exports = {
    patchSheetXml,
    patchSheetDataInner,
    splitSheetData,
    indexCells,
    sanitizeXmlText,
    buildCellXml,
    zipWriteUtf8Xml,
    createDateStyler,
    createStyleWriter,
    isDateFormatCode,
    patchCommentsXml,
    patchVmlXml,
    insertLegacyDrawing,
    applySheetNotes,
    buildPatchedXlsx,
  };
}
