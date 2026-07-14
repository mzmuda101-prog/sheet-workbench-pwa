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
const CELL_BLOCK_RE = /<c r="([A-Z]+\d+)"([^>]*)>([\s\S]*?)<\/c>|<c r="([A-Z]+\d+)"([^/>]*)\/>/g;

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

// Usuwa znaki niedozwolone w XML 1.0 (tab/LF/CR zostają). Lone UTF-16 surrogates
// (bug XMLSerializer przy emoji) też wyrzucamy — inaczej Excel odmawia otwarcia.
function sanitizeXmlText(s) {
  return String(s)
    .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F\uFFFE\uFFFF]/g, "")
    .replace(/[\uD800-\uDFFF]/g, "");
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
      const newFull = m.cell.selfClose
        ? `<c r="${m.ref}"${m.cell.attrs}>${newInner}</c>`
        : `<c r="${m.ref}"${m.cell.attrs}>${newInner}</c>`;
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

function normalizeCellAttrs(attrs, payload) {
  let a = String(attrs || "").replace(/\s*\bt="[^"]*"/g, "").trim();
  if (payload && payload.t === "s") a = (a ? a + " " : "") + 't="inlineStr"';
  return a ? " " + a : "";
}

function buildCellXml(ref, payload, attrs) {
  const attrStr = normalizeCellAttrs(attrs, payload);
  if (payload.t === "s") {
    const t = escapeXmlText(sanitizeXmlText(payload.v));
    return `<c r="${ref}"${attrStr}><is><t xml:space="preserve">${t}</t></is></c>`;
  }
  let num;
  if (payload.t === "d" && payload.v instanceof Date) num = dateToExcelSerial(payload.v);
  else num = Number(payload.v);
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

function patchSheetDataInner(inner, cells, formulaMap) {
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
      if (existing) replacements.push({ start: existing.start, len: existing.full.length, text: "" });
      return;
    }

    const cellXml = buildCellXml(ref, payload, existing ? existing.attrs : "");
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
function patchSheetXml(xml, cells, formulaMap) {
  const parts = splitSheetData(xml);
  if (!parts) return xml;
  const patchedInner = patchSheetDataInner(parts.inner, cells, formulaMap);
  if (patchedInner === parts.inner) return xml;
  return parts.before + patchedInner + parts.after;
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
    zip.file("xl/workbook.xml", xml);
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

async function buildPatchedXlsx(originalBytes, edits, formulaMaps = {}) {
  if (typeof JSZip === "undefined") throw new Error("JSZip niedostępny");
  const zip = await JSZip.loadAsync(originalBytes);
  const sheetPaths = await resolveSheetPaths(zip);

  let touched = false;
  for (const sheetName of Object.keys(edits || {})) {
    const cells = edits[sheetName];
    if (!cells || Object.keys(cells).length === 0) continue;
    const path = sheetPaths[sheetName];
    if (!path || !zip.file(path)) continue;
    const xml = await zip.file(path).async("string");
    zip.file(path, patchSheetXml(xml, cells, (formulaMaps && formulaMaps[sheetName]) || null));
    touched = true;
  }

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
  };
}
