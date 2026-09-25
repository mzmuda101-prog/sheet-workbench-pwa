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

function normalizeCellAttrs(attrs, payload) {
  // [EN] Strip stray "/" left if a self-closing cell was ever mis-parsed as open.
  let a = String(attrs || "").replace(/\/\s*$/g, "").replace(/\s*\bt="[^"]*"/g, "").trim();
  if (payload && payload.t === "s") a = (a ? a + " " : "") + 't="inlineStr"';
  return a ? " " + a : "";
}

function buildCellXml(ref, payload, attrs, dateStyler) {
  // Data zapisuje się jako liczba seryjna — bez formatu daty w stylu Excel pokazałby
  // „45424”. Styler podmienia s= na wariant tego samego stylu z formatem daty.
  if (dateStyler && payload.t === "d") attrs = withDateStyle(attrs, dateStyler);
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
      if (existing) replacements.push({ start: existing.start, len: existing.full.length, text: "" });
      return;
    }

    const cellXml = buildCellXml(ref, payload, existing ? existing.attrs : "", dateStyler);
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

async function buildPatchedXlsx(originalBytes, edits, formulaMaps = {}) {
  if (typeof JSZip === "undefined") throw new Error("JSZip niedostępny");
  const zip = await JSZip.loadAsync(originalBytes);
  const sheetPaths = await resolveSheetPaths(zip);

  // Styler dat tylko gdy jakaś edycja zapisuje datę (zwykły zapis tekstu nie rusza styles.xml).
  const hasDateEdit = Object.values(edits || {}).some((cells) => cells && Object.values(cells).some((p) => p && p.t === "d"));
  let dateStyler = null;
  let stylesPath = null;
  let stylesXml = null;
  if (hasDateEdit) {
    stylesPath = await resolveStylesPath(zip);
    if (stylesPath) {
      stylesXml = await zip.file(stylesPath).async("string");
      dateStyler = createDateStyler(stylesXml);
    }
  }

  let touched = false;
  for (const sheetName of Object.keys(edits || {})) {
    const cells = edits[sheetName];
    if (!cells || Object.keys(cells).length === 0) continue;
    const path = sheetPaths[sheetName];
    if (!path || !zip.file(path)) continue;
    const xml = await zip.file(path).async("string");
    const patched = patchSheetXml(xml, cells, (formulaMaps && formulaMaps[sheetName]) || null, dateStyler);
    zipWriteUtf8Xml(zip, path, patched);
    touched = true;
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
    isDateFormatCode,
  };
}
