// =====================================================================
// ZIP-patch zapisu XLSX — zachowuje oryginalny plik (tabele, wykresy, style,
// formuły, rysunki) i podmienia TYLKO wartości edytowanych komórek w sheet.xml.
//
// Dlaczego: xlsx-js-style (jak SheetJS) przy XLSX.write gubi tabele Excela,
// wykresy, rysunki i calcChain — co psuje formuły z odwołaniami strukturalnymi
// (np. SUBTOTAL(9,Dane[Kwota])) i formatowanie tabel. Zamiast przepisywać plik
// od zera, rozpakowujemy oryginalny ZIP (JSZip), nanosimy edycje na XML arkusza
// i pakujemy z powrotem — reszta pliku zostaje nietknięta.
// =====================================================================

const SPREADSHEETML_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const OOXML_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

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

function findRowElement(sheetData, rowNum) {
  const rows = sheetData.getElementsByTagName("row");
  for (let i = 0; i < rows.length; i++) {
    if (parseInt(rows[i].getAttribute("r"), 10) === rowNum) return rows[i];
  }
  return null;
}

// Wstawia <row> w sheetData zachowując rosnącą kolejność r.
function insertRowInOrder(doc, sheetData, rowNum) {
  const row = doc.createElementNS(SPREADSHEETML_NS, "row");
  row.setAttribute("r", String(rowNum));
  const rows = sheetData.getElementsByTagName("row");
  let ref = null;
  for (let i = 0; i < rows.length; i++) {
    if (parseInt(rows[i].getAttribute("r"), 10) > rowNum) { ref = rows[i]; break; }
  }
  sheetData.insertBefore(row, ref);
  return row;
}

// Wstawia <c> w wierszu zachowując rosnącą kolejność kolumn.
function insertCellInOrder(doc, rowEl, cellRef, colIndex0) {
  const c = doc.createElementNS(SPREADSHEETML_NS, "c");
  c.setAttribute("r", cellRef);
  const cells = rowEl.getElementsByTagName("c");
  let ref = null;
  for (let i = 0; i < cells.length; i++) {
    const otherCol = XLSX.utils.decode_cell(cells[i].getAttribute("r")).c;
    if (otherCol > colIndex0) { ref = cells[i]; break; }
  }
  rowEl.insertBefore(c, ref);
  return c;
}

// Nanosi wartość/typ na element <c>, zachowując atrybut stylu (s) i usuwając
// starą wartość/formułę.
function applyValueToCell(doc, cEl, payload) {
  // usuń dotychczasowe dzieci (v, f, is) i atrybut typu
  Array.from(cEl.childNodes).forEach((n) => cEl.removeChild(n));
  cEl.removeAttribute("t");

  if (payload.t === "s") {
    // tekst jako inline string — nie ruszamy sharedStrings
    cEl.setAttribute("t", "inlineStr");
    const is = doc.createElementNS(SPREADSHEETML_NS, "is");
    const tEl = doc.createElementNS(SPREADSHEETML_NS, "t");
    tEl.setAttribute("xml:space", "preserve");
    tEl.appendChild(doc.createTextNode(String(payload.v)));
    is.appendChild(tEl);
    cEl.appendChild(is);
    return;
  }

  // liczba lub data → wartość numeryczna (format wyświetlania pochodzi ze stylu .s)
  let num;
  if (payload.t === "d" && payload.v instanceof Date) num = dateToExcelSerial(payload.v);
  else num = Number(payload.v);
  const vEl = doc.createElementNS(SPREADSHEETML_NS, "v");
  vEl.appendChild(doc.createTextNode(String(num)));
  cEl.appendChild(vEl);
}

// Nanosi wszystkie edycje na jeden arkusz (string XML -> string XML).
function patchSheetXml(xml, cells) {
  const doc = new DOMParser().parseFromString(xml, "application/xml");
  const sheetData = doc.getElementsByTagName("sheetData")[0];
  if (!sheetData) return xml;

  // mapa istniejących komórek ref -> <c>
  const cellIndex = {};
  Array.from(sheetData.getElementsByTagName("c")).forEach((c) => {
    const r = c.getAttribute("r");
    if (r) cellIndex[r] = c;
  });

  Object.keys(cells).forEach((ref) => {
    const payload = cells[ref];
    let cEl = cellIndex[ref];

    if (payload === null) {
      // usunięcie wartości
      if (cEl && cEl.parentNode) cEl.parentNode.removeChild(cEl);
      return;
    }

    if (!cEl) {
      const { r: rowIdx0, c: colIdx0 } = XLSX.utils.decode_cell(ref);
      const rowNum = rowIdx0 + 1;
      let rowEl = findRowElement(sheetData, rowNum);
      if (!rowEl) rowEl = insertRowInOrder(doc, sheetData, rowNum);
      cEl = insertCellInOrder(doc, rowEl, ref, colIdx0);
    }
    applyValueToCell(doc, cEl, payload);
  });

  let out = new XMLSerializer().serializeToString(doc);
  // zachowaj deklarację XML, jeśli oryginał ją miał (serializer ją pomija)
  if (/^\s*<\?xml/.test(xml) && !/^\s*<\?xml/.test(out)) {
    out = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' + out;
  }
  return out;
}

// Wymusza pełne przeliczenie formuł przy otwarciu (Excel pokaże aktualne wyniki
// po zmianie wartości wejściowych). NIE usuwamy calcChain.xml — edytujemy tylko
// wartości, więc lista formuł pozostaje poprawna; usunięcie calcChain bez usunięcia
// jego wpisów w [Content_Types].xml i workbook.xml.rels psuło plik ("naprawić?").
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
  // Brak <calcPr> → nie wstawiamy (kolejność elementów w workbook.xml jest ścisła);
  // Excel w trybie automatycznym i tak przeliczy zależne formuły przy otwarciu.
}

// Główna funkcja: oryginalne bajty + edycje -> nowe bajty XLSX (zachowany plik).
async function buildPatchedXlsx(originalBytes, edits) {
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
    zip.file(path, patchSheetXml(xml, cells));
    touched = true;
  }

  if (touched) await forceRecalcOnLoad(zip);

  return zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
}
