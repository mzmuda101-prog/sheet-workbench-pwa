/* ───────────────────────────────────────────────────────────────────────────
   Data Validation — listy/słowniki jak w Excelu (Sprawdzanie poprawności → Lista)
   FAZA 1: ODCZYT z pliku + pokazywanie. Parsujemy <dataValidation type="list">
   z surowego XML arkusza (ta sama ścieżka co conditional-formatting.js: JSZip +
   DOMParser, pętla po arkuszach). Z pliku bierzemy też 3 tryby rygoru zapisane
   w atrybucie errorStyle:  stop → blokuj,  warning → ostrzegaj,  information →
   tylko podpowiadaj.  Źródło wartości (<formula1> / x14:<xm:f>):
     • lista inline:   "Jan,Ola,Ewa"
     • zakres:         Słownik!$A$2:$A$50  (także z INNEGO arkusza tego pliku)
     • nazwany zakres: Imiona  (z workbook.xml <definedNames> → wb.Workbook.Names)
     • kolumna tabeli: Tabela3[Kolumna]  (structured ref → A1 przez currentTables)
   Cross-sheet listy Excel trzyma w rozszerzeniu <x14:dataValidation> (xm:f/xm:sqref),
   więc parsujemy oba warianty. Wartości rozwiązywane LENIWIE (przy pierwszej edycji
   komórki) i cache'owane na regule. Reużywa globalnych: cfSheetPathsFromRels,
   decodeXmlEntities, XLSX, JSZip, workbook, currentTables. ───────────────── */

let dvRulesBySheet = null; // Map<sheetName, Array<rule>>  (null = brak reguł)
let dvNameRefs = null;     // Map<definedName, refString>

function dvModeFromErrorStyle(es) {
  if (es === "warning") return "warning";
  if (es === "information") return "info";
  return "stop"; // brak/„stop" = blokuj (domyślne zachowanie Excela)
}

function dvDecodeRanges(sqref) {
  return String(sqref || "")
    .trim()
    .split(/\s+/)
    .filter(Boolean)
    .map((r) => { try { return XLSX.utils.decode_range(r.replace(/\$/g, "")); } catch { return null; } })
    .filter(Boolean);
}

// Parsuj reguły z XML JEDNEGO arkusza: standardowe <dataValidation> + rozszerzenie
// <x14:dataValidation> (źródła wskazujące inny arkusz). Tylko type="list".
function parseSheetDataValidations(sheetXml) {
  const out = [];
  let doc;
  try { doc = new DOMParser().parseFromString(sheetXml, "application/xml"); } catch { return out; }

  const pushFrom = (el, sqref, formula1Raw) => {
    const type = el.getAttribute("type") || "";
    if (type && type !== "list") return; // FAZA 1: tylko listy (dropdown)
    const ranges = dvDecodeRanges(sqref);
    if (!ranges.length) return;
    out.push({
      type: type || "list",
      ranges,
      mode: dvModeFromErrorStyle(el.getAttribute("errorStyle")),
      allowBlank: el.getAttribute("allowBlank") !== "0", // domyślnie dopuść pustą
      formula1Raw: formula1Raw || "",
      _values: null, // cache rozwiązanych wartości
    });
  };

  const std = doc.getElementsByTagName("dataValidation");
  for (let i = 0; i < std.length; i++) {
    const el = std[i];
    const f1 = el.getElementsByTagName("formula1")[0];
    pushFrom(el, el.getAttribute("sqref") || "", f1 ? f1.textContent : "");
  }

  const ext = doc.getElementsByTagName("x14:dataValidation");
  for (let i = 0; i < ext.length; i++) {
    const el = ext[i];
    const sqEl = el.getElementsByTagName("xm:sqref")[0];
    const fEl = el.getElementsByTagName("xm:f")[0];
    pushFrom(el, sqEl ? sqEl.textContent : "", fEl ? fEl.textContent : "");
  }

  return out;
}

// Rozwiń structured ref Excel (Tabela[Kolumna] / Tabela[[#Data],[Kolumna]]) → "Arkusz!A2:A145".
// Bez tego decode_range traktuje "Tabela3[…]" jak śmieciowy adres A1 i lista DV jest pusta.
function dvResolveStructuredToA1(ref) {
  const s = String(ref || "").trim();
  if (s.indexOf("[") < 0) return null;
  const m = s.match(/^([A-Za-z_À-ɏ][\w.À-ɏ]*)\[(.*)\]$/);
  if (!m) return null;
  const tables = typeof currentTables !== "undefined" ? currentTables : null;
  const t = tables && tables[m[1].toLowerCase()];
  if (!t || !t.ref || !t.sheetName || !t.columns) return null;
  let inner = String(m[2] || "").trim();
  let colName = inner;
  let mode = "data"; // data = bez nagłówka (domyślne Table[Col] w Excelu)
  const fancy = inner.match(/^\[#(All|Data|Headers)\]\s*,\s*\[([^\]]+)\]$/i);
  if (fancy) {
    mode = fancy[1].toLowerCase();
    colName = fancy[2].trim();
  } else if (/^\[.+\]$/.test(inner)) {
    colName = inner.slice(1, -1).trim();
  }
  const c = t.columns[colName.toLowerCase()];
  if (c == null) return null;
  let full;
  try { full = XLSX.utils.decode_range(t.ref); } catch { return null; }
  let r0 = full.s.r;
  let r1 = full.e.r;
  if (mode === "headers") {
    r1 = r0;
  } else {
    if (mode !== "all") r0 = full.s.r + 1; // pomiń wiersz nagłówka tabeli
    if (t.totalsRowShown) r1 = full.e.r - 1;
  }
  if (r1 < r0) return null;
  const a1 = XLSX.utils.encode_range({ s: { r: r0, c }, e: { r: r1, c } });
  const sn = String(t.sheetName);
  const quoted = /[\s'!]/.test(sn) ? `'${sn.replace(/'/g, "''")}'` : sn;
  return `${quoted}!${a1}`;
}

// Odczytaj wartości z zakresu (np. "Słownik!$A$2:$A$50" lub "$A$1:$A$9").
// Sheet w refie może wskazywać inny arkusz tego samego pliku.
function dvReadRangeValues(ref, defaultSheet) {
  if (typeof XLSX === "undefined" || !workbook) return [];
  let sheetName = defaultSheet;
  let rangePart = String(ref || "");
  const bang = rangePart.lastIndexOf("!");
  if (bang >= 0) {
    let sn = rangePart.slice(0, bang);
    rangePart = rangePart.slice(bang + 1);
    sn = sn.replace(/^'(.*)'$/, "$1").replace(/''/g, "'"); // 'Nazwa z spacją' → Nazwa z spacją
    sheetName = sn;
  }
  rangePart = rangePart.replace(/\$/g, "");
  const sheet = workbook.Sheets[sheetName];
  if (!sheet) return [];
  let range;
  try { range = XLSX.utils.decode_range(rangePart); } catch { return []; }
  const vals = [];
  for (let R = range.s.r; R <= range.e.r; R++) {
    for (let C = range.s.c; C <= range.e.c; C++) {
      const cell = sheet[XLSX.utils.encode_cell({ r: R, c: C })];
      if (!cell) continue;
      const disp = cell.w != null ? cell.w : (cell.v != null ? cell.v : "");
      const s = String(disp).trim();
      if (s !== "") vals.push(s);
      if (vals.length >= 5000) return vals; // bezpiecznik
    }
  }
  return vals;
}

// Rozwiąż listę dozwolonych wartości reguły (leniwie + cache). Dedup zachowuje kolejność.
function dvResolveValues(rule, defaultSheet) {
  if (rule._values) return rule._values;
  let raw = [];
  const f = String(rule.formula1Raw || "").trim();
  if (f && f !== "#N/A" && f !== "#REF!") {
    if (f.startsWith('"') && f.endsWith('"')) {
      raw = f.slice(1, -1).split(",").map((s) => s.trim()).filter((s) => s !== "");
    } else if (dvNameRefs && dvNameRefs.has(f)) {
      let resolved = dvNameRefs.get(f);
      const expanded = dvResolveStructuredToA1(resolved);
      if (expanded) resolved = expanded;
      raw = dvReadRangeValues(resolved, defaultSheet);
    } else {
      let resolved = f;
      const expanded = dvResolveStructuredToA1(f);
      if (expanded) resolved = expanded;
      raw = dvReadRangeValues(resolved, defaultSheet);
    }
  }
  const seen = new Set();
  const dedup = [];
  for (const v of raw) {
    const k = String(v).toLowerCase();
    if (seen.has(k)) continue;
    seen.add(k);
    dedup.push(v);
    if (dedup.length >= 2000) break;
  }
  rule._values = dedup;
  return dedup;
}

// Publiczny lookup: zwróć regułę listy dla komórki (r,c absolutne w arkuszu) albo null.
// { values:[…], mode:"stop"|"warning"|"info", allowBlank:bool }.
function getCellDataValidation(sheetName, r, c) {
  if (!dvRulesBySheet || !dvRulesBySheet.has(sheetName)) return null;
  const rules = dvRulesBySheet.get(sheetName);
  for (const rule of rules) {
    if (rule.type !== "list") continue;
    const hit = rule.ranges.some((rg) => r >= rg.s.r && r <= rg.e.r && c >= rg.s.c && c <= rg.e.c);
    if (!hit) continue;
    const values = dvResolveValues(rule, sheetName);
    if (!values.length) continue;
    return { values, mode: rule.mode, allowBlank: rule.allowBlank };
  }
  return null;
}

// Wczytaj reguły z całego skoroszytu (wołane przy otwarciu pliku, obok CF).
async function buildDataValidations(bytes, wb) {
  dvRulesBySheet = null;
  dvNameRefs = new Map();
  if (typeof JSZip === "undefined" || !bytes || !wb) return;
  try {
    const names = wb.Workbook && wb.Workbook.Names;
    if (Array.isArray(names)) {
      names.forEach((n) => { if (n && n.Name && n.Ref && !dvNameRefs.has(n.Name)) dvNameRefs.set(n.Name, n.Ref); });
    }
    const zip = await JSZip.loadAsync(bytes);
    const wbXml = await zip.file("xl/workbook.xml")?.async("string");
    const relsXml = await zip.file("xl/_rels/workbook.xml.rels")?.async("string");
    if (!wbXml || !relsXml) return;
    const ridToPath = cfSheetPathsFromRels(relsXml);
    const map = new Map();
    const sheetRe = /<sheet\b[^>]*?name="([^"]+)"[^>]*?r:id="([^"]+)"[^>]*?\/?>/g;
    let sm;
    while ((sm = sheetRe.exec(wbXml))) {
      const name = decodeXmlEntities(sm[1]);
      const path = ridToPath.get(sm[2]);
      const file = path ? zip.file(path) : null;
      if (!file) continue;
      const xml = await file.async("string");
      if (xml.indexOf("dataValidation") === -1) continue;
      const rules = parseSheetDataValidations(xml);
      if (rules.length) map.set(name, rules);
    }
    dvRulesBySheet = map.size ? map : null;
  } catch {
    dvRulesBySheet = null;
  }
}

// FAZA 2 — dodaj ręczną regułę DV dla bieżącego arkusza.
// values: string[] dozwolonych wartości (już odczytane przez wywołującego).
// colIdx: 0-bazowany indeks kolumny w currentHeaders (mapowany na absolutny kolumny arkusza).
// mode: "stop"|"warning"|"info".
// startColOffset: currentStartCol z ui-controls (przesunięcie kolumn w arkuszu).
// Zwraca true jeśli dodano, false jeśli nie ma aktywnego arkusza.
function addManualDvRule(sheetName, values, colIdx, startColOffset, mode, colName) {
  if (!sheetName || !values || !values.length) return false;
  const absCol = colIdx + (startColOffset || 0);
  const rule = {
    type: "list",
    ranges: [{ s: { r: 0, c: absCol }, e: { r: 1048575, c: absCol } }],
    mode: mode || "warning",
    allowBlank: true,
    formula1Raw: "__manual__",
    _values: values.slice(),
    _manual: true,
    _colName: colName || "",
  };
  if (!dvRulesBySheet) dvRulesBySheet = new Map();
  if (!dvRulesBySheet.has(sheetName)) dvRulesBySheet.set(sheetName, []);
  const arr = dvRulesBySheet.get(sheetName);
  // Usuń poprzednią ręczną regułę dla tej samej kolumny (replace).
  const kept = arr.filter((r) => !(r._manual && r.ranges[0]?.s.c === absCol));
  kept.push(rule);
  dvRulesBySheet.set(sheetName, kept);
  return true;
}

// Zwróć listę ręcznych reguł dla arkusza jako { colIdx, colName(?), mode }[].
// colName jest ustawiany przez wywołującego lub brak.
function getDvManualRules(sheetName) {
  if (!dvRulesBySheet || !dvRulesBySheet.has(sheetName)) return [];
  return dvRulesBySheet.get(sheetName)
    .filter((r) => r._manual)
    .map((r) => ({ colIdx: r.ranges[0]?.s.c ?? -1, colName: r._colName || `kol.${(r.ranges[0]?.s.c ?? 0) + 1}`, mode: r.mode }));
}

// Usuń ręczne reguły DV dla danego arkusza i kolumny (lub wszystkie jeśli absCol=-1).
function removeManualDvRule(sheetName, absCol) {
  if (!dvRulesBySheet || !dvRulesBySheet.has(sheetName)) return;
  const arr = dvRulesBySheet.get(sheetName);
  const filtered = absCol === -1
    ? arr.filter((r) => !r._manual)
    : arr.filter((r) => !(r._manual && r.ranges[0]?.s.c === absCol));
  if (filtered.length) dvRulesBySheet.set(sheetName, filtered);
  else dvRulesBySheet.delete(sheetName);
}
