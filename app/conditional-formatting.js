// app/conditional-formatting.js
// Odczyt i ewaluacja formatowania warunkowego (CF) z surowego .xlsx:
//  - parsuje <dxfs> (style różnicowe) i <conditionalFormatting> (reguły),
//  - ewaluuje reguły expression/cellIs własnym mini-silnikiem formuł (podzbiór Excela),
//  - oraz containsText/top10/aboveAverage/duplicate/colorScale.
// Stan globalny (currentDxfs, currentCFRules, cfEvalCache) jest w core.js.
// Reużywa colorFromStyleNode() i decodeXmlEntities() z table.js (funkcje globalne).

/* ===================== Formatowanie warunkowe (CF) =====================
   Excel zmienia kolory czcionek/teł przez reguły conditionalFormatting (dxf).
   Biblioteka tego nie parsuje, więc czytamy z surowego .xlsx, a reguły typu
   `expression`/`cellIs` ewaluujemy własnym mini-silnikiem formuł (podzbiór). */

function cfDateToSerial(d) {
  return Math.round((Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()) - Date.UTC(1899, 11, 30)) / 86400000);
}
function cfToday() { return cfDateToSerial(new Date()); }
function cfSerialToDate(s) { return new Date(Date.UTC(1899, 11, 30) + s * 86400000); }

const CF_ERR = { __cfError: true };
function cfIsErr(x) { return !!(x && x.__cfError); }

function cfCellValue(sheet, r, c) {
  if (r < 0 || c < 0) return null;
  const cell = sheet[XLSX.utils.encode_cell({ r, c })];
  if (!cell || cell.v == null || cell.v === "") return null;
  const v = cell.v;
  if (v instanceof Date) return cfDateToSerial(v);
  if (typeof v === "boolean" || typeof v === "number") return v;
  return String(v);
}

function cfNum(v) {
  if (v === null) return 0;
  if (typeof v === "number") return v;
  if (typeof v === "boolean") return v ? 1 : 0;
  if (typeof v === "string") { if (v.trim() === "") return 0; const n = Number(v); return isNaN(n) ? NaN : n; }
  return NaN;
}
function cfStr(v) {
  if (v === null) return "";
  if (typeof v === "boolean") return v ? "TRUE" : "FALSE";
  return String(v);
}
function cfBool(v) {
  if (cfIsErr(v)) return v;
  if (typeof v === "boolean") return v;
  if (v === null) return false;
  if (typeof v === "number") return v !== 0;
  const s = String(v).toLowerCase();
  if (s === "true") return true;
  if (s === "false" || s === "") return false;
  return CF_ERR;
}
function cfEquals(a, b) {
  const ea = a === null, eb = b === null;
  if (ea && eb) return true;
  if (ea) return b === 0 || b === "";
  if (eb) return a === 0 || a === "";
  if (typeof a === "number" && typeof b === "number") return a === b;
  if (typeof a === "boolean" || typeof b === "boolean") return cfNum(a) === cfNum(b);
  return String(a).toLowerCase() === String(b).toLowerCase();
}
function cfCompare(op, a, b) {
  if (op === "=") return cfEquals(a, b);
  if (op === "<>") return !cfEquals(a, b);
  let cmp;
  if (typeof a !== "string" && typeof b !== "string") cmp = cfNum(a) - cfNum(b);
  else cmp = cfStr(a).toLowerCase().localeCompare(cfStr(b).toLowerCase());
  if (op === "<") return cmp < 0;
  if (op === ">") return cmp > 0;
  if (op === "<=") return cmp <= 0;
  if (op === ">=") return cmp >= 0;
  return false;
}

// ---- tokenizer / parser ----
const CF_BINPREC = { "=": 1, "<>": 1, "<": 1, ">": 1, "<=": 1, ">=": 1, "&": 2, "+": 3, "-": 3, "*": 4, "/": 4, "^": 5 };
function cfTokenize(s) {
  const toks = [];
  const re = /\s*(?:(\$?[A-Za-z]{1,3}\$?\d+)|(\d+(?:\.\d+)?(?:[eE][+-]?\d+)?)|"((?:[^"]|"")*)"|(<=|>=|<>|[-+*/^&=<>(),:])|([A-Za-z_][A-Za-z0-9_.]*))/g;
  let m;
  while ((m = re.exec(s)) !== null) {
    if (m[0].trim() === "") { if (re.lastIndex === m.index) re.lastIndex++; continue; }
    if (m[1] !== undefined) toks.push({ t: "ref", v: m[1] });
    else if (m[2] !== undefined) toks.push({ t: "num", v: parseFloat(m[2]) });
    else if (m[3] !== undefined) toks.push({ t: "str", v: m[3].replace(/""/g, '"') });
    else if (m[4] !== undefined) toks.push({ t: "op", v: m[4] });
    else if (m[5] !== undefined) {
      const up = m[5].toUpperCase();
      if (up === "TRUE") toks.push({ t: "bool", v: true });
      else if (up === "FALSE") toks.push({ t: "bool", v: false });
      else toks.push({ t: "name", v: up });
    }
  }
  return toks;
}
function cfParse(tokens) {
  let pos = 0;
  const peek = () => tokens[pos];
  const next = () => tokens[pos++];
  function parseExpr(minPrec) {
    let left = parseUnary();
    while (true) {
      const tk = peek();
      if (!tk || tk.t !== "op") break;
      const prec = CF_BINPREC[tk.v];
      if (prec == null || prec < minPrec) break;
      next();
      const right = parseExpr(tk.v === "^" ? prec : prec + 1);
      left = { k: "bin", op: tk.v, a: left, b: right };
    }
    return left;
  }
  function parseUnary() {
    const tk = peek();
    if (tk && tk.t === "op" && (tk.v === "-" || tk.v === "+")) { next(); return { k: "un", op: tk.v, a: parseUnary() }; }
    return parsePrimary();
  }
  function parsePrimary() {
    const tk = next();
    if (!tk) throw 0;
    if (tk.t === "num") return { k: "num", v: tk.v };
    if (tk.t === "str") return { k: "str", v: tk.v };
    if (tk.t === "bool") return { k: "bool", v: tk.v };
    if (tk.t === "ref") {
      if (peek() && peek().t === "op" && peek().v === ":") { next(); const r2 = next(); return { k: "range", a: tk.v, b: r2 && r2.v }; }
      return { k: "ref", v: tk.v };
    }
    if (tk.t === "name") {
      if (peek() && peek().t === "op" && peek().v === "(") {
        next();
        const args = [];
        if (!(peek() && peek().t === "op" && peek().v === ")")) {
          args.push(parseExpr(0));
          while (peek() && peek().t === "op" && peek().v === ",") { next(); args.push(parseExpr(0)); }
        }
        const cl = next();
        if (!cl || cl.v !== ")") throw 0;
        return { k: "call", name: tk.v, args };
      }
      throw 0;
    }
    if (tk.t === "op" && tk.v === "(") { const e = parseExpr(0); const c = next(); if (!c || c.v !== ")") throw 0; return e; }
    throw 0;
  }
  const ast = parseExpr(0);
  if (pos < tokens.length) throw 0;
  return ast;
}
function cfRefToRC(ref, dr, dc) {
  const m = /^(\$?)([A-Za-z]{1,3})(\$?)(\d+)$/.exec(ref);
  if (!m) return null;
  let col = 0;
  const letters = m[2].toUpperCase();
  for (let i = 0; i < letters.length; i++) col = col * 26 + (letters.charCodeAt(i) - 64);
  col -= 1;
  let row = parseInt(m[4], 10) - 1;
  if (!m[1]) col += dc;
  if (!m[3]) row += dr;
  return { r: row, c: col };
}
function cfEval(node, ctx) {
  if (!node) return CF_ERR;
  switch (node.k) {
    case "num": case "str": case "bool": return node.v;
    case "ref": { const rc = cfRefToRC(node.v, ctx.dr, ctx.dc); return rc ? cfCellValue(ctx.sheet, rc.r, rc.c) : CF_ERR; }
    case "range": return cfRangeValues(node, ctx);
    case "un": { const a = cfEval(node.a, ctx); if (cfIsErr(a)) return a; const n = cfNum(a); return isNaN(n) ? CF_ERR : (node.op === "-" ? -n : n); }
    case "bin": return cfEvalBin(node, ctx);
    case "call": return cfEvalCall(node, ctx);
  }
  return CF_ERR;
}
function cfRangeValues(node, ctx) {
  const a = cfRefToRC(node.a, ctx.dr, ctx.dc);
  const b = cfRefToRC(node.b, ctx.dr, ctx.dc);
  if (!a || !b) return CF_ERR;
  const out = [];
  for (let r = Math.min(a.r, b.r); r <= Math.max(a.r, b.r); r++)
    for (let c = Math.min(a.c, b.c); c <= Math.max(a.c, b.c); c++)
      out.push(cfCellValue(ctx.sheet, r, c));
  return out;
}
function cfEvalBin(node, ctx) {
  const a = cfEval(node.a, ctx); if (cfIsErr(a)) return a;
  const b = cfEval(node.b, ctx); if (cfIsErr(b)) return b;
  const op = node.op;
  if (op === "&") return cfStr(a) + cfStr(b);
  if (["=", "<>", "<", ">", "<=", ">="].includes(op)) return cfCompare(op, a, b);
  const x = cfNum(a), y = cfNum(b);
  if (isNaN(x) || isNaN(y)) return CF_ERR;
  if (op === "+") return x + y;
  if (op === "-") return x - y;
  if (op === "*") return x * y;
  if (op === "/") return y === 0 ? CF_ERR : x / y;
  if (op === "^") return Math.pow(x, y);
  return CF_ERR;
}
function cfFlat(args) {
  const out = [];
  for (const a of args) { if (Array.isArray(a)) out.push(...a); else out.push(a); }
  return out;
}
function cfEvalCall(node, ctx) {
  const name = node.name, A = node.args;
  if (name === "IF") { const c = cfBool(cfEval(A[0], ctx)); if (cfIsErr(c)) return c; return c ? cfEval(A[1], ctx) : (A[2] !== undefined ? cfEval(A[2], ctx) : false); }
  if (name === "AND") { for (const a of A) { const v = cfBool(cfEval(a, ctx)); if (cfIsErr(v)) return v; if (!v) return false; } return true; }
  if (name === "OR") { for (const a of A) { const v = cfBool(cfEval(a, ctx)); if (cfIsErr(v)) return v; if (v) return true; } return false; }
  if (name === "NOT") { const v = cfBool(cfEval(A[0], ctx)); if (cfIsErr(v)) return v; return !v; }
  if (name === "IFERROR") { const v = cfEval(A[0], ctx); return cfIsErr(v) ? cfEval(A[1], ctx) : v; }
  if (name === "TODAY" || name === "NOW") return cfToday();
  const args = A.map((a) => cfEval(a, ctx));
  const tolerant = name === "ISERROR" || name === "ISERR" || name === "ISNA";
  if (!tolerant) for (const a of args) if (cfIsErr(a)) return a;
  return cfFunc(name, args);
}
function cfFunc(name, args) {
  const a0 = args[0];
  const s = (i) => cfStr(args[i]);
  const n = (i) => cfNum(args[i]);
  switch (name) {
    case "ISERROR": case "ISERR": case "ISNA": return cfIsErr(a0);
    case "ISBLANK": return a0 === null;
    case "ISNUMBER": return typeof a0 === "number";
    case "ISTEXT": return typeof a0 === "string";
    case "ISLOGICAL": return typeof a0 === "boolean";
    case "ISNONTEXT": return typeof a0 !== "string";
    case "LEN": return s(0).length;
    case "LEFT": return s(0).slice(0, args[1] === undefined ? 1 : n(1));
    case "RIGHT": { const k = args[1] === undefined ? 1 : n(1); return k <= 0 ? "" : s(0).slice(-k); }
    case "MID": return s(0).substr(Math.max(0, n(1) - 1), n(2));
    case "UPPER": return s(0).toUpperCase();
    case "LOWER": return s(0).toLowerCase();
    case "TRIM": return s(0).trim().replace(/\s+/g, " ");
    case "PROPER": return s(0).replace(/\b\w/g, (m) => m.toUpperCase());
    case "EXACT": return s(0) === s(1);
    case "CONCATENATE": case "CONCAT": return cfFlat(args).map(cfStr).join("");
    case "SEARCH": { const p = s(0).toLowerCase(), w = s(1).toLowerCase(); const st = args[2] === undefined ? 0 : n(2) - 1; const i = w.indexOf(p, st); return i < 0 ? CF_ERR : i + 1; }
    case "FIND": { const p = s(0), w = s(1); const st = args[2] === undefined ? 0 : n(2) - 1; const i = w.indexOf(p, st); return i < 0 ? CF_ERR : i + 1; }
    case "ABS": return Math.abs(n(0));
    case "INT": return Math.floor(n(0));
    case "ROUND": { const f = Math.pow(10, n(1) || 0); return Math.round(n(0) * f) / f; }
    case "ROUNDDOWN": { const f = Math.pow(10, n(1) || 0); return Math.trunc(n(0) * f) / f; }
    case "MOD": { const y = n(1); return y === 0 ? CF_ERR : ((n(0) % y) + y) % y; }
    case "SQRT": return Math.sqrt(n(0));
    case "POWER": return Math.pow(n(0), n(1));
    case "MIN": { const v = cfFlat(args).map(cfNum).filter((x) => !isNaN(x)); return v.length ? Math.min(...v) : 0; }
    case "MAX": { const v = cfFlat(args).map(cfNum).filter((x) => !isNaN(x)); return v.length ? Math.max(...v) : 0; }
    case "SUM": return cfFlat(args).reduce((acc, x) => acc + (isNaN(cfNum(x)) ? 0 : cfNum(x)), 0);
    case "COUNT": return cfFlat(args).filter((x) => typeof x === "number").length;
    case "COUNTA": return cfFlat(args).filter((x) => x !== null && x !== "").length;
    case "AVERAGE": { const v = cfFlat(args).filter((x) => typeof x === "number"); return v.length ? v.reduce((a, b) => a + b, 0) / v.length : CF_ERR; }
    case "VALUE": { const x = Number(s(0).replace(/\s/g, "").replace(",", ".")); return isNaN(x) ? CF_ERR : x; }
    case "YEAR": return cfSerialToDate(n(0)).getUTCFullYear();
    case "MONTH": return cfSerialToDate(n(0)).getUTCMonth() + 1;
    case "DAY": return cfSerialToDate(n(0)).getUTCDate();
    case "WEEKDAY": { const wd = cfSerialToDate(n(0)).getUTCDay(); return wd === 0 ? 7 : wd; }
    case "DATE": return cfDateToSerial(new Date(Date.UTC(n(0), n(1) - 1, n(2))));
    case "EDATE": { const d = cfSerialToDate(n(0)); return cfDateToSerial(new Date(Date.UTC(d.getUTCFullYear(), d.getUTCMonth() + n(1), d.getUTCDate()))); }
    case "DATEDIF": {
      const a = n(0), b = n(1);
      if (isNaN(a) || isNaN(b) || b < a) return CF_ERR;
      const d1 = cfSerialToDate(a), d2 = cfSerialToDate(b);
      const y1 = d1.getUTCFullYear(), m1 = d1.getUTCMonth(), day1 = d1.getUTCDate();
      const y2 = d2.getUTCFullYear(), m2 = d2.getUTCMonth(), day2 = d2.getUTCDate();
      const unit = String(s(2)).toLowerCase();
      if (unit === "d") return Math.round(b - a);
      if (unit === "m") { let mo = (y2 - y1) * 12 + (m2 - m1); if (day2 < day1) mo--; return mo; }
      if (unit === "y") { let yr = y2 - y1; if (m2 < m1 || (m2 === m1 && day2 < day1)) yr--; return yr; }
      if (unit === "ym") { let mo = (y2 - y1) * 12 + (m2 - m1); if (day2 < day1) mo--; return ((mo % 12) + 12) % 12; }
      if (unit === "md") { let dd = day2 - day1; if (dd < 0) dd += new Date(Date.UTC(y2, m2, 0)).getUTCDate(); return dd; }
      if (unit === "yd") {
        let start = Date.UTC(y2, m1, day1);
        if (start > Date.UTC(y2, m2, day2)) start = Date.UTC(y2 - 1, m1, day1);
        return Math.round((Date.UTC(y2, m2, day2) - start) / 86400000);
      }
      return CF_ERR;
    }
    case "ROUNDUP": { const f = Math.pow(10, n(1) || 0); const x = n(0) * f; return (x < 0 ? Math.floor(x) : Math.ceil(x)) / f; }
  }
  return CF_ERR;
}

// ---- parsowanie dxf + reguł CF z XML ----
function cfColorFromEl(el) {
  if (!el) return null;
  const node = {};
  if (el.hasAttribute("rgb")) node.rgb = el.getAttribute("rgb");
  if (el.hasAttribute("theme")) node.theme = parseInt(el.getAttribute("theme"), 10);
  if (el.hasAttribute("tint")) node.tint = parseFloat(el.getAttribute("tint"));
  if (el.hasAttribute("indexed")) node.indexed = parseInt(el.getAttribute("indexed"), 10);
  if (el.hasAttribute("auto")) node.auto = el.getAttribute("auto");
  return colorFromStyleNode(node, true);
}
function parseDxfs(stylesXml) {
  const out = [];
  let doc;
  try { doc = new DOMParser().parseFromString(stylesXml, "application/xml"); } catch { return out; }
  const dxfs = doc.getElementsByTagName("dxf");
  for (let i = 0; i < dxfs.length; i++) {
    const dxf = dxfs[i];
    let fontColor = null;
    const fontEl = dxf.getElementsByTagName("font")[0];
    if (fontEl) fontColor = cfColorFromEl(fontEl.getElementsByTagName("color")[0]);
    let fillColor = null;
    const pf = dxf.getElementsByTagName("patternFill")[0];
    if (pf) fillColor = cfColorFromEl(pf.getElementsByTagName("bgColor")[0]) || cfColorFromEl(pf.getElementsByTagName("fgColor")[0]);
    out.push({ fontColor, fillColor });
  }
  return out;
}
function cfDecodeRange(ref) {
  try { return XLSX.utils.decode_range(ref); } catch { return null; }
}
function parseSheetCF(sheetXml) {
  const blocks = [];
  let doc;
  try { doc = new DOMParser().parseFromString(sheetXml, "application/xml"); } catch { return blocks; }
  const cfs = doc.getElementsByTagName("conditionalFormatting");
  for (let i = 0; i < cfs.length; i++) {
    const cf = cfs[i];
    const sqref = cf.getAttribute("sqref") || "";
    const ranges = sqref.trim().split(/\s+/).filter(Boolean).map(cfDecodeRange).filter(Boolean);
    if (!ranges.length) continue;
    const anchor = ranges[0].s;
    const ruleEls = cf.getElementsByTagName("cfRule");
    const rules = [];
    for (let j = 0; j < ruleEls.length; j++) {
      const re = ruleEls[j];
      const rule = {
        type: re.getAttribute("type"),
        dxfId: re.hasAttribute("dxfId") ? parseInt(re.getAttribute("dxfId"), 10) : null,
        priority: re.hasAttribute("priority") ? parseInt(re.getAttribute("priority"), 10) : 0,
        operator: re.getAttribute("operator") || "",
        text: re.getAttribute("text") || "",
        stopIfTrue: re.getAttribute("stopIfTrue") === "1",
        formulas: Array.from(re.getElementsByTagName("formula")).map((f) => f.textContent),
        rank: re.hasAttribute("rank") ? parseInt(re.getAttribute("rank"), 10) : 10,
        percent: re.getAttribute("percent") === "1",
        bottom: re.getAttribute("bottom") === "1",
        aboveAverage: re.getAttribute("aboveAverage") !== "0",
        equalAverage: re.getAttribute("equalAverage") === "1",
        colorScale: null,
        _ast: {},
      };
      const cs = re.getElementsByTagName("colorScale")[0];
      if (cs) {
        const cfvos = Array.from(cs.getElementsByTagName("cfvo")).map((v) => ({ type: v.getAttribute("type"), val: v.getAttribute("val") }));
        const colors = Array.from(cs.getElementsByTagName("color")).map(cfColorFromEl);
        rule.colorScale = { cfvos, colors };
      }
      rules.push(rule);
    }
    rules.sort((a, b) => a.priority - b.priority);
    blocks.push({ ranges, anchor, rules });
  }
  return blocks;
}
function cfSheetPathsFromRels(relsXml) {
  const map = new Map();
  const re = /<Relationship\b[^>]*?Id="([^"]+)"[^>]*?Target="([^"]+)"[^>]*?\/?>/g;
  let m;
  while ((m = re.exec(relsXml))) {
    const target = m[2];
    if (!/worksheets\//.test(target)) continue;
    let path = decodeXmlEntities(target);
    if (path.startsWith("/")) path = path.slice(1);
    else if (!path.startsWith("xl/")) path = "xl/" + path.replace(/^\.\//, "");
    map.set(m[1], path);
  }
  return map;
}
async function buildConditionalFormatting(bytes, wb) {
  currentDxfs = [];
  currentCFRules = null;
  cfEvalCache = new Map();
  currentTables = {};
  if (typeof JSZip === "undefined" || !bytes || !wb) return;
  try {
    const zip = await JSZip.loadAsync(bytes);
    currentTables = await parseTables(zip);
    const stylesXml = await zip.file("xl/styles.xml")?.async("string");
    if (stylesXml) currentDxfs = parseDxfs(stylesXml);
    const wbXml = await zip.file("xl/workbook.xml")?.async("string");
    const relsXml = await zip.file("xl/_rels/workbook.xml.rels")?.async("string");
    if (!wbXml || !relsXml) return;
    const ridToPath = cfSheetPathsFromRels(relsXml);
    const rules = new Map();
    const sheetRe = /<sheet\b[^>]*?name="([^"]+)"[^>]*?r:id="([^"]+)"[^>]*?\/?>/g;
    let sm;
    while ((sm = sheetRe.exec(wbXml))) {
      const name = decodeXmlEntities(sm[1]);
      const path = ridToPath.get(sm[2]);
      const file = path ? zip.file(path) : null;
      if (!file) continue;
      const xml = await file.async("string");
      if (xml.indexOf("conditionalFormatting") === -1) continue;
      const blocks = parseSheetCF(xml);
      if (blocks.length) rules.set(name, blocks);
    }
    currentCFRules = rules.size ? rules : null;
  } catch {
    currentCFRules = null;
  }
}

// ---- ewaluacja CF per arkusz (z cache) ----
function getSheetCFMap(sheetName) {
  if (!currentCFRules || !currentCFRules.has(sheetName)) return null;
  if (cfEvalCache.has(sheetName)) return cfEvalCache.get(sheetName);
  let map;
  try { map = evalSheetCF(sheetName); } catch { map = new Map(); }
  cfEvalCache.set(sheetName, map);
  return map;
}
function cfBlockStats(sheet, block, used) {
  const nums = [];
  const counts = new Map();
  for (const range of block.ranges) {
    const r1 = used ? Math.max(range.s.r, used.s.r) : range.s.r;
    const r2 = used ? Math.min(range.e.r, used.e.r) : range.e.r;
    const c1 = used ? Math.max(range.s.c, used.s.c) : range.s.c;
    const c2 = used ? Math.min(range.e.c, used.e.c) : range.e.c;
    for (let r = r1; r <= r2; r++) for (let c = c1; c <= c2; c++) {
      const v = cfCellValue(sheet, r, c);
      if (v === null) continue;
      if (typeof v === "number") nums.push(v);
      const key = typeof v === "string" ? v.toLowerCase() : v;
      counts.set(key, (counts.get(key) || 0) + 1);
    }
  }
  const n = nums.length;
  const avg = n ? nums.reduce((a, b) => a + b, 0) / n : 0;
  return { nums, avg, n, counts };
}
function cfLerpColor(h1, h2, t) {
  const p = (h) => { const m = /^#?([0-9a-f]{6})$/i.exec(h || ""); if (!m) return null; const x = parseInt(m[1], 16); return [(x >> 16) & 255, (x >> 8) & 255, x & 255]; };
  const a = p(h1), b = p(h2);
  if (!a || !b) return h1 || h2;
  const c = a.map((v, i) => Math.round(v + (b[i] - v) * t));
  return `#${c.map((v) => v.toString(16).padStart(2, "0")).join("")}`;
}
function cfColorScaleColor(rule, v, stats) {
  if (typeof v !== "number" || !rule.colorScale) return null;
  const { cfvos, colors } = rule.colorScale;
  if (colors.length < 2 || !stats.nums.length) return null;
  const sorted = stats.nums.slice().sort((a, b) => a - b);
  const lo = sorted[0], hi = sorted[sorted.length - 1];
  const valFor = (cv) => {
    if (cv.type === "min") return lo;
    if (cv.type === "max") return hi;
    if (cv.type === "percentile") return sorted[Math.round((sorted.length - 1) * parseFloat(cv.val) / 100)];
    if (cv.type === "percent") return lo + (hi - lo) * parseFloat(cv.val) / 100;
    return parseFloat(cv.val);
  };
  const stops = cfvos.map((cv, i) => ({ v: valFor(cv), color: colors[i] })).filter((st) => st.color != null && !isNaN(st.v));
  if (stops.length < 2) return null;
  if (v <= stops[0].v) return stops[0].color;
  if (v >= stops[stops.length - 1].v) return stops[stops.length - 1].color;
  for (let i = 0; i < stops.length - 1; i++) {
    if (v >= stops[i].v && v <= stops[i + 1].v) {
      const span = stops[i + 1].v - stops[i].v || 1;
      return cfLerpColor(stops[i].color, stops[i + 1].color, (v - stops[i].v) / span);
    }
  }
  return null;
}
function cfRuleResult(rule, sheet, r, c, dr, dc, stats) {
  const t = rule.type;
  if (t === "colorScale") {
    const col = cfColorScaleColor(rule, cfCellValue(sheet, r, c), stats);
    return col ? { fillColor: col } : null;
  }
  let matched = false;
  if (t === "expression") {
    let ast = rule._ast[0];
    if (ast === undefined) { try { ast = cfParse(cfTokenize(rule.formulas[0] || "")); } catch { ast = null; } rule._ast[0] = ast; }
    if (!ast) return null;
    matched = cfBool(cfEval(ast, { sheet, dr, dc })) === true;
  } else if (t === "cellIs") {
    matched = cfCellIs(rule, sheet, r, c, dr, dc);
  } else if (t === "containsText" || t === "notContainsText" || t === "beginsWith" || t === "endsWith") {
    const val = cfCellValue(sheet, r, c);
    const str = (val == null ? "" : String(val)).toLowerCase();
    const needle = (rule.text || "").toLowerCase();
    if (needle) {
      if (t === "containsText") matched = str.includes(needle);
      else if (t === "notContainsText") matched = !str.includes(needle);
      else if (t === "beginsWith") matched = str.startsWith(needle);
      else matched = str.endsWith(needle);
    }
  } else if (t === "top10") {
    const v = cfCellValue(sheet, r, c);
    if (typeof v === "number" && stats.nums.length) {
      const sorted = stats.nums.slice().sort((a, b) => rule.bottom ? a - b : b - a);
      let k = rule.rank || 10;
      if (rule.percent) k = Math.max(1, Math.round(stats.n * k / 100));
      const th = sorted[Math.min(k, sorted.length) - 1];
      matched = th != null && (rule.bottom ? v <= th : v >= th);
    }
  } else if (t === "aboveAverage") {
    const v = cfCellValue(sheet, r, c);
    if (typeof v === "number") {
      const above = rule.aboveAverage !== false;
      if (rule.equalAverage) matched = above ? v >= stats.avg : v <= stats.avg;
      else matched = above ? v > stats.avg : v < stats.avg;
    }
  } else if (t === "duplicateValues" || t === "uniqueValues") {
    const v = cfCellValue(sheet, r, c);
    if (v !== null) {
      const cnt = stats.counts.get(typeof v === "string" ? v.toLowerCase() : v) || 0;
      matched = t === "duplicateValues" ? cnt > 1 : cnt === 1;
    }
  } else {
    return null;
  }
  if (!matched) return null;
  const dxf = rule.dxfId != null ? currentDxfs[rule.dxfId] : null;
  if (!dxf) return null;
  return { fontColor: dxf.fontColor, fillColor: dxf.fillColor };
}
function cfCellIs(rule, sheet, r, c, dr, dc) {
  const val = cfCellValue(sheet, r, c);
  const ev = (i) => {
    let ast = rule._ast[i];
    if (ast === undefined) { try { ast = cfParse(cfTokenize(rule.formulas[i] || "")); } catch { ast = null; } rule._ast[i] = ast; }
    return ast ? cfEval(ast, { sheet, dr, dc }) : CF_ERR;
  };
  const v1 = ev(0);
  if (cfIsErr(v1)) return false;
  const op = rule.operator;
  if (op === "between" || op === "notBetween") {
    const v2 = ev(1);
    if (cfIsErr(v2)) return false;
    const lo = Math.min(cfNum(v1), cfNum(v2)), hi = Math.max(cfNum(v1), cfNum(v2));
    const x = cfNum(val);
    const inside = x >= lo && x <= hi;
    return op === "between" ? inside : !inside;
  }
  const opMap = { greaterThan: ">", lessThan: "<", equal: "=", notEqual: "<>", greaterThanOrEqual: ">=", lessThanOrEqual: "<=" };
  const o = opMap[op];
  return o ? cfCompare(o, val, v1) === true : false;
}
function evalSheetCF(sheetName) {
  const map = new Map();
  const sheet = workbook.Sheets[sheetName];
  const blocks = currentCFRules.get(sheetName);
  if (!sheet || !blocks) return map;
  const used = cfDecodeRange(sheet["!ref"]);
  let budget = 400000;
  for (const block of blocks) {
    const needStats = block.rules.some((r) => ["top10", "aboveAverage", "colorScale", "duplicateValues", "uniqueValues"].includes(r.type));
    const stats = needStats ? cfBlockStats(sheet, block, used) : null;
    for (const range of block.ranges) {
      const r1 = used ? Math.max(range.s.r, used.s.r) : range.s.r;
      const r2 = used ? Math.min(range.e.r, used.e.r) : range.e.r;
      const c1 = used ? Math.max(range.s.c, used.s.c) : range.s.c;
      const c2 = used ? Math.min(range.e.c, used.e.c) : range.e.c;
      for (let r = r1; r <= r2; r++) {
        for (let c = c1; c <= c2; c++) {
          if (budget-- <= 0) return map;
          const dr = r - block.anchor.r, dc = c - block.anchor.c;
          let fontColor = null, fillColor = null;
          for (const rule of block.rules) {
            const res = cfRuleResult(rule, sheet, r, c, dr, dc, stats);
            if (!res) continue;
            if (res.fontColor && !fontColor) fontColor = res.fontColor;
            if (res.fillColor && !fillColor) fillColor = res.fillColor;
            if (rule.stopIfTrue) break;
            if (fontColor && fillColor) break;
          }
          if (fontColor || fillColor) map.set(XLSX.utils.encode_cell({ r, c }), { fontColor, fillColor });
        }
      }
    }
  }
  return map;
}

// Parsuje xl/tables/*.xml → currentTables: nazwa/displayName(lower) -> { columns: {kol(lower): absColIndex} }.
async function parseTables(zip) {
  const tables = {};
  const files = Object.keys(zip.files).filter((f) => /xl\/tables\/table\d+\.xml$/i.test(f));
  for (const f of files) {
    let xml;
    try { xml = await zip.file(f).async("string"); } catch { continue; }
    const refM = xml.match(/\bref="([A-Z]+\d+:[A-Z]+\d+|[A-Z]+\d+)"/);
    if (!refM) continue;
    let startCol;
    try { startCol = XLSX.utils.decode_range(refM[1]).s.c; } catch { continue; }
    const cols = {};
    let cm, idx = 0;
    const colRe = /<tableColumn\b[^>]*\bname="([^"]+)"/g;
    while ((cm = colRe.exec(xml))) { cols[decodeXmlEntities(cm[1]).toLowerCase()] = startCol + idx; idx++; }
    const meta = { columns: cols };
    const nameM = xml.match(/<table\b[^>]*\bname="([^"]+)"/);
    const dispM = xml.match(/\bdisplayName="([^"]+)"/);
    if (nameM) tables[decodeXmlEntities(nameM[1]).toLowerCase()] = meta;
    if (dispM) tables[decodeXmlEntities(dispM[1]).toLowerCase()] = meta;
  }
  return tables;
}

// Zamienia odwołania strukturalne dla BIEŻĄCEGO wiersza na zwykłe A1, np.
// `Tabela[[#This Row],[od]]` / `Tabela[@od]` / `Tabela[@[od]]` → "C5".
// Odwołania do całej kolumny (bez #This Row/@) zostają nietknięte (nieobsługiwane w przeliczeniu wiersza).
function resolveStructuredRefs(formula, thisRow) {
  if (!formula || formula.indexOf("[") < 0) return formula;
  return String(formula).replace(
    /([A-Za-z_À-ɏ][\w.À-ɏ]*)\[((?:[^\[\]]|\[[^\]]*\])*)\]/g,
    (whole, tname, inner) => {
      const t = currentTables[tname.toLowerCase()];
      if (!t) return whole;
      const innerS = inner.trim();
      if (!/#This Row/i.test(innerS) && innerS.indexOf("@") < 0) return whole; // tylko bieżący wiersz
      let col = null;
      const lastBracket = innerS.match(/\[([^\[\]]+)\]\s*$/);
      if (lastBracket && lastBracket[1].trim()[0] !== "#") col = lastBracket[1].trim();
      if (!col) { const at = innerS.match(/@\s*\[?\s*([^\[\]@]+?)\s*\]?$/); if (at) col = at[1].trim(); }
      if (!col) return whole;
      const colIdx = t.columns[col.toLowerCase()];
      if (colIdx == null) return whole;
      return XLSX.utils.encode_cell({ r: thisRow, c: colIdx });
    }
  );
}

// Przelicza formułę komórki (cell.f) używając silnika CF — do ODŚWIEŻANIA wartości
// zależnych od TODAY()/NOW() (np. „Długość dni" do dzisiaj) bez wchodzenia do Excela.
// thisRow = absolutny wiersz komórki (dla odwołań strukturalnych [#This Row]/@).
// Zwraca number|string|boolean albo null, gdy formuła jest nieobsługiwana/błędna → zostaje wartość z pliku.
function cfRecomputeCellFormula(sheet, formulaText, thisRow) {
  if (!sheet || !formulaText) return null;
  const resolved = resolveStructuredRefs(formulaText, thisRow || 0);
  let ast;
  try { ast = cfParse(cfTokenize(resolved)); } catch { return null; }
  let v;
  try { v = cfEval(ast, { sheet, dr: 0, dc: 0 }); } catch { return null; }
  if (cfIsErr(v) || v === undefined) return null;
  if (typeof v === "number" && !isFinite(v)) return null;
  return v;
}
