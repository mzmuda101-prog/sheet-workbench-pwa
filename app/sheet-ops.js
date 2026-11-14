// sheet-ops.js — „od zera”: nowy skoroszyt oraz dodawanie / zmiana nazwy / usuwanie arkusza
// (docs/OD-ZERA.md, pozycja 0; 2026-11-14).
//
// Działa na SAMYM PLIKU (paczka ZIP w JSZip), tekstowo jak xlsx-patch.js — bez DOMParsera,
// więc da się to sprawdzić w Node (scripts/sheet-ops-test.js). Aplikacja (sheet-ops-ui.js)
// najpierw wpisuje bieżące poprawki do pliku w pamięci, potem woła tę operację i wczytuje
// wynik od nowa — zapis idzie dalej tą samą drogą co dziś (ZIP-patch).
//
// Zachowanie sprawdzone w prawdziwym Excelu (Mac, PL, 2026-11-14):
//   • nowy plik: „Zeszyt1” z jednym arkuszem „Arkusz1”; części jak z Excela (styl Normalny,
//     Aptos Narrow, motyw Office 2023) — szablon niżej to plik zapisany przez Excela, bez autora;
//   • nowy arkusz: „ArkuszN” z najniższym wolnym N (po ponownym otwarciu pliku Excel liczy tak samo);
//   • nazwa: 1–31 znaków, bez : \ / ? * [ ], bez apostrofu na początku i końcu, bez powtórki
//     (wielkość liter bez znaczenia), nie „Historia” (PL) / „History” (EN); spacje na brzegach wolno;
//   • zmiana nazwy poprawia odwołania we wszystkich formułach ('Moje dane'!A1), tekst w cudzysłowie
//     zostaje (="Arkusz3!A1");
//   • usunięcie: odwołanie do usuniętego arkusza w pliku to samo „#REF!” (=#REF!+SUM(#REF!)).
// Apostrofy wokół nazwy jak w Excelu: 'a-b', '2025', 'AB12', 'R1C1', 'TRUE'; bez nich Dane_ąę, Zł, abc.def.

const swbSheetOps = (() => {
  const FE = typeof FormulaEngine !== "undefined" ? FormulaEngine : (typeof require === "function" ? require("./formula-engine.js") : null);

  const NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
  const REL_WS = `${NS_REL}/worksheet`;
  const CT_WS = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";

  // ── nazwy ────────────────────────────────────────────────────────────────────
  const BAD_CHARS = /[:\\/?*[\]]/;
  const up = (s) => String(s).toLocaleUpperCase("pl");
  // Powód odmowy (klucz tekstu) albo null. `self` = obecna nazwa przy zmianie nazwy.
  function nameProblem(name, existing, self = null) {
    const n = String(name ?? "");
    if (!n.length) return "sheetNameEmpty";
    if (n.length > 31) return "sheetNameTooLong";
    if (BAD_CHARS.test(n)) return "sheetNameBadChars";
    if (n.startsWith("'") || n.endsWith("'")) return "sheetNameApostrophe";
    if (/^(historia|history)$/i.test(n)) return "sheetNameReserved";
    const taken = (existing || []).some((e) => up(e) === up(n) && (self == null || up(e) !== up(self)));
    return taken ? "sheetNameTaken" : null;
  }
  function nextName(existing, base) {
    const set = new Set((existing || []).map(up));
    for (let i = 1; ; i++) if (!set.has(up(base + i))) return base + i;
  }
  function quoteName(name) {
    const q = FE ? FE.quoteSheet(name) : `'${String(name).replace(/'/g, "''")}'`;
    // Excel bierze w apostrofy też nazwy wyglądające jak adres R1C1 (sprawdzone: 'R1C1')
    if (q === name && /^(R\d*C?\d*|C\d*)$/i.test(name)) return `'${name}'`;
    return q;
  }

  // ── formuły: odwołania do arkusza ─────────────────────────────────────────────
  const IDENT = /[\p{L}\p{N}_.\\]/u;
  const REF_BODY = /^(?:\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?|\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}|\$?\d+:\$?\d+|#REF!|[\p{L}_\\][\p{L}\p{N}_.]*)/u;
  // map(nazwaArkusza) → null (bez zmian) | { name: nowa } | { ref: true } (arkusz usunięty → #REF!)
  function rewriteFormula(text, map) {
    const s = String(text);
    let out = "";
    let i = 0;
    let changed = false;
    while (i < s.length) {
      const ch = s[i];
      if (ch === '"') { // tekst w cudzysłowie zostaje (Excel też go nie rusza)
        let j = i + 1;
        while (j < s.length) { if (s[j] === '"') { if (s[j + 1] === '"') { j += 2; continue; } break; } j++; }
        out += s.slice(i, j + 1); i = j + 1; continue;
      }
      let sheet = null;
      let len = 0;
      const prev = i ? s[i - 1] : "";
      if (ch === "'") {
        let j = i + 1;
        let name = "";
        while (j < s.length) { if (s[j] === "'") { if (s[j + 1] === "'") { name += "'"; j += 2; continue; } break; } name += s[j++]; }
        if (s[j] === "'" && s[j + 1] === "!") { sheet = name; len = j + 2 - i; }
      } else if (IDENT.test(ch) && !IDENT.test(prev) && prev !== "]" && prev !== "$") {
        let j = i;
        while (j < s.length && IDENT.test(s[j])) j++;
        if (s[j] === "!") { sheet = s.slice(i, j); len = j + 1 - i; }
      }
      const act = sheet != null && !sheet.startsWith("[") ? map(sheet) : null;
      if (!act) {
        if (len) { out += s.slice(i, i + len); i += len; } else { out += ch; i++; }
        continue;
      }
      changed = true;
      if (act.name != null) { out += `${quoteName(act.name)}!`; i += len; continue; }
      const body = s.slice(i + len).match(REF_BODY);
      out += "#REF!";
      i += len + (body ? body[0].length : 0);
    }
    return changed ? out : s;
  }

  const decodeXml = (s) => s.replace(/&(lt|gt|quot|apos|amp|#(\d+)|#x([0-9a-fA-F]+));/g, (m, n, d, h) => (
    d ? String.fromCodePoint(+d) : h ? String.fromCodePoint(parseInt(h, 16)) : { lt: "<", gt: ">", quot: '"', apos: "'", amp: "&" }[n]));
  const escText = (s) => String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
  const escAttr = (s) => escText(s).replace(/"/g, "&quot;");
  // Każdy element, w którym Excel trzyma formułę: komórki <f>, formatowanie warunkowe <formula>,
  // listy wyboru <formula1/2> i <xm:f> (x14), nazwy <definedName>, wykresy <c:f>, tabele.
  const FORMULA_TAGS = /<((?:[\w-]+:)?(?:f|formula|formula1|formula2|definedName|calculatedColumnFormula|totalsRowFormula))(\s[^>]*)?>([^<]*)<\/\1>/g;
  function rewriteXmlFormulas(xml, map) {
    let changed = false;
    const out = xml.replace(FORMULA_TAGS, (m, tag, attrs, body) => {
      const next = rewriteFormula(decodeXml(body), map);
      if (next === decodeXml(body)) return m;
      changed = true;
      return `<${tag}${attrs || ""}>${escText(next)}</${tag}>`;
    });
    return changed ? out : xml;
  }

  // ── części paczki ────────────────────────────────────────────────────────────
  const attr = (tag, name) => { const m = String(tag).match(new RegExp(`\\s${name}="([^"]*)"`)); return m ? m[1] : null; };
  const setAttr = (tag, name, value) => (attr(tag, name) == null
    ? tag.replace(/\s*(\/?)>$/, ` ${name}="${value}"$1>`)
    : tag.replace(new RegExp(`(\\s${name}=")[^"]*(")`), `$1${value}$2`));
  const dropAttr = (tag, name) => tag.replace(new RegExp(`\\s${name}="[^"]*"`), "");
  const read = async (zip, p) => (zip.file(p) ? zip.file(p).async("string") : null);
  const write = (zip, p, s) => zip.file(p, typeof TextEncoder !== "undefined" ? new TextEncoder().encode(s) : s);
  const relsPathOf = (part) => { const i = part.lastIndexOf("/"); return `${part.slice(0, i)}/_rels/${part.slice(i + 1)}.rels`; };
  const dirOf = (part) => part.slice(0, part.lastIndexOf("/"));
  function resolveTarget(baseDir, target) {
    const parts = target.startsWith("/") ? [] : baseDir.split("/").filter(Boolean);
    target.replace(/^\//, "").split("/").forEach((seg) => { if (seg === "..") parts.pop(); else if (seg && seg !== ".") parts.push(seg); });
    return parts.join("/");
  }
  const relsOf = (xml) => [...String(xml || "").matchAll(/<Relationship\b[^>]*>/g)].map((m) => ({
    raw: m[0], id: attr(m[0], "Id"), type: attr(m[0], "Type") || "", target: attr(m[0], "Target") || "", external: attr(m[0], "TargetMode") === "External",
  }));
  // <sheet> w workbook.xml (z prefiksem przestrzeni nazw albo bez) — w kolejności zakładek
  function sheetTags(wbXml) {
    return [...wbXml.matchAll(/<((?:[\w-]+:)?)sheet\b[^>]*\/>/g)].map((m) => ({
      raw: m[0], prefix: m[1], name: decodeXml(attr(m[0], "name") || ""), sheetId: Number(attr(m[0], "sheetId")),
      rid: (m[0].match(/\s[\w-]+:id="([^"]*)"/) || [])[1], state: attr(m[0], "state") || "visible",
    }));
  }
  async function workbookInfo(zip) {
    const wb = await read(zip, "xl/workbook.xml");
    const relsXml = await read(zip, "xl/_rels/workbook.xml.rels");
    if (!wb || !relsXml) throw new Error("brak xl/workbook.xml");
    const rels = relsOf(relsXml);
    const sheets = sheetTags(wb).map((s) => {
      const rel = rels.find((r) => r.id === s.rid);
      return { ...s, path: rel ? resolveTarget("xl", rel.target) : null };
    });
    return { wb, relsXml, rels, sheets };
  }

  // Wszystkie części XML poza arkuszem usuwanym — w nich poprawiamy formuły.
  async function rewriteAllFormulas(zip, map, skip = new Set()) {
    const paths = Object.keys(zip.files).filter((p) => !zip.files[p].dir && /^xl\/.*\.xml$/i.test(p) && !skip.has(p));
    for (const p of paths) {
      const xml = await read(zip, p);
      let next = rewriteXmlFormulas(xml, map);
      // źródło tabeli przestawnej: <worksheetSource ref="A1:D9" sheet="Dane"/>
      next = next.replace(/<((?:[\w-]+:)?)worksheetSource\b[^>]*>/g, (tag) => {
        const sh = attr(tag, "sheet");
        const act = sh != null ? map(decodeXml(sh)) : null;
        if (!act) return tag;
        return act.name != null ? setAttr(tag, "sheet", escAttr(act.name)) : tag;
      });
      if (next !== xml) write(zip, p, next);
    }
  }

  // Aktywna zakładka i jedno zaznaczone tabSelected (dwa = arkusze zgrupowane w Excelu).
  async function setActive(zip, wbXml, sheets, activeIndex) {
    let wb = wbXml.replace(/<((?:[\w-]+:)?)workbookView\b[^>]*>/, (tag) => {
      let t = setAttr(tag, "activeTab", String(activeIndex));
      const first = Number(attr(t, "firstSheet") || 0);
      if (first >= sheets.length || first > activeIndex) t = dropAttr(t, "firstSheet");
      return t;
    });
    for (let i = 0; i < sheets.length; i++) {
      const p = sheets[i].path;
      const xml = p && (await read(zip, p));
      if (!xml) continue;
      const next = xml.replace(/<((?:[\w-]+:)?)sheetView\b[^>]*>/, (tag) => (i === activeIndex ? setAttr(tag, "tabSelected", "1") : dropAttr(tag, "tabSelected")));
      if (next !== xml) write(zip, p, next);
    }
    return wb;
  }

  // docProps/app.xml: lista nazw arkuszy (Excel nadpisuje ją sam, ale niech się zgadza)
  async function syncAppXml(zip, map, added) {
    const xml = await read(zip, "docProps/app.xml");
    if (!xml) return;
    let next = xml.replace(/<vt:lpstr>([^<]*)<\/vt:lpstr>/g, (m, v) => {
      const act = map(decodeXml(v));
      if (!act) return m;
      return act.name != null ? `<vt:lpstr>${escText(act.name)}</vt:lpstr>` : "\u0000";
    });
    const removed = (next.match(/\u0000/g) || []).length;
    next = next.replace(/\u0000/g, "");
    if (!removed && !added) { if (next !== xml) write(zip, "docProps/app.xml", next); return; }
    // liczba arkuszy w HeadingPairs (pierwsza para) i rozmiar wektora TitlesOfParts
    const delta = (added ? 1 : 0) - removed;
    next = next.replace(/(<HeadingPairs>[\s\S]*?<vt:i4>)(\d+)(<\/vt:i4>)/, (m, a, n, b) => a + Math.max(0, Number(n) + delta) + b);
    next = next.replace(/(<TitlesOfParts>\s*<vt:vector\b[^>]*\bsize=")(\d+)(")/, (m, a, n, b) => a + Math.max(0, Number(n) + delta) + b);
    if (added) {
      // nazwy arkuszy stoją na początku wektora, w kolejności zakładek
      let k = 0;
      next = next.replace(/(<TitlesOfParts>\s*<vt:vector\b[^>]*>)([\s\S]*?)(<\/vt:vector>)/, (m, a, body, c) => {
        const items = body.match(/<vt:lpstr>[^<]*<\/vt:lpstr>/g) || [];
        items.splice(added.index, 0, `<vt:lpstr>${escText(added.name)}</vt:lpstr>`);
        k = items.length;
        return a + items.join("") + c;
      });
      if (!k) return; // nietypowy app.xml — Excel i tak go przepisze
    }
    write(zip, "docProps/app.xml", next);
  }

  async function dropCalcChain(zip) {
    if (!zip.file("xl/calcChain.xml")) return;
    zip.remove("xl/calcChain.xml");
    const ct = await read(zip, "[Content_Types].xml");
    if (ct) write(zip, "[Content_Types].xml", ct.replace(/<Override\b[^>]*calcChain\.xml[^>]*\/>/gi, ""));
    const rels = await read(zip, "xl/_rels/workbook.xml.rels");
    if (rels) write(zip, "xl/_rels/workbook.xml.rels", rels.replace(/<Relationship\b[^>]*calcChain\.xml[^>]*\/>/gi, ""));
  }

  const BLANK_SHEET = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><dimension ref="A1"/><sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="15"/><sheetData/><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>';

  // ── operacje ─────────────────────────────────────────────────────────────────
  // Nowy arkusz za arkuszem `afterName` (jak „+” w Excelu: na prawo od bieżącego). Zwraca nazwę.
  async function addSheet(zip, { name, afterName = null } = {}) {
    const info = await workbookInfo(zip);
    const names = info.sheets.map((s) => s.name);
    if (nameProblem(name, names)) throw new Error(nameProblem(name, names));
    const at = afterName != null && names.includes(afterName) ? names.indexOf(afterName) + 1 : names.length;
    // wysokość wiersza i szerokość jak w pierwszym arkuszu (te same czcionki stylu Normalny)
    const firstXml = info.sheets[0] && info.sheets[0].path ? await read(zip, info.sheets[0].path) : null;
    const fmt = firstXml && firstXml.match(/<sheetFormatPr\b[^>]*\/>/);
    let sheetXml = BLANK_SHEET;
    if (fmt) {
      let f = fmt[0].replace(/^<[\w-]+:/, "<").replace(/\s[\w-]+:dyDescent="[^"]*"/, "");
      ["customHeight", "zeroHeight", "thickTop", "thickBottom", "outlineLevelRow", "outlineLevelCol", "defaultColWidth"].forEach((a) => { f = dropAttr(f, a); });
      sheetXml = sheetXml.replace(/<sheetFormatPr\b[^>]*\/>/, f);
    }
    let n = 1;
    while (zip.file(`xl/worksheets/sheet${n}.xml`)) n++;
    const path = `xl/worksheets/sheet${n}.xml`;
    write(zip, path, sheetXml);
    let rid = 1;
    const ids = new Set(info.rels.map((r) => r.id));
    while (ids.has(`rId${rid}`)) rid++;
    const relsXml = info.relsXml.replace(/<\/Relationships>\s*$/, `<Relationship Id="rId${rid}" Type="${REL_WS}" Target="worksheets/sheet${n}.xml"/></Relationships>`);
    write(zip, "xl/_rels/workbook.xml.rels", relsXml);
    const ct = await read(zip, "[Content_Types].xml");
    write(zip, "[Content_Types].xml", ct.replace(/<\/Types>\s*$/, `<Override PartName="/${path}" ContentType="${CT_WS}"/></Types>`));
    const sheetId = Math.max(0, ...info.sheets.map((s) => s.sheetId || 0)) + 1;
    const pfx = (info.sheets[0] && info.sheets[0].prefix) || "";
    const relPfx = ((info.sheets[0] && info.sheets[0].raw.match(/\s([\w-]+):id="/)) || [])[1] || "r";
    const tag = `<${pfx}sheet name="${escAttr(name)}" sheetId="${sheetId}" ${relPfx}:id="rId${rid}"/>`;
    let wb = info.wb;
    if (at < info.sheets.length) wb = wb.replace(info.sheets[at].raw, tag + info.sheets[at].raw);
    else wb = wb.replace(info.sheets[info.sheets.length - 1].raw, info.sheets[info.sheets.length - 1].raw + tag);
    // nazwy lokalne (localSheetId) arkuszy za wstawionym przesuwają się o jeden
    wb = shiftLocalSheetIds(wb, (id) => (id >= at ? id + 1 : id));
    const sheets = info.sheets.slice();
    sheets.splice(at, 0, { name, path });
    wb = await setActive(zip, wb, sheets, at);
    write(zip, "xl/workbook.xml", wb);
    await syncAppXml(zip, () => null, { index: at, name });
    return name;
  }

  function shiftLocalSheetIds(wb, fn) {
    return wb.replace(/<((?:[\w-]+:)?)definedName\b[^>]*>/g, (tag) => {
      const id = attr(tag, "localSheetId");
      return id == null ? tag : setAttr(tag, "localSheetId", String(fn(Number(id))));
    });
  }

  async function renameSheet(zip, { from, to }) {
    const info = await workbookInfo(zip);
    const names = info.sheets.map((s) => s.name);
    const idx = names.indexOf(from);
    if (idx < 0) throw new Error("brak arkusza");
    if (nameProblem(to, names, from)) throw new Error(nameProblem(to, names, from));
    if (to === from) return to;
    const map = (sh) => (up(sh) === up(from) ? { name: to } : null);
    await rewriteAllFormulas(zip, map);
    let wb = await read(zip, "xl/workbook.xml");
    const s = sheetTags(wb)[idx];
    wb = wb.replace(s.raw, setAttr(s.raw, "name", escAttr(to)));
    write(zip, "xl/workbook.xml", wb);
    await syncAppXml(zip, map, null);
    return to;
  }

  // Usunięcie arkusza wraz z częściami, do których prowadzi tylko on (tabele, notatki, rysunki…).
  // Zwraca nazwę arkusza, który staje się aktywny (sąsiad z prawej, jak w Excelu, albo z lewej).
  async function deleteSheet(zip, { name }) {
    const info = await workbookInfo(zip);
    const idx = info.sheets.findIndex((s) => s.name === name);
    if (idx < 0) throw new Error("brak arkusza");
    const visible = info.sheets.filter((s) => s.state === "visible");
    if (info.sheets[idx].state === "visible" && visible.length <= 1) throw new Error("sheetDeleteLast");
    const victim = info.sheets[idx];
    // części do usunięcia: arkusz + to, do czego prowadzą jego relacje, o ile nikt inny tam nie prowadzi
    const doomed = new Set();
    const queue = [victim.path];
    while (queue.length) {
      const p = queue.shift();
      if (!p || doomed.has(p) || !zip.file(p)) continue;
      doomed.add(p);
      const rels = relsOf(await read(zip, relsPathOf(p)));
      rels.filter((r) => !r.external).forEach((r) => queue.push(resolveTarget(dirOf(p), r.target)));
    }
    // część, do której prowadzi relacja spoza usuwanego zestawu, zostaje (do skutku: zostaje też
    // to, do czego prowadzi część ocalona, np. rekordy pamięci podręcznej tabeli przestawnej)
    const relFiles = [];
    for (const rf of Object.keys(zip.files).filter((p) => /\.rels$/.test(p))) {
      const owner = rf === "_rels/.rels" ? "" : rf.replace(/_rels\/([^/]+)\.rels$/, "$1");
      relFiles.push({ owner, targets: relsOf(await read(zip, rf)).filter((r) => !r.external).map((r) => resolveTarget(owner.includes("/") ? dirOf(owner) : "", r.target)) });
    }
    for (let changed = true; changed;) {
      changed = false;
      relFiles.forEach(({ owner, targets }) => {
        if (owner && doomed.has(owner)) return;
        targets.forEach((t) => { if (t !== victim.path && doomed.delete(t)) changed = true; });
      });
    }
    doomed.add(victim.path);
    for (const p of doomed) {
      zip.remove(p);
      if (zip.file(relsPathOf(p))) zip.remove(relsPathOf(p));
    }
    let ct = await read(zip, "[Content_Types].xml");
    ct = ct.replace(/<Override\b[^>]*\/>/g, (tag) => (doomed.has(String(attr(tag, "PartName") || "").replace(/^\//, "")) ? "" : tag));
    write(zip, "[Content_Types].xml", ct);
    const wbRel = info.rels.find((r) => r.id === victim.rid);
    if (wbRel) write(zip, "xl/_rels/workbook.xml.rels", info.relsXml.replace(wbRel.raw, ""));
    // formuły w pozostałych częściach: odwołania do usuniętego arkusza → #REF!
    const map = (sh) => (up(sh) === up(name) ? { ref: true } : null);
    await rewriteAllFormulas(zip, map, doomed);
    let wb = await read(zip, "xl/workbook.xml");
    wb = wb.replace(victim.raw, "");
    // nazwy lokalne usuniętego arkusza znikają, dalsze przesuwają się o jeden w lewo
    wb = wb.replace(/<((?:[\w-]+:)?)definedName\b[^>]*>[^<]*<\/\1definedName>/g, (m) => (attr(m, "localSheetId") === String(idx) ? "" : m));
    wb = shiftLocalSheetIds(wb, (id) => (id > idx ? id - 1 : id));
    wb = wb.replace(/<((?:[\w-]+:)?)definedNames>\s*<\/\1definedNames>/, "");
    const rest = info.sheets.filter((s, i) => i !== idx);
    let next = rest.findIndex((s, i) => i >= idx && s.state === "visible");
    if (next < 0) next = rest.map((s) => s.state).lastIndexOf("visible");
    wb = await setActive(zip, wb, rest, next);
    write(zip, "xl/workbook.xml", wb);
    await dropCalcChain(zip); // łańcuch obliczeń wskazuje arkusze numerami — Excel „naprawiałby” plik
    await syncAppXml(zip, map, null);
    return rest[next].name;
  }

  // Pusty skoroszyt jak z Excela (Mac, PL): Zeszyt1 → Arkusz1. Daty utworzenia = teraz, bez autora.
  async function newWorkbook(JSZipCtor, { sheetName = "Arkusz1", headingSheets = "Arkusze", now = new Date() } = {}) {
    const zip = new JSZipCtor();
    const iso = new Date(now).toISOString().replace(/\.\d{3}Z$/, "Z");
    const x = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
    const wsName = escText(sheetName);
    write(zip, "[Content_Types].xml", `${x}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="${CT_WS}"/><Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>`);
    write(zip, "_rels/.rels", `${x}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId3" Type="${NS_REL}/extended-properties" Target="docProps/app.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId1" Type="${NS_REL}/officeDocument" Target="xl/workbook.xml"/></Relationships>`);
    write(zip, "docProps/app.xml", `${x}<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>Microsoft Excel</Application><DocSecurity>0</DocSecurity><ScaleCrop>false</ScaleCrop><HeadingPairs><vt:vector size="2" baseType="variant"><vt:variant><vt:lpstr>${escText(headingSheets)}</vt:lpstr></vt:variant><vt:variant><vt:i4>1</vt:i4></vt:variant></vt:vector></HeadingPairs><TitlesOfParts><vt:vector size="1" baseType="lpstr"><vt:lpstr>${wsName}</vt:lpstr></vt:vector></TitlesOfParts><LinksUpToDate>false</LinksUpToDate><SharedDoc>false</SharedDoc><HyperlinksChanged>false</HyperlinksChanged><AppVersion>16.0300</AppVersion></Properties>`);
    write(zip, "docProps/core.xml", `${x}<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dcterms:created xsi:type="dcterms:W3CDTF">${iso}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${iso}</dcterms:modified></cp:coreProperties>`);
    write(zip, "xl/_rels/workbook.xml.rels", `${x}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId3" Type="${NS_REL}/styles" Target="styles.xml"/><Relationship Id="rId2" Type="${NS_REL}/theme" Target="theme/theme1.xml"/><Relationship Id="rId1" Type="${REL_WS}" Target="worksheets/sheet1.xml"/></Relationships>`);
    write(zip, "xl/workbook.xml", `${x}<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="${NS_REL}"><fileVersion appName="xl" lastEdited="7" lowestEdited="7" rupBuild="11108"/><workbookPr defaultThemeVersion="202300"/><bookViews><workbookView xWindow="0" yWindow="0" windowWidth="28800" windowHeight="17000"/></bookViews><sheets><sheet name="${escAttr(sheetName)}" sheetId="1" r:id="rId1"/></sheets><calcPr calcId="191029"/></workbook>`);
    write(zip, "xl/worksheets/sheet1.xml", BLANK_SHEET.replace('<sheetView workbookViewId="0"/>', '<sheetView tabSelected="1" workbookViewId="0"/>').replace('defaultRowHeight="15"', 'baseColWidth="10" defaultRowHeight="16"'));
    write(zip, "xl/styles.xml", `${x}<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font><sz val="12"/><color theme="1"/><name val="Aptos Narrow"/><family val="2"/><charset val="238"/><scheme val="minor"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs><cellStyles count="1"><cellStyle name="Normalny" xfId="0" builtinId="0"/></cellStyles><dxfs count="0"/><tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/></styleSheet>`);
    write(zip, "xl/theme/theme1.xml", `${x}${THEME}`);
    return zip;
  }

  const THEME = "<a:theme xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" name=\"Motyw pakietu Office\"><a:themeElements><a:clrScheme name=\"Pakiet Office\"><a:dk1><a:sysClr val=\"windowText\" lastClr=\"000000\"/></a:dk1><a:lt1><a:sysClr val=\"window\" lastClr=\"FFFFFF\"/></a:lt1><a:dk2><a:srgbClr val=\"0E2841\"/></a:dk2><a:lt2><a:srgbClr val=\"E8E8E8\"/></a:lt2><a:accent1><a:srgbClr val=\"156082\"/></a:accent1><a:accent2><a:srgbClr val=\"E97132\"/></a:accent2><a:accent3><a:srgbClr val=\"196B24\"/></a:accent3><a:accent4><a:srgbClr val=\"0F9ED5\"/></a:accent4><a:accent5><a:srgbClr val=\"A02B93\"/></a:accent5><a:accent6><a:srgbClr val=\"4EA72E\"/></a:accent6><a:hlink><a:srgbClr val=\"467886\"/></a:hlink><a:folHlink><a:srgbClr val=\"96607D\"/></a:folHlink></a:clrScheme><a:fontScheme name=\"Pakiet Office\"><a:majorFont><a:latin typeface=\"Aptos Display\" panose=\"02110004020202020204\"/><a:ea typeface=\"\"/><a:cs typeface=\"\"/><a:font script=\"Jpan\" typeface=\"游ゴシック Light\"/><a:font script=\"Hang\" typeface=\"맑은 고딕\"/><a:font script=\"Hans\" typeface=\"等线 Light\"/><a:font script=\"Hant\" typeface=\"新細明體\"/><a:font script=\"Arab\" typeface=\"Times New Roman\"/><a:font script=\"Hebr\" typeface=\"Times New Roman\"/><a:font script=\"Thai\" typeface=\"Tahoma\"/><a:font script=\"Ethi\" typeface=\"Nyala\"/><a:font script=\"Beng\" typeface=\"Vrinda\"/><a:font script=\"Gujr\" typeface=\"Shruti\"/><a:font script=\"Khmr\" typeface=\"MoolBoran\"/><a:font script=\"Knda\" typeface=\"Tunga\"/><a:font script=\"Guru\" typeface=\"Raavi\"/><a:font script=\"Cans\" typeface=\"Euphemia\"/><a:font script=\"Cher\" typeface=\"Plantagenet Cherokee\"/><a:font script=\"Yiii\" typeface=\"Microsoft Yi Baiti\"/><a:font script=\"Tibt\" typeface=\"Microsoft Himalaya\"/><a:font script=\"Thaa\" typeface=\"MV Boli\"/><a:font script=\"Deva\" typeface=\"Mangal\"/><a:font script=\"Telu\" typeface=\"Gautami\"/><a:font script=\"Taml\" typeface=\"Latha\"/><a:font script=\"Syrc\" typeface=\"Estrangelo Edessa\"/><a:font script=\"Orya\" typeface=\"Kalinga\"/><a:font script=\"Mlym\" typeface=\"Kartika\"/><a:font script=\"Laoo\" typeface=\"DokChampa\"/><a:font script=\"Sinh\" typeface=\"Iskoola Pota\"/><a:font script=\"Mong\" typeface=\"Mongolian Baiti\"/><a:font script=\"Viet\" typeface=\"Times New Roman\"/><a:font script=\"Uigh\" typeface=\"Microsoft Uighur\"/><a:font script=\"Geor\" typeface=\"Sylfaen\"/><a:font script=\"Armn\" typeface=\"Arial\"/><a:font script=\"Bugi\" typeface=\"Leelawadee UI\"/><a:font script=\"Bopo\" typeface=\"Microsoft JhengHei\"/><a:font script=\"Java\" typeface=\"Javanese Text\"/><a:font script=\"Lisu\" typeface=\"Segoe UI\"/><a:font script=\"Mymr\" typeface=\"Myanmar Text\"/><a:font script=\"Nkoo\" typeface=\"Ebrima\"/><a:font script=\"Olck\" typeface=\"Nirmala UI\"/><a:font script=\"Osma\" typeface=\"Ebrima\"/><a:font script=\"Phag\" typeface=\"Phagspa\"/><a:font script=\"Syrn\" typeface=\"Estrangelo Edessa\"/><a:font script=\"Syrj\" typeface=\"Estrangelo Edessa\"/><a:font script=\"Syre\" typeface=\"Estrangelo Edessa\"/><a:font script=\"Sora\" typeface=\"Nirmala UI\"/><a:font script=\"Tale\" typeface=\"Microsoft Tai Le\"/><a:font script=\"Talu\" typeface=\"Microsoft New Tai Lue\"/><a:font script=\"Tfng\" typeface=\"Ebrima\"/></a:majorFont><a:minorFont><a:latin typeface=\"Aptos Narrow\" panose=\"02110004020202020204\"/><a:ea typeface=\"\"/><a:cs typeface=\"\"/><a:font script=\"Jpan\" typeface=\"游ゴシック\"/><a:font script=\"Hang\" typeface=\"맑은 고딕\"/><a:font script=\"Hans\" typeface=\"等线\"/><a:font script=\"Hant\" typeface=\"新細明體\"/><a:font script=\"Arab\" typeface=\"Arial\"/><a:font script=\"Hebr\" typeface=\"Arial\"/><a:font script=\"Thai\" typeface=\"Tahoma\"/><a:font script=\"Ethi\" typeface=\"Nyala\"/><a:font script=\"Beng\" typeface=\"Vrinda\"/><a:font script=\"Gujr\" typeface=\"Shruti\"/><a:font script=\"Khmr\" typeface=\"DaunPenh\"/><a:font script=\"Knda\" typeface=\"Tunga\"/><a:font script=\"Guru\" typeface=\"Raavi\"/><a:font script=\"Cans\" typeface=\"Euphemia\"/><a:font script=\"Cher\" typeface=\"Plantagenet Cherokee\"/><a:font script=\"Yiii\" typeface=\"Microsoft Yi Baiti\"/><a:font script=\"Tibt\" typeface=\"Microsoft Himalaya\"/><a:font script=\"Thaa\" typeface=\"MV Boli\"/><a:font script=\"Deva\" typeface=\"Mangal\"/><a:font script=\"Telu\" typeface=\"Gautami\"/><a:font script=\"Taml\" typeface=\"Latha\"/><a:font script=\"Syrc\" typeface=\"Estrangelo Edessa\"/><a:font script=\"Orya\" typeface=\"Kalinga\"/><a:font script=\"Mlym\" typeface=\"Kartika\"/><a:font script=\"Laoo\" typeface=\"DokChampa\"/><a:font script=\"Sinh\" typeface=\"Iskoola Pota\"/><a:font script=\"Mong\" typeface=\"Mongolian Baiti\"/><a:font script=\"Viet\" typeface=\"Arial\"/><a:font script=\"Uigh\" typeface=\"Microsoft Uighur\"/><a:font script=\"Geor\" typeface=\"Sylfaen\"/><a:font script=\"Armn\" typeface=\"Arial\"/><a:font script=\"Bugi\" typeface=\"Leelawadee UI\"/><a:font script=\"Bopo\" typeface=\"Microsoft JhengHei\"/><a:font script=\"Java\" typeface=\"Javanese Text\"/><a:font script=\"Lisu\" typeface=\"Segoe UI\"/><a:font script=\"Mymr\" typeface=\"Myanmar Text\"/><a:font script=\"Nkoo\" typeface=\"Ebrima\"/><a:font script=\"Olck\" typeface=\"Nirmala UI\"/><a:font script=\"Osma\" typeface=\"Ebrima\"/><a:font script=\"Phag\" typeface=\"Phagspa\"/><a:font script=\"Syrn\" typeface=\"Estrangelo Edessa\"/><a:font script=\"Syrj\" typeface=\"Estrangelo Edessa\"/><a:font script=\"Syre\" typeface=\"Estrangelo Edessa\"/><a:font script=\"Sora\" typeface=\"Nirmala UI\"/><a:font script=\"Tale\" typeface=\"Microsoft Tai Le\"/><a:font script=\"Talu\" typeface=\"Microsoft New Tai Lue\"/><a:font script=\"Tfng\" typeface=\"Ebrima\"/></a:minorFont></a:fontScheme><a:fmtScheme name=\"Pakiet Office\"><a:fillStyleLst><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:gradFill rotWithShape=\"1\"><a:gsLst><a:gs pos=\"0\"><a:schemeClr val=\"phClr\"><a:lumMod val=\"110000\"/><a:satMod val=\"105000\"/><a:tint val=\"67000\"/></a:schemeClr></a:gs><a:gs pos=\"50000\"><a:schemeClr val=\"phClr\"><a:lumMod val=\"105000\"/><a:satMod val=\"103000\"/><a:tint val=\"73000\"/></a:schemeClr></a:gs><a:gs pos=\"100000\"><a:schemeClr val=\"phClr\"><a:lumMod val=\"105000\"/><a:satMod val=\"109000\"/><a:tint val=\"81000\"/></a:schemeClr></a:gs></a:gsLst><a:lin ang=\"5400000\" scaled=\"0\"/></a:gradFill><a:gradFill rotWithShape=\"1\"><a:gsLst><a:gs pos=\"0\"><a:schemeClr val=\"phClr\"><a:satMod val=\"103000\"/><a:lumMod val=\"102000\"/><a:tint val=\"94000\"/></a:schemeClr></a:gs><a:gs pos=\"50000\"><a:schemeClr val=\"phClr\"><a:satMod val=\"110000\"/><a:lumMod val=\"100000\"/><a:shade val=\"100000\"/></a:schemeClr></a:gs><a:gs pos=\"100000\"><a:schemeClr val=\"phClr\"><a:lumMod val=\"99000\"/><a:satMod val=\"120000\"/><a:shade val=\"78000\"/></a:schemeClr></a:gs></a:gsLst><a:lin ang=\"5400000\" scaled=\"0\"/></a:gradFill></a:fillStyleLst><a:lnStyleLst><a:ln w=\"12700\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\"><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:prstDash val=\"solid\"/><a:miter lim=\"800000\"/></a:ln><a:ln w=\"19050\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\"><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:prstDash val=\"solid\"/><a:miter lim=\"800000\"/></a:ln><a:ln w=\"25400\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\"><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:prstDash val=\"solid\"/><a:miter lim=\"800000\"/></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad=\"57150\" dist=\"19050\" dir=\"5400000\" algn=\"ctr\" rotWithShape=\"0\"><a:srgbClr val=\"000000\"><a:alpha val=\"63000\"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:solidFill><a:schemeClr val=\"phClr\"><a:tint val=\"95000\"/><a:satMod val=\"170000\"/></a:schemeClr></a:solidFill><a:gradFill rotWithShape=\"1\"><a:gsLst><a:gs pos=\"0\"><a:schemeClr val=\"phClr\"><a:tint val=\"93000\"/><a:satMod val=\"150000\"/><a:shade val=\"98000\"/><a:lumMod val=\"102000\"/></a:schemeClr></a:gs><a:gs pos=\"50000\"><a:schemeClr val=\"phClr\"><a:tint val=\"98000\"/><a:satMod val=\"130000\"/><a:shade val=\"90000\"/><a:lumMod val=\"103000\"/></a:schemeClr></a:gs><a:gs pos=\"100000\"><a:schemeClr val=\"phClr\"><a:shade val=\"63000\"/><a:satMod val=\"120000\"/></a:schemeClr></a:gs></a:gsLst><a:lin ang=\"5400000\" scaled=\"0\"/></a:gradFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements><a:objectDefaults><a:lnDef><a:spPr/><a:bodyPr/><a:lstStyle/><a:style><a:lnRef idx=\"2\"><a:schemeClr val=\"accent1\"/></a:lnRef><a:fillRef idx=\"0\"><a:schemeClr val=\"accent1\"/></a:fillRef><a:effectRef idx=\"1\"><a:schemeClr val=\"accent1\"/></a:effectRef><a:fontRef idx=\"minor\"><a:schemeClr val=\"tx1\"/></a:fontRef></a:style></a:lnDef></a:objectDefaults><a:extraClrSchemeLst/><a:extLst><a:ext uri=\"{05A4C25C-085E-4340-85A3-A5531E510DB2}\"><thm15:themeFamily xmlns:thm15=\"http://schemas.microsoft.com/office/thememl/2012/main\" name=\"Office Theme\" id=\"{2E142A2C-CD16-42D6-873A-C26D2A0506FA}\" vid=\"{1BDDFF52-6CD6-40A5-AB3C-68EB2F1E4D0A}\"/></a:ext></a:extLst></a:theme>";

  // ── Formatuj jako tabelę ─────────────────────────────────────────────────────
  // Jak Excel (sprawdzone 2026-11-15): <table id name=displayName="TabelaN" ref totalsRowShown="0"> z <autoFilter>
  // na całym zakresie, kolumny = nagłówki z arkusza (aplikacja wpisuje je wcześniej jako tekst: „Kolumna1”,
  // „Nazwa2”, „2026”), styl TableStyleMedium2. Filtr arkusza na tym zakresie przechodzi do tabeli (Excel też
  // nie pozwala na oba naraz). Zwraca nazwę tabeli.
  const REL_TABLE = `${NS_REL}/table`;
  const CT_TABLE = "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml";
  const tableColName = (n) => escAttr(String(n)).replace(/\r?\n/g, "_x000a_");
  async function addTable(zip, { sheet, ref, name, columns, style = "TableStyleMedium2" }) {
    const info = await workbookInfo(zip);
    const idx = info.sheets.findIndex((s) => s.name === sheet);
    if (idx < 0 || !info.sheets[idx].path) throw new Error("brak arkusza");
    const sheetPath = info.sheets[idx].path;
    // id i ścieżka: największe id w skoroszycie + 1, pierwsza wolna tableN.xml
    let maxId = 0;
    for (const p of Object.keys(zip.files).filter((f) => /^xl\/tables\/[^/]+\.xml$/i.test(f))) {
      const m = (await read(zip, p)).match(/<(?:[\w-]+:)?table\b[^>]*\sid="(\d+)"/);
      if (m) maxId = Math.max(maxId, Number(m[1]));
    }
    let n = 1;
    while (zip.file(`xl/tables/table${n}.xml`)) n++;
    const path = `xl/tables/table${n}.xml`;
    const cols = columns.map((c, i) => `<tableColumn id="${i + 1}" name="${tableColName(c)}"/>`).join("");
    write(zip, path, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="${maxId + 1}" name="${escAttr(name)}" displayName="${escAttr(name)}" ref="${ref}" totalsRowShown="0"><autoFilter ref="${ref}"/><tableColumns count="${columns.length}">${cols}</tableColumns><tableStyleInfo name="${style}" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/></table>`);
    const ct = await read(zip, "[Content_Types].xml");
    write(zip, "[Content_Types].xml", ct.replace(/<\/Types>\s*$/, `<Override PartName="/${path}" ContentType="${CT_TABLE}"/></Types>`));
    // relacja arkusza
    const relsPath = relsPathOf(sheetPath);
    const relsXml = (await read(zip, relsPath)) || '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';
    const ids = new Set(relsOf(relsXml).map((r) => r.id));
    let rid = 1;
    while (ids.has(`rId${rid}`)) rid++;
    const target = `../tables/table${n}.xml`;
    write(zip, relsPath, relsXml.replace(/<\/Relationships>\s*$/, `<Relationship Id="rId${rid}" Type="${REL_TABLE}" Target="${target}"/></Relationships>`));
    // arkusz: <tablePart> (przed extLst — kolejność ze schematu) i bez filtra arkusza
    let xml = await read(zip, sheetPath);
    const pfx = (xml.match(/<((?:[\w-]+:)?)worksheet\b/) || [])[1] || "";
    // prefiks relacji z GŁÓWNEGO znacznika (bywa zadeklarowany tylko lokalnie, np. na <hyperlink>)
    const root = (xml.match(/<(?:[\w-]+:)?worksheet\b[^>]*>/) || [""])[0];
    const relPfx = (root.match(/\sxmlns:([\w-]+)="http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/relationships"/) || [])[1];
    if (!relPfx) xml = xml.replace(/<((?:[\w-]+:)?)worksheet\b/, `<$1worksheet xmlns:r="${NS_REL}"`);
    const rp = relPfx || "r";
    const hadFilter = /<(?:[\w-]+:)?autoFilter\b/.test(xml);
    xml = xml.replace(/<((?:[\w-]+:)?)autoFilter\b[^>]*\/>|<((?:[\w-]+:)?)autoFilter\b[^>]*>[\s\S]*?<\/\2autoFilter>/, "");
    const part = `<${pfx}tablePart ${rp}:id="rId${rid}"/>`;
    if (/<(?:[\w-]+:)?tableParts\b/.test(xml)) {
      xml = xml.replace(/<((?:[\w-]+:)?)tableParts\b([^>]*)>([\s\S]*?)<\/\1tableParts>/, (m, p1, attrs, body) => `<${p1}tableParts${attrs.replace(/\scount="\d+"/, ` count="${(body.match(/<(?:[\w-]+:)?tablePart\b/g) || []).length + 1}"`)}>${body}${part}</${p1}tableParts>`);
    } else {
      const block = `<${pfx}tableParts count="1">${part}</${pfx}tableParts>`;
      const ext = xml.search(new RegExp(`<${pfx}extLst\\b`));
      xml = ext >= 0 ? xml.slice(0, ext) + block + xml.slice(ext) : xml.replace(new RegExp(`</${pfx}worksheet>\\s*$`), `${block}</${pfx}worksheet>`);
    }
    write(zip, sheetPath, xml);
    // ukryta nazwa filtra arkusza (_xlnm._FilterDatabase) — filtr przeszedł do tabeli
    if (hadFilter) {
      const wb = await read(zip, "xl/workbook.xml");
      const next = wb.replace(/<((?:[\w-]+:)?)definedName\b[^>]*>[^<]*<\/\1definedName>/g, (m) => (/name="_xlnm\._FilterDatabase"/.test(m) && attr(m, "localSheetId") === String(idx) ? "" : m))
        .replace(/<((?:[\w-]+:)?)definedNames>\s*<\/\1definedNames>/, "");
      if (next !== wb) write(zip, "xl/workbook.xml", next);
    }
    return name;
  }

  // Kolumny tabeli jak Excel: pusta → „Kolumna1, Kolumna2…”, powtórzona (bez wielkości liter) → „Nazwa2”,
  // liczba / data → tekst. headers = wartości komórek nagłówka (już jako tekst albo null).
  function tableColumnNames(headers, emptyBase = "Kolumna") {
    const used = new Set();
    let k = 0;
    return headers.map((h) => {
      let base = h == null ? "" : String(h).trim() ? String(h) : "";
      let name = base;
      if (!base) { do { name = `${emptyBase}${++k}`; } while (used.has(up(name))); }
      else if (used.has(up(base))) { let i = 2; while (used.has(up(base + i))) i++; name = base + i; }
      used.add(up(name));
      return name;
    });
  }

  return { nameProblem, nextName, quoteName, rewriteFormula, rewriteXmlFormulas, addSheet, renameSheet, deleteSheet, newWorkbook, workbookInfo, addTable, tableColumnNames };
})();

if (typeof module !== "undefined" && module.exports) module.exports = swbSheetOps;
