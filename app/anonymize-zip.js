// anonymize-zip.js — etap 2 anonimizacji: ślady danych POZA komórkami i test szczelności.
//
// Plik .xlsx trzyma kopie danych w wielu miejscach: zapamiętane wyniki formuł, stare teksty
// w sharedStrings (po podmianie komórki tekst zostaje w tabeli napisów), pamięć wykresów
// (c:strCache), pamięć tabel przestawnych, listy walidacji, hiperłącza (mailto:), notatki
// i ich autorzy, nagłówki/stopki wydruku, właściwości pliku (autor, firma), ścieżka do pliku
// (x15ac:absPath). swbAnonCleanZip przepisuje je tą samą sesją (spójnie z komórkami),
// a na końcu szuka oryginałów we WSZYSTKICH częściach pliku i zwraca raport.
//
// Używa pomocników z xlsx-patch.js (globalne): parseRelsList, resolvePartPath, relsRemove,
// removeLegacyDrawing, ctRemoveOverride, zipWriteUtf8Xml, forceRecalcOnLoad, dropCalcChain,
// resolveSheetPaths, escapeXmlText, sanitizeXmlText, decodeXmlEntitiesSimple.

const SWB_ANON_ZIP_DEFAULTS = {
  metadata: true, // autor, ostatnio zmodyfikowane, firma, menedżer, tytuł…, ścieżka pliku
  notes: "anonymize", // "anonymize" | "remove" | "keep"
  headersFooters: "anonymize", // "anonymize" | "clear" | "keep"
  formulaCache: true, // przepisz/wyczyść zapamiętane wyniki formuł (Excel przeliczy przy otwarciu)
  charts: true, // pamięć wykresów i teksty w rysunkach
  pivots: true, // pamięć tabel przestawnych (+ odświeżenie przy otwarciu)
  links: true, // hiperłącza, listy walidacji, stałe w formułach i nazwach
  sharedStrings: true, // nieużywane już teksty z tabeli napisów
};

function swbAnonDecode(s) {
  return decodeXmlEntitiesSimple(String(s).replace(/&#x([0-9a-f]+);/gi, (m, h) => String.fromCodePoint(parseInt(h, 16))).replace(/&#(\d+);/g, (m, d) => String.fromCodePoint(+d)));
}
function swbAnonEncodeText(s) {
  return escapeXmlText(sanitizeXmlText(s));
}
function swbAnonEncodeAttr(s) {
  return swbAnonEncodeText(s).replace(/"/g, "&quot;");
}

// Przepisuje treść elementów tekstowych (np. <t>, <c:v>, <a:t>) przez fn(tekst).
function swbAnonRewriteTextNodes(xml, tagRe, fn) {
  let count = 0;
  const out = xml.replace(tagRe, (all, open, body, close) => {
    if (!body) return all;
    const before = swbAnonDecode(body);
    const after = fn(before);
    if (after === before) return all;
    count++;
    return `${open}${swbAnonEncodeText(after)}${close}`;
  });
  return { xml: out, count };
}
// Przepisuje wartość atrybutu (np. <s v="…"/>, display="…").
function swbAnonRewriteAttr(xml, attrRe, fn) {
  let count = 0;
  const out = xml.replace(attrRe, (all, pre, val, post) => {
    const before = swbAnonDecode(val);
    const after = fn(before);
    if (after === before) return all;
    count++;
    return `${pre}${swbAnonEncodeAttr(after)}${post}`;
  });
  return { xml: out, count };
}
// Teksty w cudzysłowach wewnątrz formuły / listy walidacji ("a,b" albo "Jan Kowalski").
function swbAnonRewriteQuoted(text, fn, splitList) {
  return text.replace(/"((?:[^"]|"")*)"/g, (all, inner) => {
    const raw = inner.replace(/""/g, '"');
    const next = splitList ? raw.split(",").map((x) => { const t = x.trim(); const r = fn(t); return r === t ? x : x.replace(t, r); }).join(",") : fn(raw);
    return next === raw ? all : `"${next.replace(/"/g, '""')}"`;
  });
}

async function swbAnonText(zip, path) {
  const f = zip.file(path);
  return f ? f.async("string") : null;
}
function swbAnonRelsPathOf(partPath) {
  const dir = partPath.slice(0, partPath.lastIndexOf("/"));
  return `${dir}/_rels/${partPath.slice(partPath.lastIndexOf("/") + 1)}.rels`;
}

// Regex wycieków: oryginały jako całe słowa, bez wielkości liter.
function swbAnonLeakRegex(terms) {
  const list = [...terms].filter((t) => String(t).length >= 3).sort((a, b) => b.length - a.length);
  if (!list.length) return null;
  const esc = list.map((t) => String(t).replace(/[.*+?^${}()|[\]\\]/g, "\\$&").replace(/\s+/g, "\\s+"));
  // paczki po 400 — jedna ogromna alternatywa bywa wolna
  const out = [];
  for (let i = 0; i < esc.length; i += 400) out.push(new RegExp(`(?<![\\p{L}\\p{N}])(?:${esc.slice(i, i + 400).join("|")})(?![\\p{L}\\p{N}])`, "giu"));
  return out;
}
function swbAnonLeaksIn(text, res) {
  if (!res) return [];
  const found = new Set();
  res.forEach((re) => { re.lastIndex = 0; for (const m of text.matchAll(re)) found.add(m[0]); });
  return [...found];
}

// session: sesja z SwbAnon.createSession (z wypełnioną pamięcią zamian), opts: jak wyżej,
// ctx: { personMethod, scopeSheets:Set|null }. Zwraca { bytes, report }.
async function swbAnonCleanZip(bytes, session, opts = {}, ctx = {}) {
  const o = { ...SWB_ANON_ZIP_DEFAULTS, ...opts };
  const zip = await JSZip.loadAsync(bytes);
  const report = { cleaned: {}, leaks: [], warnings: [] };
  const bump = (what, n = 1) => { if (n > 0) report.cleaned[what] = (report.cleaned[what] || 0) + n; };
  const rewrite = (s) => session.rewriteString(s);
  const personMethod = ctx.personMethod || session.methods.person || "pseudo";
  const sheetPaths = await resolveSheetPaths(zip);
  const pathToSheet = Object.fromEntries(Object.entries(sheetPaths).map(([n, p]) => [p, n]));
  let ct = await swbAnonText(zip, "[Content_Types].xml");

  // ── autorzy (notatki, komentarze wątkowe) jako osoby: ta sama zamiana wszędzie ─────
  const authorOut = (name) => {
    const n = String(name || "").trim();
    if (!n) return n;
    const out = session.replaceText("person", n, personMethod === "keep" ? "pseudo" : personMethod);
    session.remember(n, out);
    return out;
  };

  // ── 1. właściwości pliku ─────────────────────────────────────────────
  if (o.metadata) {
    let core = await swbAnonText(zip, "docProps/core.xml");
    if (core) {
      const before = core;
      core = core.replace(/<((?:dc|cp):(?:creator|lastModifiedBy|title|subject|description|keywords|category|contentStatus|identifier))(\s[^>]*)?>[\s\S]*?<\/\1>/g, (all, tag, attrs) => `<${tag}${attrs || ""}></${tag}>`);
      if (core !== before) { zipWriteUtf8Xml(zip, "docProps/core.xml", core); bump("metadata"); }
    }
    let app = await swbAnonText(zip, "docProps/app.xml");
    if (app) {
      const before = app;
      app = app.replace(/<(Company|Manager|HyperlinkBase)>[\s\S]*?<\/\1>/g, "<$1></$1>");
      if (app !== before) { zipWriteUtf8Xml(zip, "docProps/app.xml", app); bump("metadata"); }
    }
    if (zip.file("docProps/custom.xml")) {
      zip.remove("docProps/custom.xml");
      let rels = await swbAnonText(zip, "_rels/.rels");
      if (rels) zipWriteUtf8Xml(zip, "_rels/.rels", rels.replace(/<Relationship\b[^>]*docProps\/custom\.xml[^>]*\/>/g, ""));
      if (ct) ct = ctRemoveOverride(ct, "/docProps/custom.xml");
      bump("metadata");
    }
    let wb = await swbAnonText(zip, "xl/workbook.xml");
    if (wb) {
      const before = wb;
      wb = wb.replace(/<mc:AlternateContent\b[^>]*>(?:(?!<\/mc:AlternateContent>)[\s\S])*?absPath[\s\S]*?<\/mc:AlternateContent>/g, "")
        .replace(/<(?:[\w]+:)?absPath\b[^>]*\/>/g, "")
        .replace(/(<fileSharing\b[^>]*\buserName=")[^"]*(")/g, "$1$2");
      if (wb !== before) { zipWriteUtf8Xml(zip, "xl/workbook.xml", wb); bump("metadata"); }
    }
    for (const p of Object.keys(zip.files)) {
      if (!/^xl\/persons\/[^/]+\.xml$/i.test(p)) continue;
      let xml = await swbAnonText(zip, p);
      const r = swbAnonRewriteAttr(xml, /(<(?:[\w]+:)?person\b[^>]*\bdisplayName=")([^"]*)(")/g, authorOut);
      xml = r.xml.replace(/(\buserId=")[^"]*(")/g, "$1$2").replace(/(\bproviderId=")[^"]*(")/g, "$1None$2");
      zipWriteUtf8Xml(zip, p, xml);
      bump("authors", r.count);
    }
  }

  // ── 2. arkusze: wyniki formuł, nagłówki/stopki, walidacja, hiperłącza, stałe w formułach ──
  for (const [sheetName, path] of Object.entries(sheetPaths)) {
    let xml = await swbAnonText(zip, path);
    if (!xml) continue;
    const before = xml;
    if (o.formulaCache) {
      let n = 0;
      xml = xml.replace(/<c\b([^>]*?)(?<!\/)>((?:(?!<\/c>)[\s\S])*?<f\b(?:(?!<\/c>)[\s\S])*?)<\/c>/g, (all, attrs, inner) => {
        const isStr = /\bt="(str|inlineStr)"/.test(attrs);
        if (!isStr) return all;
        const vm = inner.match(/<v>([\s\S]*?)<\/v>/);
        if (!vm) return all;
        const v0 = swbAnonDecode(vm[1]);
        const v1 = rewrite(v0);
        if (v1 === v0) return all;
        n++;
        return `<c${attrs}>${inner.replace(vm[0], `<v>${swbAnonEncodeText(v1)}</v>`)}</c>`;
      });
      bump("formulaResults", n);
    }
    if (o.links) {
      // stałe tekstowe w formułach (=JEŻELI(A2="Jan Kowalski";…)) — zamiana spójna z komórkami
      let n = 0;
      xml = xml.replace(/(<f\b[^>]*>)([^<]+)(<\/f>)/g, (all, a, body, c) => {
        const f0 = swbAnonDecode(body);
        if (!f0.includes('"')) return all;
        const f1 = swbAnonRewriteQuoted(f0, rewrite, false);
        if (f1 === f0) return all;
        n++;
        return `${a}${swbAnonEncodeText(f1)}${c}`;
      });
      bump("formulaConstants", n);
      const dv = swbAnonRewriteTextNodes(xml, /(<(?:[\w]+:)?(?:formula1|formula2|f)>)([^<]*)(<\/(?:[\w]+:)?(?:formula1|formula2|f)>)/g, (s) => swbAnonRewriteQuoted(s, rewrite, true));
      xml = dv.xml;
      const hl = swbAnonRewriteAttr(xml, /(<hyperlink\b[^>]*\b(?:display|tooltip|location)=")([^"]*)(")/g, rewrite);
      xml = hl.xml;
      bump("links", hl.count);
    }
    if (o.headersFooters !== "keep") {
      const r = swbAnonRewriteTextNodes(xml, /(<(?:odd|even|first)(?:Header|Footer)>)([\s\S]*?)(<\/(?:odd|even|first)(?:Header|Footer)>)/g, (s) => (o.headersFooters === "clear" ? "" : rewrite(s)));
      xml = r.xml;
      bump("headersFooters", r.count);
    }
    if (xml !== before) zipWriteUtf8Xml(zip, path, xml);

    // hiperłącza zewnętrzne (mailto:, ścieżki) w relacjach arkusza
    if (o.links) {
      const relsPath = swbAnonRelsPathOf(path);
      const rels = await swbAnonText(zip, relsPath);
      if (rels) {
        const r = swbAnonRewriteAttr(rels, /(<Relationship\b[^>]*\bTargetMode="External"[^>]*?\bTarget=")([^"]*)(")|(<Relationship\b[^>]*?\bTarget=")([^"]*)("[^>]*\bTargetMode="External")/g, (s) => {
          const dec = (() => { try { return decodeURIComponent(s); } catch { return s; } })();
          const out = rewrite(dec.replace(/^mailto:/i, ""));
          return out === dec.replace(/^mailto:/i, "") ? s : (/^mailto:/i.test(dec) ? `mailto:${out}` : out);
        });
        if (r.count) { zipWriteUtf8Xml(zip, relsPath, r.xml); bump("links", r.count); }
      }
    }

    // notatki i komentarze wątkowe tego arkusza
    if (o.notes !== "keep") await swbAnonSheetNotes(zip, path, o.notes, rewrite, authorOut, bump, (x) => { ct = x(ct); });
    void sheetName;
  }

  // ── 3. tabela napisów: nieużywane teksty (zostały po podmianie komórek) → puste ──
  if (o.sharedStrings) {
    const sstPath = Object.keys(zip.files).find((p) => /^xl\/sharedStrings\.xml$/i.test(p));
    if (sstPath) {
      const used = new Set();
      for (const p of Object.values(sheetPaths)) {
        const xml = await swbAnonText(zip, p);
        if (!xml) continue;
        for (const m of xml.matchAll(/<c\b[^>]*\bt="s"[^>]*>\s*(?:<f\b[\s\S]*?(?:<\/f>|\/>)\s*)?<v>(\d+)<\/v>/g)) used.add(+m[1]);
      }
      // napisy wskazywane z innych miejsc (np. pamięć tabel przestawnych ich nie używa; zostają tylko komórki)
      let sst = await swbAnonText(zip, sstPath);
      let i = -1, blanked = 0;
      sst = sst.replace(/<si>[\s\S]*?<\/si>|<si\/>/g, (si) => {
        i++;
        if (used.has(i) || si === "<si/>" || /^<si><t\/><\/si>$/.test(si)) return si;
        blanked++;
        return "<si><t/></si>";
      });
      if (blanked) { zipWriteUtf8Xml(zip, sstPath, sst); bump("unusedStrings", blanked); }
    }
  }

  // ── 4. wykresy, rysunki, tabele przestawne, łącza zewnętrzne, nazwy ──────────
  for (const p of Object.keys(zip.files)) {
    if (zip.files[p].dir) continue;
    if (o.charts && /^xl\/(charts|drawings)\/[^/]+\.xml$/i.test(p)) {
      const xml = await swbAnonText(zip, p);
      const r = swbAnonRewriteTextNodes(xml, /(<(?:c:v|a:t)>)([^<]*)(<\/(?:c:v|a:t)>)/g, rewrite);
      if (r.count) { zipWriteUtf8Xml(zip, p, r.xml); bump("charts", r.count); }
    } else if (o.pivots && /^xl\/pivotCache\/pivotCacheDefinition\d*\.xml$/i.test(p)) {
      // Pozycje pola (sharedItems) muszą być UNIKALNE (bez wielkości liter) — rekordy wskazują je
      // numerem. Dwie formy tej samej osoby („B. Bielak”, „Bogdan Bielak”) dostają ten sam
      // pseudonim, więc kolizję numerujemy; tabela i tak odświeży się przy otwarciu.
      let xml = await swbAnonText(zip, p);
      let count = 0;
      xml = xml.replace(/(<cacheField\b[^>]*>)([\s\S]*?)(<\/cacheField>)/g, (all, a, body, z) => {
        const r = swbAnonRewriteAttr(body, /(<s\b[^>]*?\bv=")([^"]*)(")/g, rewrite);
        if (!r.count) return all;
        count += r.count;
        const used = new Set();
        const out = r.xml.replace(/(<s\b[^>]*?\bv=")([^"]*)(")/g, (m, pre, enc, post) => {
          const base = swbAnonDecode(enc);
          let v = base, n = 2;
          while (used.has(v.toLowerCase())) v = `${base} (${n++})`;
          used.add(v.toLowerCase());
          return v === base ? m : `${pre}${swbAnonEncodeAttr(v)}${post}`;
        });
        return `${a}${out}${z}`;
      });
      if (count && !/\brefreshOnLoad=/.test(xml)) xml = xml.replace(/<pivotCacheDefinition\b/, '<pivotCacheDefinition refreshOnLoad="1"');
      if (count) { zipWriteUtf8Xml(zip, p, xml); bump("pivots", count); }
    } else if (o.pivots && /^xl\/pivotCache\/pivotCacheRecords\d*\.xml$/i.test(p)) {
      // w rekordach <s v> to wartości wpisane wprost (pola bez sharedItems) — powtórzenia wolno
      const xml = await swbAnonText(zip, p);
      const r = swbAnonRewriteAttr(xml, /(<s\b[^>]*?\bv=")([^"]*)(")/g, rewrite);
      if (r.count) { zipWriteUtf8Xml(zip, p, r.xml); bump("pivots", r.count); }
    } else if (o.pivots && /^xl\/pivotTables\/[^/]+\.xml$/i.test(p)) {
      // własne nazwy pozycji (n="…") też unikalne w obrębie pola
      let xml = await swbAnonText(zip, p);
      let count = 0;
      xml = xml.replace(/(<pivotField\b[^>]*>)([\s\S]*?)(<\/pivotField>)/g, (all, a, body, z) => {
        const used = new Set();
        const out = body.replace(/(<item\b[^>]*?\bn=")([^"]*)(")/g, (m, pre, enc, post) => {
          const base = rewrite(swbAnonDecode(enc));
          let v = base, n = 2;
          while (used.has(v.toLowerCase())) v = `${base} (${n++})`;
          used.add(v.toLowerCase());
          if (v === swbAnonDecode(enc)) return m;
          count++;
          return `${pre}${swbAnonEncodeAttr(v)}${post}`;
        });
        return `${a}${out}${z}`;
      });
      if (count) { zipWriteUtf8Xml(zip, p, xml); bump("pivots", count); }
    } else if (o.links && /^xl\/externalLinks\/[^/]+\.xml$/i.test(p)) {
      const xml = await swbAnonText(zip, p);
      const r = swbAnonRewriteTextNodes(xml, /(<v>)([^<]*)(<\/v>)/g, rewrite);
      if (r.count) { zipWriteUtf8Xml(zip, p, r.xml); bump("links", r.count); }
    } else if (o.links && p === "xl/workbook.xml") {
      const xml = await swbAnonText(zip, p);
      const r = swbAnonRewriteTextNodes(xml, /(<definedName\b[^>]*>)([^<]*)(<\/definedName>)/g, (s) => swbAnonRewriteQuoted(s, rewrite, false));
      if (r.count) { zipWriteUtf8Xml(zip, p, r.xml); bump("links", r.count); }
    }
  }

  if (ct) zipWriteUtf8Xml(zip, "[Content_Types].xml", ct);
  await forceRecalcOnLoad(zip);
  await dropCalcChain(zip);

  // ── 5. tabele Excela: nazwa kolumny = tekst komórki nagłówka (inaczej Excel „naprawia” plik) ──
  bump("tablesSynced", await swbAnonSyncTables(zip, sheetPaths));

  // ── 6. test szczelności: oryginały w KAŻDEJ części tekstowej, z dokładnym miejscem ──
  const leakRes = swbAnonLeakRegex(session.leakTerms);
  report.partsChecked = Object.keys(zip.files).filter((p) => !zip.files[p].dir && /\.(xml|rels|vml)$/i.test(p)).length;
  report.leaks = await swbAnonLocateLeaks(zip, leakRes, sheetPaths, ctx);
  if (Object.keys(zip.files).some((p) => /^xl\/media\//i.test(p))) report.warnings.push("images");
  if (Object.keys(zip.files).some((p) => /^xl\/revisions\//i.test(p))) report.warnings.push("revisions");
  if (Object.keys(zip.files).some((p) => /^xl\/externalLinks\//i.test(p))) report.warnings.push("externalLinks");
  if (Object.keys(zip.files).some((p) => /^xl\/(vbaProject\.bin|embeddings\/)/i.test(p))) report.warnings.push("embedded");

  const out = await zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
  return { bytes: out, report };
}

// Teksty tabeli napisów (sharedStrings) jako tablica — do odczytu komórek t="s".
async function swbAnonSharedStrings(zip) {
  const path = Object.keys(zip.files).find((p) => /^xl\/sharedStrings\.xml$/i.test(p));
  const xml = path ? await swbAnonText(zip, path) : null;
  if (!xml) return { path: null, list: [] };
  const list = [];
  for (const m of xml.matchAll(/<si>([\s\S]*?)<\/si>|<si\/>/g)) {
    const body = (m[1] || "").replace(/<rPh\b[\s\S]*?<\/rPh>/g, "");
    list.push([...body.matchAll(/<t\b[^>]*>([^<]*)<\/t>/g)].map((x) => swbAnonDecode(x[1])).join(""));
  }
  return { path, list };
}
// Tekst komórki z XML arkusza: <is> (inline), t="s" (sharedStrings), t="str" (wynik formuły), liczba.
function swbAnonCellText(attrs, inner, sst) {
  const t = (attrs.match(/\bt="([^"]+)"/) || [])[1] || "n";
  if (t === "inlineStr") return [...inner.matchAll(/<t\b[^>]*>([^<]*)<\/t>/g)].map((x) => swbAnonDecode(x[1])).join("");
  const v = (inner.match(/<v>([^<]*)<\/v>/) || [])[1];
  if (v == null) return "";
  if (t === "s") return sst[+v] ?? "";
  return swbAnonDecode(v);
}
function swbAnonEachCell(xml, fn) {
  for (const m of xml.matchAll(/<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g)) {
    const ref = (m[1].match(/\br="([A-Z]+\d+)"/) || [])[1];
    if (ref) fn(ref, m[1], m[2] || "");
  }
}

// Tabele Excela: każda kolumna musi się nazywać dokładnie jak komórka nagłówka. Gdy zapis
// zmienił nagłówek (np. tabela poniżej wykrytego wiersza nagłówka arkusza), dostosuj nazwę
// kolumny i odwołania Tabela[Kolumna] w formułach. Zwraca liczbę poprawionych nazw.
async function swbAnonSyncTables(zip, sheetPaths) {
  const sst = (await swbAnonSharedStrings(zip)).list;
  const renames = []; // [stara, nowa]
  for (const path of Object.values(sheetPaths)) {
    const rels = await swbAnonText(zip, swbAnonRelsPathOf(path));
    if (!rels) continue;
    const sheetXml = await swbAnonText(zip, path);
    if (!sheetXml) continue;
    const texts = new Map();
    let loaded = false;
    const dir = path.slice(0, path.lastIndexOf("/"));
    for (const r of parseRelsList(rels)) {
      if (r.external || !/\/table$/i.test(r.type)) continue;
      const tp = resolvePartPath(dir, r.target);
      let txml = await swbAnonText(zip, tp);
      if (!txml) continue;
      const ref = (txml.match(/<table\b[^>]*\bref="([A-Z]+\d+):([A-Z]+\d+)"/) || []);
      const hdr = Number((txml.match(/<table\b[^>]*\bheaderRowCount="(\d+)"/) || [])[1] ?? 1);
      if (!ref[1] || hdr < 1) continue;
      if (!loaded) { swbAnonEachCell(sheetXml, (cref, attrs, inner) => texts.set(cref, swbAnonCellText(attrs, inner, sst))); loaded = true; }
      const start = XLSX.utils.decode_cell(ref[1]);
      const used = new Set();
      let i = 0, changed = false;
      txml = txml.replace(/(<tableColumn\b[^>]*?\bname=")([^"]*)(")/g, (all, a, nameEnc, b) => {
        const old = swbAnonDecode(nameEnc);
        const cellText = texts.get(XLSX.utils.encode_cell({ r: start.r, c: start.c + i })) ?? "";
        i++;
        let want = String(cellText).replace(/\r?\n/g, "_x000a_");
        if (!want.trim() || want === old) { used.add(old.toLowerCase()); return all; }
        let uniq = want, n = 2;
        while (used.has(uniq.toLowerCase())) uniq = `${want}${n++}`;
        used.add(uniq.toLowerCase());
        renames.push([old, uniq]);
        changed = true;
        return `${a}${swbAnonEncodeAttr(uniq)}${b}`;
      });
      if (changed) zipWriteUtf8Xml(zip, tp, txml);
    }
  }
  if (!renames.length) return 0;
  // odwołania strukturalne: Tabela[Stara], [#This Row],[Stara], [Stara]:[Inna]
  const esc = (x) => x.replace(/['#[\]]/g, "'$&");
  const fix = (f) => renames.reduce((acc, [o, n]) => acc.split(`[${esc(o)}]`).join(`[${esc(n)}]`), f);
  for (const p of Object.keys(zip.files)) {
    if (!/^xl\/(worksheets\/[^/]+|tables\/[^/]+|workbook)\.xml$/i.test(p)) continue;
    const xml = await swbAnonText(zip, p);
    const out = xml.replace(/(<(?:f|calculatedColumnFormula|totalsRowFormula|definedName)\b[^>]*>)([^<]*)(<\/(?:f|calculatedColumnFormula|totalsRowFormula|definedName)>)/g, (all, a, body, c) => {
      const f0 = swbAnonDecode(body);
      const f1 = fix(f0);
      return f1 === f0 ? all : `${a}${swbAnonEncodeText(f1)}${c}`;
    });
    if (out !== xml) zipWriteUtf8Xml(zip, p, out);
  }
  return renames.length;
}

// Gdzie dokładnie zostały oryginały. Zwraca listę wpisów:
// { kind, sheet?, label?, items:[{ term, refs:[{ ref, header, formula }] }] }
//  kind: "cells" (komórki arkusza), "sheetOther" (lista wyboru, nagłówek wydruku… w arkuszu),
//        "sheetName", "definedName", "tableName", "tableColumn", "appTitles", "part".
async function swbAnonLocateLeaks(zip, res, sheetPaths, ctx = {}) {
  if (!res) return [];
  const out = [];
  const sst = await swbAnonSharedStrings(zip);
  const pathToSheet = Object.fromEntries(Object.entries(sheetPaths).map(([n, p]) => [p, n]));
  const headerCells = ctx.headerCells || new Set();
  const protectedCells = ctx.protectedCells || new Set();
  const covered = new Set([sst.path, ...Object.values(sheetPaths)]);
  const addItems = (entry, terms, refInfo) => {
    terms.forEach((term) => {
      let it = entry.items.find((x) => x.term.toLowerCase() === term.toLowerCase());
      if (!it) { it = { term, refs: [] }; entry.items.push(it); }
      if (refInfo) it.refs.push(refInfo);
    });
  };
  for (const [sheetName, path] of Object.entries(sheetPaths)) {
    const xml = await swbAnonText(zip, path);
    if (!xml) continue;
    const cellsEntry = { kind: "cells", sheet: sheetName, items: [] };
    swbAnonEachCell(xml, (ref, attrs, inner) => {
      const text = swbAnonCellText(attrs, inner, sst.list);
      if (!text) return;
      const found = swbAnonLeaksIn(text, res);
      if (!found.length) return;
      const key = `${sheetName}!${ref}`;
      addItems(cellsEntry, found, { ref, header: headerCells.has(key) || protectedCells.has(key), formula: /<f\b/.test(inner) });
    });
    if (cellsEntry.items.length) out.push(cellsEntry);
    const rest = swbAnonDecode(xml.replace(/<sheetData\b[\s\S]*?<\/sheetData>/, ""));
    const other = swbAnonLeaksIn(rest, res);
    if (other.length) { const e = { kind: "sheetOther", sheet: sheetName, items: [] }; addItems(e, other); out.push(e); }
  }
  for (const p of Object.keys(zip.files).sort()) {
    if (zip.files[p].dir || covered.has(p) || !/\.(xml|rels|vml)$/i.test(p)) continue;
    const raw = await swbAnonText(zip, p);
    if (p === "xl/workbook.xml") {
      const kinds = [
        ["sheetName", [...raw.matchAll(/<sheet\b[^>]*\bname="([^"]*)"/g)].map((m) => swbAnonDecode(m[1]))],
        ["definedName", [...raw.matchAll(/<definedName\b[^>]*\bname="([^"]*)"/g)].map((m) => swbAnonDecode(m[1]))],
        ["part", [...raw.matchAll(/<definedName\b[^>]*>([^<]*)<\/definedName>/g)].map((m) => swbAnonDecode(m[1]))],
      ];
      kinds.forEach(([kind, list]) => {
        const e = { kind, label: kind === "part" ? "definedNameValue" : null, items: [] };
        list.forEach((txt) => addItems(e, swbAnonLeaksIn(txt, res)));
        if (e.items.length) out.push(e);
      });
      continue;
    }
    if (/^xl\/tables\//i.test(p)) {
      const tname = swbAnonDecode((raw.match(/<table\b[^>]*\bdisplayName="([^"]*)"/) || [])[1] || p);
      const e1 = { kind: "tableName", label: tname, items: [] };
      addItems(e1, swbAnonLeaksIn(tname, res));
      const e2 = { kind: "tableColumn", label: tname, items: [] };
      [...raw.matchAll(/<tableColumn\b[^>]*\bname="([^"]*)"/g)].forEach((m) => addItems(e2, swbAnonLeaksIn(swbAnonDecode(m[1]), res)));
      [e1, e2].forEach((e) => { if (e.items.length) out.push(e); });
      continue;
    }
    const found = swbAnonLeaksIn(swbAnonDecode(raw), res);
    if (!found.length) continue;
    const e = { kind: p === "docProps/app.xml" ? "appTitles" : "part", label: p, items: [] };
    addItems(e, found);
    out.push(e);
  }
  // teksty z sharedStrings bez żadnej komórki (nie powinno się zdarzyć po czyszczeniu)
  if (sst.path) {
    const inCells = new Set(out.filter((e) => e.kind === "cells").flatMap((e) => e.items.map((i) => i.term.toLowerCase())));
    const left = sst.list.flatMap((t) => swbAnonLeaksIn(t, res)).filter((t) => !inCells.has(t.toLowerCase()));
    if (left.length) { const e = { kind: "part", label: sst.path, items: [] }; addItems(e, [...new Set(left)]); out.push(e); }
  }
  void pathToSheet;
  return out;
}

// Notatki (comments + VML) i komentarze wątkowe jednego arkusza.
async function swbAnonSheetNotes(zip, sheetPath, mode, rewrite, authorOut, bump, editCt) {
  const dir = sheetPath.slice(0, sheetPath.lastIndexOf("/"));
  const relsPath = swbAnonRelsPathOf(sheetPath);
  let rels = await swbAnonText(zip, relsPath);
  if (!rels) return;
  const list = parseRelsList(rels);
  let sheetXml = null;
  for (const r of list) {
    if (r.external) continue;
    const isComments = /\/comments$/i.test(r.type);
    const isThreaded = /\/threadedComment$/i.test(r.type);
    const isVml = /\/vmlDrawing$/i.test(r.type);
    if (!isComments && !isThreaded && !(isVml && mode === "remove")) continue;
    const p = resolvePartPath(dir, r.target);
    const f = zip.file(p);
    if (mode === "remove") {
      if (isVml) {
        // VML samych notatek (bez kontrolek formularza) usuwamy razem z notatkami
        const vml = f ? await f.async("string") : "";
        if (/ObjectType="(?!Note)/.test(vml)) continue;
        sheetXml = removeLegacyDrawing(sheetXml ?? (await swbAnonText(zip, sheetPath)));
      }
      let removed = 0;
      if (f) { removed = isComments ? ((await f.async("string")).match(/<comment\b/g) || []).length : isThreaded ? ((await f.async("string")).match(/<threadedComment\b/g) || []).length : 0; zip.remove(p); }
      rels = relsRemove(rels, r.id);
      editCt((ct) => (ct ? ctRemoveOverride(ct, `/${p}`) : ct));
      const sub = swbAnonRelsPathOf(p);
      if (zip.file(sub)) zip.remove(sub);
      bump("notesRemoved", removed);
      continue;
    }
    if (!f) continue;
    let xml = await f.async("string");
    if (isComments) {
      const a = swbAnonRewriteTextNodes(xml, /(<author>)([^<]*)(<\/author>)/g, authorOut);
      // tekst notatki bywa w kilku przebiegach; zaczyna się często od „Autor:” — przepisujemy każdy <t>
      const t = swbAnonRewriteTextNodes(a.xml, /(<t\b[^>]*>)([^<]*)(<\/t>)/g, rewrite);
      xml = t.xml;
      bump("authors", a.count);
      bump("notes", t.count);
    } else {
      const t = swbAnonRewriteTextNodes(xml, /(<(?:[\w]+:)?text>)([^<]*)(<\/(?:[\w]+:)?text>)/g, rewrite);
      xml = t.xml;
      bump("notes", t.count);
    }
    zipWriteUtf8Xml(zip, p, xml);
  }
  if (mode === "remove") {
    if (sheetXml !== null) zipWriteUtf8Xml(zip, sheetPath, sheetXml);
    if (parseRelsList(rels).length) zipWriteUtf8Xml(zip, relsPath, rels);
    else zip.remove(relsPath);
  }
}
