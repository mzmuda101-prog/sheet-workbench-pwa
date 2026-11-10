// xlsx-grow.js — przy zapisie: zakresy w pliku rosną razem z danymi dopisanymi pod spodem
// (Etap 1a „nowe wiersze”, 2026-11-10).
//
// Dlaczego: xlsx-patch.js dopisuje nowe wiersze do <sheetData>, ale zakresy opisane gdzie
// indziej zostawały przy starym końcu danych. W Excelu dopisany wiersz wypadał wtedy poza
// tabelę (bez jej stylu, filtra i formuł kolumn), poza autofiltr, formatowanie warunkowe
// i listy wyboru. Excel przy wpisywaniu tuż pod danymi rozszerza je sam — tu robimy to samo.
//
// Zasada (jak Excel przy dopisywaniu): rośnie zakres, który KOŃCZY SIĘ na ostatnim wierszu
// danych z otwartego pliku (fromRow) i ma co najmniej dwa wiersze. Nie ruszamy:
//   • zakresów sięgających dalej niż dane (lista wyboru do wiersza 200 — już obejmuje nowe),
//   • pojedynczych komórek w ostatnim wierszu (np. jedna komórka sumy),
//   • całych kolumn (A:A) i zakresów kończących się wyżej,
//   • tabel z wierszem sumy (Excel wstawia wtedy NAD sumą — to wstawianie w środku, Etap 4).
// Tylko tekst poza <sheetData> (jak xlsx-patch.js: bez DOMParsera na całym arkuszu).
// Wiersze 1-based, jak w adresach Excela.

const swbGrow = (() => {
  const CELL_RE = /^(\$?)([A-Z]{1,3})(\$?)(\d+)$/;

  // „A3:H42” → rośnie do toRow, gdy kończy się na fromRow i zaczyna wyżej. Inne bez zmian.
  function growRange(range, fromRow, toRow) {
    const parts = String(range).split(":");
    if (parts.length !== 2) return range; // pojedyncza komórka albo coś nietypowego
    const a = CELL_RE.exec(parts[0]);
    const b = CELL_RE.exec(parts[1]);
    if (!a || !b) return range; // A:A, 3:3, odwołania do innych arkuszy…
    const r1 = Number(a[4]);
    const r2 = Number(b[4]);
    if (r2 !== fromRow || r1 >= r2) return range;
    return `${parts[0]}:${b[1]}${b[2]}${b[3]}${toRow}`;
  }

  // sqref = lista zakresów rozdzielonych spacjami
  function growSqref(sqref, fromRow, toRow) {
    return String(sqref).split(/\s+/).filter(Boolean).map((r) => growRange(r, fromRow, toRow)).join(" ");
  }

  function growAttr(tagXml, attr, fromRow, toRow, fn = growRange) {
    return tagXml.replace(new RegExp(`(\\s${attr}=")([^"]*)(")`), (m, p1, v, p3) => p1 + fn(v, fromRow, toRow) + p3);
  }

  // Część arkusza bez <sheetData> (z prefiksem przestrzeni nazw albo bez) — tylko tu szukamy.
  function splitOutsideSheetData(xml) {
    const m = /<((?:\w+:)?)sheetData[\s>/]/.exec(xml);
    if (!m) return null;
    const open = m.index;
    const prefix = m[1];
    const openEnd = xml.indexOf(">", open);
    if (openEnd < 0) return null;
    if (xml[openEnd - 1] === "/") return { before: xml.slice(0, openEnd + 1), data: "", after: xml.slice(openEnd + 1) }; // <sheetData/>
    const closeTag = `</${prefix}sheetData>`;
    const close = xml.indexOf(closeTag, openEnd);
    if (close < 0) return null;
    return { before: xml.slice(0, openEnd + 1), data: xml.slice(openEnd + 1, close), after: xml.slice(close) };
  }

  // <dimension ref="A1:K42"/> — koniec co najmniej do toRow (Excel liczy z niego „używany zakres”)
  function growDimension(before, toRow) {
    return before.replace(/(<(?:\w+:)?dimension\s[^>]*\bref=")([^"]*)(")/, (m, p1, ref, p3) => {
      const parts = ref.split(":");
      const end = CELL_RE.exec(parts[parts.length - 1]);
      if (!end || Number(end[4]) >= toRow) return m;
      return `${p1}${parts[0]}:${end[2]}${toRow}${p3}`;
    });
  }

  function growSheetXml(xml, fromRow, toRow) {
    if (!(toRow > fromRow) || !(fromRow >= 1)) return xml;
    const parts = splitOutsideSheetData(xml);
    if (!parts) return xml;
    const before = growDimension(parts.before, toRow);
    let after = parts.after;
    // autofiltr arkusza
    after = after.replace(/<(?:\w+:)?autoFilter\s[^>]*>/g, (tag) => growAttr(tag, "ref", fromRow, toRow));
    // formatowanie warunkowe i listy wyboru (zwykłe: atrybut sqref)
    after = after.replace(/<(?:\w+:)?conditionalFormatting\s[^>]*>/g, (tag) => growAttr(tag, "sqref", fromRow, toRow, growSqref));
    after = after.replace(/<(?:\w+:)?dataValidation\s[^>]*>/g, (tag) => growAttr(tag, "sqref", fromRow, toRow, growSqref));
    // rozszerzenia x14 (nowsze typy reguł): zakres w <xm:sqref>…</xm:sqref>
    after = after.replace(/(<(\w+:)?sqref>)([^<]*)(<\/\2?sqref>)/g, (m, open, _p, v, close) => open + growSqref(v, fromRow, toRow) + close);
    return before + parts.data + after;
  }

  // Tabela Excela (xl/tables/tableN.xml): ref i autoFilter rosną, gdy tabela kończy się na
  // ostatnim wierszu danych. Wiersz sumy (totalsRowCount ≥ 1) = bez zmian.
  function growTableXml(xml, fromRow, toRow) {
    if (!(toRow > fromRow)) return xml;
    const root = /<(?:\w+:)?table\s[^>]*>/.exec(xml);
    if (!root) return xml;
    const tag = root[0];
    const totals = /\stotalsRowCount="(\d+)"/.exec(tag);
    if (totals && Number(totals[1]) > 0) return xml;
    const ref = /\sref="([^"]*)"/.exec(tag);
    if (!ref || growRange(ref[1], fromRow, toRow) === ref[1]) return xml;
    let out = xml.slice(0, root.index) + growAttr(tag, "ref", fromRow, toRow) + xml.slice(root.index + tag.length);
    out = out.replace(/<(?:\w+:)?autoFilter\s[^>]*>/, (t) => growAttr(t, "ref", fromRow, toRow));
    return out;
  }

  // Ścieżki tabel arkusza z jego pliku relacji (xl/worksheets/_rels/sheetN.xml.rels).
  function tablePathsFromRels(relsXml, sheetPath) {
    const dir = sheetPath.replace(/[^/]+$/, ""); // xl/worksheets/
    const out = [];
    const re = /<Relationship\s[^>]*>/g;
    let m;
    while ((m = re.exec(relsXml))) {
      const tag = m[0];
      if (!/Type="[^"]*\/table"/.test(tag) || /TargetMode="External"/.test(tag)) continue;
      const target = (/Target="([^"]*)"/.exec(tag) || [])[1];
      if (!target) continue;
      out.push(resolvePath(dir, target));
    }
    return out;
  }
  function resolvePath(dir, target) {
    if (target.startsWith("/")) return target.slice(1);
    const segs = (dir + target).split("/");
    const res = [];
    segs.forEach((s) => { if (s === "..") res.pop(); else if (s && s !== ".") res.push(s); });
    return res.join("/");
  }

  // Cały plik: dla każdego arkusza z { fromRow, toRow } — arkusz i jego tabele.
  async function growZip(zip, sheetPaths, growBySheet) {
    for (const sheetName of Object.keys(growBySheet || {})) {
      const g = growBySheet[sheetName];
      const path = sheetPaths[sheetName];
      if (!g || !path || !(g.toRow > g.fromRow) || !zip.file(path)) continue;
      const xml = await zip.file(path).async("string");
      const grown = growSheetXml(xml, g.fromRow, g.toRow);
      if (grown !== xml) zipWriteUtf8Xml(zip, path, grown);
      const relsPath = path.replace(/([^/]+)$/, "_rels/$1.rels");
      const rels = zip.file(relsPath) ? await zip.file(relsPath).async("string") : "";
      for (const tpath of tablePathsFromRels(rels, path)) {
        if (!zip.file(tpath)) continue;
        const txml = await zip.file(tpath).async("string");
        const tgrown = growTableXml(txml, g.fromRow, g.toRow);
        if (tgrown !== txml) zipWriteUtf8Xml(zip, tpath, tgrown);
      }
    }
  }

  return { growRange, growSqref, growSheetXml, growTableXml, tablePathsFromRels, growZip };
})();

if (typeof module !== "undefined" && module.exports) module.exports = swbGrow;
